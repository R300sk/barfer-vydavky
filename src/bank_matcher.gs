/**
 * Bank matcher – spáruje posledný BankImport_* s mesiacom YYYY_MM.
 * Heuristika: suma ±0.05 (netto vs brutto preveríme 23/20/10 %),
 * dátum ±3 dni, VS/KS/SS, názov dodávateľa.
 * Výstup zapisuje do OUT bloku import listu: MATCH_STATUS, RULE, SCORE, NORMALIZED_AMOUNT, EXTRACTED_*.
 * UNMATCHED riadky exportuje do listu Unmatched_YYYY_MM.
 */

// ------- lokálne fallback helpery (ak nie sú dostupné z bank_import.gs) -------
function __vatRates__() { try { return VAT_RATES && VAT_RATES.length ? VAT_RATES : [0.23,0.20,0.10]; } catch(e){ return [0.23,0.20,0.10]; } }
function __parseMoney__(v){ if(v==null||v==='')return null; if(typeof v==='number')return v; var s=String(v).trim().replace(/\u00A0/g,' ').replace(/ /g,''); if(/\d+\.\d{3},\d{2}$/.test(s))s=s.replace(/\./g,'').replace(',', '.'); else if(/\d+,\d{3}\.\d{2}$/.test(s))s=s.replace(/,/g,''); else s=s.replace(',', '.'); var n=Number(s); return isNaN(n)?null:n; }
function __guessNetFromGross__(gross){ if(gross==null) return ''; var rates=__vatRates__(); for(var i=0;i<rates.length;i++){ var r=rates[i]; var net = Math.round((gross/(1+r))*100)/100; var back= Math.round((net*(1+r))*100)/100; if(Math.abs(back-gross)<=0.01) return net; } return gross; }
function __ensureOutBlockStrict__(sh, headers){ var lastCol=sh.getLastColumn(); if(lastCol<1)lastCol=1; var header=sh.getRange(1,1,1,lastCol).getValues()[0]; outer: for(var c=1;c<=lastCol-headers.length+1;c++){ for(var j=0;j<headers.length;j++){ if(header[c-1+j]!==headers[j]) continue outer; } return {col:c,count:headers.length}; } var start=lastCol+1; sh.getRange(1,start,1,headers.length).setValues([headers]); return {col:start,count:headers.length}; }
function __dropDuplicateHeaderColumns__(sh){ var last=sh.getLastColumn(); if(last<1) return; var hdr=sh.getRange(1,1,1,last).getValues()[0]; for(var c=last;c>=1;c--){ var name=hdr[c-1]; if(name && typeof name==='string' && /\.1$/.test(name)) sh.deleteColumn(c); } }
function __findAndComputeAmount__(header,row){ var hLower=header.map(h=>String(h||'').toLowerCase()); var amountIdx=hLower.indexOf('amount'); var creditIdx=hLower.indexOf('credit'); var debitIdx=hLower.indexOf('debit'); var cdIdx=hLower.indexOf('creditdebit'); var raw=null; if(amountIdx>=0) raw=__parseMoney__(row[amountIdx]); else if(creditIdx>=0&&debitIdx>=0){ var credit=__parseMoney__(row[creditIdx])||0; var debit=__parseMoney__(row[debitIdx])||0; raw=credit-debit; } if(cdIdx>=0){ var cd=String(row[cdIdx]||'').trim().toUpperCase(); if(cd==='DBIT' && raw>0) raw=-raw; } return raw; }
function __parseDate__(v){ if(v instanceof Date) return v; var s=String(v||'').trim(); var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/); if(m) return new Date(Number(m[1]),Number(m[2])-1,Number(m[3])); var d=new Date(s); return isNaN(d.getTime())?null:d; }
function __daysDiff__(a,b){ var MS=24*3600*1000; return Math.round((a-b)/MS); }

function __extractIds__(row, header){
  var map={}; header.forEach(function(h,i){ map[String(h||'').toLowerCase()]=i; });
  var u = row[map['ustrd']] || row[map['payername']] || row[map['endtoendid']] || '';
  u = String(u||'');
  var find = function(re){ var m=(u.match(re)||[]); return m[1]||''; };
  // VS / SS / KS – skús rôzne formy
  var vs = find(/\bVS[:\s]?(\d{2,10})\b/i) || find(/\bVar(?:\.|iabiln[yý])?\s*symb\w*\s*(\d{2,10})\b/i) || '';
  var ks = find(/\bKS[:\s]?(\d{2,10})\b/i) || '';
  var ss = find(/\bSS[:\s]?(\d{2,10})\b/i) || '';
  return { vs: vs, ks: ks, ss: ss, freeText: u };
}

// ---- ciele (stĺpce) v mesačnom liste – hľadáme flexibilne
var MONTH_COLS = {
  date:  ['Date','Dátum','Datum','BookDate','ValueDate'],
  amount:['Amount','Suma','Bez DPH','AmountNet','Suma bez DPH'],
  vendor:['Dodávateľ','Vendor','Supplier','Name','Názov'],
  vs:    ['VS','VarSymbol','Variabilný symbol'],
  ks:    ['KS','Konštantný symbol'],
  ss:    ['SS','Špecifický symbol','Specific symbol'],
  note:  ['Note','Poznámka','Text'],
  cost:  ['CostCenter','Nákladové stredisko','Stredisko'],
  rev:   ['RevenueCenter','Príjmové stredisko']
};

function __colIndex__(header, names){
  for(var i=0;i<names.length;i++){ var idx = header.indexOf(names[i]); if(idx>=0) return idx; }
  // skús-case-insensitive
  var lower = header.map(h=>String(h||'').toLowerCase());
  for(var j=0;j<names.length;j++){ var id = lower.indexOf(String(names[j]).toLowerCase()); if(id>=0) return id; }
  return -1;
}

function __scoreMatch__(imp, monthRow){
  // imp: {amount,gross,net,date,ids{vs,ks,ss,freeText}}
  // monthRow: {amountNet,date,vendor,vs,ks,ss}
  var score=0, reasons=[];
  // suma: ±0.05, porovnáme imp.net voči month.amountNet; ak net chýba, použijeme gross→net guess
  var inNet = (typeof imp.net==='number')? imp.net : __guessNetFromGross__(imp.gross);
  if(typeof inNet==='number' && typeof monthRow.amountNet==='number'){
    if(Math.abs(inNet - monthRow.amountNet) <= 0.05){ score+=60; reasons.push('amount±0.05'); }
  }
  // dátum ±3 dni
  if(imp.date && monthRow.date){ var dd=Math.abs(__daysDiff__(imp.date, monthRow.date)); if(dd<=3){ score+=25; reasons.push('date±3'); } }
  // VS/KS/SS
  if(imp.ids.vs && monthRow.vs && String(imp.ids.vs)===String(monthRow.vs)){ score+=15; reasons.push('VS'); }
  if(imp.ids.ks && monthRow.ks && String(imp.ids.ks)===String(monthRow.ks)){ score+=8; reasons.push('KS'); }
  if(imp.ids.ss && monthRow.ss && String(imp.ids.ss)===String(monthRow.ss)){ score+=8; reasons.push('SS'); }
  // vendor v texte
  if(imp.ids.freeText && monthRow.vendor){
    var ft = String(imp.ids.freeText).toLowerCase();
    var vn = String(monthRow.vendor).toLowerCase();
    if(ft && vn && (ft.indexOf(vn)>=0 || vn.indexOf(ft)>=0)){ score+=10; reasons.push('vendor~'); }
  }
  return {score:score, reason:reasons.join(', ')};
}

// ---- hlavné API ----

/** Spáruj posledný import s daným mesiacom (názov listu YYYY_MM). */
function matchLastBankImportToMonth(monthName){
  if(!monthName || !/^\d{4}_\d{2}$/.test(monthName)) throw new Error('Očakávam názov mesiaca vo formáte YYYY_MM (napr. 2025_01).');

  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var imp = null, latestTime=0;
  for(var i=0;i<sheets.length;i++){
    var sh=sheets[i], n=sh.getName();
    if(n.indexOf('BankImport_')===0){
      var t = sh.getLastUpdated ? sh.getLastUpdated().getTime() : sh.getMaxRows(); // fallback
      if(t>=latestTime){ latestTime=t; imp=sh; }
    }
  }
  if(!imp) throw new Error('Nenašiel som list začínajúci „BankImport_“.');

  var month = ss.getSheetByName(monthName);
  if(!month) throw new Error('Nenašiel som mesačný list: '+monthName);

  // --- priprav import list ---
  __dropDuplicateHeaderColumns__(imp);
  var impHdr = imp.getRange(1,1,1,imp.getLastColumn()).getValues()[0];
  var OUT_HEADERS = ['MATCH_STATUS','MATCH_RULE','MATCH_SCORE','NORMALIZED_AMOUNT','EXTRACTED_ICO','EXTRACTED_VAR','EXTRACTED_SPEC','EXTRACTED_KS'];
  var outLoc = __ensureOutBlockStrict__(imp, OUT_HEADERS);

  var impRows = imp.getRange(2,1,imp.getLastRow()-1,imp.getLastColumn()).getValues();
  var out = new Array(impRows.length);

  // --- načítaj mesiac ---
  var mHdr = month.getRange(1,1,1,month.getLastColumn()).getValues()[0];

  var idxDate = __colIndex__(mHdr, MONTH_COLS.date);
  var idxAmt  = __colIndex__(mHdr, MONTH_COLS.amount);
  var idxVend = __colIndex__(mHdr, MONTH_COLS.vendor);
  var idxVS   = __colIndex__(mHdr, MONTH_COLS.vs);
  var idxKS   = __colIndex__(mHdr, MONTH_COLS.ks);
  var idxSS   = __colIndex__(mHdr, MONTH_COLS.ss);

  if(idxDate<0 || idxAmt<0){
    throw new Error('Mesačný list '+monthName+' nemá rozpoznateľné stĺpce pre dátum a sumu (skús premenovať hlavičky – pozri MONTH_COLS v bank_matcher.gs).');
  }

  var mRows = month.getRange(2,1,month.getLastRow()-1,month.getLastColumn()).getValues().map(function(r, i){
    return {
      rowIndex: i+2,
      date: __parseDate__(r[idxDate]),
      amountNet: __parseMoney__(r[idxAmt]),
      vendor: idxVend>=0 ? r[idxVend] : '',
      vs:     idxVS>=0 ? r[idxVS] : '',
      ks:     idxKS>=0 ? r[idxKS] : '',
      ss:     idxSS>=0 ? r[idxSS] : ''
    };
  });

  // --- pre každý import riadok nájdi best match ---
  var unmatched = [];
  for(var i=0;i<impRows.length;i++){
    var r = impRows[i];

    var bookDateIdx = impHdr.indexOf('BookDate'); var valueDateIdx = impHdr.indexOf('ValueDate');
    var date = __parseDate__( r[valueDateIdx>=0?valueDateIdx:bookDateIdx] );

    var gross = __findAndComputeAmount__(impHdr, r);
    var net   = (typeof gross==='number') ? __guessNetFromGross__(gross) : '';

    var ids = __extractIds__(r, impHdr);

    var best = {score:-1, reason:'', match:null};
    for(var j=0;j<mRows.length;j++){
      var sc = __scoreMatch__({gross:gross, net:net, date:date, ids:ids}, mRows[j]);
      if(sc.score>best.score){ best={score:sc.score, reason:sc.reason, match:mRows[j]}; }
    }

    if(best.score>=70){
      out[i] = ['MATCHED','amount+date'+(ids.vs?' +VS':''), best.score, net, '', ids.vs||'', ids.ss||'', ids.ks||''];
      // voliteľne: môžeme zapísať referenciu do importu alebo do mesiaca (nechávam zatiaľ len v import liste)
    } else {
      out[i] = ['UNMATCHED', best.reason, best.score, net, '', ids.vs||'', ids.ss||'', ids.ks||''];
      // pripravíme export do Unmatched_*
      unmatched.push([
        date ? Utilities.formatDate(date, (CONFIG && CONFIG.TIMEZONE)||'Europe/Bratislava','yyyy-MM-dd') : '',
        gross, net, ids.vs||'', ids.ks||'', ids.ss||'',
        r.join(' | ')
      ]);
    }
  }

  // zapíš späť do import listu
  if(out.length) imp.getRange(2, outLoc.col, out.length, outLoc.count).setValues(out);

  // vytvor/obnov Unmatched_YYYY_MM
  var unName = 'Unmatched_'+monthName;
  var un = ss.getSheetByName(unName);
  if(!un) un = ss.insertSheet(unName); else un.clear();
  var unHdr = [['Date','Gross','Net','VS','KS','SS','SourceText']];
  un.getRange(1,1,unHdr.length,unHdr[0].length).setValues(unHdr);
  if(unmatched.length) un.getRange(2,1,unmatched.length,unmatched[0].length).setValues(unmatched);
  un.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ Párovanie dokončené.\nMatched: '+ (impRows.length - unmatched.length) + '\nUnmatched: ' + unmatched.length + '\nVýstup v importe + list '+unName);
}

/** Menu handler – vyžiada si YYYY_MM a spustí match. */
function menuMatchWithMonthPrompt(){
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Zadaj mesiac (YYYY_MM)', 'napr. 2025_01', ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton() !== ui.Button.OK) return;
  var monthName = String(resp.getResponseText()||'').trim();
  matchLastBankImportToMonth(monthName);
}
