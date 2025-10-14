/**
 * Bank matcher – spáruje posledný BankImport_* s mesačným listom.
 * Akceptuje mesiac vo formáte YYYY-MM aj YYYY_MM.
 * UNMATCHED riadky exportuje do Unmatched_<mesiac>.
 */

/*** ───────────── Helpers ───────────── ***/
function _vatRates_(){ try{ return (typeof VAT_RATES!=='undefined'&&VAT_RATES.length)?VAT_RATES:[0.23,0.20,0.10]; }catch(e){ return [0.23,0.20,0.10]; } }
function _tz_(){ try{ return (CONFIG&&CONFIG.TIMEZONE)||Session.getScriptTimeZone()||'Europe/Bratislava'; } catch(e){ return 'Europe/Bratislava'; } }
function _parseMoney_(v){ if(v==null||v==='')return null; if(typeof v==='number')return v; var s=String(v).trim().replace(/\u00A0/g,' ').replace(/ /g,''); if(/\d+\.\d{3},\d{2}$/.test(s))s=s.replace(/\./g,'').replace(',', '.'); else if(/\d+,\d{3}\.\d{2}$/.test(s))s=s.replace(/,/g,''); else s=s.replace(',', '.'); var n=Number(s); return isNaN(n)?null:n; }
function _guessNetFromGross_(gross){ if(gross==null) return ''; var rates=_vatRates_(); for (var i=0;i<rates.length;i++){ var r=rates[i], net=Math.round((gross/(1+r))*100)/100, back=Math.round((net*(1+r))*100)/100; if(Math.abs(back-gross)<=0.01) return net; } return gross; }
function _parseDate_(v){ if(v instanceof Date) return v; var s=String(v||'').trim(); var m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/); if(m) return new Date(+m[1],+m[2]-1,+m[3]); var d=new Date(s); return isNaN(d.getTime())?null:d; }
function _days_(a,b){ return Math.round((a-b)/(24*3600*1000)); }
function _dropDupHdrCols_(sh){ var last=sh.getLastColumn(); if(last<1)return; var hdr=sh.getRange(1,1,1,last).getValues()[0]; for(var c=last;c>=1;c--){ var n=hdr[c-1]; if(n&&/\.1$/.test(String(n))) sh.deleteColumn(c); } }
function _ensureOut_(sh, headers){ var last=sh.getLastColumn(); var hdr=sh.getRange(1,1,1,last).getValues()[0]; outer: for(var c=1;c<=last-headers.length+1;c++){ for(var j=0;j<headers.length;j++){ if(hdr[c-1+j]!==headers[j]) continue outer; } return {col:c,count:headers.length}; } var start=last+1; sh.getRange(1,start,1,headers.length).setValues([headers]); return {col:start,count:headers.length}; }
function _findAmount_(hdr,row){ var h=hdr.map(h=>String(h||'').toLowerCase()); var iA=h.indexOf('amount'), iC=h.indexOf('credit'), iD=h.indexOf('debit'), iCD=h.indexOf('creditdebit'); var raw=null; if(iA>=0) raw=_parseMoney_(row[iA]); else if(iC>=0&&iD>=0){ raw=(_parseMoney_(row[iC])||0)-(_parseMoney_(row[iD])||0); } if(iCD>=0){ var cd=String(row[iCD]||'').toUpperCase(); if(cd==='DBIT'&&raw>0) raw=-raw; } return raw; }
function _extractIds_(row,hdr){ var H={}; hdr.forEach(function(h,i){H[String(h||'').toLowerCase()]=i;}); var u=row[H['ustrd']]||row[H['payername']]||row[H['endtoendid']]||''; u=String(u||''); function g(re){ var m=u.match(re); return m?m[1]:''; } var vs=g(/\bVS[:\s]?(\d{2,10})\b/i)||g(/\bVar\w*\s*symb\w*\s*(\d{2,10})\b/i)||''; var ks=g(/\bKS[:\s]?(\d{2,10})\b/i)||''; var ss=g(/\bSS[:\s]?(\d{2,10})\b/i)||''; return {vs:vs,ks:ks,ss:ss,freeText:u}; }
function _monthVariants_(name){ var s=String(name).trim(); var m=s.match(/^(\d{4})[-_](\d{2})$/); if(!m) throw new Error('Očakávam mesiac vo formáte YYYY-MM alebo YYYY_MM (napr. 2025-01).'); var dash=m[1]+'-'+m[2], under=m[1]+'_'+m[2]; return [dash,under]; }
function _resolveMonthSheet_(ss,name){ var v=_monthVariants_(name); return ss.getSheetByName(v[0]) || ss.getSheetByName(v[1]) || null; }

var MONTH_COLS = {
  date:['Date','Dátum','Datum','BookDate','ValueDate'],
  amount:['Amount','Suma','Bez DPH','AmountNet','Suma bez DPH'],
  vendor:['Dodávateľ','Vendor','Supplier','Name','Názov'],
  vs:['VS','VarSymbol','Variabilný symbol'],
  ks:['KS','Konštantný symbol'],
  ss:['SS','Špecifický symbol','Specific symbol'],
  note:['Note','Poznámka','Text']
};
function _idx_(hdr, names){ for(var i=0;i<names.length;i++){ var j=hdr.indexOf(names[i]); if(j>=0) return j; } var low=hdr.map(h=>String(h||'').toLowerCase()); for(var k=0;k<names.length;k++){ var t=low.indexOf(String(names[k]).toLowerCase()); if(t>=0) return t; } return -1; }
function _score_(imp, m){ var s=0, rs=[]; var net=(typeof imp.net==='number')?imp.net:_guessNetFromGross_(imp.gross); if(typeof net==='number'&&typeof m.amountNet==='number'&&Math.abs(net-m.amountNet)<=0.05){ s+=60; rs.push('amount±0.05'); } if(imp.date&&m.date&&Math.abs(_days_(imp.date,m.date))<=3){ s+=25; rs.push('date±3'); } if(imp.ids.vs&&m.vs&&String(imp.ids.vs)===String(m.vs)){ s+=15; rs.push('VS'); } if(imp.ids.ks&&m.ks&&String(imp.ids.ks)===String(m.ks)){ s+=8; rs.push('KS'); } if(imp.ids.ss&&m.ss&&String(imp.ids.ss)===String(m.ss)){ s+=8; rs.push('SS'); } if(imp.ids.freeText&&m.vendor){ var ft=String(imp.ids.freeText).toLowerCase(), vn=String(m.vendor).toLowerCase(); if(ft&&vn&&(ft.indexOf(vn)>=0||vn.indexOf(ft)>=0)){ s+=10; rs.push('vendor~'); } } return {score:s, reason:rs.join(', ')}; }

/*** ───────────── API ───────────── ***/
function matchLastBankImportToMonth(monthName){
  var ss=SpreadsheetApp.getActive();
  // nájdi import list (posledný BankImport_*)
  var imp=null, tMax=0, shs=ss.getSheets();
  for(var i=0;i<shs.length;i++){ var sh=shs[i], n=sh.getName(); if(n.indexOf('BankImport_')===0){ var t=sh.getLastRow(); if(t>=tMax){tMax=t; imp=sh;} } }
  if(!imp) throw new Error('Nenašiel som list začínajúci „BankImport_“.');

  // resolve mesiac
  var month=_resolveMonthSheet_(ss, monthName);
  if(!month) throw new Error('Nenašiel som mesačný list pre: '+monthName+' (skúšané: '+_monthVariants_(monthName).join(', ')+')');

  // priprav hlavičky a OUT blok
  _dropDupHdrCols_(imp);
  var iHdr=imp.getRange(1,1,1,imp.getLastColumn()).getValues()[0];
  var OUT=['MATCH_STATUS','MATCH_RULE','MATCH_SCORE','NORMALIZED_AMOUNT','EXTRACTED_ICO','EXTRACTED_VAR','EXTRACTED_SPEC','EXTRACTED_KS'];
  var oLoc=_ensureOut_(imp, OUT);

  var iRows=imp.getRange(2,1,imp.getLastRow()-1,imp.getLastColumn()).getValues();
  var out=new Array(iRows.length);

  var mHdr=month.getRange(1,1,1,month.getLastColumn()).getValues()[0];
  var idDate=_idx_(mHdr, MONTH_COLS.date);
  var idAmt=_idx_(mHdr, MONTH_COLS.amount);
  var idVend=_idx_(mHdr, MONTH_COLS.vendor);
  var idVS=_idx_(mHdr, MONTH_COLS.vs), idKS=_idx_(mHdr, MONTH_COLS.ks), idSS=_idx_(mHdr, MONTH_COLS.ss);
  if(idDate<0||idAmt<0) throw new Error('Mesačný list postráda rozpoznateľný dátum/amount stĺpec.');

  var mRows=month.getRange(2,1,month.getLastRow()-1,month.getLastColumn()).getValues().map(function(r,i){
    return {rowIndex:i+2, date:_parseDate_(r[idDate]), amountNet:_parseMoney_(r[idAmt]), vendor:idVend>=0?r[idVend]:'', vs:idVS>=0?r[idVS]:'', ks:idKS>=0?r[idKS]:'', ss:idSS>=0?r[idSS]:''};
  });

  var unmatched=[];
  for (var i=0;i<iRows.length;i++){
    var r=iRows[i];
    var d = _parseDate_(r[iHdr.indexOf('ValueDate')] || r[iHdr.indexOf('BookDate')]);
    var g = _findAmount_(iHdr, r);
    var n = (typeof g==='number')? _guessNetFromGross_(g) : '';
    var ids=_extractIds_(r,iHdr);
    var best={score:-1, reason:'', match:null};
    for(var j=0;j<mRows.length;j++){ var sc=_score_({gross:g,net:n,date:d,ids:ids}, mRows[j]); if(sc.score>best.score){ best={score:sc.score, reason:sc.reason, match:mRows[j]}; } }
    if(best.score>=70){
      out[i]=['MATCHED','amount+date'+(ids.vs?' +VS':''),best.score,n,'',ids.vs||'',ids.ss||'',ids.ks||''];
    }else{
      out[i]=['UNMATCHED',best.reason,best.score,n,'',ids.vs||'',ids.ss||'',ids.ks||''];
      unmatched.push([
        d?Utilities.formatDate(d,_tz_(),'yyyy-MM-dd'):'',
        g,n,ids.vs||'',ids.ks||'',ids.ss||'',
        r.join(' | ')
      ]);
    }
  }

  if(out.length) imp.getRange(2,oLoc.col,out.length,oLoc.count).setValues(out);

  var nameVar=_monthVariants_(month.getName())[0]; // použijeme tvar s pomlčkou
  var unName='Unmatched_'+nameVar;
  var un=ss.getSheetByName(unName); if(!un) un=ss.insertSheet(unName); else un.clear();
  un.getRange(1,1,1,7).setValues([['Date','Gross','Net','VS','KS','SS','SourceText']]);
  if(unmatched.length) un.getRange(2,1,unmatched.length,unmatched[0].length).setValues(unmatched);
  un.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ Párovanie dokončené.\nMatched: '+(iRows.length-unmatched.length)+'\nUnmatched: '+unmatched.length+'\nVýstup: Import list + '+unName);
}

/** Prompt (YYYY-MM alebo YYYY_MM). */
function menuMatchWithMonthPrompt(){
  var ui=SpreadsheetApp.getUi();
  var resp=ui.prompt('Zadaj mesiac (YYYY-MM / YYYY_MM)','napr. 2025-01',ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  matchLastBankImportToMonth(String(resp.getResponseText()||'').trim());
}

/** S aktívnym listom (očakáva názov typu 2025-01 alebo 2025_01). */
function menuMatchWithActiveMonth(){
  var sh=SpreadsheetApp.getActiveSheet();
  var name=sh?sh.getName():'';
  if(!name) throw new Error('Aktívny list nemá názov.');
  matchLastBankImportToMonth(name);
}
