/**
 * Bank matcher – páruje posledný BankImport_* s mesačným listom (YYYY-MM / YYYY_MM).
 * Exportuje Unmatched_<mesiac> a Match_Diagnostics_<mesiac>.
 */

/*** Helpers ***/
function _vatRates_(){ try{ return (typeof VAT_RATES!=='undefined'&&VAT_RATES.length)?VAT_RATES:[0.23,0.20,0.10]; }catch(e){ return [0.23,0.20,0.10]; } }
function _tz_(){ try{ return (CONFIG&&CONFIG.TIMEZONE)||Session.getScriptTimeZone()||'Europe/Bratislava'; } catch(e){ return 'Europe/Bratislava'; } }
function _parseMoney_(v){
  if(v==null||v==='')return null;
  if(typeof v==='number')return v;
  var s=String(v)
    .replace(/\u00A0/g,' ')   // pevná medzera
    .replace(/EUR/gi,'')
    .replace(/€/g,'')
    .replace(/[\s]/g,'')      // všetky medzery preč
    .trim();
  // bežné európske formáty
  if (/^\-?\d{1,3}(\.\d{3})*,\d{2}$/.test(s)) s=s.replace(/\./g,'').replace(',', '.');
  else if (/^\-?\d{1,3}(,\d{3})*\.\d{2}$/.test(s)) s=s.replace(/,/g,'');
  else s=s.replace(',', '.');
  var n=Number(s);
  return isNaN(n)?null:n;
}
function _guessNetFromGross_(gross){
  if(gross==null) return '';
  var rates=_vatRates_();
  for (var i=0;i<rates.length;i++){
    var r=rates[i], net=Math.round((gross/(1+r))*100)/100, back=Math.round((net*(1+r))*100)/100;
    if(Math.abs(back-gross)<=0.01) return net;
  }
  return gross;
}
function _parseDate_(v){
  if(v instanceof Date) return v;
  var s=String(v||'').trim();
  var m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);          // yyyy-mm-dd
  if(m) return new Date(+m[1],+m[2]-1,+m[3]);
  var m2=s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/); // dd.mm.yyyy alebo dd/mm/yyyy
  if(m2) return new Date(+m2[3],+m2[2]-1,+m2[1]);
  var d=new Date(s); return isNaN(d.getTime())?null:d;
}
function _days_(a,b){ return Math.round((a-b)/(24*3600*1000)); }
function _dropDupHdrCols_(sh){
  var last=sh.getLastColumn(); if(last<1)return;
  var hdr=sh.getRange(1,1,1,last).getValues()[0];
  for(var c=last;c>=1;c--){ var n=hdr[c-1]; if(n&&/\.1$/.test(String(n))) sh.deleteColumn(c); }
}
function _ensureOut_(sh, headers){
  var last=sh.getLastColumn(); if(last<1){ sh.insertColumnAfter(1); last=1; }
  var hdr=sh.getRange(1,1,1,last).getValues()[0];
  outer: for(var c=1;c<=last-headers.length+1;c++){
    for(var j=0;j<headers.length;j++){ if(hdr[c-1+j]!==headers[j]) continue outer; }
    return {col:c,count:headers.length};
  }
  var start=last+1; sh.getRange(1,start,1,headers.length).setValues([headers]);
  return {col:start,count:headers.length};
}
function _findAmount_(hdr,row){
  var h=hdr.map(function(x){return String(x||'').toLowerCase();});
  var iA=h.indexOf('amount'), iC=h.indexOf('credit'), iD=h.indexOf('debit'), iCD=h.indexOf('creditdebit');
  var raw=null;
  if(iA>=0) raw=_parseMoney_(row[iA]);
  else if(iC>=0&&iD>=0){ raw=(_parseMoney_(row[iC])||0)-(_parseMoney_(row[iD])||0); }
  if(iCD>=0){ var cd=String(row[iCD]||'').toUpperCase(); if(cd==='DBIT'&&raw>0) raw=-raw; }
  return raw;
}
function _extractIds_(row,hdr){
  var H={}; hdr.forEach(function(h,i){H[String(h||'').toLowerCase()]=i;});
  var u=row[H['ustrd']]||row[H['payername']]||row[H['endtoendid']]||'';
  u=String(u||'');
  function g(re){ var m=u.match(re); return m?m[1]:''; }
  var vs=g(/\bVS[:\s]?(\d{2,10})\b/i)||g(/\bVar\w*\s*symb\w*\s*(\d{2,10})\b/i)||'';
  var ks=g(/\bKS[:\s]?(\d{2,10})\b/i)||'';
  var ss=g(/\bSS[:\s]?(\d{2,10})\b/i)||'';
  return {vs:vs,ks:ks,ss:ss,freeText:u};
}
function _monthVariants_(name){
  var s=String(name).trim();
  var m=s.match(/^(\d{4})[-_](\d{2})$/);
  if(!m) throw new Error('Očakávam mesiac vo formáte YYYY-MM alebo YYYY_MM (napr. 2025-01).');
  var dash=m[1]+'-'+m[2], under=m[1]+'_'+m[2]; return [dash,under];
}
function _resolveMonthSheet_(ss,name){ var v=_monthVariants_(name); return ss.getSheetByName(v[0]) || ss.getSheetByName(v[1]) || null; }

var MONTH_COLS = {
  date:   ["Date","Dátum","Datum","Dátum vystavenia","Dátum prijatia","Dátum účtovania","BookDate","ValueDate"],
  // Net (bez DPH)
  amountNet: ["Bez DPH","AmountNet","Suma bez DPH","Cena bez DPH","Celkom bez DPH","Základ DPH","Základ dane","Price (net)","Net"],
  // Gross (s DPH)
  amountGross: ["S DPH","Celkom s DPH","Suma s DPH","Total (gross)","Gross","Cena s DPH","Price (gross)"],
  vendor: ["Dodávateľ","Vendor","Supplier","Name","Názov","Od koho"],
  vs:     ["VS","VarSymbol","Variabilný symbol"],
  ks:     ["KS","Konštantný symbol"],
  ss:     ["SS","Špecifický symbol","Specific symbol"],
  note:   ["Note","Poznámka","Text","Popis"]
};
function _idxExactOrLoose_(hdr, names){
  for(var i=0;i<names.length;i++){ var j=hdr.indexOf(names[i]); if(j>=0) return j; }
  var low=hdr.map(function(h){return String(h||'').toLowerCase();});
  for(var k=0;k<names.length;k++){ var t=low.indexOf(String(names[k]).toLowerCase()); if(t>=0) return t; }
  return -1;
}
function _idxByHeuristic_(sh){ // fallback pre net/gross ak sa nenájdu hlavičky
  var rows = Math.max(0, Math.min(50, sh.getLastRow()-1));
  var cols = sh.getLastColumn();
  if (!rows || !cols) return {date:-1, amount:-1};
  var data = sh.getRange(2,1,rows,cols).getValues();
  var bestDate=-1, bestDateScore=-1, bestAmt=-1, bestAmtScore=-1;
  for (var c=0;c<cols;c++){
    var dCount=0, aCount=0;
    for (var r=0;r<rows;r++){
      var v=data[r][c];
      if (v instanceof Date) dCount++;
      var n=(typeof v==="number")?v:_parseMoney_(v);
      if (typeof n==="number" && !isNaN(n)) aCount++;
    }
    if (dCount>bestDateScore){ bestDateScore=dCount; bestDate=c; }
    if (aCount>bestAmtScore){ bestAmtScore=aCount; bestAmt=c; }
  }
  return {date: bestDate, amount: bestAmt};
}

/*** Scoring ***/
var AMT_EPS = 0.10;
var DATE_EPS_DAYS = 7;

function _amountMatch_(impGross, impNet, monthNet, monthGross){
  // porovnávame voči net aj gross z mesiaca (ak existujú)
  var candidates = [];
  function push(v, tag){ if(typeof v==='number' && !isNaN(v)) candidates.push({val:v, tag:tag}); }
  // import varianty
  var impVariants = function(y, label){
    push(y, 'amount('+label+')');
    push(Math.abs(y), 'amount(abs '+label+')');
    push(-y, 'amount(-'+label+')');
  };
  if(typeof impNet==='number')   impVariants(impNet,'net');
  if(typeof impGross==='number') impVariants(impGross,'gross');
  // mesačné ciele
  var targets = [];
  if(typeof monthNet==='number')   targets.push({val:monthNet, tag:'→month(net)'});
  if(typeof monthGross==='number') targets.push({val:monthGross, tag:'→month(gross)'});
  if(!targets.length) return {ok:false,score:0,mode:'',delta:1e9};

  var best = {ok:false, score:0, mode:'', delta:1e9};
  for (var i=0;i<candidates.length;i++){
    for (var j=0;j<targets.length;j++){
      var delta = Math.abs(candidates[i].val - targets[j].val);
      var sc = (delta<=AMT_EPS) ? 60 : Math.max(0, 60 - (delta*200));
      if (sc>best.score){
        best = {ok: delta<=AMT_EPS, score: sc, mode: candidates[i].tag+' '+targets[j].tag, delta: delta};
      }
    }
  }
  return best;
}

function _score_(imp, m){
  var rs=[];
  var amt=_amountMatch_(imp.gross, imp.net, m.amountNet, m.amountGross);
  var s = amt.score; if(amt.mode) rs.push(amt.mode);
  if(imp.date&&m.date&&Math.abs(_days_(imp.date,m.date))<=DATE_EPS_DAYS){ s+=25; rs.push('date±'+DATE_EPS_DAYS); }
  if(imp.ids.vs&&m.vs&&String(imp.ids.vs)===String(m.vs)){ s+=15; rs.push('VS'); }
  if(imp.ids.ks&&m.ks&&String(imp.ids.ks)===String(m.ks)){ s+=8;  rs.push('KS'); }
  if(imp.ids.ss&&m.ss&&String(imp.ids.ss)===String(m.ss)){ s+=8;  rs.push('SS'); }
  return {score:s, reason:rs.join(', ')};
}

/*** API ***/
function matchLastBankImportToMonth(monthName){
  var ss=SpreadsheetApp.getActive();

  // 1) posledný BankImport_*
  var imp=null, tMax=0, shs=ss.getSheets();
  for(var i=0;i<shs.length;i++){ var sh=shs[i], n=sh.getName(); if(n.indexOf('BankImport_')===0){ var t=sh.getLastRow(); if(t>=tMax){tMax=t; imp=sh;} } }
  if(!imp) throw new Error('Nenašiel som list začínajúci „BankImport_“.');

  // 2) mesiac
  var month=_resolveMonthSheet_(ss, monthName);
  if(!month) throw new Error('Nenašiel som mesačný list pre: '+monthName+' (skúšané: '+_monthVariants_(monthName).join(', ')+')');

  // 3) hlavičky + OUT
  _dropDupHdrCols_(imp);
  var iHdr=imp.getRange(1,1,1,imp.getLastColumn()).getValues()[0];
  var OUT=['MATCH_STATUS','MATCH_RULE','MATCH_SCORE','NORMALIZED_AMOUNT','EXTRACTED_ICO','EXTRACTED_VAR','EXTRACTED_SPEC','EXTRACTED_KS'];
  var oLoc=_ensureOut_(imp, OUT);
  var iRows=imp.getRange(2,1,Math.max(imp.getLastRow()-1,0),imp.getLastColumn()).getValues();
  var out=(iRows.length? new Array(iRows.length): []);

  var mHdr=month.getRange(1,1,1,month.getLastColumn()).getValues()[0];
  var idDate=_idxExactOrLoose_(mHdr, MONTH_COLS.date);

  // snaž sa nájsť NET aj GROSS zvlášť
  var idNet = _idxExactOrLoose_(mHdr, MONTH_COLS.amountNet);
  var idG   = _idxExactOrLoose_(mHdr, MONTH_COLS.amountGross);

  // fallback: heuristika na amount, ak nemáme ani net ani gross
  if (idNet<0 && idG<0){
    var guess=_idxByHeuristic_(month);
    if (idDate<0) idDate=guess.date;
    idNet = (idNet<0? guess.amount: idNet);
  }

  if (idDate<0 || (idNet<0 && idG<0)) throw new Error('Mesačný list postráda rozpoznateľné stĺpce dátum/amount (net/gross).');

  var idVend=_idxExactOrLoose_(mHdr, MONTH_COLS.vendor);
  var idVS=_idxExactOrLoose_(mHdr, MONTH_COLS.vs), idKS=_idxExactOrLoose_(mHdr, MONTH_COLS.ks), idSS=_idxExactOrLoose_(mHdr, MONTH_COLS.ss);

  var mRows=month.getRange(2,1,Math.max(month.getLastRow()-1,0),month.getLastColumn()).getValues().map(function(r,i){
    return {
      rowIndex:i+2,
      date:_parseDate_(r[idDate]),
      amountNet: idNet>=0 ? _parseMoney_(r[idNet]) : null,
      amountGross: idG>=0 ? _parseMoney_(r[idG]) : null,
      vendor:idVend>=0?r[idVend]:'',
      vs:idVS>=0?r[idVS]:'',
      ks:idKS>=0?r[idKS]:'',
      ss:idSS>=0?r[idSS]:''
    };
  });

  // 4) skórovanie + diagnostika
  var unmatched=[], diag=[['ImpRow','ImpDate','ImpGross','ImpNet','BestScore','BestReason','MonthRow','MonthDate','MonthNet','MonthGross','MonthVS','MonthVendor']];
  for (var i=0;i<iRows.length;i++){
    var r=iRows[i];
    var vIdx = Math.max(iHdr.indexOf('ValueDate'), iHdr.indexOf('BookDate'));
    var d = _parseDate_(vIdx>=0 ? r[vIdx] : null);
    var g = _findAmount_(iHdr, r);
    var n = (typeof g==='number')? _guessNetFromGross_(g) : '';
    var ids=_extractIds_(r,iHdr);

    var best={score:-1, reason:'', match:null};
    for(var j=0;j<mRows.length;j++){
      var sc=_score_({gross:g,net:n,date:d,ids:ids}, mRows[j]);
      if(sc.score>best.score){ best={score:sc.score, reason:sc.reason, match:mRows[j]}; }
    }

    if(best.score>=70){
      out[i]=['MATCHED',best.reason,best.score,n,'',ids.vs||'',ids.ss||'',ids.ks||''];
    }else{
      out[i]=['UNMATCHED',best.reason,best.score,n,'',ids.vs||'',ids.ss||'',ids.ks||''];
      unmatched.push([
        d?Utilities.formatDate(d,_tz_(),'yyyy-MM-dd'):'',
        g,n,ids.vs||'',ids.ks||'',ids.ss||'',
        r.join(' | ')
      ]);
    }

    var md = best.match || {};
    diag.push([
      i+2,
      d?Utilities.formatDate(d,_tz_(),'yyyy-MM-dd'):'',
      (typeof g==='number')?g:'',
      (typeof n==='number')?n:'',
      best.score,
      best.reason,
      md.rowIndex||'',
      md.date?Utilities.formatDate(md.date,_tz_(),'yyyy-MM-dd'):'',
      (typeof md.amountNet==='number')?md.amountNet:'',
      (typeof md.amountGross==='number')?md.amountGross:'',
      md.vs||'',
      md.vendor||''
    ]);
  }

  if(out.length) imp.getRange(2,oLoc.col,out.length,oLoc.count).setValues(out);

  // 5) unmatched
  var nameVar=_monthVariants_(month.getName())[0];
  var unName='Unmatched_'+nameVar;
  var un=ss.getSheetByName(unName); if(!un) un=ss.insertSheet(unName); else un.clear();
  un.getRange(1,1,1,7).setValues([['Date','Gross','Net','VS','KS','SS','SourceText']]);
  if(unmatched.length) un.getRange(2,1,unmatched.length,unmatched[0].length).setValues(unmatched);
  un.setFrozenRows(1);

  // 6) diagnostika
  var dgName='Match_Diagnostics_'+nameVar;
  var dg=ss.getSheetByName(dgName); if(!dg) dg=ss.insertSheet(dgName); else dg.clear();
  dg.getRange(1,1,diag.length,diag[0].length).setValues(diag);
  dg.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ Párovanie dokončené.\nMatched: '+(iRows.length-unmatched.length)+'\nUnmatched: '+unmatched.length+'\nVýstup: Import list + '+unName+'\nDiagnostika: '+dgName);
}

/** Prompt (YYYY-MM / YYYY_MM) */
function menuMatchWithMonthPrompt(){
  var ui=SpreadsheetApp.getUi();
  var resp=ui.prompt('Zadaj mesiac (YYYY-MM / YYYY_MM)','napr. 2025-01',ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  matchLastBankImportToMonth(String(resp.getResponseText()||'').trim());
}
/** Aktívny list */
function menuMatchWithActiveMonth(){
  var sh=SpreadsheetApp.getActiveSheet();
  var name=sh?sh.getName():'';
  if(!name) throw new Error('Aktívny list nemá názov.');
  matchLastBankImportToMonth(name);
}
