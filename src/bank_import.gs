/** bank_import.gs
 * Import CSV bankového výpisu a párovanie proti listu Výdavky.
 * - importBankCSVFromDrive(fileId)
 * - importBankCSVText(csvText, shortName)
 * - matchAllBankImportRows(sheetName)
 */

// Konfigurácia
const BANK_IMPORT_CONFIG = {
  BANK_SHEET_PREFIX: "BankImport_",
  TARGET_SHEET: (typeof CONFIG !== 'undefined' && CONFIG.SHEETS) ? CONFIG.SHEETS.VYDAVKY : "Výdavky",
  VAT_RATES: [0.23, 0.20, 0.10],
  DATE_WINDOW_DAYS: 3,
  AMOUNT_TOLERANCE: 0.5
};

/* ===================== 1) IMPORT CSV ===================== */

function importBankCSVFromDrive(fileId) {
  const file = DriveApp.getFileById(fileId);
  const text = file.getBlob().getDataAsString("UTF-8");
  const name = file.getName().replace(/\.[^/.]+$/, "");
  return importBankCSVText(text, name);
}

function importBankCSVText(csvText, shortName) {
  const module = "bank_import";
  logInfo(module, `Import CSV: ${shortName}`);

  const rows = parseCSV_(csvText);
  if (!rows || rows.length === 0) {
    throw new Error("CSV prázdny alebo nevalidný.");
  }

  const sheetName = `${BANK_IMPORT_CONFIG.BANK_SHEET_PREFIX}${shortName}`;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) sheet.clear(); else sheet = ss.insertSheet(sheetName);

  sheet.getRange(1, 1, 1, rows[0].length).setValues([rows[0]]);
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, rows[0].length).setValues(rows.slice(1));
  }

  const extraHeaders = [
    "MATCH_TARGET_ROW","MATCH_SCORE","MATCH_REASON",
    "NORMALIZED_AMOUNT","NORMALIZED_DATE",
    "EXTRACTED_VARIABLE_SYMBOL","EXTRACTED_PAYER"
  ];
  sheet.getRange(1, rows[0].length + 1, 1, extraHeaders.length).setValues([extraHeaders]);

  logInfo(module, `CSV importované do ${sheetName} (${rows.length - 1} riadkov).`);
  return sheetName;
}

/* ===================== 2) PÁROVANIE ===================== */

function matchAllBankImportRows(importSheetName) {
  const module = "bank_import.matchAll";
  logInfo(module, `Spúšťam párovanie pre ${importSheetName}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName(importSheetName);
  if (!importSheet) throw new Error("Import sheet neexistuje: " + importSheetName);

  const headers = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues()[0];
  const data = importSheet.getRange(2, 1, Math.max(0, importSheet.getLastRow() - 1), headers.length).getValues();

  const targetSheet = ss.getSheetByName(BANK_IMPORT_CONFIG.TARGET_SHEET);
  if (!targetSheet) throw new Error("Target sheet neexistuje: " + BANK_IMPORT_CONFIG.TARGET_SHEET);
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const targetData = targetSheet.getRange(2, 1, Math.max(0, targetSheet.getLastRow() - 1), targetHeaders.length).getValues();

  const dateCol = findDateColumnIndex(headers);
  const amountCol = findAmountColumnIndex(headers);
  const payerCol = findPayerColumnIndex(headers);

  const prepared = data.map((row, i) => {
    const rawDate = row[dateCol];
    const rawAmount = row[amountCol];
    const rawPayer = row[payerCol];

    const normDate = normalizeDateValue_(rawDate);
    const normAmount = normalizeAmount(rawAmount);
    const extractedVS = extractVariableSymbol_(row.join(" "));
    const extractedPayer = (rawPayer || "").toString().trim();

    return { rowIndex: i + 2, raw: row, normDate, normAmount, extractedVS, extractedPayer };
  });

  const results = prepared.map(r => {
    const match = findBestMatchForImportRow_(r, targetData, targetHeaders);
    return [
      match ? match.targetRow + 1 : "",
      match ? match.score : 0,
      match ? match.reason : "no-match",
      r.normAmount,
      r.normDate ? formatDate_(r.normDate) : "",
      r.extractedVS || "",
      r.extractedPayer || ""
    ];
  });

  const writeCol = headers.length + 1;
  if (results.length) {
    importSheet.getRange(2, writeCol, results.length, results[0].length).setValues(results);
  }

  logInfo(module, `Párovanie dokončené (${results.length} riadkov).`);
  return results.length;
}

/* ============ 3) MATCHER & HEURISTIKY ============ */

function findBestMatchForImportRow_(importRow, targetData, targetHeaders) {
  const candidates = [];

  for (let i = 0; i < targetData.length; i++) {
    const t = targetData[i];
    const tDate = normalizeDateValue_(t[0]) || null;          // predpoklad: dátum v prvom stĺpci
    const tAmount = findNumericInRow_(t) || 0;                // nájde prvú číselnú sumu v riadku
    const tText = t.join(" ").toString();

    let score = 0;
    const reasons = [];

    // 1) Variabilný symbol
    if (importRow.extractedVS) {
      const vsInTarget = extractVariableSymbol_(tText);
      if (vsInTarget && vsInTarget === importRow.extractedVS) {
        score = Math.max(score, 100);
        reasons.push("variable-symbol-exact");
      }
    }

    // 2) Rovnaká suma ± tolerancia + dátum v okne
    const amtDiff = Math.abs(Number(importRow.normAmount) - Number(tAmount));
    if (amtDiff <= BANK_IMPORT_CONFIG.AMOUNT_TOLERANCE) {
      if (importRow.normDate && tDate && Math.abs(dateDiffDays_(importRow.normDate, tDate)) <= BANK_IMPORT_CONFIG.DATE_WINDOW_DAYS) {
        score = Math.max(score, 80); reasons.push("amount-date-exact");
      } else {
        score = Math.max(score, 60); reasons.push("amount-only");
      }
    }

    // 3) Odstránenie DPH z bankovej sumy
    for (let rate of BANK_IMPORT_CONFIG.VAT_RATES) {
      const approxNet = tryRemoveVAT_(importRow.normAmount, rate);
      const diffNet = Math.abs(approxNet - Number(tAmount || 0));
      if (diffNet <= BANK_IMPORT_CONFIG.AMOUNT_TOLERANCE) {
        if (importRow.normDate && tDate && Math.abs(dateDiffDays_(importRow.normDate, tDate)) <= BANK_IMPORT_CONFIG.DATE_WINDOW_DAYS) {
          score = Math.max(score, 75); reasons.push(`vat-match-${Math.round(rate*100)}%`);
        } else {
          score = Math.max(score, 55); reasons.push(`vat-amount-${Math.round(rate*100)}%`);
        }
      }
    }

    // 4) Fuzzy match názvu platiteľa
    if (importRow.extractedPayer && importRow.extractedPayer.length > 2) {
      const lev = levenshtein_(importRow.extractedPayer.toLowerCase(), tText.toLowerCase());
      const maxLen = Math.max(importRow.extractedPayer.length, tText.length);
      const sim = 1 - (lev / maxLen); // 0..1
      if (sim > 0.6) {
        const bonus = Math.round(40 * sim);
        score = Math.max(score, bonus);
        reasons.push("fuzzy-payer");
      }
    }

    if (score > 0) {
      candidates.push({ targetRow: i + 1, score, reason: reasons.join(";") });
    }
  }

  if (!candidates.length) return null;
  candidates.sort((a, b) => b.score - a.score);
  return candidates[0];
}

/* ============ 4) HELPERS (CSV/normalizácia) ============ */

function parseCSV_(text) {
  const lines = String(text).replace(/\r/g, "").split("\n").filter(l => l.trim().length > 0);
  if (!lines.length) return [];
  const header = lines[0];
  const delimiter = header.includes(";") ? ";" : (header.includes(",") ? "," : ";");
  return lines.map(l => l.split(delimiter).map(c => c.trim()));
}

function normalizeAmount(raw) {
  if (raw === null || raw === undefined) return 0;
  let s = String(raw).trim();
  s = s.replace(/[^\d,.\-]/g, "");
  const comma = s.indexOf(","), dot = s.indexOf(".");
  if (comma !== -1 && dot !== -1) {
    if (comma > dot) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (comma !== -1 && dot === -1) {
    s = s.replace(",", ".");
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

function normalizeDateValue_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = v.toString().trim();
  const m1 = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/); // dd.mm.yyyy
  if (m1) return new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));
  const m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);             // yyyy-mm-dd
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function findDateColumnIndex(headers) {
  const low = headers.map(h => String(h || "").toLowerCase());
  for (let i = 0; i < low.length; i++) {
    if (/(date|datum|dátum|value date|valuta)/i.test(low[i])) return i;
  }
  return 0;
}

function findAmountColumnIndex(headers) {
  const low = headers.map(h => String(h || "").toLowerCase());
  for (let i = 0; i < low.length; i++) {
    if (/(amount|sum|suma|credit|debet|betrag)/i.test(low[i])) return i;
  }
  return Math.max(0, headers.length - 1);
}

function findPayerColumnIndex(headers) {
  const low = headers.map(h => String(h || "").toLowerCase());
  for (let i = 0; i < low.length; i++) {
    if (/(payer|name|meno|popis|vs|variable|konšt|konst)/i.test(low[i])) return i;
  }
  return Math.max(0, 1);
}

function extractVariableSymbol_(text) {
  if (!text) return "";
  const t = String(text);
  const m = t.match(/(?:VS[:\s]*|variable[_\s]*symbol[:\s]*|variabiln[ýy]\s*symbol[:\s]*|VS\D{0,3})(\d{3,12})/i);
  if (m) return m[1];
  const m2 = t.match(/(\d{6,10})/);
  return m2 ? m2[1] : "";
}

function findNumericInRow_(row) {
  for (let c of row) {
    const n = normalizeAmount(c);
    if (Math.abs(n) > 0.001) return n;
  }
  return 0;
}

function tryRemoveVAT_(amount, rate) {
  const gross = Number(amount) || 0;
  return Math.round((gross / (1 + rate)) * 100) / 100;
}

function dateDiffDays_(a, b) {
  const DAY = 1000 * 60 * 60 * 24;
  return Math.round((a.getTime() - b.getTime()) / DAY);
}

/* ============ 5) Levenshtein ============ */
function levenshtein_(a, b) {
  a = a || ""; b = b || "";
  const al = a.length, bl = b.length;
  const dp = Array(al + 1).fill(null).map(() => Array(bl + 1).fill(0));
  for (let i = 0; i <= al; i++) dp[i][0] = i;
  for (let j = 0; j <= bl; j++) dp[0][j] = j;
  for (let i = 1; i <= al; i++) {
    for (let j = 1; j <= bl; j++) {
      const cost = a[i-1] === b[j-1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i-1][j] + 1, dp[i][j-1] + 1, dp[i-1][j-1] + cost);
    }
  }
  return dp[al][bl];
}

// ====== SAFE MATCH WRAPPER (vyhne sa 0-stĺpcovým range) ======
/**
 * Bezpečné párovanie – vytvorí MATCH_* stĺpce ak chýbajú,
 * a zapisuje len keď je čo zapisovať (žiadne 0-column range).
 * Ak zatiaľ nemáme implementované reálne heuristiky, zapíše základné placeholdery,
 * takže nikdy nepadne na setValues().
 */
function safeMatchAllBankImportRows(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet nie je dostupný: " + sheetName);

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    // nič na párovanie – končíme potichu
    return;
  }

  // prečítaj hlavičky
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // požadované výstupné hlavičky
  var OUT_HEADERS = [
    'MATCH_STATUS','MATCH_RULE','MATCH_SCORE',
    'NORMALIZED_AMOUNT','EXTRACTED_ICO','EXTRACTED_VAR','EXTRACTED_SPEC','EXTRACTED_KS'
  ];

  // nájdi prvý voľný stĺpec na konci (alebo ak už hlavičky existujú, použi ich pozície)
  var startCol = header.length + 1;

  // Ak už niektoré z OUT_HEADERS existujú, presunieme startCol na najmenší existujúci index
  for (var h = 0; h < OUT_HEADERS.length; h++) {
    var idx = header.indexOf(OUT_HEADERS[h]);
    if (idx >= 0) startCol = Math.min(startCol, idx + 1);
  }

  // Ak hlavičky chýbajú, dopíšeme ich na koniec
  if (startCol === header.length + 1) {
    sh.getRange(1, startCol, 1, OUT_HEADERS.length).setValues([OUT_HEADERS]);
  } else {
    // Uisti sa, že všetky OUT_HEADERS sú na miestach – ak niektoré chýbajú medzi,
    // dopíšeme ich postupne (jednoduché, ale robustné).
    var existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    for (var oh = 0; oh < OUT_HEADERS.length; oh++) {
      if (existing.indexOf(OUT_HEADERS[oh]) === -1) {
        sh.insertColumnAfter(sh.getLastColumn());
        var newCol = sh.getLastColumn();
        sh.getRange(1, newCol).setValue(OUT_HEADERS[oh]);
      }
    }
  }

  // znovu načítaj posledný stĺpec po prípadnom vkladaní
  var outFirstCol = startCol;
  var outColCount = OUT_HEADERS.length;

  // načítaj dáta (bez hlavičky)
  var data = sh.getRange(2, 1, lastRow - 1, header.length).getValues();
  if (!data.length) return; // žiadne riadky

  // heuristiky – ak zatiaľ nevieme, dáme defaulty, aby nenastala 0-column situácia
  // (prípadné skutočné matchovanie máte v bank_import*.gs; toto je "safety net")
  var out = [];
  // pokus o detekciu stĺpca so sumou
  var amtIdx = Math.max(header.indexOf('Amount'), header.indexOf('Suma'));
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var amountRaw = (amtIdx >= 0) ? row[amtIdx] : '';
    var amountNum = (typeof amountRaw === 'number') ? amountRaw
                   : (amountRaw ? Number(String(amountRaw).replace(',', '.')) : '');
    var norm = (typeof amountNum === 'number' && !isNaN(amountNum)) ? amountNum : '';

    // default – unmatched; MATCH_RULE a SCORE necháme prázdne/0
    out.push(['UNMATCHED','',0, norm,'','','','']);
  }

  // bezpečný zápis – len ak máme aspoň 1 stĺpec
  if (out.length && out[0].length >= 1) {
    sh.getRange(2, outFirstCol, out.length, out[0].length).setValues(out);
  }
}

// ========= ROBUSTNÁ VERZIA SAFE MATCH (pevný blok stĺpcov) =========

/** Nájde/ vytvorí súvislý blok výstupných hlavičiek na konci a vráti {col, count}. */
function __ensureOutBlock__(sh, headers) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) lastCol = 1;

  // Skús nájsť existujúci súvislý blok headers v akomkoľvek mieste
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  outer:
  for (var c = 1; c <= lastCol - headers.length + 1; c++) {
    for (var j = 0; j < headers.length; j++) {
      if (headerRow[c - 1 + j] !== headers[j]) continue outer;
    }
    return { col: c, count: headers.length }; // našli sme
  }

  // Inak ich pridáme na úplný koniec v správnom poradí
  var startCol = lastCol + 1;
  sh.getRange(1, startCol, 1, headers.length).setValues([headers]);
  return { col: startCol, count: headers.length };
}

/** Parsuje peniaze z rôznych formátov (1 234,56 | 1 234,56 | 1,234.56 | -1234,56). */
function __parseMoney__(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;

  var s = String(v).trim();
  // odstraň NBSP aj bežné medzery
  s = s.replace(/\u00A0/g, ' ').replace(/ /g, '');
  // ak je vo formáte 1.234,56 => odstráň bodky (tisíce) a zmeň čiarku na bodku
  if (/\d+\.\d{3},\d{2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  // ak je vo formáte 1,234.56 => odstráň čiarky (tisíce)
  else if (/\d+,\d{3}\.\d{2}$/.test(s)) s = s.replace(/,/g, '');
  // bežný SK: 1234,56
  else s = s.replace(',', '.');

  var n = Number(s);
  return isNaN(n) ? null : n;
}

/** Vracia sumu bez DPH – skúsi viac sadzieb, vráti najbližšie 2 des. */
function __netOfVat__(gross, rates) {
  if (gross == null) return null;
  for (var i = 0; i < rates.length; i++) {
    var r = rates[i];
    var net = gross / (1 + r);
    // zaokrúhlenie na 2 des.
    net = Math.round(net * 100) / 100;
    if (isFinite(net)) return net; // prvý platný stačí (chová sa deterministicky)
  }
  return gross;
}

/** Nájdi index stĺpca so sumou podľa známych názvov; fallback: prvý číselný */
function __findAmountIndex__(header) {
  var candidates = ['Amount', 'Suma', 'AmountWithVat', 'AmountGross', 'Amount EUR', 'Suma EUR'];
  for (var i = 0; i < candidates.length; i++) {
    var idx = header.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  // fallback: prvý stĺpec, kde vo 2. riadku vyzerá hodnota ako číslo
  return -1;
}

/** Hlavný bezpečný matcher – zapisuje len do vlastného bloku na konci listu. */
function safeMatchAllBankImportRows(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet nie je dostupný: " + sheetName);

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return; // nič na párovanie

  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // Náš presný blok výstupu (nepliesť so starými NORMALIZED_DATE a pod.)
  var OUT_HEADERS = [
    'MATCH_STATUS','MATCH_RULE','MATCH_SCORE',
    'NORMALIZED_AMOUNT','EXTRACTED_ICO','EXTRACTED_VAR','EXTRACTED_SPEC','EXTRACTED_KS'
  ];

  var outLoc = __ensureOutBlock__(sh, OUT_HEADERS);
  var outFirstCol = outLoc.col;
  var outColCount = outLoc.count;

  // Načítaj dáta (bez hlavičky) v rozsahu vstupných stĺpcov (1..lastCol)
  var rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  if (!rows.length) return;

  // Nájdeme index sumy
  var amtIdx = __findAmountIndex__(header);

  // VAT sadzby – ak sú v projekte definované, použi tie; inak default
  var rates = (typeof VAT_RATES !== 'undefined' && VAT_RATES && VAT_RATES.length)
    ? VAT_RATES
    : [0.23, 0.20, 0.10];

  var out = new Array(rows.length);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];

    // zober kandidáta na sumu
    var gross = null;
    if (amtIdx >= 0) {
      gross = __parseMoney__(row[amtIdx]);
    } else {
      // fallback: prvá hodnota v riadku, ktorá vyzerá číselne
      for (var c = 0; c < row.length; c++) {
        var tryN = __parseMoney__(row[c]);
        if (tryN != null) { gross = tryN; break; }
      }
    }

    var net = __netOfVat__(gross, rates);

    // (aktuálne placeholder – ďalšie polia doplníme pravidlami neskôr)
    out[i] = [
      'UNMATCHED', // MATCH_STATUS
      '',          // MATCH_RULE
      0,           // MATCH_SCORE
      (net != null ? net : ''), // NORMALIZED_AMOUNT
      '', '', '', '' // EXTRACTED_*
    ];
  }

  // bezpečný zápis do nášho bloku
  sh.getRange(2, outFirstCol, out.length, outColCount).setValues(out);
}

// ====== CLEANUP & RECOMPUTE FOR BANK IMPORT SHEETS ======

/** Odstráni duplicitné hlavičky s príponou ".1" v prvom riadku. */
function __dropDuplicateHeaderColumns__(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  // ideme sprava -> zľava, aby sa indexy nemenili
  for (var c = lastCol; c >= 1; c--) {
    var name = header[c - 1];
    if (name && typeof name === 'string' && /\.1$/.test(name)) {
      sh.deleteColumn(c);
    }
  }
}

/** Zaistí súvislý blok výstupných hlavičiek na konci listu. */
function __ensureOutBlockStrict__(sh, headers) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) lastCol = 1;
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // pokus nájsť už existujúci súvislý blok
  outer:
  for (var c = 1; c <= lastCol - headers.length + 1; c++) {
    for (var j = 0; j < headers.length; j++) {
      if (headerRow[c - 1 + j] !== headers[j]) continue outer;
    }
    return { col: c, count: headers.length };
  }

  // inak dopíšeme úplne na koniec v danom poradí
  var startCol = lastCol + 1;
  sh.getRange(1, startCol, 1, headers.length).setValues([headers]);
  return { col: startCol, count: headers.length };
}

/** Parse peňazí z rôznych text formátov; čísla vracia priamo. */
function __parseMoney__(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  var s = String(v).trim().replace(/\u00A0/g, ' ').replace(/ /g, '');
  if (/\d+\.\d{3},\d{2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  else if (/\d+,\d{3}\.\d{2}$/.test(s)) s = s.replace(/,/g, '');
  else s = s.replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? null : n;
}

/** Vyberie sieť sadzieb: najprv z globálu VAT_RATES, inak 23/20/10 %. */
function __vatRates__() {
  return (typeof VAT_RATES !== 'undefined' && VAT_RATES && VAT_RATES.length)
    ? VAT_RATES
    : [0.23, 0.20, 0.10];
}

/** Skúsi odvodiť NET z GROSS tak, aby GROSS ≈ NET*(1+r) (tolerancia 0.01). Ak nič, vráti GROSS. */
function __guessNetFromGross__(gross) {
  if (gross == null) return '';
  var rates = __vatRates__();
  for (var i = 0; i < rates.length; i++) {
    var r = rates[i];
    var net = Math.round((gross / (1 + r)) * 100) / 100;
    var back = Math.round((net * (1 + r)) * 100) / 100;
    if (Math.abs(back - gross) <= 0.01) return net;
  }
  return gross;
}

/** Nájde index stĺpca so sumou podľa známych názvov; inak -1. */
function __findAmountIndex__(header) {
  var candidates = ['Amount', 'Suma', 'AmountWithVat', 'AmountGross', 'Amount EUR', 'Suma EUR'];
  for (var i = 0; i < candidates.length; i++) {
    var idx = header.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

/**
 * Oprav list importu: odstráň duplikátne hlavičky ".1", vytvor súvislý výstupný blok
 * a dopočítaj NORMALIZED_AMOUNT (odhad netto z GROSS).
 */
function repairAndRecomputeBankImport(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet nie je dostupný: " + sheetName);

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  // 1) odstráň duplicitné .1 stĺpce
  __dropDuplicateHeaderColumns__(sh);

  // 2) refresh hlavičky po mazaní
  lastCol = sh.getLastColumn();
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // 3) pevný výstupný blok
  var OUT_HEADERS = [
    'MATCH_STATUS','MATCH_RULE','MATCH_SCORE',
    'NORMALIZED_AMOUNT','EXTRACTED_ICO','EXTRACTED_VAR','EXTRACTED_SPEC','EXTRACTED_KS'
  ];
  var outLoc = __ensureOutBlockStrict__(sh, OUT_HEADERS);

  // 4) načítaj dáta a dopočítaj
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  if (!rows.length) return;

  var amtIdx = __findAmountIndex__(header);

  var out = new Array(rows.length);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var gross = null;

    if (amtIdx >= 0) gross = __parseMoney__(row[amtIdx]);
    else {
      // fallback: prvá číselná hodnota v riadku
      for (var c = 0; c < row.length; c++) {
        var tryN = __parseMoney__(row[c]);
        if (tryN != null) { gross = tryN; break; }
      }
    }

    var net = __guessNetFromGross__(gross);

    out[i] = ['UNMATCHED', '', 0, (net !== '' ? net : ''), '', '', '', ''];
  }

  // 5) bezpečný zápis do výstupného bloku (súvislý rozsah)
  sh.getRange(2, outLoc.col, out.length, outLoc.count).setValues(out);
}

// shortcut: spusti na aktívnom importe (poslednom vytvorenom liste BankImport_*)
function repairLastBankImport() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var latest = null, latestTime = 0;
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var n = sh.getName();
    if (n.indexOf('BankImport_') === 0) {
      var t = sh.getLastUpdated ? sh.getLastUpdated().getTime() : 0;
      if (t >= latestTime) { latestTime = t; latest = sh; }
    }
  }
  if (!latest) throw new Error('Nenašiel som žiadny list začínajúci "BankImport_".');
  repairAndRecomputeBankImport(latest.getName());
}

// ====== PATCH: robustné rozpoznanie Amount a znamienko podľa CreditDebit ======

function __findAndComputeAmount__(header, row) {
  const hLower = header.map(h => String(h || '').toLowerCase());
  const amountIdx = hLower.indexOf('amount');
  const creditIdx = hLower.indexOf('credit');
  const debitIdx = hLower.indexOf('debit');
  const cdIdx = hLower.indexOf('creditdebit');

  // 1️⃣ urči hodnotu
  let raw = null;
  if (amountIdx >= 0) raw = __parseMoney__(row[amountIdx]);
  else if (creditIdx >= 0 && debitIdx >= 0) {
    const credit = __parseMoney__(row[creditIdx]) || 0;
    const debit = __parseMoney__(row[debitIdx]) || 0;
    raw = credit - debit;
  }

  // 2️⃣ aplikuj znamienko podľa CreditDebit
  if (cdIdx >= 0) {
    const cdVal = String(row[cdIdx]).trim().toUpperCase();
    if (cdVal === 'DBIT' && raw > 0) raw = -raw;
  }

  return raw;
}

// prepíš časť repairAndRecomputeBankImport() (override)
const __oldRepair__ = repairAndRecomputeBankImport;
repairAndRecomputeBankImport = function(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet nie je dostupný: " + sheetName);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  __dropDuplicateHeaderColumns__(sh);

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const OUT_HEADERS = [
    'MATCH_STATUS','MATCH_RULE','MATCH_SCORE',
    'NORMALIZED_AMOUNT','EXTRACTED_ICO','EXTRACTED_VAR','EXTRACTED_SPEC','EXTRACTED_KS'
  ];
  const outLoc = __ensureOutBlockStrict__(sh, OUT_HEADERS);

  const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const out = new Array(rows.length);

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const gross = __findAndComputeAmount__(header, row);
    const net = __guessNetFromGross__(gross);
    out[i] = ['UNMATCHED', '', 0, net, '', '', '', ''];
  }

  sh.getRange(2, outLoc.col, out.length, outLoc.count).setValues(out);
};
