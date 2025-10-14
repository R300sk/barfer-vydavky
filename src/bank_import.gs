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
