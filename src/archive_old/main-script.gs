/**
 * Hlavný Google Apps Script pre firemné výdavky a dashboard.
 * Autor: GPT, na mieru pre Barfer.sk
 */

// ================= CONFIG =================
const config = {
  sheetName: "Form Responses 1",
  dashboardSheet: "📊 Dashboard",
  auditSheet: "📘 Audit log",
  budgetSheet: "Rozpočet",
  trzbySheet: "Tržby",
  driveRootId: "1qYSW1DHku6i-4bzGgJpp2MXn0pDa2pCG",
  columnMap: {
    date: "Dátum",
    amount: "Suma (€) bez DPH",
    category: "Kategória",
    project: "Projekt",
    expenseName: "Názov výdavku",
    note: "Poznámka",
    file: "Dokument k výdavku"
  },
  categoryColors: {
    "Materiál": "#a5d6a7",
    "Tovar (dodavatelia)": "#80cbc4",
    "Personál": "#4dd0e1",
    "Preprava objednávok(PHM,Ext.Kurier)": "#ffe082",
    "Marketing": "#90caf9",
    "Réžia(najom,servis,telefon,ucto,programator..)": "#ce93d8",
    "Leasingy a úvery": "#bcaaa4"
  }
};
// ==========================================

// ---- Pomocné formátovanie / indexy ----
function safeParseNumber(value) {
  if (typeof value === "string") value = value.replace(",", ".").replace(/\s/g, "");
  return parseFloat(value) || 0;
}
function formatCurrency(val) {
  const num = typeof val === "string"
    ? parseFloat(val.replace(",", ".").replace(/[^0-9.-]/g, ""))
    : parseFloat(val);
  if (isNaN(num)) return "—";
  return `${num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, " ").replace(".", ",")} €`;
}
function colIndexMap(headers) {
  return {
    date: headers.indexOf("Dátum"),
    expenseName: headers.indexOf("Názov výdavku"),
    amount: headers.indexOf("Suma (€) bez DPH"),
    category: headers.indexOf("Kategória"),
    project: headers.indexOf("Projekt"),
    note: headers.indexOf("Poznámka"),
    file: headers.indexOf("Dokument k výdavku"),
    status: headers.indexOf("Status"),
    supplierColumns: {
      "Materiál": headers.indexOf("Materiál"),
      "Tovar (dodavatelia)": headers.indexOf("Tovar (dodavatelia)"),
      "Personál": headers.indexOf("Personál"),
      "Preprava objednávok(PHM,Ext.Kurier)": headers.indexOf("Preprava objednávok(PHM,Ext.Kurier)"),
      "Marketing": headers.indexOf("Marketing"),
      "Réžia(najom,servis,telefon,ucto,programator..)": headers.indexOf("Réžia(najom,servis,telefon,ucto,programator..)"),
      "Leasingy a úvery": headers.indexOf("Leasingy a úvery")
    }
  };
}
function extractSupplier(row, category, idx) {
  const i = idx.supplierColumns[category];
  return i !== undefined && i >= 0 ? row[i] : "";
}

// ---- Drive pomôcky ----
function getFileIdFromUrl(url) {
  const match = url?.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}
function getOrCreateFolder(name, parent) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

// ---- Personál (zoznam aktívnych) ----
function getActiveEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Personál");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameIndex = headers.indexOf("Meno");
  const activeIndex = headers.indexOf("Aktívny");
  if (nameIndex === -1 || activeIndex === -1) return [];
  return data.slice(1)
    .filter(r => String(r[activeIndex]).toLowerCase() === "áno")
    .map(r => r[nameIndex]);
}

// ---- Form submit handler ----
function onFormSubmit(e) {
  const data = e.namedValues;
  const datum = new Date(data[config.columnMap.date][0]);
  const projekt = data[config.columnMap.project][0];
  const fileUrl = data[config.columnMap.file][0];
  const fileId = getFileIdFromUrl(fileUrl);
  const uuid = Utilities.getUuid().slice(0, 8);

  const rok = datum.getFullYear();
  const mesiac = String(datum.getMonth() + 1).padStart(2, "0");
  const root = DriveApp.getFolderById(config.driveRootId);
  const yearFolder = getOrCreateFolder(String(rok), root);
  const monthFolder = getOrCreateFolder(mesiac, yearFolder);
  const projectFolder = getOrCreateFolder(projekt, monthFolder);

  if (fileId) {
    const file = DriveApp.getFileById(fileId);
    const originalName = file.getName();
    let newName = `${rok}-${mesiac}-${uuid}_${originalName}`;
    let n = 2;
    while (projectFolder.getFilesByName(newName).hasNext()) {
      newName = `${rok}-${mesiac}-${uuid}_${originalName} (${n++})`;
    }
    file.setName(newName);
    projectFolder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}
  }

  generateStyledMonthlySheets();
  updateDashboardWithBudgetComparison();
}

// ======== GENEROVANIE MESAČNÝCH PREHĽADOV =========
function generateStyledMonthlySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(config.sheetName);
  if (!source) return;
  const data = source.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const idx = colIndexMap(headers);
  const rows = data.slice(1);

  const monthsFromTrzby = getAllMonthsFromTrzbySheet();
  const months = {};

  rows.forEach(r => {
    const d = new Date(r[idx.date]);
    if (isNaN(d)) return;
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    if (!months[key]) months[key] = [];
    months[key].push(r);
  });
  monthsFromTrzby.forEach(m => { if (!months[m]) months[m] = []; });

  const activeEmployees = getActiveEmployees();

  Object.entries(months).forEach(([monthKey, monthRows]) => {
    const sheet = ss.getSheetByName(monthKey) || ss.insertSheet(monthKey);
    sheet.clear();

    // Hlavička
    const headerRow = ["Dátum", "Názov výdavku", "Suma (€) bez DPH", "Kategória", "Projekt", "Dodávateľ", "Poznámka", "Doklad (odkaz)"];
    sheet.getRange(1, 1, 1, headerRow.length)
      .setValues([headerRow]).setFontWeight("bold").setBackground("#f1f3f4");

    let rowIdx = 2;

    // Rozdelenie podľa kategórie
    const byCategory = {};
    monthRows.forEach(r => {
      const cat = r[idx.category] || "Nezaradené";
      if (!byCategory[cat]) byCategory[cat] = [];
      byCategory[cat].push(r);
    });
    if (!byCategory["Personál"]) byCategory["Personál"] = [];

    // Sekcie
    for (const [cat, catRows] of Object.entries(byCategory)) {
      // Hlavička kategórie + „vizuálne tlačidlo“
      sheet.getRange(rowIdx, 1, 1, 8)
        .merge().setValue(cat).setFontWeight("bold").setFontColor("white").setBackground("#434343");
      sheet.getRange(rowIdx, 8).setValue("▶ Zobraziť/skryť").setFontStyle("italic").setFontColor("#cccccc");
      rowIdx++;

      let sum = 0;
      const addedEmployees = new Set();
      const isPersonal = cat === "Personál";

      catRows.forEach(r => {
        const supplier = extractSupplier(r, cat, idx);
        const row = [
          r[idx.date],
          r[idx.expenseName],
          formatCurrency(safeParseNumber(r[idx.amount])),
          r[idx.category],
          r[idx.project],
          supplier,
          r[idx.note],
          r[idx.file] ? `=HYPERLINK("${r[idx.file]}", "Doklad")` : ""
        ];
        sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
        sheet.getRange(rowIdx, 3).setHorizontalAlignment("right");
        sheet.getRange(rowIdx, 4).setBackground(config.categoryColors[cat] || "#e0e0e0").setFontColor("black");
        sum += safeParseNumber(r[idx.amount]);
        if (isPersonal && supplier) addedEmployees.add(supplier);
        rowIdx++;
      });

      // Doplníme chýbajúcich aktívnych zam.
      if (isPersonal) {
        activeEmployees.forEach(name => {
          if (!addedEmployees.has(name)) {
            const row = ["", "", formatCurrency(0), cat, "", name, "", ""];
            sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
            sheet.getRange(rowIdx, 3).setHorizontalAlignment("right").setBackground("#fde0dc");
            sheet.getRange(rowIdx, 4).setBackground(config.categoryColors[cat] || "#e0e0e0");
            sheet.getRange(rowIdx, 8).setValue(formatCurrency(0));
            rowIdx++;
          }
        });
      }

      // Sumár kategórie
      sheet.getRange(rowIdx, 1).setValue("Sumár");
      sheet.getRange(rowIdx, 3).setValue(formatCurrency(sum));
      sheet.getRange(rowIdx, 8).setValue(formatCurrency(sum));
      sheet.getRange(rowIdx, 1, 1, 8).setBackground("#d0f0c0").setFontWeight("bold");
      sheet.getRange(rowIdx, 3).setHorizontalAlignment("right");
      sheet.getRange(rowIdx, 8).setHorizontalAlignment("right");
      rowIdx += 2;
    }

    // Doplníme personál (bez ohľadu na to, či boli záznamy)
    renderPersonnelInReports(sheet, monthKey);

    // Blok súhrnu a formátovanie farieb
    insertTrzbySummaryBlock(sheet, monthKey, monthRows, headers);

    sheet.setFrozenRows(1);
    sheet.setHiddenGridlines(true);
  });
}

// ======== Súhrn za mesiac (pod tabuľkou) =========
function insertTrzbySummaryBlock(sheet, monthKey, dataRows, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trzbySheet = ss.getSheetByName(config.trzbySheet);

  const projectColIndex = headers.indexOf("Projekt");
  const categoryColIndex = headers.indexOf("Kategória");
  const amountColIndex = headers.indexOf("Suma (€) bez DPH");

  let sumAll = 0;
  let sumNoPrivate = 0;
  let sumNoTovar = 0;

  dataRows.forEach(r => {
    const amount = safeParseNumber(r[amountColIndex]);
    sumAll += amount;

    const project = r[projectColIndex];
    const category = r[categoryColIndex];

    if (project !== "Beňo") sumNoPrivate += amount;               // vyraď súkromné
    if (project === "Barfer.sk" && category !== "Tovar (dodavatelia)") {
      sumNoTovar += amount;                                       // náklad bez tovaru
    }
  });

  const [y, m] = monthKey.split("-").map(Number);
  const daysInMonth = new Date(y, m, 0).getDate();
  const dailyCost = sumNoTovar / daysInMonth;

  // Tržby / objednávky / profit
  let trzba = "Chýbajú dáta", orders = "Chýbajú dáta", profit = "Chýbajú dáta";
  if (trzbySheet) {
    const trzbyData = trzbySheet.getDataRange().getValues();
    const H = trzbyData[0];
    const idxMonth = H.indexOf("Mesiac");
    const idxTrzba = H.indexOf("Tržba celkom bez DPH");
    const idxOrders = H.indexOf("Počet objednávok");
    const idxProfit = H.indexOf("Profit z predaja tovaru");
    const row = trzbyData.find((r, i) => i > 0 && r[idxMonth] === monthKey);
    if (row) {
      trzba = safeParseNumber(row[idxTrzba]) || 0;
      orders = row[idxOrders] || 0;
      profit = safeParseNumber(row[idxProfit]) || 0;
    }
  }

  // Správny smer: Profit – Náklady bez tovaru
  const diffProfit = typeof profit === "number" ? (profit - sumNoTovar) : "Chýbajú dáta";
  const diffTrzba  = typeof trzba  === "number" ? (trzba  - sumNoPrivate) : "Chýbajú dáta";

  const fmt = v => (typeof v === "number" ? formatCurrency(v) : v);

  const tableData = [
    ["Náklady celkom bez DPH barfer.sk", fmt(sumAll)],
    ["Náklady celkom bez DPH barfer.sk bez súkromných", fmt(sumNoPrivate)],
    ["Náklady bez tovaru bez DPH barfer.sk", fmt(sumNoTovar)],
    ["Denný náklad bez tovaru", fmt(dailyCost)],
    [""],
    ["Tržba celkom bez DPH", fmt(trzba)],
    ["Počet objednávok", orders],
    ["Profit z predaja tovaru", fmt(profit)],
    [""],
    ["Náklady vs profit", fmt(diffProfit)],
    ["Tržba – komplet náklady bez súkromných", fmt(diffTrzba)]
  ];

  const startRow = sheet.getLastRow() + 2;
  const normalized = tableData.map(r => (r.length === 2 ? r : [r[0], ""]));
  sheet.getRange(startRow, 1, normalized.length, 2).setValues(normalized);

  // zvýraznenie hlavičiek riadkov bez hodnoty
  normalized.forEach((row, i) => {
    if (!row[1]) sheet.getRange(startRow + i, 1, 1, 2).merge().setFontWeight("bold");
  });

  const labelCells = sheet.getRange(startRow, 1, normalized.length, 1).setFontWeight("bold");
  const valueCells = sheet.getRange(startRow, 2, normalized.length, 1).setHorizontalAlignment("right").setFontWeight("bold");

  // Farby podľa dohody (náklady ružové s čiernym písmom, ostatné zelené/červené podľa znamienka)
  normalized.forEach((row, i) => {
    const r = startRow + i;
    const label = row[0] || "";
    const val = row[1];
    const cell = sheet.getRange(r, 2);

    if (typeof val === "string" && val.includes("Chýbajú")) {
      cell.setBackground("#fdd").setFontColor("black");
      return;
    }
    if (typeof val === "string" && val.endsWith("€")) {
      const number = parseFloat(val.replace(/\s/g, "").replace(",", "."));
      if (
        label.includes("Náklady celkom") ||
        label.includes("Náklady bez tovaru") ||
        label.includes("Denný náklad")
      ) {
        cell.setBackground("#fde0dc").setFontColor("black"); // svetločervené
      } else {
        const color = number >= 0 ? "#4caf50" : "#e53935";
        cell.setFontColor(color).setBackground("#e8f5e9");   // zelené pozadie pre metriky/rozdiely
      }
    }
  });
}

// ======== Dashboard (základ) =========
function updateDashboardWithBudgetComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(config.dashboardSheet) || ss.insertSheet(config.dashboardSheet);
  dash.clear();

  // zoberieme najnovší mesiac podľa listov „YYYY-MM“
  const monthSheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => /^\d{4}-\d{2}$/.test(n))
    .sort();
  const latest = monthSheets[monthSheets.length - 1];
  if (!latest) {
    dash.getRange(1,1).setValue("Zatiaľ nie sú k dispozícii mesačné dáta.");
    return;
  }

  // prečítaj blok súhrnu z konca mesačného listu (2 stĺpce)
  const ms = ss.getSheetByName(latest);
  const lastRow = ms.getLastRow();
  // heuristika: posledných ~20 riadkov bude súhrn
  const rng = ms.getRange(Math.max(1, lastRow - 30), 1, 30, 2).getValues()
    .filter(r => r[0]); // len ne-prázdne labely

  dash.getRange(1,1).setValue(`📅 Súhrn pre ${latest}`).setFontWeight("bold");
  if (rng.length) {
    dash.getRange(3,1, rng.length, 2).setValues(rng);
    dash.getRange(3,1, rng.length, 1).setFontWeight("bold");
    dash.getRange(1,1,1,2).merge();
  }

  dash.autoResizeColumns(1, 2);
  dash.setFrozenRows(2);
}

// ======== Ďalšie pomôcky =========
function getAllMonthsFromTrzbySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.trzbySheet);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const monthIndex = headers.indexOf("Mesiac");
  if (monthIndex === -1) return [];
  const months = new Set();
  for (let i = 1; i < data.length; i++) {
    const val = data[i][monthIndex];
    if (typeof val === "string" && /^\d{4}-\d{2}$/.test(val)) months.add(val);
  }
  return Array.from(months).sort();
}

function renderPersonnelInReports(sheet, monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const personSheet = ss.getSheetByName("Personál");
  if (!personSheet) return;

  const data = personSheet.getDataRange().getValues();
  const headers = data[0];
  const idx = { name: headers.indexOf("Meno"), active: headers.indexOf("Aktívny") };
  const activePeople = data.slice(1)
    .filter(r => r[idx.active] === true || String(r[idx.active]).toUpperCase() === "TRUE" || String(r[idx.active]).toLowerCase() === "áno")
    .map(r => r[idx.name]);

  const lastRow = sheet.getLastRow();
  const categoryRows = sheet.getRange(1, 1, lastRow, 1).getValues();
  let startRow = -1;
  for (let i = 0; i < categoryRows.length; i++) {
    const val = categoryRows[i][0];
    if (val && typeof val === "string" && val.trim().toLowerCase() === "personál") {
      startRow = i + 1; break;
    }
  }
  if (startRow === -1) return;

  // existujúce mená v časti Personál
  const existingNames = [];
  for (let i = startRow + 1; i <= Math.min(lastRow, startRow + 80); i++) {
    const categoryVal = sheet.getRange(i, 4).getValue();
    const supplier = sheet.getRange(i, 6).getValue();
    if (categoryVal === "Personál" && supplier) existingNames.push(supplier);
  }

  // doplň chýbajúcich
  let insertRow = startRow + 1;
  activePeople.forEach(name => {
    if (!existingNames.includes(name)) {
      sheet.insertRows(insertRow, 1);
      sheet.getRange(insertRow, 4).setValue("Personál");
      sheet.getRange(insertRow, 6).setValue(name);
      sheet.getRange(insertRow, 3).setValue(0).setFontColor("black").setBackground("#fde0dc");
      sheet.getRange(insertRow, 8).setValue(0);
      insertRow++;
    }
  });
}

// ======== UI: menu + ručné skrývanie kategórií =========
function toggleCategoryRows() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Zadaj názov kategórie, ktorú chceš zobraziť/skryť:");
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const categoryName = response.getResponseText().trim();
  if (!categoryName) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(1, 1, lastRow, 1).getValues();

  let startRow = -1, endRow = -1;
  for (let i = 0; i < range.length; i++) {
    if (typeof range[i][0] === "string" && range[i][0].trim().toLowerCase() === categoryName.toLowerCase()) {
      startRow = i + 1; break;
    }
  }
  if (startRow === -1) { ui.alert(`Kategória "${categoryName}" nebola nájdená.`); return; }

  for (let i = startRow + 1; i <= lastRow; i++) {
    const val = sheet.getRange(i, 1).getValue();
    if (typeof val === "string" && val.trim().toLowerCase() === "sumár") { endRow = i - 1; break; }
  }
  if (endRow === -1) { ui.alert(`Kategória "${categoryName}" nemá ukončenie (sumár).`); return; }

  const currentlyHidden = sheet.isRowHiddenByUser(startRow + 1);
  if (currentlyHidden) sheet.showRows(startRow + 1, endRow - startRow);
  else sheet.hideRows(startRow + 1, endRow - startRow);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💼 Výdavky")
    .addItem("🔄 Obnoviť mesačné prehľady", "generateStyledMonthlySheets")
    .addItem("📊 Obnoviť Dashboard", "updateDashboardWithBudgetComparison")
    .addItem("♻️ Spustiť úplnú aktualizáciu", "runFullUpdate")
    .addItem("📂 Zobraziť/Skryť kategóriu", "toggleCategoryRows")
    .addToUi();
}
function runFullUpdate() {
  generateStyledMonthlySheets();
  updateDashboardWithBudgetComparison();
}
