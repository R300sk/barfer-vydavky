function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Výdavky')
    .addItem('Aktualizovať mesačný výkaz', 'buildSummary')
    .addItem('Aktualizovať Dashboard', 'updateDashboard')
    .addSeparator()
    .addItem('💳 Importovať camt.053 z Drive ID', 'menuImportCamt')
    .addItem('💳 Importovať CSV z Drive ID', 'menuImportCSV')
    .addItem('📥 Importovať posledný súbor z Inbox priečinka', 'menuImportFromInbox')
    .addItem('🔗 Spárovať posledný import', 'menuMatchLastImport')
    .addToUi();
}

/** Prompt na camt.053 XML import podľa fileId z Drive. */
function menuImportCamt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Zadaj Google Drive fileId (camt.053 XML):");
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const fileId = resp.getResponseText().trim();
  if (!fileId) return ui.alert("Nevyplnené fileId.");

  const sheetName = importBankCamtFromDrive(fileId);
  try {
    safeMatchAllBankImportRows(sheetName);
    ui.alert("Hotovo", "Naimportované a spárované do listu:\n" + sheetName, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Import prebehol, párovanie zlyhalo:\n" + e, ui.ButtonSet.OK);
  }
}

/** Prompt na CSV import podľa fileId z Drive. */
function menuImportCSV() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Zadaj Google Drive fileId (CSV):");
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const fileId = resp.getResponseText().trim();
  if (!fileId) return ui.alert("Nevyplnené fileId.");

  const sheetName = importBankCSVFromDrive(fileId);
  try {
    safeMatchAllBankImportRows(sheetName);
    ui.alert("Hotovo", "Naimportované a spárované do listu:\n" + sheetName, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Import prebehol, párovanie zlyhalo:\n" + e, ui.ButtonSet.OK);
  }
}

/** Import najnovšieho XML/CSV z Drive „Inbox“ priečinka (CONFIG.BANK.INBOX_FOLDER_ID). */
function menuImportFromInbox() {
  const ui = SpreadsheetApp.getUi();
  const folderId = (typeof CONFIG !== 'undefined' && CONFIG.BANK) ? CONFIG.BANK.INBOX_FOLDER_ID : "";
  if (!folderId) return ui.alert("Inbox priečinok nie je nastavený v configu (CONFIG.BANK.INBOX_FOLDER_ID).");

  const folder = DriveApp.getFolderById(folderId);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const n = f.getName();
    if (/\.(xml|csv)$/i.test(n)) files.push({ f, n, t: f.getLastUpdated().getTime() });
  }
  if (!files.length) return ui.alert("V Inbox priečinku som nenašiel žiadne .xml / .csv súbory.");

  files.sort((a,b) => b.t - a.t);
  const chosen = files[0];
  const fileId = chosen.f.getId();

  let sheetName = "";
  if (/\.xml$/i.test(chosen.n)) sheetName = importBankCamtFromDrive(fileId);
  else sheetName = importBankCSVFromDrive(fileId);

  try {
    safeMatchAllBankImportRows(sheetName);
    ui.alert("Hotovo", "Naimportované a spárované:\n" + chosen.n + "\n→ " + sheetName, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Import prebehol, párovanie zlyhalo:\n" + e, ui.ButtonSet.OK);
  }
}

/** Znova spustí párovanie na „poslednom“ BankImport_ liste (podľa poradia sheetov). */
function menuMatchLastImport() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const pref = (typeof BANK_IMPORT_CONFIG !== 'undefined' && BANK_IMPORT_CONFIG.BANK_SHEET_PREFIX) ? BANK_IMPORT_CONFIG.BANK_SHEET_PREFIX : "BankImport_";
  const importSheets = ss.getSheets().filter(sh => sh.getName().startsWith(pref));
  if (!importSheets.length) return ui.alert("Nenašiel som žiadny sheet začínajúci '" + pref + "'.");

  const last = importSheets.sort((a,b) => b.getIndex() - a.getIndex())[0];
  try {
    safeMatchAllBankImportRows(last.getName());
    ui.alert("Hotovo", "Spárované znova: " + last.getName(), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Párovanie zlyhalo:\n" + e, ui.ButtonSet.OK);
  }
}
