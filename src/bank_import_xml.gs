/** bank_import_xml.gs
 * Import Tatra banka camt.053 XML (camt.053.001.02) a napojenie na párovanie.
 * Vytvorí sheet "BankImport_<shortName>" s normalizovanými stĺpcami.
 *
 * Použitie:
 *  importBankCamtFromDrive(fileId)          // načíta XML z Drive a importuje
 *  importBankCamtXmlText(xmlText, shortName) // import z textu XML
 *  matchAllBankImportRows("BankImport_<...>")// existujúca funkcia z bank_import.gs
 */

function importBankCamtFromDrive(fileId) {
  const file = DriveApp.getFileById(fileId);
  const xmlText = file.getBlob().getDataAsString("UTF-8");
  const shortName = file.getName().replace(/\.[^/.]+$/, "");
  return importBankCamtXmlText(xmlText, shortName);
}

function importBankCamtXmlText(xmlText, shortName) {
  const module = "bank_import_xml";
  logInfo(module, `Importujem camt.053 XML: ${shortName}`);

  const rows = parseCamt053_(xmlText);
  if (!rows || rows.length < 2) {
    throw new Error("XML neobsahuje žiadne transakcie (alebo neznámy formát).");
  }

  const sheetName = `${BANK_IMPORT_CONFIG.BANK_SHEET_PREFIX}${shortName}`;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(sheetName);

  sheet.getRange(1, 1, 1, rows[0].length).setValues([rows[0]]);
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, rows[0].length).setValues(rows.slice(1));
  }

  const extraHeaders = [
    "MATCH_TARGET_ROW",
    "MATCH_SCORE",
    "MATCH_REASON",
    "NORMALIZED_AMOUNT",
    "NORMALIZED_DATE",
    "EXTRACTED_VARIABLE_SYMBOL",
    "EXTRACTED_PAYER"
  ];
  sheet.getRange(1, rows[0].length + 1, 1, extraHeaders.length).setValues([extraHeaders]);

  logInfo(module, `camt.053 import hotový → ${sheetName} (${rows.length - 1} riadkov).`);
  return sheetName;
}

function parseCamt053_(xmlText) {
  const doc = XmlService.parse(xmlText);
  const ns = XmlService.getNamespace("urn:iso:std:iso:20022:tech:xsd:camt.053.001.02");
  const stmt = doc.getRootElement().getChild("BkToCstmrStmt", ns).getChild("Stmt", ns);
  const entries = stmt.getChildren("Ntry", ns);
  const headers = [
    "BookDate","ValueDate","Amount","Currency","CreditDebit","PayerName",
    "CreditorName","EndToEndId","AcctSvcrRef","Ustrd","NtryRef"
  ];
  const rows = [headers];

  entries.forEach((ntry) => {
    const amtEl = ntry.getChild("Amt", ns);
    const amt = amtEl ? Number(amtEl.getText().replace(",", ".")) : 0;
    const ccy = amtEl ? (amtEl.getAttribute("Ccy") ? amtEl.getAttribute("Ccy").getValue() : "EUR") : "EUR";
    const cd = getText_(ntry, ["CdtDbtInd"], ns);
    const bookDate = getText_(ntry, ["BookgDt","Dt"], ns) || "";
    const valDate = getText_(ntry, ["ValDt","Dt"], ns) || "";
    const ntryRef = getText_(ntry, ["NtryRef"], ns) || "";

    const ntryDtls = ntry.getChild("NtryDtls", ns);
    let endToEndId="",acctSvcrRef="",ustrd="",payerName="",creditorName="";
    if (ntryDtls) {
      const txDtls = firstChildDeep_(ntryDtls, ["TxDtls"], ns);
      if (txDtls) {
        endToEndId = getText_(txDtls, ["Refs","EndToEndId"], ns) || "";
        acctSvcrRef = getText_(txDtls, ["Refs","AcctSvcrRef"], ns) || "";
        ustrd = getText_(txDtls, ["RmtInf","Ustrd"], ns) || "";
        const dbtr = firstChildDeep_(txDtls, ["RltdPties","Dbtr"], ns);
        const cdtr = firstChildDeep_(txDtls, ["RltdPties","Cdtr"], ns);
        payerName = dbtr ? getText_(dbtr, ["Nm"], ns) : "";
        creditorName = cdtr ? getText_(cdtr, ["Nm"], ns) : "";
      }
    }

    rows.push([
      bookDate,valDate,amt,ccy,cd,payerName,creditorName,
      endToEndId,acctSvcrRef,ustrd,ntryRef
    ]);
  });

  return rows;
}

function getText_(el, pathArr, ns) {
  if (!el) return "";
  let cur = el;
  for (const p of pathArr) {
    cur = cur.getChild(p, ns);
    if (!cur) return "";
  }
  return cur.getText() || "";
}

function firstChildDeep_(el, pathArr, ns) {
  let cur = el;
  for (const p of pathArr) {
    cur = cur.getChild(p, ns);
    if (!cur) return null;
  }
  return cur;
}
