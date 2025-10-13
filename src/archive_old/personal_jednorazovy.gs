function createPersonalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Personál");

  if (!sheet) {
    sheet = ss.insertSheet("Personál");
  } else {
    sheet.clear();
  }

  const months = ["Január", "Február", "Marec", "Apríl", "Máj", "Jún", "Júl", "August", "September", "Október", "November", "December"];
  const headers = ["Meno", "Hodinová sadzba (€)", "Príplatok (€)", "Plán hodín", "Aktívny"].concat(months);
  const columnCount = headers.length;

  // Zápis hlavičky
  sheet.getRange(1, 1, 1, columnCount).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);

  // Formátovanie
  sheet.getRange("A:A").setFontWeight("bold").setHorizontalAlignment("left");
  sheet.getRange("B:D").setNumberFormat("#,##0.00").setHorizontalAlignment("right");
  sheet.getRange("E:E").setHorizontalAlignment("center");
  sheet.getRange("F:Q").setNumberFormat("0").setHorizontalAlignment("center");

  // Vzorové dáta (voliteľné, môžeš odstrániť)
  sheet.getRange(2, 1, 1, columnCount).setValues([[
    "Ján Novák", 8.5, 0.5, 160, "áno", "", "", "", "", "", "", "", "", "", "", "", ""
  ]]);
}
