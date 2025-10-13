function setupBudgetSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let budgetSheet = sheet.getSheetByName("Rozpočet");
  if (budgetSheet) sheet.deleteSheet(budgetSheet);
  budgetSheet = sheet.insertSheet("Rozpočet");

  const headers = [
    "Mesiac",
    "Kategória",
    "Projekt",
    "Plánovaný výdavok (€)",
    "Opakovanie", // Mesačne / Jednorazovo
    "Posledná splátka"
  ];
  const sampleData = [
    ["2025-01", "Marketing", "Kampaň A", 1000, "Mesačne", "2025-06"],
    ["2025-01", "IT", "Infra", 2000, "Jednorazovo", ""],
    ["2025-02", "Marketing", "Kampaň A", 1500, "Mesačne", "2025-05"]
  ];

  budgetSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#e8eaed");
  budgetSheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
}
