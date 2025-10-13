
function updateDashboardWithBudgetComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = ss.getSheetByName("Rozpočet");
  if (!budgetSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Rozpočet' neexistuje.");
    return;
  }

  const dashboardSheet = ss.getSheetByName("📊 Dashboard") || ss.insertSheet("📊 Dashboard");
  dashboardSheet.clear();
  dashboardSheet.setFrozenRows(1);

  const budgetData = budgetSheet.getDataRange().getValues().slice(1);
  const monthSheets = ss.getSheets().filter(s => s.getName().match(/^2025-\d{2}$/));

  const actuals = {};

  monthSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues().slice(1);
    data.forEach(row => {
      const mesiac = sheet.getName();
      const kategoria = row[3] || "Nezaradené";
      const projekt = row[4] || "Nezaradené";
      const suma = parseFloat(row[2]) || 0;

      if (!actuals[mesiac]) actuals[mesiac] = {};
      if (!actuals[mesiac][kategoria]) actuals[mesiac][kategoria] = {};
      if (!actuals[mesiac][kategoria][projekt]) actuals[mesiac][kategoria][projekt] = 0;

      actuals[mesiac][kategoria][projekt] += suma;
    });
  });

  const merged = [];

  budgetData.forEach(row => {
    const mesiac = row[0];
    const kategoria = row[1];
    const projekt = row[2];
    const plan = parseFloat(row[3]) || 0;
    const skutocnost = actuals?.[mesiac]?.[kategoria]?.[projekt] || 0;
    const rozdiel = skutocnost - plan;

    merged.push([
      mesiac,
      kategoria,
      projekt,
      formatCurrency(plan),
      formatCurrency(skutocnost),
      formatCurrency(rozdiel)
    ]);
  });

  const header = ["Mesiac", "Kategória", "Projekt", "Plánovaný výdavok (€)", "Skutočný výdavok (€)", "Rozdiel (€)"];
  dashboardSheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold").setBackground("#e8eaed");
  dashboardSheet.getRange(2, 1, merged.length, header.length).setValues(merged);

  const summaryByMonth = {};
  merged.forEach(r => {
    const mesiac = r[0];
    const real = parseCurrency(r[4]);
    if (!summaryByMonth[mesiac]) summaryByMonth[mesiac] = 0;
    summaryByMonth[mesiac] += real;
  });

  const monthSummaryData = Object.entries(summaryByMonth).sort().map(([m, s]) => [m, s]);
  const monthStartRow = merged.length + 4;
  dashboardSheet.getRange(monthStartRow, 1).setValue("📊 Výdavky podľa mesiacov").setFontWeight("bold");
  dashboardSheet.getRange(monthStartRow + 1, 1, monthSummaryData.length, 2).setValues(monthSummaryData);

  const chart1 = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dashboardSheet.getRange(monthStartRow + 1, 1, monthSummaryData.length, 2))
    .setPosition(monthStartRow + 1, 4, 0, 0)
    .setOption('title', 'Výdavky podľa mesiacov')
    .setOption('colors', ['#4285F4'])
    .build();
  dashboardSheet.insertChart(chart1);

  const summaryByProject = {};
  merged.forEach(r => {
    const projekt = r[2];
    const real = parseCurrency(r[4]);
    if (!summaryByProject[projekt]) summaryByProject[projekt] = 0;
    summaryByProject[projekt] += real;
  });

  const projectSummaryData = Object.entries(summaryByProject).sort().map(([p, s]) => [p, s]);
  const projectStartRow = monthStartRow + monthSummaryData.length + 10;
  dashboardSheet.getRange(projectStartRow, 1).setValue("📈 Výdavky podľa projektov").setFontWeight("bold");
  dashboardSheet.getRange(projectStartRow + 1, 1, projectSummaryData.length, 2).setValues(projectSummaryData);

  const chart2 = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dashboardSheet.getRange(projectStartRow + 1, 1, projectSummaryData.length, 2))
    .setPosition(projectStartRow + 1, 4, 0, 0)
    .setOption('title', 'Výdavky podľa projektov')
    .setOption('colors', ['#0F9D58'])
    .build();
  dashboardSheet.insertChart(chart2);
}

function formatCurrency(num) {
  if (isNaN(num)) return "";
  return num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, " ").replace(".", ",") + " €";
}

function parseCurrency(formatted) {
  return parseFloat(formatted.replace(" €", "").replace(/ /g, "").replace(",", "."));
}
