function ensureSummary2Sheet() {
  const SS = SpreadsheetApp.getActive();
  const NAME = 'summary 2';
  const EXPENSE_SHEET = 'Form Responses 1';
  const SALES_SHEET   = 'Tržby';

  let sh = SS.getSheetByName(NAME);
  if (!sh) sh = SS.insertSheet(NAME);
  sh.clear({ contentsOnly: true });

  sh.getRange('A1:C1').setValues([['Metrika', 'Hodnota', 'Vzorec']]).setFontWeight('bold');

  const rows = [
    ['Náklady celkom bez DPH barfer.sk',
     `=LET(S, '${EXPENSE_SHEET}'!A:Z,
        SUMIFS(
          INDEX(S,0, MATCH("Suma (€) bez DPH", INDEX(S,1,), 0)),
          INDEX(S,0, MATCH("Projekt", INDEX(S,1,), 0)), "barfer.sk"
        ))`],
    ['Náklady celkom bez DPH barfer.sk bez súkromných',
     `=LET(S, '${EXPENSE_SHEET}'!A:Z,
        SUMIFS(
          INDEX(S,0, MATCH("Suma (€) bez DPH", INDEX(S,1,), 0)),
          INDEX(S,0, MATCH("Projekt", INDEX(S,1,), 0)), "barfer.sk",
          INDEX(S,0, MATCH("Kategória", INDEX(S,1,), 0)), "<>Súkromné"
        ))`],
    ['Náklady bez tovaru bez DPH barfer.sk',
     `=LET(S, '${EXPENSE_SHEET}'!A:Z,
        SUMIFS(
          INDEX(S,0, MATCH("Suma (€) bez DPH", INDEX(S,1,), 0)),
          INDEX(S,0, MATCH("Projekt", INDEX(S,1,), 0)), "barfer.sk",
          INDEX(S,0, MATCH("Kategória", INDEX(S,1,), 0)), "<>Tovar (dodavatelia)"
        ))`],
    ['Denný náklad bez tovaru',
     `=LET(S, '${EXPENSE_SHEET}'!A:Z,
        SUMIFS(
          INDEX(S,0, MATCH("Suma (€) bez DPH", INDEX(S,1,), 0)),
          INDEX(S,0, MATCH("Projekt", INDEX(S,1,), 0)), "barfer.sk",
          INDEX(S,0, MATCH("Kategória", INDEX(S,1,), 0)), "<>Tovar (dodavatelia)"
        )/
        COUNTUNIQUE(
          FILTER(
            INDEX(S,0, MATCH("Dátum", INDEX(S,1,), 0)),
            INDEX(S,0, MATCH("Projekt", INDEX(S,1,), 0))="barfer.sk",
            INDEX(S,0, MATCH("Kategória", INDEX(S,1,), 0))<>"Tovar (dodavatelia)"
          )
        ))`],
    ['', ''],
    ['Tržba celkom bez DPH',
     `=LET(S, '${SALES_SHEET}'!A:Z, SUM( INDEX(S,0, MATCH("Tržba bez DPH", INDEX(S,1,), 0)) ))`],
    ['Počet objednávok',
     `=LET(S, '${SALES_SHEET}'!A:Z, SUM( INDEX(S,0, MATCH("Počet objednávok", INDEX(S,1,), 0)) ))`],
     // ak nemáš tento stĺpec a máš "Objednávka ID", zmeň na: COUNTA(... MATCH("Objednávka ID"...))
    ['Profit z predaja tovaru',
     `=LET(S, '${SALES_SHEET}'!A:Z, SUM( INDEX(S,0, MATCH("Profit z predaja tovaru", INDEX(S,1,), 0)) ))`],
    ['', ''],
    ['Náklady vs profit', '=B9 - B4'],
    ['Tržba – komplet náklady bez súkromných', '=B7 - B3'],
  ];

  const startRow = 2;
  sh.getRange('A1:C1').setValues([['Metrika', 'Hodnota', 'Vzorec']]).setFontWeight('bold');

  // popisky
  sh.getRange(startRow, 1, rows.length, 1).setValues(rows.map(r => [r[0]]));

  // hodnoty (formuly)
  rows.forEach((r, i) => {
    const f = (r[1] || '').trim();
    const row = startRow + i;
    const cell = sh.getRange(row, 2);
    if (f) cell.setFormula(f); else cell.setValue('');
    sh.getRange(row, 3).setFormula(f ? `=FORMULATEXT(B${row})` : ''); // stĺpec „Vzorec“
  });

  // formáty
  sh.setColumnWidths(1, 3, 260);
  sh.getRange('A:A').setWrap(true);
  sh.getRange('C:C').setWrap(true);
  sh.setFrozenRows(1);

  // € formát len pre peňažné riadky (nie počet objednávok)
  const moneyRows = [2,3,4,5,7,9,11,12]; // čísla riadkov v tabulke (A2..)
  moneyRows.forEach(n => sh.getRange(n, 2).setNumberFormat('€ #,##0.00'));
  sh.getRange(startRow, 2, rows.length, 1).setHorizontalAlignment('right');
}

