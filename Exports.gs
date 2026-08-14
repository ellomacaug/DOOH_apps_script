function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Export")
    .addItem("Выгрузить расчет со всеми столбцами", "exportCurrentSheetPrettyFull")
    .addItem("Выгрузить расчет готовый", "exportCurrentSheetPrettySmall")
    .addItem("Таблица для Email Sender DOOH", "exportCurrentSheetEmailSender")
    .addToUi();
}

function showFinishHTML(newSS){
//  ---------- Export URL ----------
  const url =
    "https://docs.google.com/spreadsheets/d/" +
    newSS.getId() +
    "/export?format=xlsx";

  Logger.log(url);
  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  const html2 = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, Helvetica, sans-serif; font-size: 20px; padding:15px">
      <p>
        <a href="https://docs.google.com/spreadsheets/d/${newSS.getId()}"
          target="_blank">
          Открыть GoogleSheet
        </a>
      </p>

      <p>
        <a href="${url}" target="_blank">
          Скачать файл (.xlsx)
        </a>
      </p>
    </div>
  `)
  SpreadsheetApp.getUi().showModalDialog(html2, "Готово");
}

function exportCurrentSheetEmailSender(){
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const newSS = SpreadsheetApp.create('contacts'); 
  const resultSheet = sourceSheet.copyTo(newSS);
  resultSheet.setName('contacts'); 
  const defaultSheet = newSS.getSheets().find(
    s => s.getSheetId() !== resultSheet.getSheetId()
  );
  newSS.deleteSheet(defaultSheet);
  removeCheckedRows(resultSheet);
  resultSheet.getDataRange()
    .clearFormat()
    .breakApart()
    .setBorder(false, false, false, false, false, false);

  let lastRow = resultSheet.getLastRow();
  const emails = resultSheet.getRange(2, 45, lastRow - 1, 1).getValues();

  for (let i = emails.length - 1; i >= 0; i--) {
    if (!emails[i][0]) {
      resultSheet.deleteRow(i + 2);
    }
  }

  const keep = [3, 4, 8, 9, 10, 12, 19, 20, 45, 46];

  for (let col = resultSheet.getLastColumn(); col >= 1; col--) {
    if (!keep.includes(col)) {
      resultSheet.deleteColumn(col);
    }
  }

  resultSheet.getRange(1, 1, 1, 10).setValues([[
    "city",
    "mall",
    "rim",
    "num",
    "size",
    "link",
    "min",
    "sec",
    "email",
    "name"
  ]]);

  showFinishHTML(newSS)
}

function exportCurrentSheetPrettySmall(){
  exportCurrentSheetPretty(false);
}
function exportCurrentSheetPrettyFull(){
  exportCurrentSheetPretty(true);
}

function exportCurrentSheetPretty(isFull){
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rA = sourceSheet.getRange("CA1").getValue() || "";
  const brand = sourceSheet.getRange("CB1").getValue() || "";
  const fromDate = sourceSheet.getRange("CC1").getValue() || "";
  const toDate = sourceSheet.getRange("CD1").getValue() || "";

  const cdValues = sourceSheet.getRange("C:D").getValues();
  let lastCDRow = 0;
  for (let i = cdValues.length - 1; i >= 0; i--) {
    if (cdValues[i][0] !== "" && cdValues[i][1] !== "") {
      lastCDRow = i + 1;
      break;
    }}
  const dataRows = Math.max(lastCDRow, 0);

  const newSS = SpreadsheetApp.create( Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') ); 
  const resultSheet = sourceSheet.copyTo(newSS);
  let maxRows = resultSheet.getMaxRows();

  if (maxRows > dataRows) {
    resultSheet.deleteRows(
      dataRows + 1,
      maxRows - dataRows
    );
    maxRows = resultSheet.getMaxRows();
  }
  resultSheet.setName('Malls_' + formatPeriod(fromDate, toDate)); 
  const defaultSheet = newSS.getSheets().find(
    s => s.getSheetId() !== resultSheet.getSheetId()
  );
  newSS.deleteSheet(defaultSheet);

  removeCheckedRows(resultSheet);
  let maxCols = resultSheet.getMaxColumns();

  // Save ЗАКУПОЧНАЯ И ФОРМУЛА into CA:CB
  resultSheet.insertColumnsAfter(maxCols, 3);
  resultSheet
    .getRange(1, 23, maxRows, 2) // W:X
    .copyTo(
      resultSheet.getRange(1, maxCols+1, maxRows, 2),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );
  // Remove W:X so all existing column numbers remain unchanged
  resultSheet.deleteColumns(23, 2);
  maxCols = maxCols-2;
  let lastRow = resultSheet.getLastRow();

  // Number malls
  const cities = resultSheet.getRange(2, 3, lastRow - 1, 1).getValues();
  const malls  = resultSheet.getRange(2, 4, lastRow - 1, 1).getValues();
  resultSheet.getRange(2, 2, lastRow - 1, 1).breakApart();

  let numberOfMalls = 1;
  let i = 0;
  while (i < cities.length) {
    const key = getMallKey(cities[i][0], malls[i][0]);
    let j = i + 1;
    while (
      j < cities.length &&
      getMallKey(cities[j][0], malls[j][0]) === key
    ) {
      j++;
    }
    const range = resultSheet.getRange(i + 2, 2, j - i, 1);
    range.clearContent();
    if (j - i > 1) {
      range.merge();
    }

    range.setValue(numberOfMalls++);
    i = j;
  }
  numberOfMalls--;
  newSS.setName( rA.toUpperCase() + '_' + brand.toUpperCase() + '_DOOH_MALLS (' + numberOfMalls + ')_' + formatPeriod(fromDate, toDate) + '_v1.0_');

  // Delete exsessive cols
  let endCol = maxCols;
  if (maxCols > 32) {
    endCol = isFull ? 55 : 32;
    if (maxCols > endCol) {
      resultSheet.deleteColumns(endCol + 1, maxCols - endCol);
    }
  }
  resultSheet.insertColumnsAfter(endCol, 20);
  maxCols = resultSheet.getMaxColumns();

  resultSheet.getRange(1, 1, maxRows, maxCols)
    .setBorder(
      true, true, true, true, true, true,
      "white",
      SpreadsheetApp.BorderStyle.SOLID
    );

  if (dataRows) {
  // 2. Alignment
  const ttRange = resultSheet.getRange(1, 11, dataRows, 3);
  const ttValues = ttRange.getValues();
  ttRange.setHorizontalAlignments(
    ttValues.map(row =>
      row.map(value => value === "-" ? "center" : "left")
    )
  );
  const cityRange = resultSheet.getRange(2, 3, dataRows, 1);
  cityRange.setHorizontalAlignments(
    cityRange.getValues().map(row =>
      row.map(value => String(value || "").length < 8 ? "center" : "left")
    )
  );
  // center
  [2,4,6,7,9,14,15,16,17,18,19,20,21,22,25,28,29,30,31].forEach(col => {
      resultSheet.getRange(2, col, dataRows, 1).setHorizontalAlignment("center");
  });

  // left
  [1,5,8,10,32].forEach(col => {
      resultSheet.getRange(2, col, dataRows, 1).setHorizontalAlignment("left");
  });

  // right
  [23,24,26,27].forEach(col => {
      resultSheet.getRange(2, col, dataRows, 1).setHorizontalAlignment("right");
  });

  // 3. Column widths
  resultSheet.setColumnWidth(1,6);
  resultSheet.setColumnWidth(2,25);
  resultSheet.setColumnWidth(3,70);
  resultSheet.setColumnWidth(4,225);
  resultSheet.setColumnWidth(8,225);
  resultSheet.setColumnWidth(11,110);
  resultSheet.setColumnWidth(12,160);
  resultSheet.setColumnWidth(15,150);
  resultSheet.setColumnWidth(16,70);
  resultSheet.setColumnWidth(17,85);
  resultSheet.setColumnWidth(18,90);
  resultSheet.setColumnWidth(20,110);
  resultSheet.setColumnWidth(21,85);
  resultSheet.setColumnWidth(22,85);
  resultSheet.setColumnWidth(23,200);
  resultSheet.setColumnWidth(24,200);
  if (rA.length > 6) {resultSheet.setColumnWidth(25,160);}
  resultSheet.setColumnWidth(26,200);
  resultSheet.setColumnWidth(27,210);
  resultSheet.setColumnWidth(28,148);

  resultSheet.autoResizeColumn(31);
  let currentWidth = resultSheet.getColumnWidth(31);
  if (currentWidth < 110) { resultSheet.setColumnWidth(31, 110); }

  resultSheet.autoResizeColumn(32);
  currentWidth = resultSheet.getColumnWidth(32);
  if (currentWidth < 120) { resultSheet.setColumnWidth(32, 120); } else if (currentWidth > 250){resultSheet.setColumnWidth(32, 250);}

  // 5. Font sizes
  resultSheet.getRange(1,1,1,resultSheet.getLastColumn()).setFontFamily("Palatino Linotype").setFontSize(11);
  resultSheet
    .getRange(2,1,dataRows,resultSheet.getLastColumn())
    .setFontSize(10)
    .setFontFamily("Palatino Linotype");

  const linkRange = resultSheet.getRange(2,11,dataRows-1,3);
  linkRange.setFontColors(
    linkRange.getValues().map(row =>
      row.map(value => value === "-" ? "black" : "blue")
    )
  );

  // 6. Bold columns
  [
    [[4,9,14,15,16,18,19,20,21,22,25,26,27,28,29,30,31], "bold"],
    [[1,2,3,5,6,7,8,10,11,12,13,17,23,24,32], "normal"]
  ].forEach(([cols, weight]) =>
    cols.forEach(col =>
      resultSheet.getRange(2, col, dataRows, 1).setFontWeight(weight)
    )
  );

  const rAColor1 = RA_CONFIG[rA.toUpperCase()]?.color1 || "black";
  const rAColor2 = RA_CONFIG[rA.toUpperCase()]?.color2 || "white";

  // табличка сверху
  resultSheet.insertRows(1, 6);
  resultSheet.getRange(1, 1, 6, resultSheet.getLastColumn()+20).setBackground("white");
  resultSheet.getRange(7, endCol+1, maxRows+200, resultSheet.getLastColumn()+20).setBackground("white");
  resultSheet.getRange(lastRow+1, endCol+1, maxRows+200-lastRow, resultSheet.getLastColumn()+20)
    .setBorder(
      true, true, true, true, true, true,
      "white",
      SpreadsheetApp.BorderStyle.SOLID
    );

  resultSheet.getRange("C2").setValue("РА:");
  resultSheet.getRange("E2").setValue(rA);

  resultSheet.getRange("C3").setValue("Клиент:");
  resultSheet.getRange("E3").setValue(brand);

  resultSheet.getRange("C4").setValue("Период размещения:");
  resultSheet.getRange("E4").setValue(formatPeriod(fromDate, toDate));

  resultSheet.getRange("C5").setValue("Дата расчета:");
  resultSheet.getRange("E5").setValue(new Date()).setNumberFormat("dd.MM.yyyy");
  
  resultSheet.getRange("C2:D5").mergeAcross();
  resultSheet.getRange("E2:F5").mergeAcross();

  const infoRange = resultSheet.getRange("C2:F5");
  infoRange
    .setFontFamily("Palatino Linotype")
    .setFontSize(12)
    .setFontColor("black")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  resultSheet.getRange("C2:D5").setHorizontalAlignment("left");
  resultSheet.getRange("E2:F2").setFontColor(rAColor1)

  resultSheet.getRange(1, 1, 6, maxCols)
    .setBorder(
      true, true, true, true, true, true,
      "white",
      SpreadsheetApp.BorderStyle.SOLID
    );

  infoRange.setBorder(
    true, true, true, true, true, true,
    "black",
    SpreadsheetApp.BorderStyle.SOLID
  );
  lastRow = resultSheet.getLastRow();
  maxRows = resultSheet.getMaxRows();
  const rangesToColor = [
    resultSheet.getRange(1, 1, 6, maxCols).getA1Notation(),
    resultSheet.getRange(8, 1, maxRows - 7, isFull ? 32 : maxCols).getA1Notation(),
    "A7",
    ...(maxCols > endCol
      ? [resultSheet.getRange(7, endCol + 1, 1, maxCols - endCol).getA1Notation()]
      : [])];

  resultSheet.getRangeList(rangesToColor).setBackground("white");
  const footerRow = lastRow + 2;
  
  // футэр
  resultSheet
    .getRange("D" + footerRow)
    .setValue("Адаптация контента под ТТ экранов – бесплатно.");
  resultSheet
    .getRange("D" + (lastRow + 3))
    .setValue("Фото и видео отчёт – бесплатно.");
  resultSheet
    .getRange("C" + (lastRow + 5))
    .setValue("СOMMENTS / КОММЕНТАРИИ:");
  resultSheet
    .getRange("C" + (lastRow + 6))
    .setValue("* ТЦ оставляет за собой право отказать в размещении по причине несогласования содержания ролика / макета РИМ");
  resultSheet
    .getRange("C" + (lastRow + 7))
    .setValue("* Стоимость действительна для указанного периода и может быть скорректирована при изменении параметров АП ");
  resultSheet
    .getRange("C" + (lastRow + 8))
    .setValue("* Все стороны находятся в свободной продаже");

  resultSheet.getRange("C" + footerRow + ":AA" + (lastRow + 8))
    .setFontFamily("Palatino Linotype")
    .setFontSize(12)
    .setFontColor("black")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  resultSheet.getRange("C" + (lastRow + 5))
    .setFontColor("red")
    .setFontLine("underline");

  resultSheet.getRange("V" + footerRow + ":AA" + footerRow)
    .setHorizontalAlignment("right")
    .setFontSize(10);


  maxRows = resultSheet.getMaxRows();
  resultSheet.getRange(lastRow+2, 1, maxRows, maxCols)
    .setBorder(
      true, true, true, true, true, true,
      "white",
      SpreadsheetApp.BorderStyle.SOLID
    );
  resultSheet.getRange("U" + (footerRow+1) + ":AA" + (footerRow+1))
    .setBorder(
      true, false, false, false, false, false,
      "black",
      SpreadsheetApp.BorderStyle.SOLID);

  resultSheet.getRange(`W${footerRow}`)
    .setBackground("white")
    .setFontWeight("normal")
    .setFontColor("#0033cc");

  resultSheet.getRangeList([`X${footerRow}`,`AA${footerRow}`])
    .setBackground("yellow")
    .setFontWeight("normal")
    .setFontColor("#0033cc");

  resultSheet.getRangeList([`Y${footerRow}`,`Z${footerRow}`])
    .setBackground("white")
    .setFontWeight("bold")
    .setFontColor("black");

  const mallC = String(numberOfMalls);
  const text1 = `Итого_Прайс (${mallC} ТЦ):`;
  const color1 = "green";
  const richText1 = SpreadsheetApp.newRichTextValue()
  .setText(text1)
  .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor("black").build()) 
  .setTextStyle(text1.indexOf(mallC), text1.indexOf(mallC) + mallC.length + 3, SpreadsheetApp.newTextStyle().setForegroundColor(color1).build())
  .build();
  resultSheet
    .getRange("V" + footerRow)
    .setRichTextValue(richText1);

  resultSheet
    .getRange("W" + footerRow)
    .setFormula(`=SUM(W8:W${lastRow})`);
  resultSheet
    .getRange("X" + footerRow)
    .setFormula(`=SUM(X8:X${lastRow})`);

  const text2 = `С учётом скидок для ${rA}:`;
  const richText2 = SpreadsheetApp.newRichTextValue()
  .setText(text2)
  .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor("black").build()) 
  .setTextStyle(text2.indexOf(rA), text2.indexOf(rA) + rA.length, SpreadsheetApp.newTextStyle().setForegroundColor(rAColor1).build())
  .build();
  resultSheet
    .getRange("Z" + footerRow)
    .setRichTextValue(richText2);

  resultSheet
    .getRange("AA" + footerRow)
    .setFormula(`=SUM(AA8:AA${lastRow})`);

  // 1. Borders
    resultSheet.getRange(8, 2, lastRow-7, endCol-1)
      .setBorder(
        true, true, true, true, true, true,
        "black",
        SpreadsheetApp.BorderStyle.SOLID
      );

  // 4. Row heights
  for (let r = 1; r <= 6; r++) {
    resultSheet.setRowHeight(r,26);
  }
  resultSheet.setRowHeight(7,100);
  for (let r = 8; r <= maxRows; r++) {
    resultSheet.setRowHeight(r,22);
  }

  // заголовок таблицы
  resultSheet.getRange("B7:AF7")
    .setFontFamily("Palatino Linotype")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(
      true, true, true, true, true, true,
      "black",
      SpreadsheetApp.BorderStyle.SOLID
    );
  
  
  const text3 = `Скидка для ${rA}, %`;
  const richText3 = SpreadsheetApp.newRichTextValue()
  .setText(text3)
  .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("white").build()) 
  .setTextStyle(text3.indexOf(rA), text3.indexOf(rA) + rA.length, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(rAColor2).build())
  .build();
  resultSheet
    .getRange("Y7")
    .setRichTextValue(richText3);

  const text4 = `Стоимость размещения с учётом скидок для ${rA}, руб., за ПЕРИОД, без НДС.`;
  const richText4 = SpreadsheetApp.newRichTextValue()
  .setText(text4)
  .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("white").build()) 
  .setTextStyle(text4.indexOf(rA)+ rA.length + 8, text4.indexOf(rA) + rA.length + 17,SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("yellow").build())
  .setTextStyle(text4.indexOf(rA), text4.indexOf(rA) + rA.length, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(rAColor2).build())
  .build();
  resultSheet
    .getRange("AA7")
    .setRichTextValue(richText4);
  
  const richText5 = SpreadsheetApp.newRichTextValue()
  .setText(`Стоимость размещения по ПРАЙСУ руб., за 1 месяц, без НДС.`)
  .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("white").build()) 
  .setTextStyle(37, 47,SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("yellow").build())
  .build();
  resultSheet
    .getRange("W7")
    .setRichTextValue(richText5);
  
  const richText6 = SpreadsheetApp.newRichTextValue()
  .setText('Стоимость размещения по ПРАЙСУ руб., за ПЕРИОД, без НДС.')
  .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("white").build()) 
  .setTextStyle(37, 46,SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("yellow").build())
  .build();
  resultSheet
    .getRange("X7")
    .setRichTextValue(richText6);

  }



  // Restore ЗАКУПОЧНАЯ И ФОРМУЛА from CA:CB
  resultSheet.insertColumnsBefore(23, 2);
  resultSheet
    .getRange(7, endCol+23, maxRows+7, 2) // CA:CB
    .copyTo(
      resultSheet.getRange(7, 23, maxRows+7, 2), // W:X
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );
    resultSheet.deleteColumns(endCol+23, 3);
    resultSheet.getRange(1, 23, lastRow, 2).setBackground("yellow").setFontColor("black");
  
    resultSheet.getRange("W7").setValue("Закупочная");
    resultSheet.getRange("X7").setValue("Формула");
    

  resultSheet.setFrozenColumns(4);

  showFinishHTML(newSS);
}
