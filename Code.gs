onOpen();

function exportFilteredFileAdvanced() {
  exportFilteredFile("advanced");
}

function exportFilteredFileList() {
  exportFilteredFile("list");
}

function exportFilteredFileTender() {
  exportFilteredFile("tender");
}

function exportFilteredFile(mode) {
  const cfg = CONFIG[mode];
  const filterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);

  // ---------- Read source ----------
  const sourceRange = sourceSheet.getDataRange();

  const sourceData = sourceRange.getValues();
  const sourceFormulas = sourceRange.getFormulasR1C1();

  const headers = sourceData[0];

  const cityCol = headers.indexOf('Город');
  const typeCol = headers.indexOf('DOOH / Digital Indoor');
  const isDayedCol = headers.indexOf('Делают пересчёт по дням');
  const notesCol = headers.indexOf("№");
  const mallCol = headers.indexOf("Торговый центр");
  const prohibitedBrandsCol = headers.indexOf("Кого ТОЧНО не размещают");
  const shopCols = [
    headers.indexOf("Магазин бытовой техники"),
    headers.indexOf("Кинотеатр"),
    headers.indexOf("Гипермаркет")
  ];

  if (cityCol === -1 || typeCol === -1) {
    throw new Error('В супер файле что-то не так (Нет столбцов Город или DOOH/Digital Indoor).');
  }

  if (shopCols.some(col => col === -1)) {
    throw new Error('В супер файле что-то не так (Нет столбцов с Гипермаркетом или Кинотеатром или МБТ или "Кого ТОЧНО не размещают").');
  }
  const durationCol = headers.indexOf('Продолжительность размещения');
  const periodCol = headers.indexOf('Период размещения');
  const daysCol = headers.indexOf('Кол-во дней');

  if (durationCol === -1 || periodCol === -1 || daysCol === -1) {
    throw new Error("В супер файле что-то не так (не хватает нужных столбов).");
  }

  // ---------- Read filter table ----------

  const cityFilterMap =
    mode === "advanced"
      ? buildAdvancedFilter(filterSheet, cfg)
      : mode === "tender"
        ? buildTenderFilter(filterSheet, cfg)
        : buildListFilter(filterSheet, cfg);


  const fromDate = filterSheet.getRange(cfg.fromCell).getValue();
  const toDate = filterSheet.getRange(cfg.toCell).getValue();

  if (!(fromDate instanceof Date) || !(toDate instanceof Date)) {
    throw new Error("В ячейках G9 и G10 должны быть даты формата 26.06.2026");
  }
  const brand = filterSheet.getRange(cfg.brandCell).getDisplayValue().trim();

  const rA = filterSheet.getRange(cfg.raCell).getDisplayValue().trim();

  const shops = filterSheet.getRange(cfg.shopCell || "").getDisplayValue().split(",")
    .map(s => normalize(s))
    .filter(Boolean);

  const isZaprosRaw = String(filterSheet.getRange(cfg.zaprosCell).getDisplayValue())
    .trim()
    .toLowerCase();

  const isZapros = ["да", "yes", "y", "true", "1"].includes(isZaprosRaw);

  // ---------- Filter ----------
  const result = [headers];
  const formulaRows = [];
  const matchedRows = [];

  for (let i = 1; i < sourceData.length; i++) {

    const row = sourceData[i];
    const type = String(row[typeCol]).trim();
    const isDayedValue = String(row[isDayedCol]).trim();
    const notes = String(row[notesCol] || "");
    const normalizedNotes = normalize(notes);
    const mall = String(row[mallCol]).trim();
    const rowCityRaw = String(row[cityCol]).trim();
    const city = canonicalCity(rowCityRaw);

    const rowBrands = String(row[prohibitedBrandsCol] || "")
      .split(",")
      .map(s => normalize(s))
      .filter(Boolean);

    const isNOTProhibitedBrand = !brand || !rowBrands.includes(normalize(brand));

    const rowShops = shopCols
      .flatMap(col =>
        String(row[col] || "")
          .split(",")
          .map(s => normalize(s))
          .filter(Boolean)
      );

    const shopMatch =
      shops.length === 0 ||
      shops.some(shop => rowShops.includes(shop));

    const isMO =
      rowCityRaw === "МО" ||
      rowCityRaw.startsWith("МО (");

    let filterKey =
      cityFilterMap.has(city)
        ? city
        : isMO && cityFilterMap.has("__MO_ALL__")
          ? "__MO_ALL__"
          : null;

    if (!filterKey) {
      continue;
    }
    let allowedTypes;
    let allowedDayed;
    if (mode === "tender") {

      const filters = cityFilterMap.get(filterKey);

      const exceptionMall = getExceptionMallForRow(mall);

      const filter = filters.find(f => {
        // Exception mall:
        // match ONLY the same mall, regardless of source city.
        if (exceptionMall) {
          return normalizeMall(f.mall) === normalizeMall(mall);
        }

        // Normal Tender row:
        // city + mall must both match.
        return getMallKey(city, f.mall) === getMallKey(city, mall);
      });

      if (!filter) {
        continue;
      }

      allowedTypes = filter.types;
      allowedDayed = filter.dayed;

    } else {

      const filter = cityFilterMap.get(filterKey);

      allowedTypes = filter.types;
      allowedDayed = filter.dayed;
    }

    const rowTypes = type
      .split(',')
      .map(t => t.trim());
    const rowDayed = isDayedValue
      .split(',')
      .map(v => v.trim());

    const typeMatch =
      allowedTypes.size === 0 ||
      rowTypes.some(t => allowedTypes.has(t));

    const dayedMatch =
      allowedDayed.size === 0 ||
      rowDayed.some(v => allowedDayed.has(v));
    const newRow = [...row];
    const newFormulaRow = [...sourceFormulas[i]];

    newRow[durationCol] = formatDuration(fromDate, toDate);
    newRow[periodCol] = formatPeriod(fromDate, toDate);
    newRow[daysCol] = daysInclusive(fromDate, toDate);

    // Skip "under request only"
    if (isZapros && normalizedNotes.includes(normalize("Продаём только под запрос"))) {
      continue;
    }

    // Skip forbidden RAs
    const aliases = RA_CONFIG[rA.toUpperCase()]?.aliases || [];

    const forbiddenForRA =
      normalizedNotes.includes("непредлаг") &&
      aliases.some(alias => normalizedNotes.includes(normalize(alias)));

    if (forbiddenForRA) {
      continue;
    }

    // empty type list means "all types"
    if (typeMatch && dayedMatch && shopMatch && isNOTProhibitedBrand) {
      result.push(newRow);
      formulaRows.push(newFormulaRow);
      matchedRows.push(i + 1);
    }
  }
  
  
  // ---------- Dates ----------
  function daysInclusive(fromDate, toDate) {
    const msPerDay = 24 * 60 * 60 * 1000;
    return Math.floor((toDate - fromDate) / msPerDay) + 1;
  }

  function monthDifferenceInclusive(fromDate, toDate) {
    return (
      (toDate.getFullYear() - fromDate.getFullYear()) * 12 +
      (toDate.getMonth() - fromDate.getMonth()) +
      1
    );
  }

  function formatDuration(fromDate, toDate) {

    const months = monthDifferenceInclusive(fromDate, toDate);

    // Exactly whole months (e.g. 01.03–31.03, 01.03–30.04)
    const isWholeMonths =
        fromDate.getDate() === 1 &&
        toDate.getDate() === new Date(
          toDate.getFullYear(),
          toDate.getMonth() + 1,
          0
        ).getDate();

    if (isWholeMonths) {
      return months === 1 ? "1 месяц" :
            months >= 2 && months <= 4 ? months + " месяца" :
            months + " месяцев";
    }

    const days = daysInclusive(fromDate, toDate);

    const mod10 = days % 10;
    const mod100 = days % 100;

    let suffix;

    if (mod10 === 1 && mod100 !== 11) {
      suffix = "день";
    } else if (mod10 >= 2 && mod10 <= 4 && !(mod100 >= 12 && mod100 <= 14)) {
      suffix = "дня";
    } else {
      suffix = "дней";
    }

    return `${days} ${suffix}`;
  }

  const rowsToExport = result
    .slice(1)
    .map((row, i) => ({
      row,
      formulas: formulaRows[i],
      sourceRow: matchedRows[i]
    }));

  rowsToExport.sort((a, b) => {
    const ca = citySortInfo(a.row[cityCol]);
    const cb = citySortInfo(b.row[cityCol]);

    // 1. Million-city priority
    if (ca.priority !== cb.priority) {
      return ca.priority - cb.priority;
    }

    // 2. All МО entries: ignore subcity, sort only by mall
    if (ca.priority === CITY_PRIORITY.get(normalize("МО"))) {
      return normalizeMall(a.row[mallCol])
        .localeCompare(normalizeMall(b.row[mallCol]), "ru");
    }

    // 3. Other cities: city then mall
    const byCity = ca.name.localeCompare(cb.name, "ru");
    if (byCity) return byCity;

    return normalizeMall(a.row[mallCol])
      .localeCompare(normalizeMall(b.row[mallCol]), "ru");
  });

  result.splice(1, result.length - 1, ...rowsToExport.map(x => x.row));
  formulaRows.splice(0, formulaRows.length, ...rowsToExport.map(x => x.formulas));
  matchedRows.splice(0, matchedRows.length, ...rowsToExport.map(x => x.sourceRow));

  // ---------- Create temporary spreadsheet ----------
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetName = rA.toUpperCase() + '_' + brand.toUpperCase()+ '_' + formatPeriod(fromDate, toDate) + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd HH:mm');

  // delete previous sheet with same name if exists
  const old = ss.getSheetByName(sheetName);
  // if (old) sheetName = rA.toUpperCase() + '_' + brand.toUpperCase()+ '_' + formatPeriod(fromDate, toDate) + '_v2.0_';
  if (old) ss.deleteSheet(old);

  const oldT = ss.getSheetByName('Template');
  if (oldT) ss.deleteSheet(oldT);

  const resultSheet = ss.insertSheet(sheetName);
  const templateSheet = sourceSheet.copyTo(ss);
  templateSheet.setName("Template");
  templateSheet.hideSheet();

  resultSheet
    .getRange(1, 1, result.length, result[0].length)
    .setValues(result);

  let lastRow = resultSheet.getLastRow();
  const rows = lastRow - 1;

  if (rows > 0) {
    const agAjValues = result
      .slice(1)
      .map(row => row.slice(32, 37));

    const agAjFormulas = formulaRows.map(row => row.slice(32, 37));

    const range = resultSheet.getRange(2, 33, rows, 5);

    range.setValues(agAjValues);

    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < 5; j++) {
        const formula = agAjFormulas[i][j];

        if (formula) {
          resultSheet
            .getRange(i + 2, 33 + j)
            .setFormulaR1C1(formula);
        }
      }
    }
  }


  const hours = resultSheet.getRange(2, 6, rows, 1).getValues().map(([schedule]) => getWorkingHours(schedule));

  [
    [17, (r, i) => `=60/S${r}*${hours[i].toString().replace('.', ',')}*I${r}*U${r}`],
    [18, (r) => `=Q${r}*P${r}`],
    [22, (r) => `=(T${r}*U${r})/(S${r}*60)`],
    [24, (r) => `=W${r}`],
    [27, (r) => `=X${r}*(1-Y${r})`],
  ].forEach(([col, formula]) => {
    resultSheet.getRange(2, col, rows).setFormulas(
      Array.from({ length: rows }, (_, i) => [formula(i + 2, i)])
    );
  });

  // статика
  resultSheet.getRange(2, 7, rows, 1).getValues().forEach(([value], i) => {
    if (value === "Static") {
      resultSheet
        .getRange(i + 2, 17, 1, 6) // Q:V
        .merge()
        .setValue("Статичное размещение");
    }});

  //resultSheet.getRange(1, 1, resultSheet.getLastRow()+100, resultSheet.getLastColumn()+20).setBackground("white");
  templateSheet
      .getRange(1,1,1,result[0].length)
      .copyFormatToRange(resultSheet,1,result[0].length,1,1);

  for (let i = 0; i < matchedRows.length; i++) {

      const sourceRow = matchedRows[i];

      const targetRow = i + 2; // row 2 because row 1 is header

      templateSheet
          .getRange(sourceRow, 1, 1, result[0].length)
          .copyFormatToRange(resultSheet,1,result[0].length,targetRow,targetRow);
  }
  ss.deleteSheet(templateSheet);

  // Форматирование столбца Прайса
  resultSheet.getRange(2, 24, rows, 1).copyFormatToRange(resultSheet, 23, 23, 2, rows);


  // Insert ЗАКУПОЧНАЯ И ФОРМУЛА
  const k = RA_CONFIG[rA.toUpperCase()]?.k || "1,46";
  resultSheet.insertColumnsBefore(23, 2);

  resultSheet.getRange("W1").setValue("НЕ УДАЛЯТЬ СТОЛБЕЦ\nЗакупочная");
  resultSheet.getRange("X1").setValue("НЕ УДАЛЯТЬ СТОЛБЕЦ\nФормула");

  if (lastRow > 1) {
    resultSheet.getRange(2, 24, rows, 1).setFormulas(
      Array.from(
        { length: rows },
        (_, i) => [`=if(not(W${i + 2}>0,9); "не число"; IF(W${i + 2}<12000; 29000;  IF(W${i + 2}<22000; 39000; IF(W${i + 2}<30000; 49000; ROUNDDOWN(W${i + 2}*${k}/10000;0)*10000+9000))))`] // ФОРМУЛА
      )
    );
  }
  
  resultSheet.getRange(1, 23, lastRow+20, 2).setBackground("yellow").setFontColor("black");

  resultSheet.getRange("CA1").setValue(rA);
  resultSheet.getRange("CB1").setValue(brand);
  resultSheet.getRange("CC1").setValue(fromDate);
  resultSheet.getRange("CD1").setValue(toDate);
  resultSheet.hideColumns(79, 4);

  // ---------- Checkbox column ----------
  resultSheet.getRange(1, 1, lastRow).clearContent().clearDataValidations().clearFormat();
  resultSheet.getRange("A1").setValue("Поставленная галочка\n=\nСтрока удалится при экспорте").setBackground("#356854").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Roboto").setWrap(true);
  if (lastRow > 1) {
    resultSheet.getRange(2, 1, lastRow - 1)
      .insertCheckboxes();
  }

  const rules = resultSheet.getConditionalFormatRules();

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$A2=TRUE")
      .setRanges([
        resultSheet.getRange(2, 1, lastRow - 1, resultSheet.getMaxColumns())
      ])
      .setBackground("#d9d9d9")
      .setFontColor("#808080")
      .build()
  );

  resultSheet.setConditionalFormatRules(rules);
  resultSheet.setFrozenRows(1);
  resultSheet.setFrozenColumns(9);

}

// FILTER BUILDERS

function buildAdvancedFilter(filterSheet, cfg) {

  const filterData = filterSheet.getRange(cfg.tableRange).getDisplayValues();

  const headers = filterData[0];

  const cityCol = headers.indexOf("Город");
  const typeCol = headers.indexOf("DOOH / Digital Indoor");
  const dayedCol = headers.indexOf("Делают пересчёт по дням");

  const map = new Map();

  for (let i = 1; i < filterData.length; i++) {

    const allowedTypes = String(filterData[i][typeCol])
      .split(",")
      .map(s => s.trim())
      .filter(Boolean);

    const allowedDayed = String(filterData[i][dayedCol])
      .split(",")
      .map(s => s.trim())
      .filter(Boolean);

    const rawCity = String(filterData[i][cityCol]).trim();
    const city =
      rawCity === "МО_ВСЕ_ГОРОДА"
        ? "__MO_ALL__"
        : canonicalCity(rawCity);
    if (!city) continue;

    const filter = {
      types: new Set(allowedTypes),
      dayed: new Set(allowedDayed)
    };

    map.set(city, filter);

    // If this city activates Vegas exceptions,
    // also activate those Vegas malls through the same filter.
    getExceptionMallsForCity(city).forEach(mall => {
      EXCEPTION_MALLS[mall].forEach(exceptionCity => {
        const exceptionCityKey = canonicalCity(exceptionCity);

        if (!map.has(exceptionCityKey)) {
          map.set(exceptionCityKey, {
            types: new Set(filter.types),
            dayed: new Set(filter.dayed)
          });
        }
      });
    });

  }

  return map;
}

function buildTenderFilter(filterSheet, cfg) {
  const data = filterSheet.getRange(cfg.tableRange).getDisplayValues();

  const headers = data[0];

  const cityCol = headers.indexOf("Город");
  const mallCol = headers.indexOf("Торговый центр");
  const typeCol = headers.indexOf("DOOH / Digital Indoor");
  const dayedCol = headers.indexOf("Делают пересчёт по дням");

  const map = new Map();

  for (let i = 1; i < data.length; i++) {
    const rawCity = String(data[i][cityCol]).trim();

    const city =
      rawCity === "МО_ВСЕ_ГОРОДА"
        ? "__MO_ALL__"
        : canonicalCity(rawCity);

    if (!city) continue;

    const mall = String(data[i][mallCol]).trim();

    const filter = {
      mall,
      types: new Set(
        String(data[i][typeCol])
          .split(",")
          .map(s => s.trim())
          .filter(Boolean)
      ),
      dayed: new Set(
        String(data[i][dayedCol])
          .split(",")
          .map(s => s.trim())
          .filter(Boolean)
      )
    };

    if (!map.has(city)) {
      map.set(city, []);
    }

    map.get(city).push(filter);

    // ---------- Exception mall ----------
    const exceptionCities = getExceptionCitiesForMall(mall);

    exceptionCities.forEach(exceptionCity => {
      if (!map.has(exceptionCity)) {
        map.set(exceptionCity, []);
      }

      map.get(exceptionCity).push({
        ...filter,
        mall
      });
    });
  }

  return map;
}

function buildListFilter(filterSheet, cfg) {

  const [citiesRaw, typesRaw, dayedRaw] =
    filterSheet.getRange(cfg.tableRange).getDisplayValues()[0];

  const cities = String(citiesRaw)
    .split(/[;,|\n]+/)
    .map(s => s.trim())
    .filter(Boolean);

  const types = new Set(
    String(typesRaw)
      .split(",")
      .map(s => s.trim())
      .filter(Boolean)
  );

  const dayed = new Set(
    String(dayedRaw)
      .split(",")
      .map(s => s.trim())
      .filter(Boolean)
  );

  const map = new Map();

  cities.forEach(city => {
    const key =
      city === "МО_ВСЕ_ГОРОДА"
        ? "__MO_ALL__"
        : canonicalCity(city);

    map.set(key, {
      types,
      dayed
    });

    getExceptionMallsForCity(key).forEach(exceptionMall => {
      EXCEPTION_MALLS[exceptionMall].forEach(exceptionCity => {
        const exceptionCityKey = canonicalCity(exceptionCity);

        if (!map.has(exceptionCityKey)) {
          map.set(exceptionCityKey, {
            types,
            dayed
          });
        }
      });
    });
  });

  return map;
}
