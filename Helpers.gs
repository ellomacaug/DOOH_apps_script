function formatPeriod(fromDate, toDate) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(fromDate, tz, "dd.MM.yyyy") +
        " - " +
        Utilities.formatDate(toDate, tz, "dd.MM.yyyy");
}

function normalize(str) {
  return String(str)
    .toLowerCase()
    .replace(/\((?:г\.о\.|г\.|п\.|д\.|с\.)\s*/g, "")
    .replace(/\s+/g, "")
    .replace("ё", "е")
    .replace(/[-_.,;:()]/g, "");
}

function canonicalCity(city) {
  city = String(city).trim();
  const nCity = normalize(city);

  const m = city.match(/^МО\s*\(\s*(?:г\.о\.|г\.|п\.|д\.|с\.)\s*([^)]+)\s*\)$/i);
  if (m) return normalize(m[1]);

  if (nCity === normalize("МО")) return "__MO__";

  const shortCityMap = {
    "спб": "Санкт-Петербург",
    "мск": "Москва",
    "нн": "Нижний Новгород",
    "рнд": "Ростов-на-Дону",
    "екб": "Екатеринбург"
  };

  return normalize(shortCityMap[nCity] || city);
}

function parseCity(city) {
  const raw = String(city).trim();

  const match = raw.match(
    /^(.+?)\s*\(\s*(?:г\.о\.|г\.|п\.|д\.|с\.)\s*([^)]+)\s*\)$/i
  );

  if (match) {
    return {
      parent: canonicalCity(match[1]),
      city: canonicalCity(match[2]),
      isSubCity: true
    };
  }

  return {
    parent: null,
    city: canonicalCity(raw),
    isSubCity: false
  };
}


function getFilterKey(rowCityRaw, cityFilterMap) {
  const parsed = parseCity(rowCityRaw);

  if (parsed.city && cityFilterMap.has(parsed.city)) {
    return parsed.city;
  }

  if (parsed.parent && cityFilterMap.has(parsed.parent)) {
    return parsed.parent;
  }

  return null;
}


function citySortInfo(city) {
  city = String(city).trim();

  // Treat all Московская область variants as "МО"
  const key = city.startsWith("МО (")
    ? normalize("МО")
    : normalize(city);

  return {
    priority: CITY_PRIORITY.has(key)
      ? CITY_PRIORITY.get(key)
      : 999,
    name: city
      .replace(/\((?:г\.о\.|г\.|д\.|с\.)\s*/g, "")
      .replace(/\)/g, "")
  };
}

function normalizeMall(str) {
  return normalize(
    String(str).replace("Крокус Экспо / ", "").replace(/\s*\([^)]*\)/g, "")
  );
}

function getMallKey(city, mall) {
  return `${normalize(city)}|${normalizeMall(mall)}`;
}

function getExceptionCitiesForMall(mall) {
  const mallKey = normalizeMall(mall);

  const entry = Object.entries(EXCEPTION_MALLS)
    .find(([exceptionMall]) => normalizeMall(exceptionMall) === mallKey);

  if (!entry) return [];

  return entry[1].map(city => canonicalCity(city));
}


function getExceptionMallsForCity(city) {
  const cityKey = canonicalCity(city);
  const result = [];

  Object.entries(EXCEPTION_MALLS).forEach(([mall, cities]) => {
    const cityKeys = cities.map(c => canonicalCity(c));

    if (cityKeys.includes(cityKey)) {
      result.push(mall);
    }
  });

  return result;
}

function getExceptionMallForRow(mall) {
  const mallKey = normalizeMall(mall);

  return Object.keys(EXCEPTION_MALLS)
    .find(exceptionMall =>
      normalizeMall(exceptionMall) === mallKey
    ) || null;
}

function getExceptionFilterKey(mall, cityFilterMap) {
  const exceptionMall = getExceptionMallForRow(mall);
  if (!exceptionMall) {
    return null;
  }
  for (const filterCity of cityFilterMap.keys()) {

    if (filterCity === "__MO_ALL__") {
      continue;
    }
    const exceptionMalls = getExceptionMallsForCity(filterCity);
    if (
      exceptionMalls.some(
        exceptionMallName =>
          normalizeMall(exceptionMallName) === normalizeMall(exceptionMall)
      )
    ) {
      return filterCity;
    }
  }
  return null;
}



function getWorkingHours(schedule) {
  schedule = String(schedule).trim().toLowerCase();

  if (schedule === "круглосуточно") {
    return 24;
  }

  const match = schedule.match(/(\d{2}):(\d{2}).*?(\d{2}):(\d{2})/);
  if (!match) return 0;

  const [, h1, m1, h2, m2] = match;

  const start = Number(h1) + Number(m1) / 60;
  let end = Number(h2) + Number(m2) / 60;

  // Crosses midnight
  if (end <= start) {
    end += 24;
  }

  return Math.round((end - start) * 100) / 100;
}


function removeCheckedRows(sheet) {
  let lastRow = sheet.getLastRow();

  const checked = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = checked.length - 1; i >= 0; i--) {
    if (checked[i][0]) {
      sheet.deleteRow(i + 2);
    }
  };
  sheet.getRange(1, 1, lastRow).clearContent().clearDataValidations().clearFormat();
}
