const sourceSS = SpreadsheetApp.openByUrl(
  "https://docs.google.com/spreadsheets/d/1rozN2fSu2C46Q4SLMNQL57WeS1QG36Tn/edit" // Тестовый АП: "https://docs.google.com/spreadsheets/d/1xQ5x-QUnTEPe2XiyPCOFVIVL13w9f6MPv4L6df4-TT4/edit"
);
const sourceSheet = sourceSS.getSheetByName('Malls_01.03.2024 - 31.03.2024');