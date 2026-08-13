const CONFIG = {
  advanced: {
    sheet: "Фильтр_Расширенный",
    tableRange: "B2:D",
    fromCell: "G5",
    toCell: "G6",
    brandCell: "G4",
    raCell: "G3",
    zaprosCell: "G2",
    shopCell: "G7"
  },

  tender: {
    sheet: "Фильтр_Тендер",
    tableRange: "B2:E",
    fromCell: "H5",
    toCell: "H6",
    brandCell: "H4",
    raCell: "H3",
    zaprosCell: "H2",
    shopCell: "H7"
  },

  list: {
    sheet: "Фильтр_Списком",
    tableRange: "B10:D10",
    fromCell: "C5",
    toCell: "C6",
    brandCell: "C4",
    raCell: "C3",
    zaprosCell: "C2",
    shopCell: "C7"
  }
};


const RA_CONFIG = {
  "MASTERAD": {
    aliases: ["мастерэд", "мастеред", "masterad", "mastered"],
    k: "1,63",
    color1: "#4f81bd",
    color2: "#dce6f1"
  },

  "BUYTECKNOWLOGY": {
    aliases: ["buytecknowlogy", "buytechnology"],
    k: "1,79",
    color1: "#7f7f7f",
    color2: "#d8d8d8"
  },

  "POSTERSCOPE": {
    aliases: ["posterscope", "постерскоп", "posterskop"],
    k: "1,46",
    color1: "#4f81bd",
    color2: "#dce6f1"
  },

  "SKYTECKNOWLOGY": {
    aliases: ["скайтек", "skytecknowlogy", "skytek", "skytech"],
    k: "1,79",
    color1: "#4f81bd",
    color2: "#dce6f1"
  },

  "GROUP4M": {
    aliases: ["group4m", "груп4м"],
    k: "1,59",
    color1: "#4f81bd",
    color2: "#dce6f1"
  },

  "NMI": {
    aliases: ["nmi", "нми"],
    k: "1,46",
    color1: "#0033cc",
    color2: "#d8d8d8"
  },

  "OMD": {
    aliases: ["омд", "omd"],
    k: "1,45",
    color1: "red",
    color2: "red"
  }
};


const CITY_PRIORITY = new Map([
  "Москва",
  "МО",
  "Санкт-Петербург",
  "Новосибирск",
  "Екатеринбург",
  "Казань",
  "Нижний Новгород",
  "Красноярск",
  "Челябинск",
  "Самара",
  "Уфа",
  "Ростов-на-Дону",
  "Краснодар",
  "Омск",
  "Воронеж",
  "Пермь",
  "Волгоград"
].map((city, i) => [normalize(city), i]));


EXCEPTION_MALLS = {
  "Vegas Крокус Сити" : ["г. Красногорск", "Москва", "МО", "Красногорск"],
  "Vegas Кунцево" : ["МО (г.о. Одинцовский)", "с. Немчиновка", "г.о. Одинцовский", "Москва", "МО"],
  "Vegas Каширское Шоссе" : ["МО (г.о. Ленинский)", "п. Совхоз имени Ленина", "МО (п. Совхоз имени Ленина)", "г.о. Ленинский", "Москва", "МО"]
}
