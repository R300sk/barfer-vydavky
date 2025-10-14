const CONFIG = {
  TIMEZONE: "Europe/Bratislava",
  DATE_FORMAT: "dd.MM.yyyy",
  SHEETS: {
    VYDAVKY: "Výdavky",
    TRZBY: "Tržby",
    PERSONÁL: "Personál",
    DASHBOARD: "Dashboard",
    SUMMARY: "Mesačný výkaz"
  },
  HEADERS: {
    VYDAVKY: [
      "Dátum", "Názov výdavku", "Suma (€) bez DPH", "Kategória",
      "Projekt", "Dodávateľ", "Poznámka", "Doklad (odkaz)"
    ],
    TRZBY: [
      "Mesiac", "Tržba celkom bez DPH", "Počet objednávok", "Profit z predaja tovaru"
    ],
    PERSONÁL: [
      "Meno", "Plánované hodiny", "Skutočné hodiny", "Sadzba €/hod", "Bonus", "Poznámka"
    ],
    SUMMARY: [
      "Mesiac", "Výdavky spolu", "Tržby spolu", "Profit", "Poznámka"
    ]
  }
};
// ==== Bank import config ====
if (typeof CONFIG === 'undefined') { var CONFIG = {}; }
if (!CONFIG.BANK) { CONFIG.BANK = {}; }
// TODO: nastav Drive folder ID, kam budeš hádzať výpisy (XML/CSV):
// Príklad: CONFIG.BANK.INBOX_FOLDER_ID = "1AbCdEfGhIjKlMn...";
CONFIG.BANK.INBOX_FOLDER_ID = "";
