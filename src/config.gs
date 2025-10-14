/**
 * Bank import – bezpečné rozšírenie CONFIG bez redeklarácie.
 * Nevytvárame 'var/let/const CONFIG', len ho doplníme, ak už existuje.
 * Ak by CONFIG ešte neexistoval, vytvoríme ho cez globalThis (this).
 */
(function () {
  if (typeof CONFIG === 'undefined') {
    // Vytvor CONFIG len ak neexistuje (žiadne redeklarácie).
    this.CONFIG = {};
  }
  // Teraz je CONFIG určite dostupný ako globálna referencia
  CONFIG.BANK = CONFIG.BANK || {};
  // ponecháme existujúcu hodnotu, alebo dáme prázdny string
  CONFIG.BANK.INBOX_FOLDER_ID = CONFIG.BANK.INBOX_FOLDER_ID || "197m6hr3iF39HjDBdNJdJ8HFtfFumTD5B";
})();

// ---- defaults for date/time ----
(function () {
  if (typeof CONFIG === 'undefined') this.CONFIG = {};
  CONFIG.TIMEZONE    = (typeof CONFIG.TIMEZONE    === 'string' && CONFIG.TIMEZONE)    ? CONFIG.TIMEZONE    : 'Europe/Bratislava';
  CONFIG.DATE_FORMAT = (typeof CONFIG.DATE_FORMAT === 'string' && CONFIG.DATE_FORMAT) ? CONFIG.DATE_FORMAT : 'yyyy-MM-dd';
})();
