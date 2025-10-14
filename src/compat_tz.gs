/**
 * Compat: ensure global TZ is a string.
 * Niektoré staršie funkcie používajú Utilities.formatDate(d, TZ, fmt).
 * Ak TZ nie je definované alebo nie je string, nastavíme ho bezpečne.
 */
(function () {
  var safe = 'Europe/Bratislava';
  try {
    var scr = (typeof Session !== 'undefined' && Session.getScriptTimeZone) ? Session.getScriptTimeZone() : null;
    if (typeof TZ !== 'string' || !TZ) this.TZ = (scr && typeof scr === 'string' && scr) || safe;
  } catch (e) {
    this.TZ = safe;
  }
})();
