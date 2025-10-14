/**
 * Date utils – bezpečné formátovanie s korektným timezone stringom.
 */

function __tz__() {
  // vždy vráť string; ak by nebola timezone nastavená, fallback na Europe/Bratislava
  return Session.getScriptTimeZone() || 'Europe/Bratislava';
}

/** Format date (default: yyyy-MM-dd) */
function formatDate_(d, pattern) {
  if (!d) return '';
  if (!(d instanceof Date)) d = new Date(d);
  var fmt = pattern || 'yyyy-MM-dd';
  return Utilities.formatDate(d, __tz__(), fmt);
}

/** Format date-time (yyyy-MM-dd HH:mm) */
function formatDateTime_(d) {
  return formatDate_(d, 'yyyy-MM-dd HH:mm');
}

/** Parse rôznych textových dátumov na Date alebo null */
function parseDateMaybe_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();

  var m1 = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/); // dd.mm.yyyy
  if (m1) return new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));

  var m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);              // yyyy-mm-dd
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));

  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
