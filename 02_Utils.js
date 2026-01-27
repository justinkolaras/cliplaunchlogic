function normalizeText_(v) {
  return String(v == null ? '' : v).trim().toLowerCase();
}

function isBlank_(v) {
  return v === '' || v === null || v === undefined;
}

function normalizeSheetDate_(value) {
  var d = null;
  if (value instanceof Date && !isNaN(value)) d = new Date(value);
  else if (typeof value === 'string') {
    var parsed = new Date(value);
    if (!isNaN(parsed)) d = parsed;
  }
  if (!d) return null;
  if (d.getFullYear() < 1950) d.setFullYear(d.getFullYear() + 100);
  d.setHours(0, 0, 0, 0);
  return d;
}

function sameCellValue_(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  return (a == null ? '' : a) === (b == null ? '' : b);
}

function addMonths_(date, months) {
  var d = new Date(date);
  var day = d.getDate();
  d.setMonth(d.getMonth() + months);
  if (d.getDate() < day) d.setDate(0);
  return d;
}

function safeGetSpreadsheetUrl_(ss) {
  try { return ss.getUrl(); } catch (err) { return null; }
}
