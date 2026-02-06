'use strict';

function normalizeText_(v) {
  return String(v ?? '').trim().toLowerCase();
}

function isBlank_(v) {
  return v === '' || v === null || v === undefined;
}

function normalizeSheetDate_(value) {
  let d = null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    d = new Date(value);
  } else if (typeof value === 'string') {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) d = parsed;
  }

  if (!d) return null;

  if (d.getFullYear() < 1950) d.setFullYear(d.getFullYear() + 100);

  d.setHours(0, 0, 0, 0);
  return d;
}

function sameCellValue_(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  return String(a ?? '') === String(b ?? '');
}

function addMonths_(date, months) {
  const d = new Date(date);
  const day = d.getDate();
  d.setMonth(d.getMonth() + months);
  if (d.getDate() < day) d.setDate(0);
  return d;
}

function safeGetSpreadsheetUrl_(ss) {
  try {
    return ss.getUrl();
  } catch (_) {
    return null;
  }
}
