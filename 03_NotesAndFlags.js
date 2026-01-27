function getNotesColIndex_(sheet) {
  var sid = String(sheet.getSheetId());
  if (Object.prototype.hasOwnProperty.call(__NOTES_COL_CACHE, sid)) return __NOTES_COL_CACHE[sid];

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    __NOTES_COL_CACHE[sid] = null;
    return null;
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = null;

  for (var c = 0; c < headers.length; c++) {
    var h = normalizeText_(headers[c]);
    if (h === 'notes' || h === 'note') {
      idx = c + 1;
      break;
    }
  }

  __NOTES_COL_CACHE[sid] = idx;
  return idx;
}

function isFlagCommand_(s) {
  return /^\s*!flag(\s+|$)/i.test(String(s == null ? '' : s));
}

function stripFlagCommand_(s) {
  return String(s == null ? '' : s).replace(/^\s*!flag(\s+|$)/i, '');
}

function hasFlagToken_(note) {
  return new RegExp('\\b' + CFG.TOKENS.FLAG_NOTE.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b').test(String(note || ''));
}

function ensureFlagToken_(range) {
  var current = String(range.getNote() || '');
  if (hasFlagToken_(current)) return;
  range.setNote((current ? current + '\n' : '') + CFG.TOKENS.FLAG_NOTE);
}

function readAndApplyFlagFromNotes_(sheet, row, notesCol) {
  if (!notesCol) return false;

  var cell = sheet.getRange(row, notesCol);
  var value = String(cell.getValue() == null ? '' : cell.getValue());
  var note = String(cell.getNote() || '');

  if (isFlagCommand_(value)) {
    var cleaned = stripFlagCommand_(value);
    if (cleaned !== value) cell.setValue(cleaned);
    ensureFlagToken_(cell);
    return true;
  }

  return hasFlagToken_(note);
}
