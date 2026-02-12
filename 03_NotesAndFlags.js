'use strict';

function getNotesColIndex_(sheet) {
  const sid = String(sheet.getSheetId());
  if (NOTES_COL_CACHE.has(sid)) return NOTES_COL_CACHE.get(sid);

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    NOTES_COL_CACHE.set(sid, null);
    return null;
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = null;

  for (let c = 0; c < headers.length; c++) {
    const h = normalizeText_(headers[c]);
    if (h === 'notes' || h === 'note') {
      idx = c + 1;
      break;
    }
  }

  NOTES_COL_CACHE.set(sid, idx);
  return idx;
}

function isFlagCommand_(s) {
  return /^\s*!flag(\s+|$)/i.test(String(s ?? ''));
}

function stripFlagCommand_(s) {
  return String(s ?? '').replace(/^\s*!flag(\s+|$)/i, '');
}

function isUnflagCommand_(s) {
  return /^\s*!unflag(\s+|$)/i.test(String(s ?? ''));
}

function hasFlagToken_(note) {
  const token = CFG.TOKENS.FLAG_NOTE.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`\\b${token}\\b`).test(String(note ?? ''));
}

function ensureFlagToken_(range) {
  const current = String(range.getNote() ?? '');
  if (hasFlagToken_(current)) return;
  range.setNote((current ? `${current}\n` : '') + CFG.TOKENS.FLAG_NOTE);
}

function removeFlagToken_(range) {
  const current = String(range.getNote() ?? '');
  if (!hasFlagToken_(current)) return;

  const token = CFG.TOKENS.FLAG_NOTE;
  const filtered = current
    .split('\n')
    .map((line) => String(line ?? '').trim())
    .filter((line) => line && line !== token);

  range.setNote(filtered.join('\n'));
}

function setFlaggedRichText_(cell, reason) {
  const reasonText = String(reason ?? '').trim();
  const text = reasonText ? `FLAGGED ${reasonText}` : 'FLAGGED';
  const style = SpreadsheetApp.newTextStyle().setBold(true).build();
  const rich = SpreadsheetApp.newRichTextValue().setText(text).setTextStyle(0, 7, style).build();
  cell.setRichTextValue(rich);
}

function readAndApplyFlagFromNotes_(sheet, row, notesCol) {
  if (!notesCol) return false;

  const cell = sheet.getRange(row, notesCol);
  const value = String(cell.getValue() ?? '');
  const note = String(cell.getNote() ?? '');

  if (isUnflagCommand_(value)) {
    cell.clearContent();
    removeFlagToken_(cell);
    return false;
  }

  if (isFlagCommand_(value)) {
    const reason = stripFlagCommand_(value).trim();
    setFlaggedRichText_(cell, reason);
    ensureFlagToken_(cell);
    return true;
  }

  return hasFlagToken_(note);
}
