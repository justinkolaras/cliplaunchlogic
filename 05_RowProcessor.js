'use strict';

function processRow_(sheet, row, editedColumn, ctx) {
  if (!ctx) ctx = { lastCol: sheet.getLastColumn(), notesCol: getNotesColIndex_(sheet) };

  const block = sheet.getRange(row, CFG.COL_STATUS, 1, 4).getValues()[0];
  const status0 = block[0];
  const initial0 = block[1];
  const followFlag0 = block[2];
  const follow0 = block[3];

  let status = status0;
  let followFlag = followFlag0;
  let followObj = follow0;

  const initialDateNormalized = normalizeSheetDate_(initial0);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (editedColumn === CFG.COL_INITIAL) {
    if (initialDateNormalized) {
      const followDate = new Date(initialDateNormalized);
      followDate.setDate(followDate.getDate() + 2);
      followDate.setHours(0, 0, 0, 0);
      followObj = followDate;
      followFlag = status === 'Sent' ? 'Wait' : '';
    } else {
      followFlag = '';
      followObj = '';
    }
  }

  if (String(status ?? '').trim() === 'Responded') {
    followFlag = 'Not needed';
  }

  const followDateObj = normalizeSheetDate_(followObj);
  const hasValidFollowDate = Boolean(followDateObj);
  if (hasValidFollowDate) followObj = followDateObj;

  if (status === 'Sent' && hasValidFollowDate && !String(followFlag ?? '').trim()) {
    followFlag = 'Wait';
  }

  if (status !== 'Sent' && followFlag === 'Wait') {
    followFlag = '';
  }

  if (hasValidFollowDate) {
    followDateObj.setHours(0, 0, 0, 0);
    if (followFlag === 'Wait' && status === 'Sent' && today >= followDateObj) {
      followFlag = 'Follow Up Now';
    }
  }

  if (initialDateNormalized) {
    const initialDate = new Date(initialDateNormalized);
    initialDate.setHours(0, 0, 0, 0);

    const threeMonths = addMonths_(initialDate, 3);
    const sixMonths = addMonths_(initialDate, 6);
    threeMonths.setHours(0, 0, 0, 0);
    sixMonths.setHours(0, 0, 0, 0);

    if (hasValidFollowDate && today >= threeMonths && today < sixMonths) {
      status = 'Consider again (3 months)';
    }

    if (today >= sixMonths && status === 'Consider again (3 months)') {
      status = 'Consider again! (6 months)';
    }
  }

  if (status !== status0) sheet.getRange(row, CFG.COL_STATUS).setValue(status);

  if (String(followFlag ?? '') !== String(followFlag0 ?? '')) {
    const cellFF = sheet.getRange(row, CFG.COL_FOLLOW_FLAG);
    if (!followFlag) cellFF.clearContent();
    else cellFF.setValue(followFlag);
  }

  if (!sameCellValue_(followObj, follow0)) {
    const cellFD = sheet.getRange(row, CFG.COL_FOLLOW_DATE);
    if (!followObj) cellFD.clearContent();
    else cellFD.setValue(followObj);
  }

  const isFlagged = readAndApplyFlagFromNotes_(sheet, row, ctx.notesCol);

  const rowRange = sheet.getRange(row, 1, 1, ctx.lastCol);

  if (isFlagged) {
    rowRange.setBackground(CFG.COLORS.FLAG);
    return;
  }

  const barType = getBarType_(status, followFlag);
  applyRowBar_(rowRange, barType);
}

function getBarType_(status, followFlag) {
  const s = normalizeText_(status);
  const f = normalizeText_(followFlag);

  if (s === 'sale') return 'green';
  if (s === 'responded' || s === 'offer provided, awaiting') return 'orange';

  if (f === 'follow up now' || f === 'consider again (3 months)' || s === 'consider again (3 months)') {
    return 'yellow';
  }

  return null;
}

function applyRowBar_(rowRange, barType) {
  let desired = null;

  if (barType === 'orange') desired = CFG.COLORS.ORANGE;
  else if (barType === 'yellow') desired = CFG.COLORS.YELLOW;
  else if (barType === 'green') desired = CFG.COLORS.GREEN;

  if (desired) {
    rowRange.setBackground(desired);
    return;
  }

  const bgs = rowRange.getBackgrounds();
  const flat = bgs[0] ?? [];
  if (!flat.length) return;

  const first = flat[0];
  for (let i = 1; i < flat.length; i++) if (flat[i] !== first) return;

  if (
    first === CFG.COLORS.ORANGE ||
    first === CFG.COLORS.YELLOW ||
    first === CFG.COLORS.GREEN ||
    first === CFG.COLORS.FLAG
  ) {
    rowRange.setBackground(CFG.COLORS.RESET);
  }
}
