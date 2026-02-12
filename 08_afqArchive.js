'use strict';

function archiveAfqRows() {
  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(20000);
  } catch (_) {
    return;
  }

  try {
    const ss = SpreadsheetApp.getActive();
    const mainSheet = getMainSheet_(ss);
    if (!mainSheet) return;

    const afqLogSheet = getAfqLogSheet_(ss);
    if (!afqLogSheet) return;

    const mainLastRow = mainSheet.getLastRow();
    const mainLastCol = mainSheet.getLastColumn();
    if (mainLastRow < 2 || mainLastCol < 1) return;

    const data = mainSheet.getRange(2, 1, mainLastRow - 1, mainLastCol).getValues();
    const rowsToMove = [];

    for (let i = 0; i < data.length; i++) {
      const status = normalizeText_(data[i][CFG.COL_STATUS - 1]);
      if (status === 'a/fq') {
        rowsToMove.push(i + 2);
      }
    }

    if (rowsToMove.length) {
      const logCols = Math.max(1, afqLogSheet.getLastColumn());
      const copyCols = Math.min(mainLastCol, logCols);
      const out = rowsToMove.map((r) => data[r - 2].slice(0, copyCols));
      const appendStart = Math.max(1, afqLogSheet.getLastRow() + 1);
      afqLogSheet.getRange(appendStart, 1, out.length, copyCols).setValues(out);

      for (let i = rowsToMove.length - 1; i >= 0; i--) {
        mainSheet.deleteRow(rowsToMove[i]);
      }
    }

    collapseConsecutiveSeparatorRows_(mainSheet);

    const urgentSheet = ss.getSheetByName(CFG.URGENT_NAME);
    if (urgentSheet) rebuildUrgentFromMain_(ss, mainSheet, urgentSheet);
  } finally {
    lock.releaseLock();
  }
}

function getAfqLogSheet_(ss) {
  for (const name of CFG.AFQ_LOG_NAMES) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function collapseConsecutiveSeparatorRows_(mainSheet) {
  const lastRow = mainSheet.getLastRow();
  const lastCol = mainSheet.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return;

  const sepColor = getSeparatorColorFromMainRow7_(mainSheet, lastCol) || '#ff0000';
  const vals = mainSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const bgs = mainSheet.getRange(2, 1, lastRow - 1, lastCol).getBackgrounds();

  const rowsToDelete = [];
  let prevWasSep = false;

  for (let i = 0; i < vals.length; i++) {
    const isSep = isSeparatorRow_(vals[i], bgs[i], sepColor);
    if (isSep && prevWasSep) rowsToDelete.push(i + 2);
    prevWasSep = isSep;
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    mainSheet.deleteRow(rowsToDelete[i]);
  }
}
