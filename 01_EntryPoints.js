'use strict';

function mainLogic(e) {
  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(20000);
  } catch (_) {
    return;
  }

  try {
    const ss = e?.source ?? SpreadsheetApp.getActive();
    const mainSheet = getMainSheet_(ss);
    if (!mainSheet) return;

    const urgentSheet = ss.getSheetByName(CFG.URGENT_NAME);

    if (e?.range) {
      const sheet = e.range.getSheet();
      if (!sheet) return;

      if (e.range.getRow() === 1) NOTES_COL_CACHE.clear();

      const sheetName = sheet.getName();

      if (sheetName === mainSheet.getName()) {
        const startRow = e.range.getRow();
        const endRow = startRow + e.range.getNumRows() - 1;

        const colStart = e.range.getColumn();
        const colEnd = e.range.getLastColumn();

        const editedColumn =
          colStart <= CFG.COL_INITIAL && colEnd >= CFG.COL_INITIAL
            ? CFG.COL_INITIAL
            : colStart;

        const ctx = buildSheetCtx_(mainSheet);

        for (let r = Math.max(2, startRow); r <= endRow; r++) {
          processRow_(mainSheet, r, editedColumn, ctx);
        }

        if (urgentSheet) rebuildUrgentFromMain_(ss, mainSheet, urgentSheet);
        return;
      }

      if (urgentSheet && sheetName === CFG.URGENT_NAME) {
        handleUrgentEdit_(e, ss, mainSheet, urgentSheet);
        return;
      }

      return;
    }

    const lastRow = mainSheet.getLastRow();
    const ctxAll = buildSheetCtx_(mainSheet);

    for (let row = 2; row <= lastRow; row++) {
      processRow_(mainSheet, row, undefined, ctxAll);
    }

    if (urgentSheet) rebuildUrgentFromMain_(ss, mainSheet, urgentSheet);
  } finally {
    lock.releaseLock();
  }
}

function buildSheetCtx_(sheet) {
  const lastCol = sheet.getLastColumn();
  let notesCol = getNotesColIndex_(sheet);
  if (notesCol && notesCol > lastCol) notesCol = null;
  return { lastCol, notesCol };
}

function getMainSheet_(ss) {
  for (const name of CFG.MAIN_NAMES) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function onEdit(e) {
  mainLogic(e);
}

function onOpen(e) {
  mainLogic(e);
}

function rebuildAll() {
  mainLogic();
}
