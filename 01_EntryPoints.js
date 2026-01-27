function mainLogic(e) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return;
  }

  try {
    var ss = e && e.source ? e.source : SpreadsheetApp.getActive();
    var mainSheet = getMainSheet_(ss);
    if (!mainSheet) return;

    var urgentSheet = ss.getSheetByName(CFG.URGENT_NAME);
    if (e && e.range) {
      var sheet = e.range.getSheet();
      if (!sheet) return;

      if (e.range.getRow() === 1) __NOTES_COL_CACHE = {};

      var sheetName = sheet.getName();
      if (sheetName === mainSheet.getName()) {
        var startRow = e.range.getRow();
        var endRow = startRow + e.range.getNumRows() - 1;

        var colStart = e.range.getColumn();
        var colEnd = e.range.getLastColumn();
        var editedColumn = (colStart <= CFG.COL_INITIAL && colEnd >= CFG.COL_INITIAL) ? CFG.COL_INITIAL : colStart;

        var ctx = buildSheetCtx_(mainSheet);

        for (var r = Math.max(2, startRow); r <= endRow; r++) {
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

    var lastRow = mainSheet.getLastRow();
    var ctxAll = buildSheetCtx_(mainSheet);
    for (var row = 2; row <= lastRow; row++) {
      processRow_(mainSheet, row, undefined, ctxAll);
    }
    if (urgentSheet) rebuildUrgentFromMain_(ss, mainSheet, urgentSheet);
  } finally {
    lock.releaseLock();
  }
}

function buildSheetCtx_(sheet) {
  var lastCol = sheet.getLastColumn();
  var notesCol = getNotesColIndex_(sheet);
  if (notesCol && notesCol > lastCol) notesCol = null;
  return { lastCol: lastCol, notesCol: notesCol };
}

function getMainSheet_(ss) {
  for (var i = 0; i < CFG.MAIN_NAMES.length; i++) {
    var sh = ss.getSheetByName(CFG.MAIN_NAMES[i]);
    if (sh) return sh;
  }
  return null;
}

function onEdit(e) { mainLogic(e); }
function onOpen(e) { mainLogic(e); }

function rebuildAll() { mainLogic(); }
