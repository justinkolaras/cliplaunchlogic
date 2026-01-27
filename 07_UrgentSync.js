function handleUrgentEdit_(e, ss, mainSheet, urgentSheet) {
  var range = e.range;
  var startRow = range.getRow();
  var endRow = startRow + range.getNumRows() - 1;
  var colStart = range.getColumn();
  var colEnd = range.getLastColumn();

  var idIndex = buildMainIdIndex_(mainSheet);

  var mainLastCol = mainSheet.getLastColumn();
  var urgentMaxDataCol = 1 + mainLastCol;

  var uDataStart = Math.max(2, colStart);
  var uDataEnd = Math.min(colEnd, urgentMaxDataCol);
  var numCols = (uDataStart <= uDataEnd) ? (uDataEnd - uDataStart + 1) : 0;

  var initialDateTouched = (colStart <= 5 && colEnd >= 5);
  var ctx = buildSheetCtx_(mainSheet);

  for (var uRow = Math.max(2, startRow); uRow <= Math.max(2, endRow); uRow++) {
    var mainRow = getMainRowFromUrgent_(urgentSheet, uRow, idIndex.idToRow);
    if (!mainRow) continue;

    if (numCols > 0) {
      var values = urgentSheet.getRange(uRow, uDataStart, 1, numCols).getValues();
      mainSheet.getRange(mainRow, uDataStart - 1, 1, numCols).setValues(values);
    }

    processRow_(mainSheet, mainRow, initialDateTouched ? CFG.COL_INITIAL : undefined, ctx);
  }

  rebuildUrgentFromMain_(ss, mainSheet, urgentSheet);
}

function rebuildUrgentFromMain_(ss, mainSheet, urgentSheet) {
  var mainLastCol = mainSheet.getLastColumn();
  var urgentCols = 1 + mainLastCol;
  var mainLastRow = mainSheet.getLastRow();

  var idIndex = buildMainIdIndex_(mainSheet);

  var props = PropertiesService.getDocumentProperties();
  var prevBuilt = parseInt(props.getProperty('URGENT_LAST_BUILD_ROWS') || '0', 10) || 0;

  var clearRows = Math.max(prevBuilt, Math.max(0, urgentSheet.getLastRow() - 1));
  if (clearRows > 0) {
    var clearRange = urgentSheet.getRange(2, 1, clearRows, urgentCols);
    clearRange.clearContent();
    clearRange.setBackground(CFG.COLORS.RESET);
  }

  if (mainLastRow < 2) {
    props.setProperty('URGENT_LAST_BUILD_ROWS', '0');
    return;
  }

  var separatorDetectColor = getSeparatorColorFromMainRow7_(mainSheet, mainLastCol) || '#ff0000';
  var notesCol = getNotesColIndex_(mainSheet);
  if (notesCol && notesCol > mainLastCol) notesCol = null;

  var numMainRows = mainLastRow - 1;
  var mainRange = mainSheet.getRange(2, 1, numMainRows, mainLastCol);
  var values = mainRange.getValues();
  var bgs = mainRange.getBackgrounds();

  var notesColNotes = notesCol ? mainSheet.getRange(2, notesCol, numMainRows, 1).getNotes() : null;

  var items = [];
  var category = 0;
  var inSep = false;

  for (var i = 0; i < values.length; i++) {
    var mainRow = i + 2;
    var rowVals = values[i];

    var rowEmpty = true;
    for (var k = 0; k < rowVals.length; k++) {
      if (!isBlank_(rowVals[k])) { rowEmpty = false; break; }
    }

    if (rowEmpty) {
      var isSep = isSeparatorRow_(rowVals, bgs[i], separatorDetectColor);
      if (isSep) {
        if (!inSep) category++;
        inSep = true;
      }
      continue;
    }

    inSep = false;

    var status = rowVals[CFG.COL_STATUS - 1];
    var followUp = rowVals[CFG.COL_FOLLOW_FLAG - 1];
    var barType = getBarType_(status, followUp);

    var flagged = false;
    if (notesCol) {
      flagged = isFlagCommand_(rowVals[notesCol - 1]) || (notesColNotes && hasFlagToken_(notesColNotes[i][0]));
    }

    if (barType === 'orange' || barType === 'yellow') {
      items.push({
        mainRow: mainRow,
        rowId: idIndex.rowToId[mainRow] || null,
        category: category,
        barType: barType,
        flagged: flagged,
        rowValues: rowVals
      });
    }
  }

  items.sort(function(a, b) {
    if (a.category !== b.category) return a.category - b.category;
    return a.mainRow - b.mainRow;
  });

  var out = [];
  var meta = [];
  var prevCat = null;

  for (var j = 0; j < items.length; j++) {
    var it = items[j];

    if (prevCat !== null && it.category !== prevCat) {
      out.push(new Array(urgentCols).fill(''));
      meta.push({ type: 'sep' });
    }

    var rowOut = new Array(urgentCols).fill('');
    rowOut[0] = String(it.mainRow);
    for (var c = 0; c < mainLastCol; c++) rowOut[c + 1] = it.rowValues[c];

    out.push(rowOut);
    meta.push({ type: 'data', mainRow: it.mainRow, rowId: it.rowId, barType: it.barType, flagged: it.flagged });

    prevCat = it.category;
  }

  var requiredRows = out.length + 1;
  if (urgentSheet.getMaxRows() < requiredRows) {
    urgentSheet.insertRowsAfter(urgentSheet.getMaxRows(), requiredRows - urgentSheet.getMaxRows());
  }

  if (!out.length) {
    props.setProperty('URGENT_LAST_BUILD_ROWS', '0');
    return;
  }

  applyUrgentTemplateRow_(urgentSheet, urgentCols, out.length);

  var outRange = urgentSheet.getRange(2, 1, out.length, urgentCols);
  outRange.setValues(out);

  var bgsOut = meta.map(function(m) {
    var col = CFG.COLORS.RESET;
    if (m.type === 'data') {
      if (m.flagged) col = CFG.COLORS.FLAG;
      else col = (m.barType === 'orange') ? CFG.COLORS.ORANGE : CFG.COLORS.YELLOW;
    }
    return new Array(urgentCols).fill(col);
  });
  outRange.setBackgrounds(bgsOut);

  var linkRange = urgentSheet.getRange(2, 1, out.length, 1);
  var rich = [];
  var notes = [];

  var baseUrl = safeGetSpreadsheetUrl_(ss);
  var sheetGid = mainSheet.getSheetId();

  for (var mIdx = 0; mIdx < meta.length; mIdx++) {
    var m = meta[mIdx];

    if (m.type !== 'data' || !baseUrl) {
      rich.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
      notes.push(['']);
      continue;
    }

    var deepLink = baseUrl + '#gid=' + sheetGid + '&range=A' + m.mainRow;
    rich.push([SpreadsheetApp.newRichTextValue().setText(String(m.mainRow)).setLinkUrl(deepLink).build()]);
    notes.push([m.rowId ? 'ROW_ID=' + m.rowId : '']);
  }

  linkRange.setRichTextValues(rich);
  linkRange.setNotes(notes);

  props.setProperty('URGENT_LAST_BUILD_ROWS', String(out.length));
}

function applyUrgentTemplateRow_(urgentSheet, urgentCols, outRows) {
  if (!outRows) return;
  try {
    var template = urgentSheet.getRange(2, 1, 1, urgentCols);
    var dest = urgentSheet.getRange(2, 1, outRows, urgentCols);
    template.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    var dv = template.getDataValidations();
    if (dv && dv[0]) {
      var dvs = Array.from({ length: outRows }, function() { return dv[0].slice(); });
      dest.setDataValidations(dvs);
    }
  } catch (err) {}
}

function getMainRowFromUrgent_(urgentSheet, urgentRow, idToRow) {
  var cell = urgentSheet.getRange(urgentRow, 1);
  var id = parseRowIdFromNote_(cell.getNote());
  if (id && idToRow[id]) return idToRow[id];
  return parseMainRowNumber_(cell.getDisplayValue());
}

function parseMainRowNumber_(v) {
  var s = String(v || '').trim();
  if (/^\d+$/.test(s)) return parseInt(s, 10);
  var m = s.match(/range=[A-Za-z]+(\d+)/) || s.match(/(\d+)\s*$/);
  return m ? parseInt(m[1], 10) : null;
}
