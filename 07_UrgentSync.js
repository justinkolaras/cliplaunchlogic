'use strict';

function handleUrgentEdit_(e, ss, mainSheet, urgentSheet) {
  const range = e.range;
  const startRow = range.getRow();
  const endRow = startRow + range.getNumRows() - 1;
  const colStart = range.getColumn();
  const colEnd = range.getLastColumn();

  const idIndex = buildMainIdIndex_(mainSheet);

  const mainLastCol = mainSheet.getLastColumn();
  const urgentMaxDataCol = 1 + mainLastCol;

  const uDataStart = Math.max(2, colStart);
  const uDataEnd = Math.min(colEnd, urgentMaxDataCol);
  const numCols = uDataStart <= uDataEnd ? uDataEnd - uDataStart + 1 : 0;

  const initialDateTouched = colStart <= 5 && colEnd >= 5;
  const ctx = buildSheetCtx_(mainSheet);

  for (let uRow = Math.max(2, startRow); uRow <= Math.max(2, endRow); uRow++) {
    const mainRow = getMainRowFromUrgent_(urgentSheet, uRow, idIndex.idToRow);
    if (!mainRow) continue;

    if (numCols > 0) {
      const values = urgentSheet.getRange(uRow, uDataStart, 1, numCols).getValues();
      mainSheet.getRange(mainRow, uDataStart - 1, 1, numCols).setValues(values);
    }

    processRow_(mainSheet, mainRow, initialDateTouched ? CFG.COL_INITIAL : undefined, ctx);
  }

  rebuildUrgentFromMain_(ss, mainSheet, urgentSheet);
}

function rebuildUrgentFromMain_(ss, mainSheet, urgentSheet) {
  const mainLastCol = mainSheet.getLastColumn();
  const urgentCols = 1 + mainLastCol;
  const mainLastRow = mainSheet.getLastRow();

  const idIndex = buildMainIdIndex_(mainSheet);

  const props = PropertiesService.getDocumentProperties();
  const prevBuilt = parseInt(props.getProperty('URGENT_LAST_BUILD_ROWS') || '0', 10) || 0;

  const clearRows = Math.max(prevBuilt, Math.max(0, urgentSheet.getLastRow() - 1));
  if (clearRows > 0) {
    const clearRange = urgentSheet.getRange(2, 1, clearRows, urgentCols);
    clearRange.clearContent();
    clearRange.setBackground(CFG.COLORS.RESET);
  }

  if (mainLastRow < 2) {
    props.setProperty('URGENT_LAST_BUILD_ROWS', '0');
    return;
  }

  const separatorDetectColor = getSeparatorColorFromMainRow7_(mainSheet, mainLastCol) || '#ff0000';

  let notesCol = getNotesColIndex_(mainSheet);
  if (notesCol && notesCol > mainLastCol) notesCol = null;

  const numMainRows = mainLastRow - 1;
  const mainRange = mainSheet.getRange(2, 1, numMainRows, mainLastCol);
  const values = mainRange.getValues();
  const bgs = mainRange.getBackgrounds();

  const notesColNotes = notesCol ? mainSheet.getRange(2, notesCol, numMainRows, 1).getNotes() : null;

  const items = [];
  let category = 0;
  let inSep = false;

  for (let i = 0; i < values.length; i++) {
    const mainRow = i + 2;
    const rowVals = values[i];

    let rowEmpty = true;
    for (let k = 0; k < rowVals.length; k++) {
      if (!isBlank_(rowVals[k])) {
        rowEmpty = false;
        break;
      }
    }

    if (rowEmpty) {
      const isSep = isSeparatorRow_(rowVals, bgs[i], separatorDetectColor);
      if (isSep) {
        if (!inSep) category++;
        inSep = true;
      }
      continue;
    }

    inSep = false;

    const status = rowVals[CFG.COL_STATUS - 1];
    const followUp = rowVals[CFG.COL_FOLLOW_FLAG - 1];
    const barType = getBarType_(status, followUp);

    let flagged = false;
    if (notesCol) {
      flagged =
        isFlagCommand_(rowVals[notesCol - 1]) ||
        Boolean(notesColNotes && hasFlagToken_(notesColNotes[i][0]));
    }

    if (barType === 'orange' || barType === 'yellow') {
      items.push({
        mainRow,
        rowId: idIndex.rowToId[mainRow] || null,
        category,
        barType,
        flagged,
        rowValues: rowVals,
      });
    }
  }

  items.sort((a, b) => (a.category !== b.category ? a.category - b.category : a.mainRow - b.mainRow));

  const out = [];
  const meta = [];
  let prevCat = null;

  for (let j = 0; j < items.length; j++) {
    const it = items[j];

    if (prevCat !== null && it.category !== prevCat) {
      out.push(Array(urgentCols).fill(''));
      meta.push({ type: 'sep' });
    }

    const rowOut = Array(urgentCols).fill('');
    rowOut[0] = String(it.mainRow);
    for (let c = 0; c < mainLastCol; c++) rowOut[c + 1] = it.rowValues[c];

    out.push(rowOut);
    meta.push({ type: 'data', mainRow: it.mainRow, rowId: it.rowId, barType: it.barType, flagged: it.flagged });

    prevCat = it.category;
  }

  const requiredRows = out.length + 1;
  if (urgentSheet.getMaxRows() < requiredRows) {
    urgentSheet.insertRowsAfter(urgentSheet.getMaxRows(), requiredRows - urgentSheet.getMaxRows());
  }

  if (!out.length) {
    props.setProperty('URGENT_LAST_BUILD_ROWS', '0');
    return;
  }

  applyUrgentTemplateRow_(urgentSheet, urgentCols, out.length);

  const outRange = urgentSheet.getRange(2, 1, out.length, urgentCols);
  outRange.setValues(out);

  const bgsOut = meta.map((m) => {
    let col = CFG.COLORS.RESET;
    if (m.type === 'data') {
      if (m.flagged) col = CFG.COLORS.FLAG;
      else col = m.barType === 'orange' ? CFG.COLORS.ORANGE : CFG.COLORS.YELLOW;
    }
    return Array(urgentCols).fill(col);
  });
  outRange.setBackgrounds(bgsOut);

  const linkRange = urgentSheet.getRange(2, 1, out.length, 1);
  const rich = [];
  const notes = [];

  const baseUrl = safeGetSpreadsheetUrl_(ss);
  const sheetGid = mainSheet.getSheetId();

  for (let mIdx = 0; mIdx < meta.length; mIdx++) {
    const m = meta[mIdx];

    if (m.type !== 'data' || !baseUrl) {
      rich.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
      notes.push(['']);
      continue;
    }

    const deepLink = `${baseUrl}#gid=${sheetGid}&range=A${m.mainRow}`;
    rich.push([SpreadsheetApp.newRichTextValue().setText(String(m.mainRow)).setLinkUrl(deepLink).build()]);
    notes.push([m.rowId ? `${CFG.TOKENS.ROW_ID_PREFIX}${m.rowId}` : '']);
  }

  linkRange.setRichTextValues(rich);
  linkRange.setNotes(notes);

  props.setProperty('URGENT_LAST_BUILD_ROWS', String(out.length));
}

function applyUrgentTemplateRow_(urgentSheet, urgentCols, outRows) {
  if (!outRows) return;

  try {
    const template = urgentSheet.getRange(2, 1, 1, urgentCols);
    const dest = urgentSheet.getRange(2, 1, outRows, urgentCols);

    template.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    const dv = template.getDataValidations();
    if (dv && dv[0]) {
      const dvs = Array.from({ length: outRows }, () => dv[0].slice());
      dest.setDataValidations(dvs);
    }
  } catch (_) {}
}

function getMainRowFromUrgent_(urgentSheet, urgentRow, idToRow) {
  const cell = urgentSheet.getRange(urgentRow, 1);
  const id = parseRowIdFromNote_(cell.getNote());
  if (id && idToRow[id]) return idToRow[id];
  return parseMainRowNumber_(cell.getDisplayValue());
}

function parseMainRowNumber_(v) {
  const s = String(v ?? '').trim();
  if (/^\d+$/.test(s)) return parseInt(s, 10);
  const m = s.match(/range=[A-Za-z]+(\d+)/) || s.match(/(\d+)\s*$/);
  return m ? parseInt(m[1], 10) : null;
}
