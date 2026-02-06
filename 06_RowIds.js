'use strict';

function parseRowIdFromNote_(note) {
  const m = String(note ?? '').match(/ROW_ID=([A-Za-z0-9\-]+)/);
  return m ? m[1] : null;
}

function buildMainIdIndex_(mainSheet) {
  const lastRow = mainSheet.getLastRow();
  const rowToId = {};
  const idToRow = {};
  if (lastRow < 2) return { rowToId, idToRow };

  const noteRange = mainSheet.getRange(2, 1, lastRow - 1, 1);
  const notes = noteRange.getNotes();
  const newNotes = notes.map((r) => r.slice());
  let needsWrite = false;

  for (let i = 0; i < notes.length; i++) {
    const row = i + 2;
    let id = parseRowIdFromNote_(notes[i][0]);

    if (!id) {
      id = Utilities.getUuid();
      newNotes[i][0] = (notes[i][0] ? `${notes[i][0]}\n` : '') + `${CFG.TOKENS.ROW_ID_PREFIX}${id}`;
      needsWrite = true;
    }

    rowToId[row] = id;
    idToRow[id] = row;
  }

  if (needsWrite) noteRange.setNotes(newNotes);
  return { rowToId, idToRow };
}
