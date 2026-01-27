function parseRowIdFromNote_(note) {
  var m = String(note || '').match(/ROW_ID=([A-Za-z0-9\-]+)/);
  return m ? m[1] : null;
}

function buildMainIdIndex_(mainSheet) {
  var lastRow = mainSheet.getLastRow();
  var rowToId = {};
  var idToRow = {};
  if (lastRow < 2) return { rowToId: rowToId, idToRow: idToRow };

  var noteRange = mainSheet.getRange(2, 1, lastRow - 1, 1);
  var notes = noteRange.getNotes();
  var newNotes = notes.map(function(r) { return r.slice(); });
  var needsWrite = false;

  for (var i = 0; i < notes.length; i++) {
    var row = i + 2;
    var id = parseRowIdFromNote_(notes[i][0]);
    if (!id) {
      id = Utilities.getUuid();
      newNotes[i][0] = (notes[i][0] ? notes[i][0] + '\n' : '') + 'ROW_ID=' + id;
      needsWrite = true;
    }
    rowToId[row] = id;
    idToRow[id] = row;
  }

  if (needsWrite) noteRange.setNotes(newNotes);
  return { rowToId: rowToId, idToRow: idToRow };
}
