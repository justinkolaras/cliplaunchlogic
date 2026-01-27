function getSeparatorColorFromMainRow7_(mainSheet, mainLastCol) {
  try {
    var vals = mainSheet.getRange(7, 1, 1, mainLastCol).getValues()[0];
    for (var i = 0; i < vals.length; i++) if (!isBlank_(vals[i])) return null;

    var bgs = mainSheet.getRange(7, 1, 1, mainLastCol).getBackgrounds()[0];
    var counts = {};
    for (var j = 0; j < bgs.length; j++) {
      var lc = String(bgs[j]).toLowerCase();
      if (!lc || lc === '#ffffff' || lc === '#fff') continue;
      counts[lc] = (counts[lc] || 0) + 1;
    }

    var best = null;
    var bestN = 0;
    for (var k in counts) {
      if (counts[k] > bestN) {
        bestN = counts[k];
        best = k;
      }
    }
    return best || null;
  } catch (err) {
    return null;
  }
}

function isSeparatorRow_(rowVals, bgRow, sepColor) {
  for (var i = 0; i < rowVals.length; i++) if (!isBlank_(rowVals[i])) return false;

  var target = String(sepColor).toLowerCase();
  var hits = 0;
  var total = 0;

  for (var j = 0; j < bgRow.length; j++) {
    var lc = String(bgRow[j]).toLowerCase();
    if (!lc || lc === '#ffffff' || lc === '#fff') continue;
    total++;
    if (lc === target) hits++;
  }

  return total > 0 && hits / total >= 0.6;
}
