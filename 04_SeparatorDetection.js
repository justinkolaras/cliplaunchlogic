'use strict';

function getSeparatorColorFromMainRow7_(mainSheet, mainLastCol) {
  try {
    const vals = mainSheet.getRange(7, 1, 1, mainLastCol).getValues()[0];
    for (let i = 0; i < vals.length; i++) {
      if (!isBlank_(vals[i])) return null;
    }

    const bgs = mainSheet.getRange(7, 1, 1, mainLastCol).getBackgrounds()[0];
    const counts = Object.create(null);

    for (let j = 0; j < bgs.length; j++) {
      const lc = String(bgs[j] ?? '').toLowerCase();
      if (!lc || lc === '#ffffff' || lc === '#fff') continue;
      counts[lc] = (counts[lc] ?? 0) + 1;
    }

    let best = null;
    let bestN = 0;

    for (const [color, n] of Object.entries(counts)) {
      if (n > bestN) {
        bestN = n;
        best = color;
      }
    }

    return best || null;
  } catch (_) {
    return null;
  }
}

function isSeparatorRow_(rowVals, bgRow, sepColor) {
  for (let i = 0; i < rowVals.length; i++) {
    if (!isBlank_(rowVals[i])) return false;
  }

  const target = String(sepColor ?? '').toLowerCase();
  let hits = 0;
  let total = 0;

  for (let j = 0; j < bgRow.length; j++) {
    const lc = String(bgRow[j] ?? '').toLowerCase();
    if (!lc || lc === '#ffffff' || lc === '#fff') continue;
    total++;
    if (lc === target) hits++;
  }

  return total > 0 && hits / total >= 0.6;
}
