/** Get the list of exclusions */
function get_exclusions_(): boolean[][] {
  const excludeSheet = SPREADSHEET.getSheetByName(SHEETS.EXCLUSIONS);
  const excludeData = excludeSheet.getDataRange()
    .getValues() as string[][];
  const filtered = excludeData.filter(e => e[0].length > 0)
    .map(e => e.slice(0, MAX_PLAYERS)) as string[][];

  const excludedUnits: boolean[][] = [];

  const players = filtered.shift();  // First row is player names
  players.shift();  // drop first column

  // For each unit rows
  filtered.forEach((e) => {
    const unitName = e.shift();  // first column is unit names
    excludedUnits[unitName] = [];

    // For each exclusion cell
    e.forEach((x, c) => {
      const player = players[c];
      const isExcluded = Boolean(x ? x.trim() : '');  // exclude if cell is not empty?
      excludedUnits[unitName][player] = isExcluded;
    });
  });

  return excludedUnits;
}

/** Process all the excluded units */
function processExclusions(data, excludeData) {
  /*
  // First row is player names
  var players = data.slice().shift()
  // drop first column
  players.shift()

  data.forEach(function(e, i, a) {
  })
*/

  for (let r = 1, rLen = data.length; r < rLen; r += 1) {
    for (let c = HERO_PLAYER_COL_OFFSET, cLen = data[r].length; c < cLen; c += 1) {
      try {
        const heroName = data[r][0];
        const playerName = data[0][c];
        if (
          excludeData[heroName] != null &&
          excludeData[heroName][playerName]
        ) {
          // clear the unit's data
          data[r][c] = '';
        }
      } catch (e) {}
    }
  }

  return data;
}
