/** Get the list of exclusions */
function get_exclusions_(): {[key: string]: {[key: string]: boolean}} {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.EXCLUSIONS);
  const data = sheet.getDataRange()
    .getValues() as string[][];
  const filtered = data.reduce(
    (acc: string[][], e) => {
      if (e[0].length > 0) {
        acc.push(e.slice(0, MAX_PLAYERS));
      }
      return acc;
    },
    [],
  );

  const exclusions: {[key: string]: {[key: string]: boolean}} = {};

  const players = filtered.shift();  // First row holds player names
  players.shift();  // drop first column

  // For each unit rows
  for (const e of filtered) {
    const unitName = e.shift();  // first column is unit names
    exclusions[unitName] = {};

    // For each exclusion cell
    e.forEach((x, c) => {
      const player = players[c];
      const isExcluded = Boolean(x ? x.trim() : '');  // exclude if cell is not empty?
      exclusions[unitName][player] = isExcluded;
    });
  }

  return exclusions;
}

/** Process all the excluded units */
function processExclusions_(
  data: string[][],
  exclusions: {[key: string]: {[key: string]: boolean}},
) {
  const maxRow = data.length;
  for (let row = 1; row < maxRow; row += 1) {
    const maxCol = data[row].length;
    for (let col = HERO_PLAYER_COL_OFFSET; col < maxCol; col += 1) {
      try {
        const unitName = data[row][0];
        const playerName = data[0][col];
        if (exclusions[unitName] && exclusions[unitName][playerName]) {
          // clear the unit's data
          data[row][col] = '';
        }
      } catch (e) {}
    }
  }

  return data;
}
