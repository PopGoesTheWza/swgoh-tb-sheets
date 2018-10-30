/** Get the list of exclusions */
function get_exclusions_(): KeyedType<KeyedBooleans> {

  const data = SPREADSHEET.getSheetByName(SHEETS.EXCLUSIONS).getDataRange()
    .getValues() as string[][];
  const filtered = data.reduce(
    (acc: string[][], e) => {
      if (e[0].length > 0) {
        acc.push(e.slice(0, MAX_PLAYERS + 1));
      }
      return acc;
    },
    [],
  );

  const exclusions: KeyedType<KeyedBooleans> = {};

  const players = filtered.shift();  // First row holds player names
  players.shift();  // drop first column

  // For each unit rows
  for (const e of filtered) {
    const unit = e.shift();  // first column is unit names

    // For each exclusion cell
    e.forEach((x, c) => {
      const player = players[c];
      const isExcluded = Boolean(x ? x.trim() : '');  // exclude if cell is not empty?
      if (isExcluded) {
        if (!exclusions[player]) {
          exclusions[player] = {};
        }
        exclusions[player][unit] = isExcluded;
      }
    });
  }

  return exclusions;
}

/** Process all the excluded units */
function processExclusions_(
  data: KeyedType<UnitInstances>,
  exclusions: KeyedType<KeyedBooleans>,
) {
  for (const player in exclusions) {
    const units = exclusions[player];
    for (const unit in units) {
      if (units[unit] && data[unit] && data[unit][player]) {
        delete data[unit][player];
      }
    }
  }

  return data;
}
