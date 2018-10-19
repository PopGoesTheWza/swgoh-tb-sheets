// ****************************************
// Ship Functions
// ****************************************

/** Populate the Ships list with Member data */
function populateShipsList(members) {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);

  // Build a Ship Index by BaseID
  const baseIDs = sheet
    .getRange(2, 2, getShipCount_(), 1)
    .getValues() as string[][];
  const hIdx = [];
  baseIDs.forEach((e, i) => {
    hIdx[e[0]] = i;
  });

  // Build a Member index by Name
  const mList = SPREADSHEET
    .getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 1)
    .getValues() as [string][];
  const pIdx = [];
  mList.forEach((e, i) => {
    pIdx[e[0]] = i;
  });
  const mHead = [];
  mHead[0] = [];

  // Clear out our old data, if any, including names as order may have changed
  sheet.getRange(1, SHIP_PLAYER_COL_OFFSET, baseIDs.length, MAX_PLAYERS)
    .clearContent();

  // This will hold all our data
  // Initialize our data
  const data = baseIDs.map(e => Array(mList.length).fill(undefined));

  for (const m of members) {
    mHead[0].push(m.name);
    const units = m.units;
    for (const e of baseIDs) {
      const baseId = e[0];
      const u = units[baseId];
      data[hIdx[baseId]][pIdx[m.name]] = (u && `${u.rarity}*L${u.level}P${u.power}`) || '';
    }
  }
  // Write our data
  sheet.getRange(1, SHIP_PLAYER_COL_OFFSET, 1, mList.length).setValues(mHead);
  sheet
    .getRange(2, SHIP_PLAYER_COL_OFFSET, baseIDs.length, mList.length)
    .setValues(data);
}

/** Initialize the list of ships */
function updateShipsList(ships: UnitDeclaration[]): void {

  // update the sheet
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);

  // clear the old content
  sheet.getRange(1, 1, 300, SHIP_PLAYER_COL_OFFSET - 1).clearContent();

  const result = ships.map((e, i) => {
    const hMap = [e.name, e.baseId, e.tags];

    // insert the star count formulas
    const row = i + 2;
    const rangeText = `$J${row}:$BI${row}`;

    {
      [2, 3, 4, 5, 6, 7].forEach((stars) => {
        const formula =
          `=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(${rangeText},"([${stars}-7]+)\\*")))))`;
        hMap.push(formula);
      });
    }

    // insert the needed count
    const formula = `=COUNTIF({${SHEETS.PLATOONS}!$D$2:$D$16,${SHEETS.PLATOONS}!$H$2:$H$16,
      ${SHEETS.PLATOONS}!$L$2:$L$16,${SHEETS.PLATOONS}!$P$2:$P$16,
      ${SHEETS.PLATOONS}!$T$2:$T$16,${SHEETS.PLATOONS}!$X$2:$X$16},A${row})`;

    hMap.push(formula);

    return hMap;
  });
  const header = [];
  header[0] = [
    'Ship',
    'Base ID',
    'Tags',
    2,
    3,
    4,
    5,
    6,
    7,
    '=CONCAT("# Needed P",Platoon!A2)',
  ];
  sheet.getRange(1, 1, 1, header[0].length).setValues(header);
  sheet.getRange(2, 1, result.length, SHIP_PLAYER_COL_OFFSET - 1).setValues(result);

  return;
}
