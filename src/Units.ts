// ****************************************
// Guild Unit Functions
// ****************************************

/** Populate the Hero list with Member data */
function populateHeroesList(members: PlayerData[]): void {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);

  // Build a Hero Index by BaseID
  const baseIDs = sheet
    .getRange(2, 2, getCharacterCount_(), 1)
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
  sheet.getRange(1, HERO_PLAYER_COL_OFFSET, baseIDs.length, MAX_PLAYERS)
    .clearContent();

  // This will hold all our data
  // Initialize our data
  const data = baseIDs.map(e => Array(mList.length).fill(''));

  for (const m of members) {
    mHead[0].push(m.name);
    const units = m.units;
    baseIDs.forEach((e, i) => {
      const baseId = e[0];
      const u = units[baseId];
      data[hIdx[baseId]][pIdx[m.name]] =
          (u && `${u.rarity}*L${u.level}G${u.gearLevel}P${u.power}`) || '';
    });
  }
  // Logger.log(
  //   "Last Member Data: " + JSON.stringify(members[mList[mList.length - 1]])
  // )
  // for (var key in members) {
  //   if (key == "unique") {
  //     continue
  //   }
  //   var m = members[key]
  //   mHead[0][pIdx[m["name"]]] = m["name"]
  //   // Logger.log("Parsing Units for: " + m["name"])
  //   var units = m["units"]
  //   for (r = 0; r < baseIDs.length; r++) {
  //     var uKey = baseIDs[r]
  //     var u = units[uKey]
  //     if (!u) {
  //       continue
  //     } // Means player has not unlocked unit
  //     data[hIdx[uKey]][pIdx[m["name"]]] =
  //       u["rarity"] +
  //       "*L" +
  //       u["level"] +
  //       "G" +
  //       u["gearLevel"] +
  //       "P" +
  //       u["power"]
  //   }
  //   // Logger.log("Parsed " + r + " units.")
  // }
  // Write our data
  sheet.getRange(1, HERO_PLAYER_COL_OFFSET, 1, mList.length).setValues(mHead);
  sheet
    .getRange(2, HERO_PLAYER_COL_OFFSET, baseIDs.length, mList.length)
    .setValues(data);
}

/** Initialize the list of heroes */
function updateHeroesList(heroes: UnitDeclaration[]): void {

  // update the sheet
  const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);

  const result = heroes.map((e, i) => {
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

    // insdert the needed count
    const formula = `=COUNTIF({${SHEETS.PLATOONS}!$D$20:$D$34,${SHEETS.PLATOONS}!$H$20:$H$34,
      ${SHEETS.PLATOONS}!$L$20:$L$34,${SHEETS.PLATOONS}!$P$20:$P$34,
      ${SHEETS.PLATOONS}!$T$20:$T$34,${SHEETS.PLATOONS}!$X$20:$X$34,
      ${SHEETS.PLATOONS}!$D$38:$D$52,${SHEETS.PLATOONS}!$H$38:$H$52,
      ${SHEETS.PLATOONS}!$L$38:$L$52,${SHEETS.PLATOONS}!$P$38:$P$52,
      ${SHEETS.PLATOONS}!$T$38:$T$52,${SHEETS.PLATOONS}!$X$38:$X$52},A${row})`;

    hMap.push(formula);

    return hMap;
  });
  // write our Units Header
  const header = [];
  header[0] = [
    'Hero',
    'Base ID',
    'Tags',
    '2',
    '3',
    '4',
    '5',
    '6',
    '7',
    '=CONCAT("# Needed P",Platoon!A2)',
  ];
  sheet.getRange(1, 1, 1, header[0].length).setValues(header);
  sheet.getRange(2, 1, result.length, HERO_PLAYER_COL_OFFSET - 1).setValues(result);

  return;
}
