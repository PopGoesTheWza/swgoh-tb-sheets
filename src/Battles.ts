// ****************************************
// TB functions
// ****************************************

declare function getGuildDataFromSwgohHelp_(): PlayerData[];

/** set the value and style in a cell */
function setCellValue_(
  cell: GoogleAppsScript.Spreadsheet.Range,
  value,
  bold: boolean,
  align?: 'left' | 'center' | 'right',
): void {

  cell.setFontWeight(bold ? 'bold' : 'normal')
    .setHorizontalAlignment(align)
    .setValue(value);
}

function populateEventTable_(
  data: string[][],
  members: PlayerData[],
  heroes: UnitDeclaration[],
): (string|number)[][] {

  const memberNames = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 1)
    .getValues() as [string][];

  const nameToBaseId: {[key: string]: string} = {};
  for (const e of heroes) {
    nameToBaseId[e.name] = e.baseId;
  }

  let total = 0;
  let phaseCount = 0;
  let squadCount = 0;
  let lastSquad = 0;
  let lastRequired = false;

  const table: (string|number)[][] = [];
  table[0] = [];

  for (let c = 0; c < memberNames.length; c += 1) {

    const m = members[c]; // weak
    table[0][c] = m.name;

    for (let r = 0; r < data.length; r += 1) {

      const curHero = data[r];

      if (table[r + 1] == null) {
        table[r + 1] = [];
      }

      if (curHero[0] === 'Phase Count:') {
        table[r + 1][c] = phaseCount;
        total += phaseCount;
        phaseCount = 0;
        squadCount = 0;
        continue;
      } else if (curHero[0] === 'Total:') {
        table[r + 1][c] = total;
        total = 0;
        phaseCount = 0;
        squadCount = 0;
        continue;
      } else if (curHero[0].length === 0) {
        table[r + 1][c] = '';
        continue;
      } else {
        table[r + 1][c] = '';
      }
      const squad = Number(curHero[4]);
      if (squad !== lastSquad) {
        squadCount = 0;
      }
      lastSquad = squad;

      // Get Hero for member
      const o = m['units'][nameToBaseId[curHero[0]]];
      if (o == null) {
        continue;
      }
      const requirementsMet =
        o.rarity >= Number(curHero[1]) &&
        o.gearLevel >= Number(curHero[2]) &&
        o.level >= Number(curHero[3]);
      if (requirementsMet) {
        if (curHero[5] === 'R' && !lastRequired) {
          squadCount = 0;
        }
        lastRequired = curHero[5] === 'R';
        squadCount += 1;
        if (squadCount <= 5) {
          phaseCount += 1;
        }
      }
      table[r + 1][c] = requirementsMet
        ? `${o.rarity}`
        : `${o.rarity}*L${o.level}G${o.gearLevel}`;
    }
  }
  return table;
}

/**
 * Update the Guild Roster
 *
 * @return The Roster sheet is updated.
 * @customfunction
 */
function updateGuildRoster_(members: PlayerData[]): PlayerData[] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  const unitsIndex = getHeroesTabIndex_().concat(getShipsTabIndex_());

  const add = sheet.getRange(2, META_ADD_PLAYER_COL, sheet.getLastRow(), 2)
    .getValues() as [string, number][];
  for (const e of add) {
    const allyCode = e[1];
    if (allyCode && allyCode > 0) {
      const name = ((typeof e[0] === 'string') ? e[0] : `${e[0]}`).trim();
      const index = members.findIndex(m => m.allyCode === allyCode);
      if (index === -1) {
        // get PlayerData and update members
        const member = (getPlayerData_SwgohGgApi_(allyCode, undefined, unitsIndex));
        if (member) {
          members.push(member);
        }
      } else if (name.length > 0 && members[index].name !== name) {
        members[index].name = name;  // rename member
      }
    }
  }

  const remove = sheet.getRange(2, META_REMOVE_PLAYER_COL, sheet.getLastRow(), 1)
    .getValues() as number[][];
  for (const e of remove) {
    const allyCode = e && Number(e[0]) || 0;
    if (allyCode > 0) {
      const index = members.findIndex(m => m.allyCode === allyCode);
      if (index > -1) {
        members.splice(index, 1);  // remove member
      }
    }
  }

  const sortFunction = getSortRoster_()
    // sort roster by player name
    ? (a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase())
    // sort roster by GP
    : (a, b) => b.gp - a.gp;

  members.sort(sortFunction);

  if (members.length > MAX_PLAYERS) {
    members.splice(MAX_PLAYERS);
    UI.alert(`Guild roster was truncated to the first ${MAX_PLAYERS} members.`);
  }

  // get the filter & tag
  // var POWER_TARGET = get_minimum_character_gp_()

  // cleanup the header
  const header = [['Name', 'Hyper Link', 'GP', 'GP Heroes', 'GP Ships']];

  const result = members.map(e => [
    [e.name],
    // [e.link],
    [e.allyCode],
    [e.gp],
    [e.heroesGp],
    [e.shipsGp],
  ]);

  // result.sort(sortFunction)

  // write the roster
  sheet.getRange(1, 2, 60, result[0].length).clearContent();
  sheet.getRange(1, 2, header.length, header[0].length).setValues(header);
  sheet.getRange(2, 2, result.length, result[0].length).setValues(result);

  return members;
}

/** Setup the Territory Battle for Hoth */
function setupEvent(): void {

  // const shipsSheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);

  // make sure the roster is up-to-date

  // Update Heroes and Ship Sheets
  // NOTE Currently not supported by Scorpio, so always using SWGOH.gg data
  let heroes: UnitDeclaration[];
  let ships: UnitDeclaration[];
  if (isDataSourceSwgohHelp_()) {
    // heroes = getHeroesFromSWGOHhelp();
    // ships = getShipsFromSWGOHhelp();
    heroes = getHeroListFromSwgohGg_();
    ships = getShipListFromSwgohGg_();
  } else {
    // TODO: re-read only if necessary
    heroes = getHeroListFromSwgohGg_();
    ships = getShipListFromSwgohGg_();
  }
  updateHeroesList(heroes);
  updateShipsList(ships);

  // Figure out which data source to use
  let members: PlayerData[];
  if (isDataSourceSwgohHelp_()) {
    members = getGuildDataFromSwgohHelp_();
  } else if (isDataSourceSwgohGg_()) {
    members = getGuildDataFromSwgohGg_(getSwgohGgGuildId_());
  }
  if (!members) {
    UI.alert(
      'Parsing Error',
      'Unable to parse guild data. Check source links in Meta Tab',
      UI.ButtonSet.OK,
      );
    return;
  }

  // TODO: relocate
  // fix name starting with single quote
  for (const e of members) {
    if (e.name[0] === '\'') {
      e.name = ` ${e.name}`;
    }
  }

  // This will update Roster Sheet with names and GPs,
  // will also return a new members array with added/deleted from sheet
  members = updateGuildRoster_(members);

  // find duplicate names and append allycode
  const index: { [key: string] : number[] } = {};
  members.forEach((e, i) => {
    if (index.hasOwnProperty(e.name)) {
      index[e.name].push(i);
    } else {
      index[e.name] = [i];
    }
  });
  for (const key in index) {
    const a = index[key];
    if (a.length > 1) {
      for (const i of a) {
        members[i].name += ` (${members[i].allyCode})`;
      }
    }
  }

  populateHeroesList(members);
  populateShipsList(members);

  // clear the hero data
  const tbSheet = SPREADSHEET.getSheetByName(SHEETS.TB);
  tbSheet.getRange(1, 10, 1, MAX_PLAYERS).clearContent();
  tbSheet.getRange(2, 1, 150, 9 + MAX_PLAYERS).clearContent();

    // collect the meta data for the heroes
  let row = 2;
  const col = isLight_(getSideFilter_()) ? META_HEROES_COL : META_HEROES_DS_COL;
  let tbRow = 2;
  let lastPhase = '1';
  let phaseCount = 0;
  let total = 0;
  let lastSquad = '0';
  let squadCount = 0;
  const metaSheet = SPREADSHEET.getSheetByName(SHEETS.META);
  let curMeta = metaSheet.getRange(row, col, 1, 8);
  let metaData = curMeta.getValues() as string[][];
  let event = metaData[0][0];
  const phaseList = [];
  let phase;

  while (event.length > 0) {

    phase = metaData[0][1];
    const name = metaData[0][2];
    const stars = metaData[0][3];
    const gear = metaData[0][4];
    const level = metaData[0][5];
    const squad = metaData[0][6];
    const required = metaData[0][7];

    // see if the phase ended
    if (lastPhase !== phase) {
      if (tbRow > 2) {
        // end the phase if this isn't the first row
        const curTb = tbSheet.getRange(tbRow, 3);
        setCellValue_(curTb, 'Phase Count:', true, 'right');
        curTb.offset(0, 1).setValue(Math.min(phaseCount, 5));
        curTb
          .offset(0, 6)
          .setValue(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);
        phaseList.push([lastPhase, tbRow]);
        tbRow += 2;
      }
      lastPhase = phase;
      total += phaseCount;
      phaseCount = 0;
    }

    // see if the squad changed
    if (lastSquad !== squad) {
      lastSquad = squad;
      squadCount = 0;
    }

    // store the meta data
    let curTb = tbSheet.getRange(tbRow, 1, 1, 9);
    const data = [];
    data[0] = [
      event,
      phase,
      name,
      stars,
      gear,
      level,
      squad,
      required,
      `=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`,
    ];

    curTb.setValues(data);
    curTb = tbSheet.getRange(tbRow, 1);
    setCellValue_(curTb.offset(0, 2), name, false, 'left');
    tbRow += 1;
    squadCount += 1;

    if (squadCount <= 5) {
      phaseCount += 1;
    }

    // get the next row
    row += 1;
    curMeta = metaSheet.getRange(row, col, 1, 8);
    metaData = curMeta.getValues() as string[][];
    event = metaData[0][0];
  }

  const lastHeroRow = tbRow;

  // add the final phase
  let curTb = tbSheet.getRange(tbRow, 3);
  setCellValue_(curTb, 'Phase Count:', true, 'right');
  curTb.offset(0, 1).setValue(Math.min(phaseCount, 5));
  curTb
    .offset(0, 6)
    .setValue(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);
  phaseList.push([phase, tbRow]);

  total += phaseCount;
  tbRow += 1;
  // add the total
  curTb = tbSheet.getRange(tbRow, 3);
  setCellValue_(curTb, 'Total:', true, 'right');
  curTb.offset(0, 1).setValue(total);
  curTb
    .offset(0, 6)
    .setValue(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);

  // add the readiness chart
  setCellValue_(
    tbSheet.getRange(tbRow + 2, 3),
    `=CONCATENATE("Guild Readiness ",FIXED(100*AVERAGE(J${tbRow}:BI${tbRow})/D${tbRow},1),"%")`,
    true,
    'center',
  );

  tbRow += 3;
  // list the phases
  phaseList.forEach((e, i) => {
    curTb = tbSheet.getRange(tbRow + i, 2);
    curTb.setValue(e[0]);
    setCellValue_(curTb.offset(0, 1), `=I${e[1]}`, true, 'center');
  });

  // show the legend
  tbRow += phaseList.length + 1;
  curTb = tbSheet.getRange(tbRow, 3);
  setCellValue_(curTb, 'Legend', true, 'center');
  setCellValue_(curTb.offset(1, 0), 'Meets Requirements', false, 'center');
  setCellValue_(curTb.offset(2, 0), 'Missing 1 Gear (<= G8)', false, 'center');
  setCellValue_(curTb.offset(3, 0), 'Missing Levels', false, 'center');
  setCellValue_(curTb.offset(4, 0), 'Missing 1 Gear (> G8)', false, 'center');
  setCellValue_(curTb.offset(5, 0), 'Missing 1 Star', false, 'center');

  // setup player columns
  let table: (string | number)[][] = [];
  // curTb = tbSheet.getRange(2, 10)
  // for (var i = 0; i < MAX_PLAYERS; ++i) {
  //   var range = tbSheet.getRange(2, 3, lastHeroRow - 1, 6).getValues()
  //   var playerData = populate_player_tb_(tbSheet, i, range, heroes)
  //   if (playerData != null) {
  //     for (var p = 0, pLen = playerData.length; p < pLen; ++p) {
  //       if (table[p] == null) {
  //         // first entry
  //         table[p] = [playerData[p]]
  //       } else {
  //         // append additional entries
  //         table[p].push(playerData[p])
  //       }
  //     }
  //   }
  // }
  table = populateEventTable_(
    tbSheet.getRange(2, 3, lastHeroRow, 6).getValues() as string[][],
    members,
    heroes,
  );

  // store the table of player data
  const width = table.reduce((a: number, e) => Math.max(a, e.length), 0);
  table = table.map(
    e =>
      e.length !== width ? e.concat(Array(width).fill(undefined)).slice(0, width) : e,
  );
  tbSheet.getRange(1, META_TB_COL_OFFSET, table.length, table[0].length)
    .setValues(table);
}
