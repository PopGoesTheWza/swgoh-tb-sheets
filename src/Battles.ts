// ****************************************
// TB functions
// ****************************************

declare function getGuildDataFromSwgohHelp(): PlayerData[];

// set the value and style in a cell
function set_cell_value_(
  cell: GoogleAppsScript.Spreadsheet.Range,
  value,
  bold: boolean,
  align?: 'left' | 'center' | 'right',
): void {
  cell.setFontWeight(bold ? 'bold' : 'normal')
    .setHorizontalAlignment(align)
    .setValue(value);
}

function populateTBTable_(
  data: string[][],
  members: PlayerData[],
  heroes: UnitDeclaration[],
): (string|number)[][] {
  const roster = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 1)
    .getValues() as [string][];
  const hIdx = [];
  for (const e of heroes) {
    hIdx[e.name] = e.baseId;
  }
  let total = 0;
  let phaseCount = 0;
  let squadCount = 0;
  let lastSquad = 0;
  let lastRequired = false;

  const table: (string|number)[][] = [];
  table[0] = [];

  for (let c = 0; c < roster.length; c += 1) {
    // const m = members[roster[c]]
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
      const o = m['units'][hIdx[curHero[0]]];
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

// function add_missing_members_(
//   result: PlayerData[],
//   addMembers: {[key: number]: string},
// ): PlayerData[] {
//   // for each member to add
//   for (const key in addMembers) {
//     if (addMembers.hasOwnProperty(key)) {
//       // add
//     }
//   }
//   // addMembers.filter(e => e[0].trim().length > 0)  // it must have a name
//   //   .map(e => [e[0], forceHttps(e[1])])  // the url must use TLS
//   //   // it must be unique. make sure the player's link isn't already in the list
//   //   .filter(e => !result.some(l => l[1] === e[1]))
//   //   // add member to the list. TODO: added members lack the gp information
//   //   .forEach(e => result.push([e[0], e[1], 0, 0, 0]));

//   return result;
// }

// TODO: use allycode instead of url
// function remove_members_(members: PlayerData[], removeMembers: [string][]): PlayerData[] {
//   const result: string[] = [];
//   members.forEach((m) => {
//     if (!should_remove_(m, removeMembers)) {
//       result.push(m);
//     }
//   });

//   return result;
// }

/**
 * Update the Guild Roster
 *
 * @return The Roster sheet is updated.
 * @customfunction
 */
function updateGuildRoster_(members: PlayerData[]): PlayerData[] {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  // let results = members;

  // get the list of members to add and remove
  // const addMembers = sheet.getRange(2, META_ADD_PLAYER_COL, MAX_PLAYERS, 2)
  //   .getValues() as [string, number][];
  // const addMemberList: {[key: number]: string} = {};
  // addMembers.forEach((e) => {
  //   const allyCode: number = e[1];
  //   if (allyCode && typeof allyCode === 'number' && allyCode > 0) {
  //     if (!addMemberList.hasOwnProperty(allyCode)) {
  //       const name: string = String(e[0]).trim();
  //       addMemberList[allyCode] = name;
  //     }
  //   }
  // });
  // // add missing members
  // results = add_missing_members_(results, addMemberList);

  // const removeMembers = sheet.getRange(2, META_REMOVE_PLAYER_COL, MAX_PLAYERS, 1)
  //   .getValues() as [number][];
  // const removeMemberList: {[key: number]: any} = {};
  // removeMembers.forEach((e) => {
  //   const allyCode: number = e[0];
  //   if (allyCode && typeof allyCode === 'number' && allyCode > 0) {
  //     if (!removeMemberList.hasOwnProperty(allyCode)) {
  //       removeMemberList[allyCode] = undefined;
  //     }
  //   }
  // });
  // results = remove_members_(results, removeMemberList); // TODO, Remove members from array

  const sortFunction = getSortRoster_()
    // sort roster by player name
    ? (a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase())
    // sort roster by GP
    : (a, b) => b.gp - a.gp;

  members.sort(sortFunction);

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
function setupTBSide(): void {
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
    members = getGuildDataFromSwgohHelp();
  } else if (isDataSourceSwgohGg_()) {
    members = getGuildDataFromSwgohGg_(getSwgohGgGuildId_());
    // TODO: enrich with units name and tags
  } else {
    // members = getGuildDataFromScorpio();
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
  const tagFilter = getSideFilter_();
  const col = isLight_(tagFilter) ? META_HEROES_COL : META_HEROES_DS_COL;
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
        set_cell_value_(curTb, 'Phase Count:', true, 'right');
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
    set_cell_value_(curTb.offset(0, 2), name, false, 'left');
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
  set_cell_value_(curTb, 'Phase Count:', true, 'right');
  curTb.offset(0, 1).setValue(Math.min(phaseCount, 5));
  curTb
    .offset(0, 6)
    .setValue(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);
  phaseList.push([phase, tbRow]);

  total += phaseCount;
  tbRow += 1;
  // add the total
  curTb = tbSheet.getRange(tbRow, 3);
  set_cell_value_(curTb, 'Total:', true, 'right');
  curTb.offset(0, 1).setValue(total);
  curTb
    .offset(0, 6)
    .setValue(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);

  // add the readiness chart
  set_cell_value_(
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
    set_cell_value_(curTb.offset(0, 1), `=I${e[1]}`, true, 'center');
  });

  // show the legend
  tbRow += phaseList.length + 1;
  curTb = tbSheet.getRange(tbRow, 3);
  set_cell_value_(curTb, 'Legend', true, 'center');
  set_cell_value_(curTb.offset(1, 0), 'Meets Requirements', false, 'center');
  set_cell_value_(curTb.offset(2, 0), 'Missing 1 Gear (<= G8)', false, 'center');
  set_cell_value_(curTb.offset(3, 0), 'Missing Levels', false, 'center');
  set_cell_value_(curTb.offset(4, 0), 'Missing 1 Gear (> G8)', false, 'center');
  set_cell_value_(curTb.offset(5, 0), 'Missing 1 Star', false, 'center');

  // setup player columns
  let table = [];
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
  table = populateTBTable_(
    tbSheet.getRange(2, 3, lastHeroRow, 6).getValues() as string[][],
    members,
    heroes,
  );

  // store the table of player data
  // tbSheet.getRange(1, HERO_PLAYER_COL_OFFSET + 1,
  //                  table.length, table[0].length).setValues(table);
  const width = table.reduce((a: number, e: string) => Math.max(a, e.length), 0);
  table = table.map(
    e =>
      e.length !== width ? e.concat(Array(width).fill(null)).slice(0, width) : e,
  );
  tbSheet
    .getRange(1, META_TB_COL_OFFSET, table.length, table[0].length)
    .setValues(table);
}
