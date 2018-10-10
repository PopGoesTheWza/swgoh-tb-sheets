// ****************************************
// TB functions
// ****************************************

// set the value and style in a cell
function set_cell_value_(
  cell: GoogleAppsScript.Spreadsheet.Range,
  value,
  bold: boolean,
  align?: 'left' | 'center' | 'right',
) {
  cell
    .setFontWeight(bold ? 'bold' : 'normal')
    .setHorizontalAlignment(align)
    .setValue(value);
}

function populateTBTable(data, members, heroes) {
  const roster = SPREADSHEET
    .getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 1)
    .getValues() as string[][];
  const hIdx = [];
  heroes.forEach(e => (hIdx[e.UnitName] = e.UnitId));
  let total = 0;
  let phaseCount = 0;
  let squadCount = 0;
  let lastSquad = 0;
  let lastRequired = false;

  const table = [];
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
      }
      const squad = curHero[4];
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
        o.gear_level >= Number(curHero[2]) &&
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
        ? `${o.rarity}*`
        : `${o.rarity}*L${o.level}G${o.gear_level}`;
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
function updateGuildRoster(members) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  // get the list of members to add and remove
  const addMembers = sheet
    .getRange(2, META_ADD_PLAYER_COL, MAX_PLAYERS, 2)
    .getValues();
  const removeMembers = sheet
    .getRange(2, META_REMOVE_PLAYER_COL, MAX_PLAYERS, 1)
    .getValues();

  // members = remove_members_(text, removeMembers); // TODO, Remove members from array
  // add missing members
  //  result = add_missing_members_(result, addMembers); // TODO Add members via SWGOH links

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
    [e.link],
    [e.gp],
    [e.heroes_gp],
    [e.ships_gp],
  ]);

  // result.sort(sortFunction)

  // write the roster
  sheet.getRange(1, 2, 60, result[0].length).clearContent();
  sheet.getRange(1, 2, header.length, header[0].length).setValues(header);
  sheet.getRange(2, 2, result.length, result[0].length).setValues(result);

  return members;
}

/** Setup the Territory Battle for Hoth */
function setupTBSide() {
  // const shipsSheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);

  // make sure the roster is up-to-date

  // Update Heroes and Ship Sheets
  // NOTE Currently not supported by Scorpio, so always using SWGOH.gg data
  let heroes: UnitDeclaration[];
  let ships: UnitDeclaration[];
  if (isDataSourceSwgohHelp()) {
    // heroes = getHeroesFromSWGOHhelp();
    // ships = getShipsFromSWGOHhelp();
    heroes = getHeroesFromSWGOHgg();
    ships = getShipsFromSWGOHgg();
  } else {
    heroes = getHeroesFromSWGOHgg();
    ships = getShipsFromSWGOHgg();
  }
  updateHeroesList(heroes);
  updateShipsList(ships);

  // Figure out which data source to use
  let members: PlayerData[];
  if (isDataSourceSwgohHelp()) {
    members = getGuildDataFromSwgohHelp();
  } else if (isDataSourceSwgohGg()) {
    members = getGuildDataFromSwgohGg();
  } else {
    members = getGuildDataFromScorpio();
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
  members.forEach((e) => {
    if (e.name[0] === '\'') {
      e.name = ` ${e.name}`;
    }
  });

  // This will update Roster Sheet with names and GPs,
  // will also return a new members array with added/deleted from sheet
  members = updateGuildRoster(members);

  populateHeroesList(members);
  populateShipsList(members);

  // clear the hero data
  const tbSheet = SPREADSHEET.getSheetByName(SHEETS.TB);
  tbSheet.getRange(1, 10, 1, MAX_PLAYERS).clearContent();
  tbSheet.getRange(2, 1, 150, 9 + MAX_PLAYERS).clearContent();

    // collect the meta data for the heroes
  let row = 2;
  const tagFilter = getTagFilter_();
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
        curTb.offset(0, 1).setValue(phaseCount);
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
  curTb.offset(0, 1).setValue(phaseCount);
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
  table = populateTBTable(
    tbSheet.getRange(2, 3, lastHeroRow, 6).getValues(),
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
