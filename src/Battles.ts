/** set the value and style in a cell */
function spooledSetCellValue_(
  spooler: utils.Spooler,
  range: Spreadsheet.Range | utils.SpooledRange,
  value: boolean | number | string | Date,
  bold: boolean,
  align: 'left' | 'center' | 'right' = 'left',
): utils.SpooledRange {
  const spooled = range instanceof utils.SpooledRange ? range : spooler.attach(range);

  spooled
    .setFontWeight(bold ? 'bold' : 'normal')
    .setHorizontalAlignment(align)
    .setValue(value);

  return spooled;
}

/** process and output members data for current event */
function populateEventTable_(
  data: string[][],
  members: PlayerData[],
  unitsIndex: UnitDefinition[],
): Array<Array<string | number>> {
  const memberNames = Members.getNames();
  const nameToBaseId: KeyedStrings = {};
  for (const e of unitsIndex) {
    nameToBaseId[e.name] = e.baseId;
  }

  let total = 0;
  let requiredUnits = 0;
  let missingRequiredUnits = 0;
  let squadCount = 0;
  let lastSquad = 0;

  const table: Array<Array<string | number>> = [[]];

  for (let c = 0; c < memberNames.length; c += 1) {
    const m = members[c]; // weak
    table[0][c] = m.name;

    for (let r = 0; r < data.length; r += 1) {
      const curHero = data[r];

      if (table[r + 1] == null) {
        table[r + 1] = [];
      }

      if (curHero[0] === 'Phase Count:') {
        const phaseUnits = +curHero[1];
        const readyUnits = requiredUnits - missingRequiredUnits + Math.min(phaseUnits - requiredUnits, squadCount);
        table[r + 1][c] = readyUnits;
        total += readyUnits;
        requiredUnits = 0;
        missingRequiredUnits = 0;
        squadCount = 0;
        continue;
      } else if (curHero[0] === 'Total:') {
        table[r + 1][c] = total;
        total = 0;
        requiredUnits = 0;
        missingRequiredUnits = 0;
        squadCount = 0;
        continue;
      } else if (curHero[0].length === 0) {
        table[r + 1][c] = '';
        continue;
      } else {
        table[r + 1][c] = '';
      }
      const squad = +curHero[4];
      if (squad !== lastSquad) {
        requiredUnits = 0;
        missingRequiredUnits = 0;
        squadCount = 0;
      }
      lastSquad = squad;

      // Get Hero for member
      let baseId = nameToBaseId[curHero[0]];
      if (!baseId) {
        // refresh from data source
        const definitions = Units.getDefinitionsFromDataSource();
        // replace content of unitsIndex with definitions
        unitsIndex.splice(0, unitsIndex.length, ...[...definitions.heroes, ...definitions.ships]);
        // refresh nameToBaseId with updated unitsIndex
        for (const e of unitsIndex) {
          nameToBaseId[e.name] = e.baseId;
        }
        // try again... once
        baseId = nameToBaseId[curHero[0]];
      }
      const o = m.units[baseId];
      const requirementsMet = o && (o.rarity >= +curHero[1] && o.gearLevel! >= +curHero[2] && o.level >= +curHero[3]);
      const unitIsRequired = curHero[5] === 'R';
      if (unitIsRequired) {
        requiredUnits += 1;
        if (!requirementsMet) {
          missingRequiredUnits += 1;
        }
      } else if (requirementsMet) {
        squadCount += 1;
      }
      if (o) {
        table[r + 1][c] = requirementsMet ? `${o.rarity}` : `${o.rarity}*L${o.level}G${o.gearLevel}`;
      }
    }
  }
  return table;
}

/**
 * Updates the guild roster
 *
 * @param members - array of members
 * @returns array of PlayerData
 */
function updateGuildRoster_(members: PlayerData[]): PlayerData[] {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);

  const sortFunction = config.sortRoster()
    ? // sort roster by member name
      (a: PlayerData, b: PlayerData) => utils.caseInsensitive(a.name, b.name)
    : // sort roster by GP
      (a: PlayerData, b: PlayerData) => b.gp - a.gp;

  members.sort(sortFunction);

  if (members.length > MAX_MEMBERS) {
    members.splice(MAX_MEMBERS);
    SpreadsheetApp.getUi().alert(`Guild roster was truncated to the first ${MAX_MEMBERS} members.`);
  }

  // get the filter & tag
  // var POWER_TARGET = requiredHeroGp()

  // cleanup the header
  const header = [['Name', 'Ally Code', 'GP', 'GP Heroes', 'GP Ships']];

  const result = members.map((e) => [[e.name], [e.allyCode], [e.gp], [e.heroesGp], [e.shipsGp]]);

  // write the roster
  sheet.getRange(1, 2, 60, result[0].length).clearContent();
  sheet.getRange(1, 2, header.length, header[0].length).setValues(header);
  sheet.getRange(2, 2, result.length, result[0].length).setValues(result);
  SPREADSHEET.toast('Roster data updated', 'Guild roster', 3);

  return members;
}

/** compute a hash of current settings */
function getSettingsHash_() {
  const roster = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  const meta = SPREADSHEET.getSheetByName(SHEETS.META);

  // members name & ally code
  const members = (roster.getRange(2, 2, 50, 2).getValues() as Array<[string, number]>)
    .reduce((acc: Array<[string, number]>, e) => {
      if (e[1] > 0) {
        acc.push(e);
      }
      return acc;
    }, [])
    .sort((a, b) => a[1] - b[1]);

  // rename/add/remove settings
  const rar = (roster.getRange(2, 16, roster.getMaxRows(), 3).getValues() as Array<[string, number, number]>)
    .reduce((acc: Array<[string, number, number]>, e) => {
      if (e[1] > 0 || e[2] > 0) {
        acc.push(e);
      }
      return acc;
    }, [])
    .sort((a, b) => (a[1] !== b[1] ? a[1] - b[1] : a[2] - b[2]));

  // data source
  const dataSource = meta.getRange(14, 4).getValue();
  // SwgohGg settings
  const swgohGg = meta.getRange(2, 1).getValue();
  // SwgohGg settings
  const swgohHelp = meta.getRange(16, 1, 5).getValues();

  const hash = String(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify({ members, rar, dataSource, swgohGg, swgohHelp }),
    ),
  );

  return hash;
}

/** process the Rename, Add and Remove columns */
function renameAddRemove_(members: PlayerData[]): PlayerData[] {
  const ROSTER_RENAME_ADD_PLAYER_ROW = 2;
  const ROSTER_RENAME_ADD_PLAYER_COL = 16;
  const ROSTER_REMOVE_PLAYER_ROW = 2;
  const ROSTER_REMOVE_PLAYER_COL = 18;
  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  const add = sheet
    .getRange(ROSTER_RENAME_ADD_PLAYER_ROW, ROSTER_RENAME_ADD_PLAYER_COL, sheet.getLastRow(), 2)
    .getValues() as Array<[string, number]>;
  const remove = sheet
    .getRange(ROSTER_REMOVE_PLAYER_ROW, ROSTER_REMOVE_PLAYER_COL, sheet.getLastRow(), 1)
    .getValues() as number[][];

  const definitions = Units.getDefinitions();
  const unitsIndex = [...definitions.heroes, ...definitions.ships];

  // add & rename
  for (const e of add) {
    const allyCode = e[1];
    if (allyCode && allyCode > 0) {
      const name = (typeof e[0] === 'string' ? e[0] : `${e[0]}`).trim();

      const index = members.findIndex((m) => m.allyCode === allyCode);
      if (index === -1) {
        // get PlayerData and update members
        const member = Player.getFromDataSource(allyCode, unitsIndex);
        if (member) {
          members.push(member);
        } else {
          SPREADSHEET.toast(`Player allycode ${allyCode} not found`, 'Rename/Add/Remove', 3);
        }
      }
      if (index !== -1 && name.length > 0 && members[index].name !== name) {
        members[index].name = name; // rename member
      }
    }
  }

  // remove
  for (const e of remove) {
    const allyCode = (e && +e[0]) || 0;
    if (allyCode > 0) {
      const index = members.findIndex((m) => m.allyCode === allyCode);
      if (index > -1) {
        members.splice(index, 1); // remove member
      }
    }
  }

  return members;
}

/** fix members name */
function normalizeRoster_(members: PlayerData[]): PlayerData[] {
  // fix name starting with single quote
  for (const e of members) {
    if (e.name[0] === "'") {
      e.name = ` ${e.name}`;
    }
  }

  // find duplicate names and append allycode
  const index: KeyedType<number[]> = {};
  members.forEach((e, i) => {
    if (index.hasOwnProperty(e.name)) {
      index[e.name].push(i);
    } else {
      index[e.name] = [i];
    }
  });
  for (const key of Object.keys(index)) {
    const a = index[key];
    if (a.length > 1) {
      for (const i of a) {
        members[i].name += ` (${members[i].allyCode})`;
      }
    }
  }

  return members;
}

/** get a PlayerData array of members, from sheet if possible, else from data source */
function getMembers_(): PlayerData[] {
  const cds = config.dataSource;

  let members: PlayerData[] | undefined;

  const settingsHash = getSettingsHash_();

  const cacheId = SPREADSHEET.getId();
  const cache = CacheService.getScriptCache();
  const cachedHash = cache.get(cacheId);

  if (cachedHash && cachedHash === settingsHash) {
    SPREADSHEET.toast('Using cached roster data', 'Get guild members', 3);
    return Members.getFromSheet();
  }
  // Figure out which data source to use
  if (cds.isSwgohHelp()) {
    SPREADSHEET.toast(`Fetching roster data from ${cds.getDataSource()}`, 'Get guild members', 3);
    members = SwgohHelp.getGuildData();
  } else if (cds.isSwgohGg()) {
    SPREADSHEET.toast(`Fetching roster data from ${cds.getDataSource()}`, 'Get guild members', 3);
    members = SwgohGg.getGuildData(config.SwgohGgApi.guild());
  }
  if (!members) {
    throw new Error('The datasource returned no data');
  }
  cds.setGuildDataDate();

  const definitions = Units.getDefinitions();
  const unitsIndex = [...definitions.heroes, ...definitions.ships];
  const missingUnit = members.some((m: PlayerData) => {
    for (const baseId in m.units) {
      if (unitsIndex.findIndex((e) => e.baseId === baseId) === -1) {
        return true;
      }
    }

    return false;
  });

  if (missingUnit) {
    Units.getDefinitionsFromDataSource();
  }

  const seconds = 3600; // 1 hour
  cache.put(cacheId, settingsHash, seconds);
  return normalizeRoster_(renameAddRemove_(members));
}

/** setup the current event */
function setupEvent(): void {
  const event = config.currentEvent();

  [
    SHEETS.ASSIGNMENTS,
    SHEETS.GEODSPLATOONAUDIT,
    SHEETS.GEOSQUADRONAUDIT,
    SHEETS.GEONEEDEDUNITS,
    SHEETS.DSPLATOONAUDIT,
    SHEETS.LSPLATOONAUDIT,
    SHEETS.SQUADRONAUDIT,
    SHEETS.NEEDEDUNITS,
    SHEETS.BREAKDOWN,
    SHEETS.GEODSMISSIONS,
    SHEETS.DSMISSIONS,
    SHEETS.LSMISSIONS,
    SHEETS.HEROES,
    SHEETS.SHIPS,
    SHEETS.STATICSLICES,
    SHEETS.GEODSPLATOON,
    SHEETS.GEOSQUADRON,
    SHEETS.HOTHDSPLATOON,
    SHEETS.HOTHLSPLATOON,
    SHEETS.HOTHSQUADRON,
  ].forEach((e) => SPREADSHEET.getSheetByName(e).hideSheet());

  // // make sure the roster is up-to-date
  const definitions = Units.getDefinitions();
  const unitsIndex = [...definitions.heroes, ...definitions.ships];

  let members = getMembers_();

  if (!members) {
    const UI = SpreadsheetApp.getUi();
    UI.alert('Parsing Error', 'Unable to parse guild data. Check source links in Meta Tab', UI.ButtonSet.OK);

    return;
  }

  // This will update Roster Sheet with names and GPs,
  // will also return a new members array with added/deleted from sheet
  members = updateGuildRoster_(members);

  const heroesTable = new Units.Heroes();
  const shipsTable = new Units.Ships();
  heroesTable.setInstances(members);
  shipsTable.setInstances(members);

  const spooler = new utils.Spooler();

  // clear the hero data
  SPREADSHEET.toast('Rebuilding...', 'TB sheet', 3);
  const tbSheet = SPREADSHEET.getSheetByName(SHEETS.TB);
  spooler.attach(tbSheet.getRange(1, 10, 1, MAX_MEMBERS)).clearContent();
  spooler.attach(tbSheet.getRange(2, 1, tbSheet.getMaxRows() - 1, 9 + MAX_MEMBERS)).clearContent();

  type EventData = [
    string, // eventType
    string, // phase
    string, // unit
    number, // rarity
    number, // gearLevel
    number, // level
    string, // squad
    string, // required
  ];

  // collect the meta data for the heroes
  const row = 2;
  const col = isHothLS_(event)
    ? META_SQUADS_HOTHLS_COL
    : isHothDS_(event)
    ? META_SQUADS_HOTHDS_COL
    : META_SQUADS_GEODS_COL;
  const metaSheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const eventDefinition = metaSheet.getRange(row, col, metaSheet.getLastRow() - row + 1, 8).getValues() as EventData[];

  interface EventUnit {
    name: string;
    rarity: number;
    gearLevel: number;
    level: number;
    required: string;
  }

  interface EventObject {
    squad: string;
    eventType: string;
    phase: string;
    units: EventUnit[];
  }

  const events = eventDefinition
    .reduce((acc: EventObject[], def) => {
      const phase = def[1];
      const squad = def[6];
      if (typeof phase === 'string' && typeof squad === 'string' && phase.length > 0 && squad.length > 0) {
        let obj = acc.find((o) => o.phase === phase && o.squad === squad);
        if (!obj) {
          obj = {
            eventType: def[0],
            phase,
            squad,
            units: [],
          };
          acc.push(obj);
        }
        obj.units.push({
          gearLevel: def[4],
          level: def[5],
          name: def[2],
          rarity: def[3],
          required: def[7],
        });
      }
      return acc;
    }, [])
    .sort((a, b) => {
      if (a.phase < b.phase) {
        return -1;
      }
      if (a.phase > b.phase) {
        return 1;
      }
      if (a.squad < b.squad) {
        return -1;
      }
      if (a.squad > b.squad) {
        return 1;
      }
      return 0;
    });

  let tbRow = 2;
  const phaseList = [];
  let total = 0;

  for (const e of events) {
    let phaseCount = 0;
    let squadCount = 0;

    for (const u of e.units) {
      // store the meta data
      const data = [
        e.eventType,
        e.phase,
        u.name, // see following setCellValue_
        u.rarity,
        u.gearLevel,
        u.level,
        e.squad,
        u.required,
        `=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`,
      ];
      spooler.attach(tbSheet.getRange(tbRow, 1, 1, 9)).setValues([data]);
      spooledSetCellValue_(spooler, tbSheet.getRange(tbRow, 3), u.name, false, 'left');
      tbRow += 1;
      squadCount += 1;

      if (squadCount <= 5) {
        phaseCount += 1;
      }
    }

    spooledSetCellValue_(spooler, tbSheet.getRange(tbRow, 3), 'Phase Count:', true, 'right')
      .offset(0, 1)
      .setValue(Math.min(phaseCount, 5))
      .offset(0, 5)
      .setFormula(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);

    phaseList.push([e.phase, tbRow]);
    tbRow += 2;

    total += phaseCount;
  }

  const lastHeroRow = tbRow;

  // add the total
  let curTb = tbSheet.getRange(tbRow, 3);
  spooledSetCellValue_(spooler, curTb, 'Total:', true, 'right');
  spooler.attach(curTb.offset(0, 1)).setValue(total);
  spooler.attach(curTb.offset(0, 6)).setFormula(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);

  // add the readiness chart
  spooledSetCellValue_(
    spooler,
    tbSheet.getRange(tbRow + 2, 3),
    `=CONCATENATE("Guild Readiness ",FIXED(100*AVERAGE(J${tbRow}:BI${tbRow})/D${tbRow},1),"%")`,
    true,
    'center',
  );

  tbRow += 3;
  // list the phases
  phaseList.forEach((e, i) => {
    curTb = tbSheet.getRange(tbRow + i, 2);
    spooler.attach(curTb).setValue(e[0]);
    spooledSetCellValue_(spooler, curTb.offset(0, 1), `=I${e[1]}`, true, 'center');
  });

  // show the legend
  tbRow += phaseList.length + 1;
  curTb = tbSheet.getRange(tbRow, 3);
  spooledSetCellValue_(spooler, curTb, 'Legend', true, 'center');
  spooledSetCellValue_(spooler, curTb.offset(1, 0), 'Meets Requirements', false, 'center');
  spooledSetCellValue_(spooler, curTb.offset(2, 0), 'Missing 1 Gear (<= G8)', false, 'center');
  spooledSetCellValue_(spooler, curTb.offset(3, 0), 'Missing Levels', false, 'center');
  spooledSetCellValue_(spooler, curTb.offset(4, 0), 'Missing 1 Gear (> G8)', false, 'center');
  spooledSetCellValue_(spooler, curTb.offset(5, 0), 'Missing 1 Star', false, 'center');

  spooler.commit();

  // setup member columns
  SPREADSHEET.toast('Populating...', 'TB sheet', 3);
  let table = populateEventTable_(
    tbSheet.getRange(2, 3, lastHeroRow, 6).getValues() as string[][],
    members,
    unitsIndex,
  );

  // store the table of member data
  const TB_OFFSET_ROW = 1;
  const TB_OFFSET_COL = 10;
  const width = table.reduce((a: number, e) => Math.max(a, e.length), 0);
  table = table.map((e) => (e.length !== width ? [...e, ...Array(width).fill(null)].slice(0, width) : e));
  tbSheet.getRange(TB_OFFSET_ROW, TB_OFFSET_COL, table.length, table[0].length).setValues(table);

  if (isHothDS_(event)) {
    [
      SHEETS.DSPLATOONAUDIT,
      SHEETS.SQUADRONAUDIT,
      SHEETS.NEEDEDUNITS,
      SHEETS.DSMISSIONS,
      SHEETS.ESTIMATE,
      SHEETS.HEROES,
      SHEETS.SHIPS,
    ].forEach((e) => SPREADSHEET.getSheetByName(e).showSheet());
  } else if (isHothLS_(event)) {
    [
      SHEETS.LSPLATOONAUDIT,
      SHEETS.SQUADRONAUDIT,
      SHEETS.NEEDEDUNITS,
      SHEETS.LSMISSIONS,
      SHEETS.ESTIMATE,
      SHEETS.HEROES,
      SHEETS.SHIPS,
    ].forEach((e) => SPREADSHEET.getSheetByName(e).showSheet());
  } else if (isGeoDS_(event)) {
    [
      SHEETS.GEODSPLATOONAUDIT,
      SHEETS.GEOSQUADRONAUDIT,
      SHEETS.GEONEEDEDUNITS,
      SHEETS.GEODSMISSIONS,
      SHEETS.HEROES,
      SHEETS.SHIPS,
    ].forEach((e) => SPREADSHEET.getSheetByName(e).showSheet());
  }

  SPREADSHEET.toast('Ready', 'TB sheet', 3);
}
