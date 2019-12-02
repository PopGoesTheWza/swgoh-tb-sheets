/** (Spooler) Set the value and style in a cell */
function spooledSetCellValue_(
  spoolerInstance: utils.Spooler,
  range: Spreadsheet.Range | utils.SpooledRange,
  value: boolean | number | string | Date,
  isBold: boolean,
  align: 'left' | 'center' | 'right' = 'left',
): utils.SpooledRange {
  const spooledRange = range instanceof utils.SpooledRange ? range : spoolerInstance.attach(range);

  spooledRange
    .setFontWeight(isBold ? 'bold' : 'normal')
    .setHorizontalAlignment(align)
    .setValue(value);

  return spooledRange;
}

/** Process and output members data for current event */
function populateEventTable_(
  eventLines: string[][],
  members: PlayerData[],
  unitDefinitions: UnitDefinition[],
): Array<Array<string | number>> {
  const unitNameToBaseId: KeyedStrings = {};
  for (const unitDefinition of unitDefinitions) {
    unitNameToBaseId[unitDefinition.name] = unitDefinition.baseId;
  }

  const getBaseId = (unitname: string) => {
    let result = unitNameToBaseId[unitname];
    if (!result) {
      // refresh from data source and refresh unitDefinitions & unitNameToBaseId
      const definitions = Units.getDefinitionsFromDataSource();
      unitDefinitions
        .splice(0, unitDefinitions.length, ...[...definitions.heroes, ...definitions.ships])
        .forEach((e) => (unitNameToBaseId[e.name] = e.baseId));
      result = unitNameToBaseId[unitname];
    }
    return result;
  };

  let total = 0;
  let requiredUnits = 0;
  let missingRequiredUnits = 0;
  let squadCount = 0;
  let lastSquad = 0;

  const resetCounts = () => {
    requiredUnits = 0;
    missingRequiredUnits = 0;
    squadCount = 0;
  };

  const table: Array<Array<string | number>> = [[]];

  members.forEach((member, memberIndex) => {
    table[0][memberIndex] = member.name;

    eventLines.forEach((eventLine, eventLinesIndex) => {
      const event = {
        gear: +eventLine[2],
        level: +eventLine[3],
        power: +eventLine[4]	,
        rarity: +eventLine[1],
        required: eventLine[6],
        squad: +eventLine[5],
        unitname: eventLine[0],
      };

      const tableIndex = eventLinesIndex + 1;
      if (!Array.isArray(table[tableIndex])) {
        table[tableIndex] = [];
      }
      const tableRow = table[tableIndex];

      if (event.unitname === 'Phase Count:') {
        const phaseUnits = event.rarity; // incontext, # of required units
        const readyUnits = requiredUnits - missingRequiredUnits + Math.min(phaseUnits - requiredUnits, squadCount);
        tableRow[memberIndex] = readyUnits;
        total += readyUnits;
        resetCounts();
      } else if (event.unitname === 'Total:') {
        tableRow[memberIndex] = total;
        total = 0;
        resetCounts();
      } else if (event.unitname.length === 0) {
        tableRow[memberIndex] = '';
      } else {
        if (lastSquad !== event.squad) {
          lastSquad = event.squad;
          resetCounts();
        }

        const unitInstance = member.units[getBaseId(event.unitname)];
        const requirementsMet = unitInstance &&
          (unitInstance.rarity >= event.rarity &&
            unitInstance.gearLevel! >= event.gear &&
            unitInstance.level >= event.level &&
            unitInstance.power >= event.power);
        if (event.required === 'R') {
          requiredUnits += 1;
          if (!requirementsMet) {
            missingRequiredUnits += 1;
          }
        } else if (requirementsMet) {
          squadCount += 1;
        }
        tableRow[memberIndex] = unitInstance
          ? requirementsMet
            ? `${unitInstance.rarity}`
            : `${unitInstance.rarity}*L${unitInstance.level}G${unitInstance.gearLevel}P${unitInstance.power}`
          : '';
      }
    });
  });
  return table;
}

/**
 * Updates the guild roster
 *
 * @param members - array of members
 * @returns array of PlayerData
 */
function updateGuildRoster_(members: PlayerData[]): PlayerData[] {
  const sheet = utils.getSheetByNameOrDie(SHEET.ROSTER);
  members.sort(
    config.sortRoster()
      ? // sort roster by member name
        (a: PlayerData, b: PlayerData) => utils.caseInsensitive(a.name, b.name)
      : // sort roster by GP
        (a: PlayerData, b: PlayerData) => b.gp - a.gp,
  );

  if (members.length > MAX_MEMBERS) {
    members.splice(MAX_MEMBERS);
    SpreadsheetApp.getUi().alert(`Guild roster was truncated to the first ${MAX_MEMBERS} members.`);
  }

  // cleanup the header
  const header = [['Name', 'Ally Code', 'GP', 'GP Heroes', 'GP Ships']];
  const result = members.map((e) => [[e.name], [e.allyCode], [e.gp], [e.heroesGp], [e.shipsGp]]);

  // write the roster
  sheet
    .getRange(1, 2, 60, result[0].length)
    .clearContent()
    .offset(0, 0, header.length, header[0].length)
    .setValues(header)
    .offset(1, 0, result.length, result[0].length)
    .setValues(result);
  // sheet.getRange(1, 2, header.length, header[0].length).setValues(header);
  // sheet.getRange(2, 2, result.length, result[0].length).setValues(result);
  SPREADSHEET.toast('Roster data updated', 'Guild roster', 3);
  return members;
}

/** compute a hash of current settings */
function getSettingsHash_() {
  const roster = utils.getSheetByNameOrDie(SHEET.ROSTER);
  const meta = utils.getSheetByNameOrDie(SHEET.META);
  /** members name & ally code */
  const members = (roster.getRange(2, 2, 50, 2).getValues() as Array<[string, number]>)
    .filter((e) => e[1] > 0)
    .sort((a, b) => a[1] - b[1]);

  /** rename/add/remove settings */
  const rar = (roster.getRange(2, 16, roster.getMaxRows(), 3).getValues() as Array<[string, number, number]>)
    .filter((e) => e[1] > 0 || e[2] > 0)
    .sort((a, b) => (a[1] !== b[1] ? a[1] - b[1] : a[2] - b[2]));

  const hash = String(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify({
        dataSource: /** data source */ meta.getRange(14, 4).getValue(),
        members,
        rar,
        swgohGg: /** SwgohGg settings */ meta.getRange(2, 1).getValue(),
        swgohHelp: /** SwgohHelp settings */ meta.getRange(16, 1, 5).getValues(),
      }),
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
  const sheet = utils.getSheetByNameOrDie(SHEET.ROSTER);
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
  const cachedHash = cache!.get(cacheId);

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

  SPREADSHEET.toast(`Processing data from ${cds.getDataSource()}`, 'Get guild members', 3);

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
  cache!.put(cacheId, settingsHash, seconds);
  return normalizeRoster_(renameAddRemove_(members));
}

/** setup the current event */
function setupEvent(): void {
  const event = config.currentEvent();
  const currentSheet = utils.getActiveSheet();

  // // make sure the roster is up-to-date
  const definitions = Units.getDefinitions();
  const unitsIndex = [...definitions.heroes, ...definitions.ships];

  let members = getMembers_();

  if (!members) {
    const UI = SpreadsheetApp.getUi();
    UI.alert('Parsing Error', 'Unable to parse guild data. Check source links in Meta Tab', UI.ButtonSet.OK);

    return;
  }

  [
    SHEET.BREAKDOWN,
    SHEET.GEONEEDEDUNITS,
    SHEET.HOTHNEEDEDUNITS,
    SHEET.GEOSQUADRONAUDIT,
    SHEET.HOTHSQUADRONAUDIT,
    SHEET.GEODSPLATOONAUDIT,
    SHEET.GEOLSPLATOONAUDIT,
    SHEET.HOTHDSPLATOONAUDIT,
    SHEET.HOTHLSPLATOONAUDIT,
    SHEET.TB,
    SHEET.ASSIGNMENTS,
    SHEET.GEODSMISSIONS,
    SHEET.GEOLSMISSIONS,
    SHEET.HOTHDSMISSIONS,
    SHEET.HOTHLSMISSIONS,
    SHEET.HEROES,
    SHEET.SHIPS,
    SHEET.STATICSLICES,
    SHEET.GEODSPLATOON,
    SHEET.GEOLSPLATOON,
    SHEET.GEOSQUADRON,
    SHEET.HOTHDSPLATOON,
    SHEET.HOTHLSPLATOON,
    SHEET.HOTHSQUADRON,
  ].forEach((e) => utils.getSheetByNameOrDie(e).hideSheet());

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
  const metaSheet = utils.getSheetByNameOrDie(SHEET.META);
  const tbInfoSheet = utils.getSheetByNameOrDie(SHEET.TBINFO);
  const tbSheet = utils.getSheetByNameOrDie(SHEET.TB);

  spooler.attach(tbSheet.getRange(1, 11, 1, MAX_MEMBERS)).clearContent();
  spooler.attach(tbSheet.getRange(2, 1, tbSheet.getMaxRows() - 1, 10 + MAX_MEMBERS)).clearContent();

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

  const col = tbInfoSheet.getRange(3, 21).getValue();
  const eventDefinition = metaSheet.getRange(row, col, metaSheet.getLastRow() - row + 1, tbInfoSheet.getRange(3, 22).getValue()).getValues();

  interface EventUnit {
    name: string;
    rarity: number;
    gearLevel: number;
    level: number;
    power: number;
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
      const squad = def[7];
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
          power: def[6],
          required: def[8],
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
        u.power,
        e.squad,
        u.required,
        `=COUNTIF(K${tbRow}:BJ${tbRow},CONCAT(">=",D${tbRow}))`,
      ];
      spooler.attach(tbSheet.getRange(tbRow, 1, 1, 10)).setValues([data]);
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
      .offset(0, 6)
      .setFormula(`=COUNTIF(K${tbRow}:BJ${tbRow},CONCAT(">=",D${tbRow}))`);

    phaseList.push([e.phase, tbRow]);
    tbRow += 2;

    total += phaseCount;
  }

  const lastHeroRow = tbRow;

  // add the total
  let curTb = tbSheet.getRange(tbRow, 3);
  spooledSetCellValue_(spooler, curTb, 'Total:', true, 'right');
  spooler.attach(curTb.offset(0, 1)).setValue(total);
  spooler.attach(curTb.offset(0, 7)).setFormula(`=COUNTIF(K${tbRow}:BJ${tbRow},CONCAT(">=",D${tbRow}))`);

  // add the readiness chart
  spooledSetCellValue_(
    spooler,
    tbSheet.getRange(tbRow + 2, 3),
    `=CONCATENATE("Guild Readiness ",FIXED(100*AVERAGE(K${tbRow}:BJ${tbRow})/D${tbRow},1),"%")`,
    true,
    'center',
  );

  tbRow += 3;
  // list the phases
  phaseList.forEach((e, i) => {
    curTb = tbSheet.getRange(tbRow + i, 2);
    spooler.attach(curTb).setValue(e[0]);
    spooledSetCellValue_(spooler, curTb.offset(0, 1), `=J${e[1]}`, true, 'center');
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
    tbSheet.getRange(2, 3, lastHeroRow, 7).getValues() as string[][],
    members,
    unitsIndex,
  );

  // store the table of member data
  const TB_OFFSET_ROW = 1;
  const TB_OFFSET_COL = 11;
  const width = table.reduce((a: number, e) => Math.max(a, e.length), 0);
  table = table.map((e) => (e.length !== width ? [...e, ...Array(width).fill(null)].slice(0, width) : e));
  tbSheet.getRange(TB_OFFSET_ROW, TB_OFFSET_COL, table.length, table[0].length).setValues(table);

  config.hideShowSheets(event);
  utils.setActiveSheet(currentSheet, SHEET.TB);

  SPREADSHEET.toast('Ready', 'TB sheet', 3);
}
