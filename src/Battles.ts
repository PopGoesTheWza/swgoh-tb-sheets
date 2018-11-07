/** workaround to tslint issue of namespace scope after importing type definitions */
declare namespace SwgohHelp {

  function getGuildData(): PlayerData[];

}

/** set the value and style in a cell */
function spooledSetCellValue_(
  spooler: utils.Spooler,
  range: Range | utils.SpooledRange,
  value: boolean|number|string|Date,
  bold: boolean,
  align?: 'left' | 'center' | 'right',
): utils.SpooledRange {

  const spooled = range instanceof utils.SpooledRange ? range : spooler.attach(range);
  spooled.setFontWeight(bold ? 'bold' : 'normal')
    .setHorizontalAlignment(align)
    .setValue(value);

  return spooled;
}

/** process and output members data for current event */
function populateEventTable_(
  data: string[][],
  members: PlayerData[],
  unitsIndex: UnitDefinition[],
): (string|number)[][] {

  const memberNames = Members.getNames();
  const nameToBaseId: KeyedStrings = {};
  for (const e of unitsIndex) {
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
      let baseId = nameToBaseId[curHero[0]];
      if (!baseId) {
        // refresh from data source
        const definitions = Units.getDefinitionsFromDataSource();
        // replace content of unitsIndex with definitions
        unitsIndex.splice(0, unitsIndex.length, ...definitions.heroes.concat(definitions.ships));
        // refresh nameToBaseId with updated unitsIndex
        for (const e of unitsIndex) {
          nameToBaseId[e.name] = e.baseId;
        }
        // try again... once
        baseId = nameToBaseId[curHero[0]];
      }
      const o = m['units'][baseId];
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

/** update the guild roster */
function updateGuildRoster_(members: PlayerData[]): PlayerData[] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);

  const sortFunction = config.sortRoster()
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
  // var POWER_TARGET = requiredHeroGp()

  // cleanup the header
  const header = [['Name', 'Ally Code', 'GP', 'GP Heroes', 'GP Ships']];

  const result = members.map(e => [
    [e.name],
    // [e.link],
    [e.allyCode],
    [e.gp],
    [e.heroesGp],
    [e.shipsGp],
  ]);

  // write the roster
  sheet.getRange(1, 2, 60, result[0].length).clearContent();
  sheet.getRange(1, 2, header.length, header[0].length).setValues(header);
  sheet.getRange(2, 2, result.length, result[0].length).setValues(result);

  return members;
}

function getSettingsHash_() {
  // const roster = xxxxxx
  const roster = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  const meta = SPREADSHEET.getSheetByName(SHEETS.META);

  // members name & ally code
  const members = (roster.getRange(2, 2, 50, 2)
    .getValues() as [string, number][])
    .reduce(
      (acc: [string, number][], e) => {
        if (e[1] > 0) acc.push(e);
        return acc;
      },
      [],
    ).sort((a, b) => a[1] - b[1]);

  // rename/add/remove settings
  const rar = (roster.getRange(2, 16, roster.getMaxRows(), 3)
    .getValues() as [string, number, number][])
    .reduce(
      (acc: [string, number, number][], e) => {
        if (e[1] > 0 || e[2] > 0) acc.push(e);
        return acc;
      },
      [],
    ).sort((a, b) => a[1] !== b[1] ? a[1] - b[1] : a[2] - b[2]);

  // data source
  const dataSource = meta.getRange(14, 4).getValue();
  // SwgohGg settings
  const swgohGg = meta.getRange(2, 1).getValue();
  // SwgohGg settings
  const swgohHelp = meta.getRange(20, 1, 5).getValues();

  const hash = String(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify({ members, rar, dataSource, swgohGg, swgohHelp }),
  ));

  return hash;
}

function renameAddRemove_(members: PlayerData[]): PlayerData[] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  const add = sheet.getRange(2, META_RENAME_ADD_PLAYER_COL, sheet.getLastRow(), 2)
    .getValues() as [string, number][];
  const remove = sheet.getRange(2, META_REMOVE_PLAYER_COL, sheet.getLastRow(), 1)
    .getValues() as number[][];

  const definitions = Units.getDefinitions();
  const unitsIndex = definitions.heroes.concat(definitions.ships);

  // add & rename
  for (const e of add) {
    const allyCode = e[1];
    if (allyCode && allyCode > 0) {
      const name = ((typeof e[0] === 'string') ? e[0] : `${e[0]}`).trim();
      const index = members.findIndex(m => m.allyCode === allyCode);
      if (index === -1) {
        // get PlayerData and update members
        const member = (Player.getFromDataSource(allyCode, undefined, unitsIndex));
        if (member) {
          members.push(member);
        } else {
          // TODO: UI message, AC not found
        }
      }
      if (index !== -1 && name.length > 0 && members[index].name !== name) {
        members[index].name = name;  // rename member
      }
    }
  }

  // remove
  for (const e of remove) {
    const allyCode = e && Number(e[0]) || 0;
    if (allyCode > 0) {
      const index = members.findIndex(m => m.allyCode === allyCode);
      if (index > -1) {
        members.splice(index, 1);  // remove member
      }
    }
  }

  return members;
}

function normalizeRoster_(members: PlayerData[]): PlayerData[] {

  // fix name starting with single quote
  for (const e of members) {
    if (e.name[0] === '\'') {
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
  for (const key in index) {
    const a = index[key];
    if (a.length > 1) {
      for (const i of a) {
        members[i].name += ` (${members[i].allyCode})`;
      }
    }
  }

  return members;
}

function getMembers_(): PlayerData[] {

  let members: PlayerData[];

  const settingsHash =  getSettingsHash_();

  const cacheId = SPREADSHEET.getId();
  const cache = CacheService.getScriptCache();
  const cachedHash = cache.get(cacheId);

  if (cachedHash && cachedHash === settingsHash) {
    return Members.getFromSheet();
  }
  // Figure out which data source to use
  if (config.dataSource.isSwgohHelp()) {
    members = SwgohHelp.getGuildData();
  } else if (config.dataSource.isSwgohGg()) {
    members = SwgohGg.getGuildData(config.SwgohGg.guild());
  }

  const definitions = Units.getDefinitions();
  const unitsIndex = definitions.heroes.concat(definitions.ships);
  const missingUnit = members.some((m: PlayerData) => {
    for (const baseId in m.units) {
      if (unitsIndex.findIndex(e => e.baseId === baseId) === -1) {
        return true;
      }
    }

    return false;
  });

  if (missingUnit) {
    Units.getDefinitionsFromDataSource();
  }

  const seconds = 3600;  // 1 hour
  cache.put(cacheId, settingsHash, seconds);

  return normalizeRoster_(renameAddRemove_(members));
}

/** setup the current event */
function setupEvent(): void {

  // // make sure the roster is up-to-date
  const definitions = Units.getDefinitions();
  const unitsIndex = definitions.heroes.concat(definitions.ships);

  // Figure out which data source to use
  let members = getMembers_();

  if (!members) {
    UI.alert(
      'Parsing Error',
      'Unable to parse guild data. Check source links in Meta Tab',
      UI.ButtonSet.OK,
    );

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
  const tbSheet = SPREADSHEET.getSheetByName(SHEETS.TB);
  spooler.attach(tbSheet.getRange(1, 10, 1, MAX_PLAYERS))
    .clearContent();
  spooler.attach(tbSheet.getRange(2, 1, tbSheet.getMaxRows() - 1, 9 + MAX_PLAYERS))
    .clearContent();

  type eventData = [
    string,  // eventType
    string,  // phase
    string,  // unit
    number,  // rarity
    number,  // gearLevel
    number,  // level
    string,  // squad
    string  // required
  ];

  // collect the meta data for the heroes
  const row = 2;
  const col = isLight_(config.currentEvent()) ? META_HEROES_COL : META_HEROES_DS_COL;
  const metaSheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const eventDefinition = metaSheet.getRange(row, col, metaSheet.getLastRow() - row + 1, 8)
    .getValues() as eventData[];

  type eventUnit = {
    name: string,
    rarity: number,
    gearLevel: number,
    level: number,
    required: string,
  };

  type eventObject = {
    squad: string,
    eventType: string,
    phase: string,
    units: eventUnit[],
  };

  const events = eventDefinition.reduce(
    (acc: eventObject[], e) => {
      const phase = e[1];
      const squad = e[6];
      if (
        typeof phase === 'string'
        && typeof squad === 'string'
        && phase.length > 0
        && squad.length > 0
      ) {
        let o = acc.find(e => e.phase === phase && e.squad === squad);
        if (!o) {
          o = {
            phase,
            squad,
            eventType: e[0],
            units: [],
          };
          acc.push(o);
        }
        o.units.push({
          name: e[2],
          rarity: e[3],
          gearLevel: e[4],
          level: e[5],
          required: e[7],
        });
      }
      return acc;
    },
    [],
  )
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
        u.name,  // see following setCellValue_
        u.rarity,
        u.gearLevel,
        u.level,
        e.squad,
        u.required,
        `=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`,
      ];
      spooler.attach(tbSheet.getRange(tbRow, 1, 1, 9))
        .setValues([data]);
      spooledSetCellValue_(spooler, tbSheet.getRange(tbRow, 3), u.name, false, 'left');
      tbRow += 1;
      squadCount += 1;

      if (squadCount <= 5) {
        phaseCount += 1;
      }
    }

    const curTb = tbSheet.getRange(tbRow, 3);
    spooledSetCellValue_(spooler, tbSheet.getRange(tbRow, 3), 'Phase Count:', true, 'right');
    spooler.attach(curTb.offset(0, 1))
      .setValue(Math.min(phaseCount, 5));
    spooler.attach(curTb.offset(0, 6))
      .setFormula(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);

    phaseList.push([e.phase, tbRow]);
    tbRow += 2;

    // lastPhase = e.phase;
    total += phaseCount;
    phaseCount = 0;
  }

  const lastHeroRow = tbRow;

  // add the total
  let curTb = tbSheet.getRange(tbRow, 3);
  spooledSetCellValue_(spooler, curTb, 'Total:', true, 'right');
  spooler.attach(curTb.offset(0, 1))
    .setValue(total);
  spooler.attach(curTb.offset(0, 6))
    .setFormula(`=COUNTIF(J${tbRow}:BI${tbRow},CONCAT(">=",D${tbRow}))`);

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
    spooler.attach(curTb)
      .setValue(e[0]);
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

  // setup player columns
  let table = populateEventTable_(
    tbSheet.getRange(2, 3, lastHeroRow, 6).getValues() as string[][],
    members,
    unitsIndex,
  );

  // store the table of player data
  const width = table.reduce((a: number, e) => Math.max(a, e.length), 0);
  table = table.map(
    e =>
      e.length !== width ? e.concat(Array(width).fill(null)).slice(0, width) : e,
  );
  tbSheet.getRange(1, META_TB_COL_OFFSET, table.length, table[0].length)
    .setValues(table);
}
