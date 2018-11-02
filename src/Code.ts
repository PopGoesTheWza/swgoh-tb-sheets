/** workaround to tslint issue of namespace scope after importingtype definitions */
declare function getPlayerDataFromSwgohHelp_(allyCode: number): PlayerData;

/**
 * @OnlyCurrentDoc
 */

/** string[] callback for case insensitive alphabetical sort */
function caseInsensitive_(a: string, b: string): number {
  return a.toLowerCase().localeCompare(b.toLowerCase());
}

/** [string, *][] callback for case insensitive alphabetical sort */
function firstElementCaseInsensitive_(a: [string, undefined], b: [string, undefined]): number {
  return a[0].toLowerCase().localeCompare(b[0].toLowerCase());
}

/** retrieve player's data from unit tabs or a data source */
function getSnapshopData_(
  sheet: Sheet,
  filter: string,
  unitsIndex: UnitDefinition[],
): PlayerData {

  const members = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 2)
    .getValues() as [string, number][];

  const memberName = (sheet.getRange(5, 1).getValue() as string).trim();

  // get the player's link from the Roster
  if (memberName.length > 0 && members.find(e => e[0] === memberName)) {
    return getPlayerDataFromUnitsTab_(memberName, filter.toLowerCase());
  }

  // check if ally code
  const allyCode = (sheet.getRange(2, 1).getValue() as number);

  if (allyCode > 0) {

    // check if ally code exist in roster
    const member = members.find(e => e[1] === allyCode);
    if (member) {
      return getPlayerDataFromUnitsTab_(member[0], filter.toLowerCase());
    }

    return getPlayerDataFromDataSource_(allyCode, filter, unitsIndex);
  }

  return undefined;
}

/** read player's data from a data source */
function getPlayerDataFromDataSource_(
  allyCode: number,
  tag: string = '',
  unitsIndex: UnitDefinition[] = undefined,
): PlayerData {

  const playerData = isDataSourceSwgohHelp_()
    ? getPlayerDataFromSwgohHelp_(allyCode)
    : getPlayerDataFromSwgohGg_(allyCode);

  if (playerData) {
    const units = playerData.units;
    const filteredUnits: UnitInstances = {};
    const filter = tag.toLowerCase();

    for (const baseId in units) {
      const u = units[baseId];
      // TODO: reload units definition if no match
      const d = unitsIndex.find(e => e.baseId === baseId);
      if (d && d.tags.indexOf(filter) > -1) {
        u.name = d.name;
        u.stats = `${u.rarity}* G${u.gearLevel} L${u.level} P${u.power}`;
        u.tags = d.tags;
        filteredUnits[baseId] = u;
      }
    }
    playerData.units = filteredUnits;

    return playerData;
  }

  return undefined;
}

/** read player's data from unit tabs */
function getPlayerDataFromUnitsTab_(memberName: string, tag: string): PlayerData {

  const roster = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 5).getValues() as [string, number, number, number, number][];
  const p = roster.find(e => e[0] === memberName);
  if (p) {
    const playerData: PlayerData = {
      allyCode: p[1],
      gp: p[2],
      heroesGp: p[3],
      name: memberName,
      shipsGp: p[4],
      units: {},
    };
    const filter = tag.toLowerCase();

    const heroesTable = new HeroesTable();
    const heroes = heroesTable.getMemberInstances(memberName);
    for (const baseId in heroes) {
      const u = heroes[baseId];
      if (filter.length === 0 || u.tags.indexOf(filter) > -1) {
        playerData.units[baseId] = u;
        playerData.level = Math.max(playerData.level, u.level);
      }
    }

    const shipsTable = new ShipsTable();
    const ships = shipsTable.getMemberInstances(memberName);
    for (const baseId in ships) {
      const u = ships[baseId];
      if (filter.length === 0 || u.tags.indexOf(filter) > -1) {
        playerData.units[baseId] = u;
        playerData.level = Math.max(playerData.level, u.level);
      }
    }

    return playerData;
  }

  return undefined;
}

/** is alignment 'Light Side' */
function isLight_(filter: string): boolean {
  return filter === ALIGNMENT.LIGHTSIDE;
}

/** get the current event definition */
function getEventDefinition_(filter: string): [string, string][] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const row = 2;
  const numRows = sheet.getLastRow() - row + 1;

  const col = (isLight_(filter) ? META_HEROES_COL : META_HEROES_DS_COL) + 2;
  const values = sheet.getRange(row, col, numRows).getValues() as [string][];
  const meta: [string, string][] = values.reduce(
    (acc, e) => {
      if (typeof e[0] === 'string' && e[0].trim().length > 0) {  // not empty
        acc.push(e[0]);
      }
      return acc;
    },
    [],
  )
  .unique()
  .map(e => [e, 'n/a']) as [string, string][];

  return meta;
}

/** output to Snapshot sheet */
function playerSnapshotOutput_(
  sheet: Sheet,
  rowGp: number,
  baseData: string[][],
  rowHeroes: number,
  meta: string[][],
): void {

  sheet.getRange(1, 3, 50, 2).clearContent();  // clear the sheet
  sheet.getRange(rowGp, 3, baseData.length, 2).setValues(baseData);
  sheet.getRange(rowHeroes, 3, meta.length, 2).setValues(meta);
}

/** create a snapshot of a player or guild member */
function playerSnapshot(): void {

  // cache the matrix of hero data
  const heroesTable = new HeroesTable();
  const shipsTable = new ShipsTable();
  const unitsIndex = heroesTable.getDefinitions().concat(shipsTable.getDefinitions());

  // collect the meta data for the heroes
  const filter = getEventFilter_(); // TODO: potentially broken if TB not sync
  const meta = getEventDefinition_(filter);

  // get all hero stats
  let countFiltered = 0;
  let countTagged = 0;
  const characterTag = getTagFilter_(); // TODO: potentially broken if TB not sync
  const powerTarget = getMinimumCharacterGp_();
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SNAPSHOT);
  const playerData = getSnapshopData_(sheet, filter, unitsIndex);
  if (playerData) {
    for (const baseId in playerData.units) {
      const u = playerData.units[baseId];
      const name = u.name;

      // does the hero meet the filtered requirements?
      if (u.rarity >= 7 && u.power >= powerTarget) {
        countFiltered += 1;
        // does the hero meet the tagged requirements?
        if (u.tags) {
          // TODO: check if u.tags exists
          debugger;
        }
        // TODO: reload units definition if no match
        unitsIndex.some((e) => {
          const found = e.baseId === baseId;
          if (found && e.tags.indexOf(characterTag) !== -1) {
            // the hero was tagged with the characterTag we're looking for
            countTagged += 1;
          }
          return found;
        });
      }

      // store hero if required
      const heroListIdx = meta.findIndex(e => name === e[0]);
      if (heroListIdx >= 0) {
        meta[heroListIdx][1] = `${u.rarity}* G${u.gearLevel} L${u.level} P${u.power}`;
      }
    }

    // format output
    const baseData = [];
    baseData.push(['GP', playerData.gp]);
    baseData.push(['GP Heroes', playerData.heroesGp]);
    baseData.push(['GP Ships', playerData.shipsGp]);
    baseData.push([`${filter} 7* P${powerTarget}+`, countFiltered]);
    baseData.push([`${characterTag} 7* P${powerTarget}+`, countTagged]);

    const rowGp = 1;
    const rowHeroes = 6;
    // output the results
    playerSnapshotOutput_(sheet, rowGp, baseData, rowHeroes, meta);
  } else {
    UI.alert('ERROR: Failed to retrieve player\'s data.');
  }
}

/** Setup new menu items when the spreadsheet opens */
function onOpen(): void {

  UI.createMenu('SWGoH')
    .addItem('Refresh TB', setupEvent.name)
    .addSubMenu(
      UI
        .createMenu('Platoons')
        .addItem('Reset', resetPlatoons.name)
        .addItem('Recommend', recommendPlatoons.name)
        .addSeparator()
        .addItem('Send Warning to Discord', allRareUnitsWebhook.name)
        .addItem('Send Rare by Unit', sendPlatoonSimplifiedByUnitWebhook.name)
        .addItem('Send Rare by Player', sendPlatoonSimplifiedByPlayerWebhook.name)
        .addSeparator()
        .addItem('Send Micromanaged by Platoon', sendPlatoonDepthWebhook.name)
        .addItem('Send Micromanaged by Player', sendMicroByPlayerWebhook.name)
        .addSeparator()
        .addItem('Register Warning Timer', registerWebhookTimer.name),
    )
    .addItem('Player Snapshot', playerSnapshot.name)
    .addToUi();
}
