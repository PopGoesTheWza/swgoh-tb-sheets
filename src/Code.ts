/**
 * @OnlyCurrentDoc
 */

/** Global Variables */
declare function getPlayerDataFromSwgohHelp_(allyCode: number): PlayerData;

// ****************************************
// Utility Functions
// ****************************************

// function fullClear(): void {

//   let sheet: GoogleAppsScript.Spreadsheet.Sheet;
//   sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
//   sheet.getRange(2, 2, MAX_PLAYERS, 9).clearContent();

//   sheet = SPREADSHEET.getSheetByName(SHEETS.TB);
//   sheet.getRange(1, META_TB_COL_OFFSET, 50, MAX_PLAYERS).clearContent();
//   sheet.getRange(2, 1, 50, META_TB_COL_OFFSET - 1).clearContent();

//   resetPlatoons();

//   sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);
//   sheet.getRange(1, 1, 300, MAX_PLAYERS + HERO_PLAYER_COL_OFFSET).clearContent();

//   sheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);
//   sheet.getRange(1, 1, 300, MAX_PLAYERS + SHIP_PLAYER_COL_OFFSET).clearContent();
// }

function getSubstringRe_(string: string, re: RegExp): string {

  const m = string.match(re);

  return m ? m[1] : '';
}

/**
 * Unescape single and double quotes from html source code
 *
 * @param {text} input The html source code to fix.
 * @return The fixed html source code.
 * @customfunction
 */
function fixString(input: string): string {

  const result = input.replace(/&quot;/g, '"').replace(/&#39;/g, "'");

  return result;
}

/**
 * Force an url to use TLS
 *
 * @param {text} url The url to fix.
 * @return The fixed url.
 * @customfunction
 */
function forceHttps(url: string): string {

  const result = url.replace('http:', 'https:');

  return result;
}

function lowerCase_(a: string, b: string): number {
  return a.toLowerCase().localeCompare(b.toLowerCase());
}

function firstElementToLowerCase_(a, b): number {
  return a[0].toLowerCase().localeCompare(b[0].toLowerCase());
}

function find_in_list_(name: string, list: string[][]): number {
  return list.findIndex(e => name === e[0]);
}

function getSnapshopData_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  tagFilter: string,
  heroesIndex: UnitDeclaration[],
): PlayerData {

  // try for external link
  const allyCode = (sheet.getRange(2, 1).getValue() as number);
  if (allyCode > 0) {
    const unitsData = getPlayerData_SwgohGgApi_(allyCode, tagFilter, heroesIndex);

    return unitsData;
  }

  // no player link supplied, check for guild member
  const memberName = sheet.getRange(5, 1).getValue() as string;
  const members = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 1)
    .getValues() as [string][];

  // get the player's link from the Roster
  const match = members.find(e => e[0] === memberName);
  if (match) {
    const playerData = getPlayerData_HeroesTab_(memberName, tagFilter.toLowerCase());

    return playerData;
  }
  return undefined;
}

function getPlayerData_SwgohGgApi_(
  allyCode: number,
  tagFilter: string = '',
  unitsIndex: UnitDeclaration[] = undefined,
): PlayerData {

  const playerData = isDataSourceSwgohHelp_()
    ? getPlayerDataFromSwgohHelp_(allyCode)
    : getPlayerDataFromSwgohGg_(allyCode);

  if (!playerData) {
    return undefined;
  }

  const units = playerData.units;
  const filteredUnits: UnitInstances = {};
  for (const key in units) {
    const u = units[key];
    const baseId = u.baseId;
    const h = unitsIndex.find(e => e.baseId === baseId);
    if (h && h.tags.indexOf(tagFilter.toLowerCase()) > -1) {
      u.name = h.name;
      u.stats = `${u.rarity}* G${u.gearLevel} L${u.level} P${u.power}`;
      u.tags = h.tags;
      filteredUnits[key] = u;
    }
  }
  playerData.units = filteredUnits;

  return playerData;
}

function getPlayerData_HeroesTab_(memberName: string, tagFilter: string): PlayerData {

  const roster = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 5).getValues() as [string, number, number, number, number][];
  const p = roster.find(e => e[0] === memberName);
  if (p) {
    const playerData: PlayerData = {
      allyCode: p[1],
      gp: p[2],
      heroesGp: p[3],
      // level: 0,
      name: memberName,
      shipsGp: p[4],
      units: {},
    };
    const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);
    const rows = getCharacterCount_() + 1;
    const cols = HERO_PLAYER_COL_OFFSET + getGuildSize_() - 1;
    const heroes = sheet.getRange(1, 1, rows, cols)
      .getValues() as string[][];

    // find the member column
    const headers = heroes[0];
    let playerCol = -1;
    for (let i = HERO_PLAYER_COL_OFFSET - 1; i < headers.length; i += 1) {
      if (headers[i] === memberName) {
        playerCol = i;
        break;
      }
    }
    if (playerCol > -1) {
      const filter = tagFilter.toLowerCase();
      for (let i = 1; i < heroes.length; i += 1) {
        const tags = heroes[i][2];
        if (filter.length === 0 || tags.indexOf(filter) > -1) {
          const stats = heroes[i][playerCol];
          const m = stats.match(/(\d+)\*L(\d+)G(\d+)P(\d+)/);
          if (m) {
            const baseId = heroes[i][1];
            const gearLevel = Number(m[3]);
            const level = Number(m[2]);
            const name = heroes[i][0];
            const power = Number(m[4]);
            const rarity = Number(m[1]);
            playerData.units[baseId] = {
              // baseId,
              gearLevel,
              level,
              name,
              power,
              rarity,
              stats,
              tags,
            };
            playerData.level = Math.max(playerData.level, level);
          }
        }
      }
    }

    return playerData;
  }

  return undefined;
}

function isLight_(tagFilter: string): boolean {
  return tagFilter === ALIGNMENT.LIGHTSIDE;
}

function get_metas_(tagFilter: string): [string, string][] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const row = 2;
  const numRows = sheet.getLastRow() - row + 1;

  const col = (isLight_(tagFilter) ? META_HEROES_COL : META_HEROES_DS_COL) + 2;
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

function getHeroesTabIndex_(): UnitDeclaration[] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);
  const data = sheet.getRange(2, 1, getCharacterCount_(), 3)
    .getValues() as string[][];
  const index: UnitDeclaration[] = data.map((e) => {
    return {
      name: e[0],
      baseId: e[1],
      tags: e[2],
    };
  });

  return index;
}

function getShipsTabIndex_(): UnitDeclaration[] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);
  const data = sheet.getRange(2, 1, getShipCount_(), 3)
    .getValues() as string[][];
  const index: UnitDeclaration[] = data.map((e) => {
    return {
      name: e[0],
      baseId: e[1],
      tags: e[2],
    };
  });

  return index;
}

function playerSnapshotOutput_(

  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowGp: number,
  baseData: string[][],
  rowHeroes: number,
  meta: string[][],
  ): void {
  sheet.getRange(1, 3, 50, 2).clearContent();  // clear the sheet
  sheet.getRange(rowGp, 3, baseData.length, 2).setValues(baseData);
  sheet.getRange(rowHeroes, 3, meta.length, 2).setValues(meta);
}

/** Create a Snapshot of a Player based on criteria tracked in the workbook */
function playerSnapshot(): void {

  // cache the matrix of hero data
  const heroesIndex = getHeroesTabIndex_();

  // collect the meta data for the heroes
  const tagFilter = getSideFilter_(); // TODO: potentially broken if TB not sync
  const meta = get_metas_(tagFilter);

  // get all hero stats
  let countFiltered = 0;
  let countTagged = 0;
  const characterTag = getTagFilter_(); // TODO: potentially broken if TB not sync
  const powerTarget = getMinimumCharacterGp_();
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SNAPSHOT);
  const playerData = getSnapshopData_(sheet, tagFilter, heroesIndex);
  if (playerData) {
    for (const baseId in playerData.units) {
      const u = playerData.units[baseId];
      const name = u.name;

      // does the hero meet the filtered requirements?
      if (u.rarity >= 7 && u.power >= powerTarget) {
        countFiltered += 1;
        // does the hero meet the tagged requirements?
        heroesIndex.some((e) => {
          const found = e.baseId === baseId;
          if (found && e.tags.indexOf(characterTag) !== -1) {
            // the hero was tagged with the characterTag we're looking for
            countTagged += 1;
          }
          return found;
        });
      }

      // store hero if required
      const heroListIdx = find_in_list_(name, meta);
      if (heroListIdx >= 0) {
        meta[heroListIdx][1] = `${u.rarity}* G${u.gearLevel} L${u.level} P${u.power}`;
      }
    }

    // format output
    const baseData = [];
    baseData.push(['GP', playerData.gp]);
    baseData.push(['GP Heroes', playerData.heroesGp]);
    baseData.push(['GP Ships', playerData.shipsGp]);
    baseData.push([`${tagFilter} 7* P${powerTarget}+`, countFiltered]);
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
