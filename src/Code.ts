/**
 * @OnlyCurrentDoc
 */

/**
 * Global Variables
 */

/** Constants for sheets name */
enum SHEETS {
  ROSTER = 'Roster',
  TB = 'TB',
  PLATOONS = 'Platoon',
  BREAKDOWN = 'Breakdown',
  ESTIMATE = 'Estimate',
  LSMISSIONS = 'LS Missions',
  DSMISSIONS = 'DS Missions',
  SNAPSHOT = 'Snapshot',
  EXCLUSIONS = 'Exclusions',
  HEROES = 'Heroes',
  SHIPS = 'Ships',
  RAREUNITS = 'Rare Units',
  SEARCHUNITS = 'Search Units',
  SLICES = 'Slices',
  MAP = 'map',
  DISCORD = 'Discord',
  META = 'Meta',
  INSTRUCTIONS = 'Instructions',
}

enum DATASOURCES {
  SWGOH_HELP = 'SWGoH.help',
  SWGOH_GG = 'SWGoH.gg',
  SCORPIO = 'SCORPIO',
}

const SPREADSHEET = SpreadsheetApp.getActive();
const UI = SpreadsheetApp.getUi();

// const DEBUG_HOTH = false

// const MAX_PLAYERS = 52;
const MAX_PLAYERS = 50;
// const MIN_PLAYER_LEVEL = 65
// const POWER_TARGET = 14000

// Meta tab columns
const META_GUILD_COL = 1;
const META_FILTER_COL = 2;
const META_FILTER_ROW = 2;
const META_TAG_COL = 3;
const META_TAG_ROW = 2;
const META_UNDERGEAR_COL = 4;
const META_UNDERGEAR_ROW = 2;
const META_UNIT_PER_PLAYER_ROW = 11;

const META_MIN_LEVEL_ROW = 5;
const META_UNIT_POWER_ROW = 8;

const META_HEROES_COL = 7;
const META_HEROES_DS_COL = 16;
// const META_HEROES_SIZE_COL = 25;
const META_SORT_ROSTER_COL = 5;

// Hoth tab columns
// const HERO_PLAYER_COL_OFFSET = 9
// const SHIP_PLAYER_COL_OFFSET = 9

// Roster tab columns
// const ROSTER_SHIP_COUNT_COL = 10

const META_DATASOURCE_COL = 4;
const META_DATASOURCE_ROW = 14;

const META_UNIT_COUNTS_COL = 5;
const META_HEROES_COUNT_ROW = 5;
const META_SHIPS_COUNT_ROW = 8;

const META_ADD_PLAYER_COL = 16;
const META_REMOVE_PLAYER_COL = 18;

// Hero/Ship tab columns
const HERO_PLAYER_COL_OFFSET = 11;
const SHIP_PLAYER_COL_OFFSET = 11;

// Roster Size info
const META_GUILD_SIZE_ROW = 5;
const META_GUILD_SIZE_COL = 12;

const META_TB_COL_OFFSET = 10;

interface KeyDict {
  [key: string]: string;
}

interface KeyOffset {
  [key: string]: number;
}

interface PlayerData {
  gp: number;
  heroes_gp: number;
  level: number;
  link: string;
  name: string;
  ships_gp: number;
  units: {[key: string]: UnitInstance};
}

interface UnitDeclaration {
  Tags: string;
  UnitId: string;
  UnitName: string;
}

interface UnitInstance {
  base_id: string;
  gear_level: number;
  level: number;
  power: number;
  rarity: number;
}

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

function getSideFilter_(): string {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(META_FILTER_ROW, META_FILTER_COL)
    .getValue() as string;

  return value;
}

function getCharacterCount_(): number {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(META_HEROES_COUNT_ROW, META_UNIT_COUNTS_COL)
    .getValue() as number;

  return value;
}

function getShipCount_(): number {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(META_SHIPS_COUNT_ROW, META_UNIT_COUNTS_COL)
    .getValue() as number;

  return value;
}

function getTagFilter_(): string {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(META_TAG_ROW, META_TAG_COL)
    .getValue() as string;

  return value;
}

function get_minimum_gear_level_(): number {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(META_UNDERGEAR_ROW, META_UNDERGEAR_COL)
    .getValue() as number;

  return value;
}

function get_minimum_character_gp_(): number {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(META_UNIT_POWER_ROW, META_UNDERGEAR_COL)
    .getValue() as number;

  return value;
}

// function get_minimun_player_gp_(): number {
//   const value = SPREADSHEET.getSheetByName(SHEETS.META)
//     .getRange(META_UNIT_POWER_ROW, META_UNDERGEAR_COL)
//     .getValue() as number;

//   return value;
// }

function getMaximumPlatoonDonation_(): number {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(META_UNIT_PER_PLAYER_ROW, META_UNDERGEAR_COL)
    .getValue() as number;

  return value;
}

function getSortRoster_(): boolean {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(2, META_SORT_ROSTER_COL)
    .getValue() as string;

  return value === 'Yes';
}

function getExclusionId_(): string {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(7, META_GUILD_COL)
    .getValue() as string;

  return value;
}

/** should we use the SWGoH.help API? */
function isDataSourceSwgohHelp(): boolean {
  return get_data_source_() === DATASOURCES.SWGOH_HELP;
}

/** should we use the SWGoH.gg API? */
function isDataSourceSwgohGg(): boolean {
  return get_data_source_() === DATASOURCES.SWGOH_GG;
}

function get_data_source_() {
  const value = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(META_DATASOURCE_ROW, META_DATASOURCE_COL)
    .getValue() as string;

  return value;
}

function getGuildSize_(): number {
  const value = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(META_GUILD_SIZE_ROW, META_GUILD_SIZE_COL)
    .getValue() as number;

  return value;
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

// TODO: use allycode instead of url
function should_remove_(memberLink: string, removeMembers: [string][]): boolean {
  const result = removeMembers
  .some(e => memberLink === forceHttps(e[0]));  // return true if link is found within he list

  return result;
}

// TODO: use allycode instead of url
function remove_members_(members: string[], removeMembers: [string][]): string[] {
  const result: string[] = [];
  members.forEach((m) => {
    if (!should_remove_(m, removeMembers)) {
      result.push(m);
    }
  });

  return result;
}

type RosterEntry = [
  string,  // player name
  string,  // player url (TODO: allycode)
  number,  // gp
  number,  // heroes gp
  number  // ships gp
];

// TODO: use allycode instead of url
function add_missing_members_(result: RosterEntry[], addMembers: string[][]): RosterEntry[] {
  // for each member to add
  addMembers.filter(e => e[0].trim().length > 0)  // it must have a name
    .map(e => [e[0], forceHttps(e[1])])  // the url must use TLS
    // it must be unique. make sure the player's link isn't already in the list
    .filter(e => !result.some(l => l[1] === e[1]))
    // add member to the list. TODO: added members lack the gp information
    .forEach(e => result.push([e[0], e[1], 0, 0, 0]));

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

function getSnapshopData_(sheet: GoogleAppsScript.Spreadsheet.Sheet, tagFilter: string) {
  // try for external link
  const playerLink = (sheet.getRange(2, 1).getValue() as string).trim();
  if (playerLink.length > 0) {
    // TODO: SwgohGg & SwgohHelp API
    const unitsData: UnitStats[] = getPlayerData_SwgohGg_html_(playerLink, tagFilter);

    return unitsData;
  }
  // TODO: https://github.com/PopGoesTheWza/swgoh-tb-sheets/issues/1
  // Read data from the Heroes tab

  // no player link supplied, check for guild member
  const memberName = sheet.getRange(5, 1).getValue() as string;
  const members = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, MAX_PLAYERS, 1)
    .getValues() as [string][];

  // get the player's link from the Roster
  const match = members.find(e => e[0] === memberName);
  if (match) {
    const unitsData: UnitStats[] = getPlayerData_HeroesTab_(memberName, tagFilter.toLowerCase());

    return unitsData;
  }
  return [];
}

interface UnitStats {
  name?: string;
  base_id?: string;
  rarity: number;
  level: number;
  gear?: number;
  power: number;
  stats: string;
}

function getPlayerData_SwgohGg_html_(playerLink: string, tagFilter: string): UnitStats[] {
  const results: UnitStats[] = [];
  // get the web page source
  // const characterTag = getTagFilter_(); // TODO: potentially broken if TB not sync
  const encodedTagFilter = tagFilter.replace(' ', '+');
  const tag = tagFilter.length > 0 ? `f=${encodedTagFilter}&` : '';

  let response: GoogleAppsScript.URL_Fetch.HTTPResponse;
  let text: string;
  let page = 1;
  do {
    const url = `${playerLink}characters/?${tag}page=${page}`;

    try {
      response = UrlFetchApp.fetch(url);
    } catch (e) {
      return results; // TODO: throw error?
    }

    // divide the source into lines that can be parsed
    text = response.getContentText();
    const unitsAsHtml = text.match(/collection-char-\w+-side[\s\S]+?<\/a>[\s\S]+?<\/a>/g);
    if (unitsAsHtml) {
      unitsAsHtml.forEach((e) => {
        // TODO: try/catch regex match errors
        const name = fixString(e.match(/alt="([^"]+)/)[1]);
        const rarity = Number((e.match(/star[1-7]"/g) || []).length);
        const level = Number(e.match(/char-portrait-full-level">([^<]*)/)[1]);
        const gear = Number.parseRoman(e.match(/char-portrait-full-gear-level">([^<]*)/)[1]);
        const power = Number(e.match(/title="Power (.*?) \/ /)[1].replace(',', ''));
        const stats: string = `${rarity}* G${gear} L${level} P${power}`;

        results.push({
          name,
          rarity,
          level,
          gear,
          power,
          stats,
        });
      });
    }

    page += 1;
  } while (text.match(/aria-label="Next"/g));

  return results;
}

function getPlayerData_HeroesTab_(memberName: string, tagFilter: string): UnitStats[] {
  const results: UnitStats[] = [];
  const heroCount = getCharacterCount_();
  const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);
  const data = sheet.getRange(1, 1, 1 + heroCount, HERO_PLAYER_COL_OFFSET + MAX_PLAYERS)
    .getValues() as string[][];

  // find the member column
  const headers = data[0];
  let index: number = -1;
  for (let i = HERO_PLAYER_COL_OFFSET - 1; i < headers.length; i += 1) {
    if (headers[i] === memberName) {
      index = i;
      break;
    }
  }
  if (index > -1) {
    for (let i = 1; i < data.length; i += 1) {
      const tags = data[i][2];
      if (tagFilter.length === 0 || tags.indexOf(tagFilter) > -1) {
        const stats = data[i][index];
        const m = stats.match(/(\d+)\*L(\d+)G(\d+)P(\d+)/);
        if (m) {
          const name = data[i][0];
          const rarity = Number(m[1]);
          const level = Number(m[2]);
          const gear = Number(m[3]);
          const power = Number(m[4]);

          results.push({
            name,
            rarity,
            level,
            gear,
            power,
            stats,
          });
        }
      }
    }
  }
  return results;
}

function isLight_(tagFilter: string): boolean {
  return tagFilter === 'Light Side';
}

function get_metas_(tagFilter: string): string[][] {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const row = 2;
  const numRows = sheet.getLastRow() - row + 1;

  const col = (isLight_(tagFilter) ? META_HEROES_COL : META_HEROES_DS_COL) + 2;
  const values = sheet.getRange(row, col, numRows).getValues() as string[][];
  const meta = values.filter(e => typeof e[0] === 'string' && e[0].trim().length > 0)  // not empty
    .map(e => e[0])
    .unique()
    .map(e => [e, undefined]);

  return meta;
}

interface UnitTabIndex {
  name: string;
  base_id: string;
  tags: string;
}

function getHeroesTabIndex_(): UnitTabIndex[] {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);
  const data = sheet.getRange(2, 1, getCharacterCount_(), 3)
    .getValues() as string[][];
  const index: UnitTabIndex[] = data.map((e) => {
    return {
      name: e[0],
      base_id: e[1],
      tags: e[2],
    };
  });

  return index;
}

function playerSnapshotOutput(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowGp: number,
  baseData: string[][],
  rowHeroes: number,
  meta: string[][],
  ) {
  sheet.getRange(1, 3, 50, 2).clearContent();  // clear the sheet
  sheet.getRange(rowGp, 3, baseData.length, 2).setValues(baseData);
  sheet.getRange(rowHeroes, 3, meta.length, 2).setValues(meta);
}

/** Create a Snapshot of a Player based on criteria tracked in the workbook */
function playerSnapshot() {
  // cache the matrix of hero data
  const heroesIndex = getHeroesTabIndex_();

  // collect the meta data for the heroes
  const tagFilter = getSideFilter_(); // TODO: potentially broken if TB not sync
  const meta = get_metas_(tagFilter);

  // get all hero stats
  let countFiltered = 0;
  let countTagged = 0;
  const gp: number[] = [Number.NaN, Number.NaN, Number.NaN];  // TODO: reimplemet this feature
  const characterTag = getTagFilter_(); // TODO: potentially broken if TB not sync
  const POWER_TARGET = get_minimum_character_gp_();
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SNAPSHOT);
  const unitsData = getSnapshopData_(sheet, tagFilter);
  unitsData.forEach((u) => {
    const name = u.name;
    // does the hero meet the filtered requirements?
    if (u.rarity >= 7 && u.power >= POWER_TARGET) {
      countFiltered += 1;
      // does the hero meet the tagged requirements?
      heroesIndex.some((e) => {
        const found = e.name === name;
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
      meta[heroListIdx][1] = `${u.rarity}* G${u.gear} L${u.level} P${u.power}`;
    }
  });

  // format output
  const baseData = [];
  baseData.push(['GP', gp[0]]);
  baseData.push(['GP Heroes', gp[1]]);
  baseData.push(['GP Ships', gp[2]]);
  baseData.push([`${tagFilter} 7* P${POWER_TARGET}+`, countFiltered]);
  baseData.push([`${characterTag} 7* P${POWER_TARGET}+`, countTagged]);

  const rowGp = 1;
  const rowHeroes = 6;
  // output the results
  playerSnapshotOutput(sheet, rowGp, baseData, rowHeroes, meta);
}

/** Setup new menu items when the spreadsheet opens */
function onOpen() {
  UI.createMenu('SWGoH')
    .addItem('Refresh TB', setupTBSide.name)
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
