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

// function fullClear() {
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

function getSubstringRe_(string: string, re: RegExp) {
  const m = string.match(re);
  return m ? m[1] : '';
}

function getTagFilter_() {
  const value = SPREADSHEET
    .getSheetByName(SHEETS.META)
    .getRange(META_FILTER_ROW, META_FILTER_COL)
    .getValue() as string;
  return value;
}

function getCharacterCount_() {
  const value = SPREADSHEET
    .getSheetByName(SHEETS.META)
    .getRange(META_HEROES_COUNT_ROW, META_UNIT_COUNTS_COL)
    .getValue() as number;
  return value;
}

function getShipCount_() {
  const value = SPREADSHEET
    .getSheetByName(SHEETS.META)
    .getRange(META_SHIPS_COUNT_ROW, META_UNIT_COUNTS_COL)
    .getValue() as number;
  return value;
}

function get_character_tag_() {
  const value = SPREADSHEET
    .getSheetByName(SHEETS.META)
    .getRange(META_TAG_ROW, META_TAG_COL)
    .getValue() as string;
  return value;
}

function get_minimum_gear_level_() {
  const value = SPREADSHEET
    .getSheetByName(SHEETS.META)
    .getRange(META_UNDERGEAR_ROW, META_UNDERGEAR_COL)
    .getValue() as number;
  return value;
}

function get_minimum_character_gp_() {
  const value = SPREADSHEET
    .getSheetByName(SHEETS.META)
    .getRange(META_UNIT_POWER_ROW, META_UNDERGEAR_COL)
    .getValue() as number;
  return value;
}

/*
function get_minimun_player_gp_() {

  // get the undergeared value
  var sheet = SPREADSHEET.getSheetByName("Meta")
  var value = sheet.getRange(META_UNIT_POWER_ROW, META_UNDERGEAR_COL).getValue()

  return value
}
*/

function getMaximumPlatoonDonation_() {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const value = sheet.getRange(META_UNIT_PER_PLAYER_ROW, META_UNDERGEAR_COL).getValue();

  return value;
}

function getSortRoster_() {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const value = sheet.getRange(2, META_SORT_ROSTER_COL).getValue();

  return value === 'Yes';
}

function getExclusionId_() {
  const value = SPREADSHEET
    .getSheetByName(SHEETS.META)
    .getRange(7, META_GUILD_COL)
    .getValue() as string;
  return value;
}

/** should we use the SWGoH.help API? */
function isDataSourceSwgohHelp() {
  return get_data_source_() === DATASOURCES.SWGOH_HELP;
}

/** should we use the SWGoH.gg API? */
function isDataSourceSwgohGg() {
  return get_data_source_() === DATASOURCES.SWGOH_GG;
}

// function getUseSwgohggApi_() {
//   // should we use the swgoh.gg API?
//   // var sheet = SPREADSHEET.getSheetByName("Meta")
//   // var value = sheet.getRange(14, META_UNDERGEAR_COL).getValue()
//   // return value == "Yes"
//   return get_data_source_() === DATASOURCES.SWGOH_GG;
// }

function get_data_source_() {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const value = sheet.getRange(META_DATASOURCE_ROW, META_DATASOURCE_COL).getValue();

  return value;
}

function getGuildSize_() {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  return Number(sheet.getRange(META_GUILD_SIZE_ROW, META_GUILD_SIZE_COL).getValue());
}

/**
 * Unexcape quots and double quote from html source code
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

// ****************************************
// Roster functions
// ****************************************

function should_remove_(memberLink: string, removeMembers: string[][]): boolean {
  const result = removeMembers
    // return true if link is found within he list
    .some((e) => {
      return memberLink === forceHttps(e[0]);
    });

  return result;
}

function remove_members_(members, removeMembers) {
  const result = [];
  members.forEach((m) => {
    if (!should_remove_(m, removeMembers)) {
      result.push(m);
    }
  });
  return result;
}

/*
function parse_guild_gp_(text, removeMembers) {

  var result = text
  // returns an array of every members
  .match(/<td\s+data-sort-value=[\s\S]+?<\/tr/g)
  // translate html code into an array
  .map(function(e) {
    // member's link
    var link = e
    .match(/<a\s+href="([^"]+)/)[1]
    .replace(/&#39;/g, "%27")

    link =
    Utilities.formatString(
      "https://swgoh.gg%s",
      link
    )

    // member's name
    var name = e
    .match(/<strong>([^<]+)/)[1]
    .replace(/&#39;/g, "'")
    .trim()

    // member's gp (global, characters & ships)
    var gps = e
    .match(/<td\s+class="text-center">[0-9]+<\/td/g)
    .map(function(gp) {
      return Number(gp.match(/>([0-9]+)/)[1])
    })

    // i.e. array of value for "Roster!B:H"
    return [
      name,
      link,
      gps[0],
      gps[1],
      gps[2]  //,
//      null,
//      null
    ]
  })

  result = result
  .filter(function(e) {
    return !should_remove_(e[1], removeMembers)
  })

  return result
}
*/

function add_missing_members_(result, addMembers) {
  // for each member to add
  addMembers
    // it must have a name
    .filter((e) => {
      return e[0].trim().length > 0;
    })
    // the url must use TLS
    .map((e) => {
      return [e[0], forceHttps(e[1])];
    })
    // it must be unique
    .filter((e) => {
      const url = e[1];
      // make sure the player's link isn't already in the list
      const found = result.some((l) => {
        return l[1] === url;
      });

      return !found;
    })
    .forEach((e) => {
      // add member to the list
      result.push([
        e[0],
        e[1],
        0, // TODO: added members lack the gp information
        0,
        0 /*,
      null,
      null  */,
      ]);
    });

  return result;
}

function lowerCase_(a, b) {
  return a.toLowerCase().localeCompare(b.toLowerCase());
}

function firstElementToLowerCase_(a, b) {
  return a[0].toLowerCase().localeCompare(b[0].toLowerCase());
}

/**
 * Get the Guild Roster from json
 *
 * @return The Roster sheet is updated.
 * @customfunction
 */
/*
function fixDuplicatesFromJson_(json) {

  var members = {}
  var player_names = {}

  for (var unit_id in json) {
    var instances = json[unit_id]
    instances.forEach(function(unit) {
      var player_id = unit.id
      var player = unit.player

      if (!members[player_id]) {
        members[player_id] = {
          "player" : player
        }
      }

      if (!player_names[player]) {
        player_names[player] = {}
      }
      player_names[player][player_id] = true
    })
  }

  for (var player in player_names) {
    var o = player_names[player]
    var keys = Object.keys(o)
    if (keys.length !== 1) {
      keys.forEach(function(player_id){
        var player = members[player_id].player
        members[player_id].player = Utilities.formatString(
          "%s (%s)",
          player,
          player_id
        )
      })
    }
  }

  for (var unit_id in json) {
    var instances = json[unit_id]
    instances.forEach(function(unit) {
      var player_id = unit.id

      if (members[player_id].player !== unit.player) {
        unit.player = members[player_id].player
      }
    })
  }

  return json
}
*/

/**
 * Get the Guild Roster from json
 *
 * @return The Roster sheet is updated.
 * @customfunction
 */
/*
function GetGuildRosterFromJson_(json) {

  var sheet = SPREADSHEET.getSheetByName("Roster")

  // get the list of members to add and remove
  // NOT SUPPORTED WITH SCORPIO ROSTER
  // var addMembers = sheet.getRange(2, 16, MAX_PLAYERS, 2).getValues()
  // var removeMembers = sheet.getRange(2, 18, MAX_PLAYERS, 1).getValues()

  var members = {}
  for (var unit_id in json) {
    var instances = json[unit_id]
    instances.forEach(function(unit) {
      var player_id = unit.id
      var player = unit.player
      var combat_type = unit.combat_type
      var power = unit.power || 0;  // SCORPIO ships currently have no power

      if (!members[player_id]) {
        members[player_id] = {
          // "player_id" : player_id,
          "player" : player,
          "gp" : 0,
          "gp_characters" : 0,
          "gp_ships" : 0
        }
      }
      members[player_id].gp += power
      if (combat_type === 1) {
        members[player_id].gp_characters += power
      } else {
        members[player_id].gp_ships += power
      }
    })
  }

  var result = []
  for (var player_id in members) {
    var o = members[player_id]
    result.push([
      o.player,
      player_id,
      o.gp,
      o.gp_characters,
      o.gp_ships
    ])
  }

  var sortFunction = (getSortRoster_())
  ? firstElementToLowerCase_  // sort roster by player name
  : function(a, b) { return b[2] - a[2]; };  // sort roster by GP

  result.sort(sortFunction)

  // clear out empty data
  result = result
  .map(function(e) {
    if (e[0] == null) {  // TODO: == or ===
      return [null, null, null, null, null]
    } else {
      return e
    }
  })

  // cleanup the header
  var header = [[
    "Name",
    "Hyper Link",
    "GP",
    "GP Heroes",
    "GP Ships"
  ]]

  // write the roster
  sheet.getRange(1, 2, 60, result[0].length).clearContent()
  sheet.getRange(1, 2, header.length, header[0].length).setValues(header)
  sheet.getRange(2, 2, result.length, result[0].length).setValues(result)
}
*/

/**
 * Get the Guild Roster from swgoh.gg
 *
 * @return The Roster sheet is updated.
 * @customfunction
 */
/*
function GetGuildRoster() {

  var sheet = SPREADSHEET.getSheetByName("Meta")
  var guildLink =
  Utilities.formatString(
    "%sgp/",
    sheet.getRange(2, META_GUILD_COL).getValue()
  )

  sheet = SPREADSHEET.getSheetByName("Roster")

  var response = UrlFetchApp.fetch(guildLink)
  var text = response.getContentText()

  // get the list of members to add and remove
  var addMembers = sheet.getRange(2, 16, MAX_PLAYERS, 2).getValues()
  var removeMembers = sheet.getRange(2, 18, MAX_PLAYERS, 1).getValues()

  var result = parse_guild_gp_(text, removeMembers)
  // add missing members
  result = add_missing_members_(result, addMembers)

  var sortFunction = (getSortRoster_())
  ? firstElementToLowerCase_  // sort roster by player name
  : function(a, b) { return b[2] - a[2]; };  // sort roster by GP

  result.sort(sortFunction)

  // clear out empty data
  result = result
  .map(function(e) {
    if (e[0] == null) {  // TODO: == or ===
      //return [null, null, null, null, null, null, null]
      return [null, null, null, null, null]
    } else {
      return e
    }
  })

  // get the filter & tag
  // var POWER_TARGET = get_minimum_character_gp_()

  // cleanup the header
  var header = [[
    "Name",
    "Hyper Link",
    "GP",
    "GP Heroes",
    "GP Ships"
  ]]

  // write the roster
  sheet.getRange(1, 2, 60, result[0].length).clearContent()
  sheet.getRange(1, 2, header.length, header[0].length).setValues(header)
  sheet.getRange(2, 2, result.length, result[0].length).setValues(result)
}
*/

// ****************************************
// Snapshot Functions
// ****************************************

function find_in_list_(name, list) {
  return list.findIndex(e => name === e[0]);
}

function get_player_link_() {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SNAPSHOT);
  let playerLink = sheet.getRange(2, 1).getValue() as string;

  if (playerLink.length === 0) {
    // no player link supplied, check for guild member
    const memberName = sheet.getRange(5, 1).getValue();

    const members = SPREADSHEET
      .getSheetByName(SHEETS.ROSTER)
      .getRange(2, 2, MAX_PLAYERS, 2)
      .getValues() as string[][];

    // get the player's link from the Roster
    const match = members.find(e => e[0] === memberName);
    if (match) {
      playerLink = match[1];
    }
  }
  return playerLink;
}

function isLight_(tagFilter: string): boolean {
  return tagFilter === 'Light Side';
}

function get_metas_(tagFilter) {
  const col = (isLight_(tagFilter) ? META_HEROES_COL : META_HEROES_DS_COL) + 2;
  const metaSheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const row = 2;
  const numRows = metaSheet.getLastRow() - row + 1;

  const values = metaSheet.getRange(row, col, numRows).getValues() as string[][];
  const meta = values
  .filter(e => typeof e[0] === 'string' && e[0].trim().length > 0)  // not empty
  .map(e => e[0])  // TODO: .reduce()
  .unique()
  .map(e => [e, undefined]);

  return meta;
}

// Create a Snapshot of a Player based on criteria tracked in the workbook
function playerSnapshot() {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SNAPSHOT);
  const heroesSheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);
  // const metaSheet = SPREADSHEET.getSheetByName("Meta")

  const tagFilter = getTagFilter_(); // TODO: potentially broken if TB not sync
  const encodedTagFilter = tagFilter.replace(' ', '+');
  const characterTag = get_character_tag_(); // TODO: potentially broken if TB not sync

  // clear the sheet
  sheet.getRange(1, 3, 50, 2).clearContent();

  // collect the meta data for the heroes
  const meta = get_metas_(tagFilter);

  // cache the matrix of hero data
  const heroCount = getCharacterCount_();
  const heroData = heroesSheet
    .getRange(2, 1, heroCount, 3)
    .getValues() as string[][];

  // get all hero stats
  let countFiltered = 0;
  let countTagged = 0;
  const gp: number[] = [];
  const POWER_TARGET = get_minimum_character_gp_();
  const playerLink = get_player_link_();

  // get the web page source
  let response: GoogleAppsScript.URL_Fetch.HTTPResponse;
  let page = 1;
  let text: string;
  do {
    const tag = tagFilter.length > 0 ? `f=${encodedTagFilter}&` : '';
    const url = `${playerLink}characters/?${tag}page=${page}`;

    page += 1;
    try {
      response = UrlFetchApp.fetch(url);
    } catch (e) {
      return; // Throw error?
    }

    // divide the source into lines that can be parsed
    text = response.getContentText();
    const rem: RegExpMatchArray = text
      // .match(/collection-char-\w+-side[\s\S]+?<\/a><\/div/g)
      .match(/collection-char-\w+-side[\s\S]+?<\/a>[\s\S]+?<\/a>/g) as RegExpMatchArray;
    if (rem) {
      rem.forEach((e) => {
        const name = fixString((e.match(/alt="([^"]+)/) as RegExpMatchArray)[1]);
        const stars = Number((e.match(/star[1-7]"/g) || []).length);
        const level = (e.match(/char-portrait-full-level">([^<]*)/) as RegExpMatchArray)[1];
        const gear = Number.parseRoman(
          (e.match(/char-portrait-full-gear-level">([^<]*)/) as RegExpMatchArray)[1]);
        const power = Number(
          (e.match(/title="Power (.*?) \/ /) as RegExpMatchArray)[1].replace(',', ''));

        // does the hero meet the filtered requirements?
        if (stars >= 7 && power >= POWER_TARGET) {
          countFiltered += 1;
          // does the hero meet the tagged requirements?
          heroData.findIndex((e) => {
            const found = e[0] === name;
            if (found && e[2].indexOf(characterTag) !== -1) {
              // the hero was tagged with the characterTag we're looking for
              countTagged += 1;
            }
            return found;
          });
        }

        // store hero if required
        const heroListIdx = find_in_list_(name, meta);
        if (heroListIdx >= 0) {
          meta[heroListIdx][1] = `${stars}* G${gear} L${level} P${power}`;
        }
      });
    }
  } while (text.match(/aria-label="Next"/g));

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
  sheet.getRange(rowGp, 3, baseData.length, 2).setValues(baseData);
  sheet.getRange(rowHeroes, 3, meta.length, 2).setValues(meta);
}

// Setup new menu items when the spreadsheet opens
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
