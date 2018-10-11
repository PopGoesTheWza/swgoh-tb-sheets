/**
 * @OnlyCurrentDoc
 */

// ****************************************
// Global Variables
// ****************************************

//var DEBUG_HOTH = false;

var MaxPlayers = 52;
//var MinPlayerLevel = 65;
//var PowerTarget = 14000;

// Meta tab columns
var MetaGuildCol = 1;
var MetaFilterCol = 2;
var MetaTagCol = 3;
var MetaUndergearCol = 4;
var MetaHeroesCol = 7;
var MetaHeroesDSCol = 16;
var MetaHeroSizeCol = 25;
var MetaSortRosterCol = 5;

// Hoth tab columns
var HeroPlayerColOffset = 9;

// Roster tab columns
var RosterShipCountCol = 10;


// ****************************************
// Utility Functions
// ****************************************

function get_substring_re_(string, re) {

  var m = string.match(re);

  return (m === null) ? "": m[1];
}

function get_tag_filter_() {

  // get the filter
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(2, MetaFilterCol).getValue();

  return value;
}

function get_character_count_() {

  // get the count
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(2, MetaHeroSizeCol).getValue();

  return value;
}

function get_ship_count_() {

  // get the count
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(3, MetaHeroSizeCol).getValue();

  return value;
}

function get_character_tag_() {

  // get the tag
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(2, MetaTagCol).getValue();

  return value;
}

function get_minimum_gear_level_() {

  // get the undergeared value
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(2, MetaUndergearCol).getValue();

  return value;
}

/*
function get_minimun_player_gp_() {

  // get the undergeared value
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(5, MetaUndergearCol).getValue();

  return value;
}
*/

function get_minimun_character_gp_() {

  // get the undergeared value
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(8, MetaUndergearCol).getValue();

  return value;
}

function get_maximum_platoon_donation_() {

  // get the undergeared value
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(11, MetaUndergearCol).getValue();

  return value;
}

function get_sort_roster_() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(2, MetaSortRosterCol).getValue();

  return value == "Yes";
}

function get_exclusion_id_() {

  // get the Id for the exclusions sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(7, MetaGuildCol).getValue();

  return value;
}

function get_use_swgohgg_api_() {

  // should we use the swgoh.gg API?
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var value = sheet.getRange(14, MetaUndergearCol).getValue();

  return value == "Yes";
}

/**
 * Unexcape quots and double quote from html source code
 *
 * @param {text} input The html source code to fix.
 * @return The fixed html source code.
 * @customfunction
 */
function FixString(input) {

  var result = input
  .replace(/&quot;/g, "\"")
  .replace(/&#39;/g, "'");

  return result;
}

/**
 * Force an url to use TLS
 *
 * @param {text} url The url to fix.
 * @return The fixed url.
 * @customfunction
 */
function ForceHttps(url) {

  var result = url
  .replace("http:", "https:");

  return result;
}

// ****************************************
// Roster functions
// ****************************************

function should_remove_(memberLink, removeMembers) {

  var result = removeMembers
  // return true if link is found within he list
  .some(function(e) {
    return memberLink === ForceHttps(e[0]);
  });

  return result;
}

function parse_guild_gp_(text, removeMembers) {

  var result = text
  // returns an array of every members
  .match(/<td\s+data-sort-value=[\s\S]+?<\/tr/g)
  // translate html code into an array
  .map(function(e) {
    // member's link
    var link = e
    .match(/<a\s+href="([^"]+)/)[1]
    .replace(/&#39;/g, "%27");

    link =
    Utilities.formatString(
      "https://swgoh.gg%s",
      link
    );

    // member's name
    var name = e
    .match(/<strong>([^<]+)/)[1]
    .replace(/&#39;/g, "'")
    .trim();

    // member's gp (global, characters & ships)
    var gps = e
    .match(/<td\s+class="text-center">[0-9]+<\/td/g)
    .map(function(gp) {
      return Number(gp.match(/>([0-9]+)/)[1]);
    });

    // i.e. array of value for "Roster!B:H"
    return [
      name,
      link,
      gps[0],
      gps[1],
      gps[2]  /*,
      null,
      null  */
    ];
  });

  result = result
  .filter(function(e) {
    return !should_remove_(e[1], removeMembers);
  });

  return result;
}

function add_missing_members_(result, addMembers) {

  // for each member to add
  addMembers
  // it must have a name
  .filter(function(e) {
    return e[0].trim().length > 0;
  })
  // the url must use TLS
  .map(function(e) {
    return [e[0], ForceHttps(e[1])];
  })
  // it must be unique
  .filter(function(e) {
    var url = e[1];
    // make sure the player's link isn't already in the list
    var found = result
    .some(function(l) {
      return l[1] === url;
    });

    return !found;
  })
  .forEach(function(e) {
    // add member to the list
    result.push([
      e[0],
      e[1],
      0,  // TODO: added members lack the gp information
      0,
      0  /*,
      null,
      null  */
    ]);
  });

  return result;
}

function lower_case_(a, b) {

  return a.toLowerCase().localeCompare(b.toLowerCase());
}

function first_element_to_lower_case_(a, b) {

  return a[0].toLowerCase().localeCompare(b[0].toLowerCase());
}

/**
 * Get the Guild Roster from json
 *
 * @return The Roster sheet is updated.
 * @customfunction
 */
function fixDuplicatesFromJson_(json) {

  var members = {};
  var player_names = {};

  for (var unit_id in json) {
    var instances = json[unit_id];
    instances.forEach(function(unit) {
      var player_id = unit.id;
      var player = unit.player;

      if (!members[player_id]) {
        members[player_id] = {
          "player" : player
        };
      }

      if (!player_names[player]) {
        player_names[player] = {};
      }
      player_names[player][player_id] = true;
    });
  }

  for (var player in player_names) {
    var o = player_names[player];
    var keys = Object.keys(o);
    if (keys.length !== 1) {
      keys.forEach(function(player_id){
        var player = members[player_id].player;
        members[player_id].player = Utilities.formatString(
          "%s (%s)",
          player,
          player_id
        );
      });
    }
  }

  for (var unit_id in json) {
    var instances = json[unit_id];
    instances.forEach(function(unit) {
      var player_id = unit.id;

      if (members[player_id].player !== unit.player) {
        unit.player = members[player_id].player;
      }
    });
  }

  return json;
}

/**
 * Get the Guild Roster from json
 *
 * @return The Roster sheet is updated.
 * @customfunction
 */
function GetGuildRosterFromJson_(json) {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Roster");

  // get the list of members to add and remove
  // NOT SUPPORTED WITH SCORPIO ROSTER
  // var addMembers = sheet.getRange(2, 16, MaxPlayers, 2).getValues();
  // var removeMembers = sheet.getRange(2, 18, MaxPlayers, 1).getValues();

  var members = {};
  for (var unit_id in json) {
    var instances = json[unit_id];
    instances.forEach(function(unit) {
      var player_id = unit.id;
      var player = unit.player;
      var combat_type = unit.combat_type;
      var power = unit.power || 0;  // SCORPIO ships currently have no power

      if (!members[player_id]) {
        members[player_id] = {
          // "player_id" : player_id,
          "player" : player,
          "gp" : 0,
          "gp_characters" : 0,
          "gp_ships" : 0
        };
      }
      members[player_id].gp += power;
      if (combat_type === 1) {
        members[player_id].gp_characters += power;
      } else {
        members[player_id].gp_ships += power;
      }
    });
  }

  var result = [];
  for (var player_id in members) {
    var o = members[player_id];
    result.push([
      o.player,
      player_id,
      o.gp,
      o.gp_characters,
      o.gp_ships
    ]);
  }

  var sortFunction = (get_sort_roster_())
  ? first_element_to_lower_case_  // sort roster by player name
  : function(a, b) { return b[2] - a[2]; };  // sort roster by GP

  result.sort(sortFunction);

  // clear out empty data
  result = result
  .map(function(e) {
    if (e[0] == null) {  // TODO: == or ===
      return [null, null, null, null, null];
    } else {
      return e;
    }
  });

  // cleanup the header
  var header = [[
    "Name",
    "Hyper Link",
    "GP",
    "GP Heroes",
    "GP Ships"
  ]];

  // write the roster
  sheet.getRange(1, 2, 60, result[0].length).clearContent();
  sheet.getRange(1, 2, header.length, header[0].length).setValues(header);
  sheet.getRange(2, 2, result.length, result[0].length).setValues(result);
}

/**
 * Get the Guild Roster from swgoh.gg
 *
 * @return The Roster sheet is updated.
 * @customfunction
 */
function GetGuildRoster() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var guildLink =
  Utilities.formatString(
    "%sgp/",
    sheet.getRange(2, MetaGuildCol).getValue()
  );

  sheet = SpreadsheetApp.getActive().getSheetByName("Roster");

  var response = UrlFetchApp.fetch(guildLink);
  var text = response.getContentText();

  // get the list of members to add and remove
  var addMembers = sheet.getRange(2, 16, MaxPlayers, 2).getValues();
  var removeMembers = sheet.getRange(2, 18, MaxPlayers, 1).getValues();

  var result = parse_guild_gp_(text, removeMembers);
  // add missing members
  result = add_missing_members_(result, addMembers);

  var sortFunction = (get_sort_roster_())
  ? first_element_to_lower_case_  // sort roster by player name
  : function(a, b) { return b[2] - a[2]; };  // sort roster by GP

  result.sort(sortFunction);

  // clear out empty data
  result = result
  .map(function(e) {
    if (e[0] == null) {  // TODO: == or ===
      return [null, null, null, null, null  /*, null, null  */];
    } else {
      return e;
    }
  });

  // get the filter & tag
  // var PowerTarget = get_minimun_character_gp_();

  // cleanup the header
  var header = [[
    "Name",
    "Hyper Link",
    "GP",
    "GP Heroes",
    "GP Ships"
  ]];

  // write the roster
  sheet.getRange(1, 2, 60, result[0].length).clearContent();
  sheet.getRange(1, 2, header.length, header[0].length).setValues(header);
  sheet.getRange(2, 2, result.length, result[0].length).setValues(result);
}


// ****************************************
// Snapshot Functions
// ****************************************

function find_in_list_(name, list) {

  var result = list
  .findIndex(function(e) {
    return name === e[0];
  });

  return result;
}

function get_player_link_() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Snapshot");
  var playerLink = sheet.getRange(2, 1).getValue();

  if (playerLink.length === 0) {
    // no player link supplied, check for guild member
    var memberName = sheet.getRange(5, 1).getValue();

    var members = SpreadsheetApp.getActive().getSheetByName("Roster")
    .getRange(2, 2, MaxPlayers, 2).getValues();

    // get the player's link from the Roster
    var match = members.find(function(e) { return e[0] === memberName; });
    if (match) {
      playerLink = match[1];
    }
  }
  return playerLink;
}

function get_metas_() {

  var tag_filter = get_tag_filter_();  // TODO: potentially broken if TB not sync
  var isLight = (tag_filter === "Light Side");
  var col = (isLight) ? MetaHeroesCol : MetaHeroesDSCol;
  col = col +2;

  var meta = [];
  var metaSheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var row = 2;
  var lastRow = metaSheet.getLastRow();
  var numRows = lastRow - row + 1;
  var names = metaSheet
  .getRange(row, col, numRows).getValues()
  .filter(function(e) {
    return e[0].trim().length > 0;
  })
  .forEach(function(e) {
    var name = e[0];
    if (-1 === find_in_list_(name, meta)) {
      // store the meta data
      meta.push([name, null]);
    }
  });

  return meta;
}

// Create a Snapshot of a Player based on criteria tracked in the workbook
function PlayerSnapshot() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Snapshot");
  var heroesSheet = SpreadsheetApp.getActive().getSheetByName("Heroes");
  var metaSheet = SpreadsheetApp.getActive().getSheetByName("Meta");

  // clear the sheet
  sheet.getRange(1, 3, 50, 2).clearContent();

  var tag_filter = get_tag_filter_();  // TODO: potentially broken if TB not sync
  var encoded_tag_filter = tag_filter.replace(" ", "+");
  var character_tag = get_character_tag_();  // TODO: potentially broken if TB not sync

  // collect the meta data for the heroes
  var meta = get_metas_();

  // cache the matrix of hero data
  var heroCount = get_character_count_();
  var heroData = heroesSheet.getRange(2, 1, heroCount, 2).getValues();

  // get all hero stats
  var countFiltered = 0;
  var countTagged = 0;
  var gp = [];
  var PowerTarget = get_minimun_character_gp_();
  var playerLink = get_player_link_();

  // get the web page source
  var response;
  var page = 1;
  do {
    var url =
    Utilities.formatString(
      "%scollection/",
      playerLink
    );

    if (tag_filter.length > 0) {
      url =
      Utilities.formatString(
        "%s?f=%s",
        url,
        encoded_tag_filter
      );
    }
    if (page > 1) {
      url =
      Utilities.formatString(
        "%s%s%s",
        url,
        ((tag_filter.length > 0) ? "&" : "?") + "page=",
        page
      );
    }
    page++;
    try {
      response = UrlFetchApp.fetch(url);
    } catch (e) {
      return "";
    }

    // divide the source into lines that can be parsed
    var text = response.getContentText();
    var units = text
    .match(/collection-char-\w+-side[\s\S]+?<\/a><\/div/g)
    .forEach(function(e) {
      var name = FixString(e.match(/alt="([^"]+)/)[1]);
      var stars = Number((e.match(/star[1-7]"/g) || []).length);
      var level = e.match(/char-portrait-full-level">([^<]*)/)[1];
      var gear = Number.parseRoman(e.match(/char-portrait-full-gear-level">([^<]*)/)[1]);
      var power = Number(e.match(/title="Power (.*?) \/ /)[1].replace(",", ""));

      // does the hero meet the filtered requirements?
      if (stars >= 7 && power >= PowerTarget) {
        countFiltered++;
        // does the hero meet the tagged requirements?
        heroData
        .findIndex(function(e) {
          var found = e[0] === name;
          if (found && e[1].indexOf(character_tag) !== -1) {
            // the hero was tagged with the character_tag we're looking for
            countTagged++;
          }
          return found;
        });
      }

      // store hero if required
      var heroListIdx = find_in_list_(name, meta);
      if (heroListIdx >= 0) {
        meta[heroListIdx][1] =
        Utilities.formatString(
          "%s* G%s L%s P%s",
          stars,
          gear,
          level,
          power
        );
      }
    });
  } while (text.match(/aria-label="Next"/g))

  // format output
  var baseData = [];
  baseData.push(["GP", gp[0]]);
  baseData.push(["GP Heroes", gp[1]]);
  baseData.push(["GP Ships", gp[2]]);
  baseData.push([
    Utilities.formatString(
      "%s 7* P%s+",
      tag_filter,
      PowerTarget
    ),
    countFiltered
  ]);
  baseData.push([
    Utilities.formatString(
      "%s 7* P%s+",
      character_tag,
      PowerTarget
    ),
    countTagged
]);

  var rowGp = 1;
//  var rowFilter = 4;
  var rowHeroes = 6;
  // output the results
  sheet.getRange(rowGp, 3, baseData.length, 2).setValues(baseData);
  sheet.getRange(rowHeroes, 3, meta.length, 2).setValues(meta);
}


// ****************************************
// UI Functions
// ****************************************

// Setup new menu items when the spreadsheet opens
function onOpen() {

  var ui = SpreadsheetApp.getUi();

  ui.createMenu("SWGoH")
    .addItem("Refresh TB", "SetupTBSide")
    .addSubMenu(ui.createMenu("Platoons")
      .addItem("Reset", "ResetPlatoons")
      .addItem("Recommend", "RecommendPlatoons")
      .addSeparator()
      .addItem("Send Warning to Discord", "AllRareUnitsWebhook")
      .addItem("Send Rare by Unit", "SendPlatoonSimplifiedByUnitWebhook")
      .addItem("Send Rare by Player", "SendPlatoonSimplifiedByPlayerWebhook")
      .addSeparator()
      .addItem("Send Micromanaged by Platoon", "SendPlatoonDepthWebhook")
      .addItem("Send Micromanaged by Player", "SendMicroByPlayerWebhook")
      .addSeparator()
      .addItem("Register Warning Timer", "RegisterWebhookTimer"))
    .addItem("Player Snapshot", "PlayerSnapshot")
    .addToUi();
}

// ****************************************
// Notes
// https://developers.google.com/apps-script/reference/properties/
// ****************************************
