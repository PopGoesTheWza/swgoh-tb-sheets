// ****************************************
// Guild Unit Functions
// ****************************************

// create an array to lookup player indexes by name
function get_player_indexes_(data, offset) {

  var result = [];

  data.forEach(function(e, i, a) {
    result[e[0]] = i + offset;
  });

  return result;
}

// Create a lookup table for unit code names and display names
// type = "characters" or "ships"
function load_unit_lookup_(type) {

  var link =
  Utilities.formatString(
    "https://swgoh.gg/api/%s/?format=json",
    type
  );

  var response = UrlFetchApp.fetch(link);
  var text = response.getContentText();
  var json = JSON.parse(text);

  var result = [];
  json.forEach(function(e) {
    result[e.base_id] = e.name;
  });

  return result;
}

// Populate the list of heroes
function get_hero_list_() {

  // get the filter
  var tag_filter = "";  // get_tag_filter_();

  var link = "https://swgoh.gg/";
  if (tag_filter.length > 0) {
    link =
    Utilities.formatString(
      "%scharacters/f/%s/",
      link,
      tag_filter
    );
  }

  // get the web page source
  var response;
  try {
    response = UrlFetchApp.fetch(link);
  } catch (e) {
    return "";
  }

  // divide the source into lines that can be parsed
  var text = response.getContentText();
  var json = text
  .match(/<li\s+class="media\s+list-group-item\s+p-0\s+character"[^]+?<\/li>/g)
  .map(function(e) {
    var tags = e
    .match(/<small[^>]*>([^]*?)<\/small/)[1]
    .split("·")
    .map(function(t,i,a) {
      var tag = t.match(/\s*([^>]+?)\s*$/);
      return (tag) ? tag[1] : null;
    });

    var side = tags.shift();
    var role = tags.shift();
    var o = {
      "name": e.match(/<h5>([^<]*)/)[1].replace(/&quot;/g, "\"").replace(/&#39;/g, "'"),
      "side": side,
      "role": role,
      "tags": tags,
    };
    return o;
  });

  var heroes = json
  .map(function(e, i) {
    var tags = Utilities.formatString("%s %s", e.side, e.role);
    if (e.tags) {
      tags = Utilities.formatString("%s %s", tags, e.tags.join(" "));
    }
    var result = [e.name, tags.toLowerCase()];

    // insert the star count formulas
    var row = i + 2;
    var rangeText =
    Utilities.formatString(
      "$J%s:$BI%s",
      row,
      row
    );

    [2, 3, 4, 5, 6, 7]
    .forEach(function(stars) {
      var formula =
      Utilities.formatString(
        "=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(%s,\"([%s-7]+)\\*\")))))",
        rangeText,
        stars
      );

      result.push(formula);
    });

    // insdert the needed count
    var formula =
    Utilities.formatString(
      "=COUNTIF(Platoon!$20:$52,A%s)",
      row
    );

    result.push(formula);

    return result;
  });

  return heroes;
}

// Populate the list of heroes
function populate_hero_list_() {

  var heroes = get_hero_list_();

  // update the sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName("Heroes");
  sheet.getRange(2, 1, heroes.length, 9).setValues(heroes);

  return heroes;
}

// Get the heroes for a duplicated player name directly from their link
function get_dup_heroes_(playerLink) {

  var tag_filter = "";  // get_tag_filter_()
  var encoded_tag_filter = tag_filter.replace(" ", "+");

  if (playerLink.length == 0) {
    return "";
  }

  // get all hero stats
  var heroes = [];
//  var MinPlayerLevel = get_minimun_player_gp_();

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
        ((tag_filter.length > 0) ? "&page=" : "?page="),
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
    var units = text.match(/collection-char-\w+-side[\s\S]+?<\/a><\/div/g);
    units.forEach(function(e) {
      var name = FixString(e.match(/alt="([^"]+)/)[1]);
      var stars = Number((e.match(/star[1-7]"/g) || []).length);

      // store hero
//      if (level < MinPlayerLevel && stars > 0) {
      if (stars > 0) {
        var level = e.match(/char-portrait-full-level">([^<]*)/)[1];
        var gear = Number.parseRoman(e.match(/char-portrait-full-level">([^<]*)/)[1]);
        var power = Number(e.match(/title="Power (.*?) \/ /)[1].replace(",", ""));
        var text =
        Utilities.formatString(
          "%s*L%sG%sP%s",
          stars,
          level,
          gear,
          power
        );

        heroes.push([name, text]);
      }
    });
  } while (text.match(/aria-label="Next"/g))

  return heroes;
}

// Get the ships for a duplicated player name directly from their link
function get_dup_ships_(playerLink) {

  var tag_filter = "";  // get_tag_filter_()
  var encoded_tag_filter = tag_filter.replace("-", " ").toLowerCase();

  if (playerLink.length == 0) {
    return "";
  }

  // get all hero stats
  var units = [];
//  var MinPlayerLevel = get_minimun_player_gp_();

  // get the web page source
  var response;
  var page = 1;
  do {
    var url =
    Utilities.formatString(
      "%sships/",
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
        ((tag_filter.length > 0) ? "&page=" : "?page="),
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
    var ships = text.match(/class="collection-ship collection-ship-[\s\S]+?<\/a><\/div/g);
    ships.forEach(function(e) {
      var test_t = e.split("\n");
      var name = FixString(e.match(/rel="nofollow">([^<]+)<\/a/)[1]);
      var stars = Number((e.match(/ship-portrait-full-star\s*"/g) || []).length);

      // store hero
//      if (level < MinPlayerLevel && stars > 0) {
      if (stars > 0) {
        var level = e.match(/char-portrait-full-level">([^<]*)/)[1];
        var power = Number(e.match(/title="Power (.*?) \/ /)[1].replace(",", ""));
        var text =
        Utilities.formatString(
          "%s*L%sP%s",
          stars,
          level,
          power
        );

        units.push([name, text]);
      }
    });
  } while (text.match(/aria-label="Next"/g))

  return units;
}

// Get all units of a type ("Hero" or "Ships")
function get_all_units_(json, members, dupNames, unitType, sheetName) {

  var units = [
    [unitType, "Tags", "2", "3", "4", "5", "6", "7", "=CONCAT(\"# Needed P\",Platoon!$A$2)"]
  ];

  var isHero = (unitType == "Hero");
  var lookup = load_unit_lookup_((isHero) ? "characters" : "ships");

  // clear the sheet
  var unitsSheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  unitsSheet.getRange(1, 1, 300, HeroPlayerColOffset + MaxPlayers).clearContent();

  // get shared data
  var unitsList = (isHero) ? populate_hero_list_() : PopulateShipList();

  // quick lookup for unit's row
  var unitRows = [];
  unitsList.forEach(function(e, i) {
    var row = i + 1;
    units[row] = unitsList[i];
    unitRows[e[0]] = row;

    // initialize the row
    var seed = Array(members.length).fill(null);
    units[row] = units[row].concat(seed);
  });

  // seed the first row with the headers (player names)
  members.forEach(function(e) {
    units[0].push(e[0]);
  });

  // get a list of player column indexes
  var pIdx = get_player_indexes_(members, HeroPlayerColOffset);

  // cycle through each of the units
  var combatType = (isHero) ? 1 : 2;
  var playerDataFound = [];
  for (var key in json) {
    var players = json[key];
    var unitName = lookup[key];
    var unitIdx = unitRows[unitName];
    if (unitIdx == null || !players[0] || players[0].combat_type != combatType) {
      // skip ships for now...
      continue;
    }

    // cycle through each player
    players.forEach(function(pdata) {
      var playerName = pdata.player.trim();

      // strip leading '
      if (playerName.length > 0 && playerName.indexOf("'") === 0) {
        playerName = get_substring_re_(playerName, /'(.*)/);
      }

      // store the player, so we know they had data
      playerDataFound[playerName] = 0;

      var idx = pIdx[playerName];
      if (idx >= 0 && dupNames[playerName] == null) {
        // only set the data for players that are not duplicated
        if (combatType === 1) {
          units[unitIdx][idx] =
          Utilities.formatString(
            "%s*L%sG%sP%s",
            pdata.rarity,
            pdata.level,
            pdata.gear_level,
            pdata.power
          );
        } else {
          units[unitIdx][idx] =
          Utilities.formatString(
            "%s*L%sP%s",
            pdata.rarity,
            pdata.level,
            (pdata.power || 0)  // TODO: remove this quick fix when SCORPIO finds how ship's power is computed
          );
        }
      }
    });
  }

  // get data for duplicated players
  members
  .forEach(function(e, i) {
    var memberName = e[0];
    var t_memberName = memberName.length;
    if (memberName.length > 0 && (dupNames[memberName] != null || playerDataFound[memberName] == null)) {
      // get player data based on player link
      var playerUnits = (isHero) ? get_dup_heroes_(e[1]) : get_dup_ships_(e[1]);
      var idx = i + HeroPlayerColOffset;

      // store the data
      playerUnits = playerUnits || [];
      playerUnits.forEach(function(e, j, a) {
        var t_memberName = memberName+":"+memberName.length;
        var unitName = e[0];
        var unitIdx = unitRows[unitName];
        units[unitIdx][idx] = e[1];
      });
    }
  });

  // write out the results
  unitsSheet.getRange(1, 1, units.length, units[0].length).setValues(units);

  return units;
}

// set this to true if you want to debug player names
var DEBUG_PLAYERS = false;

// Debug function to find how player names are coming in with the Units page
function DebugPlayerNames(json) {

  var playerNames = [];
  for (var key in json) {
    var players = json[key];

    // cycle through each player
    for (var p = 0; p < players.length; ++p) {
      var pdata = players[p];
      var name = pdata.player;

      // strip leading '
      if (name.length > 0 && name.indexOf("'") == 0) {
//        name = GetSubstring(name, "'", null);
        name = get_substring_re_(name, /'(.*)/);
      }

      playerNames[name] = 0;
    }
  }

  return playerNames;
}

// Get all units for the guild
function GetGuildUnitsFromJson_(json) {

  // get the member list
  var sheet = SpreadsheetApp.getActive().getSheetByName("Roster");
  var members = sheet.getRange(2, 2, MaxPlayers, 2).getValues();

  if (DEBUG_PLAYERS) {
    DebugPlayerNames(json)
    return;
  }

  // get a list of the members with duplicate names
  var dupNames = [];
  members
  .filter(function (e) { return e[0].length > 0; })
  .forEach(function(e) {
    var name = e[0];
    dupNames[name] += dupNames[name] || 0;
  });

  dupNames = dupNames
  .filter(function (e) { return e > 0; })
  /*.map(function(e) { return 0; })*/;

  var heroes = get_all_units_(json, members, dupNames, "Hero", "Heroes");
  get_all_units_(json, members, dupNames, "Ship", "Ships");

  return heroes;
}

// Get all units for the guild
function GetGuildUnits() {

  // get the guild link
  var sheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var guildLink = sheet.getRange(2, MetaGuildCol).getValue();

  // get the guild units
  // https://swgoh.gg/g/1082/dllidncks-dksltd/
  // https://swgoh.gg/api/guilds/1082/units/
  var parts = guildLink.split("/");
  var guildID = parts[4];

  if (DEBUG_PLAYERS) {
    guildID = "1080";  // replace the guild ID when debugging
  }

  var json = {};
  if (get_use_swgohgg_api_()) {
    try {
      var unitLink =
      Utilities.formatString(
        "https://swgoh.gg/api/guilds/%s/units/",
        guildID
      );

      // Early SCORPIO support
      var jsonLink = SpreadsheetApp.getActive().getSheetByName("Meta").getRange(44, MetaGuildCol).getValue();
      if (jsonLink && jsonLink.trim && jsonLink.trim().length >= 0) {
        unitLink = jsonLink;
      }

      var response = UrlFetchApp.fetch(unitLink);
      var text = response.getContentText();
      json = JSON.parse(text);
    } catch (e) {
    }
  }

  return GetGuildUnitsFromJson_(json);
}
