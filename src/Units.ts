// ****************************************
// Guild Unit Functions
// ****************************************

// create an array to lookup player indexes by name
function get_player_indexes_(data: string[][], offset: number) {
  const result: KeyOffset = {};

  data.forEach((e, i) => {
    result[e[0]] = i + offset;
  });

  return result;
}

// Create a lookup table for unit code names and display names
// type = "characters" or "ships"
// TODO: Caching? SwgohHelp?
function load_unit_lookup_(type: 'characters'|'ships'): KeyDict {
  const response = UrlFetchApp.fetch(`https://swgoh.gg/api/${type}/?format=json`);
  const json = JSON.parse(response.getContentText()) as {base_id: string; name: string}[];

  const result: KeyDict = {};
  json.forEach(e => result[e.base_id] = e.name);

  return result;
}

// Populate the Hero list with Member data
function populateHeroesList(members) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);

  // Build a Hero Index by BaseID
  const baseIDs = sheet
    .getRange(2, 2, getCharacterCount_(), 1)
    .getValues() as string[][];
  const hIdx = [];
  baseIDs.forEach((e, i) => {
    hIdx[e[0]] = i;
  });

  // Build a Member index by Name
  const mList = SPREADSHEET
    .getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, getGuildSize_(), 1)
    .getValues() as string[][];
  const pIdx = [];
  mList.forEach((e, i) => {
    pIdx[e[0]] = i;
  });
  const mHead = [];
  mHead[0] = [];

  // Clear out our old data, if any, including names as order may have changed
  sheet
    .getRange(1, HERO_PLAYER_COL_OFFSET, baseIDs.length, MAX_PLAYERS)
    .clearContent();

  // This will hold all our data
  // Initialize our data
  const data = baseIDs.map(e => Array(mList.length).fill(''));

  members.forEach((m) => {
    mHead[0].push(m.name);
    const units = m.units;
    baseIDs.forEach((e, i) => {
      const baseId = e[0];
      const u = units[baseId];
      data[hIdx[baseId]][pIdx[m.name]] =
          (u && `${u.rarity}*L${u.level}G${u.gear_level}P${u.power}`) || '';
    });
  });
  // Logger.log(
  //   "Last Member Data: " + JSON.stringify(members[mList[mList.length - 1]])
  // )
  // for (var key in members) {
  //   if (key == "unique") {
  //     continue
  //   }
  //   var m = members[key]
  //   mHead[0][pIdx[m["name"]]] = m["name"]
  //   // Logger.log("Parsing Units for: " + m["name"])
  //   var units = m["units"]
  //   for (r = 0; r < baseIDs.length; r++) {
  //     var uKey = baseIDs[r]
  //     var u = units[uKey]
  //     if (u == null) {
  //       continue
  //     } // Means player has not unlocked unit
  //     data[hIdx[uKey]][pIdx[m["name"]]] =
  //       u["rarity"] +
  //       "*L" +
  //       u["level"] +
  //       "G" +
  //       u["gear_level"] +
  //       "P" +
  //       u["power"]
  //   }
  //   // Logger.log("Parsed " + r + " units.")
  // }
  // Write our data
  sheet.getRange(1, HERO_PLAYER_COL_OFFSET, 1, mList.length).setValues(mHead);
  sheet
    .getRange(2, HERO_PLAYER_COL_OFFSET, baseIDs.length, mList.length)
    .setValues(data);
}

// Initialize the list of heroes
function updateHeroesList(heroes) {
  // update the sheet
  const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);

  const result = heroes.map((e, i) => {
    const hMap = [e.UnitName, e.UnitId, e.Tags];

    // insert the star count formulas
    const row = i + 2;
    const rangeText = Utilities.formatString('$J%s:$BI%s', row, row);

    {
      [2, 3, 4, 5, 6, 7].forEach((stars) => {
        const formula = Utilities.formatString(
          '=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(%s,"([%s-7]+)\\*")))))',
          rangeText,
          stars,
        );
        hMap.push(formula);
      });
    }

    // insdert the needed count
    const formula = Utilities.formatString(`=COUNTIF(${SHEETS.PLATOONS}!$20:$52,A%s)`, row);

    hMap.push(formula);

    return hMap;
  });
  // write our Units Header
  const header = [];
  header[0] = [
    'Hero',
    'Base ID',
    'Tags',
    '2',
    '3',
    '4',
    '5',
    '6',
    '7',
    '=CONCAT("# Needed P",Platoon!A2)',
  ];
  sheet.getRange(1, 1, 1, header[0].length).setValues(header);
  sheet.getRange(2, 1, result.length, HERO_PLAYER_COL_OFFSET - 1).setValues(result);

  return;
}

// Populate the list of heroes
function get_hero_list_() {
  // get the filter
  const tagFilter = ''; // getSideFilter_()

  let link = 'https://swgoh.gg/';
  if (tagFilter.length > 0) {
    link = Utilities.formatString('%scharacters/f/%s/', link, tagFilter);
  }

  // get the web page source
  let response;
  try {
    response = UrlFetchApp.fetch(link);
  } catch (e) {
    return '';
  }

  // divide the source into lines that can be parsed
  const text = response.getContentText();
  const json = text
    .match(
      /<li\s+class="media\s+list-group-item\s+p-0\s+character"[^]+?<\/li>/g,
    )
    .map((e) => {
      const tags = e
        .match(/<small[^>]*>([^]*?)<\/small/)[1]
        .split('·')
        .map((t, i, a) => {
          const tag = t.match(/\s*([^>]+?)\s*$/);
          return tag ? tag[1] : null;
        });

      const side = tags.shift();
      const role = tags.shift();
      const o = {
        role,
        side,
        tags,
        name: e
          .match(/<h5>([^<]*)/)[1]
          .replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'"),
      };
      return o;
    });

  const heroes = json.map((e, i) => {
    let tags = Utilities.formatString('%s %s', e.side, e.role);
    if (e.tags) {
      tags = Utilities.formatString('%s %s', tags, e.tags.join(' '));
    }
    const result = [e.name, tags.toLowerCase()];

    // insert the star count formulas
    const row = i + 2;
    const rangeText = Utilities.formatString('$J%s:$BI%s', row, row);

    {
      [2, 3, 4, 5, 6, 7].forEach((stars) => {
        const formula = Utilities.formatString(
          '=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(%s,"([%s-7]+)\\*")))))',
          rangeText,
          stars,
        );
        result.push(formula);
      });
    }

    // insdert the needed count
    const formula = Utilities.formatString(`=COUNTIF(${SHEETS.PLATOONS}!$20:$52,A%s)`, row);

    result.push(formula);

    return result;
  });

  return heroes;
}

// Populate the list of heroes
function populate_hero_list_() {
  const heroes = get_hero_list_();

  // update the sheet
  const sheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);
  sheet.getRange(2, 1, heroes.length, 9).setValues(heroes);

  return heroes;
}

// Get the heroes for a duplicated player name directly from their link
function get_dup_heroes_(playerLink) {
  const tagFilter = ''; // getSideFilter_()
  const encodedTagFilter = tagFilter.replace(' ', '+');

  if (playerLink.length === 0) {
    return '';
  }

  // get all hero stats
  const heroes = [];
  //  var MIN_PLAYER_LEVEL = get_minimun_player_gp_()

  // get the web page source
  let response;
  let page = 1;
  let text: string;
  do {
    let url = Utilities.formatString('%scollection/', playerLink);

    if (tagFilter.length > 0) {
      url = Utilities.formatString('%s?f=%s', url, encodedTagFilter);
    }
    if (page > 1) {
      url = Utilities.formatString(
        '%s%s%s',
        url,
        tagFilter.length > 0 ? '&page=' : '?page=',
        page,
      );
    }
    page += 1;
    try {
      response = UrlFetchApp.fetch(url);
    } catch (e) {
      return '';
    }

    // divide the source into lines that can be parsed
    text = response.getContentText();
    const units = text.match(/collection-char-\w+-side[\s\S]+?<\/a><\/div/g);
    units.forEach((e) => {
      const name = fixString(e.match(/alt="([^"]+)/)[1]);
      const stars = Number((e.match(/star[1-7]"/g) || []).length);

      // store hero
      //      if (level < MIN_PLAYER_LEVEL && stars > 0) {
      if (stars > 0) {
        const level = e.match(/char-portrait-full-level">([^<]*)/)[1];
        const gear = Number.parseRoman(
          e.match(/char-portrait-full-level">([^<]*)/)[1],
        );
        const power = Number(
          e.match(/title="Power (.*?) \/ /)[1].replace(',', ''),
        );
        const text = Utilities.formatString(
          '%s*L%sG%sP%s',
          stars,
          level,
          gear,
          power,
        );

        heroes.push([name, text]);
      }
    });
  } while (text.match(/aria-label="Next"/g));

  return heroes;
}

// Get the ships for a duplicated player name directly from their link
function get_dup_ships_(playerLink) {
  const tagFilter = ''; // getSideFilter_()
  const encodedTagFilter = tagFilter.replace('-', ' ').toLowerCase();

  if (playerLink.length === 0) {
    return '';
  }

  // get all hero stats
  const units = [];

  // get the web page source
  let response;
  let page = 1;
  let text: string;
  do {
    let url = Utilities.formatString('%sships/', playerLink);

    if (tagFilter.length > 0) {
      url = Utilities.formatString('%s?f=%s', url, encodedTagFilter);
    }
    if (page > 1) {
      url = Utilities.formatString(
        '%s%s%s',
        url,
        tagFilter.length > 0 ? '&page=' : '?page=',
        page,
      );
    }
    page += 1;
    try {
      response = UrlFetchApp.fetch(url);
    } catch (e) {
      return '';
    }

    // divide the source into lines that can be parsed
    text = response.getContentText();
    const ships = text.match(
      /class="collection-ship collection-ship-[\s\S]+?<\/a><\/div/g,
    );
    ships.forEach((e) => {
      const name = fixString(e.match(/rel="nofollow">([^<]+)<\/a/)[1]);
      const stars = Number((e.match(/ship-portrait-full-star\s*"/g) || []).length);

      // store hero
      //      if (level < MIN_PLAYER_LEVEL && stars > 0) {
      if (stars > 0) {
        const level = e.match(/char-portrait-full-level">([^<]*)/)[1];
        const power = Number(
          e.match(/title="Power (.*?) \/ /)[1].replace(',', ''),
        );
        const text = Utilities.formatString('%s*L%sP%s', stars, level, power);

        units.push([name, text]);
      }
    });
  } while (text.match(/aria-label="Next"/g));

  return units;
}

// Get all units of a type ("Hero" or "Ships")
function get_all_units_(json, members, dupNames, unitType, sheetName) {
  const units = [
    [
      unitType,
      'Tags',
      '2',
      '3',
      '4',
      '5',
      '6',
      '7',
      '=CONCAT("# Needed P",Platoon!$A$2)',
    ],
  ];

  const isHero = unitType === 'Hero';
  const lookup = load_unit_lookup_(isHero ? 'characters' : 'ships');

  // clear the sheet
  const unitsSheet = SPREADSHEET.getSheetByName(sheetName);
  unitsSheet
    .getRange(1, 1, 300, HERO_PLAYER_COL_OFFSET + MAX_PLAYERS)
    .clearContent();

  // get shared data
  const unitsList = isHero ? populate_hero_list_() : populateShipList();

  // quick lookup for unit's row
  const unitRows = [];
  unitsList.forEach((e, i) => {
    const row = i + 1;
    units[row] = unitsList[i];
    unitRows[e[0]] = row;

    // initialize the row
    const seed = Array(members.length).fill(null);
    units[row] = units[row].concat(seed);
  });

  // seed the first row with the headers (player names)
  members.forEach((e) => {
    units[0].push(e[0]);
  });

  // get a list of player column indexes
  const pIdx = get_player_indexes_(members, HERO_PLAYER_COL_OFFSET);

  // cycle through each of the units
  const combatType = isHero ? 1 : 2;
  const playerDataFound = [];
  for (const key in json) {
    const players = json[key];
    const unitName = lookup[key];
    const unitIdx = unitRows[unitName];
    if (
      unitIdx == null ||
      !players[0] ||
      players[0].combat_type !== combatType
    ) {
      // skip ships for now...
      continue;
    }

    // cycle through each player
    players.forEach((pdata) => {
      let playerName = pdata.player.trim();

      // strip leading '
      if (playerName.length > 0 && playerName.indexOf("'") === 0) {
        playerName = getSubstringRe_(playerName, /'(.*)/);
      }

      // store the player, so we know they had data
      playerDataFound[playerName] = 0;

      const idx = pIdx[playerName];
      if (idx >= 0 && dupNames[playerName] == null) {
        // only set the data for players that are not duplicated
        if (combatType === 1) {
          units[unitIdx][idx] = Utilities.formatString(
            '%s*L%sG%sP%s',
            pdata.rarity,
            pdata.level,
            pdata.gear_level,
            pdata.power,
          );
        } else {
          units[unitIdx][idx] = Utilities.formatString(
            '%s*L%sP%s',
            pdata.rarity,
            pdata.level,
            pdata.power || 0, // TODO: remove this when SCORPIO returns ship's power
          );
        }
      }
    });
  }

  // get data for duplicated players
  members.forEach((e, i) => {
    const memberName = e[0];
    // const t_memberName = memberName.length;
    if (
      memberName.length > 0 &&
      (dupNames[memberName] != null || playerDataFound[memberName] == null)
    ) {
      // get player data based on player link
      let playerUnits = (isHero
        ? get_dup_heroes_(e[1])
        : get_dup_ships_(e[1])) as string[];
      const idx = i + HERO_PLAYER_COL_OFFSET;

      // store the data
      playerUnits = playerUnits || [];
      playerUnits.forEach((e, j, a) => {
        // const t_memberName = `${memberName}:${memberName.length}`;
        const unitName = e[0];
        const unitIdx = unitRows[unitName];
        units[unitIdx][idx] = e[1];
      });
    }
  });

  // write out the results
  unitsSheet.getRange(1, 1, units.length, units[0].length).setValues(units);

  return units;
}

// set this to true if you want to debug player names
let DEBUG_PLAYERS = false;

// Debug function to find how player names are coming in with the Units page
function debugPlayerNames(json) {
  const playerNames = [];
  for (const key in json) {
    const players = json[key];

    // cycle through each player
    for (let p = 0; p < players.length; p += 1) {
      const pdata = players[p];
      let name = pdata.player;

      // strip leading '
      if (name.length > 0 && name.indexOf("'") === 0) {
        //        name = GetSubstring(name, "'", null)
        name = getSubstringRe_(name, /'(.*)/);
      }

      playerNames[name] = 0;
    }
  }

  return playerNames;
}

// Get all units for the guild
function getGuildUnitsFromJson_(json) {
  // get the member list
  const sheet = SPREADSHEET.getSheetByName(SHEETS.ROSTER);
  const members = sheet.getRange(2, 2, MAX_PLAYERS, 2).getValues() as string[][];

  if (DEBUG_PLAYERS) {
    debugPlayerNames(json);
    return;
  }

  // get a list of the members with duplicate names
  let dupNames: number[] = [];
  members
    .filter((e) => {
      return e[0].length > 0;
    })
    .forEach((e) => {
      const name = e[0];
      dupNames[name] += dupNames[name] || 0;
    });

  dupNames = dupNames.filter((e) => {
    return e > 0;
  });
  /*.map(function(e) { return 0; })*/

  const heroes = get_all_units_(json, members, dupNames, 'Hero', SHEETS.HEROES);
  get_all_units_(json, members, dupNames, 'Ship', SHEETS.SHIPS);

  return heroes;
}

// Get all units for the guild
// function getGuildUnits() {
//   // get the guild link
//   const sheet = SPREADSHEET.getSheetByName(SHEETS.META);
//   const guildLink = sheet.getRange(2, META_GUILD_COL).getValue() as string;
//   // get the guild units
//   // https://swgoh.gg/g/1082/dllidncks-dksltd/
//   // https://swgoh.gg/api/guilds/1082/units/
//   const parts = guildLink.split('/');
//   let guildID = parts[4];
//   if (DEBUG_PLAYERS) {
//     guildID = '1080'; // replace the guild ID when debugging
//   }
//   let json = {};
//   if (getUseSwgohggApi_()) {
//     try {
//       let unitLink = Utilities.formatString(
//         'https://swgoh.gg/api/guilds/%s/units/',
//         guildID,
//       );
//       // Early SCORPIO support
//       const jsonLink = SPREADSHEET
//         .getSheetByName(SHEETS.META)
//         .getRange(44, META_GUILD_COL)
//         .getValue() as string;
//       if (jsonLink && jsonLink.trim && jsonLink.trim().length >= 0) {
//         unitLink = jsonLink;
//       }
//       const response = UrlFetchApp.fetch(unitLink);
//       const text = response.getContentText();
//       json = JSON.parse(text);
//     } catch (e) {}
//   }
//   return getGuildUnitsFromJson_(json);
// }
