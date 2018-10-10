// ****************************************
// Ship Functions
// ****************************************

// Populate the Ships list with Member data
function populateShipsList(members) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);

  // Build a Ship Index by BaseID
  const baseIDs = sheet
    .getRange(2, 2, getShipCount_(), 1)
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
    .getRange(1, SHIP_PLAYER_COL_OFFSET, baseIDs.length, MAX_PLAYERS)
    .clearContent();

  // This will hold all our data
  // Initialize our data
  const data = baseIDs.map(e => Array(mList.length).fill(null));

  members.forEach((m) => {
    mHead[0].push(m.name);
    const units = m.units;
    baseIDs.forEach((e) => {
      const baseId = e[0];
      const u = units[baseId];
      data[hIdx[baseId]][pIdx[m.name]] = (u && `${u.rarity}*L${u.level}P${u.power}`) || '';
    });
  });
  // for ( var mKey in members) {
  //   if (mKey == "unique") { continue; }
  //   var m = members[mKey]
  //   mHead[0][pIdx[mKey]] = m['name']
  //   var units = m['units']
  //   for (var r = 0; r < baseIDs.length; r++) {
  //     var uKey = baseIDs[r]
  //     var u = units[uKey]
  //     if (u == null ) { continue; } // Means player has not unlocked unit
  //     data[hIdx[uKey]][pIdx[m['name']]] = u['rarity']+"*L"+u['level']+"P"+u['power']
  //   }
  // }
  // Write our data
  sheet.getRange(1, SHIP_PLAYER_COL_OFFSET, 1, mList.length).setValues(mHead);
  sheet
    .getRange(2, SHIP_PLAYER_COL_OFFSET, baseIDs.length, mList.length)
    .setValues(data);
}

// Initialize the list of ships
function updateShipsList(ships) {
  // update the sheet
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);

  // clear the old content
  sheet.getRange(1, 1, 300, SHIP_PLAYER_COL_OFFSET - 1).clearContent();

  const result = ships.map((e, i) => {
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

    // insert the needed count
    const formula = Utilities.formatString(`=COUNTIF(${SHEETS.PLATOONS}!$2:$16,A%s)`, row);

    hMap.push(formula);

    return hMap;
  });
  const header = [];
  header[0] = [
    'Ship',
    'Base ID',
    'Tags',
    2,
    3,
    4,
    5,
    6,
    7,
    '=CONCAT("# Needed P",Platoon!A2)',
  ];
  sheet.getRange(1, 1, 1, header[0].length).setValues(header);
  sheet.getRange(2, 1, result.length, SHIP_PLAYER_COL_OFFSET - 1).setValues(result);

  return;
}

// Populate the list of ships
function populateShipList() {
  // get the web page source
  const link = 'https://swgoh.gg/ships/';
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
      /<li\s+class="media\s+list-group-item\s+p-0\s+unit\s+character"[^]+?<\/li>/g,
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

  const units = json.map((e, i) => {
    let tags = Utilities.formatString('%s %s', e.side, e.role);

    if (e.tags) {
      tags = Utilities.formatString('%s %s', tags, e.tags.join(' '));
    }

    const result = [e.name, tags.toLowerCase()];

    // insert the star count formulas
    const row = i + 2;
    const rangeText = Utilities.formatString('$J%s:$BI%s', row, row)

    // Rarity (stars) count formulas
    ; [2, 3, 4, 5, 6, 7].forEach((stars) => {
      const formula = Utilities.formatString(
        '=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(%s,"([%s-7]+)\\*")))))',
        rangeText,
        stars,
      );
      result.push(formula);
    });

    // insdert the needed count
    const formula = Utilities.formatString(`=COUNTIF(${SHEETS.PLATOONS}!$2:$16,A%s)`, row);

    result.push(formula);
    return result;
  });

  // update the sheet
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);
  sheet.getRange(2, 1, units.length, 9).setValues(units);

  return units;
}
