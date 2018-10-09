// *****************************************
// ** Functions to parse SCORPIO JSON Data
// *****************************************

function isScorpioSource() {
  const value = String(
    SPREADSHEET
      .getSheetByName(SHEETS.META)
      .getRange(META_DATASOURCE_ROW, META_DATASOURCE_COL)
      .getValue(),
  );
  // TODO: centralize constants
  return value === 'SCORPIO';
}

// Pull Base IDs, unitType should be "Heroes" or "Ships"
function get_base_ids_(unitType) {
  const row = unitType === SHEETS.HEROES ? META_HEROES_COUNT_ROW : META_SHIPS_COUNT_ROW;
  const rows = Number(
    SPREADSHEET
      .getSheetByName(SHEETS.META)
      .getRange(row, META_UNIT_COUNTS_COL)
      .getValue(),
  );
  const baseIDs = SPREADSHEET
    .getSheetByName(unitType)
    .getRange(2, 1, rows, 2)
    .getValues() as string[][];
  const data = [];
  baseIDs.forEach((e: string[]) => {
    data[e[1]] = e[0];
  });
  return data;
}

function getGuildDataFromScorpio() {
  const metaScorpioLinkCol = 1;
  const metaScorpioLinkRow = 11;
  const members = [];

  const link = String(
    SPREADSHEET
      .getSheetByName(SHEETS.META)
      .getRange(metaScorpioLinkRow, metaScorpioLinkCol)
      .getValue(),
  );
  if (!link || link.trim().length === 0) {
    UI.alert(
      'Unable to find SCORPIO Link',
      'Check value on Meta tab',
      UI.ButtonSet.OK,
    );
    return [];
  }
  let json;

  const hIndex = get_base_ids_(SHEETS.HEROES);
  const sIndex = get_base_ids_(SHEETS.SHIPS);

  try {
    const params = {
      //      followRedirects: true,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(link, params);
    const responseObj = {
      getContentText: response.getContentText().split('\n'),
      getHeaders: response.getHeaders(),
      getResponseCode: response.getResponseCode(),
    };
    if (response.getResponseCode() !== 200) {
      debugger;
    }
    json = JSON.parse(response.getContentText());
  } catch (e) {
    UI.alert(
      'Unable to Parse SCORPIO Data',
      'Check link in Meta tab. It should be a link and not JSON data',
      UI.ButtonSet.OK,
    );
    return [];
  }

  for (const unitId in json) {
    // if (unit_id === "unique") {
    //   continue
    // }
    const instances = json[unitId];
    instances.forEach((o) => {
      // const player_id = o.player // TODO: duplicate names?  id?
      const playerId = `p/${o.id}`;
      const pname = o.player;
      let member = [];
      const q = [];
      q['base_id'] = unitId;
      q['level'] = o.level;
      q['power'] = o.power;
      q['rarity'] = o.rarity;
      if (o.combat_type === 1) {
        q['gear_level'] = o.gear_level;
        q['name'] = hIndex[unitId];
      } else {
        q['name'] = sIndex[unitId];
      }
      if (members[playerId]) {
        member = members[playerId];
      } else {
        members[playerId] = member;
        member['name'] = pname;
        member['gp'] = 0;
        member['ships_gp'] = 0;
        member['heroes_gp'] = 0;
        member['units'] = [];
        member['link'] = `p/${o.id}`;
      }
      member['units'][unitId] = q;
      member['gp'] += q['power'];
      if (o.combat_type === 1) {
        member['heroes_gp'] += q['power'];
      } else {
        member['ships_gp'] += q['power'];
      }
    });
  }
  const flat = [];
  for (const key in members) {
    flat.push(members[key]);
  }
  return flat;
  // return [...members]
}
