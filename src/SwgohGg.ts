// import { PlayerData, UnitDeclaration } from './SWGOH_gg_API';
// import { PlayerData, UnitInstance } from './SWGOH_gg_API';
/** API Functions to pull data from swgoh.gg */

interface SwgohGgUnitDefinition {
  base_id: string;
  alignment: string;
  categories: string[];
  name: string;
  role: string;
}

interface SwgohUnit {
  data: {
    base_id: string;
    gear_level: number;
    level: number;
    power: number;
    rarity: number;
  };
}

interface SwgohGgGuildMember {
  data: {
    character_galactic_power: number;
    galactic_power: number;
    level: number;
    name: string;
    ship_galactic_power: number;
    url: string;
  };
  units: SwgohUnit[];
}

interface SwgohGgGuildData {
  players: SwgohGgGuildMember[];
}

/** Get the guild ID */
function get_guild_id_(): string {
  const metaSWGOHLinkCol = 1;
  const metaSWGOHLinkRow = 2;

  const guildLink = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
    .getValue() as string;
  const parts = guildLink.split('/');
  // TODO: input check
  const guildId = parts[4];

  return guildId;
}

// Create Guild API Link
function get_guild_api_link_(): string {
  const link = `https://swgoh.gg/api/guild/${get_guild_id_()}/`;
  // TODO: data check
  return link;
}

// function isSWGOHggSource() {
//   const value = SPREADSHEET.getSheetByName(SHEETS.META)
//     .getRange(META_DATASOURCE_ROW, META_DATASOURCE_COL)
//     .getValue() as string;
//   // TODO: centralize constants
//   return value === DATASOURCES.SWGOH_GG;
// }

// Pull base Character data from SWGoH.gg
// @returns Array of Characters with [name, base_id, tags]
function getUnitsFromSWGoHgg_<T>(link: string, errorMsg: string): T {
  let json;
  try {
    // const link = "https://swgoh.gg/api/characters/?format=json"
    const params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      // followRedirects: true,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(link, params);
    // const responseObj = {
    //   getContentText: response.getContentText().split('\n'),
    //   getHeaders: response.getHeaders(),
    //   getResponseCode: response.getResponseCode(),
    // };
    // if (response.getResponseCode() !== 200) {
    //   debugger;
    // }
    json = JSON.parse(response.getContentText());
  } catch (e) {
    // TODO: centralize alerts
    UI.alert(errorMsg, e, UI.ButtonSet.OK);
  }

  return json || undefined;
}

/**
 * Pull base Character data from SWGoH.gg
 * @returns Array of Characters with [name, base_id, tags]
 */
function getHeroesFromSWGOHgg(): UnitDeclaration[] {
  const json = getUnitsFromSWGoHgg_<SwgohGgUnitDefinition[]>(
    'https://swgoh.gg/api/characters/?format=json',
    'Error when retreiving data from swgoh.gg API',
  );
  const mapping = (e: SwgohGgUnitDefinition) => {
    const tags = [e.alignment, e.role, ...e.categories]
      .join(' ')
      .toLowerCase();
    const unit: UnitDeclaration = {
      Tags: tags,
      UnitId: e.base_id,
      UnitName: e.name,
    };
    return unit;
  };
  return json.map(mapping);
}

/**
 * Pull base Ship data from SWGoH.gg
 * @returns Array of Characters with [name, base_id, tags]
 */
function getShipsFromSWGOHgg(): UnitDeclaration[] {
  const json = getUnitsFromSWGoHgg_<SwgohGgUnitDefinition[]>(
    'https://swgoh.gg/api/ships/?format=json',
    'Error when retreiving data from swgoh.gg API',
  );
  const mapping = (e: SwgohGgUnitDefinition) => {
    const tags = [e.alignment, e.role, ...e.categories]
      .join(' ')
      .toLowerCase();
    const unit: UnitDeclaration = {
      UnitId: e.base_id,
      UnitName: e.name,
      Tags: tags,
    };
    return unit;
  };
  return json.map(mapping);
}

// Pull Guild data from SWGoH.gg
// @returns Array of Guild members and their character data
function getGuildDataFromSwgohGg(): PlayerData[] {
  const json = getUnitsFromSWGoHgg_<SwgohGgGuildData>(
    get_guild_api_link_(),
    'Error when retreiving data from swgoh.gg API',
  );
  const members: PlayerData[] = [];
  json.players.forEach((member) => {
    // const player_id = member.data.name  // TODO: duplicate names? member.data.(url/ally_code)?
    // const player_id = member.data.url;
    // const ally_code = member.data.ally_code;
    const unitArray = {};
    member.units.forEach((e) => {
      const unit = e.data;
      const q = {
        level: unit.level,
        gear_level: unit.gear_level,
        power: unit.power,
        rarity: unit.rarity,
        base_id: unit.base_id,
      };
      const base_id = unit.base_id;
      unitArray[base_id] = q;
    });
    members.push({
      gp: member.data.galactic_power,
      heroes_gp: member.data.character_galactic_power,
      level: member.data.level,
      link: member.data.url,
      name: member.data.name,
      ships_gp: member.data.ship_galactic_power,
      units: unitArray,
    });
  });

  return members;
}
