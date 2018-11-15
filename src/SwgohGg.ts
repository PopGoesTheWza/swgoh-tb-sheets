/** API Functions to pull data from swgoh.gg */

namespace SwgohGg {

  interface SwgohGgUnit {
    data: {
      base_id: string;
      gear: {
        base_id: string;
        is_obtained: boolean;
        slot: number;
      }[];
      gear_level: number;
      level: number;
      power: number;
      rarity: number;
      stats: KeyedNumbers;
      url: string;
      zeta_abilities: string[];
    };
  }

  interface SwgohGgPlayerData {
    ally_code: number;
    arena_leader_base_id: string;
    arena_rank: number;
    character_galactic_power: number;
    galactic_power: number;
    level: number;
    name: string;
    ship_galactic_power: number;
    url: string;
  }

  interface SwgohGgUnitResponse {
    ability_classes: string[];
    alignment: string;
    base_id: string;
    categories: string[];
    combat_type: number;
    description: string;
    gear_levels: {
      tier: number;
      gear: string[];
    }[];
    image: string;
    name: string;
    pk: number;
    power: number;
    role: string;
    url: string;
  }

  interface SwgohGgGuildResponse {
    data: {
      name: string;
      member_count: number;
      galactic_power: number;
      rank: number;
      profile_count: number;
      id: number;
    };
    players: SwgohGgPlayerResponse[];
  }

  interface SwgohGgPlayerResponse {
    data: SwgohGgPlayerData;
    units: SwgohGgUnit[];
  }

  /**
   * Send request to SwgohGg API
   * param link API 'GET' request
   * param errorMsg Message to display on error
   * returns JSON object response
   */
  function requestApi<T>(
    link: string,
    errorMsg: string = 'Error when retreiving data from swgoh.gg API',
  ): T {

    let json;
    try {
      const params: URLFetchRequestOptions = {
        // followRedirects: true,
        muteHttpExceptions: true,
      };
      const response = UrlFetchApp.fetch(link, params);
      json = JSON.parse(response.getContentText());
    } catch (e) {
      // TODO: centralize alerts
      const UI = SpreadsheetApp.getUi();
      UI.alert(errorMsg, e, UI.ButtonSet.OK);
    }

    return json || undefined;
  }

  /**
   * Pull base Character data from SwgohGg
   * returns Array of Characters with [tags, baseId, name]
   */
  export function getHeroList(): UnitDefinition[] {

    const json = requestApi<SwgohGgUnitResponse[]>(
      'https://swgoh.gg/api/characters/',
    );
    const mapping = (e: SwgohGgUnitResponse) => {
      const tags = [e.alignment, e.role, ...e.categories]
        .join(' ')
        .toLowerCase();
      const unit: UnitDefinition = {
        tags,
        baseId: e.base_id,
        name: e.name,
      };
      return unit;
    };

    return json.map(mapping);
  }

  /**
   * Pull base Ship data from SwgohGg
   * returns Array of Characters with [tags, baseId, name]
   */
  export function getShipList(): UnitDefinition[] {

    const json = requestApi<SwgohGgUnitResponse[]>(
      'https://swgoh.gg/api/ships/',
    );
    const mapping = (e: SwgohGgUnitResponse) => {
      const tags = [e.alignment, e.role, ...e.categories]
        .join(' ')
        .toLowerCase();
      const unit: UnitDefinition = {
        tags,
        baseId: e.base_id,
        name: e.name,
      };
      return unit;
    };

    return json.map(mapping);
  }

  /** Create guild API link */
  function getGuildApiLink(guildId: number): string {

    const link = `https://swgoh.gg/api/guild/${guildId}/`;

    // TODO: data check
    return link;
  }

  /**
   * Pull Guild data from SwgohGg
   * Units name and tags are not populated
   * returns Array of Guild members and their units data
   */
  export function getGuildData(guildId: number): PlayerData[] {

    const json = requestApi<SwgohGgGuildResponse>(
      getGuildApiLink(guildId),
    );
    if (json && json.players) {
      const members: PlayerData[] = [];
      for (const member of json.players) {
        const unitArray: UnitInstances = {};
        for (const e of member.units) {
          const d = e.data;
          const baseId = d.base_id;
          unitArray[baseId] = {
            baseId,
            gearLevel: d.gear_level,
            level: d.level,
            power: d.power,
            rarity: d.rarity,
          };
        }
        members.push({
          gp: member.data.galactic_power,
          heroesGp: member.data.character_galactic_power,
          level: member.data.level,  // TODO: store and process player level minimun requirement
          allyCode: +member.data.url.match(/(\d+)/)[1],
          // link: member.data.url,
          name: member.data.name,
          shipsGp: member.data.ship_galactic_power,
          units: unitArray,
        });
      }

      return members;
    }

    return undefined;
  }

  /** Create player API link */
  function getPlayerApiLink(allyCode: number): string {

    const link = `https://swgoh.gg/api/player/${allyCode}/`;

    // TODO: data check
    return link;
  }

  /**
   * Pull Player data from SwgohGg
   * Units name and tags are not populated
   * returns Player data, including its units data
   */
  export function getPlayerData(allyCode: number): PlayerData {

    const json = requestApi<SwgohGgPlayerResponse>(
      getPlayerApiLink(allyCode),
    );

    if (json && json.data) {
      const data = json.data;
      const player: PlayerData = {
        allyCode: data.ally_code,
        gp: data.galactic_power,
        heroesGp: data.character_galactic_power,
        level: data.level,
        link: data.url,
        name: data.name,
        shipsGp: data.ship_galactic_power,
        units: {},
      };
      const units = player.units;
      for (const o of json.units) {
        const d = o.data;
        const baseId = d.base_id;
        units[baseId] = {
          baseId,
          gearLevel: d.gear_level,
          level: d.level,
          power: d.power,
          rarity: d.rarity,
        };
      }

      return player;
    }

    return undefined;
  }

}
