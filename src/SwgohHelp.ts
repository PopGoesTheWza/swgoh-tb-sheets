import { swgohhelpapi } from '../lib';

namespace SwgohHelp {

  enum categoryId {
    alignment_dark = 'dark side',
    alignment_light = 'light side',

    role_attacker = 'attacker',
    role_capital = 'capital ship',
    role_healer = 'healer',
    role_support = 'support',
    role_tank = 'tank',

    affiliation_empire = 'empire',
    affiliation_firstorder = 'first order',
    affiliation_imperialtrooper = 'imperial trooper',
    affiliation_nightsisters = 'nightsister',
    affiliation_oldrepublic = 'old republic',
    affiliation_phoenix = 'phoenix',
    affiliation_rebels = 'rebel',
    affiliation_republic = 'galactic republic',
    affiliation_resistance = 'resistance',
    affiliation_rogue_one = 'rogue one',
    affiliation_separatist = 'separatist',
    character_fleetcommander = 'fleet commander',
    profession_bountyhunter = 'bounty hunters',
    profession_clonetrooper = 'clone trooper',
    profession_jedi = 'jedi',
    profession_scoundrel = 'scoundrel',
    profession_sith = 'sith',
    profession_smuggler = 'smuggler',
    // shipclass_capitalship = 'capital ship',
    shipclass_cargoship = 'cargo ship',
    species_droid = 'droid',
    species_ewok = 'ewok',
    species_geonosian = 'geonosian',
    // species_human = 'human',
    species_jawa = 'jawa',
    species_tusken = 'tusken',
    // species_wookiee = 'wookiee',
  }

  function checkLibrary(): boolean {

    const result = Boolean(swgohhelpapi);
    if (!result) {
      UI.alert(`Library swgohhelpapi not found
  Please visit the link below to reinstall it.
  https://github.com/PopGoesTheWza/swgoh-help-api/blob/master/README.md`);
    }

    return result;
  }

  type UnitsList = {
    nameKey: string,
    combatType: number,
    baseId: string,
    categoryIdList: [string],
  };

  export function getUnitList(): UnitsDefinitions {

    if (!checkLibrary()) {
      return undefined;
    }

    const settings: swgohhelpapi.exports.Settings = {
      username: config.SwgohHelp.username(),
      password: config.SwgohHelp.password(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const units: [UnitsList] = client.fetchData({
      collection: 'unitsList',
      language: swgohhelpapi.Languages.eng_us,
      match: {
        rarity: 7,
        obtainable: true,
        obtainableTime: 0,
      },
      project: {
        nameKey: true,
        // forceAlignment: true,
        combatType: true,
        categoryIdList: true,
        baseId: true,
      },
    });

    if (units && units.length && units.length > 0) {
      return units.reduce(
        (acc: UnitsDefinitions, e) => {
          const bucket = e.combatType === 0 ? acc.heroes : acc.ships;
          const tags = e.categoryIdList.reduce(
            (a: [string], c) => {
              const tag = categoryId[c];
              if (tag) {
                a.push(tag);
              }
              return a;
            },
            [],
          );
          const definition: UnitDefinition = {
            baseId: e.baseId,
            name: e.nameKey,
            tags: tags.join(' '),
          };
          bucket.push(definition);
          return acc;
        },
        { heroes: [], ships: [] });

    }

    return undefined;
  }

  export function getGuildData(): PlayerData[] {

    if (!checkLibrary()) {
      return undefined;
    }

    const settings: swgohhelpapi.exports.Settings = {
      username: config.SwgohHelp.username(),
      password: config.SwgohHelp.password(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const allycode = config.SwgohHelp.allyCode();
    const guild: swgohhelpapi.exports.GuildResponse[] = client.fetchGuild({
      allycode,
      language: swgohhelpapi.Languages.eng_us,
      project: {
        name: true,
        members: true,
        updated: true,
        roster: {
          allyCode: true,
          gp: true,
          gpChar: true,
          gpShip: true,
          level: true,  // TODO: store and process player level minimun requirement
          name: true,
          updated: true,
        },
      },
    });

    if (guild && guild.length && guild.length === 1) {

      const roster = guild[0].roster as swgohhelpapi.exports.PlayerResponse[];

      const red = roster.reduce(
        (acc: {allyCodes: number[]; playersData: PlayerData[]}, r) => {

          const allyCode = r.allyCode;
          const p: PlayerData = {
            allyCode,
            gp: r.gp,
            heroesGp: r.gpChar,
            level: r.level,
            name: r.name,
            shipsGp: r.gpShip,
            units: {},
          };
          acc.allyCodes.push(allyCode);
          acc.playersData.push(p);

          return acc;
        },
        {
          playersData: [],
          allyCodes: [],
        } as {allyCodes: number[]; playersData: PlayerData[]},
      );

      const units = client.fetchUnits({
        allycodes: red.allyCodes,
        language: swgohhelpapi.Languages.eng_us,
        project: {
          allyCode: true,
          type: true,
          gp: true,
          starLevel: true,
          level: true,
          gearLevel: true,
        },
      });

      if (units && typeof units === 'object') {
        const playersData = red.playersData;
        for (const baseId in units) {
          const u = units[baseId];
          for (const i of u) {
            const allyCode = i.allyCode;
            const index = playersData.findIndex(e => e.allyCode === allyCode);
            if (index > -1) {
              playersData[index].units[baseId] = {
                baseId,
                gearLevel: i.gearLevel,
                level: i.level,
                // name: i.???,
                power: i.gp,
                rarity: i.starLevel,
                // stats: i.???,
                // tags: i.???,
              };
            }
          }
        }
      }

      return red.playersData;
    }

    return undefined;
  }

  /**
   * Pull Player data from SwgohHelp
   * Units name and tags are not populated
   * returns Player data, including its units data
   */
  export function getPlayerData(allyCode: number): PlayerData {

    if (!checkLibrary()) {
      return undefined;
    }

    const settings = {
      username: config.SwgohHelp.username(),
      password: config.SwgohHelp.password(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const json = client.fetchPlayer({
      allycodes: [allyCode],
      language: swgohhelpapi.Languages.eng_us,
      project: {
        allyCode: true,
        guildName: true,
        level: true,
        name: true,
        roster: {
          combatType: true,
          defId: true,
          gp: true,
          gear: true,
          level: true,
          nameKey: true,
          rarity: true,
        },
        stats: true,
        updated: true,
      },
    });

    if (json && json.length && json.length === 1 && !json[0].hasOwnProperty('error')) {
      const e = json[0];
      const getStats = (i) => {
        const s = e.stats.find(o => o.index === i);
        return s && s.value;
      };
      const player: PlayerData = {
        allyCode: e.allyCode,
        gp: getStats(1),
        heroesGp: getStats(2),
        level: e.level,
        // link: e.url,
        name: e.name,
        shipsGp: getStats(3),
        units: {},
      };
      for (const u of e.roster) {
        const baseId = u.defId;
        player.units[baseId] = {
          baseId: u.defId,
          gearLevel: u.gear,
          level: u.level,
          name: u.nameKey,
          power: u.gp,
          rarity: u.rarity,
          // stats: u.???,
          // tags: u.???,
        };
      }
      return player;
    }

    return undefined;
  }

}
