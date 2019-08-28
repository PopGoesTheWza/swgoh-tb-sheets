import { swgohhelpapi } from '../lib';

/** API Functions to pull data from swgoh.help */
namespace SwgohHelp {
  /** Constants to translate categoryId into SwgohGg tags */
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

  /** check if swgohhelpapi api is installed */
  function checkLibrary(): boolean {
    const result = !!swgohhelpapi;
    if (!result) {
      const UI = SpreadsheetApp.getUi();
      UI.alert(
        'Library swgohhelpapi not found',
        `Please visit the link below to reinstall it.
https://github.com/PopGoesTheWza/swgoh-help-api/blob/master/README.md`,
        UI.ButtonSet.OK,
      );
    }

    return result;
  }

  interface UnitsList {
    nameKey: string;
    forceAlignment: number;
    combatType: swgohhelpapi.COMBAT_TYPE;
    baseId: string;
    categoryIdList: [string];
  }

  /** Pull Units definitions from SwgohHelp */
  export function getUnitList(): UnitsDefinitions | undefined {
    if (!checkLibrary()) {
      return undefined;
    }

    const settings: swgohhelpapi.exports.Settings = {
      password: config.SwgohHelpApi.password(),
      username: config.SwgohHelpApi.username(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const units: [UnitsList] = client.fetchData({
      collection: swgohhelpapi.Collections.unitsList,
      language: swgohhelpapi.Languages.eng_us,
      match: {
        obtainable: true,
        obtainableTime: 0,
        rarity: 7,
      },
      project: {
        baseId: true,
        categoryIdList: true,
        combatType: true,
        crewList: false,
        descKey: false,
        forceAlignment: true,
        nameKey: true,
        skillReferenceList: false,
      },
    });

    if (units && units.length && units.length > 0) {
      return units.reduce(
        (acc: UnitsDefinitions, e) => {
          const bucket = e.combatType === swgohhelpapi.COMBAT_TYPE.HERO ? acc.heroes : acc.ships;

          const tags = e.categoryIdList.reduce((a: string[], c) => {
            const tag = categoryId[c as keyof typeof categoryId];
            if (tag) {
              a.push(tag);
            }
            return a;
          }, []);
          const alignment =
            e.forceAlignment === 2
              ? categoryId.alignment_light
              : e.forceAlignment === 3
              ? categoryId.alignment_dark
              : undefined;

          if (alignment) {
            tags.unshift(alignment);
          }
          const definition: UnitDefinition = {
            baseId: e.baseId,
            name: e.nameKey,
            tags: tags.unique().join(' '),
          };
          bucket.push(definition);
          return acc;
        },
        { heroes: [], ships: [] },
      );
    }

    return undefined;
  }

  /** Pull Guild data from SwgohHelp */
  export function getGuildData(): PlayerData[] | undefined {
    if (!checkLibrary()) {
      return undefined;
    }

    const settings: swgohhelpapi.exports.Settings = {
      password: config.SwgohHelpApi.password(),
      username: config.SwgohHelpApi.username(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const allycode = config.SwgohHelpApi.allyCode();
    const guild: swgohhelpapi.exports.GuildResponse[] = client.fetchGuild({
      allycode,
      enums: true,
      language: swgohhelpapi.Languages.eng_us,
      project: {
        gp: false,
        members: true,
        name: true,
        roster: {
          allyCode: true,
          gp: true,
          gpChar: true,
          gpShip: true,
          guildMemberLevel: false,
          level: true,
          name: true,
        },
      },
    });

    if (guild && guild.length && guild.length === 1) {
      const roster = guild[0].roster as swgohhelpapi.exports.PlayerResponse[];

      const red = roster.reduce(
        (acc: { allyCodes: number[]; membersData: PlayerData[] }, r) => {
          const allyCode = r.allyCode;
          if (allyCode) {
            const p: PlayerData = {
              allyCode,
              gp: r.gp!,
              heroesGp: r.gpChar!,
              level: r.level,
              name: r.name!,
              shipsGp: r.gpShip!,
              units: {},
            };
            acc.allyCodes.push(allyCode);
            acc.membersData.push(p);
          }

          return acc;
        },
        {
          allyCodes: [],
          membersData: [],
        },
      );

      if (red.allyCodes.length > 0) {
        const units = client.fetchUnits({
          allycodes: red.allyCodes,
          enums: true,
          language: swgohhelpapi.Languages.eng_us,
          project: {
            allyCode: true,
            gear: false,
            gearLevel: true,
            gp: true,
            level: true,
            mods: false,
            player: false,
            starLevel: true,
            type: false,
            zetas: true,
          },
        });

        if (units && typeof units === 'object') {
          const membersData = red.membersData;
          for (const baseId of Object.keys(units)) {
            const u = units[baseId];
            for (const i of u) {
              const allyCode = i.allyCode;
              const index = membersData.findIndex((e) => e.allyCode === allyCode);
              if (index > -1) {
                const type = i.type === swgohhelpapi.COMBAT_TYPE.HERO ? Units.TYPES.HERO : Units.TYPES.SHIP;

                membersData[index].units[baseId] = {
                  baseId,
                  gearLevel: i.gearLevel,
                  level: i.level!,
                  // name: i.???,
                  power: i.gp!,
                  rarity: i.starLevel!,
                  // stats: i.???,
                  // tags: i.???,
                  type,
                };
              }
            }
          }
        }

        return red.membersData;
      }
    }

    return undefined;
  }

  /**
   * Pull Player data from SwgohHelp
   * Units name and tags are not populated
   * returns Player data, including its units data
   */
  export function getPlayerData(allyCode: number): PlayerData | undefined {
    if (!checkLibrary()) {
      return undefined;
    }

    const settings = {
      password: config.SwgohHelpApi.password(),
      username: config.SwgohHelpApi.username(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const json = client.fetchPlayer({
      allycodes: [allyCode],
      enums: true,
      language: swgohhelpapi.Languages.eng_us,
      project: {
        allyCode: true,
        arena: false,
        // grandArena: false,
        // grandArenaLifeTime: false,
        guildName: false,
        level: true,
        name: true,
        roster: {
          combatType: true,
          crew: false,
          defId: true,
          equipped: false,
          gear: true,
          gp: true,
          level: true,
          mods: false,
          nameKey: true,
          primaryUnitStat: false,
          rarity: true,
          skills: false,
          xp: false,
        },
        stats: true,
        updated: true,
      },
    });

    if (json && json.length && json.length === 1 && !json[0].hasOwnProperty('error')) {
      const e = json[0];
      const getStats = (i: number) => {
        const s = e.stats!.find((o) => o.index === i);
        return s && s.value;
      };
      const player: PlayerData = {
        allyCode: e.allyCode!,
        gp: getStats(1)!,
        heroesGp: getStats(2)!,
        level: e.level,
        // link: e.url,
        name: e.name!,
        shipsGp: getStats(3)!,
        units: {},
      };
      for (const u of e.roster!) {
        const baseId = u.defId;
        const type = u.combatType === swgohhelpapi.COMBAT_TYPE.HERO ? Units.TYPES.HERO : Units.TYPES.SHIP;

        player.units[baseId] = {
          baseId: u.defId,
          gearLevel: u.gear,
          level: u.level,
          name: u.nameKey,
          power: u.gp,
          rarity: u.rarity,
          // stats: u.???,
          // tags: u.???,
          type,
        };
      }
      return player;
    }

    return undefined;
  }
}
