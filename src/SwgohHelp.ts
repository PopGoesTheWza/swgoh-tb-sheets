import { swgohhelpapi } from '../lib';

interface SwgohHelpUnitListResponse {
  nameKey: string;
  forceAlignment: number;
  combatType: number;
  baseId: string;
}

function checkSwgohHelpLibrary_(): boolean {

  const result = Boolean(swgohhelpapi);
  if (!result) {
    UI.alert(`Library swgohhelpapi not found
Please visit the link below to reinstall it.
https://github.com/PopGoesTheWza/swgoh-help-api/blob/master/README.md`);
  }

  return result;
}

function getUnitListFromSwgohHelp_(): UnitDefinition[] {

  if (!checkSwgohHelpLibrary_()) {
    return undefined;
  }

  const settings: swgohhelpapi.exports.Settings = {
    username: config.SwgohHelp.username(),
    password: config.SwgohHelp.password(),
  };
  const client = new swgohhelpapi.exports.Client(settings);

  const units = client.fetchData({
    collection: 'unitsList',
    language: swgohhelpapi.Languages.eng_us,
    match: {
      rarity: 7,
      obtainable: true,
      obtainableTime: 0,
    },
    project: {
      nameKey: true,
      forceAlignment: true,
      combatType: true,
      categoryIdList: true,
      baseId: true,
    },
  });

  if (units && units.length && units.length > 1 && !units[0].hasOwnProperty('error')) {

    return units;
  }

  return undefined;
}

function getGuildDataFromSwgohHelp_(): PlayerData[] {

  if (!checkSwgohHelpLibrary_()) {
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

  if (guild && guild.length && guild.length === 1 && !guild[0].hasOwnProperty('error')) {

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

    if (units && typeof units === 'object' && !units.hasOwnProperty('error')) {
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
function getPlayerDataFromSwgohHelp_(allyCode: number): PlayerData {

  if (!checkSwgohHelpLibrary_()) {
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
