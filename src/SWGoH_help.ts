import { swgohhelpapi } from '../lib';

function swgoh_help_api_setting_() {
  const settings: swgohhelpapi.exports.Settings = {
    username: 'PopGoesTheWza',
    password: 'y7wX-wKYbV4^C$NF',
  };
  return settings;
}
function test_Player() {
  const settings = {
    username: 'PopGoesTheWza',
    password: 'y7wX-wKYbV4^C$NF',
  };
  const client: swgohhelpapi.exports.Client = new swgohhelpapi.exports.Client(
    settings,
  );
  const allycode = 213176142;
  const json = client.fetchPlayer({
    language: swgohhelpapi.Languages.eng_us,
    allycodes: [allycode],
    project: {
      allyCode: true,
      name: true,
      level: true,
      stats: true,
      roster: true,
      updated: true,
    },
  });
  debugger;
  const playerObj = json.filter(o => o.allycode && o.allycode === allycode);
  if (playerObj.length && playerObj.length > 0) {
    const player = playerObj.length && playerObj.length > 0 && playerObj[0];
    const gps = player.stats.filter(stat => stat.index >= 0 && stat.index <= 3);
    const characters: any[] = [];
    const ships: any[] = [];
    player.roster.forEach((r) => {
      const unit = {
        defId: r.defId,
        rarity: r.rarity,
        level: r.level,
        gear: r.gear,
        gp: r.gp,
      }
      ; (r.type === 1 ? characters : ships).push(unit);
    });
    debugger;
  }
  debugger;
}

function test_UnitsList() {
  const settings = {
    username: 'PopGoesTheWza',
    password: 'y7wX-wKYbV4^C$NF',
  };
  const client = new swgohhelpapi.exports.Client(settings);
  let request: swgohhelpapi.exports.DataRequest;
  request = {
    collection: 'categoryList',
    language: 'eng_us',
    match: {
      visible: true,
    },
  } as swgohhelpapi.exports.DataRequest;
  const categoryList = client.fetchData(request);
  debugger;
  request = {
    collection: 'unitsList',
    language: 'eng_us',
    match: {
      obtainable: true,
      rarity: 7,
    },
    project: {
      nameKey: true,
      forceAlignment: true,
      combatType: true,
      baseId: true,
      categoryIdList: true,
      unitClass: true,
    },
  } as swgohhelpapi.exports.DataRequest;
  const json = client.fetchData(request);
  debugger;
  if (json.length && json.length > 0) {
    const characters: object = {};
    const ships: object = {};
    json.forEach((r) => {
      const unit = {
        nameKey: r.nameKey,
        forceAlignment: r.forceAlignment,
        categoryIdList: r.categoryIdList,
        unitClass: r.unitClass,
      }
      ; (r.combatType === 1 ? characters : ships)[r.baseId] = unit;
    });
    debugger;
  }
  debugger;
}
