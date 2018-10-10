// TODO: fix tslint issue where 'import' statement messes with globals
import { swgohhelpapi } from '../lib';

/** Get the SWGoH.help API username */
function get_SwgohHelp_username_(): string {
  const metaSWGOHLinkCol = 1;
  const metaSWGOHLinkRow = 16;
  const result = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
    .getValue() as string;

  return result;
}

/** Get the SWGoH.help API password */
function get_SwgohHelp_password_(): string {
  const metaSWGOHLinkCol = 1;
  const metaSWGOHLinkRow = 18;
  const result = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
    .getValue() as string;

  return result;
}

/** Get the guild member ally code */
function get_SwgohHelp_allycode_(): number {
  const metaSWGOHLinkCol = 1;
  const metaSWGOHLinkRow = 20;
  const result = SPREADSHEET.getSheetByName(SHEETS.META)
    .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
    .getValue() as number;

  return result;
}

function getGuildDataFromSwgohHelp(): PlayerData[] {
  const members: PlayerData[] = [];
  const settings: swgohhelpapi.exports.Settings = {
    username: get_SwgohHelp_username_(),
    password: get_SwgohHelp_password_(),
  };
  const client = new swgohhelpapi.exports.Client(settings);
  const allycode = get_SwgohHelp_allycode_();
  const json: swgohhelpapi.exports.GuildResponse = client.fetchGuild({
    allycode,
    roster: true,
    units: true,
    language: swgohhelpapi.Languages.eng_us,
    project: {
      id: true,
      roster: true,
    },
  });
  const roster = json.roster as swgohhelpapi.exports.UnitsResponse;
  for (const baseId in roster) {
    const instances = roster[baseId];
    instances.forEach((o) => {
      const playerId = `p/${o.allyCode}`;
      const pname = o.player;
      let member: PlayerData;
      const q = {
        base_id: baseId,
        level: o.level,
        power: o.gp,
        rarity: o.starLevel,
        gear_level: o.gearLevel || 0,
      };
      const index = members.findIndex(m => m.link === playerId);
      if (index > -1) {
        member = members[index];
      } else {
        member = {
          name: pname,
          level: 0,
          gp: 0,
          ships_gp: 0,
          heroes_gp: 0,
          units: {},
          link: `p/${o.allyCode}`,
        };
        members.push(member);
      }
      member.units[baseId] = q;
      member.level = Math.max(member.level, q.level);
      member.gp += q.power;
      if (o.type === 1) {
        member.heroes_gp += q.power;
      } else {
        member.ships_gp += q.power;
      }
    });
  }

  return members;
}

function test_Player() {
  const settings = {
    username: get_SwgohHelp_username_(),
    password: get_SwgohHelp_password_(),
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
    username: get_SwgohHelp_username_(),
    password: get_SwgohHelp_password_(),
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
