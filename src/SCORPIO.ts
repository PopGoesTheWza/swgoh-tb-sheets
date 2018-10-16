// *****************************************
// ** Functions to parse SCORPIO JSON Data
// *****************************************

// interface ScorpioUnitInstance {
//   combat_type: number;
//   gearLevel: number;
//   id: number;
//   level: number;
//   player: string;
//   power: number;
//   rarity: number;
// }
// interface ScorpioRoster {
//   [key: string]: ScorpioUnitInstance[];
// }

// function getGuildDataFromScorpio(): PlayerData[] {
//   const metaScorpioLinkCol = 1;
//   const metaScorpioLinkRow = 11;
//   const members: PlayerData[] = [];

//   const link = SPREADSHEET.getSheetByName(SHEETS.META)
//     .getRange(metaScorpioLinkRow, metaScorpioLinkCol)
//     .getValue() as string;
//   if (!link || link.trim().length === 0) {
//     UI.alert(
//       'Unable to find SCORPIO Link',
//       'Check value on Meta tab',
//       UI.ButtonSet.OK,
//     );
//     return [];
//   }
//   let json: ScorpioRoster;

//   try {
//     const params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
//       // followRedirects: true,
//       muteHttpExceptions: true,
//     };
//     const response = UrlFetchApp.fetch(link, params);
//     // const responseObj = {
//     //   getContentText: response.getContentText().split('\n'),
//     //   getHeaders: response.getHeaders(),
//     //   getResponseCode: response.getResponseCode(),
//     // };
//     // if (response.getResponseCode() !== 200) {
//     //   debugger;
//     // }
//     json = JSON.parse(response.getContentText()) as ScorpioRoster;
//   } catch (e) {
//     UI.alert(
//       'Unable to Parse SCORPIO Data',
//       'Check link in Meta tab. It should be a link and not JSON data',
//       UI.ButtonSet.OK,
//     );
//     return [];
//   }

//   for (const unitId in json) {
//     const instances = json[unitId];
//     instances.forEach((o) => {
//       const playerId = `p/${o.id}`;
//       const pname = o.player;
//       const q: UnitInstance = {
//         baseId: unitId,
//         level: o.level,
//         power: o.power,
//         rarity: o.rarity,
//         gearLevel: o.gear_level || 0,
//       };
//       let member: PlayerData;
//       const index = members.findIndex(e => e.link === playerId);
//       if (index > -1) {
//         member = members[index];
//       } else {
//         member = {
//           name: pname,
//           level: 0,
//           gp: 0,
//           shipsGp: 0,
//           heroesGp: 0,
//           units: {},
//           allyCode: o.id,
//           // link: `p/${o.id}`,
//         };
//         members.push(member);
//       }
//       member.units[unitId] = q;
//       member.level = Math.max(member.level, q.level);
//       member.gp += q.power;
//       if (o.combat_type === 1) {
//         member.heroesGp += q.power;
//       } else {
//         member.shipsGp += q.power;
//       }
//     });
//   }
//   return members;
// }
