// ****************************************
// Platoon Functions
// ****************************************

let PLATOON_PHASES: [string, string, string][] = [];
let PLATOON_HERO_NEEDED_COUNT: KeyedNumbers = {};
let PLATOON_SHIP_NEEDED_COUNT: KeyedNumbers = {};

/** Custom object for creating custom order to walk through platoons */
class PlatoonDetails {

  public readonly zone: number;
  public readonly platoon: number;
  public readonly row: number;
  public possible: boolean;
  public readonly isGround: boolean;
  public readonly exist: boolean;

  constructor(phase: number, zone: number, platoon: number) {
    this.zone = zone;
    this.platoon = platoon;
    this.row = 2 + zone * PLATOON_ZONE_ROW_OFFSET;
    this.possible = true;
    this.isGround = zone > 0;
    this.exist = zone === 0 && phase < 3;
  }
  getOffset() {
    return this.platoon * 4;
  }
}

/**
 * Custom object for platoon units
 * hero, count, player count, player list (player, gear...)
 */
class PlatoonUnit {

  public readonly name: string;
  public count: number;
  private readonly pCount: number;
  public players: string[];

  constructor(name: string, count: number, pCount:number) {
    this.name = name;
    this.count = count;
    this.pCount = pCount;
    this.players = [];
  }

  public isMissing(): boolean {
    return this.count > this.pCount;
  }

  public isRare(): boolean {
    return this.count + 3 > this.pCount;
  }

}

/** Initialize the list of Territory names */
function initPlatoonPhases_(): void {

  const filter = getEventFilter_();

  // Names of Territories, # of Platoons
  if (isLight_(filter)) {
    PLATOON_PHASES = [
      ['', 'Rebel Base', ''],
      ['', 'Ion Cannon', 'Overlook'],
      ['Rear Airspace', 'Rear Trenches', 'Power Generator'],
      ['Forward Airspace', 'Forward Trenches', 'Outer Pass'],
      ['Contested Airspace', 'Snowfields', 'Forward Stronghold'],
      ['Imperial Fleet Staging Area', 'Imperial Flank', 'Imperial Landing'],
    ];
  } else if (filter === ALIGNMENT.DARKSIDE) {
    PLATOON_PHASES = [
      ['', 'Imperial Flank', 'Imperial Landing'],
      ['', 'Snowfields', 'Forward Stronghold'],
      ['Imperial Fleet Staging Area', 'Ion Cannon', 'Outer Pass'],
      ['Contested Airspace', 'Power Generator', 'Rear Trenches'],
      ['Forward Airspace', 'Forward Trenches', 'Overlook'],
      [
        'Rear Airspace',
        'Rebel Base Main Entrance',
        'Rebel Base South Entrance',
      ],
    ];
  } else {
    PLATOON_PHASES = [
      ['???', '???', '???'],
      ['???', '???', '???'],
      ['???', '???', '???'],
      ['???', '???', '???'],
      ['???', '???', '???'],
      ['???', '???', '???'],
    ];
  }
}

/** Get the zone name and update the cell */
function setZoneName_(phase: number, zone: number, sheet: Sheet, platoonRow: number): string {

  // set the zone name
  const zoneName = PLATOON_PHASES[phase - 1][zone];
  sheet.getRange(platoonRow + 2, 1).setValue(zoneName);

  return zoneName;
}

/** retrieve current slices */
function readSlice_(phase: number, zone: number): string[][] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.SLICES);
  const filter = getEventFilter_();

  // format the cell name
  let cellName = filter === ALIGNMENT.DARKSIDE ? 'Dark' : 'Light';
  cellName += `Slice${phase}Z${zone + 1}`;

  if (phase > 2 || zone !== 0) {
    const range = sheet.getRange(cellName);
    if (range) {
      return range.getValues() as string[][];
    }
    // try {  // TODO: avoid try/catch
    //   return sheet.getRange(cellName).getValues() as string[][];
    // } catch (e) {}
  }

  return undefined;
}

/** Populate platoon with slices if available */
function writeSlice_(data: string[][], platoon: number, range: Range): void {
  const slice = data.map(e => [e[platoon]]);
  range.setValues(slice);
}

/** Clear out a platoon */
function resetPlatoon_(
  phase: number,
  zone: number,
  platoonRow: number,
  rows: number,
  show: boolean,
): void {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);

  if (show) {
    sheet.showRows(platoonRow, MAX_PLATOON_UNITS);
  } else {
    if (platoonRow === 2) {
      sheet.hideRows(platoonRow + 1, rows - 1);
    } else {
      sheet.hideRows(platoonRow, rows);
    }
  }
  const slice = readSlice_(phase, zone);

  for (let platoon = 0; platoon < MAX_PLATOONS; platoon += 1) {
    // clear the contents
    const col = (platoon * 4) + 4;
    const range = sheet.getRange(platoonRow, col, MAX_PLATOON_UNITS, 2);
    range.clearContent()
        .setFontColor(COLOR.BLACK);

    writeSlice_(slice, platoon, range.offset(0, 0, MAX_PLATOON_UNITS, 1));

    range.offset(0, 1, MAX_PLATOON_UNITS, 1)
      .clearDataValidations();
    // clear 'Skip this' checkbox
    range.offset(15, 1, 1, 1).clearContent();
  }
}

/** Clear the full chart */
function resetPlatoons(): void {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = sheet.getRange(2, 1).getValue() as number;
  initPlatoonPhases_();

  [2, 20, 38].forEach((platoonRow, zone) => {
    resetPlatoon_(phase, zone, platoonRow, MAX_PLATOON_UNITS, phase >= 3);
    setZoneName_(phase, zone, sheet, platoonRow);
  });
}

function getNeededCount_(unitName: string, isHero: boolean) {
  const count = isHero ? PLATOON_HERO_NEEDED_COUNT : PLATOON_SHIP_NEEDED_COUNT;
  if (count[unitName]) {
    count[unitName] += 1;
  } else {
    count[unitName] = 1;
  }
}

/** Get a sorted list of recommended players */
function getRecommendedPlayers_(
  unitName: string,
  phase: number,
  data: KeyedType<UnitInstances>,
): [string, number][] {

  // see how many stars are needed
  const minRarity = phase + 1;

  const rec: [string, number][] = [];

  const members = data[unitName];
  if (members) {
    for (const player in members) {

      const playerUnit = members[player];

      if (playerUnit.rarity >= minRarity) {
        rec.push([player, playerUnit.power]);
      }
    }
  }

  // sort list by power
  const playerList = rec.sort((a, b) => a[1] - b[1]);  // sorts by 2nd element ascending

  return playerList;
}

/** create the dropdown list */
function buildDropdown_(playerList: [string, number][]): DataValidation {

  const formatList = playerList.map(e => e[0]);

  return SpreadsheetApp.newDataValidation()
    .requireValueInList(formatList)
    .build();
}

/** Reset the units used */
function resetUsedUnits_(data: KeyedType<UnitInstances>): KeyedType<KeyedBooleans> {

  const result: KeyedType<KeyedBooleans> = {};

  for (const unit in data) {
    const members = data[unit];
    for (const player in members) {
      if (!result[unit]) {
        result[unit] = {};
      }
      result[unit][player] = false;
    }
  }

  return result;
}

function filterUnits_(data: KeyedType<UnitInstances>, filter: (u: UnitInstance) => boolean) {
  const units = Object.assign({}, data);
  for (const unit in units) {
    const members = Object.assign({}, data[unit]);
    for (const player in members) {
      if (!filter(members[player])) {
        delete data[unit][player];
      }
    }
    if (Object.keys(data[unit]).length === 0) {
      delete data[unit];
    }
  }
}

// function processUnavailable_(data: KeyedType<UnitInstances>, list: [string][]) {
//   for (const unit in data) {
//     const members = Object.assign({}, data[unit]);
//     for (const player in members) {
//       if (list.findIndex(e => e[0] === player) !== -1) {
//         delete data[unit][player];
//       }
//     }
//   }
// }

// function loop1_(
//   cur: PlatoonDetails,
//   platoonMatrix: PlatoonUnit[],
//   sheet: Sheet,
// ) {
//   const baseCol = 4;  // TODO: should it be a setting

//   const row = cur.row;
//   const column = cur.getOffset() + baseCol;
//   const range = sheet.getRange(row, column, MAX_PLATOON_UNITS, 1);

//   // clear previous contents
//   range.offset(0, 1)
//     .clearContent()
//     .clearDataValidations()
//     .offset(0, -1, MAX_PLATOON_UNITS, 2)
//     .setFontColor(COLOR.BLACK);

//   if (cur.exist) {

//     /** 'Skip this' checkbox */
//     const skip = sheet.getRange(row + 15, column + 1)
//       .getValue() === 'SKIP';
//     if (skip) {
//       cur.possible = false;
//     } else {

//       // cycle through the units
//       const units = range.getValues().map((e: string[]) => e[0]);

//       const dropdowns: [DataValidation][] = [];
//       const dropdownsRange = range.offset(0, 1);

//       units.forEach((unitName, h) => {
//         const idx = platoonMatrix.length;
//         if (unitName.length === 0) {  // no unit was entered, so skip it
//           platoonMatrix.push(new PlatoonUnit(unitName, 0, 0));
//           dropdowns.push([null]);
//         } else {
//           getNeededCount_(unitName, cur.isGround);

//           const rec = getRecommendedPlayers_(
//             unitName,
//             phase,
//             cur.isGround ? allHeroes : allShips,
//           );

//           platoonMatrix.push(new PlatoonUnit(unitName, 0, rec.length));

//           if (rec.length > 0) {
//             dropdowns.push([buildDropdown_(rec)]);

//             // add the players to the matrix
//             platoonMatrix[idx].players = rec.map(r => r[0]); // player name
//           } else {
//             dropdowns.push([null]);
//             // impossible to fill the platoon if no one can donate
//             cur.possible = false;
//             sheet.getRange(row + h, column).setFontColor(COLOR.RED);
//           }
//         }

//       });
//       allDropdowns.push([dropdownsRange, dropdowns]);
//     }
//   }
// }

// function loop2_(cur: PlatoonDetails, sheet: Sheet) {
//   if (cur.exist && !cur.possible) {
//     const baseCol = 4;  // TODO: should it be a setting

//     const row = cur.row;
//     const column = cur.getOffset() + baseCol;
//     const range = sheet.getRange(row, column + 1, MAX_PLATOON_UNITS, 1);

//     range.setValue('Skip')
//       .clearDataValidations()
//       .setFontColor(COLOR.RED);
//   }
// }

// function loop3_(cur: PlatoonDetails, sheet: Sheet) {
//   if (cur.exist) {
//     if (!cur.possible) {
//       // skip this platoon
//       matrixIdx += MAX_PLATOON_UNITS;
//     } else {
//       // cycle through the heroes
//       const donors: [string][] = [];
//       const colors: [COLOR, COLOR][] = [];
//       const plattonOffset = cur.getOffset();
//       for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {

//         let defaultValue = '';
//         const count = placementCount[cur.zone];

//         const unit = platoonMatrix[matrixIdx].name;
//         for (const player of platoonMatrix[matrixIdx].players) {

//           const available = count[player] == null || count[player] < maxPlayerDonations;
//           if (!available) {
//             continue;
//           }

//           // see if the recommended player's hero has been used
//           if (cur.isGround) {
//             // ground units
//             if (usedHeroes[unit]
//               && usedHeroes[unit].hasOwnProperty(player)
//               && !usedHeroes[unit][player]
//             ) {
//               usedHeroes[unit][player] = true;
//               defaultValue = player;
//               count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

//               break;
//             }
//           } else {
//             // ships
//             if (usedShips[unit]
//               && usedShips[unit].hasOwnProperty(player)
//               && !usedShips[unit][player]
//             ) {
//               usedShips[unit][player] = true;
//               defaultValue = player;
//               count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

//               break;
//             }
//           }

//           if (defaultValue.length > 0) {
//             // we already have a recommended player
//             break;
//           }
//         }

//         donors.push([defaultValue]);

//         // see if we should highlight rare units
//         if (platoonMatrix[matrixIdx].isMissing()) {
//           colors.push([COLOR.RED, COLOR.RED]);
//         } else if (defaultValue.length > 0 && platoonMatrix[matrixIdx].isRare()) {
//           colors.push([COLOR.BLUE, COLOR.BLUE]);
//         } else {
//           colors.push([COLOR.BLACK, COLOR.BLACK]);
//         }

//         matrixIdx += 1;
//       }
//       const donorsRange = sheet.getRange(cur.row, baseCol + plattonOffset + 1, MAX_PLATOON_UNITS, 1);
//       allDonors.push([donorsRange, donors]);
//       const colorsRange = sheet.getRange(cur.row, baseCol +  plattonOffset, MAX_PLATOON_UNITS, 2);
//       allColors.push([colorsRange, colors]);
//     }
//   }
// }

/** Recommend players for each Platoon */
function recommendPlatoons() {

  // setup platoon phases
  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = sheet.getRange(2, 1).getValue() as number;
  const alignment = getEventFilter_().toLowerCase();
  const unavailable = sheet.getRange(56, 4, getGuildSize_(), 1).getValues() as [string][];

  // cache the matrix of hero data
  const heroesTable = new HeroesTable();
  const allHeroes = heroesTable.getAllInstancesByUnits();
  const shipsTable = new ShipsTable();
  const allShips = shipsTable.getAllInstancesByUnits();

  filterUnits_(
    allHeroes,
    (u: UnitInstance) => {
      return u.rarity > phase
        && u.tags.indexOf(alignment) !== -1
        && unavailable.findIndex(e => e[0] === u.name) !== -1;
    },
  );
  filterUnits_(
    allShips,
    (u: UnitInstance) => {
      return u.rarity > phase
        && unavailable.findIndex(e => e[0] === u.name) !== -1;
    },
  );
  // processUnavailable_(allHeroes, unavailable);
  // processUnavailable_(allShips, unavailable);

  // remove heroes listed on Exclusions sheet
  const exclusionsId = getExclusionId_();
  if (exclusionsId.length > 0) {
    const exclusions = getExclusionList_();
    processExclusions_(allHeroes, exclusions);
    processExclusions_(allShips, exclusions, getEventFilter_());
  }

  initPlatoonPhases_();

  // reset the needed counts
  PLATOON_HERO_NEEDED_COUNT = {};
  PLATOON_SHIP_NEEDED_COUNT = {};

  // reset the used heroes
  const usedHeroes = resetUsedUnits_(allHeroes);
  const usedShips = resetUsedUnits_(allShips);

  // setup a custom order for walking the platoons
  const platoonOrder: PlatoonDetails[] = [];

  for (let platoon = MAX_PLATOONS - 1; platoon >= 0; platoon -= 1) {

    for (let zone = 0; zone < MAX_PLATOON_ZONES; zone += 1) {  // last platoon to first platoon
      platoonOrder.push(new PlatoonDetails(phase, zone, platoon));
    }
  }

  // set the zone names and show/hide rows
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {

    // for each zone
    const platoonRow = 2 + z * PLATOON_ZONE_ROW_OFFSET;

    // set the zone name
    const zoneName = setZoneName_(phase, z, sheet, platoonRow);

    // see if we should skip the zone
    if (z !== 1) {
      if (zoneName.length === 0) {
        const hideOffset = z === 0 ? 1 : 0;
        sheet.hideRows(platoonRow + hideOffset, MAX_PLATOON_UNITS - hideOffset);
      } else {
        sheet.showRows(platoonRow, MAX_PLATOON_UNITS);
      }
    }
  }

  // initialize platoon matrix
  const platoonMatrix: PlatoonUnit[] = [];
  const baseCol = 4;

  const allDropdowns: [Range, [DataValidation][]][] = [];

  for (const cur of platoonOrder) {

    const platoonOffset = cur.getOffset();

    // clear previous contents
    sheet.getRange(cur.row, baseCol + platoonOffset + 1, MAX_PLATOON_UNITS, 1)
      .clearContent()
      .clearDataValidations()
      .offset(0, -1, MAX_PLATOON_UNITS, 2)
      .setFontColor(COLOR.BLACK);

    if (!cur.exist) {  // skip this zone
      continue;
    }

    /** skip this checkbox */
    const skip = sheet.getRange(cur.row + 15, baseCol + platoonOffset + 1, 1, 1)
        .getValue() === 'SKIP';
    if (skip) {
      cur.possible = false;
    }

    // cycle through the units
    const units = sheet.getRange(cur.row, baseCol + platoonOffset, MAX_PLATOON_UNITS, 1)
      .getValues() as string[][];

    const dropdowns: [DataValidation][] = [];
    const dropdownsRange = sheet.getRange(
      cur.row,
      baseCol + platoonOffset + 1,
      MAX_PLATOON_UNITS,
      1,
    );
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {

      const unitName = units[h][0];
      const idx = platoonMatrix.length;

      if (unitName.length === 0) {
        // no unit was entered, so skip it
        platoonMatrix.push(new PlatoonUnit(unitName, 0, 0));
        dropdowns.push([null]);

        continue;
      }

      getNeededCount_(unitName, cur.isGround);

      const rec = getRecommendedPlayers_(
        unitName,
        phase,
        cur.isGround ? allHeroes : allShips,
      );

      platoonMatrix.push(new PlatoonUnit(unitName, 0, rec.length));

      if (rec.length > 0) {
        dropdowns.push([buildDropdown_(rec)]);

        // add the players to the matrix
        platoonMatrix[idx].players = rec.map(r => r[0]); // player name
      } else {
        dropdowns.push([null]);
        // impossible to fill the platoon if no one can donate
        cur.possible = false;
        sheet.getRange(cur.row + h, baseCol + platoonOffset).setFontColor(COLOR.RED);
      }
    }
    allDropdowns.push([dropdownsRange, dropdowns]);
  }

  for (const dropdown of allDropdowns) {
    dropdown[0].setDataValidations(dropdown[1]);
  }

  // update the unit counts
  for (const p of platoonMatrix) {

    const unit = p.name;

    // find the unit's count
    if (PLATOON_HERO_NEEDED_COUNT[unit]) {
      p.count = PLATOON_HERO_NEEDED_COUNT[unit];
    } else if (PLATOON_SHIP_NEEDED_COUNT[unit]) {
      p.count = PLATOON_SHIP_NEEDED_COUNT[unit];
    }
  }

  // make sure the platoon is possible to fill
  for (const cur of platoonOrder) {

    if (!cur.exist) {  // skip this zone
      continue;
    }

    if (!cur.possible) {
      const plattonOffset = cur.getOffset();
      sheet.getRange(cur.row, baseCol + 1 + plattonOffset, MAX_PLATOON_UNITS, 1)
        .setValue('Skip')
        .clearDataValidations()
        .setFontColor(COLOR.RED);
    }
  }

  // initialize the placement counts
  const placementCount = [];
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    placementCount[z] = [];
  }
  const maxPlayerDonations = getMaximumPlatoonDonation_();

  // try to find an unused player to default to
  let matrixIdx = 0;

  const allDonors:[Range, [string][]][] = [];
  const allColors: [Range, [COLOR, COLOR][]][] = [];
  for (const cur of platoonOrder) {

    if (!cur.exist) {  // skip this zone
      continue;
    }

    if (!cur.possible) {
      // skip this platoon
      matrixIdx += MAX_PLATOON_UNITS;
      continue;
    }

    // cycle through the heroes
    const donors: [string][] = [];
    const colors: [COLOR, COLOR][] = [];
    const plattonOffset = cur.getOffset();
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {

      let defaultValue = '';
      const count = placementCount[cur.zone];

      const unit = platoonMatrix[matrixIdx].name;
      for (const player of platoonMatrix[matrixIdx].players) {

        const available = count[player] == null || count[player] < maxPlayerDonations;
        if (!available) {
          continue;
        }

        // see if the recommended player's hero has been used
        if (cur.isGround) {
          // ground units
          if (usedHeroes[unit]
            && usedHeroes[unit].hasOwnProperty(player)
            && !usedHeroes[unit][player]
          ) {
            usedHeroes[unit][player] = true;
            defaultValue = player;
            count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

            break;
          }
        } else {
          // ships
          if (usedShips[unit]
            && usedShips[unit].hasOwnProperty(player)
            && !usedShips[unit][player]
          ) {
            usedShips[unit][player] = true;
            defaultValue = player;
            count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

            break;
          }
        }

        if (defaultValue.length > 0) {
          // we already have a recommended player
          break;
        }
      }

      donors.push([defaultValue]);

      // see if we should highlight rare units
      if (platoonMatrix[matrixIdx].isMissing()) {
        colors.push([COLOR.RED, COLOR.RED]);
      } else if (defaultValue.length > 0 && platoonMatrix[matrixIdx].isRare()) {
        colors.push([COLOR.BLUE, COLOR.BLUE]);
      } else {
        colors.push([COLOR.BLACK, COLOR.BLACK]);
      }

      matrixIdx += 1;
    }
    const donorsRange = sheet.getRange(cur.row, baseCol + plattonOffset + 1, MAX_PLATOON_UNITS, 1);
    allDonors.push([donorsRange, donors]);
    const colorsRange = sheet.getRange(cur.row, baseCol +  plattonOffset, MAX_PLATOON_UNITS, 2);
    allColors.push([colorsRange, colors]);
  }

  for (const donors of allDonors) {
    donors[0].setValues(donors[1]);
  }
  for (const colors of allColors) {
    colors[0].setFontColors(colors[1]);
  }
}
