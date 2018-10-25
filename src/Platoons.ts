// ****************************************
// Platoon Functions
// ****************************************

// var usedHeroes = [];
let PLATOON_PHASES: [string, string, string][] = [];
let PLATOON_HERO_NEEDED_COUNT: number[][] = [];
let PLATOON_SHIP_NEEDED_COUNT: number[][] = [];

/** Custom object for creating custom order to walk through platoons */
class PlatoonDetails {

  public readonly zone: number;
  public readonly num: number;
  public readonly row: number;
  public possible: boolean;
  public readonly isGround: boolean;
  public readonly skip: boolean;

  constructor(phase: number, zone: number, num: number) {
    this.zone = zone;
    this.num = num;
    this.row = 2 + zone * PLATOON_ZONE_ROW_OFFSET;
    this.possible = true;
    this.isGround = zone > 0;
    this.skip = zone === 0 && phase < 3;
  }
}

/**
 * Custom object for platoon units
 * hero, hero row, count, player count, player list (player, gear...)
 */
class PlatoonUnit {

  public readonly name: string;
  public row: number;
  public count: number;
  private readonly pCount: number;
  public readonly players: string[];

  constructor(name: string, row: number, count: number, pCount:number) {
    this.name = name;
    this.row = row;
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

  const tagFilter = getSideFilter_();

  // Names of Territories, # of Platoons
  if (isLight_(tagFilter)) {
    PLATOON_PHASES = [
      ['', 'Rebel Base', ''],
      ['', 'Ion Cannon', 'Overlook'],
      ['Rear Airspace', 'Rear Trenches', 'Power Generator'],
      ['Forward Airspace', 'Forward Trenches', 'Outer Pass'],
      ['Contested Airspace', 'Snowfields', 'Forward Stronghold'],
      ['Imperial Fleet Staging Area', 'Imperial Flank', 'Imperial Landing'],
    ];
  } else if (tagFilter === ALIGNMENT.DARKSIDE) {
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
function setZoneName_(
  phase: number,
  zone: number,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  platoonRow: number,
): string {

  // set the zone name
  const zoneName = PLATOON_PHASES[phase - 1][zone];
  sheet.getRange(platoonRow + 2, 1).setValue(zoneName);

  return zoneName;
}

/** Populate platoon with slices if available */
function fillSlice(
  phase: number,
  zone: number,
  platoon: number,
  range: GoogleAppsScript.Spreadsheet.Range,
): void {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.SLICES);
  const tagFilter = getSideFilter_();

  // format the cell name
  let cellName = tagFilter === ALIGNMENT.DARKSIDE ? 'Dark' : 'Light';
  cellName += `Slice${phase}Z${zone}`;

  if (phase < 3 && zone === 1) {
    return;
  }

  try {
    const data = sheet.getRange(cellName).getValues();
    const slice = [];

    // format the data
    for (let r = 0, rLen = data.length; r < rLen; r += 1) {
      slice[r] = [data[r][platoon]];
    }
    range.setValues(slice);
  } catch (e) {}
}

/** Clear out a platoon */
function resetPlatoon(
  phase: number,
  zone: number,
  platoonRow: number,
  rows: number,
  show: boolean,
): void {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);

  if (show) {
    sheet.showRows(platoonRow, MAX_PLATOON_HEROES);
  } else {
    if (platoonRow === 2) {
      sheet.hideRows(platoonRow + 1, rows - 1);
    } else {
      sheet.hideRows(platoonRow, rows);
    }
  }

  for (let platoon = 0; platoon < MAX_PLATOONS; platoon += 1) {
    // clear the contents
    let range = sheet.getRange(platoonRow, 4 + platoon * 4, MAX_PLATOON_HEROES, 1);
    range.clearContent();
    range.setFontColor(COLOR.BLACK);
    fillSlice(phase, zone + 1, platoon, range);

    range = sheet.getRange(platoonRow, 5 + platoon * 4, MAX_PLATOON_HEROES, 1);
    range.clearContent();
    range.clearDataValidations();
    range.setFontColor(COLOR.BLACK);
  }
}

/** Clear the full chart */
function resetPlatoons(): void {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = sheet.getRange(2, 1).getValue() as number;
  initPlatoonPhases_();

  // Territory 1 (Air)
  let platoonRow = 2;
  let zone = 0;
  resetPlatoon(phase, zone, platoonRow, MAX_PLATOON_HEROES, phase >= 3);
  setZoneName_(phase, zone, sheet, platoonRow);
  zone += 1;

  // Territory 2
  platoonRow = 20;
  resetPlatoon(phase, zone, platoonRow, MAX_PLATOON_HEROES, phase >= 1);
  setZoneName_(phase, zone, sheet, platoonRow);
  zone += 1;

  // Territory 3
  platoonRow = 38;
  resetPlatoon(phase, zone, platoonRow, MAX_PLATOON_HEROES, phase >= 1);
  setZoneName_(phase, zone, sheet, platoonRow);
  // zone += 1;
}

/** Check if the player is available for the current phase */
function playerAvailable(player: string, unavailable: string[][]): boolean {
  return unavailable.some(e => e[0].length > -1 && player === e[0]);
}

/** Get a sorted list of recommended players */
function getRecommendedPlayers_(
  unitName: string,
  phase: number,
  data: string[][],
  isHero: boolean,
  unavailable: string[][],
): [string, number][] {

  // see how many stars are needed
  const minStars = phase + 1;

  const rec: [string, number][] = [];

  if (unitName.length === 0) {  // no unit selected
    return rec;
  }

  // find the hero in the list
  const guildSize = getGuildSize_();

  for (let h = 1, hLen = data.length; h < hLen; h += 1) {

    if (unitName === data[h][0]) {

      // increment the number of times this unit was needed
      (isHero ? PLATOON_HERO_NEEDED_COUNT : PLATOON_SHIP_NEEDED_COUNT)[h - 1][0] += 1;

      // found the unit, now get the recommendations
      for (let p = 0; p < guildSize; p += 1) {

        const playerIdx = HERO_PLAYER_COL_OFFSET + p - 1;
        const playerName = data[0][playerIdx];

        if (playerAvailable(playerName, unavailable)) {
          // we shouldn't use this player
          continue;
        }

        const playerUnit = data[h][playerIdx];
        if (playerUnit.length === 0) {
          // player doesn't own the unit
          continue;
        }

        const playerStars = Number(playerUnit[0]);  // weak
        if (playerStars >= minStars) {
          const power = parseInt(getSubstringRe_(playerUnit, /P(.*)/), 10);
          rec.push([playerName, power]);
        }
      }

      // finished with the hero, so break
      break;
    }
  }

  // sort list by power
  const playerList = rec.sort((a, b) => {
    return a[1] - b[1]; // sorts by 2nd element ascending
  });

  return playerList;
}

/** create the dropdown list */
function buildDropdown_(
  playerList: [string, number][],
): GoogleAppsScript.Spreadsheet.DataValidation {

  const formatList = playerList.map(e => e[0]);

  return SpreadsheetApp.newDataValidation()
    .requireValueInList(formatList)
    .build();
}

/** Reset the needed counts */
function resetNeededCount(count: number): [number][] {

  const result = [];
  for (let i = 0; i < count; i += 1) {
    result[i] = [0];
  }

  return result;
}

/** Reset the units used */
function resetUsedUnits(data: string[][]): (string|boolean)[][] {

  const result = [];

  for (let r = 0, rLen = data.length; r < rLen; r += 1) {

    if (r === 0) {
      // first row, so copy it all
      result[r] = data[r];
    } else {
      result[r] = Array(data[r].length).fill(false);
    }
  }

  return result;
}

/** Recommend players for each Platoon */
function recommendPlatoons() {

  const heroesSheet = SPREADSHEET.getSheetByName(SHEETS.HEROES);
  const shipsSheet = SPREADSHEET.getSheetByName(SHEETS.SHIPS);

  // see how many heroes are listed
  const heroCount = getCharacterCount_();
  const shipCount = getShipCount_();

  // cache the matrix of hero data
  let heroData: string[][] = heroesSheet
    .getRange(1, 1, 1 + heroCount, HERO_PLAYER_COL_OFFSET + getGuildSize_())
    .getValues() as string[][];
  let shipData: string[][] = shipsSheet
    .getRange(1, 1, 1 + shipCount, SHIP_PLAYER_COL_OFFSET + getGuildSize_())
    .getValues() as string[][];

  // remove heroes listed on Exclusions sheet
  const exclusionsId = getExclusionId_();
  if (exclusionsId.length > 0) {
    const exclusions = get_exclusions_();
    heroData = processExclusions_(heroData, exclusions);
    shipData = processExclusions_(shipData, exclusions);
  }

  // reset the needed counts
  PLATOON_HERO_NEEDED_COUNT = resetNeededCount(heroCount);
  PLATOON_SHIP_NEEDED_COUNT = resetNeededCount(shipCount);

  // reset the used heroes
  const usedHeroes = resetUsedUnits(heroData);
  const usedShips = resetUsedUnits(shipData);

  // setup platoon phases
  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const unavailable = sheet.getRange(56, 4, getGuildSize_(), 1).getValues() as [string][];
  const phase = sheet.getRange(2, 1).getValue() as number;
  initPlatoonPhases_();

  // setup a custom order for walking the platoons
  const platoonOrder: PlatoonDetails[] = [];

  for (let p = MAX_PLATOONS - 1; p >= 0; p -= 1) {

    // last platoon to first platoon
    for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {

      // zone by zone
      platoonOrder.push(new PlatoonDetails(phase, z, p));
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
        sheet.hideRows(platoonRow + hideOffset, MAX_PLATOON_HEROES - hideOffset);
      } else {
        sheet.showRows(platoonRow, MAX_PLATOON_HEROES);
      }
    }
  }

  // initialize platoon matrix
  // hero, hero row, count, player count, player list (player, gear...)
  const platoonMatrix: PlatoonUnit[] = [];
  const baseCol = 4;

  const allDropdowns: [
    GoogleAppsScript.Spreadsheet.Range,
    [GoogleAppsScript.Spreadsheet.DataValidation][]
  ][] = [];
  for (let o = 0, oLen = platoonOrder.length; o < oLen; o += 1) {

    const cur = platoonOrder[o];
    const platoonOffset = cur.num * 4;

    // clear previous contents
    sheet.getRange(cur.row, baseCol + platoonOffset + 1, MAX_PLATOON_HEROES, 1)
      .clearContent()
      .clearDataValidations()
      .offset(0, -1, MAX_PLATOON_HEROES, 2)
      .setFontColor(COLOR.BLACK);

    if (cur.skip) {
      // skip this platoon
      continue;
    }

    /** skip this checkbox */
    const skip = sheet.getRange(cur.row + 15, baseCol + platoonOffset + 1, 1, 1)
        .getValue() === 'SKIP';
    if (skip) {
      cur.possible = false;
    }

    // cycle through the units
    const units = sheet.getRange(cur.row, baseCol + platoonOffset, MAX_PLATOON_HEROES, 1)
      .getValues() as string[][];

    const dropdowns: [GoogleAppsScript.Spreadsheet.DataValidation][] = [];
    const dropdownsRange = sheet.getRange(
      cur.row,
      baseCol + platoonOffset + 1,
      MAX_PLATOON_HEROES,
      1,
    );
    for (let h = 0; h < MAX_PLATOON_HEROES; h += 1) {

      const unitName = units[h][0];
      const idx = platoonMatrix.length;

      const playersRange = sheet.getRange(
        cur.row + h,
        baseCol + 1 + platoonOffset,
      );

      if (unitName.length === 0) {
        // no unit was entered, so skip it
        platoonMatrix[idx] = new PlatoonUnit(unitName, 0, 0, 0);
        dropdowns.push([null]);
        continue;
      }

      const rec = getRecommendedPlayers_(
        unitName,
        phase,
        cur.isGround ? heroData : shipData,
        cur.isGround,
        unavailable,
      );
      platoonMatrix[idx] = new PlatoonUnit(unitName, 0, 0, rec.length);

      if (rec.length > 0) {
        dropdowns.push([buildDropdown_(rec)]);

        // add the players to the matrix
        for (const r of rec) {
          platoonMatrix[idx].players.push(r[0]); // player name
        }
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

    // find the unit's count
    for (let h = 1, hLen = heroData.length; h < hLen; h += 1) {

      if (heroData[h][0] === p.name) {
        p.row = h;
        p.count = PLATOON_HERO_NEEDED_COUNT[h - 1][0];
        break;
      }
    }

    // find the unit's count
    for (let h = 1, hLen = shipData.length; h < hLen; h += 1) {
      if (shipData[h][0] === p.name) {
        p.row = h;
        p.count = PLATOON_SHIP_NEEDED_COUNT[h - 1][0];
        break;
      }
    }
  }

  // make sure the platoon is possible to fill
  for (const cur of platoonOrder) {

    if (cur.skip) {  // skip this platoon
      continue;
    }

    if (!cur.possible) {
      const plattonOffset = cur.num * 4;
      sheet.getRange(cur.row, baseCol + 1 + plattonOffset, MAX_PLATOON_HEROES, 1)
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

  const allDonors:[
    GoogleAppsScript.Spreadsheet.Range,
    [string][]
  ][] = [];
  const allColors:[
    GoogleAppsScript.Spreadsheet.Range,
    [COLOR, COLOR][]
  ][] = [];
  for (const cur of platoonOrder) {

    if (cur.skip) {
      // skip this zone
      continue;
    }

    if (cur.possible === false) {
      // skip this platoon
      matrixIdx += MAX_PLATOON_HEROES;
      continue;
    }

    // cycle through the heroes
    const donors: [string][] = [];
    const colors: [COLOR, COLOR][] = [];
    const plattonOffset = cur.num * 4;
    for (let h = 0; h < MAX_PLATOON_HEROES; h += 1) {

      let defaultValue = '';
      const count = placementCount[cur.zone];

      for (const player of platoonMatrix[matrixIdx].players) {

        // see if the recommended player's hero has been used
        const heroRow = platoonMatrix[matrixIdx].row;
        if (cur.isGround) {
          // ground units

          for (let u = 1, uLen = usedHeroes[heroRow].length; u < uLen; u += 1) {

            const available = count[player] == null || count[player] < maxPlayerDonations;
            if (available && usedHeroes[0][u] === player && usedHeroes[heroRow][u] === false) {
              usedHeroes[heroRow][u] = true;
              defaultValue = player;
              count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

              break;
            }
          }
        } else {
          // ships

          for (let u = 1, uLen = usedShips[heroRow].length; u < uLen; u += 1) {

            const available = count[player] == null || count[player] < maxPlayerDonations;

            if (available && usedShips[0][u] === player && usedShips[heroRow][u] === false) {
              usedShips[heroRow][u] = true;
              defaultValue = player;
              count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

              break;
            }
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
    const donorsRange = sheet.getRange(cur.row, baseCol + plattonOffset + 1, MAX_PLATOON_HEROES, 1);
    allDonors.push([donorsRange, donors]);
    const colorsRange = sheet.getRange(cur.row, baseCol +  plattonOffset, MAX_PLATOON_HEROES, 2);
    allColors.push([colorsRange, colors]);
  }
  for (const donors of allDonors) {
    donors[0].setValues(donors[1]);
  }
  for (const colors of allColors) {
    colors[0].setFontColors(colors[1]);
  }
}
