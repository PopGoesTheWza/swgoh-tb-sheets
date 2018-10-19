// ****************************************
// Platoon Functions
// ****************************************

// var usedHeroes = [];
let PLATOON_PHASES = [];
let PLATOON_HERO_NEEDED_COUNT = [];
let PLATOON_SHIP_NEEDED_COUNT = [];
const MAX_PLATOON_HEROES = 15;
const MAX_PLATOONS = 6;
const MAX_PLATOON_ZONES = 3;
const PLATOON_ZONE_ROW_OFFSET = 18;

/** Custom object for creating custom order to walk through platoons */
function platoonDetails(phase: number, zone: number, num: number) {
  this.zone = zone;
  this.num = num;
  this.row = 2 + zone * PLATOON_ZONE_ROW_OFFSET;
  this.possible = true;
  this.isGround = zone > 0;
  this.skip = zone === 0 && phase < 3;
}

/**
 * Custom object for platoon units
 * hero, hero row, count, player count, player list (player, gear...)
 */
function platoonUnit(name: string, row: number, count: number, pCount:number) {
  this.name = name;
  this.row = row;
  this.count = count;
  this.pCount = pCount;
  this.players = [];
}

/** Initialize the list of Territory names */
function init_platoon_phases_() {

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
  } else if (tagFilter === 'Dark Side') {
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
function setZoneName(phase, zone, sheet, platoonRow) {

  // set the zone name
  const zoneName = PLATOON_PHASES[phase - 1][zone];
  sheet.getRange(platoonRow + 2, 1).setValue(zoneName);

  return zoneName;
}

/** Populate platoon with slices if available */
function fillSlice(phase, zone, platoon, range) {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.SLICES);
  const tagFilter = getSideFilter_();

  // format the cell name
  let cellName = tagFilter === 'Dark Side' ? 'Dark' : 'Light';
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
function resetPlatoon(phase, zone, platoonRow, rows, show) {

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
    range.clearContent()
      .setFontColor('Black');
    fillSlice(phase, zone + 1, platoon, range);

    range = sheet.getRange(platoonRow, 5 + platoon * 4, MAX_PLATOON_HEROES, 1);
    range.clearContent();
    range.clearDataValidations();
    range.setFontColor('Black');
  }
}

/** Clear the full chart */
function resetPlatoons() {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = sheet.getRange(2, 1).getValue();
  init_platoon_phases_();

  // Territory 1 (Air)
  let platoonRow = 2;
  let zone = 0;
  resetPlatoon(phase, zone, platoonRow, MAX_PLATOON_HEROES, phase >= 3);
  setZoneName(phase, zone, sheet, platoonRow);
  zone += 1;

  // Territory 2
  platoonRow = 20;
  resetPlatoon(phase, zone, platoonRow, MAX_PLATOON_HEROES, phase >= 1);
  setZoneName(phase, zone, sheet, platoonRow);
  zone += 1;

  // Territory 3
  platoonRow = 38;
  resetPlatoon(phase, zone, platoonRow, MAX_PLATOON_HEROES, phase >= 1);
  setZoneName(phase, zone, sheet, platoonRow);
  // zone += 1;
}

/** Check if the player is available for the current phase */
function playerAvailable(player, unavailable) {

  return unavailable.some((e) => {
    return e[0].length > -1 && player === e[0];
  });
}

/** Get a sorted list of recommended players */
function getRecommendedPlayers(unitName, phase, data, isHero, unavailable) {

  // see how many stars are needed
  const minStars = phase + 1;

  const rec = [];

  if (unitName.length === 0) {
    // no unit selected
    return rec;
  }

  // find the hero in the list
  let unitRow = -1;
  const guildSize = getGuildSize_();
  for (let h = 1, hLen = data.length; h < hLen; h += 1) {
    if (unitName === data[h][0]) {
      // increment the number of times this unit was needed
      if (isHero) {
        PLATOON_HERO_NEEDED_COUNT[h - 1][0] += 1;
      } else {
        PLATOON_SHIP_NEEDED_COUNT[h - 1][0] += 1;
      }

      // found the unit, now get the recommendations
      for (let p = 0; p < guildSize; p += 1) {
        const playerIdx = HERO_PLAYER_COL_OFFSET + p;
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

        const playerStars = Number(playerUnit[0]);
        if (playerStars >= minStars) {
          const power = parseInt(getSubstringRe_(playerUnit, /P(.*)/), 10);
          rec[rec.length] = [playerName, power];
        }
      }

      // finished with the hero, so break
      unitRow = h;
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
function createDropdown(playerList, range) {

  const formatList = [];
  for (let p = 0, pLen = playerList.length; p < pLen; p += 1) {
    formatList[formatList.length] = playerList[p][0];
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(formatList)
    .build();
  range.setDataValidation(rule);
}

/** Reset the needed counts */
function resetNeededCount(sheet, count) {

  const result = [];
  for (let i = 0; i < count; i += 1) {
    result[i] = [0];
  }

  return result;
}

/** Reset the units used */
function resetUsedUnits(data) {

  const result = [];
  for (let r = 0, rLen = data.length; r < rLen; r += 1) {
    if (r === 0) {
      // first row, so copy it all
      result[r] = data[r];
    } else {
      result[r] = [];
      for (let c = 0, cLen = data[r].length; c < cLen; c += 1) {
        result[r].push(false);
      }
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
  let heroData = heroesSheet
    .getRange(1, 1, 1 + heroCount, HERO_PLAYER_COL_OFFSET + getGuildSize_())
    .getValues() as string[][];
  let shipData = shipsSheet
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
  PLATOON_HERO_NEEDED_COUNT = resetNeededCount(heroesSheet, heroCount);
  PLATOON_SHIP_NEEDED_COUNT = resetNeededCount(shipsSheet, shipCount);

  // reset the used heroes
  let usedHeroes = [];
  let usedShips = [];
  usedHeroes = resetUsedUnits(heroData);
  usedShips = resetUsedUnits(shipData);

  // setup platoon phases
  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const unavailable = sheet.getRange(56, 4, getGuildSize_(), 1).getValues();
  const phase = sheet.getRange(2, 1).getValue() as number;
  init_platoon_phases_();

  // setup a custom order for walking the platoons
  const platoonOrder = [];
  for (let p = MAX_PLATOONS - 1; p >= 0; p -= 1) {
    // last platoon to first platoon
    for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
      // zone by zone
      platoonOrder.push(new platoonDetails(phase, z, p));
    }
  }

  // set the zone names and show/hide rows
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    // for each zone
    const platoonRow = 2 + z * PLATOON_ZONE_ROW_OFFSET;

    // set the zone name
    const zoneName = setZoneName(phase, z, sheet, platoonRow);

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
  const platoonMatrix = [];
  const baseCol = 4;
  for (let o = 0, oLen = platoonOrder.length; o < oLen; o += 1) {
    const cur = platoonOrder[o];

    const platoonOffset = cur.num * 4;
    const platoonRange = sheet.getRange(cur.row, baseCol + platoonOffset);

    // clear previous contents
    const range = sheet.getRange(
      cur.row,
      baseCol + 1 + platoonOffset,
      MAX_PLATOON_HEROES,
      1,
    );
    range.clearContent();
    range.clearDataValidations();
    range.setFontColor('Black');
    range.offset(0, -1).setFontColor('Black');

    if (cur.skip) {
      // skip this platoon
      continue;
    }

    const skip =
      sheet
        .getRange(cur.row + 15, baseCol + platoonOffset + 1, 1, 1)
        .getValue() === 'SKIP';
    if (skip) {
      cur.possible = false;
    }

    // cycle through the units
    const units = sheet
      .getRange(cur.row, baseCol + platoonOffset, MAX_PLATOON_HEROES, 1)
      .getValues() as string[][];
    for (let h = 0; h < MAX_PLATOON_HEROES; h += 1) {
      const unitName = units[h][0];
      const idx = platoonMatrix.length;

      const playersRange = sheet.getRange(
        cur.row + h,
        baseCol + 1 + platoonOffset,
      );
      if (unitName.length === 0) {
        // no unit was entered, so skip it
        platoonMatrix[idx] = new platoonUnit(unitName, 0, 0, 0);
        continue;
      }

      const rec = getRecommendedPlayers(
        unitName,
        phase,
        cur.isGround ? heroData : shipData,
        cur.isGround,
        unavailable,
      );
      platoonMatrix[idx] = new platoonUnit(unitName, 0, 0, rec.length);
      if (rec.length > 0) {
        createDropdown(rec, playersRange);

        // add the players to the matrix
        for (let r = 0, rLen = rec.length; r < rLen; r += 1) {
          platoonMatrix[idx].players.push(rec[r][0]); // player name
        }
      } else {
        // impossible to fill the platoon if no one can donate
        cur.possible = false;
        sheet.getRange(cur.row + h, baseCol + platoonOffset).setFontColor('Red');
      }
    }
  }

  // update the unit counts
  for (let m = 0, mLen = platoonMatrix.length; m < mLen; m += 1) {
    // find the unit's count
    for (let h = 1, hLen = heroData.length; h < hLen; h += 1) {
      if (heroData[h][0] === platoonMatrix[m].name) {
        platoonMatrix[m].row = h;
        platoonMatrix[m].count = PLATOON_HERO_NEEDED_COUNT[h - 1][0];
        break;
      }
    }

    // find the unit's count
    for (let h = 1, hLen = shipData.length; h < hLen; h += 1) {
      if (shipData[h][0] === platoonMatrix[m].name) {
        platoonMatrix[m].row = h;
        platoonMatrix[m].count = PLATOON_SHIP_NEEDED_COUNT[h - 1][0];
        break;
      }
    }
  }

  // make sure the platoon is possible to fill
  for (let o = 0, oLen = platoonOrder.length; o < oLen; o += 1) {
    const cur = platoonOrder[o];
    if (cur.skip) {
      // skip this platoon
      continue;
    }

    // const idx = cur.num + cur.zone * MAX_PLATOON_HEROES
    if (cur.possible === false) {
      const plattonOffset = cur.num * 4;
      const platoonRange = sheet.getRange(
        cur.row,
        baseCol + 1 + plattonOffset,
        MAX_PLATOON_HEROES,
        1,
      );
      platoonRange.setValue('Skip');
      platoonRange.clearDataValidations();
      platoonRange.setFontColor('Red');
    }
  }

  // initialize the placement counts
  const placementCount = [];
  const maxPlayerDonations = getMaximumPlatoonDonation_();
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    placementCount[z] = [];
  }

  // try to find an unused player to default to
  let matrixIdx = 0;
  for (let o = 0, oLen = platoonOrder.length; o < oLen; o += 1) {
    const cur = platoonOrder[o];

    if (cur.skip) {
      // skip this zone
      continue;
    }

    // const idx = cur.num + cur.zone * MAX_PLATOON_HEROES
    if (cur.possible === false) {
      // skip this platoon
      matrixIdx += MAX_PLATOON_HEROES;
      continue;
    }

    // cycle through the heroes
    for (let h = 0; h < MAX_PLATOON_HEROES; h += 1) {
      const plattonOffset = cur.num * 4;
      const platoonRange = sheet.getRange(
        cur.row + h,
        baseCol + 1 + plattonOffset,
      );

      let defaultValue = '';
      for (
        let playerIdx = 0, playerLen = platoonMatrix[matrixIdx].players.length;
        playerIdx < playerLen;
        playerIdx += 1
      ) {
        // see if the recommended player's hero has been used
        const heroRow = platoonMatrix[matrixIdx].row;
        if (cur.isGround) {
          // ground units
          for (let u = 1, uLen = usedHeroes[heroRow].length; u < uLen; u += 1) {
            const playerName = platoonMatrix[matrixIdx].players[playerIdx];
            const playerAvail =
              placementCount[cur.zone][playerName] == null ||
              placementCount[cur.zone][playerName] < maxPlayerDonations;
            if (
              playerAvail &&
              usedHeroes[0][u] === playerName &&
              usedHeroes[heroRow][u] === false
            ) {
              usedHeroes[heroRow][u] = true;
              defaultValue = playerName;
              if (placementCount[cur.zone][playerName] == null) {
                placementCount[cur.zone][playerName] = 0;
              }
              placementCount[cur.zone][playerName] += 1;
              break;
            }
          }
        } else {
          // ships
          for (let u = 1, uLen = usedShips[heroRow].length; u < uLen; u += 1) {
            const playerName = platoonMatrix[matrixIdx].players[playerIdx];
            const playerAvail =
              placementCount[cur.zone][playerName] == null ||
              placementCount[cur.zone][playerName] < maxPlayerDonations;
            if (
              playerAvail &&
              usedShips[0][u] === playerName &&
              usedShips[heroRow][u] === false
            ) {
              usedShips[heroRow][u] = true;
              defaultValue = playerName;
              if (placementCount[cur.zone][playerName] == null) {
                placementCount[cur.zone][playerName] = 0;
              }
              placementCount[cur.zone][playerName] += 1;
              break;
            }
          }
        }

        if (defaultValue.length > 0) {
          // we already have a recommended player
          break;
        }
      }
      platoonRange.setValue(defaultValue);

      // see if we should highlight rare units
      if (platoonMatrix[matrixIdx].count > platoonMatrix[matrixIdx].pCount) {
        // we don't have enough of this hero, so mark it
        const color = 'Red';
        platoonRange.setFontColor(color);
        platoonRange.offset(0, -1).setFontColor(color);
      } else if (
        defaultValue.length > 0 &&
        platoonMatrix[matrixIdx].count + 3 > platoonMatrix[matrixIdx].pCount
      ) {
        // we barely have enough of this hero, so mark it
        const color = 'Blue';
        platoonRange.setFontColor(color);
        platoonRange.offset(0, -1).setFontColor(color);
      }

      matrixIdx += 1;
    }
  }
}
