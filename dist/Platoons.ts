// ****************************************
// Platoon Functions
// ****************************************

//var UsedHeroes = [];
var PlatoonPhases = [];
var PlatoonHeroNeededCount = [];
var PlatoonShipNeededCount = [];
var MaxPlatoonHeroes = 15;
var MaxPlatoons = 6;
var MaxPlatoonZones = 3;
var PlatoonZoneRowOffset = 18;

// Initialize the list of Territory names
function init_platoon_phases_() {
  var tag_filter = get_tag_filter_();

  // Names of Territories, # of Platoons
  if (tag_filter === "Light Side") {
    PlatoonPhases = [
      ["", "Rebel Base", ""],
      ["", "Ion Cannon", "Overlook"],
      ["Rear Airspace", "Rear Trenches", "Power Generator"],
      ["Forward Airspace", "Forward Trenches", "Outer Pass"],
      ["Contested Airspace", "Snowfields", "Forward Stronghold"],
      ["Imperial Fleet Staging Area", "Imperial Flank", "Imperial Landing"]
    ];
  } else if (tag_filter === "Dark Side") {
    PlatoonPhases = [
      ["", "Imperial Flank", "Imperial Landing"],
      ["", "Snowfields", "Forward Stronghold"],
      ["Imperial Fleet Staging Area", "Ion Cannon", "Outer Pass"],
      ["Contested Airspace", "Power Generator", "Rear Trenches"],
      ["Forward Airspace", "Forward Trenches", "Overlook"],
      ["Rear Airspace", "Rebel Base Main Entrance", "Rebel Base South Entrance"]
    ];
  } else {
    PlatoonPhases = [
      ["???", "???", "???"],
      ["???", "???", "???"],
      ["???", "???", "???"],
      ["???", "???", "???"],
      ["???", "???", "???"],
      ["???", "???", "???"]
    ];
  }
}

// Get the zone name and update the cell
function SetZoneName(phase, zone, sheet, platoonRow) {
  // set the zone name
  var zoneName = PlatoonPhases[phase - 1][zone];
  sheet.getRange(platoonRow + 2, 1).setValue(zoneName);

  return zoneName;
}

// Populate platoon with slices if available
function FillSlice(phase, zone, platoon, range) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Slices");
  var tag_filter = get_tag_filter_();

  // format the cell name
  var cellName = tag_filter == "Dark Side" ? "Dark" : "Light";
  cellName += "Slice" + phase + "Z" + zone;

  if (phase < 3 && zone == 1) {
    return;
  }

  try {
    var data = sheet.getRange(cellName).getValues();
    var slice = [];

    // format the data
    for (var r = 0, rLen = data.length; r < rLen; ++r) {
      slice[r] = [data[r][platoon]];
    }
    range.setValues(slice);
  } catch (e) {}
}

// Clear out a platoon
function ResetPlatoon(phase, zone, platoonRow, rows, show) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");

  if (show) {
    sheet.showRows(platoonRow, MaxPlatoonHeroes);
  } else {
    if (platoonRow == 2) {
      sheet.hideRows(platoonRow + 1, rows - 1);
    } else {
      sheet.hideRows(platoonRow, rows);
    }
  }

  for (var platoon = 0; platoon < MaxPlatoons; ++platoon) {
    // clear the contents
    var range = sheet.getRange(
      platoonRow,
      4 + platoon * 4,
      MaxPlatoonHeroes,
      1
    );
    range.setValue(null);
    range.setFontColor("Black");
    FillSlice(phase, zone + 1, platoon, range);

    range = sheet.getRange(platoonRow, 5 + platoon * 4, MaxPlatoonHeroes, 1);
    range.clearContent();
    range.setFontColor("Black");
    range.clearDataValidations();
  }
}

// Clear the full chart
function ResetPlatoons() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");
  var phase = sheet.getRange(2, 1).getValue();
  init_platoon_phases_();

  // Territory 1 (Air)
  var platoonRow = 2;
  var zone = 0;
  ResetPlatoon(phase, zone, platoonRow, MaxPlatoonHeroes, phase >= 3);
  SetZoneName(phase, zone++, sheet, platoonRow);

  // Territory 2
  platoonRow = 20;
  ResetPlatoon(phase, zone, platoonRow, MaxPlatoonHeroes, phase >= 1);
  SetZoneName(phase, zone++, sheet, platoonRow);

  // Territory 3
  platoonRow = 38;
  ResetPlatoon(phase, zone, platoonRow, MaxPlatoonHeroes, phase >= 1);
  SetZoneName(phase, zone++, sheet, platoonRow);
}

// Check if the player is available for the current phase
function PlayerAvailable(player, unavailable) {
  return unavailable.some(function(e) {
    return e[0].length > -1 && player == e[0];
  });
}

// Get a sorted list of recommended players
function GetRecommendedPlayers(unitName, phase, data, isHero, unavailable) {
  // see how many stars are needed
  var minStars = phase + 1;

  var rec = [];

  if (unitName.length == 0) {
    // no unit selected
    return rec;
  }

  // find the hero in the list
  var unitRow = -1;
  for (var h = 1, hLen = data.length; h < hLen; ++h) {
    if (unitName == data[h][0]) {
      // increment the number of times this unit was needed
      if (isHero) {
        PlatoonHeroNeededCount[h - 1][0]++;
      } else {
        PlatoonShipNeededCount[h - 1][0]++;
      }

      // found the unit, now get the recommendations
      for (var p = 0; p < MaxPlayers; ++p) {
        var playerIdx = HeroPlayerColOffset + p;
        var playerName = data[0][playerIdx];
        if (PlayerAvailable(playerName, unavailable)) {
          // we shouldn't use this player
          continue;
        }

        var playerUnit = data[h][playerIdx];
        if (playerUnit.length == 0) {
          // player doesn't own the unit
          continue;
        }

        var playerStars = Number(playerUnit[0]);
        if (playerStars >= minStars) {
          //          var power = Number(GetSubstring(playerUnit, "P", null));
          var power = Number.parseInt(get_substring_re_(playerUnit, /P(.*)/));
          rec[rec.length] = [playerName, power];
        }
      }

      // finished with the hero, so break
      unitRow = h;
      break;
    }
  }

  // sort list by power
  var playerList = rec.sort(function(a, b) {
    return a[1] - b[1]; //sorts by 2nd element ascending
  });

  return playerList;
}

// create the dropdown list
function CreateDropdown(playerList, range) {
  var formatList = [];
  for (var p = 0, pLen = playerList.length; p < pLen; ++p) {
    formatList[formatList.length] = playerList[p][0];
  }

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(formatList)
    .build();
  range.setDataValidation(rule);
}

// Reset the needed counts
function ResetNeededCount(sheet, count) {
  var result = [];
  for (var i = 0; i < count; ++i) {
    result[i] = [0];
  }

  return result;
}

// Reset the units used
function ResetUsedUnits(data) {
  var result = [];
  for (var r = 0, rLen = data.length; r < rLen; ++r) {
    if (r == 0) {
      // first row, so copy it all
      result[r] = data[r];
    } else {
      result[r] = [];
      for (var c = 0, cLen = data[r].length; c < cLen; ++c) {
        result[r].push(false);
      }
    }
  }

  return result;
}

// Custom object for creating custom order to walk through platoons
function PlatoonDetails(phase, zone, num) {
  this.zone = zone;
  this.num = num;
  this.row = 2 + zone * PlatoonZoneRowOffset;
  this.possible = true;
  this.isGround = zone > 0;
  this.skip = zone == 0 && phase < 3;
}

// Custom object for platoon units
// hero, hero row, count, player count, player list (player, gear...)
function PlatoonUnit(name, row, count, pCount) {
  this.name = name;
  this.row = row;
  this.count = count;
  this.pCount = pCount;
  this.players = [];
}

// Recommend players for each Platoon
function RecommendPlatoons() {
  var heroesSheet = SpreadsheetApp.getActive().getSheetByName("Heroes");
  var shipsSheet = SpreadsheetApp.getActive().getSheetByName("Ships");

  // see how many heroes are listed
  var heroCount = get_character_count_();
  var shipCount = get_ship_count_();

  // cache the matrix of hero data
  var heroData = heroesSheet
    .getRange(1, 1, 1 + heroCount, HeroPlayerColOffset + MaxPlayers)
    .getValues();
  var shipData = shipsSheet
    .getRange(1, 1, 1 + shipCount, ShipPlayerColOffset + MaxPlayers)
    .getValues();

  // remove heroes listed on Exclusions sheet
  var exclusionsId = get_exclusion_id_();
  if (exclusionsId.length > 0) {
    var excludeData = get_exclusions_();
    heroData = ProcessExclusions(heroData, excludeData);
    shipData = ProcessExclusions(shipData, excludeData);
  }

  // reset the needed counts
  PlatoonHeroNeededCount = ResetNeededCount(heroesSheet, heroCount);
  PlatoonShipNeededCount = ResetNeededCount(shipsSheet, shipCount);

  // reset the used heroes
  var UsedHeroes = [];
  var UsedShips = [];
  UsedHeroes = ResetUsedUnits(heroData);
  UsedShips = ResetUsedUnits(shipData);

  // setup platoon phases
  var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");
  var unavailable = sheet.getRange(56, 4, MaxPlayers, 1).getValues();
  var phase = sheet.getRange(2, 1).getValue();
  init_platoon_phases_();

  // setup a custom order for walking the platoons
  var platoonOrder = [];
  for (var p = MaxPlatoons - 1; p >= 0; --p) {
    // last platoon to first platoon
    for (var z = 0; z < MaxPlatoonZones; ++z) {
      // zone by zone
      platoonOrder.push(new PlatoonDetails(phase, z, p));
    }
  }

  // set the zone names and show/hide rows
  for (var z = 0; z < MaxPlatoonZones; ++z) {
    // for each zone
    var platoonRow = 2 + z * PlatoonZoneRowOffset;

    // set the zone name
    var zoneName = SetZoneName(phase, z, sheet, platoonRow);

    // see if we should skip the zone
    if (z != 1) {
      if (zoneName.length == 0) {
        var hideOffset = z == 0 ? 1 : 0;
        sheet.hideRows(platoonRow + hideOffset, MaxPlatoonHeroes - hideOffset);
      } else {
        sheet.showRows(platoonRow, MaxPlatoonHeroes);
      }
    }
  }

  // initialize platoon matrix
  // hero, hero row, count, player count, player list (player, gear...)
  var PlatoonMatrix = [];
  var baseCol = 4;
  for (var o = 0, oLen = platoonOrder.length; o < oLen; ++o) {
    var cur = platoonOrder[o];

    var platoonOffset = cur.num * 4;
    var platoonRange = sheet.getRange(cur.row, baseCol + platoonOffset);

    // clear previous contents
    var range = sheet.getRange(
      cur.row,
      baseCol + 1 + platoonOffset,
      MaxPlatoonHeroes,
      1
    );
    range.clearContent();
    range.clearDataValidations();
    range.setFontColor("Black");
    range.offset(0, -1).setFontColor("Black");

    if (cur.skip) {
      // skip this platoon
      continue;
    }

    var skip =
      sheet
        .getRange(cur.row + 15, baseCol + platoonOffset + 1, 1, 1)
        .getValue() === "SKIP";
    if (skip) {
      cur.possible = false;
    }

    // cycle through the units
    var units = sheet
      .getRange(cur.row, baseCol + platoonOffset, MaxPlatoonHeroes, 1)
      .getValues();
    for (var h = 0; h < MaxPlatoonHeroes; ++h) {
      var unitName = units[h][0];
      var idx = PlatoonMatrix.length;

      var playersRange = sheet.getRange(
        cur.row + h,
        baseCol + 1 + platoonOffset
      );
      if (unitName.length == 0) {
        // no unit was entered, so skip it
        PlatoonMatrix[idx] = new PlatoonUnit(unitName, 0, 0, 0);
        continue;
      }

      var rec = GetRecommendedPlayers(
        unitName,
        phase,
        cur.isGround ? heroData : shipData,
        cur.isGround,
        unavailable
      );
      PlatoonMatrix[idx] = new PlatoonUnit(unitName, 0, 0, rec.length);
      if (rec.length > 0) {
        CreateDropdown(rec, playersRange);

        // add the players to the matrix
        for (var r = 0, rLen = rec.length; r < rLen; ++r) {
          PlatoonMatrix[idx].players.push(rec[r][0]); // player name
        }
      } else {
        // impossible to fill the platoon if no one can donate
        cur.possible = false;
        sheet
          .getRange(cur.row + h, baseCol + platoonOffset)
          .setFontColor("Red");
      }
    }
  }

  // update the unit counts
  for (var m = 0, mLen = PlatoonMatrix.length; m < mLen; ++m) {
    // find the unit's count
    for (var h = 1, hLen = heroData.length; h < hLen; ++h) {
      if (heroData[h][0] == PlatoonMatrix[m].name) {
        PlatoonMatrix[m].row = h;
        PlatoonMatrix[m].count = PlatoonHeroNeededCount[h - 1][0];
        break;
      }
    }

    // find the unit's count
    for (var h = 1, hLen = shipData.length; h < hLen; ++h) {
      if (shipData[h][0] == PlatoonMatrix[m].name) {
        PlatoonMatrix[m].row = h;
        PlatoonMatrix[m].count = PlatoonShipNeededCount[h - 1][0];
        break;
      }
    }
  }

  // make sure the platoon is possible to fill
  for (var o = 0, oLen = platoonOrder.length; o < oLen; ++o) {
    var cur = platoonOrder[o];
    if (cur.skip) {
      // skip this platoon
      continue;
    }

    var idx = cur.num + cur.zone * MaxPlatoonHeroes;
    if (cur.possible == false) {
      var plattonOffset = cur.num * 4;
      var platoonRange = sheet.getRange(
        cur.row,
        baseCol + 1 + plattonOffset,
        MaxPlatoonHeroes,
        1
      );
      platoonRange.setValue("Skip");
      platoonRange.clearDataValidations();
      platoonRange.setFontColor("Red");
    }
  }

  // initialize the placement counts
  var placementCount = [];
  var maxPlayerDonations = get_maximum_platoon_donation_();
  for (var z = 0; z < MaxPlatoonZones; ++z) {
    placementCount[z] = [];
  }

  // try to find an unused player to default to
  var matrixIdx = 0;
  for (var o = 0, oLen = platoonOrder.length; o < oLen; ++o) {
    var cur = platoonOrder[o];

    if (cur.skip) {
      // skip this zone
      continue;
    }

    var idx = cur.num + cur.zone * MaxPlatoonHeroes;
    if (cur.possible == false) {
      // skip this platoon
      matrixIdx += MaxPlatoonHeroes;
      continue;
    }

    // cycle through the heroes
    for (var h = 0; h < MaxPlatoonHeroes; ++h) {
      var plattonOffset = cur.num * 4;
      var platoonRange = sheet.getRange(
        cur.row + h,
        baseCol + 1 + plattonOffset
      );

      var defaultValue = "";
      for (
        var playerIdx = 0, playerLen = PlatoonMatrix[matrixIdx].players.length;
        playerIdx < playerLen;
        ++playerIdx
      ) {
        // see if the recommended player's hero has been used
        var heroRow = PlatoonMatrix[matrixIdx].row;
        if (cur.isGround) {
          // ground units
          for (var u = 1, uLen = UsedHeroes[heroRow].length; u < uLen; ++u) {
            var playerName = PlatoonMatrix[matrixIdx].players[playerIdx];
            var playerAvail =
              placementCount[cur.zone][playerName] == null ||
              placementCount[cur.zone][playerName] < maxPlayerDonations;
            if (
              playerAvail &&
              UsedHeroes[0][u] == playerName &&
              UsedHeroes[heroRow][u] == false
            ) {
              UsedHeroes[heroRow][u] = true;
              defaultValue = playerName;
              if (placementCount[cur.zone][playerName] == null) {
                placementCount[cur.zone][playerName] = 0;
              }
              placementCount[cur.zone][playerName]++;
              break;
            }
          }
        } else {
          // ships
          for (var u = 1, uLen = UsedShips[heroRow].length; u < uLen; ++u) {
            var playerName = PlatoonMatrix[matrixIdx].players[playerIdx];
            var playerAvail =
              placementCount[cur.zone][playerName] == null ||
              placementCount[cur.zone][playerName] < maxPlayerDonations;
            if (
              playerAvail &&
              UsedShips[0][u] == playerName &&
              UsedShips[heroRow][u] == false
            ) {
              UsedShips[heroRow][u] = true;
              defaultValue = playerName;
              if (placementCount[cur.zone][playerName] == null) {
                placementCount[cur.zone][playerName] = 0;
              }
              placementCount[cur.zone][playerName]++;
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
      if (PlatoonMatrix[matrixIdx].count > PlatoonMatrix[matrixIdx].pCount) {
        // we don't have enough of this hero, so mark it
        var color = "Red";
        platoonRange.setFontColor(color);
        platoonRange.offset(0, -1).setFontColor(color);
      } else if (
        defaultValue.length > 0 &&
        PlatoonMatrix[matrixIdx].count + 3 > PlatoonMatrix[matrixIdx].pCount
      ) {
        // we barely have enough of this hero, so mark it
        var color = "Blue";
        platoonRange.setFontColor(color);
        platoonRange.offset(0, -1).setFontColor(color);
      }

      matrixIdx++;
    }
  }
}
