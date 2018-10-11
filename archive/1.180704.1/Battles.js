// ****************************************
// TB functions
// ****************************************

// set the value and style in a cell
function set_cell_value_(cell, value, bold, align) {

  cell.setFontWeight( bold ? "bold" : "normal" );
  cell.setHorizontalAlignment(align);
  cell.setValue(value);
}

/*
function set_cell_values_(cells, values, bolds, align) {

  cells.setFontWeights( bolds );
  cells.setHorizontalAlignments(aligns);
  cells.setValues(values);
}
*/

// Get the stats for all required heroes
function GetRequiredHeroStats(heroes, meta, playerIdx) {

  var result = [];

  // get stats for the required heroes
  var phaseCount = 0;
  var squadCount = 0;
  var total = 0;
  var lastRequired = false;
  var lastSquad = "0";

  for (var m = 0, mLen = meta.length; m < mLen; ++m) {
    var metaHero = meta[m];

    if (metaHero[0] == "Phase Count:") {
      // phase totals
      result[result.length] = ["", phaseCount, "", "", "", ""];
      total += phaseCount;
      phaseCount = 0;
      squadCount = 0;
      continue;
    } else if (metaHero[0].length == 0) {
      // empty row
      result[result.length] = ["", null, "", "", "", ""];
      continue;
    }

    var squad = metaHero[4];
    if (squad != lastSquad) {
      squadCount = 0;
    }
    lastSquad = squad;

    // find the metaHero in heroes
    var heroFound = false;
    for (var h = 0, hLen = heroes.length; h < hLen; ++h) {
      var hero = heroes[h];

      if (hero[0] == metaHero[0]) {
        heroFound = true;

        if (hero[playerIdx] != null) {
          var requirementsMet = false;

          // see if the hero meets the requirements
          var stars = Number(hero[playerIdx][0]);
          var level = Number.parseInt(get_substring_re_(hero[playerIdx], /L([^G]*)/));
          var gear = Number.parseInt(get_substring_re_(hero[playerIdx], /G([^P]*)/));
          var power = Number.parseInt(get_substring_re_(hero[playerIdx], /P(.*)/));
          if (stars >= Number(metaHero[1]) && gear >= Number(metaHero[2]) && level >= Number(metaHero[3])) {
            requirementsMet = true;
            if (metaHero[5] == "R") {
              if (!lastRequired) {
                squadCount = 0;
              }
            }
            lastRequired = metaHero[5] == "R";
            squadCount++;
            if (squadCount <= 5) {
              phaseCount++;
            }
          }

          // store the hero's data
          result[result.length] = [hero[0], stars, level, gear, power, requirementsMet];
        } else {
          result[result.length] = ["", "", "", "", "", ""];
        }

        break;
      }
    }

    if (heroFound == false) {
      // new hero that may not be on the website
      result[result.length] = ["", "", "", "", "", ""];
    }
  }

  result[result.length] = ["", total, "", "", "", ""];

  return result;
}

// Populate the Territory Battle table with data for a player
function populate_player_tb_(tbSheet, offset, meta, fullHeroes) {

  // get the player link from the Roster sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName("Roster");
  var playerName = sheet.getRange(2 + offset, 2).getValue();
  if (playerName.length == 0) {
    // no player, so exit early
    return null;
  }

  // player heroes
  var pIdx = HeroPlayerColOffset + offset;
  var heroes = GetRequiredHeroStats(fullHeroes, meta, pIdx);
  var result = [];
  result.push([playerName]);
  heroes
  .forEach(function(e) {
    if (e[0].length > 0) {
      if (e[5]) {
        // hero met all requirements, so set the cell to show the hero's stars
        result.push([Number(e[1])]);
      } else {
        // hero did not meet requirements, so show the hero's stars and gear
        result.push([
          Utilities.formatString(
            "%s*L%sG%s",
            e[1],
            e[2],
            e[3]
          )
        ]);
      }
    } else {
      // empty data cell
      result.push([e[1]]);
    }
  });

  return result;
}

// Setup the heroes needed for a TB
function SetupTB(tabName, tag_filter) {

  var metaSheet = SpreadsheetApp.getActive().getSheetByName("Meta");
  var tbSheet = SpreadsheetApp.getActive().getSheetByName(tabName);
  //var heroesSheet = SpreadsheetApp.getActive().getSheetByName("Heroes");

  // make sure the roster is up-to-date

  // Early SCORPIO support
  var heroes;
  var jsonLink = SpreadsheetApp.getActive().getSheetByName("Meta").getRange(44, MetaGuildCol).getValue();
  if (jsonLink && jsonLink.trim && jsonLink.trim().length >= 0) {
    var json = {};
    try {
      var response = UrlFetchApp.fetch(jsonLink);
      var text = response.getContentText();
      json = JSON.parse(text);
    } catch (e) {
    }

    json = fixDuplicatesFromJson_(json);
    GetGuildRosterFromJson_(json);
    heroes = GetGuildUnitsFromJson_(json);
  } else {
    GetGuildRoster();
    heroes = GetGuildUnits();
  }

  // clear the hero data
  tbSheet.getRange(1, 10, 1, MaxPlayers).clearContent();
  tbSheet.getRange(2, 1, 150, 9 + MaxPlayers).clearContent();

  // collect the meta data for the heroes
  var row = 2;
  var isLight = (tag_filter === "Light Side");
  var col = (isLight) ? MetaHeroesCol : MetaHeroesDSCol;
  var tbRow = 2;
  var lastPhase = "1";
  var phaseCount = 0;
  var total = 0;
  var lastSquad = "0";
  var squadCount = 0;
  var curMeta = metaSheet.getRange(row, col, 1, 8);
  var metaData = curMeta.getValues();
  var event = metaData[0][0];
  var phaseList = [];
  while (event.length > 0) {
    var phase = metaData[0][1];
    var name = metaData[0][2];
    var stars = metaData[0][3];
    var gear = metaData[0][4];
    var level = metaData[0][5];
    var squad = metaData[0][6];
    var required = metaData[0][7];

    // see if the phase ended
    if (lastPhase != phase) {
      if (tbRow > 2) {
        // end the phase if this isn't the first row
        var curTb = tbSheet.getRange(tbRow, 3);
        set_cell_value_(curTb, "Phase Count:", true, DocumentApp.HorizontalAlignment.RIGHT);
        curTb.offset(0, 1).setValue(phaseCount);
        curTb.offset(0, 6).setValue(
          Utilities.formatString(
            "=COUNTIF(J%s:BI%s,CONCAT(\">=\",D%s))",
            tbRow,
            tbRow,
            tbRow
          )
        );

        phaseList[phaseList.length] = [lastPhase, tbRow];
        tbRow += 2;
      }
      lastPhase = phase;
      total += phaseCount;
      phaseCount = 0;
    }

    // see if the squad changed
    if (lastSquad != squad) {
      lastSquad = squad;
      squadCount = 0;
    }

    // store the meta data
    var curTb = tbSheet.getRange(tbRow, 1, 1, 9);
    var data = [];
    data[0] = [
      event,
      phase,
      name,
      stars,
      gear,
      level,
      squad,
      required,
      Utilities.formatString(
        "=COUNTIF(J%s:BI%s,CONCAT(\">=\",D%s))",
        tbRow,
        tbRow,
        tbRow
      )
    ];

    curTb.setValues(data);
    curTb = tbSheet.getRange(tbRow, 1);
    set_cell_value_(curTb.offset(0, 2), name, false, DocumentApp.HorizontalAlignment.LEFT);
    tbRow++;
    squadCount++;

    if (squadCount <= 5) {
      phaseCount++;
    }

    // get the next row
    row++;
    curMeta = metaSheet.getRange(row, col, 1, 8);
    metaData = curMeta.getValues();
    event = metaData[0][0];
  }

  var lastHeroRow = tbRow;

  // add the final phase
  var curTb = tbSheet.getRange(tbRow, 3);
  set_cell_value_(curTb, "Phase Count:", true, DocumentApp.HorizontalAlignment.RIGHT);
  curTb.offset(0, 1).setValue(phaseCount);
  curTb.offset(0, 6).setValue(
    Utilities.formatString(
      "=COUNTIF(J%s:BI%s,CONCAT(\">=\",D%s))",
      tbRow,
      tbRow,
      tbRow
    )
  );
  phaseList[phaseList.length] = [phase, tbRow];
  total += phaseCount;
  tbRow += 1;

  // add the total
  curTb = tbSheet.getRange(tbRow, 3);
  set_cell_value_(curTb, "Total:", true, DocumentApp.HorizontalAlignment.RIGHT);
  curTb.offset(0, 1).setValue(total);
  curTb.offset(0, 6).setValue(
    Utilities.formatString(
      "=COUNTIF(J%s:BI%s,CONCAT(\">=\",D%s))",
      tbRow,
      tbRow,
      tbRow
    )
  );

  // add the readiness chart
  set_cell_value_(
    tbSheet.getRange(tbRow + 2, 3),
    Utilities.formatString(
      "=CONCATENATE(\"Guild Readiness \",FIXED(100*AVERAGE(J%s:BI%s)/D%s,1),\"%\")",
      tbRow,
      tbRow,
      tbRow
    )
    , true
    , DocumentApp.HorizontalAlignment.CENTER
  );

  tbRow += 3;

  // list the phases
  for (var i = 0, iLen = phaseList.length; i < iLen; ++i) {
    curTb = tbSheet.getRange(tbRow + i, 2);
    curTb.setValue(phaseList[i][0]);
    set_cell_value_(curTb.offset(0, 1), "=I" + phaseList[i][1], true, DocumentApp.HorizontalAlignment.CENTER);
  }

  // show the legend
  tbRow += phaseList.length + 1;
  curTb = tbSheet.getRange(tbRow, 3);
  set_cell_value_(curTb, "Legend", true, DocumentApp.HorizontalAlignment.CENTER);
  set_cell_value_(curTb.offset(1, 0), "Meets Requirements", false, DocumentApp.HorizontalAlignment.CENTER);
  set_cell_value_(curTb.offset(2, 0), "Missing 1 Gear (<= G8)", false, DocumentApp.HorizontalAlignment.CENTER);
  set_cell_value_(curTb.offset(3, 0), "Missing Levels", false, DocumentApp.HorizontalAlignment.CENTER);
  set_cell_value_(curTb.offset(4, 0), "Missing 1 Gear (> G8)", false, DocumentApp.HorizontalAlignment.CENTER);
  set_cell_value_(curTb.offset(5, 0), "Missing 1 Star", false, DocumentApp.HorizontalAlignment.CENTER);

  // setup player columns
  var table = [];
  curTb = tbSheet.getRange(2, 10);
  for (var i = 0; i < MaxPlayers; ++i) {
    var range = tbSheet.getRange(2, 3, lastHeroRow - 1, 6).getValues();
    var playerData = populate_player_tb_(tbSheet, i, range, heroes);
    if (playerData != null) {
      for (var p = 0, pLen = playerData.length; p < pLen; ++p) {
        if (table[p] == null) {
          // first entry
          table[p] = [playerData[p]];
        } else {
          // append additional entries
          table[p].push(playerData[p]);
        }
      }
    }
  }

  // store the table of player data
  tbSheet.getRange(1, HeroPlayerColOffset + 1, table.length, table[0].length).setValues(table);
}

// Setup the Territory Battle for Hoth
function SetupTBSide() {

  var tag_filter = get_tag_filter_();
  SetupTB("TB", tag_filter);
}
