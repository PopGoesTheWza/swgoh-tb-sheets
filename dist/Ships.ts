// ****************************************
// Ship Functions
// ****************************************

// Populate the Ships list with Member data
function PopulateShipsList (members) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Ships")
 
  // Build a Ship Index by BaseID
  const baseIDs = sheet.getRange(2, 2, get_ship_count_(), 1).getValues() as string[][]
  const hIdx = []
  baseIDs.forEach( (e, i) => { hIdx[e[0]] = i })
  
  // Build a Member index by Name
  const mList = SpreadsheetApp.getActive().getSheetByName("Roster").getRange(2, 2, get_guild_size_(), 1).getValues() as string[][]
  const pIdx = []
  mList.forEach( (e, i) => { pIdx[e[0]] = i })
  const mHead = []
  mHead[0] = []
  
  // Clear out our old data, if any, including names as order may have changed
  sheet.getRange(1, ShipPlayerColOffset, baseIDs.length, MaxPlayers).clearContent()
  
  // This will hold all our data
  // Initialize our data
  const data = baseIDs.map( e => Array(mList.length).fill(null) )

  members.forEach( m => {
    mHead[0].push(m.name)
    const units = m.units
    baseIDs.forEach( (e, i) => {
      // const u = units[e[0]]
      const u = units.filter( (eu, iu) => iu === e[0] )
      data[hIdx[e[0]]][pIdx[m.name]] = u && `${u.rarity}*L${u.level}P${u.power}`
    })
  })
  // for ( var mKey in members) {
  //   if (mKey == "unique") { continue; }
  //   var m = members[mKey];
  //   mHead[0][pIdx[mKey]] = m['name'];
  //   var units = m['units'];
  //   for (var r = 0; r < baseIDs.length; r++) {
  //     var uKey = baseIDs[r];
  //     var u = units[uKey];
  //     if (u == null ) { continue; } // Means player has not unlocked unit
  //     data[hIdx[uKey]][pIdx[m['name']]] = u['rarity']+"*L"+u['level']+"P"+u['power'];
  //   }
  // }
  //Write our data
  sheet.getRange(1, ShipPlayerColOffset, 1, mList.length).setValues(mHead);
  sheet.getRange(2, ShipPlayerColOffset, baseIDs.length, mList.length).setValues(data);
}

// Initialize the list of ships
function UpdateShipsList(ships) {
                     
  // update the sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName("Ships");
  
  // clear the old content
  sheet.getRange(1, 1, 300, ShipPlayerColOffset-1).clearContent();
  
  var result = ships.map( function (e,i) {
    var hMap = [e[0], e[1], e[2]];

    // insert the star count formulas
    var row = i + 2;
    var rangeText =
    Utilities.formatString(
      "$J%s:$BI%s",
      row,
      row
    );

    [2, 3, 4, 5, 6, 7]
    .forEach(function(stars) {
      var formula =
      Utilities.formatString(
        "=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(%s,\"([%s-7]+)\\*\")))))",
        rangeText,
        stars
      );

      hMap.push(formula);
    });

    // insert the needed count
    var formula =
    Utilities.formatString(
      "=COUNTIF(Platoon!$2:$16,A%s)",
      row
    );

    hMap.push(formula);

    return hMap;
  });
  var header =[];
  header[0] = ["Ship", "Base ID", "Tags", 2, 3,4,5,6,7,"=CONCAT(\"# Needed P\",Platoon!A2)"];
  sheet.getRange(1, 1, 1,header[0].length).setValues(header);
  sheet.getRange(2, 1, result.length, ShipPlayerColOffset-1).setValues(result);
  
  return;
}

/*
// Populate the list of ships
function PopulateShipList() {

  // get the web page source
  var link = "https://swgoh.gg/ships/";
  var response;

  try {
    response = UrlFetchApp.fetch(link);
  } catch (e) {
    return "";
  }

  // divide the source into lines that can be parsed
  var text = response.getContentText();
  var json = text
  .match(/<li\s+class="media\s+list-group-item\s+p-0\s+unit\s+character"[^]+?<\/li>/g)
  .map(function(e) {

    var tags = e
    .match(/<small[^>]*>([^]*?)<\/small/)[1]
    .split("·")
    .map(function(t,i,a) {
      var tag = t.match(/\s*([^>]+?)\s*$/);
      return (tag) ? tag[1] : null;
    });

    var side = tags.shift();
    var role = tags.shift();

    var o = {
      "name": e.match(/<h5>([^<]*)/)[1].replace(/&quot;/g, "\"").replace(/&#39;/g, "'"),
      "side": side,
      "role": role,
      "tags": tags,
    };

    return o;
  });

  var units = json
  .map(function(e, i) {
    var tags =
    Utilities.formatString(
      "%s %s",
      e.side,
      e.role
    );

    if (e.tags) {
      tags =
      Utilities.formatString(
        "%s %s",
        tags,
        e.tags.join(" ")
      );
    }

    var result = [
      e.name,
      tags.toLowerCase()
    ];

    // insert the star count formulas
    var row = i + 2;
    var rangeText =
    Utilities.formatString(
      "$J%s:$BI%s",
      row,
      row
    );

    // Rarity (stars) count formulas
    [2, 3, 4, 5, 6, 7]
    .forEach(function(stars) {
      var formula =
      Utilities.formatString(
        "=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(%s,\"([%s-7]+)\\*\")))))",
        rangeText,
        stars
      );

      result.push(formula);
    });

    // insdert the needed count
    var formula =
    Utilities.formatString(
      "=COUNTIF(Platoon!$2:$16,A%s)",
      row
    );

    result.push(formula);
    return result;
  });

  // update the sheet
  sheet = SpreadsheetApp.getActive().getSheetByName("Ships");
  sheet.getRange(2, 1, units.length, 9).setValues(units);

  return units;
}
*/
