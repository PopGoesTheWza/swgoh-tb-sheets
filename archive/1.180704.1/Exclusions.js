// Get the list of exclusions
function get_exclusions_() {

  var excludeSheet = SpreadsheetApp.getActive().getSheetByName("Exclusions");
  var excludeData = excludeSheet.getDataRange().getValues()
  .filter(function(e) {
    return e[0].length > 0;
  })
  .map(function(e) {
    return e.slice(0, MaxPlayers);
  });

  var excludeHeroes = [];

  // First row is player names
  var players = excludeData.shift();
  // drop first column
  players.shift();

  // For each unit rows
  excludeData.forEach(function(e) {
    // first column is unit names
    var name = e.shift();
    excludeHeroes[name] = [];
    
    // For each exclusion cell
    e.forEach(function(x, c) {
      var player = players[c];
      var cell = (x) ? x.trim() : "";
      // exclude if cell is not empty?
      excludeHeroes[name][player] = Boolean(cell);
    })
  });

  return excludeHeroes;
}

// Process all the excluded units
function ProcessExclusions(data, excludeData) {

/*
  // First row is player names
  var players = data.slice().shift();
  // drop first column
  players.shift();

  data.forEach(function(e, i, a) {
  });
*/

  for (var r = 1, rLen = data.length; r < rLen; ++r) {
    for (var c = HeroPlayerColOffset, cLen = data[r].length; c < cLen; ++c) {
      try {
        var heroName = data[r][0];
        var playerName = data[0][c];
        if (excludeData[heroName] != null && excludeData[heroName][playerName]) {
          // clear the unit's data
          data[r][c] = "";
        }
      } catch (e) {
      }
    }
  }

  return data;
}
