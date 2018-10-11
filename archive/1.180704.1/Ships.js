// ****************************************
// Ship Functions
// ****************************************

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
