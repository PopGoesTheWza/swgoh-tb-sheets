// ****************************************
// Micromanaged Webhook Functions
// ****************************************
var WaitTime = 2000;  // TODO: expose as config variable

// Format the player's label
function player_label_(player, mention) {

  var label;
  if (mention != null) {
    label = Utilities.formatString(
      "Assignments for **%s** (%s)",
      player,
      mention
    );
  } else {
    label = Utilities.formatString(
      "Assignments for **%s**",
      player
    );
  }

  return label;
}

// Format the platoon label with number icons
function player_label_as_icon_(label, type, platoon) {

  var platoonIcon = [":one:", ":two:", ":three:", ":four:", ":five:", ":six:"][platoon];

  return Utilities.formatString(
    "__%s__ · %s %s",
    label,
    type,
    platoonIcon
  );
}

// Check if the unit can be easily confused
function is_group_unit_(unit) {

  return unit.search(
    /X-wing|U-wing|ARC-170|Geonosian|CC-|CT-|Dathcha|Jawa|Hoth Rebel/
  ) > -1;
}

// Convert an array index to a string
function array_index_to_string_(index) {

  var indexNum = Number.parseInt(index) + 1;

  return indexNum.toString();
}

// Format the unit name
function unit_label_(unit, slot) {

  if (is_group_unit_(unit)) {
    return Utilities.formatString(
      "[slot %s] %s",
      array_index_to_string_(slot),
      unit
    );
  }

  return unit;
}

// Send a Webhook to Discord
function SendMicroByPlayerWebhook() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");
  var phase = sheet.getRange(2, 1).getValue();

  // get the webhook
  var webhookURL = GetWebhook();
  if (webhookURL.length == 0) {
    // we need a url to proceed
    var ui = SpreadsheetApp.getUi();
    /*var result =*/ ui.alert("Discord Webhook not found (Discord!E1)", ui.ButtonSet.OK);
    return;
  }

  // get data from the platoons
  var entries = [];
  for (var z = 0; z <= 2; ++z) {
    if (z == 0 && phase < 3) {
      // skip this zone
      continue;
    }

    // for each zone
    var platoonRow = 2 + z * PlatoonZoneRowOffset;
    var label = GetZoneName(phase, z, false);
    var type = (z == 0) ? "squadron" : "platoon";

    // cycle throught the platoons in a zone
    for (var p = 0; p < MaxPlatoons; ++p) {
      var platoonData = sheet.getRange(platoonRow, 4 + (p * 4), MaxPlatoonHeroes, 2).getValues();

      // cycle through the heroes
      platoonData.some(function(e, index) {
        var player = e[1];
        if (player.length === 0 || player === "Skip") {
          return true;
        }

        // remove the gear
        var endIdx = player.indexOf(" (");
        if (endIdx > 0) {
          player = player.substring(0, endIdx);
        }
        var unit = e[0];
        var entry = {
          player: player,
          zone: {
            index: z,
            label: label,
            type: type
          },
          platoon: p,
          slot: index,
          unit: unit
        };
        entries.push(entry);
        return false;
      });
    }
  }

  entries = entries.sort(function(a, b) {
    return a.player.toLowerCase().localeCompare(b.player.toLowerCase());
  });
  
  var playerMentions = GetPlayerMentions();
  while (entries.length > 0) {
    var player = entries[0].player;
    var bucket = entries.filter(function(e) { return e.player === player; });

    entries = entries.slice(bucket.length);
    var embeds = [];
    var currentZone = bucket[0].zone;
    var currentPlatoon = bucket[0].platoon;
    var currentEmbed = {};
    embeds.push(currentEmbed);
    currentEmbed.fields = [];
    var currentField = {};
    currentEmbed.fields.push(currentField);
    currentField.name = player_label_as_icon_(currentZone.label, currentZone.type, currentPlatoon);
    currentField.value = "";
    if (currentZone.label.indexOf("Top") !== -1) {
      currentEmbed.color = 3447003;
    } else if (currentZone.label.indexOf("Bottom") !== -1) {
      currentEmbed.color = 15730230;
    } else {
      currentEmbed.color = 4317713;
    }

    bucket
    .forEach(function(currentValue, index, array) {
      if (currentValue.zone.index != currentZone.index
        || currentValue.platoon != currentPlatoon) {

        currentEmbed = {};
        embeds.push(currentEmbed);
        currentEmbed.fields = [];

        currentZone = currentValue.zone;
        currentPlatoon = currentValue.platoon;
        currentField = {};
        currentEmbed.fields.push(currentField);
        currentField.name = player_label_as_icon_(currentZone.label, currentZone.type, currentPlatoon);
        currentField.value = "";
        if (currentZone.label.indexOf("Top") != -1) {
          currentEmbed.color = 3447003;
        } else if (currentZone.label.indexOf("Bottom") != -1) {
          currentEmbed.color = 15730230;
        } else {
          currentEmbed.color = 4317713;
        }
      }
      if (currentField.value != "") {
        currentField.value += "\n";
      }
      currentField.value += unit_label_(currentValue.unit, currentValue.slot);
    });

    var mention = playerMentions[player];
    var content = player_label_(player, mention);
    var jsonObject = {};
    jsonObject.content = content;
    jsonObject.embeds = embeds;
    var options = url_fetch_make_param_(jsonObject);
    url_fetch_execute_(webhookURL, options);
    Utilities.sleep(WaitTime);
  }
}

// Setup the fetch parameters
// TODO: Make generic for all Discord webhooks
function url_fetch_make_param_(jsonObject) {

  var options = {
    "method": "post",
    "contentType": "application/json",
    // Convert the JavaScript object to a JSON string.
    "payload": JSON.stringify(jsonObject),
    muteHttpExceptions: true
  };

  return options;
}

// Execute the fetch request
// TODO: Make generic for all UrlFetch calls
function url_fetch_execute_(webhookURL, params) {

  // exectute the command
  try {
    UrlFetchApp.fetch(webhookURL, params);
  } catch (e) {
    // log the error
    Logger.log(e);

    // error sending to Discord
    var ui = SpreadsheetApp.getUi();
    /*var result =*/ ui.alert(
      "Error sending webhook to Discord. Make sure Platoons are populated and can be filled by the guild.",
      ui.ButtonSet.OK);
  }
}
