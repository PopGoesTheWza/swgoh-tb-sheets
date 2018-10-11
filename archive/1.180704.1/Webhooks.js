// ****************************************
// Webhooks Functions
// ****************************************
var RareMax = 15;
var HighMin = 10;
var DiscordWebhookCol = 5;
var WebhookTBStartRow = 3;
var WebhookPhaseHoursRow = 4;
var WebhookTitleRow = 5;
var WebhookWarnRow = 6;
var WebhookRareRow = 7;
var WebhookDepthRow = 8;
var WebhookDescRow = 9;
var WebhookClearRow = 15;

// Get the webhook address
function GetWebhook() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Discord");
  var value = sheet.getRange(1, DiscordWebhookCol).getValue();

  return value;
}

// Get the role to mention
function GetRole() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Discord");
  var value = sheet.getRange(2, DiscordWebhookCol).getValue();

  return value;
}

// Get the time and date when the TB started
function GetTBStartTime() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Discord");
  var value = sheet.getRange(WebhookTBStartRow, DiscordWebhookCol).getValue();

  return value;
}

// Get the number of hours in each phase
function GetPhaseHours() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Discord");
  var value = sheet.getRange(WebhookPhaseHoursRow, DiscordWebhookCol).getValue();

  return value;
}

// Get the template for a webhooks
function GetWebhookTemplate(phase, row, defaultVal) {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Discord");
  var text = sheet.getRange(row, DiscordWebhookCol).getValue();

  if (text.length == 0) {
    text = defaultVal;
  } else {
    text = text.replace("{0}", phase);
  }

  return text;
}

// Get the title for the webhooks
function GetWebhookTitle(phase) {

  var defaultVal = "__**Territory Battle: Phase " + phase + "**__";

  return GetWebhookTemplate(phase, WebhookTitleRow, defaultVal);
}

// Get the intro for the warning webhook
function GetWebhookWarnIntro(phase, mention) {

  var defaultVal = "Here are the __Rare Units__ to watch out for in __Phase " + phase
      + "__. **Check with an officer before donating to Platoons/Squadrons that require them.**";

  return "\n\n" + GetWebhookTemplate(phase, WebhookWarnRow, defaultVal) + " " + mention;
}

// Get the intro for the rare by webhook
function GetWebhookRareIntro(phase, mention) {

  var defaultVal = "Here are the Safe Platoons and the Rare Platoon donations for __Phase " + phase
    + "__. **Do not donate heroes to the other Platoons.**";

  return "\n\n" + GetWebhookTemplate(phase, WebhookRareRow, defaultVal) + " " + mention;
}

// Get the intro for the depth webhook
function GetWebhookDepthIntro(phase, mention) {

  var defaultVal = "Here are the Platoon assignments for __Phase " + phase + "__. **Do not donate heroes to the other Platoons.**";

  return "\n\n" + GetWebhookTemplate(phase, WebhookDepthRow, defaultVal) + " " + mention;
}

// Get the Description for the phase
function GetWebhookDesc(phase) {

  var tag_filter = get_tag_filter_();  // TODO: potentially broken if TB not sync
  var isLight = (tag_filter === "Light Side");
  var sheet = SpreadsheetApp.getActive().getSheetByName("Discord");
  var text = sheet.getRange(WebhookDescRow + phase - 1, DiscordWebhookCol + (isLight ? 0 : 1)).getValue();

  return "\n\n" + text;
}

// See if the platoons should be cleared
function GetWebhookClear() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Discord");
  var value = sheet.getRange(WebhookClearRow, DiscordWebhookCol).getValue();

  return (value == "Yes");
}

// Get the player Discord IDs for mentions
function GetPlayerMentions() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Discord");
  var data = sheet.getRange(2, 1, MaxPlayers, 2).getValues();
  var result = [];

  for (var i = 0, iLen = data.length; i < iLen; ++i) {
    var name = data[i][0];

    // only stores unique names, we can't differentiate with duplicates
    if (name != null && name.length > 0 && result[name] == null) {
      // store the ID if it exists, otherwise store the player's name
      result[name] = (data[i][1] == null || data[i][1].length == 0) ? name : data[i][1];
    }
  }

  return result;
}

// Get a string representing the platoon assignements
function GetPlatoonString(platoon) {

  var result = "";

  // cycle through the heroes
  for (var h = 0; h < MaxPlatoonHeroes; ++h) {
    if (platoon[h][1].length == 0 || platoon[h][1] == "Skip") {
      // impossible platoon
      result = "";
      break;
    }

    if (result.length > 0) {
      result += "\n";
    }

    // remove the gear
    var name = platoon[h][1];
    var endIdx = name.indexOf(" (");
    if (endIdx > 0) {
      name = name.substring(0, endIdx);
    }

    // add the assignement
    result += "**" + platoon[h][0] + "**: " + name;
  }

  return result;
}

// Get the formatted zone name with location descriptor
function GetZoneName(phase, zoneNum, full) {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");
  var zone = sheet.getRange(4 + (zoneNum * PlatoonZoneRowOffset), 1).getValue();

  switch (zoneNum) {
    case 0:
      zone += (full) ? " (Top) Squadrons" : " (Top)";
      break;
    case 2:
      zone += (full) ? " (Bottom) Platoons" : " (Bottom)";
      break;
    case 1:
    default:
      if (phase == 1 && full) {
        zone += " Platoons";
      } else if (phase == 2) {
        zone += (full) ? " (Top) Platoons" : " (Top)";
      } else {
        zone += (full) ? " (Middle) Platoons" : " (Middle)";
      }
      break;
  }

  return zone;
}

// Send a Webhook to Discord
function SendPlatoonDepthWebhook() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");
  var phase = sheet.getRange(2, 1).getValue();

  // get the webhook
  var webhookURL = GetWebhook();
  if (webhookURL.length == 0) {
    // we need a url to proceed
    var ui = SpreadsheetApp.getUi();
    /*var result =*/ ui.alert(
      "Discord Webhook not found (Discord!E1)",
      ui.ButtonSet.OK);
    return;
  }

  // mentions only works if you get the id in Settings - Appearance - Enable Developer Mode, type: \@rolename, copy the value <@$####>
  var mentions = GetRole();

  var title = GetWebhookTitle(phase);
  var descriptionText = title + GetWebhookDepthIntro(phase, mentions);

  // get data from the platoons
  var fields = [];
  for (var z = 0; z < MaxPlatoonZones; ++z) {
    // for each zone
    var platoonRow = 2 + z * PlatoonZoneRowOffset;
    var zone = GetZoneName(phase, z, false);

    if (z == 0 && phase < 3) {
      // skip this zone
      continue;
    }

    // cycle throught the platoons in a zone
    for (var p = 0; p < MaxPlatoons; ++p) {
      var platoonData = sheet.getRange(platoonRow, 4 + (p * 4), MaxPlatoonHeroes, 2).getValues();
      var platoon = GetPlatoonString(platoonData);

      if (platoon.length > 0) {
        var platoonName = zone + ": #" + (p + 1);
        fields[fields.length] =
          {
            "name": platoonName,
            "value": platoon,
            "inline": true,
          };
      }
    }
  }

  var jsonString =
  {
    "content": descriptionText,
    "embeds": [
    {
      "fields": fields
    }]
  }

  var options = url_fetch_make_param_(jsonString);
  url_fetch_execute_(webhookURL, options);
}

// Get an array representing the new platoon assignements
function GetPlatoonDonations(platoon, donations, rules, playerMentions) {

  var result = [];

  // cycle through the heroes
  for (var h = 0; h < MaxPlatoonHeroes; ++h) {
    if (platoon[h][0].length == 0) {
      // no unit needed here
      continue;
    }

    if (platoon[h][1].length == 0 || platoon[h][1] == "Skip") {
      // impossible platoon
      result = null;
      break;
    }

    // see if the hero is already in donations
    var heroDonated = false;
    for (var d = 0, dLen = donations.length; d < dLen; ++d) {
      if (donations[d][0] == platoon[h][0]) {
        heroDonated = true;
        break;
      }
    }
    for (/*var*/ d = 0, dLen = result.length; d < dLen; ++d) {
      if (result[d][0] == platoon[h][0]) {
        heroDonated = true;
        break;
      }
    }

    if (!heroDonated) {
      var criteria = rules[h][0].getCriteriaValues();

      // only add rare donations
      if (criteria[0].length < RareMax) {
        var sorted = criteria[0].sort(lower_case_);
        var text = "";
        for (var c = 0, cLen = sorted.length; c < cLen; ++c) {
          var name = sorted[c];
          var mention = playerMentions[name];
          if (mention != null) {
            name += " (" + mention + ")";
          }
          text += (text.length == 0) ? name : (", " + name);
        }

        // add the recommendations
        result[result.length] = [platoon[h][0], text];
      }
    }
  }

  return result;
}

// Get a list of units that are required a high number of times
function GetHighNeedList(sheetName, unitCount) {

  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var counts = sheet.getRange(2, 1, unitCount, HeroPlayerColOffset).getValues();
  var result = "";
  var idx = HeroPlayerColOffset - 1;

  for (var i = 0, iLen = counts.length; i < iLen; ++i) {
    // unit's count is over the min bar to be too high
    if (counts[i][idx] >= HighMin) {
      if (result.length > 0) {
        result += ", ";
      }
      result += counts[i][0] + " (" + counts[i][idx] + ")";
    }
  }

  return result;
}

// See if a unit is considered in high need
function IsHighNeed(list, unit) {

  for (var i = 0, iLen = list.length; i < iLen; ++i) {
    if (list[i] == unit) {
      return true;
    }
  }

  return false;
}

// Send the message to Discord
function PostMessage(webhookURL, message) {

  var jsonString =
  {
    "content": message.trim(),
  }

  var options = url_fetch_make_param_(jsonString);
  url_fetch_execute_(webhookURL, options);
/*
  try
  {
    UrlFetchApp.fetch(webhookURL, options);
  }
  catch (e)
  {
    // this can be used to debug issues with sending the webhooks.
    // disable "muteHttpExceptions" above to allow the exception to trigger.

    // log the error
    Logger.log(e);

    // split the message, so we can see what it choked on
    var parts = message.split(",");

    // error sending to Discord
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
      "Error sending webhook to Discord. Make sure Platoons are populated and can be filled by the guild.",
      ui.ButtonSet.OK);
  }
*/
}

// Send a Webhook to Discord
function SendPlatoonSimplifiedWebhook(byType) {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");
  var phase = sheet.getRange(2, 1).getValue();

  // get the webhook
  var webhookURL = GetWebhook();
  if (webhookURL.length == 0) {
    // we need a url to proceed
    var ui = SpreadsheetApp.getUi();
    /*var result =*/ ui.alert(
      "Discord Webhook not found (Discord!E1)",
      ui.ButtonSet.OK);
    return;
  }

  // mentions only works if you get the ID on your server, type: \@rolename, copy the value <@#######>
  var playerMentions = GetPlayerMentions();
  var mentions = GetRole();

  var title = GetWebhookTitle(phase);
  var descriptionText = title + GetWebhookRareIntro(phase, mentions);

  var highNeedHeroes = GetHighNeedList("Heroes", get_character_count_());
  var highNeedShips = GetHighNeedList("Ships", get_ship_count_());

  // get data from the platoons
  var fields = "";
  var donations = [];
  var groundStart = -1;
  for (var z = 0; z < MaxPlatoonZones; ++z) {
    // for each zone
    var platoonRow = 2 + z * PlatoonZoneRowOffset;
    var validPlatoons = "";
    var zone = GetZoneName(phase, z, true);

    if (z == 1) {
        groundStart = donations.length;
    }

    if (z == 0 && phase < 3) {
      // skip this and future zones
      continue;
    }

    // cycle throught the platoons in a zone
    for (var p = 0; p < MaxPlatoons; ++p) {
      var platoonData = sheet.getRange(platoonRow, 4 + (p * 4), MaxPlatoonHeroes, 2).getValues();
      var rules = sheet.getRange(platoonRow, 5 + (p * 4), MaxPlatoonHeroes, 1).getDataValidations();
      var platoon = GetPlatoonDonations(platoonData, donations, rules, playerMentions);

      if (platoon != null) {
        var platoonName = "#" + (p + 1);
        validPlatoons += (validPlatoons.length > 0) ? (", " + platoonName): platoonName;

        if (platoon.length > 0) {
          // add the new donations to the list
          for (var i = 0, iLen = platoon.length; i < iLen; ++i) {
            donations[donations.length] = [platoon[i][0], platoon[i][1]];
          }
        }
      }
    }

    // see if all platoons are valid
    if (validPlatoons == "#1, #2, #3, #4, #5, #6") {
      validPlatoons = "All";
    }

    // format the needed platoons
    if (validPlatoons.length > 0) {
      fields += "**" + zone + "**\n" + validPlatoons + "\n\n";
    }
  }

  // format the high needed units
  if (highNeedShips.length > 0) {
    fields += "**High Need Ships**\n" + highNeedShips + "\n\n";
  }
  if (highNeedHeroes.length > 0) {
    fields += "**High Need Heroes**\n" + highNeedHeroes + "\n\n";
  }
  PostMessage(webhookURL, descriptionText + "\n\n" + fields.trim() + "\n'");

  // reformat the output if we need by player istead of by unit
  if (byType == "Player") {
    var heroLabel = "Heroes: ";
    var shipLabel = "Ships: ";
    var playerDonations = [];
    for (var d = 0, dLen = donations.length; d < dLen; ++d) {
      var unit = donations[d][0];
      var names = donations[d][1].split(",");

      for (var n = 0, nLen = names.length; n < nLen; ++n) {
        var name = names[n].trim();

        // see if the name is already listed
        var foundName = false;
        for (/*var*/ p = 0, pLen = playerDonations.length; p < pLen; ++p) {
          if (playerDonations[p][0] == name) {
            if (d >= groundStart && playerDonations[p][1].indexOf(heroLabel) < 0) {
              playerDonations[p][1] += "\n" + heroLabel + unit;
            } else {
              playerDonations[p][1] += ", " + unit;
            }
            foundName = true;
            break;
          }
        }

        if (!foundName) {
          playerDonations[playerDonations.length] = [name, (d >= groundStart) ? heroLabel + unit : shipLabel + unit];
        }
      }
    }

    // sort by player
    donations = playerDonations.sort(first_element_to_lower_case_);
  }

  // format the needed donations
  var maxUrlLen = 1000;
  var maxCount = (byType == "Unit") ? 5 : 10;
  var playerFields = "";
  var count = 0;
  donations
  .filter(function(e) { return e[1].length > 0; })
  .forEach(function(e) {
    var fieldName = (byType == "Unit") ? e[0] + " (Rare)" : e[0];
    playerFields += "**" + fieldName + "**\n" +  e[1] + "\n\n";
    count++;
    // make sure out message isn't getting too long
    if (count >= maxCount || playerFields.length > maxUrlLen) {
      PostMessage(webhookURL, playerFields);
      playerFields = "";
      count = 0;
    }
  });
/*
  for (var d = 0, dLen = donations.length; d < dLen; ++d) {
    if (donations[d][1].length > 0) {
      var fieldName = (byType == "Unit") ? donations[d][0] + " (Rare)" : donations[d][0];
      playerFields += "**" + fieldName + "**\n" +  donations[d][1] + "\n\n";
      count++;

      // make sure out message isn't getting too long
      if (count >= maxCount || playerFields.length > maxUrlLen) {
        PostMessage(webhookURL, playerFields);
        playerFields = "";
        count = 0;
      }
    }
  }
*/
  PostMessage(webhookURL, playerFields);
}

// Send a Webhook to Discord
function SendPlatoonSimplifiedByUnitWebhook() {

  SendPlatoonSimplifiedWebhook("Unit");
}

// Send a Webhook to Discord
function SendPlatoonSimplifiedByPlayerWebhook() {

  SendPlatoonSimplifiedWebhook("Player");
}

// zone: 0, 1 or 2
function GetUniquePlatoonUnits(zone) {

  var platoonRow = 2 + (zone * 18);
  //var rows = MaxPlatoonHeroes;
  var sheet = SpreadsheetApp.getActive().getSheetByName('Platoon');
  var range;
  var units = [];
  for (var platoon = 0; platoon < MaxPlatoons; ++platoon) {
    range = sheet.getRange(platoonRow, 4 + (platoon * 4), MaxPlatoonHeroes, 1);
    units = units.concat(range.getValues());
  }
  // flatten the array and keep only unique values
  units = units
  .map(function(el) {
    return el[0];
  })
  .unique();

  return units;
}

// Get the list of Rare units needed for the phase
function GetRareUnits(sheetName, phase) {
  
  var tag_filter = get_tag_filter_();  // TODO: potentially broken if TB not sync
  var isLight = (tag_filter === "Light Side");
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var count = get_character_count_();
  var data = sheet.getRange(1, 1, count + 1, 8).getValues();
  var idx = phase + 1;
  
  // Drop first line
  data = data.slice(1);
  
  // cycle through each unit
  var units = data
  .filter(function(row) {
    return row[0].length > 0 && row[idx] < RareMax;
  })
  // keep only the unit's name
  .map(function(row) { return row[0]; })
  // sort the list of units  
  .sort();

  var platoonUnits;
  if(sheetName === 'Ships') {
    platoonUnits = GetUniquePlatoonUnits(0);
  } else {
    platoonUnits = GetUniquePlatoonUnits(1);
    if(!isLight || phase > 1) {
      platoonUnits = platoonUnits.concat(GetUniquePlatoonUnits(2));
    }
  }
  // filter out rare units that do not appear in platoons
  units = units
  .filter(function(unit) {
    return platoonUnits.some(function(el) {
      return el === unit;
    });
  });

  // build the output text string
  var text = units.join('\n');
  return text;
}

// Send a message to Discord that lists all units to watch out for in the current phase
function AllRareUnitsWebhook() {

  var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");
  var phase = sheet.getRange(2, 1).getValue();

  // get the webhook
  var webhookURL = GetWebhook();
  if (webhookURL.length == 0) {
    // we need a url to proceed
    var ui = SpreadsheetApp.getUi();
    /*var result =*/ ui.alert(
      "Discord Webhook not found (Discord!E1)",
      ui.ButtonSet.OK);
    return;
  }

  // mentions only works if you get the id in Settings - Appearance - Enable Developer Mode, type: \@rolename, copy the value <@$####>
  var mentions = GetRole();

  var title = GetWebhookTitle(phase);
  var desc = GetWebhookDesc(phase);
  var descriptionText = title + GetWebhookWarnIntro(phase, mentions) + desc;

  // get the ships list
  var fields = [];
  if (phase >= 3) {
    units = GetRareUnits("Ships", phase);
    if (units.length > 0) {
      fields[fields.length] = {
          "name": "Rare Ships",
          "value": units,
          "inline": true,
        };
    }
  }

  // get the hero list
  var units = GetRareUnits("Heroes", phase);
  if (units.length > 0) {
    fields[fields.length] = {
        "name": "Rare Heroes",
        "value": units,
        "inline": true,
      };
  }

  // make sure we're not trying to send empty data
  if (fields.length == 0) {
    // no data to send
    fields[fields.length] = {
        "name": "Rare Heroes",
        "value": "There Are No Rare Units For This Phase.",
        "inline": true,
      };
  }

  var jsonString = {
    "content": descriptionText,
    "embeds": [{
      "fields": fields
    }]
  }

  var options = url_fetch_make_param_(jsonString);
  url_fetch_execute_(webhookURL, options);
}


// ****************************************
// Timer Functions
// ****************************************

// Figure out what phase the TB is in
function SetCurrentPhase() {

  // get the guild's TB start date/time and phase length in hours
  var startTime = GetTBStartTime();
  var phaseHours = GetPhaseHours();
  if (startTime && phaseHours ) {
    var msPerHour = 1000 * 60 * 60;
    var now = new Date();
    var diff = now - startTime;
    var hours = (diff / msPerHour) + 1; // add 1 hour to ensure we are in the next phase
    var phase = Math.ceil(hours / phaseHours);
    var maxPhases = 6;

    // set the phase in Platoons tab
    if (phase <= maxPhases) {
      var sheet = SpreadsheetApp.getActive().getSheetByName("Platoon");
      sheet.getRange(2, 1).setValue(phase);
    }
  }
}

// Callback function to see if we should send the webhook
function SendTimedWebhook() {

  // set the current phase based on time
  SetCurrentPhase();

  // reset the platoons if clear flag was set
  if (GetWebhookClear()) {
    ResetPlatoons();
  }

  // call the webhook
  AllRareUnitsWebhook();

  // register the next timer
  RegisterWebhookTimer();
}

// Try to create a webhook trigger
function RegisterWebhookTimer() {

  // get the guild's TB start date/time and phase length in hours
  var startTime = GetTBStartTime();
  var phaseHours = GetPhaseHours();
  if (startTime && phaseHours) {
    var msPerHour = 1000 * 60 * 60;
    var phaseMs = phaseMs * msPerHour;
    var target = new Date(startTime);
    var now = new Date();
    var maxPhases = 6;

    // remove the trigger
    ScriptApp
    .getProjectTriggers()
    .filter(function(e) { return e.getHandlerFunction() === "SendTimedWebhook"; })
    .forEach(function(e) { ScriptApp.deleteTrigger(e); });

    // see if we can set the trigger later in the phase
    for (var i = 2; i <= maxPhases; ++i) {
      target.setTime(target.getTime() + phaseMs);
      if (target > now) {  // target is in the future
        // found the start of the next phase, so set the timer
         ScriptApp
         .newTrigger("SendTimedWebhook")
         .timeBased()
         .at(target)
         .create();

        break;
      }
    }
  }
}
