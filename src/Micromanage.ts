// ****************************************
// Micromanaged Webhook Functions
// ****************************************

/** add discord mention to the member label */
function memberLabel_(member: string, mention: string) {
  const value: string = mention ? `Assignments for **${member}** (${mention})` : `Assignments for **${member}**`;

  return value;
}

/** output platoon numner as discord icon */
function platoonAsKeycapDigit_(label: string, type: string, platoon: number) {
  const keycapDigit = utils.EMOJI_KEYCAP_DIGITS[platoon + 1];

  return `__${label}__ · ${type} ${keycapDigit}`;
}

/** check if the unit can be difficult to identify */
function isUnitHardToRead_(unit: string): boolean {
  return unit.search(/X-wing|U-wing|ARC-170|Geonosian|CC-|CT-|Dathcha|Jawa|Hoth Rebel/) > -1;
}

/** convert an array index to uhman friendly string (0 => '1') */
function arrayIndexToString_(index: number | string): string {
  return ((typeof index === 'string' ? parseInt(index, 10) : index) + 1).toString();
}

/** format the unit name */
function unitLabel_(unit: string, slot: number | string, force?: boolean): string {
  if (force || isUnitHardToRead_(unit)) {
    return `[slot ${arrayIndexToString_(slot)}] ${unit}`;
  }

  return unit;
}

interface DiscordEmbeddedField {
  name?: string;
  value?: string;
}

interface DiscordEmbed {
  color?: number;
  fields?: DiscordEmbeddedField[];
}

interface DiscordPayload {
  content?: string;
  embeds?: DiscordEmbed[];
}

interface PlatoonAssignment {
  member: string;
  unit: string;
  zone: {
    label: string;
    type: string;
    index: number;
  };
  platoon: number;
  slot: number;
}

/** Send a Webhook to Discord */
function sendMicroByMemberWebhook(): void {
  const WAIT_TIME = 2000;
  const displaySetting = config.discord.displaySlots();
  const displaySlot = displaySetting !== DISPLAYSLOT.NEVER;
  const forceDisplay = displaySetting === DISPLAYSLOT.ALWAYS;

  const SKIPPED_PLATOON_LABEL = TerritoryBattles.SKIPPED_PLATOON_LABEL;
  const getZoneName = TerritoryBattles.getZoneName;
  const getPlatoonData = TerritoryBattles.getPlatoonData;
  const sheet = SPREADSHEET.getSheetByName(SHEET.PLATOON);
  const event = config.currentEvent();
  const phase = config.currentPhase();

  // get the webhook
  const webhookURL = config.discord.webhookUrl();
  if (webhookURL.length === 0) {
    // we need a url to proceed
    const UI = SpreadsheetApp.getUi();
    UI.alert('Configuration Error', 'Discord webhook not found (on Discord sheet)', UI.ButtonSet.OK);

    return;
  }

  // get data from the platoons
  let entries: PlatoonAssignment[] = [];
  for (let zoneNum: TerritoryBattles.territoryIdx = 0; zoneNum < MAX_PLATOON_ZONES; zoneNum += 1) {
    if (!discord.isTerritory(zoneNum, phase, event)) {
      // skip this zone
      continue;
    }

    // for each zone
    const label = getZoneName(zoneNum, false);
    const type = zoneNum === 0 ? 'squadron' : 'platoon';

    // cycle throught the platoons in a zone
    for (let platoonNum = 0; platoonNum < MAX_PLATOONS; platoonNum += 1) {
      const platoonData = getPlatoonData(zoneNum, platoonNum, sheet);

      // cycle through the heroes
      for (let index = 0; index < platoonData.length; index += 1) {
        const e = platoonData[index];
        let member = e[1];
        if (member.length === 0 || member === SKIPPED_PLATOON_LABEL) {
          break;
        }

        // remove the gear
        const endIdx = member.indexOf(' (');
        if (endIdx > 0) {
          member = member.substring(0, endIdx);
        }
        const unit = e[0];
        const entry = {
          member,
          platoon: platoonNum,
          slot: index,
          unit,
          zone: {
            index: zoneNum,
            label,
            type,
          },
        };
        entries.push(entry);
      }
    }
  }

  entries.sort((a, b) => {
    return utils.caseInsensitive(a.member, b.member);
  });

  const memberMentions = discord.getMemberMentions();
  while (entries.length > 0) {
    const member = entries[0].member;
    const bucket = entries.filter((e) => e.member === member);

    entries = entries.slice(bucket.length);
    const embeds: DiscordEmbed[] = [];
    let currentZone = bucket[0].zone;
    let currentPlatoon = bucket[0].platoon;
    let currentEmbed: DiscordEmbed = {};
    embeds.push(currentEmbed);
    currentEmbed.fields = [];
    let currentField: DiscordEmbeddedField = {};
    currentEmbed.fields.push(currentField);
    currentField.name = platoonAsKeycapDigit_(currentZone.label, currentZone.type, currentPlatoon);
    currentField.value = '';
    if (currentZone.label.indexOf('Top') !== -1) {
      currentEmbed.color = 3447003;
    } else if (currentZone.label.indexOf('Bottom') !== -1) {
      currentEmbed.color = 15730230;
    } else {
      currentEmbed.color = 4317713;
    }

    for (const currentValue of bucket) {
      if (currentValue.zone.index !== currentZone.index || currentValue.platoon !== currentPlatoon) {
        currentEmbed = {};
        embeds.push(currentEmbed);
        currentEmbed.fields = [];

        currentZone = currentValue.zone;
        currentPlatoon = currentValue.platoon;
        currentField = {};
        currentEmbed.fields.push(currentField);
        currentField.name = platoonAsKeycapDigit_(currentZone.label, currentZone.type, currentPlatoon);
        currentField.value = '';
        if (currentZone.label.indexOf('Top') !== -1) {
          currentEmbed.color = 3447003;
        } else if (currentZone.label.indexOf('Bottom') !== -1) {
          currentEmbed.color = 15730230;
        } else {
          currentEmbed.color = 4317713;
        }
      }
      if (currentField.value !== '') {
        currentField.value += '\n';
      }
      currentField.value += displaySlot
        ? unitLabel_(currentValue.unit, currentValue.slot, forceDisplay)
        : currentValue.unit;
    }

    const mention = memberMentions[member];
    const content = memberLabel_(member, mention);
    const jsonObject: DiscordPayload = {};
    jsonObject.content = content;
    jsonObject.embeds = embeds;
    const options = urlFetchMakeParam_(jsonObject);
    urlFetchExecute_(webhookURL, options);
    Utilities.sleep(WAIT_TIME);
  }
}

/** Setup the fetch parameters */
// TODO: Make generic for all Discord webhooks
function urlFetchMakeParam_(jsonObject: DiscordPayload): URL_Fetch.URLFetchRequestOptions {
  const options: URL_Fetch.URLFetchRequestOptions = {
    contentType: 'application/json',
    method: 'post',
    muteHttpExceptions: true,
    // Convert the JavaScript object to a JSON string.
    payload: JSON.stringify(jsonObject),
  };

  return options;
}

/** Execute the fetch request */
// TODO: Make generic for all UrlFetch calls
function urlFetchExecute_(webhookURL: string, params: object) {
  // exectute the command
  try {
    UrlFetchApp.fetch(webhookURL, params);
  } catch (e) {
    // log the error
    Logger.log(e);

    // error sending to Discord
    const UI = SpreadsheetApp.getUi();
    UI.alert('Connection Error', 'Error sending webhook to Discord.', UI.ButtonSet.OK);
  }
}
