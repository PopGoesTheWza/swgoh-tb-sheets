// ****************************************
// Micromanaged Webhook Functions
// ****************************************

/** add discord mention to the member label */
function memberLabel_(member: string, mention: string) {

  const value: string = (mention)
    ? `Assignments for **${member}** (${mention})`
    : `Assignments for **${member}**`;

  return value;
}

/** output platoon numner as discord icon */
function platoonAsIcon_(label: string, type: string, platoon: number) {

  const platoonIcon = [':one:', ':two:', ':three:', ':four:', ':five:', ':six:'][platoon];

  return `__${label}__ · ${type} ${platoonIcon}`;
}

/** check if the unit can be difficult to identify */
function isUnitHardToRead_(unit: string): boolean {

  return unit.search(
      /X-wing|U-wing|ARC-170|Geonosian|CC-|CT-|Dathcha|Jawa|Hoth Rebel/,
    ) > -1;
}

/** convert an array index to uhman friendly string (0 => '1') */
function arrayIndexToString_(index: string): string {
  return (parseInt(index, 10) + 1).toString();
}

/** format the unit name */
function unitLabel_(unit: string, slot: string, force: boolean = undefined): string {

  if (force || isUnitHardToRead_(unit)) {
    return `[slot ${arrayIndexToString_(slot)}] ${unit}`;
  }

  return unit;
}

type DiscordEmbeddedField = {
  name?: string;
  value?: string;
};

type DiscordEmbed = {
  color?: number;
  fields?: DiscordEmbeddedField[];
};

type DiscordPayload = {
  content?: string;
  embeds?: DiscordEmbed[];
};

/** Send a Webhook to Discord */
function sendMicroByMemberWebhook(): void {

  const displaySetting = config.discord.displaySlots();
  const displaySlot = displaySetting !== DISPLAYSLOT.NEVER;
  const forceDisplay = displaySetting === DISPLAYSLOT.ALWAYS;
  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = config.currentPhase();

  // get the webhook
  const webhookURL = config.discord.webhookUrl();
  if (webhookURL.length === 0) {
    // we need a url to proceed
    const UI = SpreadsheetApp.getUi();
    UI.alert(
      'Configuration Error',
      'Discord webhook not found (Discord!E1)',
      UI.ButtonSet.OK,
    );

    return;
  }

  // get data from the platoons
  let entries = [];
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    if (z === 0 && phase < 3) {
      // skip this zone
      continue;
    }

    // for each zone
    const platoonRow = 2 + z * PLATOON_ZONE_ROW_OFFSET;
    const label = discord.getZoneName(phase, z as TerritoryBattles.territoryIdx, false);
    const type = z === 0 ? 'squadron' : 'platoon';

    // cycle throught the platoons in a zone
    for (let p = 0; p < MAX_PLATOONS; p += 1) {
      const platoonData = sheet
        .getRange(platoonRow, 4 + p * 4, MAX_PLATOON_UNITS, 2)
        .getValues() as string[][];

      // cycle through the heroes
      platoonData.some((e, index) => {
        let member = e[1];
        if (member.length === 0 || member === 'Skip') {
          return true;
        }

        // remove the gear
        const endIdx = member.indexOf(' (');
        if (endIdx > 0) {
          member = member.substring(0, endIdx);
        }
        const unit = e[0];
        const entry = {
          member,
          unit,
          zone: {
            label,
            type,
            index: z,
          },
          platoon: p,
          slot: index,
        };
        entries.push(entry);
        return false;
      });
    }
  }

  entries.sort((a, b) => {
    return utils.caseInsensitive(a.member, b.member);
  });

  const memberMentions = discord.getMemberMentions();
  while (entries.length > 0) {
    const member = entries[0].member;
    const bucket = entries.filter(e => e.member === member);

    entries = entries.slice(bucket.length);
    const embeds: DiscordEmbed[] = [];
    let currentZone = bucket[0].zone;
    let currentPlatoon = bucket[0].platoon;
    let currentEmbed: DiscordEmbed = {};
    embeds.push(currentEmbed);
    currentEmbed.fields = [];
    let currentField: DiscordEmbeddedField = {};
    currentEmbed.fields.push(currentField);
    currentField.name = platoonAsIcon_(
      currentZone.label,
      currentZone.type,
      currentPlatoon,
    );
    currentField.value = '';
    if (currentZone.label.indexOf('Top') !== -1) {
      currentEmbed.color = 3447003;
    } else if (currentZone.label.indexOf('Bottom') !== -1) {
      currentEmbed.color = 15730230;
    } else {
      currentEmbed.color = 4317713;
    }

    for (const currentValue of bucket) {
      if (
        currentValue.zone.index !== currentZone.index ||
        currentValue.platoon !== currentPlatoon
      ) {
        currentEmbed = {};
        embeds.push(currentEmbed);
        currentEmbed.fields = [];

        currentZone = currentValue.zone;
        currentPlatoon = currentValue.platoon;
        currentField = {};
        currentEmbed.fields.push(currentField);
        currentField.name = platoonAsIcon_(
          currentZone.label,
          currentZone.type,
          currentPlatoon,
        );
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
    const options: URL_Fetch.URLFetchRequestOptions = urlFetchMakeParam_(jsonObject);
    urlFetchExecute_(webhookURL, options);
    Utilities.sleep(WAIT_TIME);
  }
}

/** Setup the fetch parameters */
// TODO: Make generic for all Discord webhooks
function urlFetchMakeParam_(jsonObject: DiscordPayload): URL_Fetch.URLFetchRequestOptions {

  const options: URL_Fetch.URLFetchRequestOptions = {
    method: 'post',
    contentType: 'application/json',
    // Convert the JavaScript object to a JSON string.
    payload: JSON.stringify(jsonObject),
    muteHttpExceptions: true,
  };

  return options;
}

/** Execute the fetch request */
// TODO: Make generic for all UrlFetch calls
function urlFetchExecute_(webhookURL, params) {

  // exectute the command
  try {
    UrlFetchApp.fetch(webhookURL, params);
  } catch (e) {
    // log the error
    Logger.log(e);

    // error sending to Discord
    const UI = SpreadsheetApp.getUi();
    UI.alert(
      'Connection Error',
      'Error sending webhook to Discord.',
      UI.ButtonSet.OK,
    );
  }
}
