// ****************************************
// Micromanaged Webhook Functions
// ****************************************

/** add discord mention to the player label */
function playerLabel_(player: string, mention: string) {

  const value: string = (mention)
    ? `Assignments for **${player}** (${mention})`
    : `Assignments for **${player}**`;

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

/** Send a Webhook to Discord */
function sendMicroByPlayerWebhook(): void {

  const displaySetting = config.discord.displaySlots();
  const displaySlot = displaySetting !== DISPLAYSLOT.NEVER;
  const forceDisplay = displaySetting === DISPLAYSLOT.ALWAYS;
  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = sheet.getRange(2, 1).getValue() as number;

  // get the webhook
  const webhookURL = config.discord.webhookUrl();
  if (webhookURL.length === 0) {
    // we need a url to proceed
    UI.alert('Discord Webhook not found (Discord!E1)', UI.ButtonSet.OK);
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
    const label = discord.getZoneName(phase, z, false);
    const type = z === 0 ? 'squadron' : 'platoon';

    // cycle throught the platoons in a zone
    for (let p = 0; p < MAX_PLATOONS; p += 1) {
      const platoonData = sheet
        .getRange(platoonRow, 4 + p * 4, MAX_PLATOON_UNITS, 2)
        .getValues() as string[][];

      // cycle through the heroes
      platoonData.some((e, index) => {
        let player = e[1];
        if (player.length === 0 || player === 'Skip') {
          return true;
        }

        // remove the gear
        const endIdx = player.indexOf(' (');
        if (endIdx > 0) {
          player = player.substring(0, endIdx);
        }
        const unit = e[0];
        const entry = {
          player,
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

  entries = entries.sort((a, b) => {
    return a.player.toLowerCase().localeCompare(b.player.toLowerCase());
  });

  const playerMentions = discord.getPlayerMentions();
  while (entries.length > 0) {
    const player = entries[0].player;
    const bucket = entries.filter(e => e.player === player);

    entries = entries.slice(bucket.length);
    const embeds = [];
    let currentZone = bucket[0].zone;
    let currentPlatoon = bucket[0].platoon;
    let currentEmbed: any = {};
    embeds.push(currentEmbed);
    currentEmbed.fields = [];
    let currentField: any = {};
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

    const mention = playerMentions[player];
    const content = playerLabel_(player, mention);
    const jsonObject: any = {};
    jsonObject.content = content;
    jsonObject.embeds = embeds;
    const options: URLFetchRequestOptions = urlFetchMakeParam_(jsonObject);
    urlFetchExecute_(webhookURL, options);
    Utilities.sleep(WAIT_TIME);
  }
}

/** Setup the fetch parameters */
// TODO: Make generic for all Discord webhooks
function urlFetchMakeParam_(jsonObject: object): URLFetchRequestOptions {

  const options: URLFetchRequestOptions = {
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
    UI.alert(
      `Error sending webhook to Discord.
Make sure Platoons are populated and can be filled by the guild.`,
      UI.ButtonSet.OK,
    );
  }
}
