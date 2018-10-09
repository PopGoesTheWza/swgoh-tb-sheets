// ****************************************
// Webhooks Functions
// ****************************************
const RARE_MAX: number = 15;
const HIGH_MIN: number = 10;
const DISCORD_WEBHOOK_COL: number = 5;
const WEBHOOK_TB_START_ROW: number = 3;
const WEBHOOK_PHASE_HOURS_ROW: number = 4;
const WEBHOOK_TITLE_ROW: number = 5;
const WEBHOOK_WARN_ROW: number = 6;
const WEBHOOK_RARE_ROW: number = 7;
const WEBHOOK_DEPTH_ROW: number = 8;
const WEBHOOK_DESC_ROW: number = 9;
const WEBHOOK_CLEAR_ROW: number = 15;

// Get the webhook address
function getWebhook(): string {
  const value: string = SPREADSHEET
    .getSheetByName(SHEETS.DISCORD)
    .getRange(1, DISCORD_WEBHOOK_COL)
    .getValue() as string;

  return value;
}

// Get the role to mention
function getRole(): string {
  const value: string = SPREADSHEET
    .getSheetByName(SHEETS.DISCORD)
    .getRange(2, DISCORD_WEBHOOK_COL)
    .getValue() as string;

  return value;
}

// Get the time and date when the TB started
function getTBStartTime(): Date {
  const value: Date = SPREADSHEET
    .getSheetByName(SHEETS.DISCORD)
    .getRange(WEBHOOK_TB_START_ROW, DISCORD_WEBHOOK_COL)
    .getValue() as Date;

  return value;
}

// Get the number of hours in each phase
function getPhaseHours(): number {
  const value: number = SPREADSHEET
    .getSheetByName(SHEETS.DISCORD)
    .getRange(WEBHOOK_PHASE_HOURS_ROW, DISCORD_WEBHOOK_COL)
    .getValue() as number;

  return value;
}

// Get the template for a webhooks
function getWebhookTemplate(phase: number, row: number, defaultVal: string): string {
  const text: string = SPREADSHEET
    .getSheetByName(SHEETS.DISCORD)
    .getRange(row, DISCORD_WEBHOOK_COL)
    .getValue() as string;

  return text.length > 0 ? text.replace('{0}', String(phase)) : defaultVal;
}

// Get the title for the webhooks
function getWebhookTitle(phase: number): string {
  const defaultVal: string = `__**Territory Battle: Phase ${phase}**__`;

  return `${getWebhookTemplate(phase, WEBHOOK_TITLE_ROW, defaultVal)}`;
}

// Get the intro for the warning webhook
function getWebhookWarnIntro(phase: number, mention: string): string {
  const defaultIntro: string = 'Here are the __Rare Units__ to watch out for in ';
  const defaultFootnote: string =
  '. **Check with an officer before donating to Platoons/Squadrons that require them.**';
  const defaultVal: string = `${defaultIntro}__Phase ${phase}__${defaultFootnote}`;

  return `\n\n${getWebhookTemplate(
    phase,
    WEBHOOK_WARN_ROW,
    defaultVal,
  )} ${mention}`;
}

// Get the intro for the rare by webhook
function getWebhookRareIntro(phase: number, mention: string): string {
  const defaultIntro: string = 'Here are the Safe Platoons and the Rare Platoon donations for ';
  const defaultFootnote: string = '. **Do not donate heroes to the other Platoons.**';
  const defaultVal: string = `${defaultIntro}__Phase ${phase}__${defaultFootnote}`;

  return `\n\n${getWebhookTemplate(
    phase,
    WEBHOOK_RARE_ROW,
    defaultVal,
  )} ${mention}`;
}

// Get the intro for the depth webhook
function getWebhookDepthIntro(phase: number, mention: string): string {
  const defaultIntro: string = 'Here are the Platoon assignments for ';
  const defaultFootnote: string = '. **Do not donate heroes to the other Platoons.**';
  const defaultVal: string = `${defaultIntro}__Phase ${phase}__${defaultFootnote}`;

  return `\n\n${getWebhookTemplate(
    phase,
    WEBHOOK_DEPTH_ROW,
    defaultVal,
  )} ${mention}`;
}

// Get the Description for the phase
function getWebhookDesc(phase: number): string {
  const tagFilter: string = getTagFilter_(); // TODO: potentially broken if TB not sync
  const columnOffset = isLight_(tagFilter) ? 0 : 1;
  const text: string = SPREADSHEET
    .getSheetByName(SHEETS.DISCORD)
    .getRange(WEBHOOK_DESC_ROW + phase - 1, DISCORD_WEBHOOK_COL + columnOffset)
    .getValue() as string;

  return `\n\n${text}`;
}

// See if the platoons should be cleared
function getWebhookClear(): boolean {
  const value: string = SPREADSHEET
    .getSheetByName(SHEETS.DISCORD)
    .getRange(WEBHOOK_CLEAR_ROW, DISCORD_WEBHOOK_COL)
    .getValue() as string;

  return value === 'Yes';
}

// Get the player Discord IDs for mentions
function getPlayerMentions(): string[] {
  const data: string[][] = SPREADSHEET
    .getSheetByName(SHEETS.DISCORD)
    .getRange(2, 1, MAX_PLAYERS, 2)
    .getValues() as string[][];
  const result: string[] = [];

  for (const e of data) {
    const name: string = e[0];
    // only stores unique names, we can't differentiate with duplicates
    if (name && name.length > 0 && !result[name]) {
      // store the ID if it exists, otherwise store the player's name
      result[name] = e[1] && e[1].length > 0 ? e[1] : name;
    }
  }

  return result;
}

// Get a string representing the platoon assignements
function getPlatoonString(platoon: string[][]): string {
  let result: string = '';

  // cycle through the heroes
  for (let h: number = 0; h < MAX_PLATOON_HEROES; h += 1) {
    if (platoon[h][1].length === 0 || platoon[h][1] === 'Skip') {
      // impossible platoon
      return '';  // TODO return undefined
    }

    // remove the gear
    let name: string = platoon[h][1];
    const endIdx: number = name.indexOf(' (');
    if (endIdx > 0) {
      name = name.substring(0, endIdx);
    }

    // add the assignement
    result =
      result.length > 0
        ? `${result}\n**${platoon[h][0]}**: ${name}`
        : `**${platoon[h][0]}**: ${name}`;
  }

  return result;
}

// Get the formatted zone name with location descriptor
function getZoneName(phase: number, zoneNum: number, full: boolean): string {
  let zone: string = SPREADSHEET
    .getSheetByName(SHEETS.PLATOONS)
    .getRange((zoneNum * PLATOON_ZONE_ROW_OFFSET) + 4, 1)
    .getValue() as string;
  let loc: string;

  switch (zoneNum) {
    case 0:
      // zone += full ? " (Top) Squadrons" : " (Top)"
      loc = '(Top)';
      break;
    case 2:
      // zone += full ? " (Bottom) Platoons" : " (Bottom)"
      loc = '(Bottom)';
      break;
    case 1:
    default:
      // if (phase === 1 && full) {
      //   zone += " Platoons"
      // } else if (phase === 2) {
      //   zone += full ? " (Top) Platoons" : " (Top)"
      // } else {
      //   zone += full ? " (Middle) Platoons" : " (Middle)"
      // }
      loc = phase === 2 ? '(Top)' : '(Middle)';
  }
  zone =
    full && phase !== 1
      ? `${zone} ${loc} ${zoneNum === 0 ? 'Squadrons' : 'Platoons'}`
      : `${loc} ${zone}`;

  return zone;
}

interface IDiscordMessageEmbedFields {
  name: string;
  value: string;
  inline?: boolean;
}

// Send a Webhook to Discord
function sendPlatoonDepthWebhook(): void {
  const sheet: GoogleAppsScript.Spreadsheet.Sheet
  = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase: number = sheet.getRange(2, 1).getValue() as number;

  // get the webhook
  const webhookURL: string = getWebhook();
  if (webhookURL.length === 0) {
    // we need a url to proceed
    UI.alert('Discord Webhook not found (Discord!E1)', UI.ButtonSet.OK);

    return;
  }

  // mentions only works if you get the id
  // in Settings - Appearance - Enable Developer Mode, type: \@rolename, copy the value <@$####>
  const mentions: string = getRole();

  const title: string = getWebhookTitle(phase);
  const descriptionText: string = `${title}${getWebhookDepthIntro(phase, mentions)}`;

  // get data from the platoons
  const fields: IDiscordMessageEmbedFields[] = [];
  for (let z: number = 0; z < MAX_PLATOON_ZONES; z += 1) {
    if (z === 0 && phase < 3) {
      continue; // skip this zone
    }

    // for each zone
    const platoonRow: number = (z * PLATOON_ZONE_ROW_OFFSET) + 2;
    const zone: string = getZoneName(phase, z, false);

    // cycle throught the platoons in a zone
    for (let p: number = 0; p < MAX_PLATOONS; p += 1) {
      const platoonData: string[][] = sheet
        .getRange(platoonRow, (p * 4) + 4, MAX_PLATOON_HEROES, 2)
        .getValues() as string[][];
      const platoon: string = getPlatoonString(platoonData);

      if (platoon.length > 0) {
        fields.push({
          name: `${zone}: #${p + 1}`,
          value: platoon,
          inline: true,
        });
      }
    }
  }

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = urlFetchMakeParam_({
    content: descriptionText,
    embeds: [{ fields }],
  });
  urlFetchExecute_(webhookURL, options);
}

// Get an array representing the new platoon assignements
function getPlatoonDonations(
  platoon: string[][],
  donations: string[][],
  rules: GoogleAppsScript.Spreadsheet.DataValidation[][],
  playerMentions: string[],
): string[][] {
  const result: string[][] = [];

  // cycle through the heroes
  for (let h: number = 0; h < MAX_PLATOON_HEROES; h += 1) {
    if (platoon[h][0].length === 0) {
      continue; // no unit needed here
    }

    if (platoon[h][1].length === 0 || platoon[h][1] === 'Skip') {
      return undefined; // impossible platoon
    }

    // see if the hero is already in donations
    const heroDonated: boolean =
      donations.some((e: string[]) => e[0] === platoon[h][0]) ||
      result.some((e: string[]) => e[0] === platoon[h][0]);

    if (!heroDonated) {
      const criteria: string[][] = rules[h][0].getCriteriaValues() as string[][];

      // only add rare donations
      if (criteria[0].length < RARE_MAX) {
        const sorted: string[] = criteria[0].sort(lowerCase_);
        let text: string = '';
        for (let name of sorted) {
          const mention: string = playerMentions[name];
          if (mention) {
            name = `${name} (${mention})`;
          }
          text += text.length === 0 ? `${name}` : `, ${name}`;
        }
        // add the recommendations
        result.push([platoon[h][0], text]);
      }
    }
  }

  return result;
}

// Get a list of units that are required a high number of times
function getHighNeedList(sheetName: string, unitCount: number): string {
  const counts: number[][] = SPREADSHEET
    .getSheetByName(sheetName)
    .getRange(2, 1, unitCount, HERO_PLAYER_COL_OFFSET)
    .getValues() as number[][];
  let result: string = '';
  const idx: number = HERO_PLAYER_COL_OFFSET - 1;

  for (const e of counts) {
    // unit's count is over the min bar to be too high
    if (e[idx] >= HIGH_MIN) {
      result +=
        result.length > 0 ? `, ${e[0]} (${e[idx]})` : `${e[0]} (${e[idx]})`;
    }
  }

  return result;
}

// See if a unit is considered in high need
function isHighNeed(list: string[], unit: string): boolean {
  return list.some((u: string) => u === unit);
}

// Send the message to Discord
function postMessage(webhookURL: string, message: string): void {
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
  = urlFetchMakeParam_({ content: message.trim() });
  urlFetchExecute_(webhookURL, options);
  // try
  // {
  //   UrlFetchApp.fetch(webhookURL, options)
  // }
  // catch (e)
  // {
  //   // this can be used to debug issues with sending the webhooks.
  //   // disable "muteHttpExceptions" above to allow the exception to trigger.

  //   // log the error
  //   Logger.log(e)

  //   // split the message, so we can see what it choked on
  //   var parts = message.split(",")

  //   // error sending to Discord
  //   const result = UI.alert(
  //     `Error sending webhook to Discord.
  //     Make sure Platoons are populated and can be filled by the guild.`,
  //     UI.ButtonSet.OK);
  // }
}

// Send a Webhook to Discord
function sendPlatoonSimplifiedWebhook(byType: 'Player' | 'Unit'): void {
  const sheet: GoogleAppsScript.Spreadsheet.Sheet
  = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase: number = sheet.getRange(2, 1).getValue() as number;

  // get the webhook
  const webhookURL: string = getWebhook();
  if (webhookURL.length === 0) {
    // we need a url to proceed
    UI.alert('Discord Webhook not found (Discord!E1)', UI.ButtonSet.OK);

    return;
  }

  // mentions only works if you get the ID
  // on your Discord server, type: \@rolename, copy the value <@#######>
  const playerMentions: string[] = getPlayerMentions();
  const mentions: string = getRole();

  const title: string = getWebhookTitle(phase);
  const descriptionText: string = title + getWebhookRareIntro(phase, mentions);

  const highNeedHeroes: string = getHighNeedList(SHEETS.HEROES, getCharacterCount_());
  const highNeedShips: string = getHighNeedList(SHEETS.SHIPS, getShipCount_());

  // get data from the platoons
  let fields: string = '';
  let donations: string[][] = [];
  let groundStart: number = -1;
  for (let z: number = 0; z < MAX_PLATOON_ZONES; z += 1) {
    // for each zone
    const platoonRow: number = (z * PLATOON_ZONE_ROW_OFFSET + 2);
    let validPlatoons: string = '';
    const zone: string = getZoneName(phase, z, true);

    if (z === 1) {
      groundStart = donations.length;
    }

    if (z === 0 && phase < 3) {
      continue; // skip this and future zones
    }

    // cycle throught the platoons in a zone
    for (let p: number = 0; p < MAX_PLATOONS; p += 1) {
      const platoonData: string[][] = sheet
        .getRange(platoonRow, (p * 4) + 4, MAX_PLATOON_HEROES, 2)
        .getValues() as string[][];
      const rules: GoogleAppsScript.Spreadsheet.DataValidation[][] = sheet
        .getRange(platoonRow, (p * 4) + 5, MAX_PLATOON_HEROES, 1)
        .getDataValidations();
      const platoon: string[][] = getPlatoonDonations(
        platoonData,
        donations,
        rules,
        playerMentions,
      );

      if (platoon) {
        validPlatoons += validPlatoons.length > 0 ? `, #${p + 1}` : `#${p + 1}`;
        if (platoon.length > 0) {
          // add the new donations to the list
          for (const e of platoon) {
            donations.push([e[0], e[1]]);
          }
        }
      }
    }

    // see if all platoons are valid
    if (validPlatoons === '#1, #2, #3, #4, #5, #6') {
      validPlatoons = 'All';
    }

    // format the needed platoons
    if (validPlatoons.length > 0) {
      fields += `**${zone}**\n${validPlatoons}\n\n`;
    }
  }

  // format the high needed units
  if (highNeedShips.length > 0) {
    fields += `**High Need Ships**\n${highNeedShips}\n\n`;
  }
  if (highNeedHeroes.length > 0) {
    fields += `**High Need Heroes**\n${highNeedHeroes}\n\n`;
  }
  postMessage(webhookURL, `${descriptionText}\n\n${fields.trim()}\n`);

  // reformat the output if we need by player istead of by unit
  if (byType === 'Player') {
    const heroLabel: string = 'Heroes: ';
    const shipLabel: string = 'Ships: ';
    const playerDonations: [string, string][] = [];
    donations.forEach((donation: string[], d: number) => {
      const unit: string = donation[0];
      const names: string[] = donation[1].split(',');

      for (const name of names) {
        const nameTrim: string = name.trim();
        // see if the name is already listed
        const foundName: boolean = playerDonations.some((player: string[]) => {
          const found: boolean = player[0] === nameTrim;
          if (found) {
            if (d >= groundStart && player[1].indexOf(heroLabel) < 0) {
              player[1] += `\n${heroLabel}${unit}`;
            } else {
              player[1] += `, ${unit}`;
            }
          }

          return found;
        });

        if (!foundName) {
          playerDonations.push([
            nameTrim,
            (d >= groundStart ? heroLabel : shipLabel) + unit,
          ]);
        }
      }
    });
    // sort by player
    donations = playerDonations.sort(firstElementToLowerCase_);
  }

  // format the needed donations
  ddddd(webhookURL, byType, donations);
}

function ddddd(webhookURL: string, byType: string, donations: string[][]): void {
  const maxUrlLen: number = 1000;
  const maxCount: number = byType === 'Unit' ? 5 : 10;
  let playerFields: string = '';
  let count: number = 0;
  const f: string[][] = donations
  .filter((e: string[]) => e[1].length > 0);
  for (const e of f) {
    const fieldName: string = byType === 'Unit' ? `${e[0]} (Rare)` : e[0]; // TODO: reduce
    playerFields += `**${fieldName}**\n${e[1]}\n\n`;
    count += 1;
    // make sure our message isn't getting too long
    if (count >= maxCount || playerFields.length > maxUrlLen) {
      postMessage(webhookURL, playerFields);
      playerFields = '';
      count = 0;
    }
  }

  postMessage(webhookURL, playerFields);
}

// Send a Webhook to Discord
function sendPlatoonSimplifiedByUnitWebhook(): void {
  sendPlatoonSimplifiedWebhook('Unit');
}

// Send a Webhook to Discord
function sendPlatoonSimplifiedByPlayerWebhook(): void {
  sendPlatoonSimplifiedWebhook('Player');
}

// zone: 0, 1 or 2
function getUniquePlatoonUnits(zone: number): string[] {
  const platoonRow: number = (zone * 18) + 2;
  const sheet: GoogleAppsScript.Spreadsheet.Sheet
  = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);

  let units: string[][] = [];
  for (let platoon: number = 0; platoon < MAX_PLATOONS; platoon += 1) {
    const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(
      platoonRow,
      (platoon * 4) + 4,
      MAX_PLATOON_HEROES,
      1,
    );
    const values: string[][] = range.getValues() as string[][];
    units = units.concat(values);
  }

  // flatten the array and keep only unique values
  return units
    .map((el: string[]) => el[0])
    .unique();
}

// Get the list of Rare units needed for the phase
function getRareUnits(sheetName: string, phase: number): string {
  const tagFilter: string = getTagFilter_(); // TODO: potentially broken if TB not sync
  const useBottomTerritory = !isLight_(tagFilter) || phase > 1;
  const count: number = getCharacterCount_();
  let data: [string, number][] = SPREADSHEET
    .getSheetByName(sheetName)
    .getRange(1, 1, count + 1, 8)
    .getValues() as [string, number][];
  const idx: number = phase + 1;

  // Drop first line
  data = data.slice(1);

  // cycle through each unit
  let units: string[] = data
    .filter((row: [string, number]) => row[0].length > 0 && row[idx] < RARE_MAX)
    .map((row: [string, number]) => row[0]) // keep only the unit's name
    .sort(); // sort the list of units

  let platoonUnits: string[];
  if (sheetName === SHEETS.SHIPS) {
    platoonUnits = getUniquePlatoonUnits(0);
  } else {
    platoonUnits = getUniquePlatoonUnits(1);
    if (useBottomTerritory) {
      platoonUnits = platoonUnits.concat(getUniquePlatoonUnits(2));
    }
  }

  // filter out rare units that do not appear in platoons
  units = units.filter((unit: string) => platoonUnits.some((el: string) => el === unit));

  return units.join('\n');
}

// Send a message to Discord that lists all units to watch out for in the current phase
function allRareUnitsWebhook(): void {
  const phase: number = SPREADSHEET
    .getSheetByName(SHEETS.PLATOONS)
    .getRange(2, 1)
    .getValue() as number;

  const webhookURL: string = getWebhook(); // get the webhook
  if (webhookURL.length === 0) {
    // we need a url to proceed
    UI.alert('Discord Webhook not found (Discord!E1)', UI.ButtonSet.OK);

    return;
  }

  const fields: object[] = [];

  if (phase >= 3) {
    // TODO: remove hardcode
    // get the ships list
    const ships: string = getRareUnits(SHEETS.SHIPS, phase);
    if (ships.length > 0) {
      fields.push({
        name: 'Rare Ships',
        value: ships,
        inline: true,
      });
    }
  }

  // get the hero list
  const heroes: string = getRareUnits(SHEETS.HEROES, phase);
  if (heroes.length > 0) {
    fields.push({
      name: 'Rare Heroes',
      value: heroes,
      inline: true,
    });
  }

  // make sure we're not trying to send empty data
  if (fields.length === 0) {
    // no data to send
    fields.push({
      name: 'Rare Heroes',
      value: 'There Are No Rare Units For This Phase.',
      inline: true,
    });
  }

  const title: string = getWebhookTitle(phase);
  // mentions only works if you get the id
  // in Discord: Settings - Appearance - Enable Developer Mode
  // type: \@rolename, copy the value <@$####>
  const mentions: string = getRole();
  const warnIntro: string = getWebhookWarnIntro(phase, mentions);
  const desc: string = getWebhookDesc(phase);

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = urlFetchMakeParam_({
    content: `${title}${warnIntro}${desc}`,
    embeds: [{ fields }],
  });
  urlFetchExecute_(webhookURL, options);
}

// ****************************************
// Timer Functions
// ****************************************

// Figure out what phase the TB is in
function setCurrentPhase(): void {
  // get the guild's TB start date/time and phase length in hours
  const startTime: Date = getTBStartTime();
  const phaseHours: number = getPhaseHours();
  if (startTime && phaseHours) {
    const msPerHour: number = 1000 * 60 * 60;
    const now: Date = new Date();
    const diff: number = now.getTime() - startTime.getTime();
    const hours: number = diff / msPerHour + 1; // add 1 hour to ensure we are in the next phase
    const phase: number = Math.ceil(hours / phaseHours);
    const maxPhases: number = 6;

    // set the phase in Platoons tab
    if (phase <= maxPhases) {
      SPREADSHEET
        .getSheetByName(SHEETS.PLATOONS)
        .getRange(2, 1)
        .setValue(phase);
    }
  }
}

// Callback function to see if we should send the webhook
function sendTimedWebhook(): void {
  setCurrentPhase(); // set the current phase based on time

  // reset the platoons if clear flag was set
  if (getWebhookClear()) {
    resetPlatoons();
  }
  allRareUnitsWebhook(); // call the webhook
  registerWebhookTimer(); // register the next timer
}

// Try to create a webhook trigger
function registerWebhookTimer(): void {
  // get the guild's TB start date/time and phase length in hours
  const startTime: Date = getTBStartTime();
  const phaseHours: number = getPhaseHours();
  if (startTime && phaseHours) {
    const msPerHour: number = 1000 * 60 * 60;
    const phaseMs: number = phaseHours * msPerHour;
    const target: Date = new Date(startTime);
    const now: Date = new Date();
    const maxPhases: number = 6;

    // remove the trigger
    const triggers: GoogleAppsScript.Script.Trigger[]
    = ScriptApp.getProjectTriggers()
    .filter((e: GoogleAppsScript.Script.Trigger) => e.getHandlerFunction() === 'sendTimedWebhook');
    for (const trigger of triggers) {
      ScriptApp.deleteTrigger(trigger);
    }

    // see if we can set the trigger later in the phase
    for (let i: number = 2; i <= maxPhases; i += 1) {
      target.setTime(target.getTime() + phaseMs);

      if (target > now) {
        // target is in the future
        // found the start of the next phase, so set the timer
        ScriptApp.newTrigger('sendTimedWebhook')
          .timeBased()
          .at(target)
          .create();

        break;
      }
    }
  }
}

// type Trigger = GoogleAppsScript.Script.Trigger;
