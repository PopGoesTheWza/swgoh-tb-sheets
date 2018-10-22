// ****************************************
// Webhooks Functions
// ****************************************

interface DiscordMessageEmbedFields {
  name: string;
  value: string;
  inline?: boolean;
}

/** Get the title for the webhooks */
function getWebhookTitle_(phase: number): string {

  const defaultVal = `__**Territory Battle: Phase ${phase}**__`;

  return `${getWebhookTemplate_(phase, WEBHOOK_TITLE_ROW, defaultVal)}`;
}

/** Get the intro for the warning webhook */
function getWebhookWarnIntro_(phase: number, mention: string): string {

  const defaultVal = `Here are the __Rare Units__ to watch out for in __Phase ${phase}__.
**Check with an officer before donating to Platoons/Squadrons that require them.**`;

  return `\n\n${getWebhookTemplate_(phase, WEBHOOK_WARN_ROW, defaultVal)} ${mention}`;
}

/** Get the intro for the rare by webhook */
function getWebhookRareIntro_(phase: number, mention: string): string {

  const defaultVal =
    `Here are the Safe Platoons and the Rare Platoon donations for __Phase ${phase}__.
**Do not donate heroes to the other Platoons.**`;

  return `\n\n${getWebhookTemplate_(phase, WEBHOOK_RARE_ROW, defaultVal)} ${mention}`;
}

/** Get the intro for the depth webhook */
function getWebhookDepthIntro_(phase: number, mention: string): string {

  const defaultVal = `Here are the Platoon assignments for __Phase ${phase}__.
**Do not donate heroes to the other Platoons.**`;

  return `\n\n${getWebhookTemplate_(phase, WEBHOOK_DEPTH_ROW, defaultVal)} ${mention}`;
}

/** Get the player Discord IDs for mentions */
function getPlayerMentions_(): KeyedArray {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.DISCORD);
  const data = sheet.getRange(2, 1, sheet.getLastRow(), 2)
    .getValues() as string[][];
  const result: KeyedArray = {};

  for (const e of data) {
    const name = e[0];
    // only stores unique names, we can't differentiate with duplicates
    if (name && name.length > 0 && !result[name]) {
      // store the ID if it exists, otherwise store the player's name
      result[name] = (e[1] && e[1].length > 0) ? e[1] : name;
    }
  }

  return result;
}

/** Get a string representing the platoon assignements */
function getPlatoonString_(platoon: string[][]): string {

  const results: string[] = [];

  // cycle through the heroes
  for (let h = 0; h < MAX_PLATOON_HEROES; h += 1) {
    if (platoon[h][1].length === 0 || platoon[h][1] === 'Skip') {
      // impossible platoon
      return '';  // TODO: return undefined
    }

    // remove the gear
    let name = platoon[h][1];
    const endIdx = name.indexOf(' (');
    if (endIdx > -1) {
      name = name.substring(0, endIdx);
    }

    // add the assignement
    results.push(`**${platoon[h][0]}**: ${name}`);
  }

  return results.join('\n');
}

/** Get the formatted zone name with location descriptor */
function getZoneName_(phase: number, zoneNum: number, full: boolean): string {

  const zone = SPREADSHEET.getSheetByName(SHEETS.PLATOONS)
    .getRange((zoneNum * PLATOON_ZONE_ROW_OFFSET) + 4, 1)
    .getValue() as string;
  let loc: string;

  switch (zoneNum) {
    case 0:
      loc = '(Top)';
      break;
    case 2:
      loc = '(Bottom)';
      break;
    case 1:
    default:
      loc = (phase === 2) ? '(Top)' : '(Middle)';
  }
  const result = (full && phase !== 1)
    ? `${zone} ${loc} ${zoneNum === 0 ? 'Squadrons' : 'Platoons'}`
    : `${loc} ${zone}`;

  return result;
}

/** Send a Webhook to Discord */
function sendPlatoonDepthWebhook(): void {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = sheet.getRange(2, 1).getValue() as number;

  // get the webhook
  const webhookURL = getWebhook_();
  if (webhookURL.length === 0) {
    // we need a url to proceed
    UI.alert('Discord Webhook not found (Discord!E1)', UI.ButtonSet.OK);

    return;
  }

  // mentions only works if you get the id
  // in Settings - Appearance - Enable Developer Mode, type: \@rolename, copy the value <@$####>
  const descriptionText = `${getWebhookTitle_(phase)}${getWebhookDepthIntro_(phase, getRole_())}`;

  // get data from the platoons
  const fields: DiscordMessageEmbedFields[] = [];
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    if (z === 0 && phase < 3) {
      continue; // skip this zone
    }

    // for each zone
    const platoonRow = (z * PLATOON_ZONE_ROW_OFFSET) + 2;
    const zone = getZoneName_(phase, z, false);

    // cycle throught the platoons in a zone
    for (let p = 0; p < MAX_PLATOONS; p += 1) {
      const platoonData = sheet.getRange(platoonRow, (p * 4) + 4, MAX_PLATOON_HEROES, 2)
        .getValues() as string[][];
      const platoon = getPlatoonString_(platoonData);

      if (platoon.length > 0) {
        fields.push({
          name: `${zone}: #${p + 1}`,
          value: platoon,
          inline: true,
        });
      }
    }
  }

  const options = urlFetchMakeParam_({
    content: descriptionText,
    embeds: [{ fields }],
  });
  urlFetchExecute_(webhookURL, options);
}

/** Get an array representing the new platoon assignements */
function getPlatoonDonations_(platoon: string[][],
                              donations: string[][],
                              rules: GoogleAppsScript.Spreadsheet.DataValidation[][],
                              playerMentions: KeyedArray): string[][] {

  const result: string[][] = [];

  // cycle through the heroes
  for (let h = 0; h < MAX_PLATOON_HEROES; h += 1) {
    if (platoon[h][0].length === 0) {
      continue; // no unit needed here
    }

    if (platoon[h][1].length === 0 || platoon[h][1] === 'Skip') {
      return undefined; // impossible platoon
    }

    // see if the hero is already in donations
    const heroDonated = donations.some(e => e[0] === platoon[h][0])
      || result.some(e => e[0] === platoon[h][0]);

    if (!heroDonated) {

      type RequireValueInListCriteria = [
        [string],  // Array of player name
        boolean  // true for dropdown
      ];

      const criteria = rules[h][0].getCriteriaValues() as RequireValueInListCriteria;

      // only add rare donations
      if (criteria[0].length < RARE_MAX) {
        const sorted = criteria[0].sort(lowerCase_);
        const names: string[] = [];
        for (const name of sorted) {
          const mention = playerMentions[name];
          names.push(mention ? `${name} (${mention})` : `${name}`);
        }
        // add the recommendations
        result.push([platoon[h][0], names.join(', ')]);
      }
    }
  }

  return result;
}

/** Get a list of units that are required a high number of times */
function getHighNeedList_(sheetName: string, unitCount: number): string[] {

  const counts = SPREADSHEET.getSheetByName(sheetName)
    .getRange(2, 1, unitCount, HERO_PLAYER_COL_OFFSET)
    .getValues() as [string, number][];
  const results: string[] = [];
  const idx = HERO_PLAYER_COL_OFFSET - 1;

  for (const row of counts) {
    // unit's count is over the min bar to be too high
    if (row[idx] >= HIGH_MIN) {
      results.push(`${row[0]} (${row[idx]})`);
    }
  }

  return results;
}

/** Send the message to Discord */
function postMessage_(webhookURL: string, message: string): void {

  const options = urlFetchMakeParam_({ content: message.trim() });
  urlFetchExecute_(webhookURL, options);
}

/** Send a Webhook to Discord */
function sendPlatoonSimplifiedWebhook_(byType: 'Player' | 'Unit'): void {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = sheet.getRange(2, 1).getValue() as number;

  // get the webhook
  const webhookURL = getWebhook_();
  if (webhookURL.length === 0) {  // we need a url to proceed
    UI.alert('Discord Webhook not found (Discord!E1)', UI.ButtonSet.OK);

    return;
  }

  // mentions only works if you get the ID
  // on your Discord server, type: \@rolename, copy the value <@#######>
  const playerMentions = getPlayerMentions_();
  const mentions = getRole_();

  const descriptionText = `${getWebhookTitle_(phase)}${getWebhookRareIntro_(phase, mentions)}`;

  // get data from the platoons
  const fields: string[] = [];
  let donations: string[][] = [];
  let groundStart = -1;

  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {  // for each zone
    const platoonRow = (z * PLATOON_ZONE_ROW_OFFSET + 2);
    const validPlatoons: number[] = [];
    const zone = getZoneName_(phase, z, true);

    if (z === 1) {
      groundStart = donations.length;
    }

    if (z !== 0 || phase > 2) {

      // cycle throught the platoons in a zone
      for (let p = 0; p < MAX_PLATOONS; p += 1) {
        const platoonData = sheet.getRange(platoonRow, (p * 4) + 4, MAX_PLATOON_HEROES, 2)
          .getValues() as string[][];
        const rules = sheet.getRange(platoonRow, (p * 4) + 5, MAX_PLATOON_HEROES, 1)
          .getDataValidations();
        const platoon = getPlatoonDonations_(platoonData, donations, rules, playerMentions);

        if (platoon) {
          validPlatoons.push(p);
          if (platoon.length > 0) {
            // add the new donations to the list
            for (const e of platoon) {
              donations.push([e[0], e[1]]);
            }
          }
        }
      }
    }

    // see if all platoons are valid
    let platoons: string;
    if (validPlatoons.length === MAX_PLATOONS) {
      platoons = 'All';
    } else {
      platoons = validPlatoons.map(e => `#${e + 1}`).join(', ');
    }

    // format the needed platoons
    if (validPlatoons.length > 0) {
      fields.push(`**${zone}**\n${platoons}`);
    }
  }

  // format the high needed units
  const highNeedShips = getHighNeedList_(SHEETS.SHIPS, getShipCount_());
  if (highNeedShips.length > 0) {
    fields.push(`**High Need Ships**\n${highNeedShips.join(', ')}`);
  }
  const highNeedHeroes = getHighNeedList_(SHEETS.HEROES, getCharacterCount_());
  if (highNeedHeroes.length > 0) {
    fields.push(`**High Need Heroes**\n${highNeedHeroes.join(', ')}`);
  }
  postMessage_(webhookURL, `${descriptionText}\n\n${fields.join('\n\n')}\n`);

  // reformat the output if we need by player istead of by unit
  if (byType === 'Player') {
    const heroLabel = 'Heroes: ';
    const shipLabel = 'Ships: ';

    const acc  = donations.reduce(
      (acc: [string, string][], e, i) => {

        const unit = e[0];
        const names = e[1].split(',');
        for (const name of names) {
          const nameTrim = name.trim();
          // see if the name is already listed
          const foundName = acc.some((player) => {
            const found = player[0] === nameTrim;
            if (found) {
              player[1] += (i >= groundStart && player[1].indexOf(heroLabel) < 0)
                ? `\n${heroLabel}${unit}`
                : `, ${unit}`;
            }

            return found;
          });

          if (!foundName) {
            acc.push([
              nameTrim,
              (i >= groundStart ? heroLabel : shipLabel) + unit,
            ]);
          }
        }

        return acc;
      },
      [],
    );
    // sort by player
    donations = acc.sort(firstElementToLowerCase_);
  }

  // format the needed donations
  spoolDiscordMessage_(webhookURL, byType, donations);
}

function spoolDiscordMessage_(webhookURL: string, byType: string, donations: string[][]): void {

  const typeIsUnit = byType === 'Unit';
  const maxUrlLen = 1000;
  const maxCount = typeIsUnit ? 5 : 10;
  const acc = donations.reduce(
    (acc, e) => {
      if (e[1].length > 0) {
        const f = typeIsUnit ? `${e[0]} (Rare)` : e[0];
        const s = `**${f}**\n${e[1]}\n\n`;
        acc.count += s.length;
        acc.fields.push(s);
        // make sure our message isn't getting too long
        if (acc.fields.length >= maxCount || acc.count > maxUrlLen) {
          postMessage_(webhookURL, acc.fields.join(''));
          acc.count = 0;
          acc.fields = [];
        }
      }
      return acc;
    },
    {
      count: 0,
      fields: [],
    },
  );
  if (acc.fields.length > 0) {
    postMessage_(webhookURL, acc.fields.join(''));
  }
}

/** Send a Webhook to Discord */
function sendPlatoonSimplifiedByUnitWebhook(): void {
  sendPlatoonSimplifiedWebhook_('Unit');
}

/** Send a Webhook to Discord */
function sendPlatoonSimplifiedByPlayerWebhook(): void {
  sendPlatoonSimplifiedWebhook_('Player');
}

/** zone: 0, 1 or 2 */
function getUniquePlatoonUnits_(zone: number): string[] {

  const platoonRow = (zone * 18) + 2;
  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);

  let units: string[][] = [];
  for (let platoon = 0; platoon < MAX_PLATOONS; platoon += 1) {
    const range = sheet.getRange(platoonRow, (platoon * 4) + 4, MAX_PLATOON_HEROES, 1);
    const values = range.getValues() as string[][];
    units = units.concat(values);
  }

  // flatten the array and keep only unique values
  return units
    .map(e => e[0])
    .unique();
}

/** Get the list of Rare units needed for the phase */
function getRareUnits_(sheetName: string, phase: number): string[] {

  const useBottomTerritory = !isLight_(getSideFilter_()) || phase > 1;
  const count = getCharacterCount_() + 1;
  const data = (SPREADSHEET.getSheetByName(sheetName)
    .getRange(1, 1, count, 8)
    .getValues() as [string, number][])
    .slice(1);  // Drop first line

  const idx = phase + 2;  // HEROES/SHIPS, column D

  // cycle through each unit
  const units: string[] = data.reduce(
    (acc: [string], row) => {
      if (row[0].length > 0 && row[idx] < RARE_MAX) {
        acc.push(row[0]);
      }
      return acc;
    },
    [])  // keep only the unit's name
    .sort();  // sort the list of units

  let platoonUnits: string[];
  if (sheetName === SHEETS.SHIPS) {
    platoonUnits = getUniquePlatoonUnits_(0);
  } else {
    platoonUnits = getUniquePlatoonUnits_(1);
    if (useBottomTerritory) {
      platoonUnits = platoonUnits.concat(getUniquePlatoonUnits_(2));
    }
  }

  // filter out rare units that do not appear in platoons
  const results = units.filter(unit => platoonUnits.some(el => el === unit));

  return results;
}

/** Send a message to Discord that lists all units to watch out for in the current phase */
function allRareUnitsWebhook(): void {

  const phase = SPREADSHEET.getSheetByName(SHEETS.PLATOONS)
    .getRange(2, 1)
    .getValue() as number;

  const webhookURL = getWebhook_(); // get the webhook
  if (webhookURL.length === 0) {
    // we need a url to proceed
    UI.alert('Discord Webhook not found (Discord!E1)', UI.ButtonSet.OK);

    return;
  }

  const fields: DiscordMessageEmbedFields[] = [];

  // TODO: regroup phases and zones management
  if (phase >= 3) {
    // get the ships list
    const ships = getRareUnits_(SHEETS.SHIPS, phase);
    if (ships.length > 0) {
      fields.push({
        name: 'Rare Ships',
        value: ships.join('\n'),
        inline: true,
      });
    }
  }

  // get the hero list
  const heroes = getRareUnits_(SHEETS.HEROES, phase);
  if (heroes.length > 0) {
    fields.push({
      name: 'Rare Heroes',
      value: heroes.join('\n'),
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

  const title = getWebhookTitle_(phase);
  // mentions only works if you get the id
  // in Discord: Settings - Appearance - Enable Developer Mode
  // type: \@rolename, copy the value <@$####>
  const mentions = getRole_();
  const warnIntro = getWebhookWarnIntro_(phase, mentions);
  const desc = getWebhookDesc_(phase);

  const options = urlFetchMakeParam_({
    content: `${title}${warnIntro}${desc}`,
    embeds: [{ fields }],
  });
  urlFetchExecute_(webhookURL, options);
}

// ****************************************
// Timer Functions
// ****************************************

/** Figure out what phase the TB is in */
function setCurrentPhase_(): void {

  // get the guild's TB start date/time and phase length in hours
  const startTime = getTBStartTime_();
  const phaseHours = getPhaseHours_();
  if (startTime && phaseHours) {
    const msPerHour = 1000 * 60 * 60;
    const now = new Date();
    const diff = now.getTime() - startTime.getTime();
    const hours = diff / msPerHour + 1; // add 1 hour to ensure we are in the next phase
    const phase = Math.ceil(hours / phaseHours);
    const maxPhases = 6;

    // set the phase in Platoons tab
    if (phase <= maxPhases) {
      SPREADSHEET.getSheetByName(SHEETS.PLATOONS)
        .getRange(2, 1)
        .setValue(phase);
    }
  }
}

/** Callback function to see if we should send the webhook */
function sendTimedWebhook(): void {

  setCurrentPhase_(); // set the current phase based on time

  // reset the platoons if clear flag was set
  if (getWebhookClear_()) {
    resetPlatoons();
  }
  allRareUnitsWebhook(); // call the webhook
  registerWebhookTimer(); // register the next timer
}

/** Try to create a webhook trigger */
function registerWebhookTimer(): void {

  // get the guild's TB start date/time and phase length in hours
  const startTime = getTBStartTime_();
  const phaseHours = getPhaseHours_();
  if (startTime && phaseHours) {
    const msPerHour = 1000 * 60 * 60;
    const phaseMs = phaseHours * msPerHour;
    const target = new Date(startTime);
    const now = new Date();
    const maxPhases = 6;

    // remove the trigger
    const triggers = ScriptApp.getProjectTriggers()
      .filter(e => e.getHandlerFunction() === sendTimedWebhook.name);
    for (const trigger of triggers) {
      ScriptApp.deleteTrigger(trigger);
    }

    // see if we can set the trigger later in the phase
    for (let i = 2; i <= maxPhases; i += 1) {
      target.setTime(target.getTime() + phaseMs);

      if (target > now) {
        // target is in the future
        // found the start of the next phase, so set the timer
        ScriptApp.newTrigger(sendTimedWebhook.name)
          .timeBased()
          .at(target)
          .create();

        break;
      }
    }
  }
}
