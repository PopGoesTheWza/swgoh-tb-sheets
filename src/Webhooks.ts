namespace discord {

  export interface MessageEmbedFields {
    name: string;
    value: string;
    inline?: boolean;
  }

  /** Get the title for the webhooks */
  export function getTitle(phase: TerritoryBattles.phaseIdx): string {

    const defaultVal = `__**Territory Battle: Phase ${phase}**__`;

    return `${config.discord.webhookTemplate(phase, WEBHOOK_TITLE_ROW, defaultVal)}`;
  }

  /** Get the formatted zone name with location descriptor */
  export function getZoneName(
    phase: TerritoryBattles.phaseIdx,
    zoneNum: number,
    full: boolean,
  ): string {

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

  /** Get a string representing the platoon assignements */
  export function getPlatoonString(platoon: string[][]): string {

    const results: string[] = [];

    // cycle through the heroes
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {
      if (platoon[h][1].length === 0 || platoon[h][1] === 'Skip') {
        // impossible platoon
        return undefined;
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

  /** Get the member Discord IDs for mentions */
  export function getMemberMentions(): KeyedStrings {

    const sheet = SPREADSHEET.getSheetByName(SHEETS.DISCORD);
    const data = sheet.getRange(2, 1, sheet.getLastRow(), 2)
      .getValues() as string[][];
    const result: KeyedStrings = {};

    for (const e of data) {
      const name = e[0];
      // only stores unique names, we can't differentiate with duplicates
      if (name && name.length > 0 && !result[name]) {
        // store the ID if it exists, otherwise store the member's name
        result[name] = (e[1] && e[1].length > 0) ? e[1] : name;
      }
    }

    return result;
  }

  /** Get an array representing the new platoon assignements */
  function getPlatoonDonations(
    platoon: string[][],
    donations: string[][],
    rules: DataValidation[][],
    memberMentions: KeyedStrings,
  ): string[][] {

    const result: string[][] = [];

    // cycle through the heroes
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {
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
          [string],  // Array of members name
          boolean  // true for dropdown
        ];

        const criteria = rules[h][0].getCriteriaValues() as RequireValueInListCriteria;

        // only add rare donations
        if (criteria[0].length < RARE_MAX) {
          const sorted = criteria[0].sort(caseInsensitive_);
          const names: string[] = [];
          for (const name of sorted) {
            const mention = memberMentions[name];
            names.push(mention ? `${name} (${mention})` : `${name}`);
          }
          // add the recommendations
          result.push([platoon[h][0], names.join(', ')]);
        }
      }
    }

    return result;
  }

  /** Get the intro for the depth webhook */
  export function getDepthIntro(phase: TerritoryBattles.phaseIdx, mention: string): string {

    const defaultVal = `Here are the Platoon assignments for __Phase ${phase}__.
  **Do not donate heroes to the other Platoons.**`;

    return `\n\n${config.discord.webhookTemplate(phase, WEBHOOK_DEPTH_ROW, defaultVal)} ${mention}`;
  }

  /** Get the intro for the rare by webhook */
  function getRareIntro(phase: TerritoryBattles.phaseIdx, mention: string): string {

    const defaultVal =
      `Here are the Safe Platoons and the Rare Platoon donations for __Phase ${phase}__.
  **Do not donate heroes to the other Platoons.**`;

    return `\n\n${config.discord.webhookTemplate(phase, WEBHOOK_RARE_ROW, defaultVal)} ${mention}`;
  }

  /** Get the intro for the warning webhook */
  export function getWarnIntro(phase: TerritoryBattles.phaseIdx, mention: string): string {

    const defaultVal = `Here are the __Rare Units__ to watch out for in __Phase ${phase}__.
  **Check with an officer before donating to Platoons/Squadrons that require them.**`;

    return `\n\n${config.discord.webhookTemplate(phase, WEBHOOK_WARN_ROW, defaultVal)} ${mention}`;
  }

  /** Send the message to Discord */
  function postMessage(webhookURL: string, message: string): void {

    const options = urlFetchMakeParam_({ content: message.trim() });
    urlFetchExecute_(webhookURL, options);
  }

  function messageSpooler(webhookURL: string, byType: string, donations: string[][]): void {

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
            postMessage(webhookURL, acc.fields.join(''));
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
      postMessage(webhookURL, acc.fields.join(''));
    }
  }

  /** Send a Webhook to Discord */
  export function sendPlatoonSimplified(byType: 'Player' | 'Unit'): void {

    const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
    const phase = config.currentPhase();

    // get the webhook
    const webhookURL = config.discord.webhookUrl();
    if (webhookURL.length === 0) {  // we need a url to proceed
      const UI = SpreadsheetApp.getUi();
      UI.alert(
        'Configuration Error',
        'Discord webhook not found (Discord!E1)',
        UI.ButtonSet.OK,
      );

      return;
    }

    // mentions only works if you get the ID
    // on your Discord server, type: \@rolename, copy the value <@#######>
    const memberMentions = getMemberMentions();
    const mentions = config.discord.roleId();

    const descriptionText = `${discord.getTitle(phase)}${getRareIntro(phase, mentions)}`;

    // get data from the platoons
    const fields: string[] = [];
    let donations: string[][] = [];
    let groundStart = -1;

    for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {  // for each zone
      const platoonRow = (z * PLATOON_ZONE_ROW_OFFSET + 2);
      const validPlatoons: number[] = [];
      const zone = discord.getZoneName(phase, z, true);

      if (z === 1) {
        groundStart = donations.length;
      }

      if (z !== 0 || phase > 2) {

        // cycle throught the platoons in a zone
        for (let p = 0; p < MAX_PLATOONS; p += 1) {
          const platoonData = sheet.getRange(platoonRow, (p * 4) + 4, MAX_PLATOON_UNITS, 2)
            .getValues() as string[][];
          const rules = sheet.getRange(platoonRow, (p * 4) + 5, MAX_PLATOON_UNITS, 1)
            .getDataValidations();
          const platoon = getPlatoonDonations(platoonData, donations, rules, memberMentions);

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
    const heroesTable = new Units.Heroes();
    const highNeedShips = heroesTable.getHighNeedList();
    if (highNeedShips.length > 0) {
      fields.push(`**High Need Ships**\n${highNeedShips.join(', ')}`);
    }
    const shipsTable = new Units.Ships;
    const highNeedHeroes = shipsTable.getHighNeedList();
    if (highNeedHeroes.length > 0) {
      fields.push(`**High Need Heroes**\n${highNeedHeroes.join(', ')}`);
    }
    postMessage(webhookURL, `${descriptionText}\n\n${fields.join('\n\n')}\n`);

    // reformat the output if we need by member istead of by unit
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
            const foundName = acc.some((member) => {
              const found = member[0] === nameTrim;
              if (found) {
                member[1] += (i >= groundStart && member[1].indexOf(heroLabel) < 0)
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
      // sort by member
      donations = acc.sort((a, b) => caseInsensitive_(a[0], b[0]));
    }

    // format the needed donations
    messageSpooler(webhookURL, byType, donations);
  }

  // ****************************************
  // Timer Functions
  // ****************************************

  /** Figure out what phase the TB is in */
  export function setCurrentPhase(): void {

    // get the guild's TB start date/time and phase length in hours
    const startTime = config.discord.startTime();
    const phaseHours = config.discord.phaseDuration();
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

}

/** Send a Webhook to Discord */
function sendPlatoonDepthWebhook(): void {

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

  // mentions only works if you get the id
  // in Settings - Appearance - Enable Developer Mode, type: \@rolename, copy the value <@$####>
  const descriptionText =
    `${discord.getTitle(phase)}${discord.getDepthIntro(phase, config.discord.roleId())}`;

  // get data from the platoons
  const fields: discord.MessageEmbedFields[] = [];
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    if (z === 0 && phase < 3) {
      continue; // skip this zone
    }

    // for each zone
    const platoonRow = (z * PLATOON_ZONE_ROW_OFFSET) + 2;
    const zone = discord.getZoneName(phase, z, false);

    // cycle throught the platoons in a zone
    for (let p = 0; p < MAX_PLATOONS; p += 1) {
      const platoonData = sheet.getRange(platoonRow, (p * 4) + 4, MAX_PLATOON_UNITS, 2)
        .getValues() as string[][];
      const platoon = discord.getPlatoonString(platoonData);

      if (platoon && platoon.length > 0) {
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

/** Send a Webhook to Discord */
function sendPlatoonSimplifiedByUnitWebhook(): void {
  discord.sendPlatoonSimplified('Unit');
}

/** Send a Webhook to Discord */
function sendPlatoonSimplifiedByMemberWebhook(): void {
  discord.sendPlatoonSimplified('Player');
}

/** Send a message to Discord that lists all units to watch out for in the current phase */
function allRareUnitsWebhook(): void {

  const phase = config.currentPhase();

  const webhookURL = config.discord.webhookUrl(); // get the webhook
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

  const fields: discord.MessageEmbedFields[] = [];

  // TODO: regroup phases and zones management
  if (phase >= 3) {
    // get the ships list
    const shipsTable = new Units.Ships();
    const ships = shipsTable.getNeededRareList(phase);
    if (ships.length > 0) {
      fields.push({
        name: 'Rare Ships',
        value: ships.join('\n'),
        inline: true,
      });
    }
  }

  // get the hero list
  const heroesTable = new Units.Heroes();
  const heroes = heroesTable.getNeededRareList(phase);
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

  const title = discord.getTitle(phase);
  // mentions only works if you get the id
  // in Discord: Settings - Appearance - Enable Developer Mode
  // type: \@rolename, copy the value <@$####>
  const mentions = config.discord.roleId();
  const warnIntro = discord.getWarnIntro(phase, mentions);
  const desc = config.discord.webhookDescription(phase);

  const options = urlFetchMakeParam_({
    content: `${title}${warnIntro}${desc}`,
    embeds: [{ fields }],
  });
  urlFetchExecute_(webhookURL, options);
}

/** Callback function to see if we should send the webhook */
function sendTimedWebhook(): void {

  discord.setCurrentPhase(); // set the current phase based on time

  // reset the platoons if clear flag was set
  if (config.discord.resetPlatoons()) {
    resetPlatoons();
  }
  allRareUnitsWebhook(); // call the webhook
  registerWebhookTimer(); // register the next timer
}

/** Try to create a webhook trigger */
function registerWebhookTimer(): void {

  // get the guild's TB start date/time and phase length in hours
  const startTime = config.discord.startTime();
  const phaseHours = config.discord.phaseDuration();
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
