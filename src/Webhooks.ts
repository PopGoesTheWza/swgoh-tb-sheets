namespace discord {
  export interface RichEmbedOptionsField {
    name: string;
    value: string;
    inline?: boolean;
  }

  /** Get the title for the webhooks */
  export function getTitle(phase: TerritoryBattles.phaseIdx): string {
    const WEBHOOK_TITLE_ROW = 5;
    const defaultVal = `__**Territory Battle: Phase ${phase}**__`;

    return `${config.discord.webhookTemplate(phase, WEBHOOK_TITLE_ROW, defaultVal)}`;
  }

  // TODO: rework so that it no longer rely on a single `phase` value
  export function isTerritory(
    territory: number,
    phase = config.currentPhase(),
    event = config.currentEvent(),
  ): boolean {
    return isGeoDS_(event) //Could be Static lookup from sheet
      ? territory !== 0 || phase > 1
      : isGeoLS_(event)
        ? territory !== 0 || phase > 0
        : isHothDS_(event)
          ? territory !== 0 || phase > 2
          : isHothLS_(event)
            ? (territory !== 0 || phase > 2) && (territory !== 2 || phase > 1)
            : false;
  }

  // TODO: rework so that it no longer rely on a single `phase` value
  export function requiredRarity(
    territory: number,
    phase = config.currentPhase(),
    event = config.currentEvent(),
  ): number {
    return +utils.getSheetByNameOrDie(SHEET.META).getRange(5, 3).getValue(); //Static lookup from sheet
  }

  export function getSimplifiedPlatoons(phase: TerritoryBattles.phaseIdx) {
    const sheet = utils.getSheetByNameOrDie(SHEET.PLATOON);
    const grid: string[][] = [];

    for (let zone = 0; zone < MAX_PLATOON_ZONES; zone += 1) {
      const platoons: string[] = [];
      grid[zone] = platoons;
      for (let platoon = 0; platoon < MAX_PLATOONS; platoon += 1) {
        const cur = new PlatoonDetails(phase, zone, platoon);

        const unitRange = sheet.getRange(cur.row, cur.column, MAX_PLATOON_UNITS);
        platoons[platoon] = cur.exist
          ? /** forbidden */ unitRange.offset(15, 1, 1, 1).getValue() === TerritoryBattles.SKIP_BUTTON_CHECKED ||
            /** incomplete */ unitRange
              .offset(0, 1, MAX_PLATOON_UNITS)
              .getValues()
              .findIndex((e) => `${e[0]}`.trim().length === 0) !== -1
            ? '🚫' // or UNAVAILABLE ⛔
            : utils.EMOJI_KEYCAP_DIGITS[platoon + 1] // or OPEN ✅
          : '';
      }
    }

    return grid.map((platoons) => platoons.join('')).join('\n');
  }

  /** Get a string representing the platoon assignements */
  export function getPlatoonString(platoonData: string[][]): string | undefined {
    const SKIPPED_PLATOON_LABEL = TerritoryBattles.SKIPPED_PLATOON_LABEL;
    const results: string[] = [];

    // cycle through the heroes
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {
      const row = platoonData[h];
      const playerName = row[1].trim();
      const unitName = row[0];
      if (playerName.length === 0 || playerName === SKIPPED_PLATOON_LABEL) {
        return undefined; // impossible platoon
      }

      // check to remove the gear
      const endIdx = playerName.indexOf(' (');

      // add the assignement
      results.push(`**${unitName}**: ${endIdx > -1 ? playerName.substring(0, endIdx) : playerName}`);
    }

    return results.join('\n');
  }

  /** Get the member Discord IDs for mentions */
  export function getMemberMentions(): KeyedStrings {
    const sheet = utils.getSheetByNameOrDie(SHEET.DISCORD);
    const DISCORD_MEMBERMENTIONS_ROW = 2;
    const DISCORD_MEMBERMENTIONS_COL = 1;
    const DISCORD_MEMBERMENTIONS_NUMROWS = sheet.getLastRow();
    const DISCORD_MEMBERMENTIONS_NUMCOLS = 2;
    const data = sheet
      .getRange(
        DISCORD_MEMBERMENTIONS_ROW,
        DISCORD_MEMBERMENTIONS_COL,
        DISCORD_MEMBERMENTIONS_NUMROWS,
        DISCORD_MEMBERMENTIONS_NUMCOLS,
      )
      .getValues() as string[][];

    const result: KeyedStrings = {};

    for (const e of data) {
      const name = e[0];
      // only stores unique names, we can't differentiate with duplicates
      if (name && name.length > 0 && !result[name]) {
        // store the ID if it exists, otherwise store the member's name
        result[name] = e[1] && e[1].length > 0 ? e[1] : name;
      }
    }

    return result;
  }

  /** Get an array representing the new platoon assignements */
  function getPlatoonDonations(
    platoonData: string[][],
    donations: string[][],
    rules: (Spreadsheet.DataValidation | null)[][],
    memberMentions: KeyedStrings,
    neededUnits: KeyedNumbers,
  ): string[][] | undefined {
    const SKIPPED_PLATOON_LABEL = TerritoryBattles.SKIPPED_PLATOON_LABEL;
    const result: string[][] = [];

    // cycle through the heroes
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {
      const row = platoonData[h];
      const playerName = row[1].trim();
      const unitName = row[0];

      if (playerName.length === 0 || playerName === SKIPPED_PLATOON_LABEL) {
        return undefined; // impossible platoon
      }

      // ? should incomplete platoon be considered impossible ?
      if (unitName.length === 0) {
        continue; // no unit needed here
      }

      // see if the hero is already in donations
      const unitDonated = donations.some((e) => e[0] === unitName) || result.some((e) => e[0] === unitName);

      if (!unitDonated) {
        type RequireValueInListCriteria = [
          [string], // Array of members name
          boolean, // true for dropdown
        ];

        const criteria = rules[h][0]!.getCriteriaValues() as RequireValueInListCriteria;

        // only add rare donations
        // TODO: rarity threshold
        // ! Not Available accounted for
        if (neededUnits[unitName] + 5 > criteria[0].length) {
          const sorted = criteria[0].sort(utils.caseInsensitive);
          const names: string[] = [];
          for (const name of sorted) {
            const mention = memberMentions[name];
            names.push(mention ? `${name} (${mention})` : `${name}`);
          }
          // add the recommendations
          result.push([unitName, names.join(', ')]);
        }
      }
    }

    return result;
  }

  /** Get the intro for the depth webhook */
  export function getDepthIntro(phase: TerritoryBattles.phaseIdx, mention: string): string {
    const WEBHOOK_DEPTH_ROW = 8;
    const defaultVal = `Here are the Platoon assignments for __Phase ${phase}__.
  **Do not donate heroes to the other Platoons.**`;

    return `\n\n${config.discord.webhookTemplate(phase, WEBHOOK_DEPTH_ROW, defaultVal)} ${mention}`;
  }

  /** Get the intro for the rare by webhook */
  function getRareIntro(phase: TerritoryBattles.phaseIdx, mention: string): string {
    const WEBHOOK_RARE_ROW = 7;
    const defaultVal = `Here are the Safe Platoons and the Rare Platoon donations for __Phase ${phase}__.
**Do not donate heroes to the other Platoons.**`;

    return `\n\n${config.discord.webhookTemplate(phase, WEBHOOK_RARE_ROW, defaultVal)} ${mention}`;
  }

  /** Get the intro for the warning webhook */
  export function getWarnIntro(phase: TerritoryBattles.phaseIdx, mention: string): string {
    const WEBHOOK_WARN_ROW = 6;
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
    const batch = donations.reduce(
      (acc: { count: number; fields: string[] }, e) => {
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
    if (batch.fields.length > 0) {
      postMessage(webhookURL, batch.fields.join(''));
    }
  }

  /** Send a Webhook to Discord */
  export function sendPlatoonSimplified(byType: 'Player' | 'Unit'): void {
    // get the webhook
    const webhookURL = config.discord.webhookUrl();
    if (webhookURL.length === 0) {
      // we need a url to proceed
      const UI = SpreadsheetApp.getUi();
      UI.alert('Configuration Error', 'Discord webhook not found (see Discord sheet)', UI.ButtonSet.OK);

      return;
    }

    const getZoneName = TerritoryBattles.getZoneName;
    const getPlatoonData = TerritoryBattles.getPlatoonData;
    const getPlatoonRules = TerritoryBattles.getPlatoonRules;
    const sheet = utils.getSheetByNameOrDie(SHEET.PLATOON);
    const event = config.currentEvent();
    // TODO: rework so that it no longer rely on a single `phase` value
    const phase = config.currentPhase();
    const neededUnits = TerritoryBattles.getNeededUnits(event, phase, sheet);

    // mentions only works if you get the ID
    // on your Discord server, type: \@rolename, copy the value <@#######>
    const memberMentions = getMemberMentions();
    const mentions = config.discord.roleId();

    const descriptionText = `${discord.getTitle(phase)}${getRareIntro(phase, mentions)}`;

    // get data from the platoons
    const fields: string[] = [];
    let donations: string[][] = [];
    let groundStart = -1;

    for (let zoneNum = 0; zoneNum < MAX_PLATOON_ZONES; zoneNum += 1) {
      // for each zone
      const validPlatoons: string[] = [];
      const zone = getZoneName(zoneNum, true);

      if (zoneNum === 1) {
        groundStart = donations.length;
      }

      if (discord.isTerritory(zoneNum, phase, event)) {
        // cycle throught the platoons in a zone
        for (let platoonNum = 0; platoonNum < MAX_PLATOONS; platoonNum += 1) {
          const platoon = getPlatoonDonations(
            getPlatoonData(zoneNum, platoonNum, sheet),
            donations,
            getPlatoonRules(zoneNum, platoonNum, sheet),
            memberMentions,
            neededUnits,
          );

          if (platoon) {
            validPlatoons.push(utils.EMOJI_KEYCAP_DIGITS[platoonNum + 1]);
            if (platoon.length > 0) {
              // add the new donations to the list
              for (const e of platoon) {
                donations.push([e[0], e[1]]);
              }
            }
          } else {
            validPlatoons.push('🚫');
          }
        }
      }
      if (validPlatoons.length > 0) {
        fields.push(`**${zone}**\n${validPlatoons.join('')}`);
      }
    }

    // format the high needed units
    const heroesTable = new Units.Heroes();
    const highNeedShips = heroesTable.getHighNeedList();
    if (highNeedShips.length > 0) {
      fields.push(`**High Need Ships**\n${highNeedShips.join(', ')}`);
    }
    const shipsTable = new Units.Ships();
    const highNeedHeroes = shipsTable.getHighNeedList();
    if (highNeedHeroes.length > 0) {
      fields.push(`**High Need Heroes**\n${highNeedHeroes.join(', ')}`);
    }
    postMessage(webhookURL, `${descriptionText}\n\n${fields.join('\n\n')}\n`);

    // reformat the output if we need by member istead of by unit
    if (byType === 'Player') {
      const reduced = donations.reduce((acc: Array<[string, string]>, e, i) => {
        const label = `${i >= groundStart ? 'Heroes' : 'Ships'}: `;
        const unit = e[0];
        const names = e[1].split(',');
        for (const name of names) {
          const nameTrim = name.trim();
          // see if the name is already listed
          const foundName = acc.some((member) => {
            const found = member[0] === nameTrim;
            if (found) {
              member[1] += `${member[1].indexOf(label) > -1 ? ', ' : `\n${label}`}${unit}`;
            }

            return found;
          });

          if (!foundName) {
            acc.push([nameTrim, `${label}${unit}`]);
          }
        }

        return acc;
      }, []);
      // sort by member
      reduced.sort((a, b) => utils.caseInsensitive(a[0], b[0]));
      donations = reduced;
    }

    // format the needed donations
    messageSpooler(webhookURL, byType, donations);
  }

  // ****************************************
  // Timer Functions
  // ****************************************

  /** Figure out what phase the TB is in */
  export function setCurrentPhase(): void {
    const event = config.currentEvent();
    // get the guild's TB start date/time and phase length in hours
    const startTime = config.discord.startTime();
    const phaseHours = config.discord.phaseDuration(event);
    if (startTime && phaseHours) {
      const msPerHour = 1000 * 60 * 60;
      const now = new Date();
      const diff = now.getTime() - startTime.getTime();
      const hours = diff / msPerHour + 1; // add 1 hour to ensure we are in the next phase
      const phase = Math.ceil(hours / phaseHours) as TerritoryBattles.phaseIdx;
      const maxPhases = 6;

      // set the phase in Platoons tab
      if (phase <= maxPhases) {
        config.setPhase(phase);
      }
    }
  }
}

/** Send a Webhook to Discord */
function sendPlatoonDepthWebhook(): void {
  // get the webhook
  const webhookURL = config.discord.webhookUrl();
  if (webhookURL.length === 0) {
    // we need a url to proceed
    const UI = SpreadsheetApp.getUi();
    UI.alert('Configuration Error', 'Discord webhook not found (see Discord sheet)', UI.ButtonSet.OK);

    return;
  }

  const getZoneName = TerritoryBattles.getZoneName;
  const getPlatoonData = TerritoryBattles.getPlatoonData;
  const sheet = utils.getSheetByNameOrDie(SHEET.PLATOON);
  const event = config.currentEvent();
  // TODO: rework so that it no longer rely on a single `phase` value
  const phase = config.currentPhase();

  // mentions only works if you get the id
  // in Settings - Appearance - Enable Developer Mode, type: \@rolename, copy the value <@$####>
  const descriptionText = `${discord.getTitle(phase)}${discord.getDepthIntro(phase, config.discord.roleId())}`;

  // get data from the platoons
  const fields: discord.RichEmbedOptionsField[] = [];
  for (let zoneNum = 0; zoneNum < MAX_PLATOON_ZONES; zoneNum += 1) {
    if (!discord.isTerritory(zoneNum, phase, event)) {
      continue; // skip this zone
    }

    // for each zone
    const zone = getZoneName(zoneNum, false);

    // cycle throught the platoons in a zone
    for (let platoonNum = 0; platoonNum < MAX_PLATOONS; platoonNum += 1) {
      const platoon = discord.getPlatoonString(getPlatoonData(zoneNum, platoonNum, sheet));

      if (typeof platoon === 'string' && platoon.length > 0) {
        fields.push({
          inline: true,
          name: `${zone}: ${utils.EMOJI_KEYCAP_DIGITS[platoonNum + 1]}`,
          value: platoon,
        });
      }
    }
  }

  for (let p = 0; p < fields.length; p += 6) {
    const pp = fields.slice().slice(p, p + 6);
    const options = urlFetchMakeParam_({
      content: descriptionText,
      embeds: [{ fields: pp }],
    });
    urlFetchExecute_(webhookURL, options);
  }
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
  const webhookURL = config.discord.webhookUrl(); // get the webhook
  if (webhookURL.length === 0) {
    // we need a url to proceed
    const UI = SpreadsheetApp.getUi();
    UI.alert('Configuration Error', 'Discord webhook not found (see Discord sheet)', UI.ButtonSet.OK);

    return;
  }

  const event = config.currentEvent();
  // TODO: rework so that it no longer rely on a single `phase` value
  const phase = config.currentPhase();
  const neededUnits = TerritoryBattles.getNeededUnits(event, phase);
  const fields: discord.RichEmbedOptionsField[] = [];

  fields.push({
    inline: true,
    name: 'Platoons overview',
    value: discord.getSimplifiedPlatoons(phase),
  });

  if (discord.isTerritory(0, phase, event)) {
    // get the ships list
    const shipsTable = new Units.Ships();
    const ships = shipsTable.getNeededRareList(phase, neededUnits);
    if (ships.length > 0) {
      fields.push({
        inline: true,
        name: 'Rare Ships',
        value: ships.join('\n'),
      });
    }
  }

  // get the hero list
  const heroesTable = new Units.Heroes();
  const heroes = heroesTable.getNeededRareList(phase, neededUnits);
  if (heroes.length > 0) {
    fields.push({
      inline: true,
      name: 'Rare Heroes',
      value: heroes.join('\n'),
    });
  }

  // make sure we're not trying to send empty data
  if (fields.length === 0) {
    // no data to send
    fields.push({
      inline: true,
      name: 'Rare Units',
      value: 'There Are No Rare Units For This Phase.',
    });
  }

  const title = discord.getTitle(phase);
  // mentions only works if you get the id
  // in Discord: Settings - Appearance - Enable Developer Mode
  // type: \@rolename, copy the value <@$####>
  const mentions = config.discord.roleId();
  const warnIntro = discord.getWarnIntro(phase, mentions);
  const desc = config.discord.webhookDescription(phase, event);

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
    resetPlatoonsNoUI();
  }
  allRareUnitsWebhook(); // call the webhook
  registerWebhookTimer(); // register the next timer
}

/** Try to create a webhook trigger */
function registerWebhookTimer(): void {
  const event = config.currentEvent();
  // get the guild's TB start date/time and phase length in hours
  const startTime = config.discord.startTime();
  const phaseHours = config.discord.phaseDuration(event);
  if (startTime && phaseHours) {
    const msPerHour = 1000 * 60 * 60;
    const phaseMs = phaseHours * msPerHour;
    const target = new Date(startTime);
    const now = new Date();
    const maxPhases = 6;

    // remove the trigger
    const triggers = ScriptApp.getProjectTriggers().filter((e) => e.getHandlerFunction() === sendTimedWebhook.name);
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
