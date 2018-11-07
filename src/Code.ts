/** workaround to tslint issue of namespace scope after importingtype definitions */
declare namespace SwgohHelp {

  function getPlayerData(allyCode: number): PlayerData;

}

/**
 * @OnlyCurrentDoc
 */

/** string[] callback for case insensitive alphabetical sort */
function caseInsensitive_(a: string, b: string): number {
  return a.toLowerCase().localeCompare(b.toLowerCase());
}

/** [string, *][] callback for case insensitive alphabetical sort */
function firstElementCaseInsensitive_(a: [string, undefined], b: [string, undefined]): number {
  return a[0].toLowerCase().localeCompare(b[0].toLowerCase());
}

namespace Members {

  export function getNames(): [string][] {

    return SPREADSHEET.getSheetByName(SHEETS.ROSTER)
      .getRange(2, 2, config.memberCount(), 1)
      .getValues() as [string][];
  }

  export function getAllycodes(): [string, number][] {

    return SPREADSHEET.getSheetByName(SHEETS.ROSTER)
      .getRange(2, 2, config.memberCount(), 2)
      .getValues() as [string, number][];
  }

  export function getBaseAttributes(): [string, number, number, number, number][] {

    return SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, config.memberCount(), 5)
    .getValues() as [string, number, number, number, number][];
  }

  export function getFromSheet(): PlayerData[] {

    const roster = Members.getBaseAttributes();
    const heroesTable = new Units.Heroes();
    const shipsTable = new Units.Ships();
    const heroes = heroesTable.getAllInstancesByMember();
    const ships = shipsTable.getAllInstancesByMember();

    const members = roster.map((e) => {
      const memberName = e[0];
      const playerData: PlayerData = {
        allyCode: e[1],
        gp: e[2],
        heroesGp: e[3],
        name: memberName,
        shipsGp: e[4],
        units: {},
      };

      const h = heroes[memberName];
      if (h) {
        for (const baseId in h) {
          const u = h[baseId];
          playerData.units[baseId] = u;
          playerData.level = Math.max(playerData.level, u.level);
        }
      }

      const s = ships[memberName];
      if (s) {
        for (const baseId in s) {
          const u = s[baseId];
          playerData.units[baseId] = u;
          playerData.level = Math.max(playerData.level, u.level);
        }
      }

      return playerData;
    });

    return members;
  }

}

namespace Player {

  /** read player's data from a data source */
  export function getFromDataSource(
    allyCode: number,
    tag: string = '',
    unitsIndex: UnitDefinition[] = undefined,
  ): PlayerData {

    const playerData = config.dataSource.isSwgohHelp()
      ? SwgohHelp.getPlayerData(allyCode)
      : SwgohGg.getPlayerData(allyCode);

    if (playerData) {
      const units = playerData.units;
      const filteredUnits: UnitInstances = {};
      const filter = tag.toLowerCase();

      for (const baseId in units) {
        const u = units[baseId];
        let d = unitsIndex.find(e => e.baseId === baseId);
        if (!d) {  // baseId not found
          // refresh from data source
          const definitions = Units.getDefinitionsFromDataSource();
          // replace content of unitsIndex with definitions
          unitsIndex.splice(0, unitsIndex.length, ...definitions.heroes.concat(definitions.ships));
          // try again... once
          d = unitsIndex.find(e => e.baseId === baseId);
        }
        if (d && d.tags.indexOf(filter) > -1) {
          u.name = d.name;
          u.stats = `${u.rarity}* G${u.gearLevel} L${u.level} P${u.power}`;
          u.tags = d.tags;
          filteredUnits[baseId] = u;
        }
      }
      playerData.units = filteredUnits;

      return playerData;
    }

    return undefined;
  }

  /** read player's data from unit tabs */
  export function getFromSheet(memberName: string, tag: string): PlayerData {

    const roster = Members.getBaseAttributes();
    const p = roster.find(e => e[0] === memberName);
    if (p) {
      const playerData: PlayerData = {
        allyCode: p[1],
        gp: p[2],
        heroesGp: p[3],
        name: memberName,
        shipsGp: p[4],
        units: {},
      };
      const filter = tag.toLowerCase();

      const heroesTable = new Units.Heroes();
      const heroes = heroesTable.getMemberInstances(memberName);
      for (const baseId in heroes) {
        const u = heroes[baseId];
        if (filter.length === 0 || u.tags.indexOf(filter) > -1) {
          playerData.units[baseId] = u;
          playerData.level = Math.max(playerData.level, u.level);
        }
      }

      const shipsTable = new Units.Ships();
      const ships = shipsTable.getMemberInstances(memberName);
      for (const baseId in ships) {
        const u = ships[baseId];
        if (filter.length === 0 || u.tags.indexOf(filter) > -1) {
          playerData.units[baseId] = u;
          playerData.level = Math.max(playerData.level, u.level);
        }
      }

      return playerData;
    }

    return undefined;
  }

}

/** is alignment 'Light Side' */
function isLight_(filter: string): boolean {
  return filter === ALIGNMENT.LIGHTSIDE;
}

/** get the current event definition */
function getEventDefinition_(filter: string): [string, string][] {

  const sheet = SPREADSHEET.getSheetByName(SHEETS.META);
  const row = 2;
  const numRows = sheet.getLastRow() - row + 1;

  const col = (isLight_(filter) ? META_HEROES_COL : META_HEROES_DS_COL) + 2;
  const values = sheet.getRange(row, col, numRows).getValues() as [string][];
  const meta: [string, string][] = values.reduce(
    (acc, e) => {
      if (typeof e[0] === 'string' && e[0].trim().length > 0) {  // not empty
        acc.push(e[0]);
      }
      return acc;
    },
    [],
  )
  .unique()
  .map(e => [e, 'n/a']) as [string, string][];

  return meta;
}

namespace Snapshot {

  /** retrieve player's data from unit tabs or a data source */
  export function getData(
    sheet: Sheet,
    filter: string,
    unitsIndex: UnitDefinition[],
  ): PlayerData {

    const members = Members.getAllycodes();

    const memberName = (sheet.getRange(5, 1).getValue() as string).trim();

    // get the player's link from the Roster
    if (memberName.length > 0 && members.find(e => e[0] === memberName)) {
      return Player.getFromSheet(memberName, filter.toLowerCase());
    }

    // check if ally code
    const allyCode = (sheet.getRange(2, 1).getValue() as number);

    if (allyCode > 0) {

      // check if ally code exist in roster
      const member = members.find(e => e[1] === allyCode);
      if (member) {
        return Player.getFromSheet(member[0], filter.toLowerCase());
      }

      return Player.getFromDataSource(allyCode, filter, unitsIndex);
    }

    return undefined;
  }

  /** output to Snapshot sheet */
  export function output(
    sheet: Sheet,
    rowGp: number,
    baseData: string[][],
    rowHeroes: number,
    meta: string[][],
  ): void {

    sheet.getRange(1, 3, 50, 2).clearContent();  // clear the sheet
    sheet.getRange(rowGp, 3, baseData.length, 2).setValues(baseData);
    sheet.getRange(rowHeroes, 3, meta.length, 2).setValues(meta);
  }

}

/** create a snapshot of a player or guild member */
function playerSnapshot(): void {

  const definitions = Units.getDefinitions();
  const unitsIndex = definitions.heroes.concat(definitions.ships);

  // collect the meta data for the heroes
  const filter = config.currentEvent(); // TODO: potentially broken if TB not sync
  const meta = getEventDefinition_(filter);

  // get all hero stats
  let countFiltered = 0;
  let countTagged = 0;
  const characterTag = config.tagFilter(); // TODO: potentially broken if TB not sync
  const powerTarget = config.requiredHeroGp();
  const sheet = SPREADSHEET.getSheetByName(SHEETS.SNAPSHOT);
  const playerData = Snapshot.getData(sheet, filter, unitsIndex);
  if (playerData) {
    for (const baseId in playerData.units) {
      const u = playerData.units[baseId];
      const name = u.name;

      // does the hero meet the filtered requirements?
      if (u.rarity >= 7 && u.power >= powerTarget) {
        countFiltered += 1;
        // does the hero meet the tagged requirements?
        if (u.tags.indexOf(characterTag) !== -1) {
          // the hero was tagged with the characterTag we're looking for
          countTagged += 1;
        }
      }

      // store hero if required
      const heroListIdx = meta.findIndex(e => name === e[0]);
      if (heroListIdx >= 0) {
        meta[heroListIdx][1] = `${u.rarity}* G${u.gearLevel} L${u.level} P${u.power}`;
      }
    }

    // format output
    const baseData = [];
    baseData.push(['GP', playerData.gp]);
    baseData.push(['GP Heroes', playerData.heroesGp]);
    baseData.push(['GP Ships', playerData.shipsGp]);
    baseData.push([`${filter} 7* P${powerTarget}+`, countFiltered]);
    baseData.push([`${characterTag} 7* P${powerTarget}+`, countTagged]);

    const rowGp = 1;
    const rowHeroes = 6;
    // output the results
    Snapshot.output(sheet, rowGp, baseData, rowHeroes, meta);
  } else {
    UI.alert('ERROR: Failed to retrieve player\'s data.');
  }
}

/** Setup new menu items when the spreadsheet opens */
function onOpen(): void {

  UI.createMenu('SWGoH')
    .addItem('Refresh TB', setupEvent.name)
    .addSubMenu(
      UI
        .createMenu('Platoons')
        .addItem('Reset', resetPlatoons.name)
        .addItem('Recommend', recommendPlatoons.name)
        .addSeparator()
        .addItem('Send Warning to Discord', allRareUnitsWebhook.name)
        .addItem('Send Rare by Unit', sendPlatoonSimplifiedByUnitWebhook.name)
        .addItem('Send Rare by Player', sendPlatoonSimplifiedByPlayerWebhook.name)
        .addSeparator()
        .addItem('Send Micromanaged by Platoon', sendPlatoonDepthWebhook.name)
        .addItem('Send Micromanaged by Player', sendMicroByPlayerWebhook.name)
        .addSeparator()
        .addItem('Register Warning Timer', registerWebhookTimer.name),
    )
    .addItem('Player Snapshot', playerSnapshot.name)
    .addToUi();
}

namespace utils {

  function clone<T>(mutable: T): T {
    return JSON.parse(JSON.stringify(mutable));
  }

  class Queue<T> {

    private store: T[] = [];

    push(value: T) {
      this.store.push(value);
    }

    pop(): T | undefined {
      return this.store.shift();
    }

    isEmpty() {
      return this.store.length === 0;
    }

  }

  type SpooledTask = () => void;

  export class Spooler {

    private queue: Queue<SpooledTask>;

    constructor() {
      this.queue = new Queue();
    }

    attach(range: Range): SpooledRange {
      return new SpooledRange(this, range);
    }

    add(task: SpooledTask) {
      this.queue.push(task);
    }

    commit() {
      while (!this.queue.isEmpty()) {
        this.queue.pop()();
      }
    }

  }

  class SpooledRange {

    private readonly range: Range;
    private readonly spooler: Spooler;

    constructor(spooler: Spooler, range: Range) {
      this.range = range;
      this.spooler = spooler;
    }

    protected addTask(task: SpooledTask) {
      this.spooler.add(task);
    }

    clear(options: Object|undefined) {

      let immutable = undefined;
      if (typeof options === 'object') {
        immutable = Object.assign({}, options);
      }

      const range = this.range;
      this.addTask(() => range.clear(immutable));

      return this;
    }

    clearContent() {
      const range = this.range;
      this.addTask(() => range.clearContent());

      return this;
    }

    clearDataValidations(options: Object|undefined) {
      const range = this.range;
      this.addTask(() => range.clearDataValidations());

      return this;
    }

    clearFormat(options: Object|undefined) {
      const range = this.range;
      this.addTask(() => range.clearFormat());

      return this;
    }

    clearNote(options: Object|undefined) {
      const range = this.range;
      this.addTask(() => range.clearNote());

      return this;
    }

    offset(rowOffset: number, columnOffset: number, numRows?: number, numColumn?: number) {

      const range = this.range;
      const offset = range.offset(
        rowOffset,
        columnOffset,
        numRows || range.getNumRows(),
        numColumn || range.getNumColumns(),
      );

      return new SpooledRange(this.spooler, offset);
    }

    setBackground(color: string) {
      const range = this.range;
      this.addTask(() => range.setBackground(color));

      return this;
    }

    setBackgroundRGB(red: number, green: number, blue: number) {
      const range = this.range;
      this.addTask(() => range.setBackgroundRGB(red, green, blue));

      return this;
    }

    setBackgrounds(colors: string[][]) {
      const range = this.range;
      this.addTask(() => range.setBackgrounds(clone(colors)));

      return this;
    }

    setBorder(
      top: boolean,
      left: boolean,
      bottom: boolean,
      right: boolean,
      vertical: boolean,
      horizontal: boolean,
      color?: string,
      style?: GoogleAppsScript.Spreadsheet.BorderStyle,
    ) {
      const range = this.range;
      this.addTask(() => range.setBorder(
        top,
        left,
        bottom,
        right,
        vertical,
        horizontal,
        color,
        style,
      ));

      return this;
    }

    setDataValidation(rule: GoogleAppsScript.Spreadsheet.DataValidation) {
      const range = this.range;
      this.addTask(() => range.setDataValidation(rule));

      return this;
    }

    setDataValidations(rules: GoogleAppsScript.Spreadsheet.DataValidation[][]) {
      const range = this.range;
      this.addTask(() => range.setDataValidations(rules));

      return this;
    }

    setFontColor(color: string) {
      const range = this.range;
      this.addTask(() => range.setFontColor(color));

      return this;
    }

    setFontColors(colors: string[][]) {
      const range = this.range;
      this.addTask(() => range.setFontColors(clone(colors)));

      return this;
    }

    setFontFamilies(fontFamilies: string[][]) {
      const range = this.range;
      this.addTask(() => range.setFontFamilies(clone(fontFamilies)));

      return this;
    }

    setFontFamily(fontFamily: string) {
      const range = this.range;
      this.addTask(() => range.setFontFamily(fontFamily));

      return this;
    }

    setFontLine(fontLine: 'underline'|'line-through'|'none') {
      const range = this.range;
      this.addTask(() => range.setFontLine(fontLine));

      return this;
    }

    setFontLines(fontLines: ('underline'|'line-through'|'none')[][]) {
      const range = this.range;
      this.addTask(() => range.setFontLines(clone(fontLines)));

      return this;
    }

    setFontSize(fontSize: number) {
      const range = this.range;
      this.addTask(() => range.setFontSize(fontSize));

      return this;
    }

    setFontSizes(fontSizes: number[][]) {
      const range = this.range;
      this.addTask(() => range.setFontSizes(clone(fontSizes)));

      return this;
    }

    setFontStyle(fontStyle: 'italic'|'normal') {
      const range = this.range;
      this.addTask(() => range.setFontStyle(fontStyle));

      return this;
    }

    setFontStyles(fontStyles: ('italic'|'normal')[][]) {
      const range = this.range;
      this.addTask(() => range.setFontStyles(clone(fontStyles)));

      return this;
    }

    setFontWeight(fontWeight: 'bold'|'normal') {
      const range = this.range;
      this.addTask(() => range.setFontWeight(fontWeight));

      return this;
    }

    setFontWeights(fontWeights: ('bold'|'normal')[][]) {
      const range = this.range;
      this.addTask(() => range.setFontWeights(clone(fontWeights)));

      return this;
    }

    setFormula(formula: string) {
      const range = this.range;
      this.addTask(() => range.setFormula(formula));

      return this;
    }

    setFormulaR1C1(formula: string) {
      const range = this.range;
      this.addTask(() => range.setFormulaR1C1(formula));

      return this;
    }

    setFormulas(formulas: string[][]) {
      const range = this.range;
      this.addTask(() => range.setFormulas(clone(formulas)));

      return this;
    }

    setFormulasR1C1(formulas: string[][]) {
      const range = this.range;
      this.addTask(() => range.setFormulasR1C1(clone(formulas)));

      return this;
    }

    setHorizontalAlignment(alignment: 'left'|'center'|'right') {
      const range = this.range;
      this.addTask(() => range.setHorizontalAlignment(alignment));

      return this;
    }

    setHorizontalAlignments(alignments: ('left'|'center'|'right')[][]) {
      const range = this.range;
      this.addTask(() => range.setHorizontalAlignments(clone(alignments)));

      return this;
    }

    setNote(note: string) {
      const range = this.range;
      this.addTask(() => range.setNote(note));

      return this;
    }

    setNotes(notes: string[][]) {
      const range = this.range;
      this.addTask(() => range.setNotes(clone(notes)));

      return this;
    }

    setNumberFormat(numberFormat: string) {
      const range = this.range;
      this.addTask(() => range.setNumberFormat(numberFormat));

      return this;
    }

    setNumberFormats(numberFormats: string[][]) {
      const range = this.range;
      this.addTask(() => range.setNumberFormats(clone(numberFormats)));

      return this;
    }

    setShowHyperlink(showHyperlink: boolean) {
      const range = this.range;
      this.addTask(() => range.setShowHyperlink(showHyperlink));

      return this;
    }

    setTextDirection(textDirection: GoogleAppsScript.Spreadsheet.TextDirection) {
      const range = this.range;
      this.addTask(() => range.setTextDirection(textDirection));

      return this;
    }

    setTextDirections(textDirections: GoogleAppsScript.Spreadsheet.TextDirection[][]) {
      const range = this.range;
      this.addTask(() => range.setTextDirections(clone(textDirections)));

      return this;
    }

    setTextRotation(rotation: GoogleAppsScript.Spreadsheet.TextRotation) {
      const range = this.range;
      this.addTask(() => range.setTextRotation(rotation));

      return this;
    }

    setTextRotations(rotations: GoogleAppsScript.Spreadsheet.TextRotation[][]) {
      const range = this.range;
      this.addTask(() => range.setTextRotations(clone(rotations)));

      return this;
    }

    setValue(value: any) {
      const range = this.range;
      this.addTask(() => range.setValue(value));

      return this;
    }

    setValues(values: any[][]) {
      const range = this.range;
      this.addTask(() => range.setValues(clone(values)));

      return this;
    }

    setVerticalAlignment(alignment: 'top'|'middle'|'bottom') {
      const range = this.range;
      this.addTask(() => range.setVerticalAlignment(alignment));

      return this;
    }

    setVerticalAlignments(alignments: ('top'|'middle'|'bottom')[][]) {
      const range = this.range;
      this.addTask(() => range.setVerticalAlignments(clone(alignments)));

      return this;
    }

    setVerticalText(isVertical: boolean) {
      const range = this.range;
      this.addTask(() => range.setVerticalText(isVertical));

      return this;
    }

    setWrap(isWrapEnabled: boolean) {
      const range = this.range;
      this.addTask(() => range.setWrap(isWrapEnabled));

      return this;
    }

    setWrapStrategies(strategies: GoogleAppsScript.Spreadsheet.WrapStrategy[][]) {
      const range = this.range;
      this.addTask(() => range.setWrapStrategies(clone(strategies)));

      return this;
    }

    setWrapStrategy(strategy: GoogleAppsScript.Spreadsheet.WrapStrategy) {
      const range = this.range;
      this.addTask(() => range.setWrapStrategy(strategy));

      return this;
    }

    setWraps(isWrapEnabled: boolean[][]) {
      const range = this.range;
      this.addTask(() => range.setWraps(clone(isWrapEnabled)));

      return this;
    }

  }

}
