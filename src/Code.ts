/**
 * @OnlyCurrentDoc
 */

/** string[] callback for case insensitive alphabetical sort */
function caseInsensitive_(a: string, b: string): number {
  return a.toLowerCase().localeCompare(b.toLowerCase());
}

/** guild member related functions */
namespace Members {

  /** get a row/cell array of members name */
  export function getNames(): [string][] {
    return SPREADSHEET.getSheetByName(SHEETS.ROSTER)
      .getRange(2, 2, config.memberCount(), 1)
      .getValues() as [string][];
  }

  /** get a row/cell array of members name and ally code */
  export function getAllycodes(): [string, number][] {
    return SPREADSHEET.getSheetByName(SHEETS.ROSTER)
      .getRange(2, 2, config.memberCount(), 2)
      .getValues() as [string, number][];
  }

  /**
   * get a row/cell array of members base attributes
   * [name, ally code, gp, heroes gp, ships gp]
   */
  export function getBaseAttributes(): [string, number, number, number, number][] {
    return SPREADSHEET.getSheetByName(SHEETS.ROSTER)
    .getRange(2, 2, config.memberCount(), 5)
    .getValues() as [string, number, number, number, number][];
  }

  /** get an array of members PlayerData object */
  export function getFromSheet(): PlayerData[] {

    const heroes = new Units.Heroes().getAllInstancesByMember();
    const ships = new Units.Ships().getAllInstancesByMember();

    const members = Members.getBaseAttributes().map((e) => {
      const memberName = e[0];
      const memberData: PlayerData = {
        allyCode: e[1],
        gp: e[2],
        heroesGp: e[3],
        name: memberName,
        shipsGp: e[4],
        units: {},
      };

      const addToMemberData = (e: KeyedType<UnitInstance>) => {
        if (e) {
          for (const baseId in e) {
            const u = e[baseId];
            memberData.units[baseId] = u;
            memberData.level = Math.max(memberData.level, u.level);
          }
        }
      };

      addToMemberData(heroes[memberName]);
      addToMemberData(ships[memberName]);

      return memberData;
    });

    return members;
  }

}

/** player related functions */
namespace Player {

  /** read player's data from a data source */
  export function getFromDataSource(
    allyCode: number,
    unitsIndex: UnitDefinition[],
    tag: string = '',
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

    const p = Members.getBaseAttributes()
      .find(e => e[0] === memberName);

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
      const addToPlayerData = (e: KeyedType<UnitInstance>) => {
        for (const baseId in e) {
          const u = e[baseId];
          if (filter.length === 0 || u.tags.indexOf(filter) > -1) {
            playerData.units[baseId] = u;
            playerData.level = Math.max(playerData.level, u.level);
          }
        }
      };

      const heroes = new Units.Heroes().getMemberInstances(memberName);
      addToPlayerData(heroes);
      const ships = new Units.Ships().getMemberInstances(memberName);
      addToPlayerData(ships);

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
  const col = (isLight_(filter) ? META_HEROES_COL : META_HEROES_DS_COL) + 2;
  const numRows = sheet.getLastRow() - row + 1;
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

/** Snapshot related functions */
namespace Snapshot {

  /** retrieve player's data from tabs if avaialble or from a data source */
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
    const allyCode = +sheet.getRange(2, 1).getValue();

    if (allyCode > 0) {

      // check if ally code exist in roster
      const member = members.find(e => e[1] === allyCode);
      if (member) {
        return Player.getFromSheet(member[0], filter.toLowerCase());
      }

      return Player.getFromDataSource(allyCode, unitsIndex, filter);
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
    SpreadsheetApp.getUi().alert('ERROR: Failed to retrieve player\'s data.');
  }
}

/** Setup new menu items when the spreadsheet opens */
function onOpen(): void {

  const UI = SpreadsheetApp.getUi();
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
        .addItem('Send Rare by Player', sendPlatoonSimplifiedByMemberWebhook.name)
        .addSeparator()
        .addItem('Send Micromanaged by Platoon', sendPlatoonDepthWebhook.name)
        .addItem('Send Micromanaged by Player', sendMicroByMemberWebhook.name)
        .addSeparator()
        .addItem('Register Warning Timer', registerWebhookTimer.name),
    )
    .addItem('Player Snapshot', playerSnapshot.name)
    .addToUi();
}

/** statistical functions */
namespace statistics {

  /** sum of values of a population */
  export function sum(population: any[], accessor = (e => +e)) {
    let count = 0;
    const sum: number = population.reduce(
      (acc: number, e) => {
        count += 1;
        return acc + accessor(e);
      },
      0,
    );

    return count > 0 ? sum : undefined;
  }

  export function mean(population: any[], accessor = (e => +e)) {
    let count = 0;
    const sum: number = population.reduce(
      (acc: number, e) => {
        count += 1;
        return acc + accessor(e);
      },
      0,
    );

    return count > 0 ? (sum / count) : undefined;
  }

  export function standardDeviation(population: any[], accessor = (e => +e)) {
    let count = 0;
    const sums: { sum: number; squares: number } = population.reduce(
      (acc: { sum: number; squares: number }, e) => {
        const value = accessor(e);
        count += 1;
        acc.sum += value;
        acc.squares += Math.pow(value, 2);
        return acc;
      },
      { sum: 0, squares: 0 },
    );
    if (count > 1) {
      return Math.sqrt((sums.squares - Math.pow(sums.sum, 2) / count) / (count - 1));
    }

    return undefined;
  }

  export function zScore(population: any[], accessor = (e => +e)) {
    let count = 0;
    const sums: { sum: number; squares: number; values: [number] } = population.reduce(
      (acc: { sum: number; squares: number; values: [number] }, e) => {
        const value = accessor(e);
        count += 1;
        acc.sum += value;
        acc.squares += Math.pow(value, 2);
        acc.values.push(value);
        return acc;
      },
      { sum: 0, squares: 0, values: [] },
    );
    if (count > 1) {
      const mean = sums.sum / count;
      const stdDev = Math.sqrt((sums.squares - Math.pow(sums.sum, 2) / count) / (count - 1));

      return sums.values.map((e: number) => {
        return (e - mean) / (stdDev / Math.sqrt(count));
      });
    }

    return undefined;
  }

}

namespace utils {

  export function clone<T>(mutable: T): T {
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

  export type SpooledTask = () => void;

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

  export class SpooledRange {

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

    clearDataValidations() {
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
      const immutable = clone(colors);
      this.addTask(() => range.setBackgrounds(immutable));

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
      const immutable = clone(colors);
      this.addTask(() => range.setFontColors(immutable));

      return this;
    }

    setFontFamilies(fontFamilies: string[][]) {
      const range = this.range;
      const immutable = clone(fontFamilies);
      this.addTask(() => range.setFontFamilies(immutable));

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
      const immutable = clone(fontLines);
      this.addTask(() => range.setFontLines(immutable));

      return this;
    }

    setFontSize(fontSize: number) {
      const range = this.range;
      this.addTask(() => range.setFontSize(fontSize));

      return this;
    }

    setFontSizes(fontSizes: number[][]) {
      const range = this.range;
      const immutable = clone(fontSizes);
      this.addTask(() => range.setFontSizes(immutable));

      return this;
    }

    setFontStyle(fontStyle: 'italic'|'normal') {
      const range = this.range;
      this.addTask(() => range.setFontStyle(fontStyle));

      return this;
    }

    setFontStyles(fontStyles: ('italic'|'normal')[][]) {
      const range = this.range;
      const immutable = clone(fontStyles);
      this.addTask(() => range.setFontStyles(immutable));

      return this;
    }

    setFontWeight(fontWeight: 'bold'|'normal') {
      const range = this.range;
      this.addTask(() => range.setFontWeight(fontWeight));

      return this;
    }

    setFontWeights(fontWeights: ('bold'|'normal')[][]) {
      const range = this.range;
      const immutable = clone(fontWeights);
      this.addTask(() => range.setFontWeights(immutable));

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
      const immutable = clone(formulas);
      this.addTask(() => range.setFormulas(immutable));

      return this;
    }

    setFormulasR1C1(formulas: string[][]) {
      const range = this.range;
      const immutable = clone(formulas);
      this.addTask(() => range.setFormulasR1C1(immutable));

      return this;
    }

    setHorizontalAlignment(alignment: 'left'|'center'|'right') {
      const range = this.range;
      this.addTask(() => range.setHorizontalAlignment(alignment));

      return this;
    }

    setHorizontalAlignments(alignments: ('left'|'center'|'right')[][]) {
      const range = this.range;
      const immutable = clone(alignments);
      this.addTask(() => range.setHorizontalAlignments(immutable));

      return this;
    }

    setNote(note: string) {
      const range = this.range;
      this.addTask(() => range.setNote(note));

      return this;
    }

    setNotes(notes: string[][]) {
      const range = this.range;
      const immutable = clone(notes);
      this.addTask(() => range.setNotes(immutable));

      return this;
    }

    setNumberFormat(numberFormat: string) {
      const range = this.range;
      this.addTask(() => range.setNumberFormat(numberFormat));

      return this;
    }

    setNumberFormats(numberFormats: string[][]) {
      const range = this.range;
      const immutable = clone(numberFormats);
      this.addTask(() => range.setNumberFormats(immutable));

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
      const immutable = clone(textDirections);
      this.addTask(() => range.setTextDirections(immutable));

      return this;
    }

    setTextRotation(rotation: GoogleAppsScript.Spreadsheet.TextRotation) {
      const range = this.range;
      this.addTask(() => range.setTextRotation(rotation));

      return this;
    }

    setTextRotations(rotations: GoogleAppsScript.Spreadsheet.TextRotation[][]) {
      const range = this.range;
      const immutable = clone(rotations);
      this.addTask(() => range.setTextRotations(immutable));

      return this;
    }

    setValue(value: any) {
      const range = this.range;
      this.addTask(() => range.setValue(value));

      return this;
    }

    setValues(values: any[][]) {
      const range = this.range;
      const immutable = clone(values);
      this.addTask(() => range.setValues(immutable));

      return this;
    }

    setVerticalAlignment(alignment: 'top'|'middle'|'bottom') {
      const range = this.range;
      this.addTask(() => range.setVerticalAlignment(alignment));

      return this;
    }

    setVerticalAlignments(alignments: ('top'|'middle'|'bottom')[][]) {
      const range = this.range;
      const immutable = clone(alignments);
      this.addTask(() => range.setVerticalAlignments(immutable));

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
      const immutable = clone(strategies);
      this.addTask(() => range.setWrapStrategies(immutable));

      return this;
    }

    setWrapStrategy(strategy: GoogleAppsScript.Spreadsheet.WrapStrategy) {
      const range = this.range;
      this.addTask(() => range.setWrapStrategy(strategy));

      return this;
    }

    setWraps(isWrapEnabled: boolean[][]) {
      const range = this.range;
      const immutable = clone(isWrapEnabled);
      this.addTask(() => range.setWraps(immutable));

      return this;
    }

  }

}
