/** abstract class for units (Heroes & Ships) */
abstract class UnitsTable {

  private columnOffset: number;
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor(offset: number, sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    this.columnOffset = offset;
    this.sheet = sheet;
  }

  abstract getCount(): number;

  getDefinitions(): UnitDeclaration[] {

    const data = this.sheet.getRange(2, 1, this.getCount(), 3)
      .getValues() as string[][];

    const definitions: UnitDeclaration[] = data.map((e) => {
      const name = e[0];
      const baseId = e[1];
      const tags = e[2];

      return { name, baseId, tags };
    });

    return definitions;
  }

  /** Get a list of units that are required a high number of times (HIGH_MIN) */
  getHighNeedList(): string[] {

    const data = this.sheet.getRange(2, 1, this.getCount(), this.columnOffset)
      .getValues() as [string, number][];

    const idx = this.columnOffset - 1;
    const list: string[] = data.reduce(
      (acc: string[], row) => {
        if (row[idx] >= HIGH_MIN) {
          acc.push(`${row[0]} (${row[idx]})`);
        }

        return acc;
      },
      [],
    );

    return list;
  }

  /** Get the list of Rare units needed for the phase */
  protected getAllUnits(phase: number): string[] {

    const data = this.sheet.getRange(1, 1, this.getCount() + 1, 8)
      .getValues()
      .slice(1) as [string, number][];  // Drop first line

    const column = phase + 2;  // HEROES/SHIPS, column D

    // cycle through each unit
    const units: string[] = data.reduce(
      (acc: [string], row) => {
        if (row[0].length > 0 && row[column] < RARE_MAX) {
          acc.push(row[0]);
        }
        return acc;
      },
      [])  // keep only the unit's name
      .sort();  // sort the list of units

    return units;
  }

  /** Get the list of Rare units needed for the phase */
  protected getRares(phase: number, platoonUnits: string[]): string[] {

    const units = this.getAllUnits(phase);

    // filter out rare units that do not appear in platoons
    const list = units.filter(u => platoonUnits.some(e => e === u));

    return list;
  }

  /** Populate the unit list with Member data */
  protected populateList(members: PlayerData[], toString: (u: UnitInstance) => string): void {

    // Build a Hero Index by BaseID
    const baseIDs = this.sheet.getRange(2, 2, this.getCount(), 1)
      .getValues() as string[][];

    const hIdx: number[] = [];
    baseIDs.forEach((e, i) => hIdx[e[0]] = i);

    // Build a Member index by Name
    const mList = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
      .getRange(2, 2, getGuildSize_(), 1)
      .getValues() as [string][];

    const pIdx = [];
    mList.forEach((e, i) => pIdx[e[0]] = i);

    const mHead = [[]];
    // mHead[0] = [];

    // This will hold all our data
    const data = baseIDs.map(e => Array(mList.length).fill(null));  // Initialize our data

    for (const m of members) {
      mHead[0].push(m.name);
      const units = m.units;
      for (const e of baseIDs) {
        const baseId = e[0];
        const u = units[baseId];
        data[hIdx[baseId]][pIdx[m.name]] = toString(u);
      }
    }

    // Clear out our old data, if any, including names as order may have changed
    this.sheet.getRange(1, this.columnOffset, baseIDs.length, MAX_PLAYERS)
      .clearContent();

    // Write our data
    this.sheet.getRange(1, this.columnOffset, 1, mList.length)
      .setValues(mHead);
    this.sheet.getRange(2, this.columnOffset, baseIDs.length, mList.length)
      .setValues(data);
  }

  /** Initialize the list of units */
  protected updateList(units: UnitDeclaration[], formula: (u: number) => string): void {

    // update the sheet
    const result = units.map((e, i) => {
      const hMap = [e.name, e.baseId, e.tags];

      // insert the star count formulas
      const row = i + 2;
      const rangeText = `$K${row}:$BH${row}`;

      {
        [2, 3, 4, 5, 6, 7].forEach((stars) => {
          const formula =
            `=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(${rangeText},"([${stars}-7]+)\\*")))))`;
          hMap.push(formula);
        });
      }

      // insert the needed count
      hMap.push(formula(row));

      return hMap;
    });
    const header = [];
    header[0] = [
      'Name',
      'Base Id',
      'Tags',
      '2',
      '3',
      '4',
      '5',
      '6',
      '7',
      '=CONCAT("# Needed P",Platoon!A2)',
    ];
    this.sheet.getRange(1, 1, 1, header[0].length)
      .setValues(header);
    this.sheet.getRange(2, 1, result.length, this.columnOffset - 1)
      .setValues(result);

    return;
  }
}

class HeroesTable extends UnitsTable {

  constructor() {
    super(HERO_PLAYER_COL_OFFSET, SPREADSHEET.getSheetByName(SHEETS.HEROES));
  }

  getCount(): number {

    const value = SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_HEROES_COUNT_ROW, META_HEROES_COUNT_COL)
      .getValue() as number;

    return value;
  }

/** Get the list of Rare units needed for the phase */
  getRares(phase: number): string[] {

    const platoonUnits: string[] = (!isLight_(getSideFilter_()) || phase > 1)
      ? getUniquePlatoonUnits_(1).concat(getUniquePlatoonUnits_(2))
      : getUniquePlatoonUnits_(1);

    return super.getRares(phase, platoonUnits);
  }

  /** Populate the Hero list with Member data */
  populateList(members: PlayerData[]): void {

    const toString = (u: UnitInstance) =>
      (u && `${u.rarity}*L${u.level}G${u.gearLevel}P${u.power}`) || '';

    super.populateList(members, toString);
  }

  /** Initialize the list of heroes */
  updateList(units: UnitDeclaration[]): void {

    const formula = row => `=COUNTIF({${SHEETS.PLATOONS}!$D$20:$D$34,${SHEETS.PLATOONS}!$H$20:$H$34,
${SHEETS.PLATOONS}!$L$20:$L$34,${SHEETS.PLATOONS}!$P$20:$P$34,
${SHEETS.PLATOONS}!$T$20:$T$34,${SHEETS.PLATOONS}!$X$20:$X$34,
${SHEETS.PLATOONS}!$D$38:$D$52,${SHEETS.PLATOONS}!$H$38:$H$52,
${SHEETS.PLATOONS}!$L$38:$L$52,${SHEETS.PLATOONS}!$P$38:$P$52,
${SHEETS.PLATOONS}!$T$38:$T$52,${SHEETS.PLATOONS}!$X$38:$X$52},A${row})`;

    return super.updateList(units, formula);
  }
}

class ShipsTable extends UnitsTable {

  constructor() {
    super(SHIP_PLAYER_COL_OFFSET, SPREADSHEET.getSheetByName(SHEETS.SHIPS));
  }

  getCount(): number {

    const value = SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_SHIPS_COUNT_ROW, META_SHIPS_COUNT_COL)
      .getValue() as number;

    return value;
  }

  /** Get the list of Rare units needed for the phase */
  getRares(phase: number): string[] {

    const platoonUnits: string[] = getUniquePlatoonUnits_(0);

    return super.getRares(phase, platoonUnits);
  }

  /** Populate the Ship list with Member data */
  populateList(members: PlayerData[]): void {

    const toString = (u: UnitInstance) => (u && `${u.rarity}*L${u.level}P${u.power}`) || '';

    super.populateList(members, toString);
  }

  /** Initialize the list of ships */
  updateList(units: UnitDeclaration[]): void {

    const formula = row => `=COUNTIF({${SHEETS.PLATOONS}!$D$2:$D$16,${SHEETS.PLATOONS}!$H$2:$H$16,
${SHEETS.PLATOONS}!$L$2:$L$16,${SHEETS.PLATOONS}!$P$2:$P$16,
${SHEETS.PLATOONS}!$T$2:$T$16,${SHEETS.PLATOONS}!$X$2:$X$16},A${row})`;

    return super.updateList(units, formula);
  }
}
