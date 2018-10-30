/** abstract class for unit sheets (Heroes & Ships) */
abstract class UnitsTable {

  /** column which holds units for the first member */
  private columnOffset: number;
  /** name of the sheet on which the units are stored */
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor(offset: number, sheet: GoogleAppsScript.Spreadsheet.Sheet) {

    this.columnOffset = offset;
    this.sheet = sheet;
  }

  /** return the number of units defined */
  abstract getCount(): number;

  /** return the UnitInstance object for the cell value */
  protected abstract toUnitInstance(stats: string): UnitInstance;

  /** return an array of all units definition (name, baseId, tags) */
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

  /** return a list of units that are required a high number of times (HIGH_MIN) */
  getHighNeedList(): string[] {

    const data = this.sheet.getRange(2, 1, this.getCount(), this.columnOffset - 1)
      .getValues() as [string, number][];

    const idx = this.columnOffset - 1;
    const list: string[] = data.reduce(
      (acc: string[], row) => {
        const needed = row[idx];
        if (needed >= HIGH_MIN) {
          const name = row[0];
          acc.push(`${name} (${needed})`);
        }

        return acc;
      },
      [],
    );

    return list;
  }

  /** return a list of Rare units for a phase */
  protected getRareList(phase: number): string[] {

    const data = this.sheet.getRange(2, 1, this.getCount() + 1, this.columnOffset - 2)
      .getValues() as [string, number][];

    const column = phase + 3;  // HEROES/SHIPS, column D to I

    // cycle through each unit
    const units: string[] = data.reduce(
      (acc: [string], row) => {
        const name = row[0];
        const available = row[column];
        if (name.length > 0 && available < RARE_MAX) {
          acc.push(name);
        }
        return acc;
      },
      [])  // keep only the unit's name
      .sort();  // sort the list of units

    return units;
  }

  /** return a list of Rare units needed for a phase */
  protected getNeededRareList(phase: number, platoonUnits: string[]): string[] {

    const units = this.getRareList(phase);

    // filter out rare units that do not appear in platoons
    const list = units.filter(u => platoonUnits.some(e => e === u));

    return list;
  }

  /** get the unit instances for all members */
  getAllInstancesByMember(): KeyedType<UnitInstances> {

    const units: KeyedType<UnitInstances> = {};
    const rows = this.getCount() + 1;
    const cols = this.columnOffset + getGuildSize_() - 1;
    const data = this.sheet.getRange(1, 1, rows, cols)
      .getValues() as string[][];
    const members = data.shift().slice(this.columnOffset - 1);

    for (const row of data) {

      const name = row[0];
      const baseId = row[1];
      const tags = row[2];
      const instances = row.slice(this.columnOffset - 1);

      instances.forEach((e, i) => {

        const u = this.toUnitInstance(e);
        if (u) {
          u.baseId = baseId;
          u.name = name;
          u.tags = tags;
          const member = members[i];
          if (!units[member]) {
            units[member] = {};
          }
          units[member][baseId] = u;
        }
      });
    }

    return units;
  }

  /** get the unit instances for all members */
  getAllInstancesByUnits(): KeyedType<UnitInstances> {

    const units: KeyedType<UnitInstances> = {};
    const rows = this.getCount() + 1;
    const cols = this.columnOffset + getGuildSize_() - 1;
    const data = this.sheet.getRange(1, 1, rows, cols)
      .getValues() as string[][];
    const members = data.shift().slice(this.columnOffset - 1);

    for (const row of data) {

      const name = row[0];
      const baseId = row[1];
      const tags = row[2];
      const instances = row.slice(this.columnOffset - 1);

      instances.forEach((e, i) => {

        const u = this.toUnitInstance(e);
        if (u) {
          u.baseId = baseId;
          u.name = name;
          u.tags = tags;
          const member = members[i];
          if (!units[name]) {
            units[name] = {};
          }
          units[name][member] = u;
        }
      });
    }

    return units;
  }

  /** get the unit instances for all members */
  getMemberInstances(member: string): UnitInstances {

    const units: UnitInstances = {};
    const rows = this.getCount() + 1;
    const cols = this.columnOffset + getGuildSize_() - 1;
    const data = this.sheet.getRange(1, 1, rows, cols)
      .getValues() as string[][];
    const members = data.shift().slice(this.columnOffset - 1);
    const column = members.indexOf(member);
    if (column !== -1) {

      for (const row of data) {

        const name = row[0];
        const baseId = row[1];
        const tags = row[2];
        const instances = row.slice(this.columnOffset - 1);
        const u = this.toUnitInstance(instances[column]);
        if (u) {
          u.baseId = baseId;
          u.name = name;
          u.tags = tags;
          units[baseId] = u;
        }
      }
    }

    return units;
  }

  /** set the unit instances for all members */
  protected setInstances(members: PlayerData[], toString: (u: UnitInstance) => string): void {

    // Build a Hero Index by BaseID
    const baseIDs = this.sheet.getRange(2, 2, this.getCount(), 1)
      .getValues() as string[][];

    // Build a Member index by Name
    const memberList = SPREADSHEET.getSheetByName(SHEETS.ROSTER)
      .getRange(2, 2, getGuildSize_(), 1)
      .getValues() as [string][];

    // This will hold all our data
    const data = baseIDs.map(e => Array(memberList.length).fill(null));  // Initialize our data

    const headers: string[] = [];
    const nameIndex: KeyedNumbers = {};
    baseIDs.forEach((e, i) => nameIndex[e[0]] = i);

    const memberIndex: KeyedNumbers = {};
    memberList.forEach((e, i) => memberIndex[e[0]] = i);

    for (const m of members) {
      headers.push(m.name);
      const units = m.units;
      for (const e of baseIDs) {
        const baseId = e[0];
        const u = units[baseId];
        data[nameIndex[baseId]][memberIndex[m.name]] = toString(u);
      }
    }

    // Clear out our old data, if any, including names as order may have changed
    this.sheet.getRange(1, this.columnOffset, baseIDs.length, MAX_PLAYERS)
      .clearContent();

    // Write our data
    this.sheet.getRange(1, this.columnOffset, baseIDs.length + 1, memberList.length)
      .setValues([headers].concat(data));
  }

  /** set the unit definitions and phase count formulas */
  protected setDefinitions(units: UnitDeclaration[], formula: (u: number) => string): void {

    // update the sheet
    const result = units.map((e, i) => {
      const cells = [e.name, e.baseId, e.tags];

      // insert the star count formulas
      const row = i + 2;
      const rangeText = `$K${row}:$BH${row}`;

      {
        [2, 3, 4, 5, 6, 7].forEach((stars) => {
          cells.push(
            `=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(${rangeText},"([${stars}-7]+)\\*")))))`,
            );
        });
      }

      // insert the needed count
      cells.push(formula(row));

      return cells;
    });
    const headers: string[] = [
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

    this.sheet.getRange(1, 1, result.length + 1, headers.length)
      .setValues([headers].concat(result));
    return;
  }
}

/** class to interact with Heroes sheet */
class HeroesTable extends UnitsTable {

  constructor() {
    super(HERO_PLAYER_COL_OFFSET, SPREADSHEET.getSheetByName(SHEETS.HEROES));
  }

  /** return the number of heroes defined */
  getCount(): number {

    const value = SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_HEROES_COUNT_ROW, META_HEROES_COUNT_COL)
      .getValue() as number;

    return value;
  }

  /** return the UnitInstance object for the cell value */
  protected toUnitInstance(stats: string): UnitInstance {
    let result: UnitInstance;
    const m = stats.match(/(\d+)\*L(\d+)G(\d+)P(\d+)/);
    if (m) {
      const rarity = Number(m[1]);
      const level = Number(m[2]);
      const gearLevel = Number(m[3]);
      const power = Number(m[4]);
      result = { gearLevel, level, power, rarity, stats };
    }

    return result;
  }

  /** return a list of Rare heroes needed for a phase */
  getNeededRareList(phase: number): string[] {

    const platoonUnits: string[] = (!isLight_(getSideFilter_()) || phase > 1)
      ? getUniquePlatoonUnits_(1).concat(getUniquePlatoonUnits_(2))
      : getUniquePlatoonUnits_(1);

    return super.getNeededRareList(phase, platoonUnits);
  }

  /** set the hero instances for all members */
  setInstances(members: PlayerData[]): void {

    const toString = (u: UnitInstance) =>
      (u && `${u.rarity}*L${u.level}G${u.gearLevel}P${u.power}`) || '';

    super.setInstances(members, toString);
  }

  /** set the hero definitions and phase count formulas */
  setDefinitions(units: UnitDeclaration[]): void {

    const formula = row => `=COUNTIF({${SHEETS.PLATOONS}!$D$20:$D$34,${SHEETS.PLATOONS}!$H$20:$H$34,
${SHEETS.PLATOONS}!$L$20:$L$34,${SHEETS.PLATOONS}!$P$20:$P$34,
${SHEETS.PLATOONS}!$T$20:$T$34,${SHEETS.PLATOONS}!$X$20:$X$34,
${SHEETS.PLATOONS}!$D$38:$D$52,${SHEETS.PLATOONS}!$H$38:$H$52,
${SHEETS.PLATOONS}!$L$38:$L$52,${SHEETS.PLATOONS}!$P$38:$P$52,
${SHEETS.PLATOONS}!$T$38:$T$52,${SHEETS.PLATOONS}!$X$38:$X$52},A${row})`;

    return super.setDefinitions(units, formula);
  }
}

/** class to interact with Ships sheet */
class ShipsTable extends UnitsTable {

  constructor() {
    super(SHIP_PLAYER_COL_OFFSET, SPREADSHEET.getSheetByName(SHEETS.SHIPS));
  }

  /** return the number of ships defined */
  getCount(): number {

    const value = SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_SHIPS_COUNT_ROW, META_SHIPS_COUNT_COL)
      .getValue() as number;

    return value;
  }

  /** return the UnitInstance object for the cell value */
  protected toUnitInstance(stats: string): UnitInstance {
    let result: UnitInstance;
    const m = stats.match(/(\d+)\*L(\d+)P(\d+)/);
    if (m) {
      const rarity = Number(m[1]);
      const level = Number(m[2]);
      const power = Number(m[3]);
      result = { level, power, rarity, stats };
    }

    return result;
  }

  /** return a list of Rare ships needed for a phase */
  getNeededRareList(phase: number): string[] {

    const platoonUnits: string[] = getUniquePlatoonUnits_(0);

    return super.getNeededRareList(phase, platoonUnits);
  }

  /** set the ship instances for all members */
  setInstances(members: PlayerData[]): void {

    const toString = (u: UnitInstance) => (u && `${u.rarity}*L${u.level}P${u.power}`) || '';

    super.setInstances(members, toString);
  }

  /** set the ship definitions and phase count formulas */
  setDefinitions(units: UnitDeclaration[]): void {

    const formula = row => `=COUNTIF({${SHEETS.PLATOONS}!$D$2:$D$16,${SHEETS.PLATOONS}!$H$2:$H$16,
${SHEETS.PLATOONS}!$L$2:$L$16,${SHEETS.PLATOONS}!$P$2:$P$16,
${SHEETS.PLATOONS}!$T$2:$T$16,${SHEETS.PLATOONS}!$X$2:$X$16},A${row})`;

    return super.setDefinitions(units, formula);
  }
}
