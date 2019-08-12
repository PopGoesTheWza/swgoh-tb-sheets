// tslint:disable: max-classes-per-file

interface UnitsDefinitions {
  heroes: UnitDefinition[];
  ships: UnitDefinition[];
}

/** forces reloading units definitions and guild roster from selected data source */
function reloadUnitDefinitions() {
  Units.getDefinitionsFromDataSource();
  setupEvent();
}

/** units related classes and functions */
namespace Units {
  export enum TYPES {
    HERO = 1,
    SHIP = 2,
  }

  const sortUnits = (a: UnitDefinition, b: UnitDefinition) => {
    return utils.caseInsensitive(a.name, b.name);
  };

  /** request units definitions from data source (and cache them for 6 hours) */
  export function getDefinitionsFromDataSource(): UnitsDefinitions {
    let definitions: UnitsDefinitions;
    const cacheId = 'cachedUnits';
    const seconds = 21600; // 6 hours (maximum value)

    definitions = config.dataSource.isSwgohHelp()
      ? SwgohHelp.getUnitList()
      : { heroes: SwgohGg.getHeroList(), ships: SwgohGg.getShipList() };

    config.dataSource.setUnitDefinitionsDate();
    definitions.heroes.sort(sortUnits);
    definitions.ships.sort(sortUnits);

    const heroesTable = new Heroes();
    const shipsTable = new Ships();
    heroesTable.setDefinitions(definitions.heroes);
    shipsTable.setDefinitions(definitions.ships);

    CacheService.getScriptCache().put(cacheId, JSON.stringify(definitions), seconds);

    return definitions;
  }

  /** retrieve units definition (name, baseId, tags) from cache */
  function getDefinitionsFromCache(): UnitsDefinitions | undefined {
    const cacheId = 'cachedUnits';
    const cachedUnits = CacheService.getScriptCache().get(cacheId);

    if (cachedUnits) {
      return JSON.parse(cachedUnits) as UnitsDefinitions;
    }

    return undefined;
  }

  /** return an array of all units definition (name, baseId, tags) */
  function getDefinitionsFromSheet(sheetName: string): UnitDefinition[] {
    const sheet = SPREADSHEET.getSheetByName(sheetName);
    const UNITS_DEFINITIONS_ROW = 2;
    const UNITS_DEFINITIONS_COL = 1;
    const UNITS_DEFINITIONS_NUMROWS = sheet.getMaxRows() - 1;
    const UNITS_DEFINITIONS_NUMCOLS = 3;
    const data = sheet
      .getRange(UNITS_DEFINITIONS_ROW, UNITS_DEFINITIONS_COL, UNITS_DEFINITIONS_NUMROWS, UNITS_DEFINITIONS_NUMCOLS)
      .getValues() as string[][];

    const definitions = data.reduce((acc: UnitDefinition[], e) => {
      const name = e[0];
      const baseId = e[1];
      const tags = e[2];
      if (typeof name === 'string' && typeof baseId === 'string' && name.length > 0 && baseId.length > 0) {
        acc.push({ name, baseId, tags });
      }
      return acc;
    }, []);

    return definitions;
  }

  /** return an array of all units definition (name, baseId, tags) */
  export function getDefinitions(): UnitsDefinitions {
    const definitions = getDefinitionsFromCache() || {
      heroes: getDefinitionsFromSheet(SHEETS.HEROES),
      ships: getDefinitionsFromSheet(SHEETS.SHIPS),
    };

    return definitions;
  }

  /** abstract class for unit sheets (Heroes & Ships) */
  abstract class UnitsTable {
    /** zone: 0, 1 or 2 */
    public static getUniquePlatoonUnits(zone: number): string[] {
      const platoonRow = zone * 18 + 2;
      const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOON);
      const PLATOON_UNITS_ROW = platoonRow;
      const PLATOON_UNITS_NUMROWS = MAX_PLATOON_UNITS;
      const PLATOON_UNITS_NUMCOLS = 1;

      let units: string[][] = [];
      for (let platoon = 0; platoon < MAX_PLATOONS; platoon += 1) {
        const PLATOON_UNITS_COL = platoon * 4 + 4;
        const range = sheet.getRange(
          PLATOON_UNITS_ROW,
          PLATOON_UNITS_COL,
          PLATOON_UNITS_NUMROWS,
          PLATOON_UNITS_NUMCOLS,
        );
        const values = range.getValues() as string[][];
        units = [...units, ...values];
      }

      // flatten the array and keep only unique values
      return units.map((e) => e[0]).unique();
    }
    /** name of the sheet on which the units are stored */
    protected sheet: Spreadsheet.Sheet;
    /** column which holds units for the first member */
    private columnOffset: number;

    constructor(offset: number, sheet: Spreadsheet.Sheet) {
      this.columnOffset = offset;
      this.sheet = sheet;
    }

    /** return the number of units defined */
    public abstract getCount(): number;

    /** return a list of units that are required a high number of times (HIGHLY_NEEDED) */
    public getHighNeedList(): string[] {
      const UNITS_DATA_ROW = 2;
      const UNITS_DATA_COL = 1;
      const UNITS_DATA_NUMROWS = this.getCount();
      const UNITS_DATA_NUMCOLS = this.columnOffset - 1;
      const data = this.sheet
        .getRange(UNITS_DATA_ROW, UNITS_DATA_COL, UNITS_DATA_NUMROWS, UNITS_DATA_NUMCOLS)
        .getValues() as Array<[string, number]>;

      const idx = this.columnOffset - 1;
      const list: string[] = data.reduce((acc: string[], row) => {
        const needed = row[idx];
        if (needed >= HIGHLY_NEEDED) {
          const name = row[0];
          acc.push(`${name} (${needed})`);
        }

        return acc;
      }, []);

      return list;
    }

    /** get the unit instances for all members */
    public getAllInstancesByMember(): MemberUnitInstances {
      const units: MemberUnitInstances = {};
      const rows = this.getCount() + 1;
      const cols = this.columnOffset + config.memberCount() - 1;
      const data = this.sheet.getRange(1, 1, rows, cols).getValues() as string[][];
      const members = data.shift()!.slice(this.columnOffset - 1);

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

    /**
     * get all the unit instances for all members
     * keyed objects units[name][member] = UnitInstances
     */
    public getAllInstancesByUnits(): UnitMemberInstances {
      const units: UnitMemberInstances = {};
      const rows = this.getCount() + 1;
      const cols = this.columnOffset + config.memberCount() - 1;
      const data = this.sheet.getRange(1, 1, rows, cols).getValues() as string[][];
      const members = data.shift()!.slice(this.columnOffset - 1);

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
    public getMemberInstances(member: string): UnitInstances {
      const units: UnitInstances = {};
      const rows = this.getCount() + 1;
      const cols = this.columnOffset + config.memberCount() - 1;
      const data = this.sheet.getRange(1, 1, rows, cols).getValues() as string[][];
      const members = data.shift()!.slice(this.columnOffset - 1);
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

    /** return the UnitInstance object for the cell value */
    protected abstract toUnitInstance(stats: string): UnitInstance | undefined;

    /** return a list of Rare units for a phase */
    protected getRareList(phase: TerritoryBattles.phaseIdx): string[] {
      const data = this.sheet.getRange(2, 1, this.getCount() + 1, this.columnOffset - 2).getValues() as Array<
        [string, number]
      >;

      const column = discord.requiredRarity(1, phase) + 2; // HEROES/SHIPS, column D to I

      // cycle through each unit
      const units: string[] = data
        .reduce((acc: string[], row) => {
          const name = row[0];
          const available = row[column];
          // TODO: define RARE
          if (name.length > 0 && available < RARE_MAX) {
            acc.push(name);
          }
          return acc;
        }, []) // keep only the unit's name
        .sort(); // sort the list of units

      return units;
    }

    /** return a list of Rare units needed for a phase */
    protected getNeededRareList(phase: TerritoryBattles.phaseIdx, platoonUnits: string[]): string[] {
      const units = this.getRareList(phase);

      // filter out rare units that do not appear in platoons
      const list = units.filter((u) => platoonUnits.some((e) => e === u));

      return list;
    }

    /** set the unit instances for all members */
    protected setInstances(
      members: PlayerData[],
      definitions: UnitDefinition[],
      toString: (u: UnitInstance) => string,
    ): void {
      // Build a Member index by Name
      const memberList = Members.getNames();

      // This will hold all our data
      const data = definitions.map((e) => Array(memberList.length).fill(null));

      const headers: string[] = [];
      const nameIndex: KeyedNumbers = {};
      definitions.forEach((e, i) => (nameIndex[e.baseId] = i));

      const memberIndex: KeyedNumbers = {};
      memberList.forEach((e, i) => (memberIndex[e[0]] = i));

      for (const m of members) {
        headers.push(m.name);
        const units = m.units;
        for (const e of definitions) {
          const baseId = e.baseId;
          const u = units[baseId];
          data[nameIndex[baseId]][memberIndex[m.name]] = toString(u);
        }
      }

      // Clear out our old data, if any, including names as order may have changed
      this.sheet.getRange(1, this.columnOffset, definitions.length, MAX_MEMBERS).clearContent();

      // Write our data
      this.sheet
        .getRange(1, this.columnOffset, definitions.length + 1, memberList.length)
        .setValues([...[headers], ...data]);
    }

    /** set the unit definitions and phase count formulas */
    protected setDefinitions(units: UnitDefinition[], formula: (u: number) => string): void {
      // update the sheet
      const result = units.map((e, i) => {
        const cells = [e.name, e.baseId, e.tags];

        // insert the star count formulas
        const row = i + 2;
        const rangeText = `$K${row}:$BH${row}`;

        {
          [2, 3, 4, 5, 6, 7].forEach((stars) => {
            // tslint:disable-next-line:max-line-length
            cells.push(`=COUNT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(${rangeText},"([${stars}-7]+)\\*")))))`);
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

      this.sheet.getRange(1, 1, result.length + 1, headers.length).setValues([...[headers], ...result]);
      return;
    }
  }

  /** class to interact with Heroes sheet */
  export class Heroes extends UnitsTable {
    constructor() {
      const HERO_MEMBER_COL_OFFSET = 11;
      super(HERO_MEMBER_COL_OFFSET, SPREADSHEET.getSheetByName(SHEETS.HEROES));
    }

    /** return the number of heroes defined */
    public getCount(): number {
      const META_HEROES_COUNT_ROW = 5;
      const META_HEROES_COUNT_COL = 5;
      return +SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(META_HEROES_COUNT_ROW, META_HEROES_COUNT_COL)
        .getValue();
    }

    /** return a list of Rare heroes needed for a phase */
    public getNeededRareList(phase: TerritoryBattles.phaseIdx): string[] {
      const hgupu = Heroes.getUniquePlatoonUnits;
      // if Hoth LS phase 1: only Middle territory
      const platoonUnits: string[] = isHothLS_() && phase === 1 ? hgupu(1) : [...hgupu(1), ...hgupu(2)];

      return super.getNeededRareList(phase, platoonUnits);
    }

    /** set the hero instances for all members */
    public setInstances(members: PlayerData[]): void {
      const toString = (u: UnitInstance) => (u && `${u.rarity}*L${u.level}G${u.gearLevel}P${u.power}`) || '';

      const definitions = Units.getDefinitions().heroes;

      const baseIDs = this.sheet.getRange(2, 2, this.getCount(), 1).getValues() as string[][];
      const missingUnit = definitions.some((e) => baseIDs.findIndex((b) => b[0] === e.baseId) === -1);

      if (missingUnit) {
        this.setDefinitions(definitions);
      }

      super.setInstances(members, definitions, toString);
    }

    /** set the hero definitions and phase count formulas */
    public setDefinitions(units: UnitDefinition[]): void {
      // tslint:disable-next-line:max-line-length
      const formula = (row: number) =>
        `=COUNTIF({${SHEETS.PLATOON}!$D$20:$D$34,${SHEETS.PLATOON}!$H$20:$H$34,${SHEETS.PLATOON}!$L$20:$L$34,${SHEETS.PLATOON}!$P$20:$P$34,${SHEETS.PLATOON}!$T$20:$T$34,${SHEETS.PLATOON}!$X$20:$X$34,${SHEETS.PLATOON}!$D$38:$D$52,${SHEETS.PLATOON}!$H$38:$H$52,${SHEETS.PLATOON}!$L$38:$L$52,${SHEETS.PLATOON}!$P$38:$P$52,${SHEETS.PLATOON}!$T$38:$T$52,${SHEETS.PLATOON}!$X$38:$X$52},A${row})`;

      return super.setDefinitions(units, formula);
    }

    /** return the UnitInstance object for the cell value */
    protected toUnitInstance(stats: string): UnitInstance | undefined {
      const m = stats.match(/(\d+)\*L(\d+)G(\d+)P(\d+)/);
      if (m) {
        const rarity = +m[1];
        const level = +m[2];
        const gearLevel = +m[3];
        const power = +m[4];
        return { gearLevel, level, power, rarity, stats, type: Units.TYPES.HERO };
      }
    }
  }

  /** class to interact with Ships sheet */
  export class Ships extends UnitsTable {
    constructor() {
      const SHIP_MEMBER_COL_OFFSET = 11;
      super(SHIP_MEMBER_COL_OFFSET, SPREADSHEET.getSheetByName(SHEETS.SHIPS));
    }

    /** return the number of ships defined */
    public getCount(): number {
      const META_SHIPS_COUNT_ROW = 8;
      const META_SHIPS_COUNT_COL = 5;
      return +SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(META_SHIPS_COUNT_ROW, META_SHIPS_COUNT_COL)
        .getValue();
    }

    /** return a list of Rare ships needed for a phase */
    public getNeededRareList(phase: TerritoryBattles.phaseIdx): string[] {
      const platoonUnits: string[] = Ships.getUniquePlatoonUnits(0);

      return super.getNeededRareList(phase, platoonUnits);
    }

    /** set the ship instances for all members */
    public setInstances(members: PlayerData[]): void {
      const toString = (u: UnitInstance) => (u && `${u.rarity}*L${u.level}P${u.power}`) || '';

      const definitions = Units.getDefinitions().ships;

      const baseIDs = this.sheet.getRange(2, 2, this.getCount(), 1).getValues() as string[][];
      const missingUnit = definitions.some((e) => baseIDs.findIndex((b) => b[0] === e.baseId) === -1);

      if (missingUnit) {
        this.setDefinitions(definitions);
      }

      super.setInstances(members, definitions, toString);
    }

    /** set the ship definitions and phase count formulas */
    public setDefinitions(units: UnitDefinition[]): void {
      // tslint:disable-next-line:max-line-length
      const formula = (row: number) =>
        `=COUNTIF({${SHEETS.PLATOON}!$D$2:$D$16,${SHEETS.PLATOON}!$H$2:$H$16,${SHEETS.PLATOON}!$L$2:$L$16,${SHEETS.PLATOON}!$P$2:$P$16,${SHEETS.PLATOON}!$T$2:$T$16,${SHEETS.PLATOON}!$X$2:$X$16},A${row})`;

      return super.setDefinitions(units, formula);
    }

    /** return the UnitInstance object for the cell value */
    protected toUnitInstance(stats: string): UnitInstance | undefined {
      const m = stats.match(/(\d+)\*L(\d+)P(\d+)/);
      if (m) {
        const rarity = +m[1];
        const level = +m[2];
        const power = +m[3];
        return { level, power, rarity, stats, type: Units.TYPES.SHIP };
      }
    }
  }
}
