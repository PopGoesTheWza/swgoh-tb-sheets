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
    const sheet = utils.getSheetByNameOrDie(sheetName);
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
      heroes: getDefinitionsFromSheet(SHEET.HEROES),
      ships: getDefinitionsFromSheet(SHEET.SHIPS),
    };

    return definitions;
  }

  /** abstract class for unit sheets (Heroes & Ships) */
  abstract class UnitsTable {
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

    /** return a list of Rare units needed for a phase */
    public getNeededRareList(phase: TerritoryBattles.phaseIdx, neededUnits: KeyedNumbers): string[] {
      const data = this.sheet.getRange(2, 1, this.getCount() + 1, this.columnOffset - 2).getValues() as Array<
        [string, number]
      >;

      const column = discord.requiredRarity(1, phase) + 1; // HEROES/SHIPS, column D to I

      const result: string[] = [];

      for (const unitName of Object.keys(neededUnits)) {
        const needed = neededUnits[unitName];
        const row = data.find((e) => e[0] === unitName);
        if (row) {
          const available = +row[column];
          // TODO: rarity threshold
          // ! Not Available not accounted for
          if (needed + 5 > available) {
            result.push(unitName);
          }
        }
      }

      return result;
    }

    /** return the UnitInstance object for the cell value */
    protected abstract toUnitInstance(stats: string): UnitInstance | undefined;

    /** set the unit instances for all members */
    protected setInstances(
      members: PlayerData[],
      definitions: UnitDefinition[],
      toString: (u: UnitInstance) => string,
    ): void {
      // Build a Member index by Name
      const memberNames = Members.getNames();

      // This will hold all our data
      const data = definitions.map((e) => Array(memberNames.length).fill(null));

      const headers: string[] = [];
      const nameIndex: KeyedNumbers = {};
      definitions.forEach((e, i) => (nameIndex[e.baseId] = i));

      const memberIndex: KeyedNumbers = {};
      memberNames.forEach((e, i) => (memberIndex[e[0]] = i));

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
        .getRange(1, this.columnOffset, definitions.length + 1, memberNames.length)
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
      super(HERO_MEMBER_COL_OFFSET, utils.getSheetByNameOrDie(SHEET.HEROES));
    }

    /** return the number of heroes defined */
    public getCount(): number {
      const META_HEROES_COUNT_ROW = 5;
      const META_HEROES_COUNT_COL = 5;
      return +utils
        .getSheetByNameOrDie(SHEET.META)
        .getRange(META_HEROES_COUNT_ROW, META_HEROES_COUNT_COL)
        .getValue();
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
      const sheetName = SHEET.PLATOON;
      // tslint:disable-next-line:max-line-length
      const formula = (row: number) =>
        `=COUNTIF({${sheetName}!$D$20:$D$34,${sheetName}!$H$20:$H$34,${sheetName}!$L$20:$L$34,${sheetName}!$P$20:$P$34,${sheetName}!$T$20:$T$34,${sheetName}!$X$20:$X$34,${sheetName}!$D$38:$D$52,${sheetName}!$H$38:$H$52,${sheetName}!$L$38:$L$52,${sheetName}!$P$38:$P$52,${sheetName}!$T$38:$T$52,${sheetName}!$X$38:$X$52},A${row})`;

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
      super(SHIP_MEMBER_COL_OFFSET, utils.getSheetByNameOrDie(SHEET.SHIPS));
    }

    /** return the number of ships defined */
    public getCount(): number {
      const META_SHIPS_COUNT_ROW = 8;
      const META_SHIPS_COUNT_COL = 5;
      return +utils
        .getSheetByNameOrDie(SHEET.META)
        .getRange(META_SHIPS_COUNT_ROW, META_SHIPS_COUNT_COL)
        .getValue();
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
      const sheetName = SHEET.PLATOON;
      // tslint:disable-next-line:max-line-length
      const formula = (row: number) =>
        `=COUNTIF({${sheetName}!$D$2:$D$16,${sheetName}!$H$2:$H$16,${sheetName}!$L$2:$L$16,${sheetName}!$P$2:$P$16,${sheetName}!$T$2:$T$16,${sheetName}!$X$2:$X$16},A${row})`;

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
