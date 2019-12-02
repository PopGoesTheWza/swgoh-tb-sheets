// tslint:disable: max-classes-per-file

namespace TB {
  const SKIPPED_PLATOON_LABEL = '🚫';
  const SKIP_BUTTON_CHECKED = 'SKIP';

  /** supported TB events */
  type TBEvent = EVENT.GEONOSISDS | EVENT.GEONOSISLS | EVENT.HOTHDS | EVENT.HOTHLS;
  /** TB phases */
  type phaseIdx = 1 | 2 | 3 | 4 | 5 | 6;
  /** TB territories (zero-based) */
  type territoryIdx = 0 | 1 | 2;
  /** TB platoons/squadrons (zero-based) */
  type platoonIdx = 0 | 1 | 2 | 3 | 4 | 5;

  type slice = string[];
  type slices = [slice, slice, slice, slice, slice, slice];

  /**
   * primary class to instanciated current TB phase
   * it instantiates the relevant Territory and Platoon subclasses
   */
  class Phase {
    /** TB event */
    public readonly event: TBEvent;
    /** phase number */
    public readonly index: phaseIdx;
    /** use the Exclusions add-on spreadsheet to filter available units */
    public readonly useExclusions: boolean = true;
    /** use the Not Available list to filter available units */
    public readonly useNotAvailable: boolean = true;
    /** array of Territory objects for this phase */
    protected readonly territories: Territory[] = [];

    constructor(ev: TBEvent, index: phaseIdx) {
      this.event = ev;
      this.index = index;

      type tDef = Array<(phase: Phase, index: territoryIdx) => Territory>;
      const territories: tDef = definitions[ev][index];
      territories.forEach((e, i) => (this.territories[i] = e(this, i as territoryIdx)));
    }

    public recommend(): void {
      const spooler = new utils.Spooler();
      const territories = this.territories;

      const exclusionsId = config.exclusionId();
      const exclusions: MemberUnitBooleans = exclusionsId.length > 0 ? Exclusions.getList(this.index) : {};
      // if (this.useExclusions) {
      // }

      const notAvailable = getNotAvailableList();

      let shipsPool: ShipsPool | undefined;
      let heroesPool: HeroesPool | undefined;
      for (const territory of territories) {
        let pool: HeroesPool | ShipsPool | undefined;
        if (territory instanceof AirspaceTerritory) {
          shipsPool = shipsPool || new ShipsPool(this, exclusions, notAvailable);
          pool = shipsPool;
        }
        if (territory instanceof GroundTerritory) {
          heroesPool = heroesPool || new HeroesPool(this, exclusions, notAvailable);
          pool = heroesPool;
        }
        // get allUnits
        // init  UnitPools
        if (pool) {
          pool.addTerritory(territory);
        }
        // get neededUnits
        // platoon: clear content, data validation, set color black
        // check 'skip this'
        // create dropdown
        // mark impossible slots
        // mark impossible platoons

        // assign unit
      }

      // init neededUnit (n[unit][territory][platoon])

      // define platoonOrder(?)

      for (const territory of territories) {
        spooler.add(territory.writerName());
        territory.showHide();
      }

      spooler.commit();
    }

    public reset(): void {
      const spooler = new utils.Spooler();

      const territories = this.territories;
      for (const territory of territories) {
        territory.readSlices();
        spooler.add(territory.writerName());
        territory.writerSlices().forEach((e) => spooler.add(e));
        territory.writerResetDonors().forEach((e) => spooler.add(e));
        territory.writerResetButtons().forEach((e) => spooler.add(e));
        territory.showHide();
      }

      spooler.commit();
    }
  }

  abstract class Territory {
    public readonly index: territoryIdx;
    public platoons: Platoon[] = [];
    protected readonly sheet = utils.getSheetByNameOrDie(SHEET.PLATOON);
    protected readonly phase: Phase;
    protected readonly name: string;

    constructor(phase: Phase, index: territoryIdx, name: string) {
      this.index = index;
      this.phase = phase;
      this.name = name;
    }

    public showHide(): void {
      const sheet = this.sheet;
      const index = this.index;
      const name = this.name;
      const row = this.index * PLATOON_ZONE_ROW_OFFSET + 2;
      if (index !== 1) {
        if (name.length === 0) {
          const hideOffset = index === 0 ? 1 : 0;
          sheet.hideRows(row + hideOffset, MAX_PLATOON_UNITS - hideOffset);
        } else {
          sheet.showRows(row, MAX_PLATOON_UNITS);
        }
      }
    }

    public readSlices(): void {
      const ev = this.phase.event;
      const def = Units.getDefinitions();
      const unitsIndex = [...def.heroes, ...def.ships];

      const row = 3 + this.index * MAX_PLATOON_UNITS * MAX_PLATOONS;
      const column = (isHothDS_(ev) ? 2 : isHothLS_(ev) ? 8 : isGeoDS_(ev) ? 14 : -1) + this.phase.index;
      const data = utils
        .getSheetByNameOrDie(SHEET.STATICSLICES)
        .getRange(row, column, MAX_PLATOON_UNITS * MAX_PLATOONS, 1)
        .getValues() as string[][];
      const result: slices = [[], [], [], [], [], []];
      data.forEach((cell, i) => {
        const e = cell[0];
        const match = unitsIndex.find((u) => u.baseId === e);
        const pool = Math.floor(i / MAX_PLATOON_UNITS);
        result[pool].push(match ? match.name : e);
      });
      result.forEach((e, i) => {
        this.platoons[i].setSlice(e);
      });
    }

    public writerName(): utils.SpooledTask {
      const sheet = this.sheet;
      const name = this.name;
      const row = this.index * PLATOON_ZONE_ROW_OFFSET + 4;
      const range = sheet.getRange(row, 1);
      const writer = () => {
        range.setValue(name);
      };

      return writer;
    }

    public writerResetDonors(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map((platoon) => platoon.writerResetDonors());

      return result;
    }

    public writerResetButtons(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map((platoon) => platoon.writerResetSkipButton());

      return result;
    }

    public writerSlices(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map((platoon) => platoon.writerSlice());

      return result;
    }
  }

  class AirspaceTerritory extends Territory {
    public readonly isOpen = true;
    public readonly isGround = false;

    constructor(phase: Phase, index: territoryIdx, name: string, tp: number[]) {
      super(phase, index, name);
      this.platoons = tp.map((value, i) => new AirspacePlatoon(this, i as platoonIdx, value));
    }
  }

  class ClosedTerritory extends Territory {
    public readonly isOpen = false;
    public readonly isGround = false;

    constructor(phase: Phase, index: territoryIdx) {
      super(phase, index, '');
      for (let i = 0; i < MAX_PLATOONS; i += 1) {
        this.platoons[i] = new ClosedPlatoon(this, i as platoonIdx);
      }
    }

    public readSlices(): void {
      //
    }
  }

  class GroundTerritory extends Territory {
    public readonly isOpen = true;
    public readonly isGround = true;

    constructor(phase: Phase, index: territoryIdx, name: string, tp: number[]) {
      super(phase, index, name);
      this.platoons = tp.map((value, i) => new GroundPlatoon(this, i as platoonIdx, value));
    }
  }

  abstract class UnitPool {
    protected readonly phase: Phase;
    protected readonly allUnits: UnitMemberInstances; // units[name][member]
    protected units: UnitMemberInstances; // units[name][member]
    protected readonly exclusions: MemberUnitBooleans; // excluded[member][unit] = boolean
    protected readonly notAvailable: string[];
    protected territories: Territory[] = [];
    protected platoons: Platoon[] = [];

    constructor(phase: Phase, allUnits: UnitMemberInstances, exclusions: MemberUnitBooleans, notAvailable: string[]) {
      this.phase = phase;
      this.allUnits = utils.clone(allUnits);
      this.exclusions = utils.clone(exclusions);
      this.notAvailable = utils.clone(notAvailable);

      this.units = {};
    }

    public addTerritory(territory: Territory) {
      this.territories.push(territory);
      const platoons = territory.platoons;
      for (const platoon of platoons) {
        this.platoons.push(platoon);
      }
    }

    protected filter(filter: (member: string, u: UnitInstance) => boolean): void {
      const data = utils.clone(this.allUnits);
      const exclusions = this.exclusions;
      const units = { ...data };
      for (const unit of Object.keys(units)) {
        const members = { ...data[unit] };
        for (const member in members) {
          if ((exclusions[member] && exclusions[member][unit]) || !filter(member, members[member])) {
            delete data[unit][member];
          }
        }
        if (data[unit] && Object.keys(data[unit]).length === 0) {
          delete data[unit];
        }
      }
      this.units = data; // BAD
    }

    protected zScore() {
      const zScored: KeyedType<{
        units: Array<{
          squaredDifference: number;
          power: number;
          unit: UnitInstance;
        }>;
      }> = {};

      const units = this.units;
      for (const unitName of Object.keys(units)) {
        const perMember = units[unitName];
        for (const memberName of Object.keys(perMember)) {
          const unit = perMember[memberName];
          if (!zScored[memberName]) {
            zScored[memberName] = {
              // average: 0,
              // count: 0,
              // sum: 0,
              units: [],
            };
          }
          zScored[memberName].units.push({
            // difference: 0,
            power: unit.power,
            squaredDifference: 0,
            unit,
          });
        }
      }
      for (const memberName of Object.keys(zScored)) {
        const o = zScored[memberName];
        o.units = o.units.map((e) => e);
      }
    }
  }

  class ShipsPool extends UnitPool {
    constructor(phase: Phase, exclusions: MemberUnitBooleans = {}, notAvailable: string[] = []) {
      const shipsTable = new Units.Ships();
      const allUnits = shipsTable.getAllInstancesByUnits();
      super(phase, allUnits, exclusions, notAvailable);
    }

    protected filter() {
      const meta = utils.getSheetByNameOrDie('META');

      // filter Ships by rarity
      const filter = (member: string, u: UnitInstance) => u.rarity > +meta.getRange(5, 3).getValue();
      // && this.notAvailable.findIndex(e => e[0] === member) === -1

      super.filter(filter);
    }
  }

  class HeroesPool extends UnitPool {
    constructor(phase: Phase, exclusions: MemberUnitBooleans = {}, notAvailable: string[] = []) {
      const heroesTable = new Units.Heroes();
      const allUnits = heroesTable.getAllInstancesByUnits();
      super(phase, allUnits, exclusions, notAvailable);
    }

    protected filter() {
      const meta = utils.getSheetByNameOrDie('META');
      const alignment = config.currentAlignment();
      const rarityThreshold = +meta.getRange(5, 3).getValue();

      // filter Heroes by rarity and alignment
      const filter = (member: string, u: UnitInstance) =>
        u.rarity >= rarityThreshold && u.tags!.indexOf(alignment.toLowerCase()) > -1;
      // && this.notAvailable.findIndex(e => e[0] === member) === -1

      super.filter(filter);
    }
  }

  abstract class Platoon {
    public readonly territory: Territory;
    public readonly index: platoonIdx;
    // public readonly isGround: boolean;
    // public readonly isOpen: boolean;
    // public possible: boolean;
    public slots: Slot[] = [];
    public readonly value: number;
    protected readonly unitsRange: Spreadsheet.Range;
    protected readonly donorsRange: Spreadsheet.Range;
    protected readonly skipButtonRange: Spreadsheet.Range;
    // protected readonly sheet = SPREADSHEET.getSheetByName(SHEET.PLATOONS);
    protected slice: slice;

    constructor(territory: Territory, index: platoonIdx, value: number) {
      this.index = index;
      this.territory = territory;
      const row = PLATOON_ZONE_ROW_ORIGIN + territory.index * PLATOON_ZONE_ROW_OFFSET;
      const column = PLATOON_ZONE_COLUMN_ORIGIN + index * PLATOON_ZONE_COLUMN_OFFSET;
      const range = utils.getSheetByNameOrDie(SHEET.PLATOON).getRange(row, column, MAX_PLATOON_UNITS);
      this.unitsRange = range;
      this.donorsRange = range.offset(0, 1);
      this.skipButtonRange = range.offset(MAX_PLATOON_UNITS, 1, 1, 1);
      this.value = value;
      for (let i = 0; i <= MAX_PLATOON_UNITS; i += 1) {
        this.slots.push(new Slot(this, i));
      }

      this.slice = [];
    }

    public getDonorList(): Array<string | undefined> {
      const values = this.donorsRange.getValues() as Array<[string]>;
      const list = values.map((e) => {
        const value = e[0];
        return typeof value === 'string' && value.trim().length > 0 ? value : undefined;
      });

      return list;
    }

    public getUnitList(): Array<string | undefined> {
      const values = this.unitsRange.getValues() as Array<[string]>;
      const list = values.map((e) => {
        const value = e[0];
        return typeof value === 'string' && value.trim().length > 0 ? value : undefined;
      });

      return list;
    }

    public getSkipButton(): boolean {
      return (this.skipButtonRange.getValue() as string) === SKIP_BUTTON_CHECKED;
    }

    public setSlice(sl: slice): void {
      this.slice = sl;
    }

    public writerResetDonors(): utils.SpooledTask {
      const range = this.donorsRange;

      return () => {
        range
          .clearContent()
          .clearDataValidations()
          .setFontColor(COLOR.BLACK);
      };
    }

    public writerResetSkipButton(): utils.SpooledTask {
      const range = this.skipButtonRange;

      return () => {
        range.clearContent();
      };
    }

    public writerSlice(): utils.SpooledTask {
      const data = this.slice.map((e) => [e]);
      const range = this.unitsRange;

      return () => {
        range.setValues(data).setFontColor(COLOR.BLACK);
      };
    }
  }

  class ClosedPlatoon extends Platoon {
    constructor(territory: Territory, index: platoonIdx) {
      super(territory, index, 0);
    }

    public writerSlice(): utils.SpooledTask {
      const range = this.unitsRange;

      return () => {
        range.clearContent().setFontColor(COLOR.BLACK);
      };
    }
  }

  class AirspacePlatoon extends Platoon {
    constructor(territory: Territory, index: platoonIdx, value: number) {
      super(territory, index, value);
    }
  }

  class GroundPlatoon extends Platoon {
    constructor(territory: Territory, index: platoonIdx, value: number) {
      super(territory, index, value);
    }
  }

  class Slot {
    public unit: string;
    public member: string;
    // protected readonly unitRange: Spreadsheet.Range;
    // protected readonly donorRange: Spreadsheet.Range;
    protected readonly platoon: Platoon;
    protected readonly index: number;

    constructor(platoon: Platoon, index: number) {
      this.index = index;
      this.platoon = platoon;

      // const row = PLATOON_ZONE_ROW_ORIGIN + territory.index * PLATOON_ZONE_ROW_OFFSET;
      // const column = PLATOON_ZONE_COLUMN_ORIGIN + index * PLATOON_ZONE_COLUMN_OFFSET;

      // const range = SPREADSHEET.getSheetByName(SHEET.PLATOONS)
      //   .getRange(row, column);
      // this.unitRange = range;
      // this.donorRange = range.offset(0, 1);

      this.unit = '';
      this.member = '';
    }
  }

  type TerritoryConstructor = (phase: Phase, index: territoryIdx) => Territory;
  const closed = (phase: Phase, index: territoryIdx) => new ClosedTerritory(phase, index);
  const airspace = (n: string, tp: number[]) => (phase: Phase, index: territoryIdx) =>
    new AirspaceTerritory(phase, index, n, tp);
  const ground = (n: string, tp: number[]) => (phase: Phase, index: territoryIdx) =>
    new GroundTerritory(phase, index, n, tp);

  interface Definitions {
    [key: string]: {
      [key: number]: [TerritoryConstructor, TerritoryConstructor, TerritoryConstructor];
    };
  }

  const definitions: Definitions = {
    /** EVENT.GEONOSISDS */ 'Geo DS': {
      1: [
        closed,
        ground('Droid Factory', [166.7, 166.7, 166.7, 166.7, 166.7, 166.7]),
        ground('Canyons', [166.7, 166.7, 166.7, 166.7, 166.7, 166.7]),
      ],
      2: [
        airspace('Core Ship Yards', [166.7, 166.7, 166.7, 166.7, 166.7, 166.7]),
        ground('Separatist Command', [166.7, 166.7, 166.7, 166.7, 166.7, 166.7]),
        ground('Patranaki Arena', [166.7, 166.7, 166.7, 166.7, 166.7, 166.7]),
      ],
      3: [
        airspace('Contested Airspace', [250, 250, 250, 250, 250, 250]),
        ground('Battleground', [208.3, 208.3, 208.3, 208.3, 208.3, 208.3]),
        ground('Sand Dunes', [208.3, 208.3, 208.3, 208.3, 208.3, 208.3]),
      ],
      4: [
        airspace('Republic Fleet', [333.3, 333.3, 333.3, 333.3, 333.3, 333.3]),
        ground('Count Dookus Hangar', [250, 250, 250, 250, 250, 250]),
        ground('Rear Flank', [250, 250, 250, 250, 250, 250]),
      ],
    },
        /** EVENT.GEONOSISLS */ 'Geo LS': {
      1: [
        airspace('Galactic Republic Fleet', [208.3, 208.3, 208.3, 208.3, 208.3, 208.3]),
        ground('Count Dooku Hangar', [208.3, 208.3, 208.3, 208.3, 208.3, 208.3]),
        ground('Rear Flank', [208.3, 208.3, 208.3, 208.3, 208.3, 208.3]),
      ],
      2: [
        airspace('Contested Airspace (Republic)', [208.3, 208.3, 208.3, 208.3, 208.3, 208.3]),
        ground('Battleground', [208.3, 208.3, 208.3, 208.3, 208.3, 208.3]),
        ground('Sand Dunes', [208.3, 208.3, 208.3, 208.3, 208.3, 208.3]),
      ],
      3: [
        airspace('Contested Airspace (Separatist)', [250, 250, 250, 250, 250, 250]),
        ground('Separatist Command', [250, 250, 250, 250, 250, 250]),
        ground('Patranaki Arena', [250, 250, 250, 250, 250, 250]),
      ],
      4: [
        airspace('Separatist Armada', [333.3, 333.3, 333.3, 333.3, 333.3, 333.3]),
        ground('Factory Waste', [333.3, 333.3, 333.3, 333.3, 333.3, 333.3]),
        ground('Canyons', [333.3, 333.3, 333.3, 333.3, 333.3, 333.3]),
      ]
    },
    /** EVENT.HOTHDS */ 'Hoth DS': {
      1: [
        closed,
        ground('Imperial Flank', [102, 102, 102, 102, 153, 153]),
        ground('Imperial Landing', [102, 102, 102, 102, 153, 153]),
      ],
      2: [
        closed,
        ground('Snowfields', [126, 126, 126, 126, 189, 189]),
        ground('Forward Stronghold', [126, 126, 126, 126, 189, 189]),
      ],
      3: [
        airspace('Imperial Fleet Staging Area', [151, 151, 151, 151, 151, 151]),
        ground('Ion Cannon', [151, 151, 151, 151, 151, 151]),
        ground('Outer Pass', [151, 151, 151, 151, 151, 151]),
      ],
      4: [
        airspace('Contested Airspace', [176, 176, 176, 176, 264, 264]),
        ground('Power Generator', [176, 176, 176, 176, 176, 176]),
        ground('Rear Trenches', [176, 176, 176, 176, 176, 176]),
      ],
      5: [
        airspace('Forward Airspace', [207, 207, 207, 207, 301.5, 301.5]),
        ground('Forward Trenches', [207, 207, 207, 207, 207, 207]),
        ground('Overlook', [207, 207, 207, 207, 207, 207]),
      ],
      6: [
        airspace('Rear Airspace', [260, 260, 260, 260, 390, 390]),
        ground('Rebel Base Main Entrance', [260, 260, 260, 260, 260, 260]),
        ground('Rebel Base South Entrance', [260, 260, 260, 260, 260, 260]),
      ],
    },
    /** EVENT.HOTHLS */ 'Hoth LS': {
      1: [closed, ground('Rebel Base', [100, 100, 100, 100, 150, 150]), closed],
      2: [
        closed,
        ground('Ion Cannon', [120, 120, 120, 120, 180, 180]),
        ground('Overlook', [120, 120, 120, 120, 180, 180]),
      ],
      3: [
        airspace('Rear Airspace', [140, 140, 140, 140, 140, 140]),
        ground('Rear Trenches', [140, 140, 140, 140, 210, 210]),
        ground('Power Generator', [140, 140, 140, 140, 210, 210]),
      ],
      4: [
        airspace('Forward Airspace', [160, 160, 160, 160, 160, 160]),
        ground('Forward Trenches', [160, 160, 160, 160, 240, 240]),
        ground('Outer Pass', [160, 160, 160, 160, 240, 240]),
      ],
      5: [
        airspace('Contested Airspace', [180, 180, 180, 180, 180, 180]),
        ground('Snowfields', [180, 180, 180, 180, 270, 270]),
        ground('Forward Stronghold', [180, 180, 180, 180, 270, 270]),
      ],
      6: [
        airspace('Imperial Fleet Staging Area', [200, 200, 200, 200, 200, 200]),
        ground('Imperial Flank', [200, 200, 200, 200, 300, 300]),
        ground('Imperial Landing', [200, 200, 200, 200, 300, 300]),
      ],
    },
  };

  function getNotAvailableList(sheet = utils.getSheetByNameOrDie(SHEET.PLATOON)): string[] {
    const PLATOON_NOTAVAILABLE_ROW = 56;
    const PLATOON_NOTAVAILABLE_COL = 4;
    const PLATOON_NOTAVAILABLE_NUMROWS = MAX_MEMBERS;
    const PLATOON_NOTAVAILABLE_NUMCOLS = 1;
    const notAvailable = sheet
      .getRange(
        PLATOON_NOTAVAILABLE_ROW,
        PLATOON_NOTAVAILABLE_COL,
        PLATOON_NOTAVAILABLE_NUMROWS,
        PLATOON_NOTAVAILABLE_NUMCOLS,
      )
      .getValues()
      .reduce((acc: string[], e: string[]) => {
        const name = `${e}`.trim();
        if (name.length > 0 && acc.indexOf(name) === -1) {
          acc.push(name);
        }
        return acc;
      }, []);
    //
    return notAvailable;
  }
}
