// tslint:disable: max-classes-per-file

let PLATOON_PHASES: Array<[string, string, string]> = [];
let PLATOON_NEEDED_COUNT: KeyedNumbers = {};

/** Custom object for creating custom order to walk through platoons */
class PlatoonDetails {
  public readonly zone: number;
  public readonly platoon: number;
  public readonly row: number;
  public possible: boolean;
  public readonly isGround: boolean;
  public readonly exist: boolean;

  constructor(phase: number, zone: number, platoon: number) {
    this.zone = zone;
    this.platoon = platoon;
    this.row = 2 + zone * PLATOON_ZONE_ROW_OFFSET;
    this.possible = true;
    this.isGround = zone > 0;
    this.exist = discord.isTerritory(zone, phase as TerritoryBattles.phaseIdx);
  }

  public getOffset() {
    return this.platoon * 4;
  }
}

/**
 * Custom object for platoon units
 * hero, count, member count, member list (member, gear...)
 */
class PlatoonUnit {
  public readonly name: string;
  public count: number;
  public members: string[];
  private readonly pCount: number;

  constructor(name: string, count: number, pCount: number) {
    this.name = name;
    this.count = count;
    this.pCount = pCount;
    this.members = [];
  }

  public isMissing(): boolean {
    return this.count > this.pCount;
  }

  // TODO: define RARE
  public isRare(): boolean {
    return this.count + 3 > this.pCount;
  }
}

/** Initialize the list of Territory names */
function initPlatoonPhases_(): void {
  const event = config.currentEvent();

  // Names of Territories, # of Platoons
  if (isHothLS_(event)) {
    PLATOON_PHASES = [
      ['', 'Rebel Base', ''],
      ['', 'Ion Cannon', 'Overlook'],
      ['Rear Airspace', 'Rear Trenches', 'Power Generator'],
      ['Forward Airspace', 'Forward Trenches', 'Outer Pass'],
      ['Contested Airspace', 'Snowfields', 'Forward Stronghold'],
      ['Imperial Fleet Staging Area', 'Imperial Flank', 'Imperial Landing'],
    ];
  } else if (isHothDS_(event)) {
    PLATOON_PHASES = [
      ['', 'Imperial Flank', 'Imperial Landing'],
      ['', 'Snowfields', 'Forward Stronghold'],
      ['Imperial Fleet Staging Area', 'Ion Cannon', 'Outer Pass'],
      ['Contested Airspace', 'Power Generator', 'Rear Trenches'],
      ['Forward Airspace', 'Forward Trenches', 'Overlook'],
      ['Rear Airspace', 'Rebel Base Main Entrance', 'Rebel Base South Entrance'],
    ];
  } else if (isGeoDS_(event)) {
    PLATOON_PHASES = [
      ['', 'Droid Factory', 'Canyons'],
      ['Core Ship Yards', 'Separatist Command', 'Petranaki Arena'],
      ['Contested Airspace', 'Battleground', 'Sand Dunes'],
      ['Republic Fleet', 'Count Dooku Hangar', 'Rear Flank'],
      ['???', '???', '???'],
      ['???', '???', '???'],
    ];
  } else {
    PLATOON_PHASES = [
      ['???', '???', '???'],
      ['???', '???', '???'],
      ['???', '???', '???'],
      ['???', '???', '???'],
      ['???', '???', '???'],
      ['???', '???', '???'],
    ];
  }
}

/** Get the zone name and update the cell */
function setZoneName_(
  spooler: utils.Spooler,
  phase: TerritoryBattles.phaseIdx,
  zone: number,
  sheet: Spreadsheet.Sheet,
  platoonRow: number,
): string {
  // set the zone name
  const zoneName = PLATOON_PHASES[phase - 1][zone];
  spooler.attach(sheet.getRange(platoonRow + 2, 1)).setValue(zoneName);

  return zoneName;
}

/** Clear the full chart */
function resetPlatoons(): void {
  const event = config.currentEvent();
  const phase = config.currentPhase();

  SPREADSHEET.toast(`Platoons for ${event} phase ${phase} reset.`, 'Reset platoons', 3);
  resetPlatoonsNoUI();
}

function resetPlatoonsNoUI(): void {
  const event = config.currentEvent();
  const phase = config.currentPhase();
  new TerritoryBattles.Phase(event, phase).reset();
}

function getNeededCount_(unitName: string) {
  PLATOON_NEEDED_COUNT[unitName] = (PLATOON_NEEDED_COUNT[unitName] || 0) + 1;
}

/** Get a sorted list of recommended members */
function getRecommendedMembers_(
  unitName: string,
  phase: TerritoryBattles.phaseIdx,
  data: UnitMemberInstances,
): Array<[string, number]> {
  // see how many stars are needed
  const minRarity = isGeoDS_() ? (phase < 3 ? 6 : 7) : phase + 1;

  const rec: Array<[string, number]> = [];

  const members = data[unitName];
  if (members) {
    for (const member of Object.keys(members)) {
      const memberUnit = members[member];

      if (memberUnit.rarity >= minRarity) {
        rec.push([member, memberUnit.power]);
      }
    }
  }

  // sort list by power
  rec.sort((a, b) => a[1] - b[1]); // sorts by 2nd element ascending

  return rec;
}

/** create the dropdown list */
function buildDropdown_(memberList: Array<[string, number]>): Spreadsheet.DataValidation {
  const formatList = memberList.map((e) => e[0]);

  return SpreadsheetApp.newDataValidation()
    .requireValueInList(formatList)
    .build();
}

/** Reset the units used */
function resetUsedUnits_(data: UnitMemberInstances): UnitMemberBooleans {
  const result: UnitMemberBooleans = {};

  for (const unit of Object.keys(data)) {
    const members = data[unit];
    for (const member of Object.keys(members)) {
      if (!result[unit]) {
        result[unit] = {};
      }
      result[unit][member] = false;
    }
  }

  return result;
}

function filterUnits_(data: UnitMemberInstances, filter: (member: string, u: UnitInstance) => boolean) {
  const units = { ...data };
  for (const unit of Object.keys(units)) {
    const members = { ...data[unit] };
    for (const member in members) {
      if (!filter(member, members[member])) {
        delete data[unit][member];
      }
    }
    if (data[unit] && Object.keys(data[unit]).length === 0) {
      delete data[unit];
    }
  }
}

/** Territory Battles related classes and functions */
namespace TerritoryBattles {
  /** supported events for TB */
  export type event = EVENT.GEONOSISDS | EVENT.GEONOSISLS | EVENT.HOTHDS | EVENT.HOTHLS;
  /** TB phases */
  export type phaseIdx = 1 | 2 | 3 | 4 | 5 | 6;
  /** TB territories (zero-based) */
  export type territoryIdx = 0 | 1 | 2;
  /** TB platoons/squadrons (zero-based) */
  type platoonIdx = 0 | 1 | 2 | 3 | 4 | 5;

  /**
   * primary class to instanciated current TB phase
   * it instantiates the relevant Territory and Platoon subclasses
   */
  export class Phase {
    /** event type (LS/DS) */
    public readonly event: event;
    /** phase number */
    public readonly index: phaseIdx;
    /** use the Exclusions add-on spreadsheet to filter available units */
    public readonly useExclusions: boolean = true;
    /** use the Not Available list to filter available units */
    public readonly useNotAvailable: boolean = true;
    /** array of Territory objects for this phase */
    protected readonly territories: Territory[] = [];

    constructor(ev: event, index: phaseIdx) {
      this.event = ev;
      this.index = index;

      type tDef = Array<(p: Phase, i: territoryIdx) => Territory>;
      const territories: tDef = definitions[ev][index];
      territories.forEach((e, i) => (this.territories[i] = e(this, i as territoryIdx)));
    }

    public recommend(): void {
      const spooler = new utils.Spooler();
      const territories = this.territories;

      let exclusions: MemberUnitBooleans;
      const exclusionsId = config.exclusionId();
      if (exclusionsId.length > 0) {
        exclusions = Exclusions.getList(this.index);
      }
      // if (this.useExclusions) {
      // }

      const notAvailable: string[] = [];
      const data = SPREADSHEET.getSheetByName(SHEETS.PLATOON)
        .getRange(56, 4, MAX_MEMBERS)
        .getValues() as Array<[string]>;
      for (const e of data) {
        const name = e[0];
        if (name.length > 0) {
          notAvailable.push(name);
        }
      }
      // if (this.useNotAvailable) {
      // }

      let shipsPool: ShipsPool;
      let heroesPool: HeroesPool;
      for (const territory of territories) {
        let pool: HeroesPool | ShipsPool;
        if (territory instanceof AirspaceTerritory) {
          shipsPool = shipsPool! || new ShipsPool(this, exclusions!, notAvailable);
          pool = shipsPool;
        }
        if (territory instanceof GroundTerritory) {
          heroesPool = heroesPool! || new HeroesPool(this, exclusions!, notAvailable);
          pool = heroesPool;
        }
        // get allUnits
        // init  UnitPools
        if (pool!) {
          pool!.addTerritory(territory);
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

  type slice = string[];
  type slices = [slice, slice, slice, slice, slice, slice];

  abstract class Territory {
    public readonly index: territoryIdx;
    public platoons: Platoon[] = [];
    protected readonly sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOON);
    protected readonly phase: Phase;
    protected readonly name: string;

    constructor(phase: Phase, index: territoryIdx, name: string) {
      this.index = index;
      this.phase = phase;
      this.name = name;
    }

    public showHide(): void {
      const index = this.index;
      const name = this.name;
      const row = this.index * PLATOON_ZONE_ROW_OFFSET + 2;
      const sheet = this.sheet;
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
      const data = SPREADSHEET.getSheetByName(SHEETS.STATICSLICES)
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
      const name = this.name;
      const row = this.index * PLATOON_ZONE_ROW_OFFSET + 4;
      const range = this.sheet.getRange(row, 1);
      const writer = () => {
        range.setValue(name);
      };

      return writer;
    }

    public writerResetDonors(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map((p) => p.writerResetDonors());

      return result;
    }

    public writerResetButtons(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map((p) => p.writerResetSkipButton());

      return result;
    }

    public writerSlices(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map((p) => p.writerSlice());

      return result;
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

  class AirspaceTerritory extends Territory {
    public readonly isOpen = true;
    public readonly isGround = false;

    constructor(phase: Phase, index: territoryIdx, name: string, tp: number[]) {
      super(phase, index, name);
      this.platoons = tp.map((value, i) => new AirspacePlatoon(this, i as platoonIdx, value));
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
    // protected readonly sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
    protected slice: slice;

    constructor(territory: Territory, index: platoonIdx, value: number) {
      this.index = index;
      this.territory = territory;
      const row = territory.index * PLATOON_ZONE_ROW_OFFSET + 2;
      const column = index * PLATOON_ZONE_COLUMN_OFFSET + 4;
      const range = SPREADSHEET.getSheetByName(SHEETS.PLATOON).getRange(row, column, MAX_PLATOON_UNITS);
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
      return (this.skipButtonRange.getValue() as string) === 'SKIP';
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
      // const row = platoon.territory.index * PLATOON_ZONE_ROW_OFFSET + index + 2;
      // const column = platoon.index * PLATOON_ZONE_COLUMN_OFFSET + 4;
      // const range = SPREADSHEET.getSheetByName(SHEETS.PLATOONS)
      //   .getRange(row, column);
      // this.unitRange = range;
      // this.donorRange = range.offset(0, 1);

      this.unit = '';
      this.member = '';
    }
  }

  type TerritoryConstructor = (p: Phase, i: territoryIdx) => Territory;
  const closed = (p: Phase, i: territoryIdx) => new ClosedTerritory(p, i);
  const airspace = (n: string, tp: number[]) => (p: Phase, i: territoryIdx) => new AirspaceTerritory(p, i, n, tp);
  const ground = (n: string, tp: number[]) => (p: Phase, i: territoryIdx) => new GroundTerritory(p, i, n, tp);

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

  class HeroesPool extends UnitPool {
    constructor(phase: Phase, exclusions: MemberUnitBooleans = {}, notAvailable: string[] = []) {
      const heroesTable = new Units.Heroes();
      const allUnits = heroesTable.getAllInstancesByUnits();
      super(phase, allUnits, exclusions, notAvailable);
    }

    protected filter() {
      const alignment = config.currentAlignment().toLowerCase();
      const rarityThreshold = isGeoDS_() ? (this.phase.index < 3 ? 6 : 7) : this.phase.index + 1;

      // filter Heroes by rarity and alignment
      const filter = (member: string, u: UnitInstance) =>
        u.rarity >= rarityThreshold && u.tags!.indexOf(alignment) !== -1;
      // && this.notAvailable.findIndex(e => e[0] === member) === -1

      super.filter(filter);
    }
  }

  class ShipsPool extends UnitPool {
    constructor(phase: Phase, exclusions: MemberUnitBooleans = {}, notAvailable: string[] = []) {
      const shipsTable = new Units.Ships();
      const allUnits = shipsTable.getAllInstancesByUnits();
      super(phase, allUnits, exclusions, notAvailable);
    }

    protected filter() {
      const phase = this.phase.index;

      // filter Ships by rarity
      const filter = (member: string, u: UnitInstance) => u.rarity > phase;
      // && this.notAvailable.findIndex(e => e[0] === member) === -1

      super.filter(filter);
    }
  }
}

function loop1_(
  spooler: utils.Spooler,
  cur: PlatoonDetails,
  platoonMatrix: PlatoonUnit[],
  sheet: Spreadsheet.Sheet,
  phase: TerritoryBattles.phaseIdx,
  allUnits: UnitMemberInstances,
) {
  const baseCol = 4; // TODO: should it be a setting

  const row = cur.row;
  const column = cur.getOffset() + baseCol;
  const range = sheet.getRange(row, column, MAX_PLATOON_UNITS);

  // clear previous contents
  spooler
    .attach(range.offset(0, 1))
    .clearContent()
    .clearDataValidations()
    .offset(0, -1, MAX_PLATOON_UNITS, 2)
    .setFontColor(COLOR.BLACK);

  if (cur.exist) {
    /** skip this checkbox */
    const skip = range.offset(15, 1, 1, 1).getValue() === 'SKIP';
    if (skip) {
      cur.possible = false;
      // return;
    }

    // cycle through the units
    const units = range.getValues() as string[][];

    const dropdowns: Array<[Spreadsheet.DataValidation]> = [];
    const dropdownsRange = range.offset(0, 1);
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {
      const unitName = units[h][0];
      const idx = platoonMatrix.length;

      if (unitName.length === 0) {
        // no unit was entered, so skip it
        platoonMatrix.push(new PlatoonUnit(unitName, 0, 0));
        dropdowns.push([(null as unknown) as Spreadsheet.DataValidation]);

        continue;
      }

      getNeededCount_(unitName);

      const rec = getRecommendedMembers_(unitName, phase, allUnits);

      platoonMatrix.push(new PlatoonUnit(unitName, 0, rec.length));

      // TODO: investigate bad color
      if (rec.length > 0) {
        dropdowns.push([buildDropdown_(rec)]);

        // add the members to the matrix
        platoonMatrix[idx].members = rec.map((r) => r[0]); // member name
      } else {
        dropdowns.push([(null as unknown) as Spreadsheet.DataValidation]);
        // impossible to fill the platoon if no one can donate
        cur.possible = false;
        spooler.attach(sheet.getRange(row + h, column)).setFontColor(COLOR.RED); // Impossible to fill
      }
    }
    spooler.attach(dropdownsRange).setDataValidations(dropdowns);
  }
}

function loop2_(spooler: utils.Spooler, cur: PlatoonDetails, sheet: Spreadsheet.Sheet) {
  if (cur.exist && !cur.possible) {
    const baseCol = 4; // TODO: should it be a setting

    const row = cur.row;
    const column = cur.getOffset() + baseCol;

    spooler
      .attach(sheet.getRange(row, column + 1, MAX_PLATOON_UNITS))
      .setValue('Skip')
      .clearDataValidations()
      .setFontColor(COLOR.RED); // Recommended is 'Skip'
  }
}

function loop3_(
  spooler: utils.Spooler,
  cur: PlatoonDetails,
  sheet: Spreadsheet.Sheet,
  matrixIdx: number,
  placementCount: number[][],
  platoonMatrix: PlatoonUnit[],
  maxMemberDonations: number,
  used: UnitMemberBooleans,
) {
  if (cur.exist) {
    if (!cur.possible) {
      // skip this platoon
      // tslint:disable-next-line:no-parameter-reassignment
      matrixIdx += MAX_PLATOON_UNITS;
    } else {
      const baseCol = 4; // TODO: should it be a setting

      const row = cur.row;
      const column = cur.getOffset() + baseCol;

      // cycle through the heroes
      const donors: Array<[string]> = [];
      const colors: Array<[COLOR, COLOR]> = [];
      for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {
        let defaultValue;
        const count = placementCount[cur.zone];

        const unit = platoonMatrix[matrixIdx].name;
        for (const member of platoonMatrix[matrixIdx].members) {
          const current = count[(member as unknown) as number] as number;
          const available = current == null || current < maxMemberDonations;
          if (available) {
            // see if the recommended member's hero has been used
            if (used[unit] && used[unit].hasOwnProperty(member) && !used[unit][member]) {
              used[unit][member] = true;
              defaultValue = member;
              count[(member as unknown) as number] = (typeof current === 'number' ? current : 0) + 1;

              break;
            }
          }
        }

        donors.push([defaultValue || '']);

        // see if we should highlight rare units
        if (platoonMatrix[matrixIdx].isMissing()) {
          colors.push([COLOR.RED, COLOR.RED]); // More needed than ready unit
        } else if (defaultValue && platoonMatrix[matrixIdx].isRare()) {
          colors.push([COLOR.BLUE, COLOR.BLUE]);
        } else {
          colors.push([COLOR.BLACK, COLOR.BLACK]);
        }

        // tslint:disable-next-line:no-parameter-reassignment
        matrixIdx += 1;
      }
      spooler.attach(sheet.getRange(row, column + 1, MAX_PLATOON_UNITS)).setValues(donors);
      spooler.attach(sheet.getRange(row, column, MAX_PLATOON_UNITS, 2)).setFontColors(colors);
    }
  }

  return matrixIdx;
}

/** Recommend members for each Platoon */
function recommendPlatoons() {
  const event = config.currentEvent();
  const alignment = config.currentAlignment(event).toLowerCase();

  // const p = new TerritoryBattles.Phase(event(), config.currentPhase());

  const spooler = new utils.Spooler();

  // setup platoon phases
  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOON);
  const PLATOON_NOTAVAILABLE_ROW = 56;
  const PLATOON_NOTAVAILABLE_COL = 4;
  const PLATOON_NOTAVAILABLE_NUMROWS = MAX_MEMBERS;
  const PLATOON_NOTAVAILABLE_NUMCOLS = 1;
  const phase = config.currentPhase();
  const rarityThreshold = isGeoDS_(event) ? (phase < 3 ? 5 : 6) : phase;

  const notAvailable = sheet
    .getRange(
      PLATOON_NOTAVAILABLE_ROW,
      PLATOON_NOTAVAILABLE_COL,
      PLATOON_NOTAVAILABLE_NUMROWS,
      PLATOON_NOTAVAILABLE_NUMCOLS,
    )
    .getValues() as Array<[string]>;

  SPREADSHEET.toast(
    `Using units of rarity ${rarityThreshold + 1}⭐ for ${event} phase ${phase} ready.`,
    'Recommend platoons',
    3,
  );

  // cache the matrix of hero data
  const heroesTable = new Units.Heroes();
  const allHeroes = heroesTable.getAllInstancesByUnits();
  const shipsTable = new Units.Ships();
  const allShips = shipsTable.getAllInstancesByUnits();

  filterUnits_(allHeroes, (member: string, u: UnitInstance) => {
    return (
      u.tags!.indexOf(alignment) !== -1 &&
      u.rarity > rarityThreshold &&
      notAvailable.findIndex((e) => e[0] === member) === -1
    );
  });
  filterUnits_(allShips, (member: string, u: UnitInstance) => {
    return u.rarity > rarityThreshold && notAvailable.findIndex((e) => e[0] === member) === -1;
  });

  // remove heroes listed on Exclusions sheet
  const exclusionsId = config.exclusionId();
  if (exclusionsId.length > 0) {
    const exclusions = Exclusions.getList(phase);
    Exclusions.process(allHeroes, exclusions, alignment);
    Exclusions.process(allShips, exclusions);
  }

  initPlatoonPhases_();

  // reset the needed counts
  PLATOON_NEEDED_COUNT = {};

  // reset the used heroes
  const usedHeroes = resetUsedUnits_(allHeroes);
  const usedShips = resetUsedUnits_(allShips);

  // setup a custom order for walking the platoons
  const platoonOrder: PlatoonDetails[] = [];

  for (let platoon = MAX_PLATOONS - 1; platoon >= 0; platoon -= 1) {
    for (let zone = 0; zone < MAX_PLATOON_ZONES; zone += 1) {
      // last platoon to first platoon
      platoonOrder.push(new PlatoonDetails(phase, zone, platoon));
    }
  }

  // set the zone names and show/hide rows
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    // for each zone
    const platoonRow = 2 + z * PLATOON_ZONE_ROW_OFFSET;

    // set the zone name
    const zoneName = setZoneName_(spooler, phase, z, sheet, platoonRow);

    // see if we should skip the zone
    if (z !== 1) {
      if (zoneName.length === 0) {
        const hideOffset = z === 0 ? 1 : 0;
        sheet.hideRows(platoonRow + hideOffset, MAX_PLATOON_UNITS - hideOffset);
      } else {
        sheet.showRows(platoonRow, MAX_PLATOON_UNITS);
      }
    }
  }

  // initialize platoon matrix
  const platoonMatrix: PlatoonUnit[] = [];

  for (const cur of platoonOrder) {
    loop1_(spooler, cur, platoonMatrix, sheet, phase, cur.isGround ? allHeroes : allShips);
  }

  // update the unit counts
  for (const p of platoonMatrix) {
    const unit = p.name;

    // find the unit's count
    if (PLATOON_NEEDED_COUNT[unit]) {
      p.count = PLATOON_NEEDED_COUNT[unit];
    }
  }

  // make sure the platoon is possible to fill
  for (const cur of platoonOrder) {
    loop2_(spooler, cur, sheet);
  }

  // initialize the placement counts
  const placementCount: number[][] = [];
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    placementCount[z] = [];
  }

  const maxMemberDonations = config.maxDonationsPerTerritory();

  // try to find an unused member to default to
  let matrixIdx = 0;

  for (const cur of platoonOrder) {
    matrixIdx = loop3_(
      spooler,
      cur,
      sheet,
      matrixIdx,
      placementCount,
      platoonMatrix,
      maxMemberDonations,
      cur.isGround ? usedHeroes : usedShips,
    );
  }

  spooler.commit();
  SPREADSHEET.toast(`Platoons for ${event} phase ${phase} ready.`, 'Recommend platoons', 3);
}
