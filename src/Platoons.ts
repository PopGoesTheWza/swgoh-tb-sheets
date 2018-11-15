// ****************************************
// Platoon Functions
// ****************************************

let PLATOON_PHASES: [string, string, string][] = [];
let PLATOON_HERO_NEEDED_COUNT: KeyedNumbers = {};
let PLATOON_SHIP_NEEDED_COUNT: KeyedNumbers = {};

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
    this.exist = zone !== 0 || phase > 2;
  }

  getOffset() {
    return this.platoon * 4;
  }
}

/**
 * Custom object for platoon units
 * hero, count, player count, player list (player, gear...)
 */
class PlatoonUnit {

  public readonly name: string;
  public count: number;
  private readonly pCount: number;
  public players: string[];

  constructor(name: string, count: number, pCount:number) {
    this.name = name;
    this.count = count;
    this.pCount = pCount;
    this.players = [];
  }

  public isMissing(): boolean {
    return this.count > this.pCount;
  }

  public isRare(): boolean {
    return this.count + 3 > this.pCount;
  }

}

/** Initialize the list of Territory names */
function initPlatoonPhases_(): void {

  const filter = config.currentEvent();

  // Names of Territories, # of Platoons
  if (isLight_(filter)) {
    PLATOON_PHASES = [
      ['', 'Rebel Base', ''],
      ['', 'Ion Cannon', 'Overlook'],
      ['Rear Airspace', 'Rear Trenches', 'Power Generator'],
      ['Forward Airspace', 'Forward Trenches', 'Outer Pass'],
      ['Contested Airspace', 'Snowfields', 'Forward Stronghold'],
      ['Imperial Fleet Staging Area', 'Imperial Flank', 'Imperial Landing'],
    ];
  } else if (filter === ALIGNMENT.DARKSIDE) {
    PLATOON_PHASES = [
      ['', 'Imperial Flank', 'Imperial Landing'],
      ['', 'Snowfields', 'Forward Stronghold'],
      ['Imperial Fleet Staging Area', 'Ion Cannon', 'Outer Pass'],
      ['Contested Airspace', 'Power Generator', 'Rear Trenches'],
      ['Forward Airspace', 'Forward Trenches', 'Overlook'],
      [
        'Rear Airspace',
        'Rebel Base Main Entrance',
        'Rebel Base South Entrance',
      ],
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
  phase: number,
  zone: number,
  sheet: Sheet,
  platoonRow: number,
): string {

  // set the zone name
  const zoneName = PLATOON_PHASES[phase - 1][zone];
  spooler.attach(sheet.getRange(platoonRow + 2, 1))
    .setValue(zoneName);

  return zoneName;
}

/** Clear the full chart */
function resetPlatoons(): void {

  const event = config.currentEvent() as TerritoryBattles.event;
  const phase = config.currentPhase() as TerritoryBattles.phaseIdx;
  new TerritoryBattles.Phase(event, phase).reset();
}

function getNeededCount_(unitName: string, isHero: boolean) {
  const count = isHero ? PLATOON_HERO_NEEDED_COUNT : PLATOON_SHIP_NEEDED_COUNT;
  if (count[unitName]) {
    count[unitName] += 1;
  } else {
    count[unitName] = 1;
  }
}

/** Get a sorted list of recommended players */
function getRecommendedPlayers_(
  unitName: string,
  phase: number,
  data: KeyedType<UnitInstances>,
): [string, number][] {

  // see how many stars are needed
  const minRarity = phase + 1;

  const rec: [string, number][] = [];

  const members = data[unitName];
  if (members) {
    for (const player in members) {

      const playerUnit = members[player];

      if (playerUnit.rarity >= minRarity) {
        rec.push([player, playerUnit.power]);
      }
    }
  }

  // sort list by power
  const playerList = rec.sort((a, b) => a[1] - b[1]);  // sorts by 2nd element ascending

  return playerList;
}

/** create the dropdown list */
function buildDropdown_(playerList: [string, number][]): DataValidation {

  const formatList = playerList.map(e => e[0]);

  return SpreadsheetApp.newDataValidation()
    .requireValueInList(formatList)
    .build();
}

/** Reset the units used */
function resetUsedUnits_(data: KeyedType<UnitInstances>): KeyedType<KeyedBooleans> {

  const result: KeyedType<KeyedBooleans> = {};

  for (const unit in data) {
    const members = data[unit];
    for (const player in members) {
      if (!result[unit]) {
        result[unit] = {};
      }
      result[unit][player] = false;
    }
  }

  return result;
}

function filterUnits_(
  data: KeyedType<UnitInstances>,
  filter: (player: string, u: UnitInstance) => boolean,
) {
  const units = Object.assign({}, data);
  for (const unit in units) {
    const members = Object.assign({}, data[unit]);
    for (const player in members) {
      if (!filter(player, members[player])) {
        delete data[unit][player];
      }
    }
    if (data[unit] && Object.keys(data[unit]).length === 0) {
      delete data[unit];
    }
  }
}

namespace TerritoryBattles {

  export type event = ALIGNMENT.LIGHTSIDE|ALIGNMENT.LIGHTSIDE;
  export type phaseIdx = 1|2|3|4|5|6;
  type territoryIdx = 0|1|2;
  type platoonIdx = 0|1|2|3|4|5;

  export class Phase {

    protected readonly sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
    public readonly event: event;
    public readonly index: phaseIdx;
    public readonly useExclusions: boolean = true;
    public readonly useUnavailableq: boolean = true;
    protected readonly territories: Territory[] = [];
    public allHeroes: KeyedType<UnitInstances>;  // units[name][member]
    public allShips: KeyedType<UnitInstances>;  // units[name][member]
    public availableHeroes: KeyedType<UnitInstances>;  // units[name][member]
    public availableShips: KeyedType<UnitInstances>;  // units[name][member]
    public exclusions: KeyedType<KeyedBooleans> = {};  // exclusions[player][unit] = boolean
    public unavailable: string[] = [];

    constructor(event: event, index: phaseIdx) {

      this.event = event;
      this.index = index;

      type tDef = ((p: Phase, i: territoryIdx) => Territory)[];
      const territories: tDef = definitions[event][index];
      territories.forEach((e, i: territoryIdx) => this.territories[i] = e(this, i));
    }

    protected readUnits(): void {
      // cache the matrix of hero data
      const heroesTable = new Units.Heroes();
      this.allHeroes = heroesTable.getAllInstancesByUnits();
      const shipsTable = new Units.Ships();
      this.allShips = shipsTable.getAllInstancesByUnits();
    }

    protected readExclusions(): void {
      const exclusionsId = config.exclusionId();
      if (exclusionsId.length > 0) {
        this.exclusions = Exclusions.getList();
      }
    }

    protected readUnavailable(): void {
      const unavailable = this.sheet
        .getRange(56, 4, MAX_PLAYERS, 1)
        .getValues() as [string][];
      for (const e of unavailable) {
        const name = e[0];
        if (name.length > 0) {
          this.unavailable.push(name);
        }
      }
    }

    recommend(): void {

      const spooler = new utils.Spooler();

      this.readUnits();
      this.readExclusions();
      this.readUnavailable();

      spooler.commit();
    }

    reset(): void {

      const spooler = new utils.Spooler();

      const territories = this.territories;
      for (const territory of territories) {
        territory.readSlices();
        spooler.add(territory.writerName());
        territory.writerSlices().forEach(e => spooler.add(e));
        territory.writerResetDonors().forEach(e => spooler.add(e));
        territory.writerResetButtons().forEach(e => spooler.add(e));
        territory.showHide();
      }

      spooler.commit();
    }

  }

  type slice = string[];
  type slices = [slice, slice, slice, slice, slice, slice];

  abstract class Territory {

    protected readonly sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
    protected readonly phase: Phase;
    public readonly index: territoryIdx;
    protected readonly name: string;
    protected readonly platoons: Platoon[] = [];

    constructor(phase: Phase, index: territoryIdx, name: string) {

      this.index = index;
      this.phase = phase;
      this.name = name;
    }

    showHide(): void {
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

    readSlices(): void {
      const rowOffset = this.phase.event === ALIGNMENT.LIGHTSIDE ? 56 : 2;
      const row = this.index * PLATOON_SLICE_ROW_OFFSET + rowOffset;
      const column = (this.phase.index - 1) * PLATOON_SLICE_COLUMN_OFFSET + 2;
      const data = SPREADSHEET.getSheetByName(SHEETS.SLICES)
        .getRange(row, column, MAX_PLATOON_UNITS, MAX_PLATOONS)
        .getValues() as string[][];
      const result: slices = [[], [], [], [], [], []];
      for (const row of data) {
        row.forEach((e, i) => { result[i].push(e); });
      }
      result.forEach((e, i) => { this.platoons[i].setSlice(e); });
    }

    writerName(): utils.SpooledTask {
      const name = this.name;
      const row = this.index * PLATOON_ZONE_ROW_OFFSET + 4;
      const range = this.sheet.getRange(row, 1);
      const writer = () => { range.setValue(name); };

      return writer;
    }

    writerResetDonors(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map(p => p.writerResetDonors());

      return result;
    }

    writerResetButtons(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map(p => p.writerResetButton());

      return result;
    }

    writerSlices(): utils.SpooledTask[] {
      const platoons = this.platoons;
      const result = platoons.map(p => p.writerSlice());

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

    readSlices(): void {}

  }

  class AirspaceTerritory extends Territory {

    public readonly isOpen = true;
    public readonly isGround = false;

    constructor(phase: Phase, index: territoryIdx, name: string, tp: number[]) {
      super(phase, index, name);
      for (let i = 0; i < MAX_PLATOONS; i += 1) {
        this.platoons[i] = new AirspacePlatoon(this, i as platoonIdx);
      }
    }

  }

  class GroundTerritory extends Territory {

    public readonly isOpen = true;
    public readonly isGround = true;

    constructor(phase: Phase, index: territoryIdx, name: string, tp: number[]) {
      super(phase, index, name);
      for (let i = 0; i < MAX_PLATOONS; i += 1) {
        this.platoons[i] = new GroundPlatoon(this, i as platoonIdx);
      }
    }

  }

  abstract class Platoon {

    protected readonly sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
    protected slice: slice;
    protected readonly territory: Territory;
    protected readonly index: platoonIdx;
    // public readonly isGround: boolean;
    // public readonly isOpen: boolean;
    // public possible: boolean;
    protected readonly row: number;
    protected readonly offset: number;
    protected readonly column: number;

    constructor(territory: Territory, index: platoonIdx) {
      this.index = index;
      this.territory = territory;
      this.row = this.territory.index * PLATOON_ZONE_ROW_OFFSET + 2;
      this.offset = this.index * PLATOON_ZONE_COLUMN_OFFSET;
      this.column = this.offset + 4;
    }

    getResetButton(): boolean {
      const range = this.sheet
        .getRange(this.row + MAX_PLATOON_UNITS, this.column + 1, MAX_PLATOON_UNITS);

      return range.getValue() as string === 'SKIP';
    }

    setSlice(slice: slice): void {
      this.slice = slice;
    }

    writerResetButton(): utils.SpooledTask {
      const range = this.sheet
        .getRange(this.row + MAX_PLATOON_UNITS, this.column + 1, MAX_PLATOON_UNITS);
      const writer = () => { range.clearContent(); };

      return writer;
    }

    writerResetDonors(): utils.SpooledTask {
      const range = this.sheet
        .getRange(this.row, this.column + 1, MAX_PLATOON_UNITS);
      const writer = () =>
        { range.clearContent().clearDataValidations().setFontColor(COLOR.BLACK); };

      return writer;
    }

    writerSlice(): utils.SpooledTask {
      const data = this.slice.map(e => [e]);
      const range = this.sheet
        .getRange(this.row, this.column, MAX_PLATOON_UNITS);
      const writer = () => { range.setValues(data); };

      return writer;
    }

  }

  class ClosedPlatoon extends Platoon {

    constructor(territory: Territory, index: platoonIdx) {
      super(territory, index);
    }

    writerSlice(): utils.SpooledTask {
      const range = this.sheet
        .getRange(this.row, this.column, MAX_PLATOON_UNITS);
      const writer = () => { range.clearContent().setFontColor(COLOR.BLACK); };

      return writer;
    }

  }

  class AirspacePlatoon extends Platoon {

    constructor(territory: Territory, index: platoonIdx) {
      super(territory, index);
    }

  }

  class GroundPlatoon extends Platoon {

    constructor(territory: Territory, index: platoonIdx) {
      super(territory, index);
    }

  }

  const closed = (p: Phase, i: territoryIdx) => new ClosedTerritory(p, i);
  const airspace = (n: string, tp: number[]) =>
    (p: Phase, i: territoryIdx) => new AirspaceTerritory(p, i, n, tp);
  const ground = (n: string, tp: number[]) =>
    (p: Phase, i: territoryIdx) => new GroundTerritory(p, i, n, tp);

  const definitions = {
    'Light Side': {
      1: [
        closed,
        ground('Rebel Base', [100, 100, 100, 100, 150, 150]),
        closed,
      ],
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
    'Dark Side': {
      1: [
        [ClosedTerritory, ''],
        ground('Imperial Flank', [102, 102, 102, 102, 153, 153]),
        ground('Imperial Landing', [102, 102, 102, 102, 153, 153]),
      ],
      2: [
        [ClosedTerritory, ''],
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
  };

  // abstract class UnitPool {

  //   protected readonly phase: Phase;
  //   protected readonly allUnits: KeyedType<UnitInstances>;  // units[name][member]
  //   protected units: KeyedType<UnitInstances>;  // units[name][member]
  //   protected readonly exclusions: KeyedType<KeyedBooleans>;  // excluded[player][unit] = boolean
  //   protected readonly unavailable: string[];

  //   constructor(
  //     phase: Phase,
  //     allUnits: KeyedType<UnitInstances>,
  //     exclusions: KeyedType<KeyedBooleans>,
  //     unavailable: string[],
  //   ) {
  //     this.phase = phase;
  //     this.allUnits = utils.clone(allUnits);
  //     this.exclusions = utils.clone(exclusions);
  //     this.unavailable = utils.clone(unavailable);
  //   }

  //   protected filter(filter: (player: string, u: UnitInstance) => boolean): void {
  //     const data = utils.clone(this.allUnits);
  //     const exclusions = this.exclusions;
  //     const units = Object.assign({}, data);
  //     for (const unit in units) {
  //       const members = Object.assign({}, data[unit]);
  //       for (const player in members) {
  //         if (
  //           (exclusions[player] && exclusions[player][unit])
  //           || !filter(player, members[player])
  //         ) {
  //           delete data[unit][player];
  //         }
  //       }
  //       if (data[unit] && Object.keys(data[unit]).length === 0) {
  //         delete data[unit];
  //       }
  //     }
  //     this.units = data;  // BAD
  //   }

  //   protected zScore() {
  //     const zScored: KeyedType<{
  //       // average: number;
  //       // count: number;
  //       // sum: number;
  //       units: {
  //         // difference: number;
  //         squaredDifference: number;
  //         power: number;
  //         unit: UnitInstance;
  //       }[];
  //     }> = {};
  //     const units = this.units;
  //     for (const unitName in units) {
  //       const perPlayer = units[unitName];
  //       for (const playerName in perPlayer) {
  //         const unit = perPlayer[playerName];
  //         if (!zScored[playerName]) {
  //           zScored[playerName] = {
  //             // average: 0,
  //             // count: 0,
  //             // sum: 0,
  //             units: [],
  //           };
  //         }
  //         zScored[playerName].units.push({
  //           unit,
  //           // difference: 0,
  //           squaredDifference: 0,
  //           power: unit.power,
  //         });
  //       }
  //     }
  //     // const accessor = (u: UnitInstance) => u.power;
  //     for (const playerName in zScored) {
  //       const o = zScored[playerName];
  //       const units = o.units;
  //       const average = statistics.average(units, u => u.power);
  //       // const count = units.length;
  //       // o.count = count;
  //       // const sum = units.reduce((acc, u) => acc + u.power, 0);
  //       // o.sum = sum;
  //       // const average = sum / count;
  //       // o.average = average;
  //       o.units = units.map((e) => {
  //         const difference = e.power - average;
  //         // e.difference = difference;
  //         e.squaredDifference = Math.pow(difference, 2);
  //         return e;
  //       });
  //       const averageSquaredDifference = statistics.average(units, u => u.squaredDifference);
  //     }
  //     const avg = { count: 0, sum: 0 };
  //   }

  // }

  // class HeroesPool extends UnitPool {

  //   constructor(
  //     phase: Phase,
  //     exclusions: KeyedType<KeyedBooleans> = {},
  //     unavailable: string[] = [],
  //   ) {
  //     const heroesTable = new Units.Heroes();
  //     const allUnits = heroesTable.getAllInstancesByUnits();
  //     super(phase, allUnits, exclusions, unavailable);
  //   }

  //   protected filter() {
  //     const phase = this.phase.index;
  //     const alignment = config.currentEvent().toLowerCase();
  //     const filter = (player: string, u: UnitInstance) => {
  //       return u.rarity > phase
  //         && u.tags.indexOf(alignment) !== -1
  //         && this.unavailable.findIndex(e => e[0] === player) === -1;
  //     };
  //     super.filter(filter);
  //   }

  // }

  // class ShipsPool extends UnitPool {

  //   constructor(
  //     phase: Phase,
  //     exclusions: KeyedType<KeyedBooleans> = {},
  //     unavailable: string[] = [],
  //   ) {
  //     const shipsTable = new Units.Ships();
  //     const allUnits = shipsTable.getAllInstancesByUnits();
  //     super(phase, allUnits, exclusions, unavailable);
  //   }

  //   protected filter() {
  //     const phase = this.phase.index;
  //     const filter = (player: string, u: UnitInstance) => {
  //       return u.rarity > phase
  //         && this.unavailable.findIndex(e => e[0] === player) === -1;
  //     };
  //     super.filter(filter);
  //   }

  // }

}

// function loop1_(
//   cur: PlatoonDetails,
//   platoonMatrix: PlatoonUnit[],
//   sheet: Sheet,
//   phase: number,
//   allHeroes: KeyedType<UnitInstances>,
//   allShips: KeyedType<UnitInstances>,
//   allDropdowns: [Range, [DataValidation][]][],
// ) {
//   const baseCol = 4;  // TODO: should it be a setting

//   const row = cur.row;
//   const column = cur.getOffset() + baseCol;
//   const range = sheet.getRange(row, column, MAX_PLATOON_UNITS, 1);

//   // clear previous contents
//   range.offset(0, 1)
//     .clearContent()
//     .clearDataValidations()
//     .offset(0, -1, MAX_PLATOON_UNITS, 2)
//     .setFontColor(COLOR.BLACK);

//   if (cur.exist) {

//     /** 'Skip this' checkbox */
//     const skip = sheet.getRange(row + 15, column + 1)
//       .getValue() === 'SKIP';
//     if (skip) {
//       cur.possible = false;
//     } else {

//       // cycle through the units
//       const units = range.getValues().map((e: string[]) => e[0]);

//       const dropdowns: [DataValidation][] = [];
//       const dropdownsRange = range.offset(0, 1);

//       units.forEach((unitName, h) => {
//         const idx = platoonMatrix.length;
//         if (unitName.length === 0) {  // no unit was entered, so skip it
//           platoonMatrix.push(new PlatoonUnit(unitName, 0, 0));
//           dropdowns.push([null]);
//         } else {
//           getNeededCount_(unitName, cur.isGround);

//           const rec = getRecommendedPlayers_(
//             unitName,
//             phase,
//             cur.isGround ? allHeroes : allShips,
//           );

//           platoonMatrix.push(new PlatoonUnit(unitName, 0, rec.length));

//           if (rec.length > 0) {
//             dropdowns.push([buildDropdown_(rec)]);

//             // add the players to the matrix
//             platoonMatrix[idx].players = rec.map(r => r[0]); // player name
//           } else {
//             dropdowns.push([null]);
//             // impossible to fill the platoon if no one can donate
//             cur.possible = false;
//             sheet.getRange(row + h, column).setFontColor(COLOR.RED);
//           }
//         }

//       });
//       allDropdowns.push([dropdownsRange, dropdowns]);
//     }
//   }
// }

// function loop2_(cur: PlatoonDetails, sheet: Sheet) {
//   if (cur.exist && !cur.possible) {
//     const baseCol = 4;  // TODO: should it be a setting

//     const row = cur.row;
//     const column = cur.getOffset() + baseCol;
//     const range = sheet.getRange(row, column + 1, MAX_PLATOON_UNITS, 1);

//     range.setValue('Skip')
//       .clearDataValidations()
//       .setFontColor(COLOR.RED);
//   }
// }

// function loop3_(
//   cur: PlatoonDetails,
//   sheet: Sheet,
//   matrixIdx: number,
//   placementCount: number[][],
//   platoonMatrix: PlatoonUnit[],
//   maxPlayerDonations: number,
//   usedHeroes: KeyedType<KeyedBooleans>,
//   usedShips: KeyedType<KeyedBooleans>,
//   baseCol: number,
//   allDonors:[Range, [string][]][],
//   allColors: [Range, [COLOR, COLOR][]][],
// ) {
//   if (cur.exist) {
//     if (!cur.possible) {
//       // skip this platoon
//       matrixIdx += MAX_PLATOON_UNITS;
//     } else {
//       // cycle through the heroes
//       const donors: [string][] = [];
//       const colors: [COLOR, COLOR][] = [];
//       const plattonOffset = cur.getOffset();
//       for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {

//         let defaultValue = '';
//         const count = placementCount[cur.zone];

//         const unit = platoonMatrix[matrixIdx].name;
//         for (const player of platoonMatrix[matrixIdx].players) {

//           const available = count[player] == null || count[player] < maxPlayerDonations;
//           if (!available) {
//             continue;
//           }

//           // see if the recommended player's hero has been used
//           if (cur.isGround) {
//             // ground units
//             if (usedHeroes[unit]
//               && usedHeroes[unit].hasOwnProperty(player)
//               && !usedHeroes[unit][player]
//             ) {
//               usedHeroes[unit][player] = true;
//               defaultValue = player;
//               count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

//               break;
//             }
//           } else {
//             // ships
//             if (usedShips[unit]
//               && usedShips[unit].hasOwnProperty(player)
//               && !usedShips[unit][player]
//             ) {
//               usedShips[unit][player] = true;
//               defaultValue = player;
//               count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

//               break;
//             }
//           }

//           if (defaultValue.length > 0) {
//             // we already have a recommended player
//             break;
//           }
//         }

//         donors.push([defaultValue]);

//         // see if we should highlight rare units
//         if (platoonMatrix[matrixIdx].isMissing()) {
//           colors.push([COLOR.RED, COLOR.RED]);
//         } else if (defaultValue.length > 0 && platoonMatrix[matrixIdx].isRare()) {
//           colors.push([COLOR.BLUE, COLOR.BLUE]);
//         } else {
//           colors.push([COLOR.BLACK, COLOR.BLACK]);
//         }

//         matrixIdx += 1;
//       }
//       const donorsRange =
//         sheet.getRange(cur.row, baseCol + plattonOffset + 1, MAX_PLATOON_UNITS, 1);
//       allDonors.push([donorsRange, donors]);
//       const colorsRange =
//         sheet.getRange(cur.row, baseCol +  plattonOffset, MAX_PLATOON_UNITS, 2);
//       allColors.push([colorsRange, colors]);
//     }
//   }
// }

/** Recommend players for each Platoon */
function recommendPlatoons() {

  const spooler = new utils.Spooler();

  // setup platoon phases
  const sheet = SPREADSHEET.getSheetByName(SHEETS.PLATOONS);
  const phase = config.currentPhase();
  const alignment = config.currentEvent().toLowerCase();
  const unavailable = sheet.getRange(56, 4, MAX_PLAYERS, 1).getValues() as [string][];

  // cache the matrix of hero data
  const heroesTable = new Units.Heroes();
  const allHeroes = heroesTable.getAllInstancesByUnits();
  const shipsTable = new Units.Ships();
  const allShips = shipsTable.getAllInstancesByUnits();

  filterUnits_(
    allHeroes,
    (player: string, u: UnitInstance) => {
      return u.rarity > phase
        && u.tags.indexOf(alignment) !== -1
        && unavailable.findIndex(e => e[0] === player) === -1;
    },
  );
  filterUnits_(
    allShips,
    (player: string, u: UnitInstance) => {
      return u.rarity > phase
        && unavailable.findIndex(e => e[0] === player) === -1;
    },
  );

  // remove heroes listed on Exclusions sheet
  const exclusionsId = config.exclusionId();
  if (exclusionsId.length > 0) {
    const exclusions = Exclusions.getList();
    Exclusions.process(allHeroes, exclusions);
    Exclusions.process(allShips, exclusions, config.currentEvent());
  }

  initPlatoonPhases_();

  // reset the needed counts
  PLATOON_HERO_NEEDED_COUNT = {};
  PLATOON_SHIP_NEEDED_COUNT = {};

  // reset the used heroes
  const usedHeroes = resetUsedUnits_(allHeroes);
  const usedShips = resetUsedUnits_(allShips);

  // setup a custom order for walking the platoons
  const platoonOrder: PlatoonDetails[] = [];

  for (let platoon = MAX_PLATOONS - 1; platoon >= 0; platoon -= 1) {

    for (let zone = 0; zone < MAX_PLATOON_ZONES; zone += 1) {  // last platoon to first platoon
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
  const baseCol = 4;

  for (const cur of platoonOrder) {

    const platoonOffset = cur.getOffset();

    // clear previous contents
    spooler.attach(sheet.getRange(cur.row, baseCol + platoonOffset + 1, MAX_PLATOON_UNITS, 1))
      .clearContent()
      .clearDataValidations()
      .offset(0, -1, MAX_PLATOON_UNITS, 2)
      .setFontColor(COLOR.BLACK);

    if (!cur.exist) {  // skip this zone
      continue;
    }

    /** skip this checkbox */
    const skip = sheet.getRange(cur.row + 15, baseCol + platoonOffset + 1, 1, 1)
        .getValue() === 'SKIP';
    if (skip) {
      cur.possible = false;
    }

    // cycle through the units
    const units = sheet.getRange(cur.row, baseCol + platoonOffset, MAX_PLATOON_UNITS, 1)
      .getValues() as string[][];

    const dropdowns: [DataValidation][] = [];
    const dropdownsRange = sheet.getRange(
      cur.row,
      baseCol + platoonOffset + 1,
      MAX_PLATOON_UNITS,
      1,
    );
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {

      const unitName = units[h][0];
      const idx = platoonMatrix.length;

      if (unitName.length === 0) {
        // no unit was entered, so skip it
        platoonMatrix.push(new PlatoonUnit(unitName, 0, 0));
        dropdowns.push([null]);

        continue;
      }

      getNeededCount_(unitName, cur.isGround);

      const rec = getRecommendedPlayers_(
        unitName,
        phase,
        cur.isGround ? allHeroes : allShips,
      );

      platoonMatrix.push(new PlatoonUnit(unitName, 0, rec.length));

      if (rec.length > 0) {
        dropdowns.push([buildDropdown_(rec)]);

        // add the players to the matrix
        platoonMatrix[idx].players = rec.map(r => r[0]); // player name
      } else {
        dropdowns.push([null]);
        // impossible to fill the platoon if no one can donate
        cur.possible = false;
        spooler.attach(sheet.getRange(cur.row + h, baseCol + platoonOffset))
          .setFontColor(COLOR.RED);
      }
    }
    spooler.attach(dropdownsRange)
      .setDataValidations(dropdowns);
  }

  // update the unit counts
  for (const p of platoonMatrix) {

    const unit = p.name;

    // find the unit's count
    if (PLATOON_HERO_NEEDED_COUNT[unit]) {
      p.count = PLATOON_HERO_NEEDED_COUNT[unit];
    } else if (PLATOON_SHIP_NEEDED_COUNT[unit]) {
      p.count = PLATOON_SHIP_NEEDED_COUNT[unit];
    }
  }

  // make sure the platoon is possible to fill
  for (const cur of platoonOrder) {

    if (!cur.exist) {  // skip this zone
      continue;
    }

    if (!cur.possible) {
      const plattonOffset = cur.getOffset();

      spooler.attach(sheet.getRange(cur.row, baseCol + 1 + plattonOffset, MAX_PLATOON_UNITS, 1))
        .setValue('Skip')
        .clearDataValidations()
        .setFontColor(COLOR.RED);
    }
  }

  // initialize the placement counts
  const placementCount: number[][] = [];
  for (let z = 0; z < MAX_PLATOON_ZONES; z += 1) {
    placementCount[z] = [];
  }

  const maxPlayerDonations = config.maxDonationsPerTerritory();

  // try to find an unused player to default to
  let matrixIdx = 0;

  for (const cur of platoonOrder) {

    if (!cur.exist) {  // skip this zone
      continue;
    }

    if (!cur.possible) {
      // skip this platoon
      matrixIdx += MAX_PLATOON_UNITS;
      continue;
    }

    // cycle through the heroes
    const donors: [string][] = [];
    const colors: [COLOR, COLOR][] = [];
    const plattonOffset = cur.getOffset();
    for (let h = 0; h < MAX_PLATOON_UNITS; h += 1) {

      let defaultValue = '';
      const count = placementCount[cur.zone];

      const unit = platoonMatrix[matrixIdx].name;
      for (const player of platoonMatrix[matrixIdx].players) {

        const available = count[player] == null || count[player] < maxPlayerDonations;
        if (!available) {
          continue;
        }

        // see if the recommended player's hero has been used
        if (cur.isGround) {
          // ground units
          if (usedHeroes[unit]
            && usedHeroes[unit].hasOwnProperty(player)
            && !usedHeroes[unit][player]
          ) {
            usedHeroes[unit][player] = true;
            defaultValue = player;
            count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

            break;
          }
        } else {
          // ships
          if (usedShips[unit]
            && usedShips[unit].hasOwnProperty(player)
            && !usedShips[unit][player]
          ) {
            usedShips[unit][player] = true;
            defaultValue = player;
            count[player] = (typeof count[player] === 'number') ? count[player] + 1 : 0;

            break;
          }
        }

        if (defaultValue.length > 0) {
          // we already have a recommended player
          break;
        }
      }

      donors.push([defaultValue]);

      // see if we should highlight rare units
      if (platoonMatrix[matrixIdx].isMissing()) {
        colors.push([COLOR.RED, COLOR.RED]);
      } else if (defaultValue.length > 0 && platoonMatrix[matrixIdx].isRare()) {
        colors.push([COLOR.BLUE, COLOR.BLUE]);
      } else {
        colors.push([COLOR.BLACK, COLOR.BLACK]);
      }

      matrixIdx += 1;
    }
    spooler.attach(sheet.getRange(cur.row, baseCol + plattonOffset + 1, MAX_PLATOON_UNITS, 1))
      .setValues(donors);
    spooler.attach(sheet.getRange(cur.row, baseCol +  plattonOffset, MAX_PLATOON_UNITS, 2))
      .setFontColors(colors);
  }

  spooler.commit();
}
