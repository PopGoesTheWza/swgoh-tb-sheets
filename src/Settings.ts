/** workaround to tslint issue of namespace scope after importingtype definitions */
declare namespace SwgohHelp {

  function getGuildData(): PlayerData[];
  function getPlayerData(allyCode: number): PlayerData;
  function getUnitList(): UnitsDefinitions;

}

/** Shortcuts for Google Apps Script classes */
const SPREADSHEET = SpreadsheetApp.getActive();
// const UI = SpreadsheetApp.getUi();

type DataValidation = GoogleAppsScript.Spreadsheet.DataValidation;
type Range = GoogleAppsScript.Spreadsheet.Range;
type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
type URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

/** Global constants */
const MAX_MEMBERS = 50;
// const MIN_PLAYER_LEVEL = 65
// const POWER_TARGET = 14000

// Meta tab columns
const META_FILTER_ROW = 2;
const META_FILTER_COL = 2;
const META_TAG_ROW = 2;
const META_TAG_COL = 3;
const META_UNDERGEAR_ROW = 2;
const META_UNDERGEAR_COL = 4;
const META_SORT_ROSTER_ROW = 2;
const META_SORT_ROSTER_COL = 5;
const META_EXCLUSSIONS_ROW = 7;
const META_EXCLUSSIONS_COL = 1;
const META_MIN_LEVEL_ROW = 5;
const META_MIN_LEVEL_COL = 4;
const META_MIN_GP_ROW = 8;
const META_MIN_GP_COL = 4;
const META_UNIT_PER_MEMBER_ROW = 11;
const META_UNIT_PER_MEMBER_COL = 4;

const META_DATASOURCE_ROW = 14;
const META_DATASOURCE_COL = 4;

const META_GUILDDATADATE_ROW = 24;
const META_GUILDDATADATE_COL = 1;
const META_UNITDEFINITIONSDATE_ROW = 26;
const META_UNITDEFINITIONSDATE_COL = 1;

const META_HEROES_COL = 7;
const META_HEROES_DS_COL = 16;

const META_HEROES_COUNT_ROW = 5;
const META_HEROES_COUNT_COL = 5;
const META_SHIPS_COUNT_ROW = 8;
const META_SHIPS_COUNT_COL = 5;

const META_RENAME_ADD_PLAYER_COL = 16;
const META_REMOVE_PLAYER_COL = 18;

// Hero/Ship tab columns
const HERO_MEMBER_COL_OFFSET = 11;
const SHIP_MEMBER_COL_OFFSET = 11;

// Roster Size info
const META_GUILD_SIZE_ROW = 5;
const META_GUILD_SIZE_COL = 12;

const META_TB_COL_OFFSET = 10;

const WAIT_TIME = 2000;

const MAX_PLATOON_UNITS = 15;
const MAX_PLATOONS = 6;
const MAX_PLATOON_ZONES = 3;
const PLATOON_SLICE_ROW_OFFSET = 17;  // MAX_PLATOON_UNITS + 2;
const PLATOON_SLICE_COLUMN_OFFSET = 7;  // MAX_PLATOONS + 1;
const PLATOON_ZONE_ROW_OFFSET = 18;  // MAX_PLATOON_UNITS + 3;
const PLATOON_ZONE_COLUMN_OFFSET = 4;

const RARE_MAX = 15;
const HIGH_MIN = 10;
const DISCORD_WEBHOOK_COL = 5;
const WEBHOOK_TB_START_ROW = 3;
const WEBHOOK_PHASE_HOURS_ROW = 4;
const WEBHOOK_TITLE_ROW = 5;
const WEBHOOK_WARN_ROW = 6;
const WEBHOOK_RARE_ROW = 7;
const WEBHOOK_DEPTH_ROW = 8;
const WEBHOOK_DESC_ROW = 9;
const WEBHOOK_CLEAR_ROW = 15;
const WEBHOOK_DISPLAY_SLOT_ROW = 16;

const META_UPDATETIME_COL = 1;  // ADDED FOR TIMESTAMP
const META_UPDATETIME_ROW = 11;  // ADDED FOR TIMESTAMP

type KeyedType<T> = {
  [key: string]: T;
};

type KeyedBooleans = KeyedType<boolean>;
type KeyedNumbers = KeyedType<number>;
type KeyedStrings = KeyedType<string>;

interface PlayerData {
  allyCode: number;
  gp: number;
  heroesGp: number;
  level?: number;
  link?: string;
  name: string;
  shipsGp: number;
  units: UnitInstances;
}

/** A unit's name, baseId and tags */
interface UnitDefinition {
  /** Unit Id */
  baseId: string;
  name: string;
  /** Alignment, role and tags */
  tags: string;
  type?: number;
}

/**
 * A unit instance attributes
 * (baseId, gearLevel, level, name, power, rarity, stats, tags)
 */
interface UnitInstance {
  baseId?: string;
  gearLevel?: number;
  level: number;
  name?: string;
  power: number;
  /** calculated and used in platoons recommendation */
  zScore?: number;
  rarity: number;
  stats?: string;
  tags?: string;
}

type MemberInstances = {
  [key: string]: UnitInstance;
};

type UnitInstances = {
  [key: string]: UnitInstance;
};

type MemberUnitBooleans = {
  [key: string]: {
    [key: string]: boolean;
  };
};

type MemberUnitInstances = {
  [key: string]: {
    [key: string]: UnitInstance;
  };
};

type UnitMemberBooleans = {
  [key: string]: {
    [key: string]: boolean;
  };
};

type UnitMemberInstances = {
  [key: string]: {
    [key: string]: UnitInstance;
  };
};

/** Constants for alignment */
enum ALIGNMENT {
  DARKSIDE = 'Dark Side',
  LIGHTSIDE = 'Light Side',
}

/** Constants for background colors */
enum COLOR {
  BLACK = 'Black',
  BLUE = 'Blue',
  RED = 'Red',
}

/** Constants for data source */
enum DATASOURCES {
  /** Use swgoh.help API as data source */
  SWGOH_HELP = 'SWGoH.help',
  /** Use swgoh.gg API as data source */
  SWGOH_GG = 'SWGoH.gg',
}

/** Constants for display options */
enum DISPLAYSLOT {
  ALWAYS = 'Always',
  DEFAULT = 'Default',
  NEVER = 'Never',
}

/** Constants for sheets name */
enum SHEETS {
  ROSTER = 'Roster',
  TB = 'TB',
  PLATOONS = 'Platoon',
  BREAKDOWN = 'Breakdown',
  ESTIMATE = 'Estimate',
  LSMISSIONS = 'LS Missions',
  DSMISSIONS = 'DS Missions',
  SNAPSHOT = 'Snapshot',
  EXCLUSIONS = 'Exclusions',
  HEROES = 'Heroes',
  SHIPS = 'Ships',
  RAREUNITS = 'Rare Units',
  SEARCHUNITS = 'Search Units',
  SLICES = 'Slices',
  MAP = 'map',
  DISCORD = 'Discord',
  META = 'Meta',
  INSTRUCTIONS = 'Instructions',
}

/** settings related functions */
namespace config {

  /** get current Territory Battles phase */
  export function currentPhase(): TerritoryBattles.phaseIdx {

    const value = +SPREADSHEET.getSheetByName(SHEETS.PLATOONS)
      .getRange(2, 1)
      .getValue() as TerritoryBattles.phaseIdx;

    return value;
  }

  /** get current event */
  export function currentEvent(): ALIGNMENT {

    const value = SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_FILTER_ROW, META_FILTER_COL)
      .getValue() as ALIGNMENT;

    return value;
  }

  /** get the tag/faction of current event */
  export function tagFilter(): string {

    const value = SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_TAG_ROW, META_TAG_COL)
      .getValue() as string;

    return value;
  }

  /** get required minimum player level */
  export function requiredHeroGp(): number {

    const value = +SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_MIN_GP_ROW, META_MIN_GP_COL)
      .getValue();

    return value;
  }

  /** get required minimum player level */
  export function requiredMemberLevel(): number {

    const value = +SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_MIN_LEVEL_ROW, META_MIN_LEVEL_COL)
      .getValue();

    return value;
  }

  /** get maximum allowed donation per territory */
  export function maxDonationsPerTerritory(): number {

    const value = +SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_UNIT_PER_MEMBER_ROW, META_UNIT_PER_MEMBER_COL)
      .getValue();

    return value;
  }

  /** get roster sorting setting */
  export function sortRoster(): boolean {

    const value = SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_SORT_ROSTER_ROW, META_SORT_ROSTER_COL)
      .getValue() as string;

    return value === 'Yes';
  }

  /** get Id of exclusions */
  export function exclusionId(): string {

    const value = SPREADSHEET.getSheetByName(SHEETS.META)
      .getRange(META_EXCLUSSIONS_ROW, META_EXCLUSSIONS_COL)
      .getValue() as string;

    return value;
  }

  /** get count of members in the roster */
  export function memberCount(): number {

    const value = +SPREADSHEET.getSheetByName(SHEETS.ROSTER)
      .getRange(META_GUILD_SIZE_ROW, META_GUILD_SIZE_COL)
      .getValue();

    return value;
  }

  /** data source related settings */
  export namespace dataSource {

    /** should we use the SWGoH.help API? */
    export function isSwgohHelp(): boolean {
      return getDataSource() === DATASOURCES.SWGOH_HELP;
    }

    /** should we use the SWGoH.gg API? */
    export function isSwgohGg(): boolean {
      return getDataSource() === DATASOURCES.SWGOH_GG;
    }

    /** get selected data source */
    export function getDataSource(): string {

      const value = SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(META_DATASOURCE_ROW, META_DATASOURCE_COL)
        .getValue() as string;

      return value;
    }

    export function setGuildDataDate(): void {
      SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(META_GUILDDATADATE_ROW, META_GUILDDATADATE_COL)
        .setValue(new Date());
    }

    export function setUnitDefinitionsDate(): void {
      SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(META_UNITDEFINITIONSDATE_ROW, META_UNITDEFINITIONSDATE_COL)
        .setValue(new Date());
    }

  }

  /** SwgohGg related settings */
  export namespace SwgohGg {

    /** Get the guild id */
    export function guild(): number {

      const metaSWGOHLinkCol = 1;
      const metaSWGOHLinkRow = 2;
      const guildLink = SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
        .getValue() as string;
      const parts = guildLink.split('/');
      // TODO: input check
      const guildId = +parts[4];

      return guildId;
    }

  }

  /** SwgohHelp related settings */
  export namespace SwgohHelp {

    /** Get the SwgohHelp API username */
    export function username(): string {

      const metaSWGOHLinkCol = 1;
      const metaSWGOHLinkRow = 16;
      const result = SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
        .getValue() as string;

      return result;
    }

    /** Get the SwgohHelp API password */
    export function password(): string {

      const metaSWGOHLinkCol = 1;
      const metaSWGOHLinkRow = 18;
      const result = SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
        .getValue() as string;

      return result;
    }

    /** Get the guild member ally code */
    export function allyCode(): number {

      const metaSWGOHLinkCol = 1;
      const metaSWGOHLinkRow = 20;
      const result = +SPREADSHEET.getSheetByName(SHEETS.META)
        .getRange(metaSWGOHLinkRow, metaSWGOHLinkCol)
        .getValue();

      return result;
    }

  }

  /** discord related settings */
  export namespace discord {

    /** Get the webhook address */
    export function webhookUrl(): string {

      const value = SPREADSHEET.getSheetByName(SHEETS.DISCORD)
        .getRange(1, DISCORD_WEBHOOK_COL)
        .getValue() as string;

      return value;
    }

    /** Get the role to mention */
    export function roleId(): string {

      const value = SPREADSHEET.getSheetByName(SHEETS.DISCORD)
        .getRange(2, DISCORD_WEBHOOK_COL)
        .getValue() as string;

      return value;
    }

    /** Get the time and date when the TB started */
    export function startTime(): Date {

      const value = SPREADSHEET.getSheetByName(SHEETS.DISCORD)
        .getRange(WEBHOOK_TB_START_ROW, DISCORD_WEBHOOK_COL)
        .getValue() as Date;

      return value;
    }

    /** Get the number of hours in each phase */
    export function phaseDuration(): number {

      const value = +SPREADSHEET.getSheetByName(SHEETS.DISCORD)
        .getRange(WEBHOOK_PHASE_HOURS_ROW, DISCORD_WEBHOOK_COL)
        .getValue();

      return value;
    }

    /** Get the template for a webhooks */
    export function webhookTemplate(
      phase: TerritoryBattles.phaseIdx,
      row: number,
      defaultVal: string,
    ): string {

      const text = SPREADSHEET.getSheetByName(SHEETS.DISCORD)
        .getRange(row, DISCORD_WEBHOOK_COL)
        .getValue() as string;

      return text.length > 0 ? text.replace('{0}', String(phase)) : defaultVal;
    }

    /** Get the Description for the phase */
    export function webhookDescription(phase: TerritoryBattles.phaseIdx): string {

      const columnOffset = isLight_(config.currentEvent()) ? 0 : 1;
      const text = SPREADSHEET.getSheetByName(SHEETS.DISCORD)
        .getRange(WEBHOOK_DESC_ROW + phase - 1, DISCORD_WEBHOOK_COL + columnOffset)
        .getValue() as string;

      return `\n\n${text}`;
    }

    /** See if the platoons should be cleared */
    export function resetPlatoons(): boolean {

      const value = SPREADSHEET.getSheetByName(SHEETS.DISCORD)
        .getRange(WEBHOOK_CLEAR_ROW, DISCORD_WEBHOOK_COL)
        .getValue() as string;

      return value === 'Yes';
    }

    /** See if the slot number should be displayed */
    export function displaySlots(): string {

      const value = SPREADSHEET.getSheetByName(SHEETS.DISCORD)
        .getRange(WEBHOOK_DISPLAY_SLOT_ROW, DISCORD_WEBHOOK_COL)
        .getValue() as string;

      return value;
    }

  }

}
