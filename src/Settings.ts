/// <reference types="google-apps-script" />
/** workaround to tslint issue of namespace scope after importingtype definitions */
declare namespace SwgohHelp {
  function getGuildData(): PlayerData[];
  function getPlayerData(allyCode: number): PlayerData;
  function getUnitList(): UnitsDefinitions;
}

/** Shortcuts for Google Apps Script classes */
const SPREADSHEET = SpreadsheetApp.getActive();
// const UI = SpreadsheetApp.getUi();

import Spreadsheet = GoogleAppsScript.Spreadsheet;
import URL_Fetch = GoogleAppsScript.URL_Fetch;

/** Global constants */
const MAX_MEMBERS = 50;

// Meta tab columns
// TODO: use this meta setting
// const META_UNDERGEAR_ROW = 2;
// const META_UNDERGEAR_COL = 4;

const META_SQUADS_GEODS_COL = 25;
const META_SQUADS_HOTHDS_COL = 16;
const META_SQUADS_HOTHLS_COL = 7;

const MAX_PLATOON_UNITS = 15;
const MAX_PLATOONS = 6;
const MAX_PLATOON_ZONES = 3;
const PLATOON_ZONE_ROW_OFFSET = 18; // MAX_PLATOON_UNITS + 3;
const PLATOON_ZONE_COLUMN_OFFSET = 4;

// TODO: define RARE
const RARE_MAX = 15;
const HIGHLY_NEEDED = 8;
const DISCORD_WEBHOOK_COL = 5;

interface KeyedType<T> {
  [key: string]: T;
}

type KeyedBooleans = KeyedType<boolean>;
type KeyedNumbers = KeyedType<number>;
type KeyedStrings = KeyedType<string>;
interface PlayerData {
  /** allycode */
  allyCode: number;
  /** gp */
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
  type: Units.TYPES;
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

interface MemberInstances {
  [key: string]: UnitInstance;
}

interface UnitInstances {
  [key: string]: UnitInstance;
}

interface MemberUnitBooleans {
  [key: string]: {
    [key: string]: boolean;
  };
}

interface MemberUnitInstances {
  [key: string]: {
    [key: string]: UnitInstance;
  };
}

interface UnitMemberBooleans {
  [key: string]: {
    [key: string]: boolean;
  };
}

interface UnitMemberInstances {
  [key: string]: {
    [key: string]: UnitInstance;
  };
}

/** Constants for event */
enum EVENT {
  GEONOSISDS = 'Geo DS',
  GEONOSISLS = 'Geo LS',
  HOTHDS = 'Hoth DS',
  HOTHLS = 'Hoth LS',
}

/** Constants for alignment */
enum ALIGNMENT {
  DARKSIDE = 'Dark Side',
  LIGHTSIDE = 'Light Side',
  UNSPECIFIED = '(unspecified)',
}

/** Constants for background colors */
enum COLOR {
  MAROON = 'Maroon',
  RED = 'Red',
  ORANGE = 'Orange',
  YELLOW = 'Yellow',
  OLIVE = 'Olive',
  GREEN = 'Green',
  PURPLE = 'Purple',
  FUCHSIA = 'Fuchsia',
  LIME = 'Lime',
  TEAL = 'Teal',
  AQUA = 'Aqua',
  BLUE = 'Blue',
  NAVY = 'Navy',
  BLACK = 'Black',
  GRAY = 'Gray',
  SILVER = 'Silver',
  WHITE = 'White',
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
enum SHEET {
  ROSTER = 'Roster',
  TB = 'TB',
  PLATOON = 'Platoon',
  ASSIGNMENTS = 'Assignments',
  GEODSPLATOONAUDIT = 'GeoDSPlatoonAudit',
  GEOSQUADRONAUDIT = 'GeoSquadronAudit',
  GEONEEDEDUNITS = 'GeoNeededUnits',
  DSPLATOONAUDIT = 'HothDSPlatoonAudit',
  LSPLATOONAUDIT = 'HothLSPlatoonAudit',
  SQUADRONAUDIT = 'HothSquadronAudit',
  NEEDEDUNITS = 'HothNeededUnits',
  BREAKDOWN = 'Breakdown',
  ESTIMATE = 'Estimate',
  GEODSMISSIONS = 'GeoDSMissions',
  DSMISSIONS = 'HothDSMissions',
  LSMISSIONS = 'HothLSMissions',
  SNAPSHOT = 'Snapshot',
  EXCLUSIONS = 'Exclusions',
  HEROES = 'Heroes',
  SHIPS = 'Ships',
  RAREUNITS = 'Rare Units',
  SEARCHUNITS = 'Search Units',
  DISCORD = 'Discord',
  META = 'Meta',
  INSTRUCTIONS = 'Instructions',
  STATICSLICES = 'StaticSlices',
  GEODSPLATOON = 'GeoDSPlatoon',
  GEOSQUADRON = 'GeoSquadron',
  HOTHDSPLATOON = 'HothDSPlatoon',
  HOTHLSPLATOON = 'HothLSPlatoon',
  HOTHSQUADRON = 'HothSquadron',
}

/** settings related functions */
namespace config {
  /** get current event */
  export function currentAlignment(event = currentEvent()): ALIGNMENT {
    return isDarkSide_(event) ? ALIGNMENT.DARKSIDE : isLightSide_(event) ? ALIGNMENT.LIGHTSIDE : ALIGNMENT.UNSPECIFIED;
  }

  /** get current event */
  export function currentEvent(): EVENT {
    const META_CURRENTEVENT_ROW = 2;
    const META_CURRENTEVENT_COL = 2;
    return SPREADSHEET.getSheetByName(SHEET.META)
      .getRange(META_CURRENTEVENT_ROW, META_CURRENTEVENT_COL)
      .getValue();
  }

  const PLATOON_CURRENTPHASE_ROW = 2;
  const PLATOON_CURRENTPHASE_COL = 1;

  /** get current Territory Battles phase */
  export function currentPhase(): TerritoryBattles.phaseIdx {
    return +SPREADSHEET.getSheetByName(SHEET.PLATOON)
      .getRange(PLATOON_CURRENTPHASE_ROW, PLATOON_CURRENTPHASE_COL)
      .getValue() as TerritoryBattles.phaseIdx;
  }

  export function setPhase(phase: TerritoryBattles.phaseIdx) {
    SPREADSHEET.getSheetByName(SHEET.PLATOON)
      .getRange(PLATOON_CURRENTPHASE_ROW, PLATOON_CURRENTPHASE_COL)
      .setValue(phase);
  }

  /** get the tag/faction of current event */
  export function tagFilter(): string {
    const META_TAG_ROW = 2;
    const META_TAG_COL = 3;
    return SPREADSHEET.getSheetByName(SHEET.META)
      .getRange(META_TAG_ROW, META_TAG_COL)
      .getValue();
  }

  /** get required minimum player level */
  export function requiredHeroGp(): number {
    const META_MIN_GP_ROW = 8;
    const META_MIN_GP_COL = 4;
    return +SPREADSHEET.getSheetByName(SHEET.META)
      .getRange(META_MIN_GP_ROW, META_MIN_GP_COL)
      .getValue();
  }

  /** get required minimum player level */
  export function requiredMemberLevel(): number {
    const META_MIN_LEVEL_ROW = 5;
    const META_MIN_LEVEL_COL = 4;
    return +SPREADSHEET.getSheetByName(SHEET.META)
      .getRange(META_MIN_LEVEL_ROW, META_MIN_LEVEL_COL)
      .getValue();
  }

  /** get maximum allowed donation per territory */
  export function maxDonationsPerTerritory(): number {
    const META_UNIT_PER_MEMBER_ROW = 11;
    const META_UNIT_PER_MEMBER_COL = 4;
    return +SPREADSHEET.getSheetByName(SHEET.META)
      .getRange(META_UNIT_PER_MEMBER_ROW, META_UNIT_PER_MEMBER_COL)
      .getValue();
  }

  /** get roster sorting setting */
  export function sortRoster(): boolean {
    const META_SORT_ROSTER_ROW = 2;
    const META_SORT_ROSTER_COL = 5;
    return (
      SPREADSHEET.getSheetByName(SHEET.META)
        .getRange(META_SORT_ROSTER_ROW, META_SORT_ROSTER_COL)
        .getValue() === 'Yes'
    );
  }

  /** get Id of exclusions */
  export function exclusionId(): string {
    const META_EXCLUSSIONS_ROW = 7;
    const META_EXCLUSSIONS_COL = 1;
    return SPREADSHEET.getSheetByName(SHEET.META)
      .getRange(META_EXCLUSSIONS_ROW, META_EXCLUSSIONS_COL)
      .getValue();
  }

  /** get count of members in the roster */
  export function memberCount(): number {
    const META_GUILD_SIZE_ROW = 5;
    const META_GUILD_SIZE_COL = 12;
    return +SPREADSHEET.getSheetByName(SHEET.ROSTER)
      .getRange(META_GUILD_SIZE_ROW, META_GUILD_SIZE_COL)
      .getValue();
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
      const META_DATASOURCE_ROW = 14;
      const META_DATASOURCE_COL = 4;
      return SPREADSHEET.getSheetByName(SHEET.META)
        .getRange(META_DATASOURCE_ROW, META_DATASOURCE_COL)
        .getValue();
    }

    export function setGuildDataDate(): void {
      const META_GUILDDATADATE_ROW = 24;
      const META_GUILDDATADATE_COL = 1;
      SPREADSHEET.getSheetByName(SHEET.META)
        .getRange(META_GUILDDATADATE_ROW, META_GUILDDATADATE_COL)
        .setValue(new Date());
    }

    export function setUnitDefinitionsDate(): void {
      const META_UNITDEFINITIONSDATE_ROW = 26;
      const META_UNITDEFINITIONSDATE_COL = 1;
      SPREADSHEET.getSheetByName(SHEET.META)
        .getRange(META_UNITDEFINITIONSDATE_ROW, META_UNITDEFINITIONSDATE_COL)
        .setValue(new Date());
    }
  }

  /** SwgohGg related settings */
  export namespace SwgohGgApi {
    /** Get the guild id */
    export function guild(): number {
      const META_DOTGG_LINK_ROW = 2;
      const META_DOTGG_LINK_COL = 1;
      const guildLink = SPREADSHEET.getSheetByName(SHEET.META)
        .getRange(META_DOTGG_LINK_ROW, META_DOTGG_LINK_COL)
        .getValue() as string;
      // TODO: input check
      const parts = guildLink.split('/');
      const guildId = +parts[4];

      return guildId;
    }
  }

  /** SwgohHelp related settings */
  export namespace SwgohHelpApi {
    /** Get the SwgohHelp API username */
    export function username(): string {
      const META_DOTHELP_USERNAME_ROW = 16;
      const META_DOTHELP_USERNAME_COL = 1;
      return SPREADSHEET.getSheetByName(SHEET.META)
        .getRange(META_DOTHELP_USERNAME_ROW, META_DOTHELP_USERNAME_COL)
        .getValue();
    }

    /** Get the SwgohHelp API password */
    export function password(): string {
      const META_DOTHELP_PASSWORD_ROW = 18;
      const META_DOTHELP_PASSWORD_COL = 1;
      return SPREADSHEET.getSheetByName(SHEET.META)
        .getRange(META_DOTHELP_PASSWORD_ROW, META_DOTHELP_PASSWORD_COL)
        .getValue();
    }

    /** Get the guild member ally code */
    export function allyCode(): number {
      const META_DOTHELP_LINK_ROW = 20;
      const META_DOTHELP_LINK_COL = 1;
      return +SPREADSHEET.getSheetByName(SHEET.META)
        .getRange(META_DOTHELP_LINK_ROW, META_DOTHELP_LINK_COL)
        .getValue();
    }
  }

  /** discord related settings */
  export namespace discord {
    /** Get the webhook address */
    export function webhookUrl(): string {
      const DISCORD_WEBHOOKURL_ROW = 1;
      const DISCORD_WEBHOOKURL_COL = DISCORD_WEBHOOK_COL;
      return SPREADSHEET.getSheetByName(SHEET.DISCORD)
        .getRange(DISCORD_WEBHOOKURL_ROW, DISCORD_WEBHOOKURL_COL)
        .getValue();
    }

    /** Get the role to mention */
    export function roleId(): string {
      const DISCORD_ROLEID_ROW = 2;
      const DISCORD_ROLEID_COL = DISCORD_WEBHOOK_COL;
      return SPREADSHEET.getSheetByName(SHEET.DISCORD)
        .getRange(DISCORD_ROLEID_ROW, DISCORD_ROLEID_COL)
        .getValue();
    }

    /** Get the time and date when the TB started */
    export function startTime(): Date {
      const WEBHOOK_TB_START_ROW = 3;
      const WEBHOOK_TB_START_COL = DISCORD_WEBHOOK_COL;
      return SPREADSHEET.getSheetByName(SHEET.DISCORD)
        .getRange(WEBHOOK_TB_START_ROW, WEBHOOK_TB_START_COL)
        .getValue();
    }

    /** Get the number of hours in each phase */
    export function phaseDuration(): number {
      const columnOffset = isGeoDS_() ? 2 : isHothLS_() ? 0 : isHothDS_() ? 1 : NaN;
      const WEBHOOK_PHASE_HOURS_ROW = 4;
      const WEBHOOK_PHASE_HOURS_COL = DISCORD_WEBHOOK_COL + columnOffset;
      return +SPREADSHEET.getSheetByName(SHEET.DISCORD)
        .getRange(WEBHOOK_PHASE_HOURS_ROW, WEBHOOK_PHASE_HOURS_COL)
        .getValue();
    }

    /** Get the template for a webhooks */
    export function webhookTemplate(phase: TerritoryBattles.phaseIdx, row: number, defaultVal: string): string {
      const text = SPREADSHEET.getSheetByName(SHEET.DISCORD)
        .getRange(row, DISCORD_WEBHOOK_COL)
        .getValue() as string;

      return text.length > 0 ? text.replace('{0}', String(phase)) : defaultVal;
    }

    /** Get the Description for the phase */
    export function webhookDescription(phase: TerritoryBattles.phaseIdx): string {
      const columnOffset = isGeoDS_() ? 2 : isHothLS_() ? 0 : isHothDS_() ? 1 : NaN;
      const WEBHOOK_DESC_ROW = 9;
      const META_WEBHOOKDESC_ROW = WEBHOOK_DESC_ROW + phase - 1;
      const META_WEBHOOKDESC_COL = DISCORD_WEBHOOK_COL + columnOffset;
      return `\n\n${SPREADSHEET.getSheetByName(SHEET.DISCORD)
        .getRange(META_WEBHOOKDESC_ROW, META_WEBHOOKDESC_COL)
        .getValue()}`;
    }

    /** See if the platoons should be cleared */
    export function resetPlatoons(): boolean {
      const WEBHOOK_CLEAR_ROW = 15;
      const WEBHOOK_CLEAR_COL = DISCORD_WEBHOOK_COL;
      return (
        SPREADSHEET.getSheetByName(SHEET.DISCORD)
          .getRange(WEBHOOK_CLEAR_ROW, WEBHOOK_CLEAR_COL)
          .getValue() === 'Yes'
      );
    }

    /** See if the slot number should be displayed */
    export function displaySlots(): string {
      const WEBHOOK_DISPLAY_SLOT_ROW = 16;
      const WEBHOOK_DISPLAY_SLOT_COL = DISCORD_WEBHOOK_COL;
      return SPREADSHEET.getSheetByName(SHEET.DISCORD)
        .getRange(WEBHOOK_DISPLAY_SLOT_ROW, WEBHOOK_DISPLAY_SLOT_COL)
        .getValue() as string;
    }
  }
}
