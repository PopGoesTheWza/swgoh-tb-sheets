/** SWGoH.help API client for Google Apps Script (GAS) */
export namespace  swgohhelpapi {
  export namespace exports {
    /**
     * SWGoH.help API client class.
     *
     * Available methods:
     * - client.fetchPlayer(payload: PlayerRequest): PlayerResponse
     * - client.fetchGuild(payload: GuildRequest): GuildResponse
     * - client.fetchUnits(payload: UnitsRequest): UnitsResponse
     * - client.fetchEvents(payload: EventsRequest): EventsResponse
     * - client.fetchBattles(payload: BattlesRequest): BattlesResponse
     * - client.fetchData(payload: <DataRequest>): <DataResponse>
     */
    export class Client {
      /** Sealed copy of the Client instance Settings */
      private readonly settings;
      /** URL to the login endpoint */
      private readonly signinUrl;
      /** URL to the player endpoint */
      private readonly playersUrl;
      /** URL to the guild endpoint */
      private readonly guildsUrl;
      /** URL to the units endpoint */
      private readonly unitsUrl;
      /** URL to the events endpoint */
      private readonly eventsUrl;
      /** URL to the battles endpoint */
      private readonly battlesUrl;
      /** URL to the data endpoint */
      private readonly dataUrl;
      /** URL to the zetas endpoint */
      private readonly zetasUrl;
      /** URL to the squads endpoint */
      private readonly squadsUrl;
      /** URL to the roster endpoint */
      private readonly rosterUrl;
      /** token key is a SHA256 digest of the credentials used to access the API */
      private readonly tokenId;
      /** Creates a SWGoH.help API client. */
      constructor(settings: Settings);
      /**
       * Attempts to log into the SWGoH.help API.
       * returns A SWGoH.help API session token
       */
      login(): string;
      /** Fetch Player data */
      fetchPlayer(payload: PlayerRequest): PlayerResponse[];
      /** Fetch Guild data */
      fetchGuild(payload: GuildRequest): GuildResponse[];
      /** Fetch Units data */
      fetchUnits(payload: UnitsRequest): UnitsResponse;
      /** Fetch Events data */
      fetchEvents(payload: EventsRequest): EventsResponse;
      /** Fetch Battles data */
      fetchBattles(payload: BattlesRequest): BattlesResponse;
      /**
       * Fetch Data data
       * returns The structure of the response depends on the collection used.
       */
      fetchData(payload: DataRequest): any;
      /** Fetch Zetas data */
      fetchZetas(payload: any): any;
      /** Fetch Squads data */
      fetchSquads(payload: any): any;
      /** Fetch Roster data */
      fetchRoster(payload: any): any;
      protected fetchAPI<T>(url: string, payload: any): T;
      /**
       * Retrieve the API session token from Google CacheService.
       * The key is a SHA256 digest of the credentials used to access the API.
       * Caching period is one hour.
       *
       * If there is no valid API session token, one is created by invoking the login() method.
       */
      private getToken;
      /**
       * Store the API session token into Google CacheService.
       * The key is a SHA256 digest of the credentials used to access the API.
       * Caching period is one hour.
       */
      private setToken;
    }
    /** Interfaces and Types declarations */
    /** Settings for creating a new Client */
    export interface Settings {
      /** Registered username for swgoh.help API */
      readonly username: string;
      /** Registered password for swgoh.help API */
      readonly password: string;
      /** currently unused */
      readonly client_id?: string;
      /** currently unused */
      readonly client_secret?: string;
      /** default to 'https' */
      readonly protocol?: string;
      /** default to 'api.swgoh.help' */
      readonly host?: string;
      /** default to '' (80) */
      readonly port?: string;
    }
  }
  /** Supported languages for localized data */
  export enum Languages {
    /** Chinese (Simplified) */
    chs_cn = "chs_cn",
    /** Chinese (Traditional) */
    cht_cn = "cht_cn",
    /** English (USA) */
    eng_us = "eng_us",
    /** French (France) */
    fre_fr = "fre_fr",
    /** German (Germany) */
    ger_de = "ger_de",
    /** Indonesian (Indonesia) */
    ind_id = "ind_id",
    /** Italian (italy) */
    ita_it = "ita_it",
    /** Japanese (Japan) */
    jpn_jp = "jpn_jp",
    /** Korean (South Korea) */
    kor_kr = "kor_kr",
    /** Portugese (Brazil) */
    por_br = "por_br",
    /** Russian (Russia) */
    rus_ru = "rus_ru",
    /** Spanish? (???) */
    spa_xm = "spa_xm",
    /** Thai (Thailand) */
    tha_th = "tha_th",
    /** Turkish (Turkey) */
    tur_tr = "tur_tr"
  }
  export namespace exports {
    /** Sub-types from Responses */
    /** Player stats */
    type PlayerStats = {
      nameKey: string;
      value: number;
      /** can be used as key */
      index: number;
    };
    /** Bare Units properties */
    type BaseUnit = {
      id: string;
      defId: string;
    };
    /** Equipped properties for Roster units */
    type Equipped = {
      equipmentId: string;
      slot: number;
    };
    /** Skills properties for Roster units */
    type Skills = {
      id: string;
      tier: number;
      isZeta: boolean;
    };
    /** Bare Mod properties */
    type BaseMod = {
      id: string;
      set: number;
      level: number;
      pips: number;
      tier: number;
    };
    /** Mod properties */
    interface Mod extends BaseMod {
      slot: number;
      primaryBonusType: number;
      primaryBonusValue: number;
      secondaryType_1: number;
      secondaryValue_1: number;
      secondaryRoll_1: number;
      secondaryType_2: number;
      secondaryValue_2: number;
      secondaryRoll_2: number;
      secondaryType_3: number;
      secondaryValue_3: number;
      secondaryRoll_3: number;
      secondaryType_4: number;
      secondaryValue_4: number;
      secondaryRoll_4: number;
    }
    /** Mod properties */
    interface ModInstance extends BaseMod {
      stat: [number, (string | number), number][];
    }
    /** Crew properties */
    type Crew = {
      unitId: string;
      slot: number;
      skillReferenceList: string[];
      cp: number;
      gp: number;
    };

  }
  export enum COMBAT_TYPE {
    HERO = 1,
    SHIP = 2
  }
  export namespace exports {
    /** Roster properties */
    interface Roster extends BaseUnit {
      nameKey: string;
      rarity: number;
      level: number;
      xp: number;
      gear: number;
      equipped: Equipped[];
      combatType: COMBAT_TYPE;
      skills: Skills[];
      mods: Mod[];
      crew: Crew[];
      gp: number;
    }
    /** Arena properties */
    type Arena = {
      rank: number;
      squad: BaseUnit[];
    };
    /** Common properties for all Request interfaces */
    type CommonRequest = {
      /**
       * Optional language to return translated names
       * If no language specified, no translations will be applied
       */
      language?: Languages;
      /** Optionally return enumerated items as their string equivalents */
      enums?: boolean;
    };
    /** Request interface for /swgoh/player endpoint */
    export interface PlayerRequest extends CommonRequest {
      /** This field is mandatory */
      allycodes: number | number[];
      project?: PlayerOptions;
    }
    /** Optional projection of PlayerResponse properties (first layer) you want returned */
    type PlayerOptions = {
      allyCode?: boolean;
      name?: boolean;
      level?: boolean;
      guildName?: boolean;
      gp?: boolean;
      gpChar?: boolean;
      gpShip?: boolean;
      updated?: boolean;
      stats?: boolean;
      roster?: any;  // boolean;
      arena?: boolean;
    };
    /** Response from PlayerRequest */
    export interface PlayerResponse {
      allyCode?: number;
      name?: string;
      level?: number;
      guildname?: string;
      gp?: number;
      gpChar?: number;
      gpShip?: number;
      updated?: number;
      guildRefId?: string;
      stats?: PlayerStats[];
      roster?: Roster[];
      arena?: {
          char: Arena[];
          ship: Arena[];
      };
    }
    /** Optional projection of GuildResponse properties (first layer) you want returned */
    type GuildOptions = {
      id?: boolean;
      name?: boolean;
      desc?: boolean;
      members?: boolean;
      status?: boolean;
      required?: boolean;
      bannerColor?: boolean;
      banerLogo?: boolean;
      message?: boolean;
      gp?: boolean;
      raid?: boolean;
      roster?: any;
      updated?: boolean;
    };
    /** Request interface for /swgoh/guild endpoint */
    export interface GuildRequest extends CommonRequest {
      allycode: number;
      /** Optionally replace guild roster with full array of player profiles */
      roster?: boolean;
      /** (in conjunction with roster) Optionally replace guild roster with guild-wide units-report */
      units?: boolean;
      /** (in conjunction with units) Optionally include unit mods in units-report */
      mods?: boolean;
      project?: GuildOptions;
    }
    /** Response from GuildRequest */
    export interface GuildResponse {
      updated?: number;
      id: string;
      roster?: PlayerResponse[] | UnitsResponse;
      name?: string;
      desc?: string;
      members?: number;
      status?: number;
      required?: number;
      bannerColor?: string;
      bannerLogo?: string;
      message?: string;
      gp?: number;
      raid?: {
          aat: string;
          rancor: string;
          sith_raid: string;
      };
    }
    /** Optional projection of UnitsdResponse properties (first layer) you want returned */
    type UnitsOptions = {
      player?: boolean;
      allyCode?: boolean;
      starLevel?: boolean;
      level?: boolean;
      gearLevel?: boolean;
      gear?: boolean;
      zetas?: boolean;
      type?: boolean;
      mods?: boolean;
      gp?: boolean;
      updated?: boolean;
    };
    /** Request interface for /swgoh/units endpoint */
    export interface UnitsRequest extends CommonRequest {
      allycodes: number | number[];
      /** Optionally include unit-mods in report */
      mods?: boolean;
      project?: UnitsOptions;
    }
    /** Response from UnitsRequest */
    export interface UnitsResponse {
      [key: string]: UnitsInstance[];
    }
    /** Optional projection of EventsdResponse properties (first layer) you want returned */
    type EventsOptions = {
      id?: boolean;
      priority?: boolean;
      nameKey?: boolean;
      summaryKey?: boolean;
      descKey?: boolean;
      instances?: boolean;
      squadType?: boolean;
      defensiveSquadType?: boolean;
    };
    /** Request interface for /swgoh/events endpoint */
    export interface EventsRequest extends CommonRequest {
      project?: EventsOptions;
    }
    /** Response from EventsRequest */
    export interface EventsResponse {
      updated: number;
      id?: string;
      priority?: number;
      nameKey?: string;
      summaryKey?: string;
      descKey?: string;
      instances?: EventsInstance[];
      squadType?: number;
      defensiveSquad?: number;
    }
    /** Optional projection of BattlesResponse properties (first layer) you want returned */
    type BattlesOptions = {
      id?: boolean;
      nameKey?: boolean;
      descriptionKey?: boolean;
      campaignType?: boolean;
      campaignMapList?: boolean;
    };
    /** Request interface for /swgoh/battles endpoint */
    export interface BattlesRequest extends CommonRequest {
      project?: BattlesOptions;
    }
    /** Response from BattlesRequest */
    export interface BattlesResponse {
      updated: number;
      id?: string;
      nameKey?: string;
      descriptionKey?: string;
      campaignType?: number;
      campaignMapList?: CampaignMap[];
    }
  }
  /** Supported collections with /swgoh/data endpoint */
  export enum Collections {
    abilityList = "abilityList",
    battleEnvironmentsList = "battleEnvironmentsList",
    battleTargetingRuleList = "battleTargetingRuleList",
    categoryList = "categoryList",
    challengeList = "challengeList",
    challengeStyleList = "challengeStyleList",
    effectList = "effectList",
    environmentCollectionList = "environmentCollectionList",
    equipmentList = "equipmentList",
    eventSamplingList = "eventSamplingList",
    guildExchangeItemList = "guildExchangeItemList",
    guildRaidList = "guildRaidList",
    helpEntryList = "helpEntryList",
    materialList = "materialList",
    playerTitleList = "playerTitleList",
    powerUpBundleList = "powerUpBundleList",
    raidConfigList = "raidConfigList",
    recipeList = "recipeList",
    requirementList = "requirementList",
    skillList = "skillList",
    starterGuildList = "starterGuildList",
    statModList = "statModList",
    statModSetList = "statModSetList",
    statProgressionList = "statProgressionList",
    tableList = "tableList",
    targetingSetList = "targetingSetList",
    territoryBattleDefinitionList = "territoryBattleDefinitionList",
    territoryWarDefinitionList = "territoryWarDefinitionList",
    unitsList = "unitsList",
    unlockAnnouncementDefinitionList = "unlockAnnouncementDefinitionList",
    warDefinitionList = "warDefinitionList",
    xpTableList = "xpTableList"
  }
  export namespace exports {
    /** Request interface for /swgoh/data endpoint */
    export interface DataRequest extends CommonRequest {
      collection: Collections;
      match?: object;
      project?: object;
    }
    type AbilityListMatch = {};
    type AbilityListOptions = {};
    /** Request interface for /swgoh/data endpoint */
    export interface AbilityListRequest extends DataRequest {
      match?: AbilityListMatch;
      project?: AbilityListOptions;
    }
    export interface AbilityListResponse {
      _id: string;
      id?: string;
      nameKey?: string;
      descKey?: string;
      prefabName?: string;
      triggerConditionList?: {
          conditionType: number;
          conditionValue: string;
      }[];
      stackingLineOverride?: string;
      tierList?: Tier[];
      cooldown?: number;
      icon?: string;
      applyTypeTooltipKey?: string;
      descriptiveTagList?: {
          tag: string;
      }[];
      effectReferenceList?: EffectReference[];
      confirmationMessage?: {
          titleKey: string;
          descKey: string;
      };
      buttonLocation?: number;
      shortDescKey?: string;
      abilityType?: number;
      detailLocation?: number;
      allyTargetingRuleId?: string;
      useAsReinforcementDesc?: boolean;
      preferredAllyTargetingRuleId?: string;
      interactsWithTagList?: {
          tag: string;
      }[];
      subIcon?: string;
      aiParams?: AiParams;
      cooldownType?: number;
      alwaysDisplayInBattleUi?: boolean;
      highlightWhenReadyInBattleUi?: boolean;
      hideCooldownDescription?: boolean;
    }
    /** Response from UnitsRequest */
    type UnitsInstance = {
      updated?: number;
      player?: string;
      allyCode?: number;
      gp?: number;
      starLevel?: number;
      level?: number;
      gearLevel?: number;
      gear?: string[];
      zetas?: string[];
      type?: number;
      mods?: ModInstance[];
    };
    /** Event instance properties */
    export type EventsInstance = {
      startTime: number;
      endTime: number;
      displayStartTime: number;
      displayEndTime: number;
      rewardTime: number;
    };
    type BaseItem = {
      id: string;
      type: number;
      weight: number;
      minQuantity: number;
      maxQuantity: number;
      rarity: number;
      statMod: undefined;
    };
    type EnemyUnit = {
      baseEnemyItem: BaseItem;
      enemyLevel: number;
      enemyTier: number;
      threatLevel: number;
      thumbnailName: string;
      prefabName: string;
      displayedEnemy: boolean;
      unitClass: number;
    };
    type CampaignNodeMission = {
      id: string;
      nameKey: string;
      descKey: string;
      combatType: number;
      rewardPreviewList: BaseItem[];
      shortNameKey: string;
      groupNameKey: string;
      firstCompleteRewardPreviewList: BaseItem[];
      enemyUnitPreviewList: EnemyUnit[];
    };
    type CampaignNode = {
      id: string;
      nameKey: string;
      campaignNodeMissionList: CampaignNodeMission[];
    };
    type CampaignNodeDifficultyGroup = {
      campaignNodeDifficulty: number;
      campaignNodeList: CampaignNode;
      unlockRequirementLocalizationKey: string;
    };
    /** CampaignMap properties */
    type CampaignMap = {
      id: string;
      campaignNodeDifficultyGroupList: CampaignNodeDifficultyGroup[];
      entryOwnershipRequirementList: {
          id: string;
      }[];
      unlockRequirementLocalizationKey: string;
    };
    export type EffectReference = {
      id: string;
      contextIndexList: number[];
      maxBonusMove: number;
    };
    export type Tier = {
      descKey: string;
      upgradeDescKey: string;
      cooldownMaxOverride: number;
      effectReferenceList: EffectReference[];
      interactsWithTagList: {
          tag: string;
      }[];
    };
    export type AiParams = {
      preferredTargetAllyTargetingRuleId: string;
      preferredEnemyTargetingRuleId: string;
      requireEnemyPreferredTargets: boolean;
      requireAllyTargets: boolean;
    };
  }
}
