// *******************************************
// ** API Functions to pull data from swgoh.gg
// *******************************************

// Get the guild ID
function get_guild_id_() {
  const MetaSWGOHLinkRow = 2
  const MetaSWGOHLinkCol = 1

  const guildLink = String(SpreadsheetApp.getActive().getSheetByName("Meta").getRange(MetaSWGOHLinkRow, MetaSWGOHLinkCol).getValue())
  const parts = guildLink.split("/")
  // TODO: input check
  const guildId = parts[4]

  return guildId
}

// Create Guild API Link
function get_guild_api_link_() {
  const link = `https://swgoh.gg/api/guild/${get_guild_id_()}/`
  // TODO: data check
  return link
}

function IsSWGOHggSource() {
  const value = String(SpreadsheetApp.getActive().getSheetByName("Meta").getRange(MetaDataSourceRow, MetaDataSourceCol).getValue())
  // TODO: centralize constants
  return value === "SWGOH.gg"
}

// Pull base Character data from SWGOH.gg
// @returns Array of Characters with [name, base_id, tags]
function getUnitsFromSWGoHgg_(link, errorMsg) {
  let json
  try {
    // const link = "https://swgoh.gg/api/characters/?format=json"
    const params = {
      // followRedirects: true,
      muteHttpExceptions: true
    }
    const response = UrlFetchApp.fetch(link, params)
    const responseObj = {
      getContentText: response.getContentText().split("\n"),
      getHeaders: response.getHeaders(),
      getResponseCode: response.getResponseCode()
    }
    if (response.getResponseCode() !== 200) {
      debugger;
    }
    json = JSON.parse(response.getContentText())
    // Logger.log(`Retrieving Hero Data from SWGOH.gg: ${json.length}`)
  } catch (e) {
    // TODO: centralize alerts
    const ui = SpreadsheetApp.getUi()
    ui.alert(errorMsg, e, ui.ButtonSet.OK)
  }

  return json || []
}

// Pull base Character data from SWGOH.gg
// @returns Array of Characters with [name, base_id, tags]
function GetHeroesFromSWGOHgg() {
  const json = getUnitsFromSWGoHgg_(
    "https://swgoh.gg/api/characters/?format=json",
    "Error when retreiving data from swgoh.gg API"
  )
  const mapping = unit => {
    const tags = [unit.alignment, unit.role, ...unit.categories].join(" ").toLowerCase()
    return [unit.name, unit.base_id, tags]
  }
  return json.map(mapping)
}

// Pull base Ship data from SWGOH.gg
// @returns Array of Characters with [name, base_id, tags]
function GetShipsFromSWGOHgg() {
  const json = getUnitsFromSWGoHgg_(
    "https://swgoh.gg/api/ships/?format=json",
    "Error when retreiving data from swgoh.gg API"
  )
  const mapping = unit => {
    const tags = [unit.alignment, unit.role, ...unit.categories].join(" ").toLowerCase()
    return [unit.name, unit.base_id, tags]
  }
  return json.map(mapping)
}

// Pull Guild data from SWGOH.gg
// @returns Array of Guild members and their character data
function GetGuildDataFromSWGOHgg() {
  const json = getUnitsFromSWGoHgg_(
    get_guild_api_link_(),
    "Error when retreiving data from swgoh.gg API"
  )
  const members = []
  json.players.forEach( member => {
    // const player_id = member.data.name  // TODO: duplicate names? member.data.url?
    // const player_id = member.data.url
    let unitArray = []
    member.units.forEach( e => {
      const unit = e.data
      const base_id = unit.base_id
      const q = []
      q["level"] = unit.level
      q["gear_level"] = unit.gear_level
      q["power"] = unit.power
      q["rarity"] = unit.rarity
      q["base_id"] = unit.base_id
      unitArray[base_id] = q
    })
    // members[player_id] = {
    //   name: member.data.name,
    //   level: member.data.level,
    //   gp: member.data.galactic_power,
    //   heroes_gp: member.data.character_galactic_power,
    //   ships_gp: member.data.ship_galactic_power,
    //   link: member.data.url,
    //   units: unitArray
    // }
    members.push({
      name: member.data.name,
      level: member.data.level,
      gp: member.data.galactic_power,
      heroes_gp: member.data.character_galactic_power,
      ships_gp: member.data.ship_galactic_power,
      link: member.data.url,
      units: unitArray
    })
  })

  return members;
}
