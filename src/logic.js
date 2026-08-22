/**
 * ============================================================================
 * Pure Business Logic Layer (logic.js)
 * Laburnum CC Availability Tracker
 *
 * RULE: Functions here MUST NOT call SpreadsheetApp, DriveApp, Ui, or any
 * Google Apps Script service. They take plain JS primitives / arrays /
 * objects and return transformed data, so they run unchanged in Node (Jest).
 *
 * This file is part of the clasp project AND is required by the Jest tests:
 *   - Apps Script: functions are hoisted as globals when pushed to Google.
 *   - Node/Jest:   the guarded module.exports at the bottom enables
 *                  `const { normalizePhone } = require('./logic.js')`.
 * ============================================================================
 */

/**
 * Normalizes phone numbers to E164 format (+614...).
 *   "0400 123 456" -> "+61400123456"
 *   "0412345678"   -> "+61412345678"
 *   "412345678"    -> "+61412345678"
 *   "+61412345678" -> "+61412345678"
 */
function normalizePhone(phone) {
  if (!phone) return "";
  let clean = String(phone).replace(/[^\d]/g, '');

  if (clean.startsWith('04') && clean.length === 10) {
    clean = '61' + clean.slice(1);
  } else if (clean.length === 9 && clean.startsWith('4')) {
    clean = '61' + clean;
  }
  return '+' + clean;
}

/**
 * Returns the household for an input phone: every player whose number appears
 * in ANY of the 4 phone columns (Phone, Phone2, Phone3, Phone4).
 *
 * @param {Array<Object>} playersRows Row objects with Phone/Phone2/Phone3/Phone4.
 * @param {string} inputPhone Raw user input (normalized here).
 * @returns {Array<Object>} Matching player rows.
 */
function findHouseholdPlayers(playersRows, inputPhone) {
  const targetPhone = normalizePhone(inputPhone);
  if (!targetPhone) return [];

  return playersRows.filter(player => {
    const phones = [
      normalizePhone(player.Phone),
      normalizePhone(player.Phone2),
      normalizePhone(player.Phone3),
      normalizePhone(player.Phone4)
    ];
    return phones.includes(targetPhone);
  });
}

/**
 * Moves a 3-cell block [ProfileID, Name, Notes] from one list to another,
 * compacting the source list upward. Pure array manipulation.
 *
 * @param {Array<Array>} sourceList
 * @param {Array<Array>} destList
 * @param {string} targetProfileId
 * @param {string|null} newNotes Optional note update applied on move.
 * @returns {{ sourceList: Array, destList: Array, moved: boolean }}
 */
function moveItemBetweenLists(sourceList, destList, targetProfileId, newNotes = null) {
  const sourceIndex = sourceList.findIndex(row => String(row[0]).trim() === targetProfileId);
  if (sourceIndex === -1) return { sourceList, destList, moved: false };

  // Copy item
  const movedItem = [...sourceList[sourceIndex]];
  if (newNotes !== null) {
    movedItem[2] = newNotes; // Update Notes column
  }

  // Remove from source (compacting)
  const newSource = sourceList.filter((_, idx) => idx !== sourceIndex);

  // Append to destination
  const newDest = [...destList, movedItem];

  return { sourceList: newSource, destList: newDest, moved: true };
}

/**
 * Appends a junior-class suffix tag to the player display name.
 *   Adult       → "First Last"
 *   U18         → "First Last (U18)"
 *   U16 / U16_Y2 / U16_Y1  → "First Last (U16)"
 *   U14         → "First Last (U14)"
 */
function formatNameWithJuniorTag(fullName, juniorLevel) {
  if (juniorLevel === "U18") return fullName + " (U18)";
  if (juniorLevel === "U16" || juniorLevel === "U16_Y2" || juniorLevel === "U16_Y1") return fullName + " (U16)";
  if (juniorLevel === "U14") return fullName + " (U14)";
  return fullName;
}

/**
 * Picks the Preferred Name over First Name if it's genuinely a nickname
 * (shorter than the legal first name and not a First+Last copy).
 */
function pickFirstName(firstName, preferredName) {
  if (!preferredName) return firstName;
  if (preferredName.length < firstName.length) return preferredName;
  return firstName;
}

/**
 * Default team configurations for Laburnum CC.
 */
var DEFAULT_TEAM_CONFIGS = [
  { internalName: "1st XI", competition: "BHRDCA Senior Competition", playCricketName: "Laburnum - 1st XI" },
  { internalName: "2nd XI", competition: "BHRDCA Senior Competition", playCricketName: "Laburnum - 2nd XI" },
  { internalName: "3rd XI", competition: "BHRDCA Senior Competition", playCricketName: "Laburnum - 3rd XI" },
  { internalName: "4th XI", competition: "BHRDCA Senior Competition", playCricketName: "Laburnum - 4th XI" },
  { internalName: "5th XI", competition: "BHRDCA Senior Competition", playCricketName: "Laburnum - 5th XI" }
];

/**
 * Robust RFC-4180 compliant CSV parser.
 * Handles quoted fields, embedded quotes (""), commas, and CRLF/LF newlines.
 *
 * @param {string} text Raw CSV text.
 * @returns {Array<Array<string>>} 2D array of string cells.
 */
function parseCsvString(text) {
  if (!text || typeof text !== "string") return [];
  var rows = [];
  var row = [];
  var cur = "";
  var inQuotes = false;

  for (var i = 0; i < text.length; i++) {
    var c = text[i];
    if (c === '"') {
      if (inQuotes && text[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === "," && !inQuotes) {
      row.push(cur.trim());
      cur = "";
    } else if ((c === "\r" || c === "\n") && !inQuotes) {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(cur.trim());
      if (row.length > 1 || (row.length === 1 && row[0] !== "")) {
        rows.push(row);
      }
      row = [];
      cur = "";
    } else {
      cur += c;
    }
  }

  if (cur || row.length > 0) {
    row.push(cur.trim());
    if (row.length > 1 || (row.length === 1 && row[0] !== "")) {
      rows.push(row);
    }
  }
  return rows;
}

/**
 * Normalizes team internal name to standard column prefix (e.g. "1st XI" -> "1st").
 */
function formatTeamPrefix(name) {
  if (!name) return "";
  var clean = String(name).trim();
  clean = clean.replace(/\s*XI\s*$/i, "");
  return clean;
}

/**
 * Builds the standard header row for the Fixtures tab based on team configs.
 * Game Date is the primary key (Column 1), followed by per-team Round, Format, Opponent, Venue.
 * Format: ["Game Date", "1st Round", "1st Format", "1st Opponent", "1st Venue", "2nd Round", "2nd Format", ...]
 *
 * @param {Array<Object>} teamConfigs
 * @returns {Array<string>}
 */
function buildFixtureHeaders(teamConfigs) {
  var configs = (teamConfigs && teamConfigs.length > 0) ? teamConfigs : DEFAULT_TEAM_CONFIGS;
  var headers = ["Game Date"];
  configs.forEach(function(c) {
    var prefix = formatTeamPrefix(c.internalName);
    headers.push(prefix + " Round", prefix + " Format", prefix + " Opponent", prefix + " Venue");
  });
  return headers;
}

/**
 * Extracts team configurations from the 2D values of the Config tab.
 *
 * @param {Array<Array>} configValues 2D array from Config sheet.
 * @returns {Array<{internalName: string, competition: string, playCricketName: string}>}
 */
function readTeamConfigsFromSheet(configValues) {
  if (!configValues || !Array.isArray(configValues) || configValues.length === 0) {
    return DEFAULT_TEAM_CONFIGS;
  }

  var headerRowIdx = -1;
  var internalCol = -1;
  var compCol = -1;
  var playCricketCol = -1;

  for (var r = 0; r < configValues.length; r++) {
    var row = configValues[r];
    for (var c = 0; c < row.length; c++) {
      var val = String(row[c] || "").trim().toLowerCase();
      if (val === "internal team name" || val === "internal name" || val === "team name") {
        headerRowIdx = r;
        internalCol = c;
      } else if (val === "competition" || val === "comp") {
        compCol = c;
      } else if (val === "play cricket team name" || val === "playhq team name" || val === "play cricket name") {
        playCricketCol = c;
      }
    }
    if (headerRowIdx !== -1 && internalCol !== -1 && playCricketCol !== -1) {
      break;
    }
  }

  if (headerRowIdx === -1) {
    return DEFAULT_TEAM_CONFIGS;
  }

  var teams = [];
  for (var r = headerRowIdx + 1; r < configValues.length; r++) {
    var row = configValues[r];
    var internal = internalCol !== -1 && row[internalCol] ? String(row[internalCol]).trim() : "";
    var playCricket = playCricketCol !== -1 && row[playCricketCol] ? String(row[playCricketCol]).trim() : "";
    var comp = compCol !== -1 && row[compCol] ? String(row[compCol]).trim() : "";

    if (!internal && !playCricket) {
      if (teams.length > 0) break;
      else continue;
    }
    if (internal.toUpperCase().includes("TEMPLATE") || internal.toUpperCase().includes("WHATSAPP")) {
      break;
    }

    teams.push({
      internalName: internal,
      competition: comp,
      playCricketName: playCricket
    });
  }

  return teams.length > 0 ? teams : DEFAULT_TEAM_CONFIGS;
}

/**
 * Extracts message templates from the 2D values of the Config tab.
 *
 * @param {Array<Array>} configValues 2D array from Config sheet.
 * @returns {Object<string, string>}
 */
function readTemplatesFromSheet(configValues) {
  var templates = {
    "Availability Callout": "🏏 *LCC ROUND AVAILABILITY* 🏏\nPlease submit your availability for this round: {url}",
    "Wall of Shame": "🚨 *WALL OF SHAME* 🚨\nThe following players have not yet entered their availability: {players}",
    "Selection Announcement": "🏏 *LABURNUM CC TEAMS - {date}* 🏏"
  };

  if (!configValues || !Array.isArray(configValues) || configValues.length === 0) {
    return templates;
  }

  var headerRowIdx = -1;
  var nameCol = -1;
  var textCol = -1;

  for (var r = 0; r < configValues.length; r++) {
    var row = configValues[r];
    for (var c = 0; c < row.length; c++) {
      var val = String(row[c] || "").trim().toLowerCase();
      if (val === "template name" || val === "template") {
        headerRowIdx = r;
        nameCol = c;
      } else if (val === "template text" || val === "template content" || val === "text" || val === "content") {
        textCol = c;
      }
    }
    if (headerRowIdx !== -1 && nameCol !== -1 && textCol !== -1) {
      break;
    }
  }

  if (headerRowIdx !== -1) {
    for (var r = headerRowIdx + 1; r < configValues.length; r++) {
      var row = configValues[r];
      var name = nameCol !== -1 && row[nameCol] ? String(row[nameCol]).trim() : "";
      var text = textCol !== -1 && row[textCol] ? String(row[textCol]).trim() : "";
      if (name && text) {
        templates[name] = text;
      }
    }
  }

  return templates;
}

/**
 * Normalizes any date representation (JS Date object, YYYY-MM-DD, DD/MM/YYYY, multi-date string)
 * into a strict YYYY-MM-DD string.
 *
 * @param {string|Date} val Raw date value.
 * @returns {string} Normalized YYYY-MM-DD string or empty string.
 */
function normalizeDateToYYYYMMDD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    var y = val.getFullYear();
    var m = String(val.getMonth() + 1);
    if (m.length < 2) m = "0" + m;
    var d = String(val.getDate());
    if (d.length < 2) d = "0" + d;
    return y + "-" + m + "-" + d;
  }

  var str = String(val).trim();
  if (!str) return "";

  // If multiple dates (e.g. "04/10/2025, 11/10/2025"), take the first match day
  if (str.indexOf(",") > -1) {
    str = str.split(",")[0].trim();
  }

  // YYYY-MM-DD format
  var ymdMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (ymdMatch) {
    var y = ymdMatch[1];
    var m = ymdMatch[2].length < 2 ? "0" + ymdMatch[2] : ymdMatch[2];
    var d = ymdMatch[3].length < 2 ? "0" + ymdMatch[3] : ymdMatch[3];
    return y + "-" + m + "-" + d;
  }

  // DD/MM/YYYY format
  var dmyMatch = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (dmyMatch) {
    var d = dmyMatch[1].length < 2 ? "0" + dmyMatch[1] : dmyMatch[1];
    var m = dmyMatch[2].length < 2 ? "0" + dmyMatch[2] : dmyMatch[2];
    var y = dmyMatch[3];
    return y + "-" + m + "-" + d;
  }

  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    var y = parsed.getFullYear();
    var m = String(parsed.getMonth() + 1);
    if (m.length < 2) m = "0" + m;
    var d = String(parsed.getDate());
    if (d.length < 2) d = "0" + d;
    return y + "-" + m + "-" + d;
  }

  return str;
}

/**
 * Parses PlayHQ fixture CSV data and matches rows to configured teams.
 * Supports both Home and Away fixture CSV files.
 *
 * @param {string|Array<Array>} csvContent Raw CSV text or parsed 2D array.
 * @param {Array<Object>} teamConfigs Array of team configs.
 * @returns {Array<Object>} Parsed game objects.
 */
function parseFixtureCsv(csvContent, teamConfigs) {
  var configs = (teamConfigs && teamConfigs.length > 0) ? teamConfigs : DEFAULT_TEAM_CONFIGS;
  var rows = typeof csvContent === "string" ? parseCsvString(csvContent) : csvContent;
  if (!rows || rows.length < 2) return [];

  var headers = rows[0];
  var colIdx = {};
  headers.forEach(function(h, i) {
    colIdx[String(h).trim().toLowerCase()] = i;
  });

  var getVal = function(row, key) {
    var idx = colIdx[key.toLowerCase()];
    return idx !== undefined && row[idx] !== undefined ? String(row[idx]).trim() : "";
  };

  var parsedGames = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var homeTeam = getVal(row, "Home Team");
    var awayTeam = getVal(row, "Away Team");
    var comp = getVal(row, "Competition");
    var gameType = getVal(row, "Game Type");
    var round = getVal(row, "Round");
    var venue = getVal(row, "Venue") || getVal(row, "Playing Surface");
    var startDate = getVal(row, "Game Start Date");
    var rawDate = getVal(row, "Game Date");
    var normalizedDate = normalizeDateToYYYYMMDD(startDate || rawDate);

    if (!normalizedDate) continue;

    for (var cIdx = 0; cIdx < configs.length; cIdx++) {
      var c = configs[cIdx];
      var targetName = c.playCricketName.trim().toLowerCase();
      var compTarget = (c.competition || "").trim().toLowerCase();
      var compMatches = !compTarget || !comp || comp.toLowerCase().indexOf(compTarget) > -1 || compTarget.indexOf(comp.toLowerCase()) > -1;

      if (homeTeam.toLowerCase() === targetName && compMatches) {
        parsedGames.push({
          internalTeamName: c.internalName,
          playCricketName: c.playCricketName,
          competition: comp,
          startDate: normalizedDate,
          rawDate: rawDate,
          round: round,
          opponent: awayTeam,
          venue: venue,
          format: gameType,
          side: "Home"
        });
      } else if (awayTeam.toLowerCase() === targetName && compMatches) {
        parsedGames.push({
          internalTeamName: c.internalName,
          playCricketName: c.playCricketName,
          competition: comp,
          startDate: normalizedDate,
          rawDate: rawDate,
          round: round,
          opponent: homeTeam,
          venue: venue,
          format: gameType,
          side: "Away"
        });
      }
    }
  }

  return parsedGames;
}

/**
 * Merges parsed fixture games into the 2D matrix of the Fixtures tab.
 * Idempotent: Can be run multiple times or in phases (Home CSV then Away CSV).
 * Guarantees a SINGLE ROW per match date/round, combining all team matches together.
 *
 * @param {Array<Array>} existingMatrix Current 2D values from Fixtures sheet.
 * @param {Array<Object>} parsedGames Game objects from parseFixtureCsv.
 * @param {Array<Object>} teamConfigs Team configuration array.
 * @returns {Array<Array>} Complete 2D array [headers, ...dataRows] ready for setValues().
 */
function mergeFixturesIntoMatrix(existingMatrix, parsedGames, teamConfigs) {
  var configs = (teamConfigs && teamConfigs.length > 0) ? teamConfigs : DEFAULT_TEAM_CONFIGS;
  var headers = buildFixtureHeaders(configs);

  var headerColMap = {};
  headers.forEach(function(h, i) {
    headerColMap[h.toLowerCase()] = i;
  });

  var rowMap = new Map();

  // Load existing matrix if present, strictly normalizing date keys
  if (existingMatrix && existingMatrix.length > 1) {
    var existingHeaders = existingMatrix[0];
    var exColMap = {};
    existingHeaders.forEach(function(h, i) {
      exColMap[String(h).trim().toLowerCase()] = i;
    });
    var dateIdx = exColMap["game date"];

    if (dateIdx !== undefined) {
      for (var r = 1; r < existingMatrix.length; r++) {
        var row = existingMatrix[r];
        var dateVal = normalizeDateToYYYYMMDD(row[dateIdx]);
        if (!dateVal) continue;

        var newRow = new Array(headers.length).fill("");
        newRow[headerColMap["game date"]] = dateVal;

        existingHeaders.forEach(function(eh, ei) {
          var targetCol = headerColMap[String(eh).trim().toLowerCase()];
          if (targetCol !== undefined && targetCol !== headerColMap["game date"] && row[ei] !== undefined) {
            newRow[targetCol] = row[ei];
          }
        });
        rowMap.set(dateVal, newRow);
      }
    }
  }

  // Merge parsed games into date rows
  parsedGames.forEach(function(game) {
    var dateKey = normalizeDateToYYYYMMDD(game.startDate);
    if (!dateKey) return;

    var row = rowMap.get(dateKey);
    if (!row) {
      row = new Array(headers.length).fill("");
      row[headerColMap["game date"]] = dateKey;
      rowMap.set(dateKey, row);
    }

    var prefix = formatTeamPrefix(game.internalTeamName).toLowerCase();
    var rndCol = headerColMap[prefix + " round"];
    var fmtCol = headerColMap[prefix + " format"];
    var oppCol = headerColMap[prefix + " opponent"];
    var venCol = headerColMap[prefix + " venue"];

    if (rndCol !== undefined) row[rndCol] = game.round;
    if (fmtCol !== undefined) row[fmtCol] = game.format;
    if (oppCol !== undefined) row[oppCol] = game.opponent;
    if (venCol !== undefined) row[venCol] = game.venue;
  });

  // Sort rows chronologically by date
  var sortedDates = Array.from(rowMap.keys()).sort();
  var resultRows = [headers];
  sortedDates.forEach(function(d) {
    var row = rowMap.get(d);
    // Check that row has at least one opponent or round data
    var hasData = false;
    for (var c = 1; c < row.length; c++) {
      if (row[c]) { hasData = true; break; }
    }
    if (hasData) {
      resultRows.push(row);
    }
  });

  return resultRows;
}


/**
 * Simulates player availability for a round based on format:
 * - 1-Day format: 90% Available, 10% Unavailable
 * - 2-Day format: 85% Available (both days), 5% Day 1 only, 5% Day 2 only, 5% Unavailable
 * 
 * @param {Array<Object>} players - List of player objects { profileId, fullName, globalStatus }
 * @param {string} matchFormat - 'One Day', 'Two Day', 'T20', etc.
 * @param {Function} [rng] - Optional custom random function (returns [0, 1)) for deterministic testing
 * @returns {Array<Object>} Array of simulated availability objects { profileId, fullName, response, notes }
 */
function simulatePlayerAvailability(players, matchFormat, rng) {
  var random = rng || Math.random;
  var formatStr = String(matchFormat || "").toLowerCase();
  var isOneDay = formatStr.indexOf("one") > -1 || formatStr.indexOf("t20") > -1 || formatStr.indexOf("1-day") > -1;

  return players.map(function(player) {
    var profileId = String(player.profileId || player.ProfileID || player.id || "").trim();
    var name = player.fullName || player.FullName || player.name || "";
    var status = String(player.globalStatus || player.GlobalStatus || "Active").trim();

    // Respect Injured or Inactive status
    if (status === "Inactive") {
      return {
        profileId: profileId,
        fullName: name,
        response: "Unavailable",
        notes: "Inactive in Master"
      };
    }
    if (status === "Injured") {
      return {
        profileId: profileId,
        fullName: name,
        response: "Unavailable",
        notes: "Injured"
      };
    }

    var r = random();

    if (isOneDay) {
      // 90% Available, 10% Unavailable
      if (r < 0.90) {
        return {
          profileId: profileId,
          fullName: name,
          response: "Available",
          notes: ""
        };
      } else {
        return {
          profileId: profileId,
          fullName: name,
          response: "Unavailable",
          notes: "Unavailable for Round"
        };
      }
    } else {
      // 2-Day: 85% Available, 5% Day 1 only, 5% Day 2 only, 5% Unavailable
      if (r < 0.85) {
        return {
          profileId: profileId,
          fullName: name,
          response: "Available",
          notes: ""
        };
      } else if (r < 0.90) {
        return {
          profileId: profileId,
          fullName: name,
          response: "Available",
          notes: "Day 1 Only"
        };
      } else if (r < 0.95) {
        return {
          profileId: profileId,
          fullName: name,
          response: "Available",
          notes: "Day 2 Only"
        };
      } else {
        return {
          profileId: profileId,
          fullName: name,
          response: "Unavailable",
          notes: "Unavailable for Round"
        };
      }
    }
  });
}

/**
 * Formats a full name to short name with last initial: "{First Name} {LastNameInitial}".
 * Strips out any junior / age group tags (e.g. "(U16)", "(U18)").
 * 
 * @param {string} fullName - e.g. "Liam Wootten", "Shahmeer Hassaan (U16)"
 * @returns {string} - e.g. "Liam W", "Shahmeer H"
 */
function formatShortPlayerName(fullName) {
  if (!fullName || typeof fullName !== "string") return "";
  var trimmed = fullName.trim();
  if (!trimmed) return "";

  // Strip any parenthetical suffix (e.g. "(U16)", "(U18)", "(wk)", "(c)")
  var cleanName = trimmed.replace(/\s*\([^)]*\)/g, "").trim();

  var parts = cleanName.split(/\s+/);
  if (parts.length === 1) {
    return parts[0];
  }
  var firstName = parts[0];
  var lastName = parts.slice(1).join(" ");
  var lastInitial = lastName.charAt(0).toUpperCase();

  return firstName + " " + lastInitial;
}

/**
 * Builds the Initial Availability Callout WhatsApp message.
 * 
 * @param {string|Date} roundDate - e.g. "2025-10-18"
 * @returns {string} Formatted WhatsApp message
 */
function generateAvailabilityCalloutMessage(roundDate) {
  var dateStr = normalizeDateToYYYYMMDD(roundDate) || String(roundDate || "").trim();

  return "🏏 *LABURNUM CC ROUND AVAILABILITY* 🏏\n\n" +
    "Please submit your availability for the upcoming round (" + dateStr + "):\n" +
    "https://lcc-availability.web.app/?round=" + dateStr;
}

/**
 * Builds the Wall of Shame WhatsApp message with the updated template.
 * 
 * @param {string|Date} roundDate - e.g. "2025-10-18"
 * @param {number} declaredCount - number of players who have declared
 * @param {Array<string>} unknownPlayerNames - list of player names still outstanding
 * @returns {string} Formatted WhatsApp message
 */
function generateWallOfShameMessage(roundDate, declaredCount, unknownPlayerNames) {
  var dateStr = normalizeDateToYYYYMMDD(roundDate) || String(roundDate || "").trim();
  var numDeclared = Number(declaredCount) || 0;

  var names = [];
  if (Array.isArray(unknownPlayerNames)) {
    names = unknownPlayerNames.map(function(n) {
      if (Array.isArray(n)) n = n[0];
      return formatShortPlayerName(String(n || ""));
    }).filter(function(n) { return n.length > 0; });
  }

  var nameListStr = names.length > 0 ? names.join(", ") : "None! Everyone has responded!";

  return "✅ Thanks to the " + numDeclared + " players who have confirmed\n\n" +
    "🏏 Yet to Get off the Mark 🏏\n" +
    "The following players currently won't be troubling the scorers this round: " + nameListStr + "\n\n" +
    "⚡ Declare your availability in 10 seconds:\n" +
    "https://lcc-availability.web.app/?round=" + dateStr;
}

/**
 * Strips junior / age group tags (e.g. "(U16)", "(U18)", "(wk)", "(c)") from player name.
 * 
 * @param {string} fullName - e.g. "Shahmeer Hassaan (U16)", "Heath Elias (U18)"
 * @returns {string} - e.g. "Shahmeer Hassaan", "Heath Elias"
 */
function stripJuniorTag(fullName) {
  if (!fullName || typeof fullName !== "string") return "";
  var trimmed = fullName.trim();
  if (!trimmed) return "";
  return trimmed.replace(/\s*\([^)]*\)/g, "").trim();
}

/**
 * Formats a player name with presentation role tags: (C), (VC), (Wk).
 * Supports dual roles (e.g. "VC & WK" -> "Neil Kloot (VC) (Wk)") and split keeping.
 *
 * @param {string} name - Player full name
 * @param {string} role - Slot role string
 * @returns {string} - Formatted name with badges
 */
function formatPlayerPresentationName(name, role) {
  if (!name || typeof name !== "string" || name.trim() === "") return "";
  var cleanName = name.replace(/\s*\(.*?\)/g, "").trim();
  var rLower = String(role || "").toLowerCase();
  
  var roleTags = [];
  
  // 1. Captain (ensure not Vice Captain)
  if ((rLower.indexOf("captain") > -1 || rLower.indexOf("(c)") > -1 || rLower.indexOf("1. captain") > -1) &&
      rLower.indexOf("vice") === -1 && rLower.indexOf("vc") === -1) {
    roleTags.push("(C)");
  }
  
  // 2. Vice Captain
  if (rLower.indexOf("vc") > -1 || rLower.indexOf("vice") > -1 || rLower.indexOf("(vc)") > -1 || rLower.indexOf("2. vc") > -1) {
    roleTags.push("(VC)");
  }
  
  // 3. Wicket Keeper
  if (rLower.indexOf("wk") > -1 || rLower.indexOf("keeper") > -1 || rLower.indexOf("(wk)") > -1 || rLower.indexOf("3. wk") > -1) {
    roleTags.push("(Wk)");
  }
  
  if (roleTags.length > 0) {
    return cleanName + " " + roleTags.join(" ");
  }
  return cleanName;
}

/**
 * Formats combined round and opponent text for presentation.
 * e.g. "Round 1: LCC 1st XI vs Mitcham - 2nd XI"
 *
 * @param {string} round - Round number or name (e.g. "1" or "Round 1")
 * @param {string} prefix - Team prefix (e.g. "1ST", "2ND")
 * @param {string} opponent - Opponent team name
 * @returns {string}
 */
function formatRoundOpponent(round, prefix, opponent) {
  var rStr = String(round || "").trim();
  if (rStr && !/^round/i.test(rStr) && /^\d+/.test(rStr)) {
    rStr = "Round " + rStr;
  }
  
  var xiLabel = "LCC 1st XI";
  var pLower = String(prefix || "").toLowerCase();
  if (pLower.indexOf("t20") > -1) {
    if (pLower.indexOf("1") > -1) xiLabel = "LCC T20 1st XI";
    else if (pLower.indexOf("2") > -1) xiLabel = "LCC T20 2nd XI";
    else xiLabel = "LCC T20 XI";
  } else if (pLower.indexOf("1") > -1) {
    xiLabel = "LCC 1st XI";
  } else if (pLower.indexOf("2") > -1) {
    xiLabel = "LCC 2nd XI";
  } else if (pLower.indexOf("3") > -1) {
    xiLabel = "LCC 3rd XI";
  } else if (pLower.indexOf("4") > -1) {
    xiLabel = "LCC 4th XI";
  } else if (pLower.indexOf("5") > -1) {
    xiLabel = "LCC 5th XI";
  }
  
  var oppStr = String(opponent || "").trim();
  if (rStr && oppStr) {
    return rStr + ": " + xiLabel + " vs " + oppStr;
  } else if (oppStr) {
    return xiLabel + " vs " + oppStr;
  } else if (rStr) {
    return rStr + ": " + xiLabel;
  }
  return xiLabel;
}

/**
 * Extracts slot number (1-13) for a specific team prefix from an Alt Text title or description.
 * Prevents cross-matching between teams (e.g. 1ST vs T20_1ST, 2ND vs T20_2ND).
 *
 * @param {string} tag Alt Text title or description tag.
 * @param {string} teamPrefix Team prefix (e.g. 'T20_1ST', 'T20_2ND', '1ST', '2ND', etc.).
 * @returns {number|null} 1-based slot index or null if not a match.
 */
function extractSlotNumberForTeam(tag, teamPrefix) {
  if (!tag) return null;
  var upperTag = String(tag).trim().toUpperCase();
  var pUpper = String(teamPrefix || "").trim().toUpperCase();

  // Normalize delimiters (replace spaces, colons, dashes with underscores)
  var normalizedTag = upperTag.replace(/[\s\-:]+/g, "_");
  var normalizedPrefix = pUpper.replace(/[\s\-:]+/g, "_");

  var isT20 = normalizedPrefix.indexOf("T20") > -1;
  var isT20_1st = normalizedPrefix === "T20_1ST" || (isT20 && (normalizedPrefix.indexOf("1") > -1 || normalizedPrefix.indexOf("FIRST") > -1));
  var isT20_2nd = normalizedPrefix === "T20_2ND" || (isT20 && (normalizedPrefix.indexOf("2") > -1 || normalizedPrefix.indexOf("SECOND") > -1));

  // Check if tag explicitly belongs to a DIFFERENT team
  if (isT20_1st) {
    if (/^T20_?(?:2ND|2|SECOND)/i.test(normalizedTag)) return null;
    if (/^(?:1ST|2ND|3RD|4TH|5TH|FIRST|SECOND|THIRD|FOURTH|FIFTH)_/i.test(normalizedTag) && !/^T20/i.test(normalizedTag)) return null;
  } else if (isT20_2nd) {
    if (/^T20_?(?:1ST|1|FIRST)/i.test(normalizedTag)) return null;
    if (/^(?:1ST|2ND|3RD|4TH|5TH|FIRST|SECOND|THIRD|FOURTH|FIFTH)_/i.test(normalizedTag) && !/^T20/i.test(normalizedTag)) return null;
  } else {
    // Standard Saturday teams
    if (/^T20/i.test(normalizedTag)) return null;
    var otherSatPrefixes = ["1ST", "2ND", "3RD", "4TH", "5TH"];
    for (var i = 0; i < otherSatPrefixes.length; i++) {
      if (normalizedPrefix !== otherSatPrefixes[i] && normalizedTag.indexOf(otherSatPrefixes[i] + "_") === 0) {
        return null;
      }
    }
  }

  // Strip team prefix from the start of tag if present
  var cleanTag = normalizedTag;
  if (isT20_1st) {
    cleanTag = cleanTag.replace(/^T20_?(?:1ST|1|FIRST)?(?:_XI|_ELEVEN)?_+/i, "");
  } else if (isT20_2nd) {
    cleanTag = cleanTag.replace(/^T20_?(?:2ND|2|SECOND)?(?:_XI|_ELEVEN)?_+/i, "");
  } else {
    cleanTag = cleanTag.replace(new RegExp("^(?:" + normalizedPrefix + "|FIRST|SECOND|THIRD|FOURTH|FIFTH)(?:_XI|_ELEVEN)?_+", "i"), "");
  }

  // Extract slot number from remaining tag (e.g. SLOT_1, PLAYER_1, PHOTO_1, IMAGE_1, 1)
  var slotMatch = cleanTag.match(/^(?:SLOT|PLAYER|PHOTO|IMAGE|IMG|PIC|AVATAR|HEADSHOT|P|PH)?_?(\d+)$/i);
  if (slotMatch) {
    return parseInt(slotMatch[1], 10);
  }

  return null;
}

/**
 * Returns vertical team frame configurations for a round tab or staging hub.
 *
 * @param {boolean} isT20 Whether the round is a T20 round.
 * @returns {Array<Object>}
 */
function getRoundFrames(isT20) {
  if (isT20) {
    return [
      { name: "T20 1st ELEVEN", prefix: "T20 1st", altPrefix: "T20_1ST", defaultIdx: 0, start: 4, slots: 12 },
      { name: "T20 SECOND ELEVEN", prefix: "T20 2nd", altPrefix: "T20_2ND", defaultIdx: 1, start: 22, slots: 12 }
    ];
  }
  return [
    { name: "FIRST ELEVEN", prefix: "1st", altPrefix: "1ST", defaultIdx: 0, start: 4, slots: 12 },
    { name: "SECOND ELEVEN", prefix: "2nd", altPrefix: "2ND", defaultIdx: 1, start: 22, slots: 12 },
    { name: "THIRD ELEVEN", prefix: "3rd", altPrefix: "3RD", defaultIdx: 2, start: 40, slots: 13 },
    { name: "FOURTH ELEVEN", prefix: "4th", altPrefix: "4TH", defaultIdx: 3, start: 59, slots: 13 },
    { name: "FIFTH ELEVEN", prefix: "5th", altPrefix: "5TH", defaultIdx: 4, start: 78, slots: 13 }
  ];
}

/**
 * Detects whether a given fixture row object, tab name, or date corresponds to a T20 round.
 *
 * @param {Object|string} fixInfoOrDate Either a fixInfo object (keys like '1st format', 't20 1st opponent') or date string.
 * @param {Array<Array>} [fixtures2DData] Optional 2D array of Fixtures sheet to lookup date.
 * @returns {boolean} True if T20 round.
 */
function isT20Fixture(fixInfoOrDate, fixtures2DData) {
  if (!fixInfoOrDate) return false;

  if (typeof fixInfoOrDate === "object" && !Array.isArray(fixInfoOrDate)) {
    for (var k in fixInfoOrDate) {
      var keyLower = String(k || "").toLowerCase();
      var valLower = String(fixInfoOrDate[k] || "").toLowerCase().trim();
      if (keyLower.indexOf("t20") > -1 && valLower !== "") {
        return true;
      }
      if (keyLower.indexOf("format") > -1 && valLower.indexOf("t20") > -1) {
        return true;
      }
    }
    return false;
  }

  var dateStr = String(fixInfoOrDate || "").trim();
  if (/t20/i.test(dateStr)) return true;

  if (fixtures2DData && Array.isArray(fixtures2DData) && fixtures2DData.length > 1) {
    var headers = fixtures2DData[0];
    var dateColIdx = -1;
    headers.forEach(function(h, i) {
      var ch = String(h || "").trim().toLowerCase();
      if (ch === "game date" || ch === "date" || ch.indexOf("date") > -1) {
        dateColIdx = i;
      }
    });

    if (dateColIdx !== -1) {
      var targetNorm = normalizeDateToYYYYMMDD(dateStr);
      for (var r = 1; r < fixtures2DData.length; r++) {
        var row = fixtures2DData[r];
        var rowDateNorm = normalizeDateToYYYYMMDD(row[dateColIdx]);
        if (rowDateNorm === targetNorm || String(row[dateColIdx] || "").trim() === dateStr) {
          for (var c = 0; c < headers.length; c++) {
            var hLower = String(headers[c] || "").toLowerCase();
            var val = String(row[c] || "").trim().toLowerCase();
            if (!val) continue;
            if (hLower.indexOf("t20") > -1) return true;
            if (hLower.indexOf("format") > -1 && val.indexOf("t20") > -1) return true;
          }
        }
      }
    }
  }

  return false;
}

/**
 * Extracts and filters players from the 2D data of the Players tab for a round snapshot.
 *
 * @param {Array<Array>} players2DData 2D array from Players tab (including headers in row 0).
 * @param {boolean} isT20 Whether to filter exclusively for T20 squad players (T20Squad === "Yes").
 * @param {Date|string} [targetDate] Date of the match for checking return dates.
 * @returns {{ unavailableSnapshot: Array<Array<string>>, unknownSnapshot: Array<Array<string>> }}
 */
function filterPlayersForRoundSnapshot(players2DData, isT20, targetDate) {
  if (!players2DData || !Array.isArray(players2DData) || players2DData.length < 2) {
    return { unavailableSnapshot: [], unknownSnapshot: [] };
  }

  var headers = players2DData[0];
  var idCol = 0;
  var fNameCol = 1;
  var lNameCol = 2;
  var fullCol = 3;
  var juniorCol = 4;
  var t20Col = 5;
  var statusCol = 6;
  var returnCol = 7;

  headers.forEach(function(h, i) {
    var ch = String(h || "").trim().toLowerCase();
    if (ch === "profileid" || ch === "profile id" || ch === "id") idCol = i;
    else if (ch === "firstname" || ch === "first name") fNameCol = i;
    else if (ch === "lastname" || ch === "last name") lNameCol = i;
    else if (ch === "fullname" || ch === "full name") fullCol = i;
    else if (ch === "juniorlevel" || ch === "junior level") juniorCol = i;
    else if (ch.indexOf("t20") > -1) t20Col = i;
    else if (ch === "globalstatus" || ch === "global status" || ch === "status") statusCol = i;
    else if (ch === "expectedreturndate" || ch === "return date") returnCol = i;
  });

  var matchDate = targetDate ? (targetDate instanceof Date ? targetDate : new Date(targetDate)) : null;
  var unavailableSnapshot = [];
  var unknownSnapshot = [];

  for (var r = 1; r < players2DData.length; r++) {
    var row = players2DData[r];
    var profileId = String(row[idCol] || "").trim();
    var fullName = String(row[fullCol] || (String(row[fNameCol] || "") + " " + String(row[lNameCol] || ""))).trim();
    if (!profileId || !fullName) continue;

    var t20Raw = (t20Col !== -1 && t20Col < row.length) ? row[t20Col] : "";
    var t20Val = String(t20Raw || "").trim().toLowerCase();
    var isExplicitYes = (t20Val === "yes" || t20Val === "y" || t20Val === "true");

    // If T20 round, MUST have an explicit "Yes" in T20Squad.
    // Blank, "No", null, undefined, or any other value is EXCLUDED.
    if (isT20 && !isExplicitYes) {
      continue;
    }

    var globalStatus = String(row[statusCol] || "Active").trim();
    var returnDateRaw = row[returnCol];
    var juniorLevel = String(row[juniorCol] || "").trim();

    var displayName = formatNameWithJuniorTag(fullName, juniorLevel);

    var isExempt = false;
    if (globalStatus === "Injured" || globalStatus === "Long-Term Away" || globalStatus === "Inactive") {
      if (returnDateRaw && globalStatus !== "Inactive" && matchDate && !isNaN(matchDate.getTime())) {
        var returnDate = new Date(returnDateRaw);
        if (!isNaN(returnDate.getTime()) && returnDate >= matchDate) {
          isExempt = true;
        }
      } else {
        isExempt = true;
      }
    }

    if (isExempt) {
      unavailableSnapshot.push([profileId, displayName, "Global Exemption: " + globalStatus, ""]);
    } else if (globalStatus === "Active") {
      unknownSnapshot.push([profileId, displayName, "", ""]);
    }
  }

  return {
    unavailableSnapshot: unavailableSnapshot,
    unknownSnapshot: unknownSnapshot
  };
}

/**
 * Formats combined format and venue text for presentation.
 * e.g. "Two Day game at Kalang Park"
 *
 * @param {string} format - Match format (e.g. "Two Day", "One Day", "T20")
 * @param {string} venue - Ground name (e.g. "Kalang Park")
 * @returns {string}
 */
function formatFormatVenue(format, venue) {
  var fStr = String(format || "").trim();
  var vStr = String(venue || "").trim();
  
  if (fStr && vStr) {
    if (/game|match/i.test(fStr)) {
      return fStr + " at " + vStr;
    }
    return fStr + " game at " + vStr;
  } else if (vStr) {
    return "at " + vStr;
  } else if (fStr) {
    return /game|match/i.test(fStr) ? fStr : (fStr + " game");
  }
  return "";
}

/**
 * Processes 2D data from the Fixtures tab into structured, chronological fixture
 * options for round tab initialisation.
 *
 * Formats dropdown labels as:
 * Day and date ({list of teams with an entry for the fixture})
 * e.g. "Sat 16 Nov 2025 (1st XI)"
 * e.g. "Sat 10 Jan 2026 (1st XI, 2nd XI, 3rd XI, 4th XI, 5th XI)"
 * e.g. "Tue 20 Jan 2026 (T20 1st XI, T20 2nd XI)"
 *
 * @param {Array<Array>} fixturesData 2D array of values from the Fixtures tab.
 * @param {Array<string>} existingSheetNames List of current sheet names in the workbook.
 * @returns {Array<Object>} List of structured fixture options sorted by date.
 */
function getFixtureOptionsForRoundInit(fixturesData, existingSheetNames) {
  if (!fixturesData || !Array.isArray(fixturesData) || fixturesData.length < 2) {
    return [];
  }

  var existingTabs = (existingSheetNames && Array.isArray(existingSheetNames))
    ? existingSheetNames.map(function(s) { return String(s || "").trim(); })
    : [];

  var headers = fixturesData[0];
  var dateColIdx = -1;

  headers.forEach(function(h, i) {
    var cleanHeader = String(h || "").trim().toLowerCase();
    if (cleanHeader === "game date" || cleanHeader === "date" || cleanHeader === "match date") {
      dateColIdx = i;
    }
  });

  if (dateColIdx === -1) {
    return [];
  }

  // Dynamically discover all teams and their field column indices from headers
  var teamDescriptors = [];
  var seenPrefixes = {};

  headers.forEach(function(h, i) {
    var cleanHeader = String(h || "").trim();
    var match = cleanHeader.match(/^(.*?)\s+(Round|Format|Opponent|Venue)$/i);
    if (match) {
      var prefix = match[1].trim();
      var key = prefix.toLowerCase();
      if (!seenPrefixes[key]) {
        seenPrefixes[key] = {
          prefix: prefix,
          name: (/XI$|Team$/i.test(prefix) ? prefix : (prefix + " XI")),
          roundIdx: -1,
          formatIdx: -1,
          oppIdx: -1,
          venueIdx: -1
        };
        teamDescriptors.push(seenPrefixes[key]);
      }
      var field = match[2].toLowerCase();
      if (field === "round") seenPrefixes[key].roundIdx = i;
      else if (field === "format") seenPrefixes[key].formatIdx = i;
      else if (field === "opponent") seenPrefixes[key].oppIdx = i;
      else if (field === "venue") seenPrefixes[key].venueIdx = i;
    }
  });

  var daysOfWeek = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var fullDays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var fullMonths = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  var optionsMap = {};

  for (var r = 1; r < fixturesData.length; r++) {
    var row = fixturesData[r];
    var rawDate = row[dateColIdx];
    var normalizedDate = normalizeDateToYYYYMMDD(rawDate);
    if (!normalizedDate) continue;

    var parts = normalizedDate.split("-");
    var year = parts[0];
    var monthIdx = parseInt(parts[1], 10) - 1;
    var dayNum = parseInt(parts[2], 10);
    var dObj = new Date(parseInt(parts[0], 10), monthIdx, dayNum);
    var dayName = !isNaN(dObj.getTime()) ? daysOfWeek[dObj.getDay()] : "";
    var fullDayName = !isNaN(dObj.getTime()) ? fullDays[dObj.getDay()] : "";
    var monthName = (monthIdx >= 0 && monthIdx < 12) ? months[monthIdx] : "";
    var fullMonthName = (monthIdx >= 0 && monthIdx < 12) ? fullMonths[monthIdx] : "";

    // E.g. "Sat 16 Nov 2025" or "Sat 10 Jan 2026"
    var dayDateStr = dayName ? (dayName + " " + dayNum + " " + monthName + " " + year) : (dayNum + " " + monthName + " " + year);
    var fullDisplayDate = fullDayName ? (fullDayName + ", " + dayNum + " " + fullMonthName + " " + year) : dayDateStr;

    var tabExists = existingTabs.some(function(tabName) {
      return tabName === normalizedDate || tabName === String(rawDate).trim() || normalizeDateToYYYYMMDD(tabName) === normalizedDate;
    });

    var teamMatches = [];
    var teamsWithEntries = [];

    teamDescriptors.forEach(function(t) {
      var tRound = t.roundIdx !== -1 && row[t.roundIdx] ? String(row[t.roundIdx]).trim() : "";
      var tFmt = t.formatIdx !== -1 && row[t.formatIdx] ? String(row[t.formatIdx]).trim() : "";
      var tOpp = t.oppIdx !== -1 && row[t.oppIdx] ? String(row[t.oppIdx]).trim() : "";
      var tVen = t.venueIdx !== -1 && row[t.venueIdx] ? String(row[t.venueIdx]).trim() : "";

      if (tOpp || tVen || tFmt || tRound) {
        teamsWithEntries.push(t.name);
        teamMatches.push({
          team: t.name,
          round: tRound,
          format: tFmt,
          opponent: tOpp,
          venue: tVen
        });
      }
    });

    // Build label: Day and date ({list of teams with an entry for the fixture})
    // E.g. "Sat 16 Nov 2025 (1st XI)"
    // E.g. "Sat 10 Jan 2026 (1st XI, 2nd XI, 3rd XI, 4th XI, 5th XI)"
    // E.g. "Tue 20 Jan 2026 (T20 1st XI, T20 2nd XI)"
    var label = dayDateStr;
    if (teamsWithEntries.length > 0) {
      label += " (" + teamsWithEntries.join(", ") + ")";
    }
    if (tabExists) {
      label += " [Tab Initialised]";
    }

    if (!optionsMap[normalizedDate]) {
      optionsMap[normalizedDate] = {
        date: normalizedDate,
        rawDate: String(rawDate).trim(),
        displayDate: dayDateStr,
        fullDisplayDate: fullDisplayDate,
        teams: teamsWithEntries,
        tabExists: tabExists,
        label: label,
        teamMatches: teamMatches
      };
    }
  }

  var optionsList = Object.keys(optionsMap).map(function(k) {
    return optionsMap[k];
  });

  optionsList.sort(function(a, b) {
    return a.date.localeCompare(b.date);
  });

  return optionsList;
}

// Node/Jest interop — Apps Script ignores this guard.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    normalizePhone,
    normalizeDateToYYYYMMDD,
    findHouseholdPlayers,
    moveItemBetweenLists,
    formatNameWithJuniorTag,
    formatShortPlayerName,
    stripJuniorTag,
    pickFirstName,
    formatPlayerPresentationName,
    formatRoundOpponent,
    formatFormatVenue,
    DEFAULT_TEAM_CONFIGS,
    parseCsvString,
    formatTeamPrefix,
    buildFixtureHeaders,
    readTeamConfigsFromSheet,
    readTemplatesFromSheet,
    parseFixtureCsv,
    mergeFixturesIntoMatrix,
    getFixtureOptionsForRoundInit,
    getRoundFrames,
    extractSlotNumberForTeam,
    isT20Fixture,
    filterPlayersForRoundSnapshot,
    simulatePlayerAvailability,
    generateAvailabilityCalloutMessage,
    generateWallOfShameMessage
  };
}


