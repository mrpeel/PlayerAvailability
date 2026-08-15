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

// Node/Jest interop — Apps Script ignores this guard.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    normalizePhone,
    normalizeDateToYYYYMMDD,
    findHouseholdPlayers,
    moveItemBetweenLists,
    formatNameWithJuniorTag,
    pickFirstName,
    DEFAULT_TEAM_CONFIGS,
    parseCsvString,
    formatTeamPrefix,
    buildFixtureHeaders,
    readTeamConfigsFromSheet,
    readTemplatesFromSheet,
    parseFixtureCsv,
    mergeFixturesIntoMatrix,
    simulatePlayerAvailability
  };
}


