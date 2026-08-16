

var LCC_PALETTE = {
  maroonBg: "#6A1B29",
  maroonFg: "#ffffff",
  goldBg: "#F4B41A",
  goldFg: "#111111",
  inputHighlight: "#FFF6E5",
  grayBorder: "#e5e7eb",
  zebraLight: "#f9fafb"
};


/**
 * Diagnostic: lists all files in the headshot folder.
 * Run from the editor to verify folder access and see what's there.
 */
function checkHeadshotFolder() {
  var folderId = getHeadshotFolderId();
  if (!folderId) {
    Logger.log("HEADSHOT_FOLDER_ID not set. Set it in File > Project properties > Script properties.");
    return;
  }
  
  try {
    var folder = DriveApp.getFolderById(folderId);
    Logger.log("Folder: " + folder.getName());
    Logger.log("URL: " + folder.getUrl());
    Logger.log("---");
    
    var files = folder.getFiles();
    var count = 0;
    while (files.hasNext()) {
      var file = files.next();
      Logger.log("  " + file.getName() + " | ID: " + file.getId());
      count++;
    }
    Logger.log("---");
    Logger.log("Total files: " + count);
  } catch (e) {
    Logger.log("ERROR: " + e.message);
  }
}


/**
 * WORKSPACE GENERATOR: Spawns a standalone date snapshot ledger sheet.
 */
function deployVerticalRoundSheet(ss, dateStr) {
  if (typeof ss === "string" && !dateStr) {
    dateStr = ss;
    ss = getSS();
  }
  if (!ss || typeof ss.getSheetByName !== "function") {
    ss = getSS();
  }
  var sheetName = dateStr;
  if (ss.getSheetByName(sheetName)) {
    throw new Error("Tab '" + sheetName + "' already exists in this spreadsheet.");
  }

  var ws = ss.insertSheet(sheetName);
  var targetDate = new Date(dateStr);

  // Lookup fixture details for this date if present
  var fixSheet = ss.getSheetByName("Fixtures");
  var fixInfo = null;
  if (fixSheet && fixSheet.getLastRow() > 1) {
    var fValues = fixSheet.getDataRange().getValues();
    var fHeaders = fValues[0];
    var dateColIdx = -1;
    fHeaders.forEach(function(h, i) {
      if (String(h).trim().toLowerCase() === "game date") dateColIdx = i;
    });
    if (dateColIdx !== -1) {
      var targetNormalized = normalizeDateToYYYYMMDD(dateStr) || String(dateStr).trim();
      for (var fRow = 1; fRow < fValues.length; fRow++) {
        var rowDateNormalized = normalizeDateToYYYYMMDD(fValues[fRow][dateColIdx]) || String(fValues[fRow][dateColIdx]).trim();
        if (rowDateNormalized === targetNormalized) {
          fixInfo = {};
          fHeaders.forEach(function(h, i) {
            fixInfo[String(h).trim().toLowerCase()] = fValues[fRow][i];
          });
          break;
        }
      }
    }
  }

  // Header Metadata (Date is key, clean single Match Date header)
  ws.getRange("A1").setValue("Match Date:").setFontWeight("bold");
  ws.getRange("B1").setValue(targetDate).setNumberFormat("dd/mm/yyyy").setBackground(LCC_PALETTE.inputHighlight).setHorizontalAlignment("center").setFontWeight("bold");
  ws.getRange("A2:B2").clear();
  ws.getRange("G1").setValue("Roster Last Synced:").setFontWeight("bold");
  ws.getRange("H1").setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm").setFontColor("#666666");

  // Vertical Team Stack (Cols A & B)
  var grades = ["FIRST ELEVEN", "SECOND ELEVEN", "THIRD ELEVEN", "FOURTH ELEVEN", "FIFTH ELEVEN"];
  var teamPrefixes = ["1st", "2nd", "3rd", "4th", "5th"];
  var currentTeamRow = 4;

  grades.forEach(function(gName, gIdx) {
    var prefix = teamPrefixes[gIdx];
    var rnd = fixInfo ? (fixInfo[prefix + " round"] || "") : "";
    var opp = fixInfo ? (fixInfo[prefix + " opponent"] || "") : "";
    var ven = fixInfo ? (fixInfo[prefix + " venue"] || "") : "";
    var fmt = fixInfo ? (fixInfo[prefix + " format"] || "") : "";

    // Team Banner
    ws.getRange(currentTeamRow, 1, 1, 2).merge().setValue(gName).setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
    
    // 4 Metadata rows: Round, Opponent, Venue, Format
    ws.getRange(currentTeamRow + 1, 1).setValue("Round:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 1, 2).setValue(rnd);
    ws.getRange(currentTeamRow + 2, 1).setValue("Opponent:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 2, 2).setValue(opp);
    ws.getRange(currentTeamRow + 3, 1).setValue("Venue:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 3, 2).setValue(ven);
    ws.getRange(currentTeamRow + 4, 1).setValue("Format:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 4, 2).setValue(fmt);

    // 13 Player Slot Rows (currentTeamRow + 5 to currentTeamRow + 17)
    var structure = [
      ["1. Captain", ""], ["2. VC", ""], ["3. WK", ""],
      ["4. Player", ""], ["5. Player", ""], ["6. Player", ""],
      ["7. Player", ""], ["8. Player", ""], ["9. Player", ""],
      ["10. Player", ""], ["11. Player", ""], ["12. Player", ""], ["13. Player", ""]
    ];
    ws.getRange(currentTeamRow + 5, 1, 13, 2).setValues(structure);
    ws.getRange(currentTeamRow + 5, 1, 13, 1).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);
    ws.getRange(currentTeamRow, 1, 18, 2).setBorder(true, true, true, true, true, true, LCC_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);

    currentTeamRow += 19;
  });

  // Col D: Dynamic Virtual "Available for Selection" Pool
  ws.getRange("D3").setValue("AVAILABLE FOR SELECTION").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
  ws.getRange("D4").setValue("(Dynamic Unselected Pool)").setFontStyle("italic").setFontColor("#666666");
  ws.getRange("D5").setFormula(`=IFERROR(SORT(FILTER(G5:G, G5:G<>"", COUNTIF(B:B, G5:G)=0)), "")`);
  ws.getRange("D3:D150").setBorder(true, true, true, true, false, false, LCC_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);

  // Extract Global Profiles and Populate Hard Snapshot Registries
  var playerSheet = ss.getSheetByName("Players");
  var lastPlayer = playerSheet ? playerSheet.getLastRow() : 0;

  var unavailableSnapshot = [];
  var unknownSnapshot = [];

  if (lastPlayer > 1) {
    var pData = playerSheet.getRange(2, 1, lastPlayer - 1, playerSheet.getLastColumn()).getValues();
    pData.forEach(function(row) {
      var profileId = String(row[0]).trim();            // Col A: ProfileID
      var fullName = row[3] || (row[1] + " " + row[2]); // Col D or B+C
      if (fullName.trim() === "" || profileId === "") return;

      var globalStatus = String(row[6]).trim();          // Col G: GlobalStatus
      var returnDateRaw = row[7];                      // Col H: ExpectedReturnDate
      var juniorLevel = row[4] || "";                  // Col E: JuniorLevel

      var displayName = formatNameWithJuniorTag(fullName, juniorLevel);

      var isExempt = false;
      if (globalStatus === "Injured" || globalStatus === "Long-Term Away" || globalStatus === "Inactive") {
        if (returnDateRaw && globalStatus !== "Inactive") {
          var returnDate = new Date(returnDateRaw);
          if (!isNaN(returnDate.getTime()) && returnDate >= targetDate) {
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
    });
  }

  // List 2: Available for Round (Cols F-I)
  ws.getRange("F3:I3").merge().setValue("AVAILABLE FOR ROUND").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
  ws.getRange("F4:I4").setValues([["Profile ID", "Player Name", "Notes Context", "Switcher"]]).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);

  // List 3: Unavailable for Round (Cols K-N)
  ws.getRange("K3:N3").merge().setValue("UNAVAILABLE FOR ROUND").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
  ws.getRange("K4:N4").setValues([["Profile ID", "Player Name", "Exemption Notes", "Switcher"]]).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);
  if (unavailableSnapshot.length > 0) {
    ws.getRange(5, 11, unavailableSnapshot.length, 4).setValues(unavailableSnapshot);
  }

  // List 4: Unknown Status Pool Ledger (Cols P-S)
  ws.getRange("P3:S3").merge().setValue("UNKNOWN AVAILABILITY STATUS").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
  ws.getRange("P4:S4").setValues([["Profile ID", "Player Name", "Notes Context", "Switcher"]]).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);
  if (unknownSnapshot.length > 0) {
    ws.getRange(5, 16, unknownSnapshot.length, 4).setValues(unknownSnapshot);
  }

  // Action Dropdown Validations
  var availRule = SpreadsheetApp.newDataValidation().requireValueInList(["🚫 Move to Unavailable"], true).build();
  var unavailRule = SpreadsheetApp.newDataValidation().requireValueInList(["✅ Move to Available"], true).build();
  var unknownRule = SpreadsheetApp.newDataValidation().requireValueInList(["✅ Move to Available", "🚫 Move to Unavailable"], true).build();

  ws.getRange("I5:I150").setDataValidation(availRule);
  ws.getRange("N5:N150").setDataValidation(unavailRule);
  ws.getRange("S5:S150").setDataValidation(unknownRule);

  // Col B Selection Dropdowns linked to team-specific dynamic pools in hidden Cols AA-CQ
  applySelectionValidationRules(ws);

  // Setup Column U and Column V WhatsApp messaging (Merged cards so Row 4 height is unaffected)
  setupWhatsAppMessageColumns(ws, targetDate);

  // Column Widths & Hide Technical ProfileID and Helper Columns
  applyRoundTabColumnWidths(ws);

  ws.hideColumns(6);  // Col F (ProfileID)
  ws.hideColumns(11); // Col K (ProfileID)
  ws.hideColumns(16); // Col P (ProfileID)

  // Apply clean Hanken Grotesk font across round sheet
  try {
    ws.getRange(1, 1, Math.min(ws.getMaxRows(), 150), Math.min(ws.getMaxColumns(), 26)).setFontFamily("Hanken Grotesk");
  } catch (e) {}

  // Sync Presentation Staging Hub with this newly deployed round
  syncPresentationStagingHub(ss, sheetName);

  ss.setActiveSheet(ws);
  return "Selection tab '" + sheetName + "' initialised successfully!\n\nSpreadsheet has switched to the new round tab.";
}


/**
 * Backend logic invoked by Datepicker dialog to deploy a new round tab.
 */
function executeTabDeployment(dateStr) {
  var ss = getSS();
  if (!ss) throw new Error("Active spreadsheet not found.");
  return deployVerticalRoundSheet(ss, dateStr);
}



// ============================================================================
// SECTION 1: WEB APP ENTRY POINTS (doGet / doPost)
// ============================================================================

/**
 * Health-check endpoint.  The real frontend lives on Firebase hosting.
 */
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: "running", version: "2.0" })
  ).setMimeType(ContentService.MimeType.JSON);
}


/**
 * JSON API for the Firebase-hosted frontend.
 *
 * SECURITY MODEL: The web app runs as the deploying (owner) account
 * (appsscript.json: executeAs=USER_DEPLOYING, access=ANYONE_ANONYMOUS).
 * Sheet/Drive access comes from that one-time owner authorization — the
 * frontend needs NO service account, API key, or spreadsheet ID.
 * App-level auth still applies: phone matching / Admins tab.
 *
 * CORS: browsers must POST with Content-Type text/plain (no preflight);
 * GAS replies with Access-Control-Allow-Origin: * and JSON.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Prevent concurrent write collisions
  } catch (lockErr) {
    Logger.log("Lock acquisition failed: " + lockErr.message);
  }

  var responseOutput = { status: "success", message: "Roster data state synchronized." };
  try {
    var contents = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    var payload = JSON.parse(contents);
    var action = payload.action || null;
    Logger.log("doPost Action: " + action + ", payload: " + contents);

    if (action === 'getInitialData') {
      var result = getInitialData(payload.phone, payload.dateStr || payload.date || payload.roundNum);
      return jsonResponse(result);
    }

    if (action === 'getAdminRounds') {
      var roundsResult = getAdminRounds(payload.phone);
      return jsonResponse(roundsResult);
    }

    if (action === 'getAdminData') {
      var adminResult = getAdminData(payload.phone, payload.dateStr || payload.roundNum || payload.date);
      return jsonResponse(adminResult);
    }

    if (action === 'simulateRoundAvailability' || action === 'simulateAvailability') {
      var simResult = simulateRoundAvailability(payload.dateStr || payload.date);
      return jsonResponse(simResult);
    }

    if (action === 'getPlayersListForStudio') {
      var studioPlayers = getPlayersListForStudio();
      return jsonResponse({ status: "success", players: studioPlayers });
    }

    if (action === 'savePlayerHeadshot') {
      var saveRes = savePlayerHeadshot(payload.profileId, payload.base64Data);
      return jsonResponse(saveRes);
    }

    if (action === 'updateGlobalStatus') {
      var profileId = String(payload.profileId || payload.playerId).trim();
      var newStatus = payload.status || 'Active';
      var returnDate = payload.returnDate || '';
      var statusResult = updateGlobalStatus(profileId, newStatus, returnDate);
      return jsonResponse(statusResult);
    }

    // Default / saveAvailability path — save availability / list movement
    var pId = String(payload.profileId || payload.playerId).trim();
    var targetDateStr = String(payload.date || payload.dateStr || payload.roundNum).trim();
    var playerResponse = String(payload.response).trim();
    var playerNotes = payload.notes || "";

    var saveResult = saveAvailability(pId, targetDateStr, playerResponse, playerNotes);
    return jsonResponse(saveResult);
  } catch (err) {
    Logger.log("doPost Error: " + err.message);
    responseOutput = { status: "error", error: err.message || err.toString(), message: err.toString() };
    return jsonResponse(responseOutput);
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}


/**
 * Updates player global status and return date in Master Players sheet.
 */
function updateGlobalStatus(profileId, status, returnDate) {
  var ss = getSS();
  var playerSheet = ss.getSheetByName("Players");
  if (!playerSheet) throw new Error("Players tab not found.");

  var pData = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, 1).getValues();
  var foundRow = -1;
  for (var i = 0; i < pData.length; i++) {
    if (String(pData[i][0]).trim() === String(profileId).trim()) {
      foundRow = 2 + i;
      break;
    }
  }

  if (foundRow === -1) throw new Error("Player ID " + profileId + " not found in Master directory.");

  playerSheet.getRange(foundRow, 7).setValue(status);
  if (returnDate !== undefined && returnDate !== null) {
    playerSheet.getRange(foundRow, 8).setValue(returnDate);
  }

  return { status: "success", success: true };
}


/**
 * Returns available rounds for admin selector console.
 */
function getAdminRounds(phone) {
  var ss = getSS();
  if (!ss) return { error: "Spreadsheet not found." };

  var normPhone = normalizePhone(phone);
  var adminSheet = ss.getSheetByName("Admins");
  if (adminSheet && adminSheet.getLastRow() > 1) {
    var adminData = adminSheet.getDataRange().getValues();
    var aHeaders = adminData[0];
    var phoneIdx = aHeaders.indexOf("Phone");
    if (phoneIdx !== -1) {
      var isAdmin = adminData.slice(1).some(function(row) {
        return normalizePhone(row[phoneIdx]) === normPhone;
      });
      if (!isAdmin) {
        return { error: "Access Denied: Phone number " + normPhone + " is not registered in the Admins tab." };
      }
    }
  }

  var fixSheet = ss.getSheetByName("Fixtures");
  var rounds = [];
  var currentRound = "";

  if (fixSheet && fixSheet.getLastRow() > 1) {
    var fValues = fixSheet.getDataRange().getValues();
    var fHeaders = fValues[0];
    var dateIdx = fHeaders.indexOf("Game Date");
    var roundIdx = fHeaders.indexOf("1st Round");
    var formatIdx = fHeaders.indexOf("1st Format");

    for (var i = 1; i < fValues.length; i++) {
      var d = String(fValues[i][dateIdx]).trim();
      var r = roundIdx !== -1 ? String(fValues[i][roundIdx]).trim() : ("Round " + i);
      var fmt = formatIdx !== -1 ? String(fValues[i][formatIdx]).trim() : "Two Day";
      if (d) {
        rounds.push({
          num: r.replace(/Round\s*/i, '').trim() || String(i),
          date: d,
          roundNum: r.replace(/Round\s*/i, '').trim() || String(i),
          roundTitle: r,
          format: fmt,
          label: r + " (" + d + " - " + fmt + ")"
        });
      }
    }
  }

  if (rounds.length > 0) currentRound = rounds[0].date;

  return {
    status: "success",
    rounds: rounds,
    current: currentRound
  };
}


/**
 * Fetches admin operational summary and WhatsApp templates for a given round.
 */
function getAdminData(phone, dateStrOrRoundNum) {
  var ss = getSS();
  if (!ss) return { error: "Spreadsheet not found." };

  var fixSheet = ss.getSheetByName("Fixtures");
  var targetDate = String(dateStrOrRoundNum || "").trim();

  // If dateStrOrRoundNum is a round number (e.g. "1"), find the matching date in Fixtures
  if (fixSheet && fixSheet.getLastRow() > 1) {
    var fValues = fixSheet.getDataRange().getValues();
    var fHeaders = fValues[0];
    var dateIdx = fHeaders.indexOf("Game Date");
    var roundIdx = fHeaders.indexOf("1st Round");

    for (var f = 1; f < fValues.length; f++) {
      var d = String(fValues[f][dateIdx]).trim();
      var r = roundIdx !== -1 ? String(fValues[f][roundIdx]).trim() : "";
      if (d === targetDate || r === "Round " + targetDate || r === targetDate || r.replace(/Round\s*/i, '').trim() === targetDate) {
        targetDate = d;
        break;
      }
    }
  }

  var roundSheet = ss.getSheetByName(targetDate);
  var availNames = [];
  var unavailNames = [];
  var unknownNames = [];

  if (roundSheet) {
    var aVals = roundSheet.getRange(5, 7, 100, 1).getValues();
    var uVals = roundSheet.getRange(5, 12, 100, 1).getValues();
    var unkWVals = roundSheet.getRange(5, 17, 100, 1).getValues();

    aVals.forEach(function(r) { if (r[0]) availNames.push(String(r[0]).trim()); });
    uVals.forEach(function(r) { if (r[0]) unavailNames.push(String(r[0]).trim()); });
    unkWVals.forEach(function(r) { if (r[0]) unknownNames.push(String(r[0]).trim()); });
  }

  var availMsg = generateAvailabilityCalloutMessage(targetDate);
  var shameMsg = generateWallOfShameMessage(targetDate, availNames.length + unavailNames.length, unknownNames);

  return {
    status: "success",
    date: targetDate,
    roundNum: targetDate,
    availMsg: availMsg,
    shameMsg: shameMsg,
    counts: {
      available: availNames.length,
      unavailable: unavailNames.length,
      unknown: unknownNames.length,
      total: availNames.length + unavailNames.length + unknownNames.length
    }
  };
}


/**
 * Simulates player responses for a round based on match format:
 * - 1-Day: 90% Available, 10% Unavailable
 * - 2-Day: 85% Available (both days), 5% Day 1 only, 5% Day 2 only, 5% Unavailable
 */
function simulateRoundAvailability(dateStr) {
  var ss = getSS();
  if (!ss) throw new Error("Spreadsheet not found.");

  var targetDate = dateStr;
  var ws = null;
  if (targetDate) {
    ws = ss.getSheetByName(targetDate);
  }
  if (!ws) {
    ws = ss.getActiveSheet();
    var forbidden = ["Presentation_Staging", "Players", "Fixtures", "Config", "Availability_Log", "Admins"];
    if (forbidden.indexOf(ws.getName()) === -1) {
      targetDate = ws.getName();
    } else {
      var fixSheet = ss.getSheetByName("Fixtures");
      if (fixSheet && fixSheet.getLastRow() > 1) {
        targetDate = String(fixSheet.getRange(2, 1).getValue()).trim();
        ws = ss.getSheetByName(targetDate);
      }
    }
  }

  if (!ws || !targetDate) {
    throw new Error("No active round selection tab found for date " + (targetDate || "") + ". Please initialise a round tab first.");
  }

  var matchFormat = "Two Day";
  var fixSheet2 = ss.getSheetByName("Fixtures");
  if (fixSheet2 && fixSheet2.getLastRow() > 1) {
    var fValues = fixSheet2.getDataRange().getValues();
    var fHeaders = fValues[0];
    var dateIdx = fHeaders.indexOf("Game Date");
    var fmtIdx = fHeaders.indexOf("1st Format");
    if (dateIdx !== -1 && fmtIdx !== -1) {
      for (var f = 1; f < fValues.length; f++) {
        if (String(fValues[f][dateIdx]).trim() === targetDate) {
          matchFormat = String(fValues[f][fmtIdx]).trim() || "Two Day";
          break;
        }
      }
    }
  }

  var playerSheet = ss.getSheetByName("Players");
  if (!playerSheet || playerSheet.getLastRow() < 2) {
    throw new Error("No players found in Players tab.");
  }

  var pData = playerSheet.getDataRange().getValues();
  var pHeaders = pData[0];
  var pRows = pData.slice(1);

  var idIdx = pHeaders.indexOf("ProfileID");
  var nameIdx = pHeaders.indexOf("FullName");
  var firstIdx = pHeaders.indexOf("FirstName");
  var lastIdx = pHeaders.indexOf("LastName");
  var statusIdx = pHeaders.indexOf("GlobalStatus");

  var players = pRows.map(function(r) {
    var fn = r[nameIdx] || (r[firstIdx] + " " + r[lastIdx]);
    return {
      profileId: String(r[idIdx]).trim(),
      fullName: fn,
      globalStatus: String(r[statusIdx] || "Active").trim()
    };
  }).filter(function(p) { return p.profileId !== ""; });

  var simulated = simulatePlayerAvailability(players, matchFormat);
  var availCount = 0;
  var unavailCount = 0;
  var partialCount = 0;

  var logSheet = ss.getSheetByName("Availability_Log");

  simulated.forEach(function(sim) {
    executeRealtimeAppListMovement(ws, sim.profileId, sim.response, sim.notes);
    if (sim.response === "Available") {
      if (sim.notes && (sim.notes === "Day 1 Only" || sim.notes === "Day 2 Only")) {
        partialCount++;
      } else {
        availCount++;
      }
    } else {
      unavailCount++;
    }

    if (logSheet) {
      logSheet.appendRow([new Date(), sim.profileId, targetDate, sim.response, sim.notes || "Simulated"]);
    }
  });

  return {
    status: "success",
    date: targetDate,
    format: matchFormat,
    total: players.length,
    available: availCount,
    partial: partialCount,
    unavailable: unavailCount,
    message: "Simulated availability for " + players.length + " players on round " + targetDate + " (" + matchFormat + "):\n\n" +
      "• Full Availability: " + availCount + " players (" + Math.round((availCount/players.length)*100) + "%)\n" +
      (partialCount > 0 ? "• Partial Availability: " + partialCount + " players (" + Math.round((partialCount/players.length)*100) + "%)\n" : "") +
      "• Unavailable / Injured / Inactive: " + unavailCount + " players (" + Math.round((unavailCount/players.length)*100) + "%)\n\n" +
      "Rosters on tab '" + targetDate + "' and Availability_Log updated."
  };
}


// ============================================================================
// SECTION 5: PURE UTILITIES (also in logic.js for testing)
// ============================================================================

/**
 * Junior suffix tag generator.  Appends (U14), (U16), or (U18).
 * Legacy values U16_Y2 / U16_Y1 both map to (U16).
 */
function formatNameWithJuniorTag(fullName, juniorLevel) {
  if (juniorLevel === "U18") return fullName + " (U18)";
  if (juniorLevel === "U16" || juniorLevel === "U16_Y2" || juniorLevel === "U16_Y1") return fullName + " (U16)";
  if (juniorLevel === "U14") return fullName + " (U14)";
  return fullName;
}



/**
 * Returns the Drive folder ID for player headshots.
 * Set it once via: Apps Script editor > File > Project properties > Script properties
 * Add a property named HEADSHOT_FOLDER_ID with your folder ID.
 */
function getHeadshotFolderId() {
  return PropertiesService.getScriptProperties().getProperty('HEADSHOT_FOLDER_ID') || '';
}


/**
 * Fetches initial data payload for the Web App frontend
 */
function getInitialData(phone, dateStr) {
  try {
    const ss = getSS();
    if (!ss) return { error: 'CRITICAL: Spreadsheet not found.' };

    const playersSheet = ss.getSheetByName('Players');
    const fixturesSheet = ss.getSheetByName('Fixtures');

    if (!playersSheet || !fixturesSheet) {
      return { error: 'DATABASE ERROR: Players or Fixtures tab missing.' };
    }

    const normalizedInputPhone = normalizePhone(phone);
    if (!normalizedInputPhone) return { error: 'Please enter a valid phone number.' };

    // 1. MATCH HOUSEHOLD PLAYERS ACROSS ALL 4 PHONE COLUMNS
    const pData = playersSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    const pRows = pData.slice(1);

    const profileIdIdx = pHeaders.indexOf('ProfileID');
    const firstNameIdx = pHeaders.indexOf('FirstName');
    const lastNameIdx = pHeaders.indexOf('LastName');
    const fullNameIdx = pHeaders.indexOf('FullName');
    const statusIdx = pHeaders.indexOf('GlobalStatus');
    const returnDateIdx = pHeaders.indexOf('ExpectedReturnDate');
    const juniorLevelIdx = pHeaders.indexOf('JuniorLevel') !== -1 ? pHeaders.indexOf('JuniorLevel') : pHeaders.indexOf('JuniorClass');

    const phoneIndexes = [
      pHeaders.indexOf('Phone'),
      pHeaders.indexOf('Phone2'),
      pHeaders.indexOf('Phone3'),
      pHeaders.indexOf('Phone4')
    ].filter(idx => idx !== -1);

    const household = pRows.filter(row => {
      return phoneIndexes.some(pIdx => normalizePhone(row[pIdx]) === normalizedInputPhone);
    }).map(row => {
      const pId = String(row[profileIdIdx]).trim();
      return {
        profileId: pId,
        firstName: row[firstNameIdx],
        lastName: row[lastNameIdx],
        fullName: row[fullNameIdx] || (row[firstNameIdx] + " " + row[lastNameIdx]),
        globalStatus: row[statusIdx],
        expectedReturnDate: row[returnDateIdx] ? String(row[returnDateIdx]) : "",
        juniorLevel: row[juniorLevelIdx] || "",
        photoUrl: getPlayerPhotoUrl(pId)
      };
    });

    if (household.length === 0) {
      return { error: 'No registered players found matching phone number ' + normalizedInputPhone };
    }

    // 2. FETCH FIXTURE / ROUND DETAILS FOR THE DATE
    const fData = fixturesSheet.getDataRange().getValues();
    const fHeaders = fData[0];
    const fRows = fData.slice(1);
    const gameDateIdx = fHeaders.indexOf('Game Date');

    let targetFixture = null;
    if (dateStr) {
      targetFixture = fRows.find(row => String(row[gameDateIdx]).indexOf(dateStr) > -1);
      const roundTabExists = ss.getSheetByName(dateStr) !== null;
      if (!targetFixture && !roundTabExists) {
        return {
          error: "Round '" + dateStr + "' was not found. Please verify the link or check with your club selector.",
          invalidRound: true,
          requestedDate: dateStr
        };
      }
    } else {
      targetFixture = fRows.length > 0 ? fRows[0] : null;
    }

    const fixtureInfo = {};
    if (targetFixture) {
      fHeaders.forEach((h, i) => fixtureInfo[h] = targetFixture[i] ? String(targetFixture[i]) : "");
    }

    // 3. FETCH CURRENT AVAILABILITY & NOTES FROM THE TARGET DATE TAB
    const targetSheetName = dateStr || (fixtureInfo['Game Date'] ? fixtureInfo['Game Date'] : "");
    const dateSheet = ss.getSheetByName(targetSheetName);

    const householdAvail = {};
    const householdNotes = {};

    if (dateSheet) {
      const availIds = dateSheet.getRange(5, 6, 100, 3).getValues();
      const unavailIds = dateSheet.getRange(5, 11, 100, 3).getValues();
      const unknownIds = dateSheet.getRange(5, 16, 100, 3).getValues();

      household.forEach(player => {
        const pId = player.profileId;
        let status = "Unknown";
        let note = "";

        let match = availIds.find(r => String(r[0]).trim() === pId);
        if (match) {
          status = "Available";
          note = match[2] || "";
        } else {
          match = unavailIds.find(r => String(r[0]).trim() === pId);
          if (match) {
            status = "Unavailable";
            note = match[2] || "";
          } else {
            match = unknownIds.find(r => String(r[0]).trim() === pId);
            if (match) {
              status = "Unknown";
              note = match[2] || "";
            }
          }
        }

        householdAvail[pId] = status;
        householdNotes[pId] = note;
      });
    } else {
      household.forEach(player => {
        householdAvail[player.profileId] = "Unknown";
        householdNotes[player.profileId] = "";
      });
    }

    var roundNumStr = fixtureInfo['1st Round'] || fixtureInfo['RoundID'] || 'Round 1';
    var formatStr = fixtureInfo['1st Format'] || 'Two Day';
    var gameDateStr = fixtureInfo['Game Date'] || targetSheetName;

    return {
      status: "success",
      players: household.map(function(p) {
        return {
          id: p.profileId,
          profileId: p.profileId,
          name: p.fullName,
          fullName: p.fullName,
          firstName: p.firstName,
          lastName: p.lastName,
          globalStatus: p.globalStatus || 'Active',
          expectedReturnDate: p.expectedReturnDate || '',
          juniorLevel: p.juniorLevel || '',
          photoUrl: p.photoUrl || ''
        };
      }),
      fixtureInfo: fixtureInfo,
      roundInfo: {
        RoundNum: roundNumStr.replace(/Round\s*/i, '').trim() || '1',
        RoundTitle: roundNumStr,
        Date1: gameDateStr,
        Date2: '',
        Format: formatStr,
        Opponent: fixtureInfo['1st Opponent'] || '',
        Venue: fixtureInfo['1st Venue'] || ''
      },
      availability: householdAvail,
      notes: householdNotes,
      matchDate: targetSheetName
    };

  } catch (e) {
    return { error: `SERVER ERROR: ${e.message}` };
  }
}


/**
 * Helper to construct the Drive thumbnail URL for player headshots
 */
function getPlayerPhotoUrl(profileId) {
  var folderId = getHeadshotFolderId();
  if (!folderId) return "";
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(`${profileId}.jpg`);
    if (files.hasNext()) {
      const file = files.next();
      return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
    }
  } catch (e) {
    // Return empty if folder/file not found
  }
  return "";
}

/**
 * ============================================================================
 * FILE 2: SheetLogic.gs
 * Laburnum CC - Selection Night Logic, Interactive Dialogs, Triggers & Webhooks
 * ============================================================================
 */


/**
 * Returns the bound spreadsheet, falling back to the stored ID for standalone scripts.
 * setupInitialSystem creates the sheet and stores its ID on first run.
 */
function getSS() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { }
  }
  return null;
}


/**
 * Parses a PlayHQ CSV export and appends new players to the Players tab.
 * Skips rows whose ProfileID already exists.
 */
function importPlayHQPlayers(csvContent) {
  var ss = getSS();
  var playerSheet = ss.getSheetByName("Players");
  if (!playerSheet) return "ERROR: Players tab not found.";

  var rows = Utilities.parseCsv(csvContent);
  if (rows.length < 2) return "ERROR: CSV has no data rows.";

  var headers = rows[0];
  var colIdx = {};
  headers.forEach(function(h, i) { colIdx[h.trim()] = i; });

  // Build set of existing ProfileIDs — scan column A only
  var lastDataRow = 1;
  var colAVals = playerSheet.getRange(1, 1, playerSheet.getMaxRows(), 1).getValues();
  for (var r = colAVals.length - 1; r >= 0; r--) {
    if (colAVals[r][0] && String(colAVals[r][0]).trim() !== "") {
      lastDataRow = r + 1;
      break;
    }
  }

  var existingIds = {};
  if (lastDataRow > 1) {
    playerSheet.getRange(2, 1, lastDataRow - 1, 1).getValues().forEach(function(r) {
      if (r[0]) existingIds[String(r[0]).trim()] = true;
    });
  }

  Logger.log("Sheet: " + ss.getUrl());
  Logger.log("Last data row (ProfileID scan): " + lastDataRow);
  Logger.log("Existing IDs found: " + Object.keys(existingIds).length);

  var newRows = [];
  var skipped = 0;

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var profileId = (row[colIdx["Profile ID"]] || "").trim();
    var firstName = (row[colIdx["First Name"]] || "").trim();
    var lastName = (row[colIdx["Last Name"]] || "").trim();
    var preferredName = (row[colIdx["Preferred Name"]] || "").trim();
    var ageGroup = (row[colIdx["Age Group"]] || "").trim();
    var phone1 = (row[colIdx["Account Holder Mobile"]] || "").trim();
    var phone2 = (row[colIdx["Parent/Guardian 1 Mobile"]] || "").trim();
    var phone3 = (row[colIdx["Parent/Guardian 2 Mobile"]] || "").trim();
    var email = (row[colIdx["Account Holder Email"]] || "").trim();

    if (!profileId || !firstName || !lastName) {
      skipped++;
      continue;
    }

    if (existingIds[profileId]) {
      skipped++;
      continue;
    }

    var displayFirstName = pickFirstName(firstName, preferredName);

    var juniorLevel = "";
    if (ageGroup === "U18") juniorLevel = "U18";
    else if (ageGroup === "U16") juniorLevel = "U16";
    else if (ageGroup === "U14") juniorLevel = "U14";

    existingIds[profileId] = true;

    newRows.push([
      profileId,                // A: ProfileID
      displayFirstName,         // B: FirstName
      lastName,                 // C: LastName
      displayFirstName + " " + lastName,  // D: FullName
      juniorLevel,              // E: JuniorLevel
      "",                       // F: T20Squad
      "Active",                 // G: GlobalStatus
      "",                       // H: ExpectedReturnDate
      normalizePhone(phone1),   // I: Phone (Account Holder Mobile)
      normalizePhone(phone2),   // J: Phone2 (Parent/Guardian 1 Mobile)
      normalizePhone(phone3),   // K: Phone3 (Parent/Guardian 2 Mobile)
      "",                       // L: Phone4
      email                     // M: Email (Account Holder Email)
    ]);
  }

  Logger.log("New rows to write: " + newRows.length);
  Logger.log("Skipped (existing/incomplete): " + skipped);

  if (newRows.length === 0) {
    return "No new players found. " + skipped + " skipped.\n\nSheet: " + ss.getUrl();
  }

  var nextRow = lastDataRow + 1;
  playerSheet.getRange(nextRow, 1, newRows.length, 13).setValues(newRows);

  // Apply generous, sensible column widths
  applyAllStandardColumnWidths(ss);

  return "Imported " + newRows.length + " new player(s). " + skipped + " skipped.\n\nSheet: " + ss.getUrl();
}


/**
 * Parses a PlayHQ Fixtures CSV export (Home or Away) and merges into Fixtures tab.
 * Uses team configurations from Config tab.
 */
function importFixturesCsv(csvContent) {
  var ss = getSS();
  if (!ss) return "ERROR: Spreadsheet not found.";

  var fixSheet = ss.getSheetByName("Fixtures") || ss.insertSheet("Fixtures");

  // 1. Read team configs from Config tab
  var configSheet = ss.getSheetByName("Config");
  var teamConfigs = DEFAULT_TEAM_CONFIGS;
  if (configSheet && configSheet.getLastRow() > 1) {
    var configData = configSheet.getDataRange().getValues();
    teamConfigs = readTeamConfigsFromSheet(configData);
  }

  // 2. Parse fixtures CSV
  var parsedGames = parseFixtureCsv(csvContent, teamConfigs);
  if (parsedGames.length === 0) {
    return "No matching fixture games found in CSV.\n\nPlease verify that the 'Play Cricket Team Name' entries in your Config tab match the team names in your export file.";
  }

  // 3. Read current fixtures matrix
  var existingMatrix = [];
  if (fixSheet.getLastRow() > 0) {
    existingMatrix = fixSheet.getDataRange().getValues();
  }

  // 4. Merge using pure business logic
  var mergedMatrix = mergeFixturesIntoMatrix(existingMatrix, parsedGames, teamConfigs);

  // 5. Write back to Fixtures tab
  fixSheet.clear();
  var numRows = mergedMatrix.length;
  var numCols = mergedMatrix[0].length;

  fixSheet.getRange(1, 1, numRows, numCols).setValues(mergedMatrix);

  // Format headers
  fixSheet.getRange(1, 1, 1, numCols)
    .setFontWeight("bold")
    .setBackground(LCC_PALETTE.maroonBg)
    .setFontColor(LCC_PALETTE.maroonFg);

  // Format borders
  if (numRows > 1) {
    fixSheet.getRange(1, 1, numRows, numCols)
      .setBorder(true, true, true, true, true, true, LCC_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);
    // Explicitly format Date column (Col A) as plain text YYYY-MM-DD
    fixSheet.getRange(2, 1, numRows - 1, 1).setNumberFormat("@");
  }

  // Apply generous, sensible column widths
  applyAllStandardColumnWidths(ss);

  var uniqueDates = {};
  parsedGames.forEach(function(g) { uniqueDates[g.startDate] = true; });
  var dateCount = Object.keys(uniqueDates).length;

  return "Successfully imported " + parsedGames.length + " match fixture(s) across " + dateCount + " game date(s).\nTotal scheduled round dates in Fixtures tab: " + (numRows - 1) + "\n\nSheet: " + ss.getUrl();
}


function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Shifts 3-cell blocks [ProfileID, Name, Notes] between list columns.
 */
function movePlayerRowsBetweenLists(sheet, sourceRow, sourceStartCol, destStartCol) {
  var playerData = sheet.getRange(sourceRow, sourceStartCol, 1, 3).getValues()[0];
  var profileId = playerData[0];
  var name = playerData[1];
  var note = playerData[2];

  if (!name && !profileId) return;

  var destLastRow = 5;
  var destNames = sheet.getRange(5, destStartCol + 1, 100, 1).getValues();
  for (var i = 0; i < destNames.length; i++) {
    if (destNames[i][0] !== "") destLastRow = 5 + i + 1;
    else break;
  }

  sheet.getRange(destLastRow, destStartCol, 1, 3).setValues([[profileId, name, note]]);
  sheet.getRange(sourceRow, sourceStartCol, 1, 4).clearContent();

  var currentActiveRows = [];
  var sourceNames = sheet.getRange(5, sourceStartCol + 1, 100, 1).getValues();

  for (var k = 0; k < sourceNames.length; k++) {
    var checkRow = 5 + k;
    var rowValues = sheet.getRange(checkRow, sourceStartCol, 1, 3).getValues()[0];
    if (rowValues[1] !== "") {
      currentActiveRows.push([rowValues[0], rowValues[1], rowValues[2], ""]);
    }
  }

  sheet.getRange(5, sourceStartCol, 100, 4).clearContent();
  if (currentActiveRows.length > 0) {
    sheet.getRange(5, sourceStartCol, currentActiveRows.length, 4).setValues(currentActiveRows);
  }
}


/**
 * Searches hidden ProfileID columns (Cols F, K, P) to route players.
 * Called by web app (doPost), simulateRoundAvailability, and injury/inactive dialogs.
 */
function executeRealtimeAppListMovement(ws, profileId, response, notes) {
  var colMap = [6, 11, 16]; // Col F (Available), Col K (Unavailable), Col P (Unknown)
  var foundRow = -1;
  var foundCol = -1;
  var targetProfileId = String(profileId || "").trim();

  // Find where the player currently sits across the 3 lists
  for (var cIdx = 0; cIdx < colMap.length; cIdx++) {
    var checkCol = colMap[cIdx];
    var idList = ws.getRange(5, checkCol, 100, 1).getValues();
    for (var rIdx = 0; rIdx < idList.length; rIdx++) {
      if (String(idList[rIdx][0]).trim() === targetProfileId) {
        foundRow = 5 + rIdx;
        foundCol = checkCol;
        break;
      }
    }
    if (foundRow !== -1) break;
  }

  // Determine destination column: Col 6 (Available), Col 11 (Unavailable), Col 16 (Unknown)
  var targetDestCol = 6;
  var respLower = String(response || "").toLowerCase();
  if (respLower === "unavailable" || respLower.indexOf("unavail") > -1) {
    targetDestCol = 11;
  } else if (respLower === "unknown") {
    targetDestCol = 16;
  } else {
    // "Available", "Both Days", "Day 1 Only", "Day 2 Only"
    targetDestCol = 6;
  }

  // If found in a list
  if (foundRow !== -1) {
    if (foundCol !== targetDestCol) {
      movePlayerRowsBetweenLists(ws, foundRow, foundCol, targetDestCol);
    }
    // Now locate in destination list and update the note (Col destStartCol + 2)
    var updatedIds = ws.getRange(5, targetDestCol, 100, 1).getValues();
    for (var k = 0; k < updatedIds.length; k++) {
      if (String(updatedIds[k][0]).trim() === targetProfileId) {
        if (notes !== undefined && notes !== null) {
          ws.getRange(5 + k, targetDestCol + 2).setValue(notes);
        }
        break;
      }
    }
  } else {
    // Player not in round tab yet — lookup in Players master tab and insert
    var ss = getSS();
    var playerSheet = ss.getSheetByName("Players");
    if (playerSheet) {
      var pData = playerSheet.getDataRange().getValues();
      var pHeaders = pData[0];
      var idIdx = pHeaders.indexOf("ProfileID");
      var nameIdx = pHeaders.indexOf("FullName");
      var firstIdx = pHeaders.indexOf("FirstName");
      var lastIdx = pHeaders.indexOf("LastName");
      var matchRow = pData.slice(1).find(function(r) { return String(r[idIdx]).trim() === targetProfileId; });
      if (matchRow) {
        var pName = matchRow[nameIdx] || (matchRow[firstIdx] + " " + matchRow[lastIdx]);
        var destNames = ws.getRange(5, targetDestCol + 1, 100, 1).getValues();
        var destRow = 5;
        for (var dIdx = 0; dIdx < destNames.length; dIdx++) {
          if (destNames[dIdx][0] !== "") destRow = 5 + dIdx + 1;
          else break;
        }
        ws.getRange(destRow, targetDestCol, 1, 3).setValues([[targetProfileId, pName, notes || ""]]);
      }
    }
  }
}



// ============================================================================
// SECTION 2: PLAYER-FACING WEB APP LOGIC
// ============================================================================

/**
 * Normalizes phone numbers to E164 format (+614...)
 * Pure logic — also lives in logic.js for Jest testability.
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



// ============================================================================
// SECTION 4: ON-EDIT TRIGGER & SHEET-LEVEL LIST MOVEMENT
// ============================================================================

/**
 * Intercepts emoji dropdown actions in Cols I, N, S.
 */
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var col = range.getColumn();
  var row = range.getRow();
  var value = e.value;

  if (sheet.getName() === "Presentation_Staging" || sheet.getName() === "Players" || sheet.getName() === "Fixtures") return;
  if (row < 5 || !value) return;

  if (col === 9 && value === "🚫 Move to Unavailable") {
    movePlayerRowsBetweenLists(sheet, row, 6, 11);
    range.setValue("");
  } else if (col === 14 && value === "✅ Move to Available") {
    movePlayerRowsBetweenLists(sheet, row, 11, 6);
    range.setValue("");
  } else if (col === 19) {
    if (value === "✅ Move to Available") movePlayerRowsBetweenLists(sheet, row, 16, 6);
    else if (value === "🚫 Move to Unavailable") movePlayerRowsBetweenLists(sheet, row, 16, 11);
    range.setValue("");
  }
}



// ============================================================================
// SECTION 3: TOOLBAR & SELECTOR DIALOGS
// ============================================================================

/**
 * Australised Toolbar Menu Hook.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🏏 LCC Selection")
    .addItem("Initialise new round", "showDatePickerDialog")
    .addSeparator()
    .addItem("Pull new players into round from Master Players list", "showPullNewPlayersDialog")
    .addItem("Mark player as inactive", "showMarkInactiveDialog")
    .addSeparator()
    .addItem("📸 Player Photo Studio", "showPhotoStudioDialog")
    .addItem("Sync player headshots from Google Drive", "menuSyncPlayerHeadshots")
    .addSeparator()
    .addItem("🚀 Sync to Google Slides presentation", "showSyncSlidesDialog")
    .addItem("Assign permanent IDs to Google Slides elements", "menuAutoTagSlides")
    .addItem("Inspect Google Slides layout", "menuInspectSlidesLayout")
    .addSeparator()
    .addItem("Import fixtures from PlayHQ export", "showImportFixturesDialog")
    .addItem("Import players from PlayHQ export", "showImportPlayHQDialog")
    .addToUi();

  // Auto-sync selection validation rules & WhatsApp columns on existing round sheets
  try {
    var ss = getSS();
    if (ss) {
      var sheets = ss.getSheets();
      var nonRoundNames = ["Players", "Fixtures", "Config", "Admins", "Presentation_Staging", "Availability_Log"];
      sheets.forEach(function(s) {
        if (nonRoundNames.indexOf(s.getName()) === -1) {
          applySelectionValidationRules(s);
          setupWhatsAppMessageColumns(s);
          applyRoundTabColumnWidths(s);
        }
      });

      // Auto-sync Presentation Staging Hub and Headshots
      syncPlayerHeadshots(ss);
      syncPresentationStagingHub(ss);
    }
  } catch (e) {
    Logger.log("onOpen auto-sync warning: " + e.message);
  }
}


/**
 * Applies per-slot dynamic unselected dropdown validation to a round sheet.
 * For every individual slot (Row R), its dynamic dropdown list is:
 * =IFERROR(FILTER(G$5:G, (COUNTIF($B$8:$B$92, G$5:G)=0) + (G$5:G=B{R})), "")
 *
 * Result:
 * 1. A selected player is immediately removed from all other dropdown choices.
 * 2. The player in cell B{R} remains valid within their own slot -> ZERO RED ERROR TRIANGLES!
 * 3. Prevents selecting the same player multiple times anywhere on the sheet.
 */
function applySelectionValidationRules(ws) {
  if (!ws) return;
  try {
    // Col D: Global unselected pool (Alphabetical)
    ws.getRange("D5").setFormula(`=IFERROR(SORT(FILTER(G5:G, G5:G<>"", COUNTIF(B:B, G5:G)=0)), "")`);

    var teamStarts = [9, 28, 47, 66, 85];
    var maxCols = ws.getMaxColumns();
    if (maxCols < 95) {
      ws.insertColumnsAfter(maxCols, 95 - maxCols);
    }

    var startCol = 27; // Col AA (Col 27) through Col CQ (Col 91)
    var slotIndex = 0;

    teamStarts.forEach(function(startRow) {
      for (var s = 0; s < 13; s++) {
        var row = startRow + s;
        var helperCol = startCol + slotIndex;
        slotIndex++;

        // Set helper header and dynamic formula (Alphabetical A-Z)
        ws.getRange(3, helperCol).setValue("SLOT_" + row);
        ws.getRange(5, helperCol).setFormula(
          '=IFERROR(SORT(FILTER(G$5:G, (COUNTIF($B$9:$B$97, G$5:G)=0) + (G$5:G=B' + row + '))), "")'
        );

        // Apply Data Validation to this specific slot
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(ws.getRange(5, helperCol, 100, 1), true)
          .setAllowInvalid(true)
          .build();
        ws.getRange(row, 2).setDataValidation(rule);
      }
    });

    // Hide all 65 helper columns (Col 27 to Col 91)
    ws.hideColumns(27, 65);
  } catch (err) {
    Logger.log("applySelectionValidationRules error: " + err.message);
  }
}


/**
 * Menu action to refresh WhatsApp message columns on active sheet.
 */
function menuRefreshWhatsAppColumns() {
  var ss = getSS();
  if (!ss) return;
  var ws = ss.getActiveSheet();
  var nonRoundNames = ["Players", "Fixtures", "Config", "Admins", "Presentation_Staging", "Availability_Log"];
  if (nonRoundNames.indexOf(ws.getName()) > -1) {
    ss.toast("Please switch to an active round tab (e.g. 2025-10-18) first.", "LCC Selection", 4);
    return;
  }
  setupWhatsAppMessageColumns(ws);
  cleanConfigTabTemplates(ss);
  ss.toast("WhatsApp message columns U & V refreshed with compact row height!", "LCC Selection", 3);
}


/**
 * Cleans legacy WhatsApp template rows from Config tab.
 */
function cleanConfigTabTemplates(ss) {
  try {
    var s = ss || getSS();
    if (!s) return;
    var configSheet = s.getSheetByName("Config");
    if (!configSheet) return;
    var lastRow = configSheet.getLastRow();
    if (lastRow >= 11) {
      for (var r = 11; r <= lastRow; r++) {
        var val = String(configSheet.getRange(r, 1).getValue());
        if (val.indexOf("WHATSAPP") > -1 || val.indexOf("Availability Callout") > -1 || val.indexOf("Wall of Shame") > -1) {
          configSheet.getRange(r, 1, lastRow - r + 1, configSheet.getLastColumn()).clear();
          break;
        }
      }
    }
  } catch (e) {
    Logger.log("cleanConfigTabTemplates warning: " + e.message);
  }
}


/**
 * Sets up Column U and Column V on a round sheet for WhatsApp messages.
 * Uses merged ranges U4:U20 and V4:V20 so text flows downward without increasing Row 4 height!
 */
function setupWhatsAppMessageColumns(ws, dateStr) {
  if (!ws) return;
  var dStr = dateStr || ws.getName();

  // Column U (Col 21) Header & Merged Card
  ws.getRange("U3").setValue(dStr + " Availability Request Message")
    .setFontWeight("bold")
    .setBackground(LCC_PALETTE.maroonBg)
    .setFontColor(LCC_PALETTE.maroonFg)
    .setHorizontalAlignment("center");

  // Unmerge first if previously merged to allow clean formula assignment
  try {
    ws.getRange("U4:U20").breakApart();
  } catch (e) {}

  var uCard = ws.getRange("U4:U20");
  uCard.merge();
  uCard.setFormula(
    '="🏏 *LABURNUM CC ROUND AVAILABILITY* 🏏" & CHAR(10) & CHAR(10) & ' +
    '"Please submit your availability for the upcoming round (" & TEXT(B1, "yyyy-mm-dd") & "):" & CHAR(10) & ' +
    '"https://lcc-availability.web.app/?round=" & TEXT(B1, "yyyy-mm-dd")'
  ).setWrap(true).setVerticalAlignment("top");

  // Column V (Col 22) Header & Merged Card
  ws.getRange("V3").setValue(dStr + " Wall of Shame Message")
    .setFontWeight("bold")
    .setBackground(LCC_PALETTE.maroonBg)
    .setFontColor(LCC_PALETTE.maroonFg)
    .setHorizontalAlignment("center");

  try {
    ws.getRange("V4:V20").breakApart();
  } catch (e) {}

  var vCard = ws.getRange("V4:V20");
  vCard.merge();
  vCard.setFormula(
    '=LCC_WALL_OF_SHAME(B1, COUNTA(G5:G) + COUNTA(L5:L), Q5:Q150)'
  ).setWrap(true).setVerticalAlignment("top");

  // Card Borders and Background Styling
  try {
    ws.getRange("U3:U20").setBorder(true, true, true, true, false, false, LCC_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange("V3:V20").setBorder(true, true, true, true, false, false, LCC_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);
    uCard.setBackground(LCC_PALETTE.zebraLight);
    vCard.setBackground(LCC_PALETTE.zebraLight);
  } catch (e) {}

  // Explicitly reset Row 4 height to standard compact height (24px)
  try {
    ws.setRowHeight(4, 24);
  } catch (e) {}

  // Apply column widths
  applyRoundTabColumnWidths(ws);

  // Unhide Columns 21 & 22 if they were previously hidden
  try {
    ws.unhideColumn(ws.getRange("U1"));
    ws.unhideColumn(ws.getRange("V1"));
  } catch (e) {}
}


/**
 * Standard column width calculator for round selection tabs.
 * Switchers 57px, Notes 143px, Names 153px, Notes font size 9pt.
 */
function applyRoundTabColumnWidths(ws) {
  if (!ws) return;
  ws.setColumnWidth(1, 110); ws.setColumnWidth(2, 190); ws.setColumnWidth(3, 20);
  ws.setColumnWidth(4, 190); ws.setColumnWidth(5, 20);
  ws.setColumnWidth(6, 80);  ws.setColumnWidth(7, 153); ws.setColumnWidth(8, 143); ws.setColumnWidth(9, 57); ws.setColumnWidth(10, 20);
  ws.setColumnWidth(11, 80); ws.setColumnWidth(12, 153); ws.setColumnWidth(13, 143); ws.setColumnWidth(14, 57); ws.setColumnWidth(15, 20);
  ws.setColumnWidth(16, 80); ws.setColumnWidth(17, 153); ws.setColumnWidth(18, 143); ws.setColumnWidth(19, 57); ws.setColumnWidth(20, 20);
  ws.setColumnWidth(21, 330); // Col U (Availability Message)
  ws.setColumnWidth(22, 330); // Col V (Wall of Shame Message)
  ws.setColumnWidth(23, 20);  // Col W (Spacer)

  // Update header text to "Switcher" if existing
  try {
    if (String(ws.getRange("I4").getValue()).indexOf("Switcher") > -1) ws.getRange("I4").setValue("Switcher");
    if (String(ws.getRange("N4").getValue()).indexOf("Switcher") > -1) ws.getRange("N4").setValue("Switcher");
    if (String(ws.getRange("S4").getValue()).indexOf("Switcher") > -1) ws.getRange("S4").setValue("Switcher");
  } catch (e) {}

  // Format Notes values (Row 5 and below) with 9pt font
  try {
    ws.getRange("H5:H150").setFontSize(9);
    ws.getRange("M5:M150").setFontSize(9);
    ws.getRange("R5:R150").setFontSize(9);
  } catch (e) {}
}


/**
 * Custom function for Google Sheets to generate the Wall of Shame WhatsApp message.
 *
 * @param {Date|string} roundDate - The round date cell (e.g. B2)
 * @param {number} declaredCount - Number of declared players (e.g. COUNTA(G5:G) + COUNTA(L5:L))
 * @param {Array<Array<string>>} unknownRange - Range of unknown player names (e.g. Q5:Q150)
 * @returns {string} Formatted WhatsApp message
 * @customfunction
 */
function LCC_WALL_OF_SHAME(roundDate, declaredCount, unknownRange) {
  var names = [];
  if (Array.isArray(unknownRange)) {
    unknownRange.forEach(function(row) {
      if (Array.isArray(row)) {
        if (row[0] && String(row[0]).trim()) names.push(String(row[0]).trim());
      } else if (row && String(row).trim()) {
        names.push(String(row).trim());
      }
    });
  }
  return generateWallOfShameMessage(roundDate, declaredCount, names);
}


/**
 * Custom function for Google Sheets to generate the Initial Availability WhatsApp message.
 *
 * @param {Date|string} roundDate - The round date cell (e.g. B2)
 * @returns {string} Formatted WhatsApp message
 * @customfunction
 */
function LCC_AVAILABILITY_MESSAGE(roundDate) {
  return generateAvailabilityCalloutMessage(roundDate);
}


/**
 * Custom function for Google Sheets to clean player names and strip junior tags for presentation.
 *
 * @param {string} name - The player name cell
 * @returns {string} Clean player name without junior tags
 * @customfunction
 */
function LCC_CLEAN_PLAYER_NAME(name) {
  if (!name) return "";
  return stripJuniorTag(String(name));
}


/**
 * Custom function for Google Sheets to retrieve public photo thumbnail URL from Google Drive.
 *
 * @param {string} profileId - Player profile ID (e.g. PL-001)
 * @returns {string} Public Drive thumbnail URL
 * @customfunction
 */
function LCC_PLAYER_PHOTO_URL(profileId) {
  if (!profileId) return "";
  return getPlayerPhotoUrl(String(profileId).trim());
}


/**
 * Fully builds or updates the Presentation_Staging sheet layout, formulas, and dropdown.
 *
 * @param {Spreadsheet} ss - Bound SpreadsheetApp instance
 * @param {string} [targetRoundName] - Optional round name to select in B1
 */
function syncPresentationStagingHub(ss, targetRoundName) {
  try {
    var s = ss || getSS();
    if (!s) return;
    var sh = s.getSheetByName("Presentation_Staging") || s.insertSheet("Presentation_Staging");
    
    // Rebuild/refresh Presentation_Staging layout and formulas
    sh.clear();
    
    // Row 1: Header and Selection Controls
    sh.getRange("A1").setValue("Round to present").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
    sh.getRange("B1").setValue("").setNumberFormat("@").setBackground(LCC_PALETTE.inputHighlight).setFontWeight("bold").setHorizontalAlignment("center");
      
      var frames = [
        { name: "FIRST ELEVEN", start: 4 }, 
        { name: "SECOND ELEVEN", start: 23 }, 
        { name: "THIRD ELEVEN", start: 42 }, 
        { name: "FOURTH ELEVEN", start: 61 }, 
        { name: "FIFTH ELEVEN", start: 80 }
      ];
      
      var slotRoles = [
        "1. Captain", "2. VC", "3. WK",
        "4. Player", "5. Player", "6. Player",
        "7. Player", "8. Player", "9. Player",
        "10. Player", "11. Player", "12. Player", "13. Player"
      ];
      
      var targetTabExpr = 'IF(ISNUMBER($B$1), TEXT($B$1, "yyyy-mm-dd"), TO_TEXT($B$1))';
      
      frames.forEach(function(f) {
        var rowIdx = f.start;
        
        // Team Banner
        sh.getRange(rowIdx, 1, 1, 4).merge().setValue(f.name).setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
        
        // Metadata rows: Round, Opponent, Venue, Format (Exact 1-to-1 match with Round sheet)
        sh.getRange(rowIdx + 1, 1).setValue("Round:").setFontStyle("italic");
        sh.getRange(rowIdx + 1, 2).setFormula('=IFERROR(INDIRECT("\'" & ' + targetTabExpr + ' & "\'!B' + (rowIdx + 1) + '"), "")');
        
        sh.getRange(rowIdx + 2, 1).setValue("Opponent:").setFontStyle("italic");
        sh.getRange(rowIdx + 2, 2).setFormula('=IFERROR(INDIRECT("\'" & ' + targetTabExpr + ' & "\'!B' + (rowIdx + 2) + '"), "")');
        
        sh.getRange(rowIdx + 3, 1).setValue("Venue:").setFontStyle("italic");
        sh.getRange(rowIdx + 3, 2).setFormula('=IFERROR(INDIRECT("\'" & ' + targetTabExpr + ' & "\'!B' + (rowIdx + 3) + '"), "")');
        
        sh.getRange(rowIdx + 4, 1).setValue("Format:").setFontStyle("italic");
        sh.getRange(rowIdx + 4, 2).setFormula('=IFERROR(INDIRECT("\'" & ' + targetTabExpr + ' & "\'!B' + (rowIdx + 4) + '"), "")');
        
        // 13 Player Slot Rows (start + 5 to start + 17) -> Exact 1-to-1 match with Round sheet
        for (var idx = 0; idx < 13; idx++) {
          var currRow = rowIdx + 5 + idx; // 9..21, 28..40, 47..59, 66..78, 85..97
          
          // Col A: Slot Role
          sh.getRange(currRow, 1).setValue(slotRoles[idx]).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);
          
          // Col B: Clean Player Name (Stripped of junior tags like (U16))
          sh.getRange(currRow, 2).setFormula('=IFERROR(IF(INDIRECT("\'" & ' + targetTabExpr + ' & "\'!B' + currRow + '")="", "", TRIM(REGEXREPLACE(TO_TEXT(INDIRECT("\'" & ' + targetTabExpr + ' & "\'!B' + currRow + '")), "\\s*\\(.*?\\)", ""))), "")');
          
          // Col C: Dynamic Profile ID Lookup from Players tab (Hidden)
          sh.getRange(currRow, 3).setFormula('=IFERROR(IF(B' + currRow + '="", "", INDEX(Players!$A$2:$A, MATCH(B' + currRow + ', Players!$D$2:$D, 0))), IFERROR(INDEX(Players!$A$2:$A, MATCH(B' + currRow + ', Players!$B$2:$B & " " & Players!$C$2:$C, 0)), ""))');
          
          // Col D: Dynamic Headshot Image from Players tab PhotoUrl
          sh.getRange(currRow, 4).setFormula('=IFERROR(IF(B' + currRow + '="", "", IMAGE(INDEX(Players!$N$2:$N, MATCH(C' + currRow + ', Players!$A$2:$A, 0)))), "")');
        }
        
        // Set borders for this team frame (18 rows: 1 banner + 4 metadata + 13 slots)
        sh.getRange(rowIdx, 1, 18, 4).setBorder(true, true, true, true, true, true, LCC_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);
      });
      
      sh.setColumnWidth(1, 120); // Col A (Slot Role)
      sh.setColumnWidth(2, 210); // Col B (Clean Player Name)
      sh.setColumnWidth(3, 20);  // Col C (Hidden Profile ID)
      sh.setColumnWidth(4, 100); // Col D (Photo)
      
      sh.hideColumns(3); // Hide Column C
      
      try {
        sh.getRange(1, 1, 100, 4).setFontFamily("Hanken Grotesk");
      } catch (e) {}
    
    // Sync dropdown validation list in B1 with all round tabs
    var nonRoundNames = ["Players", "Fixtures", "Config", "Admins", "Presentation_Staging", "Availability_Log"];
    var sheets = s.getSheets();
    var roundNames = [];
    sheets.forEach(function(ws) {
      var name = ws.getName();
      if (nonRoundNames.indexOf(name) === -1) {
        roundNames.push(name);
      }
    });
    
    if (roundNames.length > 0) {
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(roundNames, true)
        .setAllowInvalid(true)
        .build();
      sh.getRange("B1").setDataValidation(rule);
      
      var curVal = String(sh.getRange("B1").getValue()).trim();
      if (targetRoundName) {
        sh.getRange("B1").setValue(targetRoundName);
      } else if (!curVal || roundNames.indexOf(curVal) === -1) {
        sh.getRange("B1").setValue(roundNames[0]);
      }
    }
  } catch (err) {
    Logger.log("syncPresentationStagingHub error: " + err.message);
  }
}


/**
 * Scans the Google Drive headshots folder and syncs photo URLs into the Players tab (Col N: PhotoUrl).
 */
function syncPlayerHeadshots(ss) {
  var results = { folderFound: false, totalFiles: 0, matchedPlayers: 0, message: "" };
  try {
    var s = ss || getSS();
    if (!s) return results;
    var folderId = getHeadshotFolderId();
    if (!folderId) {
      results.message = "HEADSHOT_FOLDER_ID is not configured in Script Properties.";
      return results;
    }

    var folder = DriveApp.getFolderById(folderId);
    results.folderFound = true;
    var files = folder.getFiles();
    var photoMap = {};
    var defaultPhotoUrl = "";
    while (files.hasNext()) {
      var file = files.next();
      results.totalFiles++;
      var fileName = file.getName();
      var baseId = fileName.replace(/\.[^/.]+$/, "").trim().toLowerCase();
      var fileId = file.getId();

      // Google CDN direct link
      var photoUrl = "https://lh3.googleusercontent.com/d/" + fileId;
      if (baseId === "default" || baseId === "transparent" || baseId === "placeholder" || baseId === "blank" || baseId === "none") {
        defaultPhotoUrl = photoUrl;
      } else {
        photoMap[baseId] = photoUrl;
      }
    }

    if (defaultPhotoUrl) {
      PropertiesService.getScriptProperties().setProperty('DEFAULT_PHOTO_URL', defaultPhotoUrl);
    }

    var playerSheet = s.getSheetByName("Players");
    if (!playerSheet) return results;
    
    // Ensure Col N Header is "PhotoUrl"
    playerSheet.getRange(1, 14).setValue("PhotoUrl").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg);
    
    var lastRow = playerSheet.getLastRow();
    if (lastRow > 1) {
      var profileIds = playerSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var photoColValues = [];
      for (var r = 0; r < profileIds.length; r++) {
        var pid = String(profileIds[r][0] || "").trim().toLowerCase();
        var url = photoMap[pid] || "";
        if (url) results.matchedPlayers++;
        photoColValues.push([url]);
      }
      playerSheet.getRange(2, 14, photoColValues.length, 1).setValues(photoColValues);
    }
    results.message = "Successfully synced " + results.matchedPlayers + " headshot(s) from Drive!";
  } catch (e) {
    results.message = "Error syncing headshots: " + e.message;
    Logger.log("syncPlayerHeadshots error: " + e.message);
  }
  return results;
}


/**
 * Menu action to manually sync player headshots from Google Drive.
 */
function menuSyncPlayerHeadshots() {
  var ss = getSS();
  if (!ss) return;
  var res = syncPlayerHeadshots(ss);
  syncPresentationStagingHub(ss);
  
  var ui = SpreadsheetApp.getUi();
  ui.alert("🏏 Headshot Sync", res.message + "\n\nFiles found in folder: " + res.totalFiles + "\nMatched players in database: " + res.matchedPlayers, ui.ButtonSet.OK);
}


/**
 * Menu action to re-apply standard column widths across all tabs.
 */
function menuResetColumnWidths() {
  var ss = getSS();
  if (!ss) return;
  applyAllStandardColumnWidths(ss);
  syncPlayerHeadshots(ss);
  syncPresentationStagingHub(ss);
  ss.toast("Standard column widths and Presentation Staging Hub refreshed.", "LCC Selection Engine", 4);
}


/**
 * Picks the Preferred Name over First Name if it's genuinely a nickname
 * (shorter than the legal first name, and not a redundant First+Last copy).
 */
function pickFirstName(firstName, preferredName) {
  if (!preferredName) return firstName;
  if (preferredName.length < firstName.length) return preferredName;
  return firstName;
}


/**
 * DIALOG 2: Pull new players from Master Players tab.
 */
function pullNewPlayersToActiveTab() {
  var ss = getSS();
  var ws = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (ws.getName() === "Presentation_Staging" || ws.getName() === "Players" || ws.getName() === "Fixtures") {
    ui.alert("Please open an active round selection tab first.");
    return;
  }

  var playerSheet = ss.getSheetByName("Players");
  var masterData = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, playerSheet.getLastColumn()).getValues();

  var existingIds = {};
  [6, 11, 16].forEach(function(col) {
    var ids = ws.getRange(5, col, 100, 1).getValues();
    ids.forEach(function(r) {
      if (r[0] !== "") existingIds[String(r[0]).trim()] = true;
    });
  });

  var newPlayersAdded = 0;
  masterData.forEach(function(row) {
    var profileId = String(row[0]).trim();
    var fullName = row[3] || (row[1] + " " + row[2]);
    var globalStatus = String(row[6]).trim();
    var juniorLevel = row[4] || "";

    if (profileId !== "" && fullName.trim() !== "" && globalStatus === "Active" && !existingIds[profileId]) {
      var destLastRow = 5;
      var destNames = ws.getRange(5, 17, 100, 1).getValues();
      for (var i = 0; i < destNames.length; i++) {
        if (destNames[i][0] !== "") destLastRow = 5 + i + 1;
        else break;
      }

      var displayName = formatNameWithJuniorTag(fullName, juniorLevel);
      ws.getRange(destLastRow, 16, 1, 3).setValues([[profileId, displayName, "Newly Added from Master"]]);
      newPlayersAdded++;
    }
  });

  ui.alert("Sync Complete", newPlayersAdded + " new player(s) pulled into Unknown Status list.", ui.ButtonSet.OK);
}


/**
 * Wipes the Players tab and rewrites headers. Use before re-importing.
 */
function resetPlayersTab() {
  var ss = getSS();
  var playerSheet = ss.getSheetByName("Players");
  if (!playerSheet) return "ERROR: Players tab not found.";

  playerSheet.clear();
  var playerHeaders = [[
    "ProfileID", "FirstName", "LastName", "FullName",
    "JuniorLevel", "T20Squad", "GlobalStatus", "ExpectedReturnDate",
    "Phone", "Phone2", "Phone3", "Phone4", "Email"
  ]];
  playerSheet.getRange(1, 1, 1, 13).setValues(playerHeaders)
    .setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg);

  return "Players tab cleared. Ready for import.\n\nSheet: " + ss.getUrl();
}


/**
 * Saves availability submitted from the Web App
 */
function saveAvailability(profileId, dateStr, response, notes) {
  try {
    const ss = getSS();
    let cleanDateStr = String(dateStr || '').trim();
    if (cleanDateStr.match(/^\d{4}-\d{2}-\d{2}/)) {
      cleanDateStr = cleanDateStr.slice(0, 10);
    }

    let ws = ss.getSheetByName(cleanDateStr);
    if (!ws && cleanDateStr) {
      const sheets = ss.getSheets();
      for (var s = 0; s < sheets.length; s++) {
        if (sheets[s].getName().indexOf(cleanDateStr) > -1) {
          ws = sheets[s];
          break;
        }
      }
    }

    if (!ws) {
      // Fallback: pick the first date/round selection tab
      const sheets = ss.getSheets();
      for (var s = 0; s < sheets.length; s++) {
        const sName = sheets[s].getName();
        if (sName !== 'Players' && sName !== 'Fixtures' && sName !== 'Config' && sName !== 'Admins' && sName !== 'Presentation_Staging' && sName !== 'Availability_Log') {
          ws = sheets[s];
          break;
        }
      }
    }

    if (!ws) {
      throw new Error(`Round selection tab for date ${dateStr} has not been initialised yet.`);
    }

    executeRealtimeAppListMovement(ws, String(profileId).trim(), response, notes);

    const logSheet = ss.getSheetByName('Availability_Log');
    if (logSheet) {
      logSheet.appendRow([new Date(), String(profileId).trim(), ws.getName(), response, notes || ""]);
    }

    return { status: "success", success: true, round: ws.getName() };
  } catch (e) {
    Logger.log("saveAvailability Error: " + e.message);
    throw e;
  }
}


/**
 * DIALOG 1: HTML5 Datepicker Modal Window.
 */
function showDatePickerDialog() {
  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><head>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;600;700;800&display=swap" rel="stylesheet">' +
    '<style>' +
    '  body { font-family: "Hanken Grotesk", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; color: #222; margin: 0; background: #fff; box-sizing: border-box; }' +
    '  input, button, select { font-family: inherit; }' +
    '  h3 { color: #6A1B29; margin-top: 0; margin-bottom: 6px; font-size: 16px; font-weight: 800; }' +
    '  p { font-size: 13px; color: #555; line-height: 1.4; margin: 0 0 12px; }' +
    '  .card { background: #fdfdfd; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; margin-bottom: 12px; }' +
    '  .card-title { font-weight: bold; font-size: 13px; color: #6A1B29; margin-bottom: 6px; display: flex; align-items: center; justify-content: space-between; }' +
    '  input[type=date] { width: 100%; font-size: 13px; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; }' +
    '  .btn { padding: 9px 16px; border-radius: 6px; border: none; font-weight: bold; cursor: pointer; font-size: 13px; transition: all 0.2s; }' +
    '  .btn-primary { background: #6A1B29; color: white; }' +
    '  .btn-primary:hover { background: #52131e; }' +
    '  .btn-secondary { background: #e5e7eb; color: #333; margin-right: 8px; }' +
    '  .actions { display: flex; justify-content: space-between; align-items: center; margin-top: 14px; }' +
    '  .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #6A1B29; border-radius: 50%; width: 36px; height: 36px; animation: spin 1s linear infinite; margin: 20px auto; }' +
    '  @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
    '  .result-box { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; padding: 12px; font-size: 13px; white-space: pre-wrap; color: #333; max-height: 180px; overflow-y: auto; box-sizing: border-box; }' +
    '</style>' +
    '</head><body>' +
    '  <h3>Initialise Round Selection Tab</h3>' +
    '  <p>Select the first match date for this round to spawn a standalone selection tab.</p>' +
    '' +
    '  <div id="inputSection">' +
    '    <div class="card">' +
    '      <div class="card-title"><span>Match Date (Day 1)</span></div>' +
    '      <input type="date" id="matchDate">' +
    '    </div>' +
    '    <div class="actions">' +
    '      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '      <button id="createBtn" class="btn btn-primary" onclick="submitDate()">Create Round Tab</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <div id="spinnerSection" style="display:none; text-align:center; padding:15px;">' +
    '    <div class="spinner"></div>' +
    '    <p style="font-weight:bold; color:#6A1B29;">Generating round selection tab...</p>' +
    '  </div>' +
    '' +
    '  <div id="resultSection" style="display:none;">' +
    '    <div id="resultMessage" class="result-box"></div>' +
    '    <div class="actions" style="justify-content: flex-end; margin-top: 12px;">' +
    '      <button class="btn btn-primary" onclick="google.script.host.close()">Close</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <script>' +
    '    function submitDate() {' +
    '      var d = document.getElementById("matchDate").value;' +
    '      if (!d) { alert("Please pick a valid calendar date."); return; }' +
    '      document.getElementById("inputSection").style.display = "none";' +
    '      document.getElementById("spinnerSection").style.display = "block";' +
    '      google.script.run' +
    '        .withSuccessHandler(function(msg) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = msg;' +
    '        })' +
    '        .withFailureHandler(function(err) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = "Creation failed: " + (err.message || err);' +
    '        })' +
    '        .executeTabDeployment(d);' +
    '    }' +
    '  </script>' +
    '</body></html>'
  ).setWidth(460).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Initialise Round Selection Tab");
}


/**
 * DIALOG: PlayHQ CSV import — file-picker modal.
 */
function showImportPlayHQDialog() {
  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><head>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;600;700;800&display=swap" rel="stylesheet">' +
    '<style>' +
    '  body { font-family: "Hanken Grotesk", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; color: #222; margin: 0; background: #fff; box-sizing: border-box; }' +
    '  input, button, select { font-family: inherit; }' +
    '  h3 { color: #6A1B29; margin-top: 0; margin-bottom: 6px; font-size: 16px; font-weight: 800; }' +
    '  p { font-size: 13px; color: #555; line-height: 1.4; margin: 0 0 12px; }' +
    '  .card { background: #fdfdfd; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; margin-bottom: 12px; }' +
    '  .card-title { font-weight: bold; font-size: 13px; color: #6A1B29; margin-bottom: 6px; display: flex; align-items: center; justify-content: space-between; }' +
    '  .badge { background: #eee; color: #666; font-size: 11px; padding: 2px 8px; border-radius: 10px; font-weight: bold; }' +
    '  .badge.ready { background: #dcfce7; color: #166534; }' +
    '  input[type=file] { width: 100%; font-size: 12px; cursor: pointer; }' +
    '  .btn { padding: 9px 16px; border-radius: 6px; border: none; font-weight: bold; cursor: pointer; font-size: 13px; transition: all 0.2s; }' +
    '  .btn-primary { background: #6A1B29; color: white; }' +
    '  .btn-primary:hover { background: #52131e; }' +
    '  .btn-secondary { background: #e5e7eb; color: #333; margin-right: 8px; }' +
    '  .actions { display: flex; justify-content: space-between; align-items: center; margin-top: 14px; }' +
    '  .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #6A1B29; border-radius: 50%; width: 36px; height: 36px; animation: spin 1s linear infinite; margin: 20px auto; }' +
    '  @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
    '  .result-box { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; padding: 12px; font-size: 13px; white-space: pre-wrap; color: #333; max-height: 180px; overflow-y: auto; box-sizing: border-box; }' +
    '</style>' +
    '</head><body>' +
    '  <h3>Import Players (PlayHQ)</h3>' +
    '  <p>Select the master registration export CSV from PlayHQ to register and update club players.</p>' +
    '' +
    '  <div id="inputSection">' +
    '    <div class="card">' +
    '      <div class="card-title"><span>PlayHQ Player Export</span><span id="fileBadge" class="badge">Select CSV</span></div>' +
    '      <input type="file" id="csvFile" accept=".csv">' +
    '    </div>' +
    '    <div class="actions">' +
    '      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '      <button id="importBtn" class="btn btn-primary" onclick="startImport()">Import Players</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <div id="spinnerSection" style="display:none; text-align:center; padding:15px;">' +
    '    <div class="spinner"></div>' +
    '    <p style="font-weight:bold; color:#6A1B29;">Importing players...</p>' +
    '  </div>' +
    '' +
    '  <div id="resultSection" style="display:none;">' +
    '    <div id="resultMessage" class="result-box"></div>' +
    '    <div class="actions" style="justify-content: flex-end; margin-top: 12px;">' +
    '      <button class="btn btn-primary" onclick="google.script.host.close()">Close</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <script>' +
    '    var fileInput = document.getElementById("csvFile");' +
    '    var fileBadge = document.getElementById("fileBadge");' +
    '' +
    '    function updateBadge() {' +
    '      var file = fileInput.files && fileInput.files[0];' +
    '      if (file) {' +
    '        fileBadge.className = "badge ready";' +
    '        fileBadge.innerText = "Ready (" + file.name + ")";' +
    '      } else {' +
    '        fileBadge.className = "badge";' +
    '        fileBadge.innerText = "Select CSV";' +
    '      }' +
    '    }' +
    '' +
    '    fileInput.addEventListener("change", updateBadge);' +
    '    setInterval(updateBadge, 500);' +
    '' +
    '    function startImport() {' +
    '      var file = fileInput.files && fileInput.files[0];' +
    '      if (!file) {' +
    '        alert("Please select a PlayHQ CSV file to import.");' +
    '        return;' +
    '      }' +
    '      document.getElementById("inputSection").style.display = "none";' +
    '      document.getElementById("spinnerSection").style.display = "block";' +
    '' +
    '      var reader = new FileReader();' +
    '      reader.onload = function(e) {' +
    '        google.script.run' +
    '          .withSuccessHandler(function(msg) {' +
    '            document.getElementById("spinnerSection").style.display = "none";' +
    '            document.getElementById("resultSection").style.display = "block";' +
    '            document.getElementById("resultMessage").innerText = msg;' +
    '          })' +
    '          .withFailureHandler(function(err) {' +
    '            document.getElementById("spinnerSection").style.display = "none";' +
    '            document.getElementById("resultSection").style.display = "block";' +
    '            document.getElementById("resultMessage").innerText = "Import failed: " + (err.message || err);' +
    '          })' +
    '          .importPlayHQPlayers(e.target.result);' +
    '      };' +
    '      reader.onerror = function() {' +
    '        document.getElementById("spinnerSection").style.display = "none";' +
    '        document.getElementById("resultSection").style.display = "block";' +
    '        document.getElementById("resultMessage").innerText = "Failed to read file.";' +
    '      };' +
    '      reader.readAsText(file);' +
    '    }' +
    '  </script>' +
    '</body></html>'
  ).setWidth(460).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "PlayHQ Import");
}


/**
 * DIALOG: PlayHQ Fixtures CSV import — phased Home and Away file picker modal.
 */
function showImportFixturesDialog() {
  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><head>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;600;700;800&display=swap" rel="stylesheet">' +
    '<style>' +
    '  body { font-family: "Hanken Grotesk", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; color: #222; margin: 0; background: #fff; box-sizing: border-box; }' +
    '  input, button, select { font-family: inherit; }' +
    '  h3 { color: #6A1B29; margin-top: 0; margin-bottom: 6px; font-size: 16px; font-weight: 800; }' +
    '  p { font-size: 13px; color: #555; line-height: 1.4; margin: 0 0 12px; }' +
    '  .card { background: #fdfdfd; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; margin-bottom: 12px; }' +
    '  .card-title { font-weight: bold; font-size: 13px; color: #6A1B29; margin-bottom: 6px; display: flex; align-items: center; justify-content: space-between; }' +
    '  .badge { background: #eee; color: #666; font-size: 11px; padding: 2px 8px; border-radius: 10px; font-weight: bold; }' +
    '  .badge.ready { background: #dcfce7; color: #166534; }' +
    '  input[type=file] { width: 100%; font-size: 12px; cursor: pointer; }' +
    '  .btn { padding: 9px 16px; border-radius: 6px; border: none; font-weight: bold; cursor: pointer; font-size: 13px; transition: all 0.2s; }' +
    '  .btn-primary { background: #6A1B29; color: white; }' +
    '  .btn-primary:hover { background: #52131e; }' +
    '  .btn-secondary { background: #e5e7eb; color: #333; margin-right: 8px; }' +
    '  .actions { display: flex; justify-content: space-between; align-items: center; margin-top: 14px; }' +
    '  .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #6A1B29; border-radius: 50%; width: 36px; height: 36px; animation: spin 1s linear infinite; margin: 20px auto; }' +
    '  @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
    '  .result-box { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; padding: 12px; font-size: 13px; white-space: pre-wrap; color: #333; max-height: 180px; overflow-y: auto; box-sizing: border-box; }' +
    '</style>' +
    '</head><body>' +
    '  <h3>Import Fixtures (PlayHQ)</h3>' +
    '  <p>Import season fixtures in two phases (Home & Away CSVs) or upload both at once.</p>' +
    '' +
    '  <div id="inputSection">' +
    '    <div class="card">' +
    '      <div class="card-title"><span>Phase 1: Home Fixtures</span><span id="homeBadge" class="badge">Optional</span></div>' +
    '      <input type="file" id="homeFile" accept=".csv">' +
    '    </div>' +
    '    <div class="card">' +
    '      <div class="card-title"><span>Phase 2: Away Fixtures</span><span id="awayBadge" class="badge">Optional</span></div>' +
    '      <input type="file" id="awayFile" accept=".csv">' +
    '    </div>' +
    '    <div class="actions">' +
    '      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '      <button id="importBtn" class="btn btn-primary" onclick="startImport()">Import Fixtures</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <div id="spinnerSection" style="display:none; text-align:center; padding:15px;">' +
    '    <div class="spinner"></div>' +
    '    <p id="spinnerText" style="font-weight:bold; color:#6A1B29;">Importing fixtures...</p>' +
    '  </div>' +
    '' +
    '  <div id="resultSection" style="display:none;">' +
    '    <div id="resultMessage" class="result-box"></div>' +
    '    <div class="actions" style="justify-content: flex-end; margin-top: 12px;">' +
    '      <button class="btn btn-primary" onclick="google.script.host.close()">Close</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <script>' +
    '    var homeInput = document.getElementById("homeFile");' +
    '    var awayInput = document.getElementById("awayFile");' +
    '    var homeBadge = document.getElementById("homeBadge");' +
    '    var awayBadge = document.getElementById("awayBadge");' +
    '' +
    '    function updateBadges() {' +
    '      var hFile = homeInput.files && homeInput.files[0];' +
    '      var aFile = awayInput.files && awayInput.files[0];' +
    '      if (hFile) {' +
    '        homeBadge.className = "badge ready";' +
    '        homeBadge.innerText = "Ready (" + hFile.name + ")";' +
    '      } else {' +
    '        homeBadge.className = "badge";' +
    '        homeBadge.innerText = "Optional";' +
    '      }' +
    '      if (aFile) {' +
    '        awayBadge.className = "badge ready";' +
    '        awayBadge.innerText = "Ready (" + aFile.name + ")";' +
    '      } else {' +
    '        awayBadge.className = "badge";' +
    '        awayBadge.innerText = "Optional";' +
    '      }' +
    '    }' +
    '' +
    '    homeInput.addEventListener("change", updateBadges);' +
    '    awayInput.addEventListener("change", updateBadges);' +
    '    setInterval(updateBadges, 500);' +
    '' +
    '    function startImport() {' +
    '      var hFile = homeInput.files && homeInput.files[0];' +
    '      var aFile = awayInput.files && awayInput.files[0];' +
    '' +
    '      if (!hFile && !aFile) {' +
    '        alert("Please select at least one CSV file (Home Fixtures or Away Fixtures) before importing.");' +
    '        return;' +
    '      }' +
    '' +
    '      document.getElementById("inputSection").style.display = "none";' +
    '      document.getElementById("spinnerSection").style.display = "block";' +
    '' +
    '      var allResults = [];' +
    '' +
    '      function doAwayImport() {' +
    '        if (aFile) {' +
    '          document.getElementById("spinnerText").innerText = "Processing Phase 2: Away Fixtures...";' +
    '          var aReader = new FileReader();' +
    '          aReader.onload = function(e) {' +
    '            google.script.run' +
    '              .withSuccessHandler(function(res) {' +
    '                allResults.push("[Phase 2 - Away Fixtures]:\\n" + res);' +
    '                finishImport();' +
    '              })' +
    '              .withFailureHandler(function(err) {' +
    '                allResults.push("[Phase 2 - Away Fixtures Error]:\\n" + (err.message || err));' +
    '                finishImport();' +
    '              })' +
    '              .importFixturesCsv(e.target.result);' +
    '          };' +
    '          aReader.onerror = function() {' +
    '            allResults.push("[Phase 2 - Away Fixtures Error]: Failed to read file.");' +
    '            finishImport();' +
    '          };' +
    '          aReader.readAsText(aFile);' +
    '        } else {' +
    '          finishImport();' +
    '        }' +
    '      }' +
    '' +
    '      function finishImport() {' +
    '        document.getElementById("spinnerSection").style.display = "none";' +
    '        document.getElementById("resultSection").style.display = "block";' +
    '        document.getElementById("resultMessage").innerText = allResults.join("\\n\\n---\\n\\n");' +
    '      }' +
    '' +
    '      if (hFile) {' +
    '        document.getElementById("spinnerText").innerText = "Processing Phase 1: Home Fixtures...";' +
    '        var hReader = new FileReader();' +
    '        hReader.onload = function(e) {' +
    '          google.script.run' +
    '            .withSuccessHandler(function(res) {' +
    '              allResults.push("[Phase 1 - Home Fixtures]:\\n" + res);' +
    '              doAwayImport();' +
    '            })' +
    '            .withFailureHandler(function(err) {' +
    '              allResults.push("[Phase 1 - Home Fixtures Error]:\\n" + (err.message || err));' +
    '              doAwayImport();' +
    '            })' +
    '            .importFixturesCsv(e.target.result);' +
    '        };' +
    '        hReader.onerror = function() {' +
    '          allResults.push("[Phase 1 - Home Fixtures Error]: Failed to read file.");' +
    '          doAwayImport();' +
    '        };' +
    '        hReader.readAsText(hFile);' +
    '      } else {' +
    '        doAwayImport();' +
    '      }' +
    '    }' +
    '  </script>' +
    '</body></html>'
  ).setWidth(460).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "PlayHQ Fixture Import");
}


/**
 * Returns round tabs list and active sheet context for the Pull Players modal.
 */
function getRoundTabsForDialog() {
  var ss = getSS();
  if (!ss) return { activeIsRound: false, activeTab: "", roundTabs: [] };
  var nonRoundNames = ["Players", "Fixtures", "Config", "Admins", "Presentation_Staging", "Availability_Log"];
  var activeTab = ss.getActiveSheet().getName();
  var activeIsRound = nonRoundNames.indexOf(activeTab) === -1;
  var roundTabs = [];
  ss.getSheets().forEach(function(s) {
    var name = s.getName();
    if (nonRoundNames.indexOf(name) === -1) {
      roundTabs.push(name);
    }
  });
  return {
    activeIsRound: activeIsRound,
    activeTab: activeTab,
    roundTabs: roundTabs
  };
}


/**
 * Returns sorted list of all players for modal dropdowns.
 */
function getPlayersForDropdown() {
  var ss = getSS();
  if (!ss) return [];
  var pSheet = ss.getSheetByName("Players");
  if (!pSheet || pSheet.getLastRow() < 2) return [];
  var data = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, pSheet.getLastColumn()).getValues();
  var list = [];
  data.forEach(function(row) {
    var pId = String(row[0]).trim();
    var fName = row[3] || (row[1] + " " + row[2]);
    var status = String(row[6] || "Active").trim();
    var junior = row[4] || "";
    if (pId && fName.trim()) {
      list.push({
        profileId: pId,
        name: formatNameWithJuniorTag(fName, junior),
        rawName: fName,
        status: status
      });
    }
  });
  list.sort(function(a, b) { return a.name.localeCompare(b.name); });
  return list;
}


/**
 * DIALOG 2: Pull new players modal dialog.
 */
function showPullNewPlayersDialog() {
  var ss = getSS();
  var nonRoundNames = ["Players", "Fixtures", "Config", "Admins", "Presentation_Staging", "Availability_Log"];
  var activeTab = ss ? ss.getActiveSheet().getName() : "";
  var activeIsRound = nonRoundNames.indexOf(activeTab) === -1;
  var roundTabs = [];
  if (ss) {
    ss.getSheets().forEach(function(s) {
      var name = s.getName();
      if (nonRoundNames.indexOf(name) === -1) roundTabs.push(name);
    });
  }

  var optionsHtml = "";
  if (roundTabs.length === 0) {
    optionsHtml = '<option value="">No round selection tabs found</option>';
  } else {
    roundTabs.forEach(function(tab) {
      var isSelected = (tab === activeTab || (activeIsRound && tab === activeTab)) ? " selected" : "";
      var label = tab + (tab === activeTab ? " (Active Tab)" : "");
      optionsHtml += '<option value="' + tab + '"' + isSelected + '>' + label + '</option>';
    });
  }

  var disabledAttr = roundTabs.length === 0 ? " disabled" : "";

  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><head>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;600;700;800&display=swap" rel="stylesheet">' +
    '<style>' +
    '  body { font-family: "Hanken Grotesk", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; color: #222; margin: 0; background: #fff; box-sizing: border-box; }' +
    '  input, button, select { font-family: inherit; }' +
    '  h3 { color: #6A1B29; margin-top: 0; margin-bottom: 6px; font-size: 16px; font-weight: 800; }' +
    '  p { font-size: 13px; color: #555; line-height: 1.4; margin: 0 0 12px; }' +
    '  .card { background: #fdfdfd; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; margin-bottom: 12px; font-size: 13px; color: #444; }' +
    '  .form-group { margin-bottom: 0; }' +
    '  .form-label { font-weight: bold; font-size: 12px; color: #6A1B29; margin-bottom: 6px; display: block; }' +
    '  select { width: 100%; font-size: 13px; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; background: #fff; }' +
    '  .btn { padding: 9px 16px; border-radius: 6px; border: none; font-weight: bold; cursor: pointer; font-size: 13px; transition: all 0.2s; }' +
    '  .btn-primary { background: #6A1B29; color: white; }' +
    '  .btn-primary:hover { background: #52131e; }' +
    '  .btn-secondary { background: #e5e7eb; color: #333; margin-right: 8px; }' +
    '  .actions { display: flex; justify-content: space-between; align-items: center; margin-top: 14px; }' +
    '  .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #6A1B29; border-radius: 50%; width: 36px; height: 36px; animation: spin 1s linear infinite; margin: 20px auto; }' +
    '  @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
    '  .result-box { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; padding: 12px; font-size: 13px; white-space: pre-wrap; color: #333; max-height: 180px; overflow-y: auto; box-sizing: border-box; }' +
    '</style>' +
    '</head><body>' +
    '  <h3>Pull New Players from Master</h3>' +
    '  <p>Scan the Master Players directory for new registrations and pull them into the Unknown status list on the selection tab.</p>' +
    '' +
    '  <div id="inputSection">' +
    '    <div class="card">' +
    '      <div class="form-group">' +
    '        <label class="form-label">Target Round Tab</label>' +
    '        <select id="roundTabSelect">' + optionsHtml + '</select>' +
    '      </div>' +
    '    </div>' +
    '    <div class="actions">' +
    '      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '      <button id="syncBtn" class="btn btn-primary" onclick="startSync()"' + disabledAttr + '>Pull Players</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <div id="spinnerSection" style="display:none; text-align:center; padding:15px;">' +
    '    <div class="spinner"></div>' +
    '    <p style="font-weight:bold; color:#6A1B29;">Syncing players from Master...</p>' +
    '  </div>' +
    '' +
    '  <div id="resultSection" style="display:none;">' +
    '    <div id="resultMessage" class="result-box"></div>' +
    '    <div class="actions" style="justify-content: flex-end; margin-top: 12px;">' +
    '      <button class="btn btn-primary" onclick="google.script.host.close()">Close</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <script>' +
    '    function startSync() {' +
    '      var selTab = document.getElementById("roundTabSelect").value;' +
    '      if (!selTab) { alert("Please select a round tab or initialise a round tab first."); return; }' +
    '      document.getElementById("inputSection").style.display = "none";' +
    '      document.getElementById("spinnerSection").style.display = "block";' +
    '      google.script.run' +
    '        .withSuccessHandler(function(msg) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = msg;' +
    '        })' +
    '        .withFailureHandler(function(err) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = "Sync failed: " + (err.message || err);' +
    '        })' +
    '        .syncNewPlayersToActiveTab(selTab);' +
    '    }' +
    '  </script>' +
    '</body></html>'
  ).setWidth(460).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Pull New Players");
}


/**
 * Backend logic to sync new players from Master to active selection tab.
 */
function syncNewPlayersToActiveTab(targetTabName) {
  var ss = getSS();
  var ws = targetTabName ? ss.getSheetByName(targetTabName) : ss.getActiveSheet();
  if (!ws) throw new Error("Round tab '" + targetTabName + "' not found.");
  var forbiddenTabs = ["Presentation_Staging", "Players", "Fixtures", "Config", "Availability_Log", "Admins"];
  if (forbiddenTabs.indexOf(ws.getName()) > -1) {
    throw new Error("Please switch to an active round selection tab (e.g. 2025-10-04) first.");
  }

  var playerSheet = ss.getSheetByName("Players");
  if (!playerSheet || playerSheet.getLastRow() < 2) {
    throw new Error("No players found in Master Players tab.");
  }

  var masterData = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, playerSheet.getLastColumn()).getValues();
  var existingIds = {};
  [6, 11, 16].forEach(function(col) {
    var ids = ws.getRange(5, col, 100, 1).getValues();
    ids.forEach(function(r) {
      if (r[0] !== "") existingIds[String(r[0]).trim()] = true;
    });
  });

  var newPlayersAdded = 0;
  masterData.forEach(function(row) {
    var profileId = String(row[0]).trim();
    var fullName = row[3] || (row[1] + " " + row[2]);
    var globalStatus = String(row[6]).trim();
    var juniorLevel = row[4] || "";

    if (profileId !== "" && fullName.trim() !== "" && globalStatus === "Active" && !existingIds[profileId]) {
      var destLastRow = 5;
      var destNames = ws.getRange(5, 17, 100, 1).getValues();
      for (var i = 0; i < destNames.length; i++) {
        if (destNames[i][0] !== "") destLastRow = 5 + i + 1;
        else break;
      }

      var displayName = formatNameWithJuniorTag(fullName, juniorLevel);
      ws.getRange(destLastRow, 16, 1, 3).setValues([[profileId, displayName, "Newly Added from Master"]]);
      newPlayersAdded++;
    }
  });

  // Update Roster Last Synced timestamp in G1 / H1
  try {
    ws.getRange("G1").setValue("Roster Last Synced:").setFontWeight("bold");
    ws.getRange("H1").setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm").setFontColor("#666666");
  } catch (e) {}

  return "Sync Complete for tab '" + ws.getName() + "'.\n\n" + newPlayersAdded + " new player(s) pulled into Unknown Status list.";
}


/**
 * DIALOG 3: Record player injury/absence modal.
 */
function showRecordInjuryDialog() {
  var players = getPlayersForDropdown();
  var playerOptionsHtml = "";
  if (players.length === 0) {
    playerOptionsHtml = '<option value="">No players found in Master</option>';
  } else {
    players.forEach(function(p) {
      var tag = p.status !== "Active" ? " (" + p.status + ")" : "";
      playerOptionsHtml += '<option value="' + p.profileId + '">' + p.name + tag + '</option>';
    });
  }

  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><head>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;600;700;800&display=swap" rel="stylesheet">' +
    '<style>' +
    '  body { font-family: "Hanken Grotesk", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; color: #222; margin: 0; background: #fff; box-sizing: border-box; }' +
    '  input, button, select { font-family: inherit; }' +
    '  h3 { color: #6A1B29; margin-top: 0; margin-bottom: 6px; font-size: 16px; font-weight: 800; }' +
    '  p { font-size: 13px; color: #555; line-height: 1.4; margin: 0 0 12px; }' +
    '  .card { background: #fdfdfd; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; margin-bottom: 12px; }' +
    '  .form-group { margin-bottom: 10px; }' +
    '  .form-label { font-weight: bold; font-size: 12px; color: #6A1B29; margin-bottom: 4px; display: block; }' +
    '  select, input[type=text], input[type=date] { width: 100%; font-size: 13px; padding: 7px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; background: #fff; }' +
    '  .btn { padding: 9px 16px; border-radius: 6px; border: none; font-weight: bold; cursor: pointer; font-size: 13px; transition: all 0.2s; }' +
    '  .btn-primary { background: #6A1B29; color: white; }' +
    '  .btn-primary:hover { background: #52131e; }' +
    '  .btn-secondary { background: #e5e7eb; color: #333; margin-right: 8px; }' +
    '  .actions { display: flex; justify-content: space-between; align-items: center; margin-top: 14px; }' +
    '  .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #6A1B29; border-radius: 50%; width: 36px; height: 36px; animation: spin 1s linear infinite; margin: 20px auto; }' +
    '  @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
    '  .result-box { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; padding: 12px; font-size: 13px; white-space: pre-wrap; color: #333; max-height: 180px; overflow-y: auto; box-sizing: border-box; }' +
    '</style>' +
    '</head><body>' +
    '  <h3>Record Player Injury / Absence</h3>' +
    '  <p>Update player status to Injured in Master directory and mark them Unavailable on the active round.</p>' +
    '' +
    '  <div id="inputSection">' +
    '    <div class="card">' +
    '      <div class="form-group">' +
    '        <label class="form-label">Select Player</label>' +
    '        <select id="playerSelect">' + playerOptionsHtml + '</select>' +
    '      </div>' +
    '      <div class="form-group">' +
    '        <label class="form-label">Injury / Absence Notes</label>' +
    '        <input type="text" id="notes" placeholder="e.g. Hamstring strain - 3 weeks">' +
    '      </div>' +
    '      <div class="form-group" style="margin-bottom:0;">' +
    '        <label class="form-label">Expected Return Date (Optional)</label>' +
    '        <input type="date" id="returnDate">' +
    '      </div>' +
    '    </div>' +
    '    <div class="actions">' +
    '      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '      <button class="btn btn-primary" onclick="submitInjury()">Record Injury</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <div id="spinnerSection" style="display:none; text-align:center; padding:15px;">' +
    '    <div class="spinner"></div>' +
    '    <p style="font-weight:bold; color:#6A1B29;">Recording injury & updating rosters...</p>' +
    '  </div>' +
    '' +
    '  <div id="resultSection" style="display:none;">' +
    '    <div id="resultMessage" class="result-box"></div>' +
    '    <div class="actions" style="justify-content: flex-end; margin-top: 12px;">' +
    '      <button class="btn btn-primary" onclick="google.script.host.close()">Close</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <script>' +
    '    function submitInjury() {' +
    '      var pId = document.getElementById("playerSelect").value;' +
    '      var n = document.getElementById("notes").value.trim();' +
    '      var d = document.getElementById("returnDate").value.trim();' +
    '      if (!pId) { alert("Please select a player."); return; }' +
    '      document.getElementById("inputSection").style.display = "none";' +
    '      document.getElementById("spinnerSection").style.display = "block";' +
    '      google.script.run' +
    '        .withSuccessHandler(function(msg) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = msg;' +
    '        })' +
    '        .withFailureHandler(function(err) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = "Failed: " + (err.message || err);' +
    '        })' +
    '        .recordPlayerInjury(pId, n, d);' +
    '    }' +
    '  </script>' +
    '</body></html>'
  ).setWidth(460).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Record Injury / Absence");
}


/**
 * Backend logic to record player injury in Master and active tab.
 */
function recordPlayerInjury(query, notes, returnDate) {
  var ss = getSS();
  var ws = ss.getActiveSheet();
  var playerSheet = ss.getSheetByName("Players");
  if (!playerSheet || playerSheet.getLastRow() < 2) {
    throw new Error("Players tab not found or empty.");
  }

  var pData = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, playerSheet.getLastColumn()).getValues();
  var targetProfileId = "";
  var foundPlayerName = "";

  for (var i = 0; i < pData.length; i++) {
    var profileId = String(pData[i][0]).trim();
    var fName = pData[i][3] || (pData[i][1] + " " + pData[i][2]);
    if (profileId.toLowerCase() === String(query).toLowerCase() || fName.toLowerCase().indexOf(String(query).toLowerCase()) > -1) {
      targetProfileId = profileId;
      foundPlayerName = fName;
      var rowNum = 2 + i;
      playerSheet.getRange(rowNum, 7).setValue("Injured");
      if (returnDate) playerSheet.getRange(rowNum, 8).setValue(returnDate);
      break;
    }
  }

  if (!targetProfileId) {
    throw new Error("Player '" + query + "' not found in Master Players directory.");
  }

  var msg = "Recorded injury for " + foundPlayerName + " in Players directory.";
  var forbiddenTabs = ["Presentation_Staging", "Players", "Fixtures", "Config", "Availability_Log", "Admins"];
  if (forbiddenTabs.indexOf(ws.getName()) === -1) {
    executeRealtimeAppListMovement(ws, targetProfileId, "Unavailable", "Injured: " + (notes || ""));
    msg += "\n\nMoved " + foundPlayerName + " to 'Unavailable for Round' on active tab (" + ws.getName() + ").";
  }

  return msg;
}


/**
 * DIALOG 4: Mark player as inactive modal.
 */
function showMarkInactiveDialog() {
  var players = getPlayersForDropdown();
  var playerOptionsHtml = "";
  if (players.length === 0) {
    playerOptionsHtml = '<option value="">No players found in Master</option>';
  } else {
    players.forEach(function(p) {
      var tag = p.status !== "Active" ? " (" + p.status + ")" : "";
      playerOptionsHtml += '<option value="' + p.profileId + '">' + p.name + tag + '</option>';
    });
  }

  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><head>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;600;700;800&display=swap" rel="stylesheet">' +
    '<style>' +
    '  body { font-family: "Hanken Grotesk", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; color: #222; margin: 0; background: #fff; box-sizing: border-box; }' +
    '  input, button, select { font-family: inherit; }' +
    '  h3 { color: #6A1B29; margin-top: 0; margin-bottom: 6px; font-size: 16px; font-weight: 800; }' +
    '  p { font-size: 13px; color: #555; line-height: 1.4; margin: 0 0 12px; }' +
    '  .card { background: #fdfdfd; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; margin-bottom: 12px; }' +
    '  .form-group { margin-bottom: 6px; }' +
    '  .form-label { font-weight: bold; font-size: 12px; color: #6A1B29; margin-bottom: 4px; display: block; }' +
    '  select { width: 100%; font-size: 13px; padding: 7px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; background: #fff; }' +
    '  .btn { padding: 9px 16px; border-radius: 6px; border: none; font-weight: bold; cursor: pointer; font-size: 13px; transition: all 0.2s; }' +
    '  .btn-primary { background: #6A1B29; color: white; }' +
    '  .btn-primary:hover { background: #52131e; }' +
    '  .btn-secondary { background: #e5e7eb; color: #333; margin-right: 8px; }' +
    '  .actions { display: flex; justify-content: space-between; align-items: center; margin-top: 14px; }' +
    '  .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #6A1B29; border-radius: 50%; width: 36px; height: 36px; animation: spin 1s linear infinite; margin: 20px auto; }' +
    '  @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
    '  .result-box { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; padding: 12px; font-size: 13px; white-space: pre-wrap; color: #333; max-height: 180px; overflow-y: auto; box-sizing: border-box; }' +
    '</style>' +
    '</head><body>' +
    '  <h3>Mark Player as Inactive</h3>' +
    '  <p>Mark a player as Inactive in the Master directory (excluded from future availability callouts).</p>' +
    '' +
    '  <div id="inputSection">' +
    '    <div class="card">' +
    '      <div class="form-group">' +
    '        <label class="form-label">Select Player to Mark Inactive</label>' +
    '        <select id="playerSelect">' + playerOptionsHtml + '</select>' +
    '      </div>' +
    '    </div>' +
    '    <div class="actions">' +
    '      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '      <button class="btn btn-primary" onclick="submitInactive()">Mark Inactive</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <div id="spinnerSection" style="display:none; text-align:center; padding:15px;">' +
    '    <div class="spinner"></div>' +
    '    <p style="font-weight:bold; color:#6A1B29;">Updating player status...</p>' +
    '  </div>' +
    '' +
    '  <div id="resultSection" style="display:none;">' +
    '    <div id="resultMessage" class="result-box"></div>' +
    '    <div class="actions" style="justify-content: flex-end; margin-top: 12px;">' +
    '      <button class="btn btn-primary" onclick="google.script.host.close()">Close</button>' +
    '    </div>' +
    '  </div>' +
    '' +
    '  <script>' +
    '    function submitInactive() {' +
    '      var pId = document.getElementById("playerSelect").value;' +
    '      if (!pId) { alert("Please select a player."); return; }' +
    '      document.getElementById("inputSection").style.display = "none";' +
    '      document.getElementById("spinnerSection").style.display = "block";' +
    '      google.script.run' +
    '        .withSuccessHandler(function(msg) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = msg;' +
    '        })' +
    '        .withFailureHandler(function(err) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = "Failed: " + (err.message || err);' +
    '        })' +
    '        .markPlayerInactive(pId);' +
    '    }' +
    '  </script>' +
    '</body></html>'
  ).setWidth(460).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Mark Player Inactive");
}


/**
 * Backend logic to mark a player inactive.
 */
function markPlayerInactive(query) {
  var ss = getSS();
  var ws = ss.getActiveSheet();
  var playerSheet = ss.getSheetByName("Players");
  if (!playerSheet || playerSheet.getLastRow() < 2) {
    throw new Error("Players tab not found or empty.");
  }

  var pData = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, playerSheet.getLastColumn()).getValues();
  var targetProfileId = "";
  var foundPlayerName = "";

  for (var i = 0; i < pData.length; i++) {
    var profileId = String(pData[i][0]).trim();
    var fName = pData[i][3] || (pData[i][1] + " " + pData[i][2]);
    if (profileId.toLowerCase() === String(query).toLowerCase() || fName.toLowerCase().indexOf(String(query).toLowerCase()) > -1) {
      targetProfileId = profileId;
      foundPlayerName = fName;
      playerSheet.getRange(2 + i, 7).setValue("Inactive");
      break;
    }
  }

  if (!targetProfileId) {
    throw new Error("Player '" + query + "' not found in Master Players directory.");
  }

  var msg = "Marked " + foundPlayerName + " as Inactive in Players directory.";
  var forbiddenTabs = ["Presentation_Staging", "Players", "Fixtures", "Config", "Availability_Log", "Admins"];
  if (forbiddenTabs.indexOf(ws.getName()) === -1) {
    executeRealtimeAppListMovement(ws, targetProfileId, "Unavailable", "Marked Inactive");
    msg += "\n\nMoved " + foundPlayerName + " to 'Unavailable for Round' on active tab (" + ws.getName() + ").";
  }

  return msg;
}


/**
 * Diagnostic: prints the URL of the spreadsheet the script is using.
 * Run this from the editor to confirm you're looking at the right sheet.
 */
/**
 * ============================================================================
 * SECTION 11: GOOGLE SLIDES PRESENTATION AUTOMATION
 * ============================================================================
 */

/**
 * Helper to get the Google Slides presentation ID.
 */
function getSlidesPresentationId() {
  return PropertiesService.getScriptProperties().getProperty('SLIDES_PRESENTATION_ID') || '1Ma1zQiu7n8jnu34xueC2wTUHf3AsmNW-FPM_EuJFHtM';
}


/**
 * Diagnostically inspects the Google Slides structure, elements, and placeholders.
 */
function inspectSelectionSlides() {
  var presId = getSlidesPresentationId();
  if (!presId) {
    throw new Error("SLIDES_PRESENTATION_ID is not configured.");
  }
  
  var pres = SlidesApp.openById(presId);
  var slides = pres.getSlides();
  var report = "Google Slides Inspection Report\n";
  report += "Presentation: " + pres.getName() + "\n";
  report += "Total Slides: " + slides.length + "\n\n";

  slides.forEach(function(slide, sIdx) {
    report += "=== SLIDE " + (sIdx + 1) + " (ID: " + slide.getObjectId() + ") ===\n";
    var elements = slide.getPageElements();
    report += "Elements count: " + elements.length + "\n";
    elements.forEach(function(el, eIdx) {
      var type = el.getPageElementType();
      var pos = " (x: " + Math.round(el.getLeft()) + ", y: " + Math.round(el.getTop()) + ", w: " + Math.round(el.getWidth()) + ", h: " + Math.round(el.getHeight()) + ")";
      var desc = "  [" + (eIdx + 1) + "] " + type + pos;
      if (type === SlidesApp.PageElementType.SHAPE) {
        var shape = el.asShape();
        if (shape.getText()) {
          var t = shape.getText().asString().trim();
          desc += " | Text: '" + t.replace(/\n/g, " ↵ ") + "'";
        }
      } else if (type === SlidesApp.PageElementType.TABLE) {
        var tbl = el.asTable();
        desc += " | Table (" + tbl.getNumRows() + "x" + tbl.getNumColumns() + ")";
        var cells = [];
        for (var r = 0; r < tbl.getNumRows(); r++) {
          for (var c = 0; c < tbl.getNumColumns(); c++) {
            var ct = tbl.getCell(r, c).getText().asString().trim();
            if (ct) cells.push("[" + r + "," + c + "]: '" + ct + "'");
          }
        }
        if (cells.length > 0) desc += " | Cells: " + cells.join(", ");
      } else if (type === SlidesApp.PageElementType.IMAGE) {
        desc += " | Image";
      } else if (type === SlidesApp.PageElementType.GROUP) {
        desc += " | Group (" + el.asGroup().getChildren().length + " items)";
      }
      report += desc + "\n";
    });
    report += "\n";
  });

  return report;
}


/**
 * Menu action to run inspection and show in a modal dialog.
 */
function menuInspectSlidesLayout() {
  try {
    var report = inspectSelectionSlides();
    var html = HtmlService.createHtmlOutput(
      '<pre style="font-family: monospace; font-size: 11px; white-space: pre-wrap; word-break: break-all; height: 420px; overflow-y: scroll; background: #f9f9f9; padding: 12px; border: 1px solid #ccc; border-radius: 4px;">' +
      report.replace(/</g, '&lt;').replace(/>/g, '&gt;') +
      '</pre>' +
      '<div style="text-align: right; margin-top: 10px;"><button style="background: #4d0012; color: #fff; border: none; padding: 8px 16px; border-radius: 4px; font-weight: bold; cursor: pointer;" onclick="google.script.host.close()">Close</button></div>'
    ).setWidth(650).setHeight(520);
    SpreadsheetApp.getUi().showModalDialog(html, "📊 Google Slides Inspection");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Google Slides Inspection Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/**
 * Helper to sort slide elements into 2 vertical columns (Left: 1..6, Right: 7..12).
 */
function sortElementsTwoColumns(elements) {
  var leftCol = [];
  var rightCol = [];
  
  elements.forEach(function(el) {
    if (el.getLeft() < 360) {
      leftCol.push(el);
    } else {
      rightCol.push(el);
    }
  });

  leftCol.sort(function(a, b) { return a.getTop() - b.getTop(); });
  rightCol.sort(function(a, b) { return a.getTop() - b.getTop(); });

  return leftCol.concat(rightCol);
}


/**
 * Automatically assigns permanent Alt Text IDs (PLAYER_1..13, PHOTO_1..13, ROUND, etc.) to slide elements.
 */
function autoTagSlideElements(ss) {
  var presId = getSlidesPresentationId();
  if (!presId) throw new Error("SLIDES_PRESENTATION_ID is not configured.");

  var pres = SlidesApp.openById(presId);
  var slides = pres.getSlides();

  var frames = [
    { name: "FIRST ELEVEN", prefix: "1ST", slideIdx: 1 }, 
    { name: "SECOND ELEVEN", prefix: "2ND", slideIdx: 2 }, 
    { name: "THIRD ELEVEN", prefix: "3RD", slideIdx: 3 }, 
    { name: "FOURTH ELEVEN", prefix: "4TH", slideIdx: 4 }, 
    { name: "FIFTH ELEVEN", prefix: "5TH", slideIdx: 5 }
  ];

  var taggedCount = 0;
  frames.forEach(function(f) {
    if (f.slideIdx >= slides.length) return;
    var slide = slides[f.slideIdx];
    var elements = slide.getPageElements();

    var textShapes = [];
    var avatarImages = [];

    elements.forEach(function(el) {
      var type = el.getPageElementType();
      if (type === SlidesApp.PageElementType.SHAPE) {
        var shp = el.asShape();
        var txt = shp.getText() ? shp.getText().asString().trim() : "";
        if (txt.toLowerCase() !== f.name.toLowerCase() &&
            txt.toLowerCase().indexOf("first eleven") === -1 &&
            txt.toLowerCase().indexOf("second eleven") === -1 &&
            txt.toLowerCase().indexOf("third eleven") === -1 &&
            txt.toLowerCase().indexOf("fourth eleven") === -1 &&
            txt.toLowerCase().indexOf("fifth eleven") === -1 &&
            txt.indexOf("{{") === -1 &&
            shp.getTop() < 420) {
          textShapes.push(el);
        }
      } else if (type === SlidesApp.PageElementType.IMAGE) {
        // Only small avatar images, exclude large central club crest
        if (el.getWidth() < 120 && el.getHeight() < 120) {
          avatarImages.push(el);
        }
      }
    });

    var sortedShapes = sortElementsTwoColumns(textShapes);
    var sortedImages = sortElementsTwoColumns(avatarImages);

    sortedShapes.forEach(function(shp, idx) {
      shp.setTitle("PLAYER_" + (idx + 1));
      shp.setDescription(f.prefix + "_PLAYER_" + (idx + 1));
      taggedCount++;
    });

    sortedImages.forEach(function(img, idx) {
      img.setTitle("PHOTO_" + (idx + 1));
      img.setDescription(f.prefix + "_PHOTO_" + (idx + 1));
      taggedCount++;
    });
  });

  return "Successfully assigned permanent Alt Text IDs to " + taggedCount + " elements across team slides!";
}


/**
 * Menu action to automatically assign permanent IDs to slide elements.
 */
function menuAutoTagSlides() {
  var ss = getSS();
  if (!ss) return;
  var ui = SpreadsheetApp.getUi();
  try {
    var res = autoTagSlideElements(ss);
    ui.alert("🏏 Slide Element IDs Assigned", res + "\n\nAll text boxes and photo cards now have permanent IDs (PLAYER_1..13, PHOTO_1..13) that persist across all rounds!", ui.ButtonSet.OK);
  } catch (err) {
    ui.alert("Auto-Tag Error", err.message, ui.ButtonSet.OK);
  }
}


/**
 * Formats a player name with presentation role tags: (C), (VC), (Wk).
 */
function formatPlayerPresentationName(name, role) {
  if (!name || name.trim() === "") return "";
  var cleanName = name.replace(/\s*\(.*?\)/g, "").trim();
  var rLower = (role || "").toLowerCase();
  
  var roleTags = [];
  if (rLower.indexOf("captain") > -1 || rLower.indexOf("1. captain") > -1) {
    roleTags.push("(C)");
  } else if (rLower.indexOf("vc") > -1 || rLower.indexOf("vice") > -1) {
    roleTags.push("(VC)");
  } else if (rLower.indexOf("wk") > -1 || rLower.indexOf("keeper") > -1) {
    roleTags.push("(Wk)");
  }
  
  if (roleTags.length > 0) {
    return cleanName + " " + roleTags.join(" ");
  }
  return cleanName;
}


/**
 * Synchronizes Presentation_Staging team rosters and metadata directly to Google Slides using permanent Alt Text IDs.
 */
function syncPresentationStagingToSlides(ss) {
  var s = ss || getSS();
  if (!s) return { success: false, message: "Spreadsheet not found." };
  
  var presId = getSlidesPresentationId();
  if (!presId) return { success: false, message: "SLIDES_PRESENTATION_ID is not configured." };

  var staging = s.getSheetByName("Presentation_Staging");
  if (!staging) return { success: false, message: "Presentation_Staging tab not found." };

  // Read player photos map from Players tab
  var playerSheet = s.getSheetByName("Players");
  var photoUrlMap = {};
  if (playerSheet && playerSheet.getLastRow() > 1) {
    var pRows = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, 14).getValues();
    pRows.forEach(function(pr) {
      var pid = String(pr[0] || "").trim().toLowerCase();
      var url = String(pr[13] || "").trim(); // Col N
      if (pid && url) photoUrlMap[pid] = url;
    });
  }

  var pres = SlidesApp.openById(presId);
  var slides = pres.getSlides();

  // Read data for all 5 teams from Presentation_Staging
  var frames = [
    { name: "FIRST ELEVEN", prefix: "1ST", slideIdx: 1, start: 4 }, 
    { name: "SECOND ELEVEN", prefix: "2ND", slideIdx: 2, start: 23 }, 
    { name: "THIRD ELEVEN", prefix: "3RD", slideIdx: 3, start: 42 }, 
    { name: "FOURTH ELEVEN", prefix: "4TH", slideIdx: 4, start: 61 }, 
    { name: "FIFTH ELEVEN", prefix: "5TH", slideIdx: 5, start: 80 }
  ];

  var teamData = [];
  frames.forEach(function(f) {
    var roundVal = String(staging.getRange(f.start + 1, 2).getValue() || "").trim();
    var oppVal = String(staging.getRange(f.start + 2, 2).getValue() || "").trim();
    var venVal = String(staging.getRange(f.start + 3, 2).getValue() || "").trim();
    var fmtVal = String(staging.getRange(f.start + 4, 2).getValue() || "").trim();
    
    var players = [];
    for (var p = 0; p < 13; p++) {
      var pRow = f.start + 5 + p;
      var role = String(staging.getRange(pRow, 1).getValue() || "").trim();
      var name = String(staging.getRange(pRow, 2).getValue() || "").trim();
      var profileId = String(staging.getRange(pRow, 3).getValue() || "").trim();
      var photoUrl = photoUrlMap[profileId.toLowerCase()] || "";
      players.push({ role: role, name: name, profileId: profileId, photoUrl: photoUrl });
    }

    teamData.push({
      teamName: f.name,
      prefix: f.prefix,
      slideIdx: f.slideIdx,
      round: roundVal,
      opponent: oppVal,
      venue: venVal,
      format: fmtVal,
      players: players
    });
  });

  var transparentPngUrl = PropertiesService.getScriptProperties().getProperty('DEFAULT_PHOTO_URL') || "https://upload.wikimedia.org/wikipedia/commons/c/ca/1x1.png";

  // Update each team's slide using permanent Alt Text IDs & spatial fallbacks
  teamData.forEach(function(t) {
    if (t.slideIdx >= slides.length) return;
    var slide = slides[t.slideIdx];
    var elements = slide.getPageElements();

    // 1. Element-by-element Alt Text Tag Matching (No text replacement)
    var matchedPlayerTags = {};
    elements.forEach(function(el) {
      var tag = String(el.getTitle() || el.getDescription() || "").trim().toUpperCase();
      if (!tag) return;

      // Match info Alt Text tags
      if (tag === t.prefix + "_ROUND" || tag === "ROUND") {
        if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) el.asShape().getText().setText(t.round);
      } else if (tag === t.prefix + "_OPPONENT" || tag === "OPPONENT") {
        if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) el.asShape().getText().setText(t.opponent);
      } else if (tag === t.prefix + "_VENUE" || tag === "VENUE") {
        if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) el.asShape().getText().setText(t.venue);
      } else if (tag === t.prefix + "_FORMAT" || tag === "FORMAT") {
        if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) el.asShape().getText().setText(t.format);
      }

      // Match PLAYER_1 .. PLAYER_13
      var pMatch = tag.match(/^(?:.*_)?PLAYER_?(\d+)$/);
      if (pMatch) {
        var pNum = parseInt(pMatch[1], 10);
        if (pNum >= 1 && pNum <= t.players.length) {
          var pInfo = t.players[pNum - 1];
          matchedPlayerTags[pNum] = true;
          if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
            var formattedName = formatPlayerPresentationName(pInfo.name, pInfo.role);
            el.asShape().getText().setText(formattedName);
          }
        }
      }

      // Match PHOTO_1 .. PHOTO_13
      var photoMatch = tag.match(/^(?:.*_)?PHOTO_?(\d+)$/);
      if (photoMatch) {
        var photoNum = parseInt(photoMatch[1], 10);
        if (photoNum >= 1 && photoNum <= t.players.length) {
          var pPhoto = t.players[photoNum - 1];
          if (el.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
            try {
              if (pPhoto.photoUrl) {
                el.asImage().replace(pPhoto.photoUrl);
              } else {
                el.asImage().replace(transparentPngUrl);
              }
            } catch (err) {
              Logger.log("Image replace warning for " + pPhoto.name + ": " + err.message);
            }
          }
        }
      }
    });

    // 3. Fallback: Spatial two-column matching for player text shapes and avatar images
    var textShapes = [];
    var avatarImages = [];

    elements.forEach(function(el) {
      var type = el.getPageElementType();
      if (type === SlidesApp.PageElementType.SHAPE) {
        var shp = el.asShape();
        var txt = shp.getText() ? shp.getText().asString().trim() : "";
        if (txt.toLowerCase() !== t.teamName.toLowerCase() &&
            txt.toLowerCase().indexOf("first eleven") === -1 &&
            txt.toLowerCase().indexOf("second eleven") === -1 &&
            txt.toLowerCase().indexOf("third eleven") === -1 &&
            txt.toLowerCase().indexOf("fourth eleven") === -1 &&
            txt.toLowerCase().indexOf("fifth eleven") === -1 &&
            txt.indexOf("{{") === -1 &&
            shp.getTop() < 420) {
          textShapes.push(el);
        }
      } else if (type === SlidesApp.PageElementType.IMAGE) {
        // Exclude central Club Crest (width/height < 120)
        if (el.getWidth() < 120 && el.getHeight() < 120) {
          avatarImages.push(el);
        }
      }
    });

    var sortedShapes = sortElementsTwoColumns(textShapes);
    var sortedImages = sortElementsTwoColumns(avatarImages);

    // Apply player names if not already set by Alt text
    if (Object.keys(matchedPlayerTags).length === 0) {
      sortedShapes.forEach(function(shpEl, sIdx) {
        if (sIdx < t.players.length) {
          shpEl.setTitle("PLAYER_" + (sIdx + 1));
          shpEl.setDescription(t.prefix + "_PLAYER_" + (sIdx + 1));
          var formattedName = formatPlayerPresentationName(t.players[sIdx].name, t.players[sIdx].role);
          shpEl.asShape().getText().setText(formattedName);
        }
      });
    }

    // Apply player headshots or transparent fallback for all avatars
    sortedImages.forEach(function(imgEl, sIdx) {
      if (sIdx < t.players.length) {
        imgEl.setTitle("PHOTO_" + (sIdx + 1));
        imgEl.setDescription(t.prefix + "_PHOTO_" + (sIdx + 1));
        var pPhoto = t.players[sIdx];
        try {
          if (pPhoto && pPhoto.photoUrl) {
            imgEl.asImage().replace(pPhoto.photoUrl);
          } else {
            imgEl.asImage().replace(transparentPngUrl);
          }
        } catch (imgErr) {
          Logger.log("Avatar replace warning for slot " + (sIdx + 1) + ": " + imgErr.message);
        }
      }
    });

    // 4. Update any 2x4 Match Info Tables (T20 slides or fixture summary tables)
    elements.forEach(function(el) {
      if (el.getPageElementType() === SlidesApp.PageElementType.TABLE) {
        var tbl = el.asTable();
        for (var r = 0; r < tbl.getNumRows(); r++) {
          for (var c = 0; c < tbl.getNumColumns(); c++) {
            var label = tbl.getCell(r, c).getText().asString().trim().toLowerCase();
            if (label === "round" && c + 1 < tbl.getNumColumns()) {
              tbl.getCell(r, c + 1).getText().setText(t.round);
            } else if ((label === "playing" || label === "opponent" || label === "versus" || label === "vs") && c + 1 < tbl.getNumColumns()) {
              tbl.getCell(r, c + 1).getText().setText(t.opponent);
            } else if ((label === "ground" || label === "venue") && c + 1 < tbl.getNumColumns()) {
              tbl.getCell(r, c + 1).getText().setText(t.venue);
            } else if (label === "format" && c + 1 < tbl.getNumColumns()) {
              tbl.getCell(r, c + 1).getText().setText(t.format);
            }
          }
        }
      }
    });
  });

  return { success: true, message: "Successfully synced all 5 teams to Google Slides using permanent element IDs!" };
}


/**
 * Retrieves summary info for the Google Slides sync dialog.
 */
function getSlidesSyncSummary() {
  var ss = getSS();
  var presId = getSlidesPresentationId();
  var presUrl = "https://docs.google.com/presentation/d/" + presId + "/edit";
  var staging = ss ? ss.getSheetByName("Presentation_Staging") : null;
  var roundVal = "";
  if (staging) {
    var rawB1 = staging.getRange("B1").getValue();
    if (rawB1 instanceof Date) {
      roundVal = Utilities.formatDate(rawB1, "Australia/Melbourne", "yyyy-MM-dd");
    } else {
      roundVal = String(rawB1 || "").trim();
    }
  }
  return {
    presentationId: presId,
    presentationUrl: presUrl,
    roundToPresent: roundVal || "(None selected in B1)",
    teamCount: 5
  };
}


/**
 * Opens the interactive Google Slides Sync modal dialog.
 */
function showSyncSlidesDialog() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html><head><base target="_top">' +
    '<style>' +
    '  * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }' +
    '  body { padding: 24px; background: #fafafa; color: #333; }' +
    '  .header { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; border-bottom: 2px solid #4d0012; padding-bottom: 12px; }' +
    '  .header h2 { color: #4d0012; font-size: 20px; font-weight: 700; }' +
    '  .card { background: #fff; border: 1px solid #e0e0e0; border-radius: 8px; padding: 16px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }' +
    '  .info-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #f0f0f0; font-size: 14px; }' +
    '  .info-row:last-child { border-bottom: none; }' +
    '  .label { color: #666; font-weight: 500; }' +
    '  .value { color: #111; font-weight: 600; }' +
    '  .actions { display: flex; gap: 12px; justify-content: flex-end; margin-top: 24px; }' +
    '  button { padding: 10px 20px; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; border: none; transition: all 0.2s; }' +
    '  .btn-primary { background: #4d0012; color: #fff; }' +
    '  .btn-primary:hover { background: #35000c; }' +
    '  .btn-secondary { background: #e0e0e0; color: #444; }' +
    '  .btn-secondary:hover { background: #d0d0d0; }' +
    '  .btn-gold { background: #fac218; color: #4d0012; border: 1px solid #dfac13; }' +
    '  .btn-gold:hover { background: #e6b216; }' +
    '  .spinner { display: inline-block; width: 32px; height: 32px; border: 3px solid #f3f3f3; border-top: 3px solid #4d0012; border-radius: 50%; animation: spin 1s linear infinite; margin-bottom: 12px; }' +
    '  @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
    '  .state-box { text-align: center; padding: 24px 0; }' +
    '</style>' +
    '</head><body>' +
    '<div class="header">' +
    '  <h2>🏏 Sync to Google Slides</h2>' +
    '</div>' +
    '<div id="confirmState">' +
    '  <p style="font-size: 14px; color: #555; margin-bottom: 16px;">This will read the active round selection from <strong>Presentation_Staging</strong> and populate all 5 team slides in your Google Slides deck.</p>' +
    '  <div class="card">' +
    '    <div class="info-row"><span class="label">Round to Present:</span><span class="value" id="roundVal">Loading...</span></div>' +
    '    <div class="info-row"><span class="label">Teams:</span><span class="value">1st, 2nd, 3rd, 4th, 5th Elevens</span></div>' +
    '    <div class="info-row"><span class="label">Headshot Sync:</span><span class="value">Drive Photos & Transparent fallback</span></div>' +
    '  </div>' +
    '  <div class="actions">' +
    '    <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '    <button class="btn-primary" id="syncBtn" onclick="startSync()">🚀 Start Sync</button>' +
    '  </div>' +
    '</div>' +
    '<div id="loadingState" style="display: none;" class="state-box">' +
    '  <div class="spinner"></div>' +
    '  <p style="font-size: 15px; font-weight: 600; color: #4d0012;">Updating Google Slides presentation...</p>' +
    '  <p style="font-size: 13px; color: #777; margin-top: 4px;">Populating team rosters, match metadata, and player headshots.</p>' +
    '</div>' +
    '<div id="successState" style="display: none;" class="state-box">' +
    '  <div style="font-size: 40px; margin-bottom: 12px;">✅</div>' +
    '  <h3 style="color: #4d0012; font-size: 18px; margin-bottom: 8px;">Sync Complete!</h3>' +
    '  <p id="successMsg" style="font-size: 14px; color: #555; margin-bottom: 24px;">All 5 team slides are now up to date.</p>' +
    '  <div class="actions" style="justify-content: center;">' +
    '    <button class="btn-gold" id="openSlidesBtn" onclick="openSlides()">🔗 Open Google Slides</button>' +
    '    <button class="btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '  </div>' +
    '</div>' +
    '<script>' +
    '  var presentationUrl = "";' +
    '  google.script.run.withSuccessHandler(function(summary) {' +
    '    document.getElementById("roundVal").innerText = summary.roundToPresent;' +
    '    presentationUrl = summary.presentationUrl;' +
    '  }).getSlidesSyncSummary();' +
    '  function startSync() {' +
    '    document.getElementById("confirmState").style.display = "none";' +
    '    document.getElementById("loadingState").style.display = "block";' +
    '    google.script.run.withSuccessHandler(function(res) {' +
    '      document.getElementById("loadingState").style.display = "none";' +
    '      document.getElementById("successState").style.display = "block";' +
    '      if (res && res.message) document.getElementById("successMsg").innerText = res.message;' +
    '    }).withFailureHandler(function(err) {' +
    '      alert("Sync Error: " + err.message);' +
    '      google.script.host.close();' +
    '    }).syncPresentationStagingToSlides();' +
    '  }' +
    '  function openSlides() {' +
    '    if (presentationUrl) window.open(presentationUrl, "_blank");' +
    '  }' +
    '</script>' +
    '</body></html>'
  ).setWidth(520).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, "🏏 Google Slides Sync");
}


/**
 * Retrieves the full players list for the Player Photo Studio.
 */
function getPlayersListForStudio() {
  var ss = getSS();
  if (!ss) return [];
  var playerSheet = ss.getSheetByName("Players");
  if (!playerSheet || playerSheet.getLastRow() < 2) return [];

  var data = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, 14).getValues();
  var players = [];
  data.forEach(function(row) {
    var profileId = String(row[0] || "").trim();
    var fName = String(row[3] || (row[1] + " " + row[2])).trim();
    var status = String(row[6] || "").trim();
    var photoUrl = String(row[13] || "").trim();
    if (profileId && fName && status !== "Inactive") {
      players.push({
        profileId: profileId,
        fullName: fName,
        hasPhoto: photoUrl !== "",
        photoUrl: photoUrl
      });
    }
  });

  players.sort(function(a, b) {
    return a.fullName.localeCompare(b.fullName);
  });
  return players;
}


/**
 * Saves a player headshot directly to Google Drive as <ProfileID>.png and updates Players sheet.
 */
function savePlayerHeadshot(profileId, base64Data) {
  if (!profileId || !base64Data) {
    throw new Error("Missing player Profile ID or image data.");
  }
  var ss = getSS();
  var folderId = getHeadshotFolderId();
  if (!folderId) throw new Error("HEADSHOT_FOLDER_ID is not configured in Script Properties.");

  var cleanBase64 = base64Data.replace(/^data:image\/\w+;base64,/, "");
  var decoded = Utilities.base64Decode(cleanBase64);
  var blob = Utilities.newBlob(decoded, "image/png", profileId + ".png");

  var folder = DriveApp.getFolderById(folderId);

  // Trash any previous versions of this player's photo
  var existingPng = folder.getFilesByName(profileId + ".png");
  while (existingPng.hasNext()) {
    existingPng.next().setTrashed(true);
  }
  var existingJpg = folder.getFilesByName(profileId + ".jpg");
  while (existingJpg.hasNext()) {
    existingJpg.next().setTrashed(true);
  }

  // Create new transparent PNG
  var newFile = folder.createFile(blob);
  var fileId = newFile.getId();
  var photoUrl = "https://lh3.googleusercontent.com/d/" + fileId;

  // Update Players sheet Col N
  if (ss) {
    var pSheet = ss.getSheetByName("Players");
    if (pSheet && pSheet.getLastRow() > 1) {
      var pIds = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < pIds.length; i++) {
        if (String(pIds[i][0]).trim().toLowerCase() === profileId.toLowerCase()) {
          pSheet.getRange(2 + i, 14).setValue(photoUrl);
          break;
        }
      }
    }
  }

  return {
    success: true,
    profileId: profileId,
    photoUrl: photoUrl,
    message: "Headshot saved successfully!"
  };
}


/**
 * Opens the Player Photo Studio interactive modal dialog with Google MediaPipe AI background removal.
 */
function showPhotoStudioDialog() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html><head><base target="_top">' +
    '<script src="https://cdn.jsdelivr.net/npm/@mediapipe/camera_utils/camera_utils.js" crossorigin="anonymous"></script>' +
    '<script src="https://cdn.jsdelivr.net/npm/@mediapipe/selfie_segmentation/selfie_segmentation.js" crossorigin="anonymous"></script>' +
    '<style>' +
    '  * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }' +
    '  body { padding: 14px 18px; background: #fafafa; color: #222; overflow-y: hidden; user-select: none; }' +
    '  .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }' +
    '  .card { background: #fff; border: 1px solid #e0e0e0; border-radius: 8px; padding: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }' +
    '  .card h3 { font-size: 13px; font-weight: 700; color: #4d0012; margin-bottom: 8px; }' +
    '  label { font-size: 11px; font-weight: 600; color: #555; display: block; margin-bottom: 3px; }' +
    '  select, input[type="text"] { width: 100%; padding: 7px 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; margin-bottom: 8px; }' +
    '  .drop-zone { border: 2px dashed #4d0012; border-radius: 8px; padding: 16px; text-align: center; background: #fff9e6; cursor: pointer; transition: all 0.2s; }' +
    '  .drop-zone:hover { background: #fff2cc; }' +
    '  .canvas-container { position: relative; width: 100%; height: 210px; background: repeating-conic-gradient(#eee 0% 25%, #fff 0% 50%) 50% / 12px 12px; border-radius: 6px; overflow: hidden; display: flex; align-items: center; justify-content: center; border: 1px solid #ddd; }' +
    '  canvas { max-width: 100%; max-height: 100%; cursor: grab; }' +
    '  .preview-circle { width: 80px; height: 80px; border-radius: 50%; border: 3px solid #fac218; box-shadow: 0 2px 6px rgba(0,0,0,0.15); overflow: hidden; margin: 0 auto 4px; background: repeating-conic-gradient(#eee 0% 25%, #fff 0% 50%) 50% / 12px 12px; }' +
    '  .preview-circle img { width: 100%; height: 100%; object-fit: cover; }' +
    '  .controls { display: flex; gap: 8px; align-items: center; margin-top: 6px; }' +
    '  .controls button { padding: 5px 10px; font-size: 11px; font-weight: 600; border-radius: 5px; cursor: pointer; }' +
    '  .actions { display: flex; justify-content: space-between; align-items: center; gap: 10px; margin-top: 10px; padding-top: 8px; border-top: 1px solid #eee; }' +
    '  button { padding: 8px 16px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; border: none; }' +
    '  .btn-primary { background: #4d0012; color: #fff; }' +
    '  .btn-primary:hover { background: #35000c; }' +
    '  .btn-primary:disabled { background: #ccc; cursor: not-allowed; }' +
    '  .btn-secondary { background: #e0e0e0; color: #333; }' +
    '  .btn-gold { background: #fac218; color: #4d0012; font-weight: bold; border: 1px solid #dfac13; }' +
    '  .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; }' +
    '  .badge-has { background: #e6f4ea; color: #137333; }' +
    '  .badge-none { background: #fce8e6; color: #c5221f; }' +
    '  .toast { padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; }' +
    '  .toast-success { background: #e6f4ea; color: #137333; border: 1px solid #ceead6; }' +
    '  .toast-error { background: #fce8e6; color: #c5221f; border: 1px solid #fad2cf; }' +
    '</style>' +
    '</head><body>' +
    '<div class="grid">' +
    '  <div class="card">' +
    '    <h3>1. Select Player</h3>' +
    '    <label for="playerSelect">Player:</label>' +
    '    <select id="playerSelect" onchange="onPlayerChange()"><option value="">Loading players...</option></select>' +
    '    <div id="playerStatus" style="font-size: 11px; margin-bottom: 8px;"></div>' +
    '    <h3>2. Choose Photo</h3>' +
    '    <div class="drop-zone" onclick="document.getElementById(\'fileInput\').click()">' +
    '      <div style="font-size: 22px; margin-bottom: 2px;">📷</div>' +
    '      <p style="font-size: 12px; font-weight: 600; color: #4d0012;">Click to Upload or Snap</p>' +
    '      <p style="font-size: 10px; color: #777;">PNG or JPG</p>' +
    '    </div>' +
    '    <input type="file" id="fileInput" accept="image/*" style="display:none;" onchange="handleFile(this.files[0])">' +
    '  </div>' +
    '  <div class="card">' +
    '    <h3>3. Position & AI Cutout</h3>' +
    '    <div class="canvas-container">' +
    '      <canvas id="cropCanvas" width="400" height="400"></canvas>' +
    '    </div>' +
    '    <div style="font-size: 10px; color: #777; margin-top: 4px; text-align: center;">🖐️ Drag photo to center face • Scroll/Slider to zoom</div>' +
    '    <div class="controls">' +
    '      <label style="margin:0; font-size:11px;">Zoom:</label>' +
    '      <input type="range" id="zoomSlider" min="0.5" max="3.5" step="0.02" value="1" oninput="drawCanvas()" style="flex:1;">' +
    '      <button class="btn-gold" id="bgBtn" onclick="removeBackground()" title="AI Human Portrait Background Removal">🪄 AI Cutout BG</button>' +
    '    </div>' +
    '    <div style="text-align: center; margin-top: 4px;">' +
    '      <div class="preview-circle">' +
    '        <img id="previewImg" src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7">' +
    '      </div>' +
    '      <span style="font-size: 10px; color: #666;">Circle Avatar Output (400×400 Transparent PNG)</span>' +
    '    </div>' +
    '  </div>' +
    '</div>' +
    '<div class="actions">' +
    '  <div id="inlineFeedback"></div>' +
    '  <div style="display:flex; gap:10px;">' +
    '    <button class="btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '    <button class="btn-primary" id="saveBtn" onclick="savePhoto()" disabled>💾 Save to Player Profile</button>' +
    '  </div>' +
    '</div>' +
    '<script>' +
    '  var playersData = [];' +
    '  var rawImg = new Image();' +
    '  var imgLoaded = false;' +
    '  var baseScale = 1;' +
    '  var zoomMultiplier = 1;' +
    '  var panX = 0, panY = 0;' +
    '  var isDragging = false;' +
    '  var startMouseX = 0, startMouseY = 0;' +
    '  var startPanX = 0, startPanY = 0;' +
    '  var canvas = document.getElementById("cropCanvas");' +
    '  var ctx = canvas.getContext("2d");' +
    '  var selfieSegmentation = null;' +
    '  try {' +
    '    selfieSegmentation = new SelfieSegmentation({ locateFile: function(f) { return "https://cdn.jsdelivr.net/npm/@mediapipe/selfie_segmentation/" + f; } });' +
    '    selfieSegmentation.setOptions({ modelSelection: 1 });' +
    '    selfieSegmentation.onResults(onSegmentationResults);' +
    '  } catch (e) { console.warn("MediaPipe init error:", e); }' +
    '  google.script.run.withSuccessHandler(function(list) {' +
    '    playersData = list;' +
    '    var sel = document.getElementById("playerSelect");' +
    '    sel.innerHTML = "<option value=\'\'>-- Select a Player --</option>";' +
    '    list.forEach(function(p) {' +
    '      var opt = document.createElement("option");' +
    '      opt.value = p.profileId;' +
    '      opt.innerText = p.fullName + (p.hasPhoto ? " (Has Photo)" : " (No Photo)");' +
    '      sel.appendChild(opt);' +
    '    });' +
    '  }).getPlayersListForStudio();' +
    '  function onPlayerChange() {' +
    '    var pid = document.getElementById("playerSelect").value;' +
    '    var p = playersData.find(function(item) { return item.profileId === pid; });' +
    '    var statusDiv = document.getElementById("playerStatus");' +
    '    if (!p) { statusDiv.innerHTML = ""; checkReady(); return; }' +
    '    if (p.hasPhoto) {' +
    '      statusDiv.innerHTML = "<span class=\'badge badge-has\'>Photo Active</span> <a href=\'" + p.photoUrl + "\' target=\'_blank\' style=\'color:#4d0012; font-size:11px;\'>View current</a>";' +
    '    } else {' +
    '      statusDiv.innerHTML = "<span class=\'badge badge-none\'>No Photo Uploaded</span>";' +
    '    }' +
    '    checkReady();' +
    '  }' +
    '  function handleFile(file) {' +
    '    if (!file) return;' +
    '    var reader = new FileReader();' +
    '    reader.onload = function(e) {' +
    '      rawImg.onload = function() {' +
    '        imgLoaded = true;' +
    '        baseScale = 400 / Math.max(rawImg.width, rawImg.height);' +
    '        zoomMultiplier = 1;' +
    '        panX = 0; panY = 0;' +
    '        document.getElementById("zoomSlider").value = 1;' +
    '        drawCanvas();' +
    '        checkReady();' +
    '      };' +
    '      rawImg.src = e.target.result;' +
    '    };' +
    '    reader.readAsDataURL(file);' +
    '  }' +
    '  function drawCanvas() {' +
    '    if (!imgLoaded) return;' +
    '    zoomMultiplier = parseFloat(document.getElementById("zoomSlider").value);' +
    '    var curScale = baseScale * zoomMultiplier;' +
    '    var w = rawImg.width * curScale;' +
    '    var h = rawImg.height * curScale;' +
    '    var cx = 200 + panX;' +
    '    var cy = 200 + panY;' +
    '    ctx.clearRect(0, 0, canvas.width, canvas.height);' +
    '    ctx.drawImage(rawImg, cx - w / 2, cy - h / 2, w, h);' +
    '    updatePreview();' +
    '  }' +
    '  canvas.addEventListener("mousedown", function(e) {' +
    '    if (!imgLoaded) return;' +
    '    isDragging = true;' +
    '    startMouseX = e.clientX; startMouseY = e.clientY;' +
    '    startPanX = panX; startPanY = panY;' +
    '    canvas.style.cursor = "grabbing";' +
    '  });' +
    '  window.addEventListener("mousemove", function(e) {' +
    '    if (!isDragging) return;' +
    '    panX = startPanX + (e.clientX - startMouseX);' +
    '    panY = startPanY + (e.clientY - startMouseY);' +
    '    drawCanvas();' +
    '  });' +
    '  window.addEventListener("mouseup", function() {' +
    '    if (isDragging) { isDragging = false; canvas.style.cursor = "grab"; }' +
    '  });' +
    '  canvas.addEventListener("touchstart", function(e) {' +
    '    if (!imgLoaded || e.touches.length !== 1) return;' +
    '    isDragging = true;' +
    '    startMouseX = e.touches[0].clientX; startMouseY = e.touches[0].clientY;' +
    '    startPanX = panX; startPanY = panY;' +
    '  }, { passive: true });' +
    '  window.addEventListener("touchmove", function(e) {' +
    '    if (!isDragging || e.touches.length !== 1) return;' +
    '    panX = startPanX + (e.touches[0].clientX - startMouseX);' +
    '    panY = startPanY + (e.touches[0].clientY - startMouseY);' +
    '    drawCanvas();' +
    '  }, { passive: true });' +
    '  window.addEventListener("touchend", function() { isDragging = false; });' +
    '  canvas.addEventListener("wheel", function(e) {' +
    '    if (!imgLoaded) return;' +
    '    e.preventDefault();' +
    '    var slider = document.getElementById("zoomSlider");' +
    '    var val = parseFloat(slider.value) + (e.deltaY < 0 ? 0.08 : -0.08);' +
    '    slider.value = Math.max(0.5, Math.min(3.5, val));' +
    '    drawCanvas();' +
    '  }, { passive: false });' +
    '  function updatePreview() {' +
    '    document.getElementById("previewImg").src = canvas.toDataURL("image/png");' +
    '  }' +
    '  function removeBackground() {' +
    '    if (!imgLoaded) return;' +
    '    var btn = document.getElementById("bgBtn");' +
    '    btn.innerText = "⏳ AI Working...";' +
    '    btn.disabled = true;' +
    '    var tempCanvas = document.createElement("canvas");' +
    '    tempCanvas.width = 400; tempCanvas.height = 400;' +
    '    var tempCtx = tempCanvas.getContext("2d");' +
    '    var curScale = baseScale * zoomMultiplier;' +
    '    var w = rawImg.width * curScale;' +
    '    var h = rawImg.height * curScale;' +
    '    var cx = 200 + panX;' +
    '    var cy = 200 + panY;' +
    '    tempCtx.drawImage(rawImg, cx - w / 2, cy - h / 2, w, h);' +
    '    if (selfieSegmentation) {' +
    '      selfieSegmentation.send({ image: tempCanvas }).catch(function(err) {' +
    '        console.error(err);' +
    '        fallbackCutout();' +
    '      });' +
    '    } else {' +
    '      fallbackCutout();' +
    '    }' +
    '  }' +
    '  function onSegmentationResults(results) {' +
    '    var w = canvas.width, h = canvas.height;' +
    '    ctx.save();' +
    '    ctx.clearRect(0, 0, w, h);' +
    '    ctx.drawImage(results.segmentationMask, 0, 0, w, h);' +
    '    ctx.globalCompositeOperation = "source-in";' +
    '    var curScale = baseScale * zoomMultiplier;' +
    '    var sw = rawImg.width * curScale;' +
    '    var sh = rawImg.height * curScale;' +
    '    var cx = 200 + panX;' +
    '    var cy = 200 + panY;' +
    '    ctx.drawImage(rawImg, cx - sw / 2, cy - sh / 2, sw, sh);' +
    '    ctx.restore();' +
    '    updatePreview();' +
    '    var btn = document.getElementById("bgBtn");' +
    '    btn.innerText = "✨ AI Cutout Done!";' +
    '    btn.disabled = false;' +
    '  }' +
    '  function fallbackCutout() {' +
    '    var imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);' +
    '    var data = imgData.data;' +
    '    var cornerR = data[0], cornerG = data[1], cornerB = data[2];' +
    '    for (var i = 0; i < data.length; i += 4) {' +
    '      var r = data[i], g = data[i+1], b = data[i+2];' +
    '      var diff = Math.abs(r - cornerR) + Math.abs(g - cornerG) + Math.abs(b - cornerB);' +
    '      if (diff < 75) data[i+3] = 0;' +
    '    }' +
    '    ctx.putImageData(imgData, 0, 0);' +
    '    updatePreview();' +
    '    var btn = document.getElementById("bgBtn");' +
    '    btn.innerText = "🪄 AI Cutout BG";' +
    '    btn.disabled = false;' +
    '  }' +
    '  function checkReady() {' +
    '    var pid = document.getElementById("playerSelect").value;' +
    '    document.getElementById("saveBtn").disabled = !(pid && imgLoaded);' +
    '  }' +
    '  function savePhoto() {' +
    '    var pid = document.getElementById("playerSelect").value;' +
    '    if (!pid || !imgLoaded) return;' +
    '    var btn = document.getElementById("saveBtn");' +
    '    var fb = document.getElementById("inlineFeedback");' +
    '    btn.disabled = true;' +
    '    btn.innerText = "⏳ Saving...";' +
    '    fb.innerHTML = "";' +
    '    var base64 = canvas.toDataURL("image/png");' +
    '    google.script.run.withSuccessHandler(function(res) {' +
    '      btn.innerText = "✅ Saved!";' +
    '      fb.innerHTML = "<span class=\'toast toast-success\'>✅ Saved to Google Drive!</span>";' +
    '      setTimeout(function() { google.script.host.close(); }, 1200);' +
    '    }).withFailureHandler(function(err) {' +
    '      btn.disabled = false;' +
    '      btn.innerText = "💾 Save to Player Profile";' +
    '      fb.innerHTML = "<span class=\'toast toast-error\'>❌ " + err.message + "</span>";' +
    '    }).savePlayerHeadshot(pid, base64);' +
    '  }' +
    '</script>' +
    '</body></html>'
  ).setWidth(740).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "📸 Player Photo Studio");
}