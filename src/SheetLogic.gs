

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
      for (var fRow = 1; fRow < fValues.length; fRow++) {
        if (String(fValues[fRow][dateColIdx]).trim() === dateStr) {
          fixInfo = {};
          fHeaders.forEach(function(h, i) {
            fixInfo[String(h).trim().toLowerCase()] = fValues[fRow][i];
          });
          break;
        }
      }
    }
  }

  // Header Metadata
  var initialRoundId = (fixInfo && (fixInfo["1st round"] || fixInfo["roundid"] || fixInfo["round"])) ? (fixInfo["1st round"] || fixInfo["roundid"] || fixInfo["round"]) : "";
  ws.getRange("A1").setValue("Round ID:").setFontWeight("bold");
  ws.getRange("B1").setValue(initialRoundId).setBackground(LCC_PALETTE.inputHighlight).setHorizontalAlignment("center").setFontWeight("bold");
  ws.getRange("G1").setValue("Roster Last Synced:").setFontWeight("bold");
  ws.getRange("H1").setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm").setFontColor("#666666");
  ws.getRange("A2").setValue("Date").setFontWeight("bold");
  ws.getRange("B2").setValue(targetDate).setNumberFormat("dd/mm/yyyy").setBackground(LCC_PALETTE.inputHighlight).setHorizontalAlignment("center").setFontWeight("bold");

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

    var headerTitle = gName + (rnd ? " (" + rnd + ")" : "");
    ws.getRange(currentTeamRow, 1, 1, 2).merge().setValue(headerTitle).setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
    ws.getRange(currentTeamRow + 1, 1).setValue("Opponent:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 1, 2).setValue(opp);
    ws.getRange(currentTeamRow + 2, 1).setValue("Venue:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 2, 2).setValue(ven);
    ws.getRange(currentTeamRow + 3, 1).setValue("Format:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 3, 2).setValue(fmt);

    var structure = [
      ["1. Captain", ""], ["2. VC", ""], ["3. WK", ""],
      ["4. Player", ""], ["5. Player", ""], ["6. Player", ""],
      ["7. Player", ""], ["8. Player", ""], ["9. Player", ""],
      ["10. Player", ""], ["11. Player", ""], ["12. Player", ""], ["13. Player", ""]
    ];
    ws.getRange(currentTeamRow + 4, 1, 13, 2).setValues(structure);
    ws.getRange(currentTeamRow + 4, 1, 13, 1).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);
    ws.getRange(currentTeamRow, 1, 17, 2).setBorder(true, true, true, true, true, true, LCC_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);

    currentTeamRow += 18;
  });

  // Col D: Dynamic Virtual "Available for Selection" Pool
  ws.getRange("D3").setValue("AVAILABLE FOR SELECTION").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
  ws.getRange("D4").setValue("(Dynamic Unselected Pool)").setFontStyle("italic").setFontColor("#666666");
  ws.getRange("D5").setFormula(`=IFERROR(FILTER(G5:G, G5:G<>"", COUNTIF(B:B, G5:G)=0), "")`);
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
  ws.getRange("F4:I4").setValues([["Profile ID", "Player Name", "Notes Context", "Action Switcher"]]).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);

  // List 3: Unavailable for Round (Cols K-N)
  ws.getRange("K3:N3").merge().setValue("UNAVAILABLE FOR ROUND").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
  ws.getRange("K4:N4").setValues([["Profile ID", "Player Name", "Exemption Notes", "Action Switcher"]]).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);
  if (unavailableSnapshot.length > 0) {
    ws.getRange(5, 11, unavailableSnapshot.length, 4).setValues(unavailableSnapshot);
  }

  // List 4: Unknown Status Pool Ledger (Cols P-S)
  ws.getRange("P3:S3").merge().setValue("UNKNOWN AVAILABILITY STATUS").setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
  ws.getRange("P4:S4").setValues([["Profile ID", "Player Name", "Notes Context", "Action Switcher"]]).setFontWeight("bold").setBackground(LCC_PALETTE.zebraLight);
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

  var staging = ss.getSheetByName("Presentation_Staging");
  if (staging) staging.getRange("B1").setValue(sheetName);

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
    .addItem("Record player injury / absence", "showRecordInjuryDialog")
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
    // Col D: Global unselected pool
    ws.getRange("D5").setFormula(`=IFERROR(FILTER(G5:G, G5:G<>"", COUNTIF(B:B, G5:G)=0), "")`);

    var teamStarts = [8, 26, 44, 62, 80];
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

        // Set helper header and dynamic formula
        ws.getRange(3, helperCol).setValue("SLOT_" + row);
        ws.getRange(5, helperCol).setFormula(
          '=IFERROR(FILTER(G$5:G, (COUNTIF($B$8:$B$92, G$5:G)=0) + (G$5:G=B' + row + ')), "")'
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
    '"Please submit your availability for the upcoming round (" & TEXT(B2, "yyyy-mm-dd") & "):" & CHAR(10) & ' +
    '"https://lcc-availability.web.app/?round=" & TEXT(B2, "yyyy-mm-dd")'
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
    '=LCC_WALL_OF_SHAME(B2, COUNTA(G5:G) + COUNTA(L5:L), Q5:Q150)'
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
 * Action switchers 108px (-5px), Notes 148px, Names 153px, Notes font size 9pt.
 */
function applyRoundTabColumnWidths(ws) {
  if (!ws) return;
  ws.setColumnWidth(1, 110); ws.setColumnWidth(2, 190); ws.setColumnWidth(3, 20);
  ws.setColumnWidth(4, 190); ws.setColumnWidth(5, 20);
  ws.setColumnWidth(6, 80);  ws.setColumnWidth(7, 153); ws.setColumnWidth(8, 148); ws.setColumnWidth(9, 108); ws.setColumnWidth(10, 20);
  ws.setColumnWidth(11, 80); ws.setColumnWidth(12, 153); ws.setColumnWidth(13, 148); ws.setColumnWidth(14, 108); ws.setColumnWidth(15, 20);
  ws.setColumnWidth(16, 80); ws.setColumnWidth(17, 153); ws.setColumnWidth(18, 148); ws.setColumnWidth(19, 108); ws.setColumnWidth(20, 20);
  ws.setColumnWidth(21, 330); // Col U (Availability Message)
  ws.setColumnWidth(22, 330); // Col V (Wall of Shame Message)
  ws.setColumnWidth(23, 20);  // Col W (Spacer)

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
 * Menu action to re-apply standard column widths across all tabs.
 */
function menuResetColumnWidths() {
  var ss = getSS();
  if (!ss) return;
  applyAllStandardColumnWidths(ss);
  ss.toast("Standard column widths applied across all database tabs.", "LCC Selection Engine", 4);
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
function showSheetUrl() {
  var ss = getSS();
  if (!ss) {
    Logger.log("ERROR: No spreadsheet found. Run setupInitialSystem first.");
    return;
  }
  Logger.log("Sheet URL: " + ss.getUrl());
  return ss.getUrl();
}


/**
 * Diagnostic: tests photo lookup for a specific profile ID.
 * Run from the editor to verify a specific player's photo.
 */
function testPhotoLookup() {
  var profileId = "ENTER_PROFILE_ID_HERE";  // Replace with a real profile ID
  var url = getPlayerPhotoUrl(profileId);
  if (url) {
    Logger.log("Photo URL for " + profileId + ": " + url);
  } else {
    Logger.log("No photo found for " + profileId);
  }
}