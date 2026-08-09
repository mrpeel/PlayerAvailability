

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
  var sheetName = dateStr;
  if (ss.getSheetByName(sheetName)) {
    SpreadsheetApp.getUi().alert("Tab '" + sheetName + "' already exists.");
    return;
  }

  var ws = ss.insertSheet(sheetName);
  var targetDate = new Date(dateStr);

  // Header Metadata
  ws.getRange("A1").setValue("Round ID:").setFontWeight("bold");
  ws.getRange("B1").setValue("").setBackground(LCC_PALETTE.inputHighlight).setHorizontalAlignment("center").setFontWeight("bold");
  ws.getRange("D1").setValue("Roster Last Synced:").setFontWeight("bold");
  ws.getRange("E1").setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm").setFontColor("#666666");
  ws.getRange("A2").setValue("Date").setFontWeight("bold");
  ws.getRange("B2").setValue(targetDate).setNumberFormat("dd/mm/yyyy").setBackground(LCC_PALETTE.inputHighlight).setHorizontalAlignment("center").setFontWeight("bold");

  // Vertical Team Stack (Cols A & B)
  var grades = ["FIRST ELEVEN", "SECOND ELEVEN", "THIRD ELEVEN", "FOURTH ELEVEN", "FIFTH ELEVEN"];
  var currentTeamRow = 4;

  grades.forEach(function(gName) {
    ws.getRange(currentTeamRow, 1, 1, 2).merge().setValue(gName).setFontWeight("bold").setBackground(LCC_PALETTE.maroonBg).setFontColor(LCC_PALETTE.maroonFg).setHorizontalAlignment("center");
    ws.getRange(currentTeamRow + 1, 1).setValue("Opponent:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 2, 1).setValue("Venue:").setFontStyle("italic");
    ws.getRange(currentTeamRow + 3, 1).setValue("Format:").setFontStyle("italic");

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
  ws.getRange("D5").setFormula(`=IFERROR(FILTER(G5:G, G5:G<>"", COUNTIF(B:B, G5:G)=0), "All players selected")`);
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

  // Col B Selection Dropdown linked to Col G (Available for Round Names)
  var teamRows = [8, 26, 44, 62, 80];
  teamRows.forEach(function(start) {
    ws.getRange(start, 2, 13, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(ws.getRange("G5:G150"), true).build()
    );
  });

  // Column Widths & Hide Technical ProfileID Columns
  ws.setColumnWidth(1, 130); ws.setColumnWidth(2, 210); ws.setColumnWidth(3, 20);
  ws.setColumnWidth(4, 210); ws.setColumnWidth(5, 20);
  ws.setColumnWidth(6, 80);  ws.setColumnWidth(7, 210); ws.setColumnWidth(8, 210); ws.setColumnWidth(9, 180); ws.setColumnWidth(10, 20);
  ws.setColumnWidth(11, 80); ws.setColumnWidth(12, 210); ws.setColumnWidth(13, 210); ws.setColumnWidth(14, 180); ws.setColumnWidth(15, 20);
  ws.setColumnWidth(16, 80); ws.setColumnWidth(17, 210); ws.setColumnWidth(18, 210); ws.setColumnWidth(19, 180);

  ws.hideColumns(6);  // Col F (ProfileID)
  ws.hideColumns(11); // Col K (ProfileID)
  ws.hideColumns(16); // Col P (ProfileID)

  var staging = ss.getSheetByName("Presentation_Staging");
  if (staging) staging.getRange("B1").setValue(sheetName);
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
  var responseOutput = { status: "success", message: "Roster data state synchronized." };
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || null;

    if (action === 'getInitialData' || action === 'getAdminRounds' || action === 'getAdminData') {
      var result = getInitialData(payload.phone, payload.dateStr || payload.roundNum);
      return jsonResponse(result);
    }

    // Default path — save availability / list movement
    var profileId = String(payload.profileId || payload.playerId).trim();
    var targetDateStr = String(payload.date || payload.roundNum).trim();
    var playerResponse = String(payload.response).trim();
    var playerNotes = payload.notes || "";

    var ss = getSS();
    var ws = ss.getSheetByName(targetDateStr);

    if (!ws) {
      return jsonResponse({ status: "error", message: "Target date ledger tab not found." });
    }

    executeRealtimeAppListMovement(ws, profileId, playerResponse, playerNotes);
  } catch (err) {
    responseOutput = { status: "error", message: err.toString() };
  }
  return jsonResponse(responseOutput);
}


/**
 * Searches hidden ProfileID columns (Cols F, K, P) to route players.
 * Called by both the web app (doPost) and the injury/inactive dialogs.
 */
function executeRealtimeAppListMovement(ws, profileId, response, notes) {
  var colMap = [6, 11, 16];
  var foundRow = -1;
  var foundCol = -1;

  for (var cIdx = 0; cIdx < colMap.length; cIdx++) {
    var checkCol = colMap[cIdx];
    var idList = ws.getRange(5, checkCol, 100, 1).getValues();
    for (var rIdx = 0; rIdx < idList.length; rIdx++) {
      if (String(idList[rIdx][0]).trim() === profileId) {
        foundRow = 5 + rIdx;
        foundCol = checkCol;
        break;
      }
    }
    if (foundRow !== -1) break;
  }

  var targetDestCol = (response === "Available") ? 6 : 11;

  if (foundRow !== -1) {
    movePlayerRowsBetweenLists(ws, foundRow, foundCol, targetDestCol);
    var updatedIds = ws.getRange(5, targetDestCol, 100, 1).getValues();
    for (var k = 0; k < updatedIds.length; k++) {
      if (String(updatedIds[k][0]).trim() === profileId) {
        ws.getRange(5 + k, targetDestCol + 2).setValue(notes);
        break;
      }
    }
  }
}


function executeTabDeployment(dateStringYYYYMMDD) {
  var ss = getSS();
  deployVerticalRoundSheet(ss, dateStringYYYYMMDD);
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
    const juniorLevelIdx = pHeaders.indexOf('JuniorClass');

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

    let targetFixture = fRows.find(row => String(row[fHeaders.indexOf('Game Date')]).indexOf(dateStr) > -1);
    if (!targetFixture && fRows.length > 0) {
      targetFixture = fRows[0];
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

    return {
      players: household,
      fixtureInfo: fixtureInfo,
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

  return "Imported " + newRows.length + " new player(s). " + skipped + " skipped.\n\nSheet: " + ss.getUrl();
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
    .addItem("Initialise new round tab", "showDatePickerDialog")
    .addSeparator()
    .addItem("Import players from PlayHQ export", "showImportPlayHQDialog")
    .addSeparator()
    .addItem("Pull new players from Master", "pullNewPlayersToActiveTab")
    .addItem("Record player injury / absence", "showRecordInjuryDialog")
    .addItem("Mark player as inactive", "showMarkInactiveDialog")
    .addToUi();
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
    const ws = ss.getSheetByName(dateStr);

    if (!ws) {
      throw new Error(`Round selection tab for date ${dateStr} has not been initialised yet.`);
    }

    executeRealtimeAppListMovement(ws, String(profileId).trim(), response, notes);

    const logSheet = ss.getSheetByName('Availability_Log');
    if (logSheet) {
      logSheet.appendRow([new Date(), profileId, dateStr, response, notes]);
    }

    return { success: true };
  } catch (e) {
    throw e;
  }
}


/**
 * DIALOG 1: HTML5 Datepicker Modal Window.
 */
function showDatePickerDialog() {
  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><body style="font-family: Arial, sans-serif; padding: 15px; color: #333;">' +
    '<h3 style="color:#6A1B29; margin-top:0;">Initialise Selection Ledger</h3>' +
    '<p style="font-size:13px;">Select the first day of play for this fixture track:</p>' +
    '<input type="date" id="matchDate" style="width:100%; padding:8px; margin-bottom:15px; border:1px solid #ccc; border-radius:4px;">' +
    '<br>' +
    '<div style="text-align: right;">' +
    '  <button onclick="google.script.host.close()" style="background:#ccc; border:none; padding:8px 12px; border-radius:4px; margin-right:5px; cursor:pointer;">Cancel</button>' +
    '  <button onclick="submitDate()" style="background:#6A1B29; color:white; border:none; padding:8px 12px; border-radius:4px; cursor:pointer; font-weight:bold;">Create Tab</button>' +
    '</div>' +
    '<script>' +
    '  function submitDate() {' +
    '    var d = document.getElementById("matchDate").value;' +
    '    if(!d) { alert("Please pick a valid calendar date."); return; }' +
    '    google.script.run.withSuccessHandler(function() { google.script.host.close(); }).executeTabDeployment(d);' +
    '  }' +
    '</script>' +
    '</body></html>'
  ).setWidth(350).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "LCC Selection Engine");
}


/**
 * DIALOG: PlayHQ CSV import — file-picker modal.
 */
function showImportPlayHQDialog() {
  var htmlOutput = HtmlService.createHtmlOutput(
    '<html><body style="font-family: Arial, sans-serif; padding: 15px; color: #333;">' +
    '<h3 style="color:#6A1B29; margin-top:0;">Import PlayHQ Export</h3>' +
    '<div id="inputSection">' +
    '<p style="font-size:13px;">Select the CSV file exported from PlayHQ:</p>' +
    '<input type="file" id="csvFile" accept=".csv" style="width:100%; margin-bottom:15px;">' +
    '<br>' +
    '<div style="text-align: right;">' +
    '  <button onclick="google.script.host.close()" style="background:#ccc; border:none; padding:8px 12px; border-radius:4px; margin-right:5px; cursor:pointer;">Cancel</button>' +
    '  <button onclick="handleImport()" style="background:#6A1B29; color:white; border:none; padding:8px 12px; border-radius:4px; cursor:pointer; font-weight:bold;">Import</button>' +
    '</div>' +
    '</div>' +
    '<div id="spinnerSection" style="display:none; text-align:center; padding:20px;">' +
    '<div style="border:4px solid #f3f3f3; border-top:4px solid #6A1B29; border-radius:50%; width:40px; height:40px; animation:spin 1s linear infinite; margin:0 auto;"></div>' +
    '<p style="margin-top:15px; font-weight:bold;">Importing players...</p>' +
    '</div>' +
    '<div id="resultSection" style="display:none; padding:15px; background:#f0f0f0; border-radius:4px; margin-top:15px;">' +
    '<p id="resultMessage" style="margin:0; white-space:pre-wrap;"></p>' +
    '<div style="text-align: right; margin-top:15px;">' +
    '  <button onclick="google.script.host.close()" style="background:#6A1B29; color:white; border:none; padding:8px 12px; border-radius:4px; cursor:pointer; font-weight:bold;">Close</button>' +
    '</div>' +
    '</div>' +
    '<style>@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }</style>' +
    '<script>' +
    '  function handleImport() {' +
    '    var file = document.getElementById("csvFile").files[0];' +
    '    if (!file) { alert("Please select a CSV file."); return; }' +
    '    document.getElementById("inputSection").style.display = "none";' +
    '    document.getElementById("spinnerSection").style.display = "block";' +
    '    var reader = new FileReader();' +
    '    reader.onload = function(e) {' +
    '      google.script.run' +
    '        .withSuccessHandler(function(msg) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = msg;' +
    '        })' +
    '        .withFailureHandler(function(err) {' +
    '          document.getElementById("spinnerSection").style.display = "none";' +
    '          document.getElementById("resultSection").style.display = "block";' +
    '          document.getElementById("resultMessage").innerText = "Import failed: " + err.message;' +
    '        })' +
    '        .importPlayHQPlayers(e.target.result);' +
    '    };' +
    '    reader.readAsText(file);' +
    '  }' +
    '</script>' +
    '</body></html>'
  ).setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "PlayHQ Import");
}


/**
 * DIALOG 4: Mark player as inactive.
 */
function showMarkInactiveDialog() {
  var ss = getSS();
  var ws = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var pResp = ui.prompt("Mark Player as Inactive", "Enter Player Name or Profile ID to mark as Inactive:", ui.ButtonSet.OK_CANCEL);
  if (pResp.getSelectedButton() !== ui.Button.OK) return;
  var inputQuery = pResp.getResponseText().trim();

  var playerSheet = ss.getSheetByName("Players");
  var pData = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, 4).getValues();
  var targetProfileId = "";

  for (var i = 0; i < pData.length; i++) {
    var profileId = String(pData[i][0]).trim();
    var fName = pData[i][3] || (pData[i][1] + " " + pData[i][2]);
    if (profileId.toLowerCase() === inputQuery.toLowerCase() || fName.toLowerCase().indexOf(inputQuery.toLowerCase()) > -1) {
      targetProfileId = profileId;
      playerSheet.getRange(2 + i, 7).setValue("Inactive");
      break;
    }
  }

  if (targetProfileId === "") {
    ui.alert("Player not found in Master Players directory.");
    return;
  }

  if (ws.getName() !== "Presentation_Staging" && ws.getName() !== "Players" && ws.getName() !== "Fixtures") {
    executeRealtimeAppListMovement(ws, targetProfileId, "Unavailable", "Marked Inactive");
  }

  ui.alert("Player marked Inactive in Master directory and moved to Unavailable.");
}


/**
 * DIALOG 3: Record injury/absence during selection.
 */
function showRecordInjuryDialog() {
  var ss = getSS();
  var ws = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (ws.getName() === "Presentation_Staging" || ws.getName() === "Players" || ws.getName() === "Fixtures") {
    ui.alert("Please open an active round selection tab first.");
    return;
  }

  var pResp = ui.prompt("Record Injury / Long-Term Absence", "Enter Player Name or Profile ID:", ui.ButtonSet.OK_CANCEL);
  if (pResp.getSelectedButton() !== ui.Button.OK) return;
  var inputQuery = pResp.getResponseText().trim();

  var noteResp = ui.prompt("Injury / Absence Notes", "Enter details (e.g. 'Torn calf - 4 weeks'):", ui.ButtonSet.OK_CANCEL);
  if (noteResp.getSelectedButton() !== ui.Button.OK) return;
  var injuryNotes = noteResp.getResponseText().trim();

  var dateResp = ui.prompt("Expected Return Date (Optional)", "Enter return date (YYYY-MM-DD) or leave blank if indefinite:", ui.ButtonSet.OK_CANCEL);
  var returnDateStr = (dateResp.getSelectedButton() === ui.Button.OK) ? dateResp.getResponseText().trim() : "";

  var playerSheet = ss.getSheetByName("Players");
  var pData = playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, 6).getValues();
  var targetProfileId = "";

  for (var i = 0; i < pData.length; i++) {
    var profileId = String(pData[i][0]).trim();
    var fName = pData[i][3] || (pData[i][1] + " " + pData[i][2]);
    if (profileId.toLowerCase() === inputQuery.toLowerCase() || fName.toLowerCase().indexOf(inputQuery.toLowerCase()) > -1) {
      targetProfileId = profileId;
      var rowNum = 2 + i;
      playerSheet.getRange(rowNum, 7).setValue("Injured");
      if (returnDateStr !== "") playerSheet.getRange(rowNum, 8).setValue(returnDateStr);
      break;
    }
  }

  if (targetProfileId === "") {
    ui.alert("Player not found in Master Players directory.");
    return;
  }

  executeRealtimeAppListMovement(ws, targetProfileId, "Unavailable", "Injured: " + injuryNotes);
  ui.alert("Injury recorded. Master Players tab and active round tab updated.");
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