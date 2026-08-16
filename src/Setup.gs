/**
 * ============================================================================
 * FILE 1: Setup.gs
 * Laburnum CC - System Initialization & One-Off Structural Setup
 * ============================================================================
 */


var LCC_SETUP_PALETTE = {
  maroonBg: "#6A1B29",
  maroonFg: "#ffffff",
  goldBg: "#F4B41A",
  goldFg: "#111111",
  inputHighlight: "#FFF6E5",
  grayBorder: "#e5e7eb",
  zebraLight: "#f9fafb"
};


/**
 * SINGLE MASTER FUNCTION
 * Run this function once to build the entire base workbook infrastructure.
 */
function setupInitialSystem() {
  var ss = getSS();
  
  if (!ss) {
    ss = SpreadsheetApp.create("Laburnum CC Availability Tracker");
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  }
  
  createCoreDatabaseSheets(ss);
  buildPresentationStagingHub(ss);
  
  return "Setup complete. Ready for data.";
}


/**
 * Sub-task 1: Creates core data tables with exact headers including Phone1-4, Config, and Fixtures.
 */
function createCoreDatabaseSheets(ss) {
  // 1. PLAYERS MASTER TAB (Includes Phone, Phone2, Phone3, Phone4 for juniors/parents)
  var playerSheet = ss.getSheetByName("Players") || ss.insertSheet("Players");
  playerSheet.clear();
  var playerHeaders = [[
    "ProfileID", "FirstName", "LastName", "FullName",
    "JuniorLevel", "T20Squad", "GlobalStatus", "ExpectedReturnDate",
    "Phone", "Phone2", "Phone3", "Phone4", "Email", "PhotoUrl"
  ]];
  playerSheet.getRange(1, 1, 1, 14).setValues(playerHeaders).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);
  
  // 2. CONFIG TAB (Team configuration)
  var configSheet = ss.getSheetByName("Config") || ss.insertSheet("Config");
  configSheet.clear();
  
  // Team Configuration
  configSheet.getRange(1, 1, 1, 3).merge().setValue("TEAM CONFIGURATION").setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);
  var teamHeaders = [["Internal Team Name", "Competition", "Play Cricket Team Name"]];
  configSheet.getRange(2, 1, 1, 3).setValues(teamHeaders).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.zebraLight);
  var defaultTeams = [
    ["1st XI", "BHRDCA Senior Competition", "Laburnum - 1st XI"],
    ["2nd XI", "BHRDCA Senior Competition", "Laburnum - 2nd XI"],
    ["3rd XI", "BHRDCA Senior Competition", "Laburnum - 3rd XI"],
    ["4th XI", "BHRDCA Senior Competition", "Laburnum - 4th XI"],
    ["5th XI", "BHRDCA Senior Competition", "Laburnum - 5th XI"],
    ["T20 1st XI", "BHRDCA T20 Competition", "Laburnum T20 1st XI"],
    ["T20 2nd XI", "BHRDCA T20 Competition", "Laburnum T20 2nd XI"]
  ];
  configSheet.getRange(3, 1, defaultTeams.length, 3).setValues(defaultTeams);
  configSheet.getRange(1, 1, 2 + defaultTeams.length, 3).setBorder(true, true, true, true, true, true, LCC_SETUP_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);
  
  // 3. FIXTURES TAB (Game Date as PK, per-team Round, Format, Opponent, Venue)
  var fixSheet = ss.getSheetByName("Fixtures") || ss.insertSheet("Fixtures");
  fixSheet.clear();
  var fixHeaders = [[
    "Game Date",
    "1st Round", "1st Format", "1st Opponent", "1st Venue",
    "2nd Round", "2nd Format", "2nd Opponent", "2nd Venue",
    "3rd Round", "3rd Format", "3rd Opponent", "3rd Venue",
    "4th Round", "4th Format", "4th Opponent", "4th Venue",
    "5th Round", "5th Format", "5th Opponent", "5th Venue"
  ]];
  fixSheet.getRange(1, 1, 1, 21).setValues(fixHeaders).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);
  
  // 4. AVAILABILITY LOG TAB (Auditing only)
  var logSheet = ss.getSheetByName("Availability_Log") || ss.insertSheet("Availability_Log");
  logSheet.clear();
  var logHeaders = [["Timestamp", "ProfileID", "MatchDate", "Response", "Notes"]];
  logSheet.getRange(1, 1, 1, 5).setValues(logHeaders).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);

  // 5. ADMINS TAB
  var adminSheet = ss.getSheetByName("Admins") || ss.insertSheet("Admins");
  if (adminSheet.getLastRow() === 0) {
    var adminHeaders = [["Name", "Phone"]];
    adminSheet.getRange(1, 1, 1, 2).setValues(adminHeaders).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);
  }

  // Pre-set sensible column widths for all tabs
  applyAllStandardColumnWidths(ss);
}


/**
 * Applies generous, sensible column widths across all core database sheets.
 */
function applyAllStandardColumnWidths(ss) {
  if (!ss) ss = getSS();
  if (!ss) return;

  var playerSheet = ss.getSheetByName("Players");
  if (playerSheet) {
    var playerWidths = [120, 110, 110, 150, 90, 80, 110, 120, 120, 120, 120, 120, 180, 200];
    for (var p = 0; p < playerWidths.length; p++) {
      playerSheet.setColumnWidth(p + 1, playerWidths[p]);
    }
  }

  var configSheet = ss.getSheetByName("Config");
  if (configSheet) {
    configSheet.setColumnWidth(1, 200); // Internal Team Name / Template Name
    configSheet.setColumnWidth(2, 280); // Competition / Template Text
    configSheet.setColumnWidth(3, 240); // Play Cricket Team Name
  }

  var fixSheet = ss.getSheetByName("Fixtures");
  if (fixSheet) {
    fixSheet.setColumnWidth(1, 110); // Game Date
    var numCols = fixSheet.getLastColumn();
    for (var c = 2; c <= numCols; c += 4) {
      fixSheet.setColumnWidth(c, 90);      // Round
      fixSheet.setColumnWidth(c + 1, 100); // Format
      fixSheet.setColumnWidth(c + 2, 220); // Opponent
      fixSheet.setColumnWidth(c + 3, 200); // Venue
    }
  }

  var logSheet = ss.getSheetByName("Availability_Log");
  if (logSheet) {
    logSheet.setColumnWidth(1, 170); // Timestamp
    logSheet.setColumnWidth(2, 280); // ProfileID
    logSheet.setColumnWidth(3, 110); // MatchDate
    logSheet.setColumnWidth(4, 110); // Response
    logSheet.setColumnWidth(5, 260); // Notes
  }

  var adminSheet = ss.getSheetByName("Admins");
  if (adminSheet) {
    adminSheet.setColumnWidth(1, 180); // Name
    adminSheet.setColumnWidth(2, 150); // Phone
  }

  // Apply clean Hanken Grotesk typography across all tabs
  try {
    ss.getSheets().forEach(function(s) {
      s.getRange(1, 1, Math.min(s.getMaxRows(), 150), Math.min(s.getMaxColumns(), 26)).setFontFamily("Hanken Grotesk");
    });
  } catch (e) {}
}


/**
 * Sub-task 2: Builds the permanent Presentation Staging Hub for Google Slides linking.
 */
function buildPresentationStagingHub(ss) {
  var sh = ss.getSheetByName("Presentation_Staging") || ss.insertSheet("Presentation_Staging");
  sh.clear();
  
  // Row 1: Header and Dropdown Selection
  sh.getRange("A1").setValue("Round to present").setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg).setHorizontalAlignment("center");
  sh.getRange("B1").setValue("").setNumberFormat("@").setBackground(LCC_SETUP_PALETTE.inputHighlight).setFontWeight("bold").setHorizontalAlignment("center");
  
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
    
    // Team Banner (Cols A to D)
    sh.getRange(rowIdx, 1, 1, 4).merge().setValue(f.name).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg).setHorizontalAlignment("center");
    
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
      sh.getRange(currRow, 1).setValue(slotRoles[idx]).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.zebraLight);
      
      // Col B: Clean Player Name (Stripped of junior tags like (U16))
      sh.getRange(currRow, 2).setFormula('=IFERROR(IF(INDIRECT("\'" & ' + targetTabExpr + ' & "\'!B' + currRow + '")="", "", TRIM(REGEXREPLACE(TO_TEXT(INDIRECT("\'" & ' + targetTabExpr + ' & "\'!B' + currRow + '")), "\\s*\\(.*?\\)", ""))), "")');
      
      // Col C: Dynamic Profile ID Lookup from Players tab (Hidden)
      sh.getRange(currRow, 3).setFormula('=IFERROR(IF(B' + currRow + '="", "", INDEX(Players!$A$2:$A, MATCH(B' + currRow + ', Players!$D$2:$D, 0))), IFERROR(INDEX(Players!$A$2:$A, MATCH(B' + currRow + ', Players!$B$2:$B & " " & Players!$C$2:$C, 0)), ""))');
      
      // Col D: Dynamic Headshot Image from Players tab PhotoUrl
      sh.getRange(currRow, 4).setFormula('=IFERROR(IF(B' + currRow + '="", "", IMAGE(INDEX(Players!$N$2:$N, MATCH(C' + currRow + ', Players!$A$2:$A, 0)))), "")');
    }
    
    // Set borders for this team frame (18 rows: 1 banner + 4 metadata + 13 slots)
    sh.getRange(rowIdx, 1, 18, 4).setBorder(true, true, true, true, true, true, LCC_SETUP_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);
  });
  
  sh.setColumnWidth(1, 120); // Col A (Slot Role)
  sh.setColumnWidth(2, 210); // Col B (Clean Player Name)
  sh.setColumnWidth(3, 20);  // Col C (Hidden Profile ID)
  sh.setColumnWidth(4, 100); // Col D (Photo)
  
  sh.hideColumns(3); // Hide Column C (Profile ID)
  
  try {
    sh.getRange(1, 1, 100, 4).setFontFamily("Hanken Grotesk");
  } catch (e) {}
}

