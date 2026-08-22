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
  
  // Data Validation for T20Squad (Col F) & GlobalStatus (Col G)
  var t20Rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No"], true)
    .setAllowInvalid(true)
    .build();
  playerSheet.getRange("F2:F200").setDataValidation(t20Rule);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Active", "Injured", "Long-Term Away", "Inactive"], true)
    .setAllowInvalid(true)
    .build();
  playerSheet.getRange("G2:G200").setDataValidation(statusRule);
  
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
  syncPresentationStagingHub(ss);
}

