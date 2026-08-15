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
    "Phone", "Phone2", "Phone3", "Phone4", "Email"
  ]];
  playerSheet.getRange(1, 1, 1, 13).setValues(playerHeaders).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);
  
  // 2. CONFIG TAB (Team configuration and WhatsApp message templates)
  var configSheet = ss.getSheetByName("Config") || ss.insertSheet("Config");
  configSheet.clear();
  
  // Section 1: Team Configuration
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
  
  // Section 2: WhatsApp & System Templates
  var templateStartRow = 4 + defaultTeams.length;
  configSheet.getRange(templateStartRow, 1, 1, 2).merge().setValue("WHATSAPP & MESSAGE TEMPLATES").setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);
  configSheet.getRange(templateStartRow + 1, 1, 1, 2).setValues([["Template Name", "Template Text"]]).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.zebraLight);
  var defaultTemplates = [
    ["Availability Callout", "🏏 *LCC ROUND AVAILABILITY* 🏏\nPlease submit your availability for this round: {url}"],
    ["Wall of Shame", "🚨 *WALL OF SHAME* 🚨\nThe following players have not yet entered their availability: {players}"],
    ["Selection Announcement", "🏏 *LABURNUM CC TEAMS - {date}* 🏏\n\n*1st XI vs {1st_opponent}* ({1st_venue}, {1st_format})\n{1st_team}\n\n*2nd XI vs {2nd_opponent}* ({2nd_venue}, {2nd_format})\n{2nd_team}\n\n*3rd XI vs {3rd_opponent}* ({3rd_venue}, {3rd_format})\n{3rd_team}\n\n*4th XI vs {4th_opponent}* ({4th_venue}, {4th_format})\n{4th_team}\n\n*5th XI vs {5th_opponent}* ({5th_venue}, {5th_format})\n{5th_team}"]
  ];
  configSheet.getRange(templateStartRow + 2, 1, defaultTemplates.length, 2).setValues(defaultTemplates);
  configSheet.getRange(templateStartRow, 1, 2 + defaultTemplates.length, 2).setBorder(true, true, true, true, true, true, LCC_SETUP_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);
  
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
    var playerWidths = [280, 120, 140, 180, 100, 90, 120, 150, 140, 140, 140, 140, 240];
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
  
  sh.getRange("A1").setValue("Active Presentation Target Sheet:").setFontWeight("bold");
  sh.getRange("B1").setValue("Placeholder").setBackground(LCC_SETUP_PALETTE.inputHighlight).setFontWeight("bold");
  
  var frames = [
    { name: "1st XI", start: 4 }, 
    { name: "2nd XI", start: 22 }, 
    { name: "3rd XI", start: 40 }, 
    { name: "4th XI", start: 58 }, 
    { name: "5th XI", start: 76 }
  ];
  
  var rowIdx = 3;
  frames.forEach(function(f) {
    sh.getRange(rowIdx, 1, 1, 2).merge().setValue(f.name + " Team").setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);
    sh.getRange(rowIdx + 1, 1).setValue("Opponent:");
    sh.getRange(rowIdx + 1, 2).setFormula(`=IFERROR(INDIRECT("'" & $B$1 & "'!B" & ${f.start + 1}), "")`);
    sh.getRange(rowIdx + 2, 1).setValue("Venue:");
    sh.getRange(rowIdx + 2, 2).setFormula(`=IFERROR(INDIRECT("'" & $B$1 & "'!B" & ${f.start + 2}), "")`);
    
    for (var s = 1; s <= 13; s++) {
      sh.getRange(rowIdx + 2 + s, 1).setValue("Slot " + s).setFontColor("#777777");
      sh.getRange(rowIdx + 2 + s, 2).setFormula(`=IFERROR(INDIRECT("'" & $B$1 & "'!B" & ${f.start + 3 + s}), "")`);
    }
    
    sh.getRange(rowIdx, 1, 16, 2).setBorder(true, true, true, true, true, true, LCC_SETUP_PALETTE.grayBorder, SpreadsheetApp.BorderStyle.SOLID);
    rowIdx += 17;
  });
  
  sh.autoResizeColumn(1);
  sh.autoResizeColumn(2);
}
