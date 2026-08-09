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
 * Sub-task 1: Creates core data tables with exact headers including Phone1-4.
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
  
  // 2. FIXTURES TAB
  var fixSheet = ss.getSheetByName("Fixtures") || ss.insertSheet("Fixtures");
  fixSheet.clear();
  var fixHeaders = [["RoundID", "Game Date", "Match Format", "1st Opponent", "1st Venue", "2nd Opponent", "2nd Venue", "3rd Opponent", "3rd Venue", "4th Opponent", "4th Venue", "5th Opponent", "5th Venue"]];
  fixSheet.getRange(1, 1, 1, 13).setValues(fixHeaders).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);
  
  // 3. AVAILABILITY LOG TAB (Auditing only)
  var logSheet = ss.getSheetByName("Availability_Log") || ss.insertSheet("Availability_Log");
  logSheet.clear();
  var logHeaders = [["Timestamp", "ProfileID", "MatchDate", "Response", "Notes"]];
  logSheet.getRange(1, 1, 1, 5).setValues(logHeaders).setFontWeight("bold").setBackground(LCC_SETUP_PALETTE.maroonBg).setFontColor(LCC_SETUP_PALETTE.maroonFg);


  // Resize columns for readability
  [playerSheet, fixSheet, logSheet].forEach(function(s) {
    for (var col = 1; col <= s.getLastColumn(); col++) {
      s.autoResizeColumn(col);
    }
  });
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
