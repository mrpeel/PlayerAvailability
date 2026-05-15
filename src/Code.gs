/**
 * Laburnum CC Availability Tracker - Backend (Code.gs)
 * v2.3 - The "Serialisation Fix" (No Date objects in return)
 */
const APP_VERSION = "2.3";

function log(msg) {
  console.log(`[LCC ${APP_VERSION}] ${msg}`);
}

function doGet(e) {
  const roundNum = e.parameter.r || "";
  const template = HtmlService.createTemplateFromFile('Index');
  template.roundNum = roundNum;
  template.version = APP_VERSION;
  return template.evaluate()
      .setTitle('Laburnum CC Availability')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/**
 * Diagnostic tool to check system health
 */
function runDiagnostics() {
  const results = {
    version: APP_VERSION,
    spreadsheetLinked: false,
    spreadsheetUrl: null,
    tabsFound: [],
    errors: []
  };

  try {
    const ss = getSS();
    if (ss) {
      results.spreadsheetLinked = true;
      results.spreadsheetUrl = ss.getUrl();
      ['Players', 'Rounds', 'Availability'].forEach(name => {
        if (ss.getSheetByName(name)) results.tabsFound.push(name);
      });
    }
  } catch (e) {
    results.errors.push(e.toString());
  }
  return results;
}

/**
 * Creates a menu in the Google Sheet for administration
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('LCC Tracker')
      .addItem('Initialize / Fix Tabs', 'initializeSpreadsheet')
      .addItem('Run Diagnostics', 'runDiagnostics')
      .addItem('Get Web App URL', 'logDeploymentInfo')
      .addToUi();
}

function logDeploymentInfo() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert("Your Web App URL is: " + url);
}

/**
 * Normalizes phone numbers to E164 format (+61...)
 */
function normalizePhone(phone) {
  if (!phone) return "";
  let clean = String(phone).replace(/[^\d]/g, ''); // Strip all non-digits
  
  // AU Mobile logic: if starts with 04..., convert to +614...
  if (clean.startsWith('04') && clean.length === 10) {
    clean = '61' + clean.slice(1);
  }
  // If it doesn't start with 61 but is 9 digits (no leading 0), assume AU
  else if (clean.length === 9 && clean.startsWith('4')) {
    clean = '61' + clean;
  }
  
  return '+' + clean;
}

/**
 * Trigger: Automatically normalizes phone numbers when edited in the sheet
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Players') return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const phoneIdx = headers.indexOf('Phone');
  if (phoneIdx === -1) return;
  
  if (range.getColumn() === phoneIdx + 1 && range.getRow() > 1) {
    const val = range.getValue();
    if (val) {
      range.setValue(normalizePhone(val));
    }
  }
}

/**
 * Helper to get the spreadsheet, supporting both container-bound and standalone scripts
 */
function getSS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  
  const ssId = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (ssId) return SpreadsheetApp.openById(ssId);
  
  return null;
}

/**
 * Fetches all necessary data for the initial load
 */
function getInitialData(phone, roundNum) {
  try {
    log(`Fetching initial data for ${phone}, round ${roundNum}`);
    const ss = getSS();
    if (!ss) return { error: 'CRITICAL: Spreadsheet not found. Please open the Google Sheet and run LCC Tracker > Initialize.' };
    
    const playersSheet = ss.getSheetByName('Players');
    const roundsSheet = ss.getSheetByName('Rounds');
    const availabilitySheet = ss.getSheetByName('Availability');

    if (!playersSheet || !roundsSheet || !availabilitySheet) {
      return { error: 'DATABASE ERROR: Required tabs (Players, Rounds, Availability) are missing. Please run LCC Tracker > Initialize.' };
    }

    const normalizedInputPhone = normalizePhone(phone);

  // 1. Get Players for this phone number
  const playersData = playersSheet.getDataRange().getValues();
  const headers = playersData[0];
  const playerRows = playersData.slice(1);
  const pPhoneIdx = headers.indexOf('Phone');
  
  const household = playerRows
    .filter(row => normalizePhone(row[pPhoneIdx]) === normalizedInputPhone)
    .map(row => {
      const p = {};
      headers.forEach((h, i) => p[h] = row[i]);
      return p;
    });

  if (household.length === 0) {
    return { error: 'No players found for ' + normalizedInputPhone };
  }

  // 2. Get Round Info
  const roundsData = roundsSheet.getDataRange().getValues();
  const rHeaders = roundsData[0];
  const rRows = roundsData.slice(1);
  
  let currentRound;
  if (roundNum) {
    currentRound = rRows.find(row => String(row[rHeaders.indexOf('RoundNum')]) === String(roundNum));
  } else {
    currentRound = rRows[rRows.length - 1];
  }

  const roundInfo = {};
  if (currentRound) {
    rHeaders.forEach((h, i) => roundInfo[h] = currentRound[i]);
  } else {
    return { error: 'Round ' + roundNum + ' not found.' };
  }

  // 3. Get Availability
  const availData = availabilitySheet.getDataRange().getValues();
  const aHeaders = availData[0];
  const aRows = availData.slice(1);
  
  const householdAvail = {};
  household.forEach(player => {
    // Search backwards for the latest response
    const responseRow = aRows.slice().reverse().find(row => 
      String(row[aHeaders.indexOf('PlayerID')]) === String(player.ID) && 
      String(row[aHeaders.indexOf('RoundNum')]) === String(roundInfo.RoundNum)
    );
    householdAvail[player.ID] = responseRow ? responseRow[aHeaders.indexOf('Response')] : null;
  });

    // CRITICAL: GAS cannot return Date objects to the client. Must stringify.
    const roundInfoSafe = roundInfo ? {
      ...roundInfo,
      Date1: roundInfo.Date1 ? roundInfo.Date1.toString() : "",
      Date2: roundInfo.Date2 ? roundInfo.Date2.toString() : ""
    } : null;

    return {
      players: household,
      roundInfo: roundInfoSafe,
      availability: householdAvail,
      version: APP_VERSION
    };
  } catch (e) {
    log(`ERROR in getInitialData: ${e.message}`);
    return { error: `SERVER ERROR: ${e.message}` };
  }
}

/**
 * Saves availability
 */
function saveAvailability(playerId, roundNum, response) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Availability');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const pIdIdx = headers.indexOf('PlayerID');
  const rNumIdx = headers.indexOf('RoundNum');
  const respIdx = headers.indexOf('Response');
  const tsIdx = headers.indexOf('Timestamp');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][pIdIdx]) === String(playerId) && String(data[i][rNumIdx]) === String(roundNum)) {
      rowIndex = i + 1;
      break;
    }
  }

  const timestamp = new Date();
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, respIdx + 1).setValue(response);
    sheet.getRange(rowIndex, tsIdx + 1).setValue(timestamp);
  } else {
    const newRow = new Array(headers.length).fill("");
    newRow[tsIdx] = timestamp;
    newRow[pIdIdx] = playerId;
    newRow[rNumIdx] = roundNum;
    newRow[respIdx] = response;
    sheet.appendRow(newRow);
  }
  return true;
  } catch (e) {
    log(`ERROR in saveAvailability: ${e.message}`);
    throw e;
  }
}

/**
 * Updates status
 */
function setPlayerActive(playerId) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Players');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('ID');
  const statusIdx = headers.indexOf('GlobalStatus');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(playerId)) {
      sheet.getRange(i + 1, statusIdx + 1).setValue('Active');
      return true;
    }
  }
  return false;
}

/**
 * Wall of Shame
 */
function getWallOfShame(roundNum) {
  const ss = getSS();
  const players = ss.getSheetByName('Players').getDataRange().getValues();
  const pHeaders = players[0];
  const activePlayers = players.slice(1).filter(row => row[pHeaders.indexOf('GlobalStatus')] === 'Active');
  
  const availability = ss.getSheetByName('Availability').getDataRange().getValues();
  const aHeaders = availability[0];
  const respondedIds = new Set(
    availability.slice(1)
      .filter(row => String(row[aHeaders.indexOf('RoundNum')]) === String(roundNum))
      .map(row => String(row[aHeaders.indexOf('PlayerID')]))
  );
  
  return activePlayers
    .filter(row => !respondedIds.has(String(row[pHeaders.indexOf('ID')])))
    .map(row => row[pHeaders.indexOf('Name')]);
}

/**
 * INITIALIZATION: Run this once to setup the spreadsheet
 */
function initializeSpreadsheet() {
  let ss = getSS();
  
  if (!ss) {
    // Create a new spreadsheet if one doesn't exist
    ss = SpreadsheetApp.create("Laburnum CC Availability Tracker");
    PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
    console.log("Created new spreadsheet: " + ss.getUrl());
  }
  
  const tabs = {
    'Players': ['ID', 'Name', 'Phone', 'GlobalStatus'],
    'Rounds': ['RoundNum', 'Date1', 'Date2', 'Format'],
    'Availability': ['Timestamp', 'PlayerID', 'RoundNum', 'Response']
  };
  
  for (let name in tabs) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    sheet.getRange(1, 1, 1, tabs[name].length).setValues([tabs[name]])
         .setFontWeight('bold')
         .setBackground('#4d0012')
         .setFontColor('#ffffff');
    
    // Freeze header
    sheet.setFrozenRows(1);
  }
  
  // Add some sample data if empty
  const playersSheet = ss.getSheetByName('Players');
  if (playersSheet.getLastRow() === 1) {
    playersSheet.appendRow(['P001', 'Sample Player', '+61400000000', 'Active']);
  }
  
  const roundsSheet = ss.getSheetByName('Rounds');
  if (roundsSheet.getLastRow() === 1) {
    roundsSheet.appendRow(['1', new Date(), '', '1-Day']);
  }
  
  return "Spreadsheet initialized successfully. URL: " + ss.getUrl();
}
