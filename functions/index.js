const { onRequest } = require("firebase-functions/v2/https");
const { google } = require('googleapis');
const cors = require('cors')({ origin: true });

async function getSheetsClient() {
  let credentials;
  if (process.env.GOOGLE_SERVICE_ACCOUNT) {
    credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  } else {
    try {
      credentials = require('./service-account.json');
    } catch (e) {
      throw new Error("No Google Credentials found. Set GOOGLE_SERVICE_ACCOUNT secret or add service-account.json locally.");
    }
  }

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

function normalizePhone(phone) {
  if (!phone) return "";
  let clean = phone.replace(/\D/g, "");
  if (clean.startsWith("04")) {
    clean = "61" + clean.substring(1);
  } else if (clean.startsWith("4")) {
    clean = "61" + clean;
  }
  return clean.startsWith("+") ? clean : "+" + clean;
}

exports.api = onRequest({ 
  secrets: ["SPREADSHEET_ID", "GOOGLE_SERVICE_ACCOUNT"],
  region: "australia-southeast1",
  maxInstances: 5,
  memory: "256MiB",
  timeoutSeconds: 30
}, async (req, res) => {
  return cors(req, res, async () => {
    try {
      const { action, phone, roundNum, playerId, response } = req.body;
      const sheets = await getSheetsClient();
      const spreadsheetId = process.env.SPREADSHEET_ID;

      if (!spreadsheetId) {
        throw new Error("SPREADSHEET_ID secret is not set.");
      }

      // --- ADMIN ACTIONS ---
      if (action === 'getAdminRounds' || action === 'getAdminData') {
        const normalizedInputPhone = normalizePhone(phone);
        const [adminsRes, roundsRes] = await Promise.all([
          sheets.spreadsheets.values.get({ spreadsheetId, range: 'Admins!A:B' }),
          sheets.spreadsheets.values.get({ spreadsheetId, range: 'Rounds!A:Z' })
        ]);

        const adminRows = adminsRes.data.values || [];
        const adminHeaders = adminRows[0] || [];
        const phoneIdx = adminHeaders.indexOf('Phone');
        
        const isAdmin = adminRows.slice(1).some(row => normalizePhone(row[phoneIdx]) === normalizedInputPhone);
        if (!isAdmin) {
          return res.status(403).json({ error: 'You do not have access to the Admin Dashboard.' });
        }

        const roundsData = roundsRes.data.values || [];
        const rHeaders = roundsData[0] || [];
        const roundNumIdx = rHeaders.indexOf('RoundNum');
        const date1Idx = rHeaders.indexOf('Date1');

        if (action === 'getAdminRounds') {
          // Filter out rounds that don't have a Date1
          const validRounds = roundsData.slice(1).filter(r => r[date1Idx] && r[date1Idx].trim() !== '');
          const rounds = validRounds.map(r => ({ num: r[roundNumIdx] }));
          const current = validRounds.length > 0 ? validRounds[validRounds.length - 1][roundNumIdx] : null;
          return res.json({ rounds, current });
        }

        let roundRow = roundNum ? roundsData.slice(1).find(r => r[roundNumIdx] == roundNum) : roundsData[roundsData.length - 1];
        if (!roundRow) return res.status(404).json({ error: 'Round not found' });

        return res.json({
          roundNum: roundRow[roundNumIdx],
          availMsg: roundRow[rHeaders.indexOf('AvailabilityMessage')],
          shameMsg: roundRow[rHeaders.indexOf('WallOfShameMessage')]
        });
      }

      // --- PLAYER ACTIONS ---
      if (action === 'getInitialData') {
        const normalizedInputPhone = normalizePhone(phone);
        
        const [playersRes, roundsRes, availRes] = await Promise.all([
          sheets.spreadsheets.values.get({ spreadsheetId, range: 'Players!A:Z' }),
          sheets.spreadsheets.values.get({ spreadsheetId, range: 'Rounds!A:Z', valueRenderOption: 'FORMATTED_VALUE' }),
          sheets.spreadsheets.values.get({ spreadsheetId, range: 'Availability!A:D' })
        ]);

        const playersData = playersRes.data.values || [];
        const roundsData = roundsRes.data.values || [];
        const availData = availRes.data.values || [];
        const headers = playersData[0] || [];

        const firstNameIdx = headers.indexOf('FirstName');
        const familyNameIdx = headers.indexOf('FamilyName');
        const phoneCols = ['Phone', 'Phone2', 'Phone3', 'Phone4']
          .map(c => headers.indexOf(c))
          .filter(i => i !== -1);
          
        const household = [];
        playersData.slice(1).forEach(row => {
          const isMatch = phoneCols.some(idx => normalizePhone(row[idx]) === normalizedInputPhone);
          if (isMatch) {
            household.push({
              ID: row[headers.indexOf('ID')],
              Name: `${row[firstNameIdx] || ''} ${row[familyNameIdx] || ''}`.trim(),
              GlobalStatus: row[headers.indexOf('GlobalStatus')] || 'Active',
              ExpectedReturnDate: row[headers.indexOf('ExpectedReturnDate')] || ''
            });
          }
        });

        if (household.length === 0) return res.status(404).json({ error: 'No players found' });

        const rHeaders = roundsData[0] || [];
        let roundRow = roundNum ? roundsData.slice(1).find(r => r[rHeaders.indexOf('RoundNum')] == roundNum) : roundsData[roundsData.length - 1];
        if (!roundRow) return res.status(404).json({ error: 'Round not found' });

        const roundInfo = {
          RoundNum: roundRow[rHeaders.indexOf('RoundNum')],
          Date1: roundRow[rHeaders.indexOf('Date1')],
          Date2: roundRow[rHeaders.indexOf('Date2')],
          Format: (roundRow[rHeaders.indexOf('Date2')] && roundRow[rHeaders.indexOf('Date2')].trim() !== '') ? '2-Day' : '1-Day'
        };

        const aHeaders = availData[0] || ['Timestamp', 'PlayerID', 'RoundNum', 'Response'];
        const householdAvail = {};
        household.forEach(player => {
          const rows = availData.slice(1).filter(r => r[aHeaders.indexOf('PlayerID')] == player.ID && r[aHeaders.indexOf('RoundNum')] == roundInfo.RoundNum);
          householdAvail[player.ID] = rows.length > 0 ? rows[rows.length - 1][aHeaders.indexOf('Response')] : null;
        });

        return res.json({ players: household, roundInfo, availability: householdAvail });
      }

      if (action === 'saveAvailability') {
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: 'Availability!A:D',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[new Date().toISOString(), playerId, roundNum, response]] }
        });
        return res.json({ success: true });
      }

      if (action === 'updateGlobalStatus') {
        const { status, returnDate } = req.body;
        const playersRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Players!A:Z' });
        const rows = playersRes.data.values;
        const headers = rows[0];
        const rowIndex = rows.findIndex(r => r[headers.indexOf('ID')] == playerId);
        if (rowIndex === -1) throw new Error('Player not found');

        const rowNum = rowIndex + 1;
        const statusIdx = headers.indexOf('GlobalStatus');
        const returnIdx = headers.indexOf('ExpectedReturnDate');

        const updates = [{ range: `Players!${String.fromCharCode(65 + statusIdx)}${rowNum}`, values: [[status]] }];
        if (returnIdx !== -1) updates.push({ range: `Players!${String.fromCharCode(65 + returnIdx)}${rowNum}`, values: [[returnDate || '']] });

        await sheets.spreadsheets.values.batchUpdate({ spreadsheetId, requestBody: { valueInputOption: 'USER_ENTERED', data: updates } });
        return res.json({ success: true });
      }

      res.status(400).json({ error: 'Invalid action' });
    } catch (error) {
      console.error('API Error:', error);
      res.status(500).json({ error: error.message });
    }
  });
});
