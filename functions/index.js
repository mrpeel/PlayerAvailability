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

      if (action === 'getInitialData') {
        const normalizedInputPhone = normalizePhone(phone);
        
        const [playersRes, roundsRes, availRes] = await Promise.all([
          sheets.spreadsheets.values.get({ spreadsheetId, range: 'Players!A:Z' }),
          sheets.spreadsheets.values.get({ spreadsheetId, range: 'Rounds!A:D', valueRenderOption: 'FORMATTED_VALUE' }),
          sheets.spreadsheets.values.get({ spreadsheetId, range: 'Availability!A:D' })
        ]);

        const playersData = playersRes.data.values;
        const roundsData = roundsRes.data.values;
        const availData = availRes.data.values;
        const headers = playersData[0];

        const firstNameIdx = headers.indexOf('FirstName');
        const familyNameIdx = headers.indexOf('FamilyName');
        const phoneCols = ['Phone', 'Phone2', 'Phone3', 'Phone4']
          .map(c => headers.indexOf(c))
          .filter(i => i !== -1);
          
        const household = [];
        
        playersData.slice(1).forEach(row => {
          const isMatch = phoneCols.some(idx => normalizePhone(row[idx]) === normalizedInputPhone);
          
          if (isMatch) {
            const firstName = row[firstNameIdx] || '';
            const familyName = row[familyNameIdx] || '';
            household.push({
              ID: row[headers.indexOf('ID')],
              Name: `${firstName} ${familyName}`.trim(),
              FirstName: firstName,
              FamilyName: familyName,
              GlobalStatus: row[headers.indexOf('GlobalStatus')] || 'Active',
              ExpectedReturnDate: row[headers.indexOf('ExpectedReturnDate')] || ''
            });
          }
        });

        if (household.length === 0) {
          return res.status(404).json({ error: 'No players found for ' + phone });
        }

        const rHeaders = roundsData[0];
        let roundRow = null;
        if (roundNum) {
          roundRow = roundsData.slice(1).find(r => r[rHeaders.indexOf('RoundNum')] == roundNum);
        } else {
          roundRow = roundsData[roundsData.length - 1];
        }

        const roundInfo = roundRow ? {
          RoundNum: roundRow[rHeaders.indexOf('RoundNum')],
          Date1: roundRow[rHeaders.indexOf('Date1')],
          Date2: roundRow[rHeaders.indexOf('Date2')],
          Format: (roundRow[rHeaders.indexOf('Date2')] && roundRow[rHeaders.indexOf('Date2')].trim() !== '') ? '2-Day' : '1-Day'
        } : null;

        const aHeaders = availData ? availData[0] : ['Timestamp', 'PlayerID', 'RoundNum', 'Response'];
        const householdAvail = {};
        household.forEach(player => {
          const rows = availData ? availData.slice(1).filter(r => 
            r[aHeaders.indexOf('PlayerID')] == player.ID && 
            r[aHeaders.indexOf('RoundNum')] == roundInfo.RoundNum
          ) : [];
          householdAvail[player.ID] = rows.length > 0 ? rows[rows.length - 1][aHeaders.indexOf('Response')] : null;
        });

        return res.json({
          players: household,
          roundInfo,
          availability: householdAvail
        });
      }

      if (action === 'saveAvailability') {
        const timestamp = new Date().toISOString();
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: 'Availability!A:D',
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [[timestamp, playerId, roundNum, response]]
          }
        });
        return res.json({ success: true });
      }

      if (action === 'updateGlobalStatus') {
        const { status, returnDate } = req.body;
        const playersRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Players!A:Z' });
        const rows = playersRes.data.values;
        const headers = rows[0];
        const idIdx = headers.indexOf('ID');
        const statusIdx = headers.indexOf('GlobalStatus');
        const returnIdx = headers.indexOf('ExpectedReturnDate');
        
        const rowIndex = rows.findIndex(r => r[idIdx] == playerId);
        if (rowIndex === -1) throw new Error('Player not found');

        const rowNum = rowIndex + 1;
        const updates = [
          {
            range: `Players!${String.fromCharCode(65 + statusIdx)}${rowNum}`,
            values: [[status]]
          }
        ];

        if (returnIdx !== -1) {
          updates.push({
            range: `Players!${String.fromCharCode(65 + returnIdx)}${rowNum}`,
            values: [[returnDate || '']]
          });
        }

        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId,
          requestBody: {
            valueInputOption: 'USER_ENTERED',
            data: updates
          }
        });
        return res.json({ success: true });
      }

      res.status(400).json({ error: 'Invalid action' });
    } catch (error) {
      console.error('API Error:', error);
      res.status(500).json({ error: error.message });
    }
  });
});
