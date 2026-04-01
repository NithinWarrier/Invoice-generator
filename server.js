require('dotenv').config({ path: '.env.local' });
const express = require('express');
const { google } = require('googleapis');

const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

// Column config (0-based index)
const COL_CHECKMARK = 6;  // G
const COL_DESCRIPTION = 1;  // B
const COL_HOURS = 4;  // E
const BALANCE_RANGE = 'K16';

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/documents',
  'https://www.googleapis.com/auth/drive',
];

function getAuthClient() {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
  oauth2Client.setCredentials({
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
  });
  return oauth2Client;
}

// ── One-time auth flow ────────────────────────────────────────────────────
// Step 1: visit http://localhost:3000/auth → redirects to Google consent screen
app.get('/auth', (req, res) => {
  const client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
  const url = client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',   // forces Google to return a refresh_token every time
    scope: SCOPES,
  });
  res.redirect(url);
});

// Step 2: Google redirects back here with ?code=...
// The page shows you the refresh token — copy it into .env.local
app.get('/api/auth/callback', async (req, res) => {
  const { code } = req.query;
  if (!code) return res.status(400).send('Missing code parameter');

  try {
    const client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      process.env.GOOGLE_REDIRECT_URI
    );
    const { tokens } = await client.getToken(code);
    res.send(`
      <html><body style="font-family:monospace;padding:2rem;background:#0d0f14;color:#e8eaf0">
        <h2 style="color:#51cf66">✅ Auth successful!</h2>
        <p>Copy the line below into your <code>.env.local</code> file:</p>
        <pre style="background:#1e2330;padding:1rem;border-radius:8px;color:#748ffc;word-break:break-all">GOOGLE_REFRESH_TOKEN=${tokens.refresh_token}</pre>
        <p style="color:#7a8094;margin-top:1rem">Then restart the server with <code>node server.js</code> and go to <a href="/" style="color:#5c7cfa">the app</a>.</p>
      </body></html>
    `);
  } catch (err) {
    res.status(500).send('Error getting token: ' + err.message);
  }
});
// ─────────────────────────────────────────────────────────────────────────

app.get('/api/generate', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;

    // Get the first sheet name
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetName = meta.data.sheets[0].properties.title;

    // Fetch balance and all data rows in parallel
    const [balanceResp, rowsResp] = await Promise.all([
      sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!${BALANCE_RANGE}`,
      }),
      sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A2:G1000`,
      }),
    ]);

    // Parse balance
    const rawBalance = balanceResp.data.values?.[0]?.[0] ?? '0';
    const balance = rawBalance.replace(/[^0-9.-]/g, '') || '0';

    // Process rows: filter out checked rows
    const rows = rowsResp.data.values || [];
    const summaries = [];
    const unbilledRowNums = [];  // 1-based spreadsheet row numbers
    let totalHours = 0;
    let checkedCount = 0;
    const monthCounts = {};

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const sheetRowNum = i + 2; // data starts at row 2

      // Stop at the first row where column A is empty
      if (!(row[0] || '').trim()) break;

      const checkmark = (row[COL_CHECKMARK] || '').trim().toUpperCase();
      const isBilled = checkmark === 'TRUE';
      if (isBilled) { checkedCount++; continue; }

      // Track this as an unbilled row
      unbilledRowNums.push(sheetRowNum);

      // Track month from column A date (format: dd/mm/yyyy) for unbilled rows
      const dateStr = (row[0] || '').trim();
      if (dateStr) {
        const parts = dateStr.split('/');
        if (parts.length === 3) {
          const d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
          if (!isNaN(d)) {
            const key = d.toLocaleString('en-US', { month: 'long', year: 'numeric' });
            monthCounts[key] = (monthCounts[key] || 0) + 1;
          }
        }
      }

      const description = (row[COL_DESCRIPTION] || '').trim();
      const hours = parseFloat(row[COL_HOURS]) || 0;

      if (description) summaries.push(description);
      if (hours > 0) totalHours += hours;
    }

    // Pick the most frequent month among unbilled rows
    const month = Object.entries(monthCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || null;

    res.json({
      balance: parseFloat(balance),
      totalHours: Math.round(totalHours * 100) / 100,
      summaries,
      month,
      unbilledRowNums,
    });
  } catch (err) {
    console.error('Error fetching sheet:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── Lightweight month detection (called on page load) ─────────────────────
app.get('/api/month', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;

    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetName = meta.data.sheets[0].properties.title;

    const rowsResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A2:G1000`,
    });

    const rows = rowsResp.data.values || [];
    const monthCounts = {};

    for (const row of rows) {
      if (!(row[0] || '').trim()) break;
      const checkmark = (row[COL_CHECKMARK] || '').trim().toUpperCase();
      if (checkmark === 'TRUE') continue;

      const parts = (row[0] || '').trim().split('/');
      if (parts.length === 3) {
        const d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        if (!isNaN(d)) {
          const key = d.toLocaleString('en-US', { month: 'long', year: 'numeric' });
          monthCounts[key] = (monthCounts[key] || 0) + 1;
        }
      }
    }

    const month = Object.entries(monthCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || null;
    res.json({ month });
  } catch (err) {
    console.error('Error detecting month:', err.message);
    res.status(500).json({ month: null });
  }
});
// ─────────────────────────────────────────────────────────────────────────

// ── Create Invoice: populate Google Doc template ─────────────────────────
app.post('/api/create-invoice', async (req, res) => {
  const { balance, totalHours, summaries, description, month, unbilledRowNums = [] } = req.body;
  const docId = process.env.GOOGLE_DOC_ID;

  if (!docId) {
    return res.status(400).json({ error: 'GOOGLE_DOC_ID is not set in .env.local' });
  }

  try {
    const auth = getAuthClient();
    const drive = google.drive({ version: 'v3', auth });
    const docs = google.docs({ version: 'v1', auth });

    // 1. Copy the template so the original stays pristine
    const copyRes = await drive.files.copy({
      fileId: docId,
      requestBody: {
        name: `Understood_Invoice_${month || new Date().toLocaleString('en-US', { month: 'long', year: 'numeric' })}`,
      },
    });
    const newDocId = copyRes.data.id;

    // 2. Format values
    const balanceStr = '$' + Number(balance).toLocaleString('en-IN', { minimumFractionDigits: 2 });
    const hoursStr = totalHours + ' hrs';
    const dateStr = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

    // 3. Replace placeholders in the copied doc
    await docs.documents.batchUpdate({
      documentId: newDocId,
      requestBody: {
        requests: [
          {
            replaceAllText: {
              containsText: { text: '{{BALANCE_DUE}}', matchCase: true },
              replaceText: balanceStr,
            },
          },
          {
            replaceAllText: {
              containsText: { text: '{{TOTAL_HOURS}}', matchCase: true },
              replaceText: hoursStr,
            },
          },
          {
            replaceAllText: {
              containsText: { text: '{{DATE}}', matchCase: true },
              replaceText: dateStr,
            },
          },
          {
            replaceAllText: {
              containsText: { text: '{{DESCRIPTION}}', matchCase: true },
              replaceText: description || '',
            },
          },

        ],
      },
    });

    const docUrl = `https://docs.google.com/document/d/${newDocId}/edit`;

    // 4. Tick the checkbox (column G) for each invoiced row in the sheet
    if (unbilledRowNums.length > 0) {
      const sheets = google.sheets({ version: 'v4', auth });
      const spreadsheetId = process.env.GOOGLE_SHEET_ID;
      const meta = await sheets.spreadsheets.get({ spreadsheetId });
      const sheetId = meta.data.sheets[0].properties.sheetId;  // numeric ID for batchUpdate

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: unbilledRowNums.map(rowNum => ({
            updateCells: {
              range: {
                sheetId,
                startRowIndex: rowNum - 1,  // 0-based
                endRowIndex: rowNum,
                startColumnIndex: 6,        // G (0-based)
                endColumnIndex: 7,
              },
              rows: [{ values: [{ userEnteredValue: { boolValue: true } }] }],
              fields: 'userEnteredValue',
            },
          })),
        },
      });
    }

    res.json({ docUrl });
  } catch (err) {
    const detail = err.response?.data || err.message;
    console.error('Error creating invoice doc:', JSON.stringify(detail, null, 2));
    res.status(500).json({ error: err.response?.data?.error?.message || err.message });
  }
});
// ─────────────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Invoice generator running at http://localhost:${PORT}`);
  console.log(`To get a refresh token, visit: http://localhost:${PORT}/auth`);
});
