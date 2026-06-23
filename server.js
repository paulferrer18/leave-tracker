const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Google Sheets Auth ────────────────────────────────────
function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

const SHEET_ID = process.env.SHEET_ID;

// ── Helper: ensure a sheet tab exists ────────────────────
async function ensureSheet(sheets, title) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const exists = meta.data.sheets.some(s => s.properties.title === title);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title } } }]
      }
    });
    // Add headers
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${title}!A1:C1`,
      valueInputOption: 'RAW',
      requestBody: { values: [['date', 'employee', 'type']] }
    });
  }
}

// ── Routes ────────────────────────────────────────────────

// GET /api/employees
app.get('/api/employees', async (req, res) => {
  try {
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const empSheet = meta.data.sheets.find(s => s.properties.title === 'employees');
    if (!empSheet) {
      return res.json({ employees: ['Maria', 'Juan', 'Ana', 'Carlo', 'Sofia'] });
    }
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'employees!A2:A',
    });
    const rows = result.data.values || [];
    res.json({ employees: rows.map(r => r[0]).filter(Boolean) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/employees
app.post('/api/employees', async (req, res) => {
  try {
    const { employees } = req.body;
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const exists = meta.data.sheets.some(s => s.properties.title === 'employees');
    if (!exists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: [{ addSheet: { properties: { title: 'employees' } } }] }
      });
    }
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SHEET_ID, range: 'employees!A:A'
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: 'employees!A1',
      valueInputOption: 'RAW',
      requestBody: { values: [['name'], ...employees.map(e => [e])] }
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// GET /api/leaves?month=2026-06
app.get('/api/leaves', async (req, res) => {
  try {
    const { month } = req.query; // e.g. "2026-06"
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    await ensureSheet(sheets, 'leaves');
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'leaves!A2:C',
    });
    const rows = result.data.values || [];
    // Filter by month and build leaveData object
    const leaveData = {};
    rows.forEach(([date, employee, type]) => {
      if (!date || !employee || !type) return;
      if (month && !date.startsWith(month)) return;
      if (!leaveData[date]) leaveData[date] = [];
      leaveData[date].push({ emp: employee, type });
    });
    res.json({ leaveData });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/leaves/add
app.post('/api/leaves/add', async (req, res) => {
  try {
    const { date, employee, type } = req.body;
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    await ensureSheet(sheets, 'leaves');
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: 'leaves!A:C',
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[date, employee, type]] }
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/leaves/remove
app.post('/api/leaves/remove', async (req, res) => {
  try {
    const { date, employee } = req.body;
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    await ensureSheet(sheets, 'leaves');
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: 'leaves!A:C'
    });
    const rows = result.data.values || [];
    // Find row index (1-based, skip header at row 1)
    const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const leavesSheet = sheetMeta.data.sheets.find(s => s.properties.title === 'leaves');
    const sheetId = leavesSheet.properties.sheetId;
    // Collect all row indices to delete (reverse order so indices don't shift)
    const toDelete = [];
    rows.forEach((row, i) => {
      if (i === 0) return; // skip header
      if (row[0] === date && row[1] === employee) toDelete.push(i);
    });
    if (toDelete.length === 0) return res.json({ ok: true });
    // Delete in reverse
    toDelete.reverse();
    const requests = toDelete.map(i => ({
      deleteDimension: {
        range: { sheetId, dimension: 'ROWS', startIndex: i, endIndex: i + 1 }
      }
    }));
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests }
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Leave tracker running on port ${PORT}`));
