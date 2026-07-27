/**
 * lib/sheets.js
 * ------------------------------------------------------------------
 * Node port of the Apps Script Code.gs logic, using the Sheets API v4
 * with a service account (no Apps Script, no OAuth popup for users).
 *
 * Every function here mirrors a same-named function in the original
 * Code.gs as closely as possible so the two stay easy to compare.
 * ------------------------------------------------------------------
 */
const { google } = require('googleapis');

// ====================== CONFIG — set these as Netlify env vars ======================
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_DATA_NAME = process.env.SHEET_DATA_NAME || 'Sheet1';
const SHEET_MAIL_NAME = process.env.SHEET_MAIL_NAME || 'Mail';
const SHEET_PLANNED_NAME = process.env.SHEET_PLANNED_NAME || 'Planned';
const SHEET_RESPONSE_NAME = process.env.SHEET_RESPONSE_NAME || 'Response';
const HEADER_SCAN_MAX_ROWS = 5;
const DATE_START_COL = 23; // column W, 1-based
const PENDING_GRACE_HOURS = 48;
const TIMEZONE = process.env.SHEET_TIMEZONE || 'Asia/Kolkata';
// =======================================================================================

let _sheetsClient = null;
async function getSheetsApi() {
  if (_sheetsClient) return _sheetsClient;
  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  await auth.authorize();
  _sheetsClient = google.sheets({ version: 'v4', auth });
  return _sheetsClient;
}

// FORMATTED_VALUE gives us the string exactly as displayed in the sheet UI,
// which does for us in one step what Code.gs had to do manually (checking
// both "is this a real Date object" and "is this D-Month-YYYY text").
async function readRange(range) {
  const api = await getSheetsApi();
  const res = await api.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueRenderOption: 'FORMATTED_VALUE'
  });
  return res.data.values || [];
}

async function appendRow(sheetName, rowValues) {
  const api = await getSheetsApi();
  await api.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: sheetName + '!A1',
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [rowValues] }
  });
}

// ------------------------------------------------------------------
// Header auto-detection (mirrors findHeaderRowByLabel_ / findHeaderRow_)
// ------------------------------------------------------------------
function findHeaderRowIndex(block, label) {
  const target = label.trim().toLowerCase();
  for (let r = 0; r < block.length; r++) {
    const has = (block[r] || []).some(cell => String(cell || '').trim().toLowerCase() === target);
    if (has) return r + 1; // 1-based
  }
  return null;
}

async function findHeaderRow() {
  const block = await readRange(`${SHEET_DATA_NAME}!A1:ZZ${HEADER_SCAN_MAX_ROWS}`);
  return findHeaderRowIndex(block, 'mb code') || 1; // fallback: row 1
}

async function getHeaders() {
  const headerRow = await findHeaderRow();
  const rows = await readRange(`${SHEET_DATA_NAME}!A${headerRow}:ZZ${headerRow}`);
  return rows[0] || [];
}

async function getDataRows() {
  const headerRow = await findHeaderRow();
  const rows = await readRange(`${SHEET_DATA_NAME}!A${headerRow + 1}:ZZ100000`);
  return rows;
}

function normalizeCode(code) {
  return String(code || '').trim().toUpperCase().replace(/\s+/g, '');
}
function normalizeMail(mail) {
  return String(mail || '').trim().toLowerCase();
}

// Exact match first, then "header contains name" fallback — same as colIndex_
function colIndex(headers, nameOrNames) {
  const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim().toLowerCase();
    for (const n of names) if (h === n.toLowerCase()) return i;
  }
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim().toLowerCase();
    for (const n of names) if (h.indexOf(n.toLowerCase()) !== -1) return i;
  }
  return -1;
}

const MONTH_MAP = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
function parseHeaderDate(label) {
  const parts = String(label || '').split('-');
  if (parts.length !== 3) return null;
  const day = parseInt(parts[0], 10);
  const month = MONTH_MAP[parts[1].toLowerCase().slice(0, 3)];
  const year = parseInt(parts[2], 10);
  if (isNaN(day) || month === undefined || isNaN(year)) return null;
  return new Date(year, month, day);
}

// Same primary-scan-then-fallback-to-full-row behaviour as getDateColumns_
function getDateColumns(headers) {
  const re = /^\d{1,2}-[A-Za-z]+-\d{4}$/;
  const scanFrom = (startIdx) => {
    const cols = [];
    for (let i = startIdx; i < headers.length; i++) {
      const val = String(headers[i] || '').trim();
      if (re.test(val)) {
        const d = parseHeaderDate(val);
        if (d) cols.push({ index: i, label: val, date: d });
      }
    }
    return cols;
  };
  let cols = scanFrom(DATE_START_COL - 1);
  if (cols.length === 0) cols = scanFrom(0);
  cols.sort((a, b) => a.date - b.date);
  return cols;
}

// ------------------------------------------------------------------
// Mail lookup + row lookup
// ------------------------------------------------------------------
async function findCodeForMail(mail) {
  const rows = await readRange(`${SHEET_MAIL_NAME}!A1:B100000`);
  const target = normalizeMail(mail);
  for (const [mbCode, mailVal] of rows) {
    if (mailVal && normalizeMail(mailVal) === target) return String(mbCode || '').trim();
  }
  return null;
}

async function findRowByCode(code) {
  const headers = await getHeaders();
  const rows = await getDataRows();
  const vendorIdx = colIndex(headers, ['Vendor Emp. Code', 'Vendor Code', 'Vendor']);
  const mbIdx = colIndex(headers, ['MB Code']);
  const target = normalizeCode(code);
  for (const row of rows) {
    const vendorVal = vendorIdx >= 0 ? normalizeCode(row[vendorIdx]) : '';
    const mbVal = mbIdx >= 0 ? normalizeCode(row[mbIdx]) : '';
    if ((vendorVal && vendorVal === target) || (mbVal && mbVal === target)) {
      return { headers, row };
    }
  }
  return null;
}

function buildProfile(headers, row) {
  const idx = (n) => colIndex(headers, n);
  const get = (n) => { const i = idx(n); return i >= 0 ? (row[i] || '') : ''; };
  return {
    mbCode: get(['MB Code']),
    vendorCode: get(['Vendor Emp. Code', 'Vendor Code', 'Vendor']),
    name: get(['Associate Name']),
    city: get(['City']),
    state: get(['State']),
    designation: get(['Designation']),
    teamLeader: get(['Team Leader Name']),
    status: get(['Status'])
  };
}

// ------------------------------------------------------------------
// PLANNED sheet
// ------------------------------------------------------------------
async function findPlannedHeaderRow() {
  const block = await readRange(`${SHEET_PLANNED_NAME}!A1:ZZ${HEADER_SCAN_MAX_ROWS}`).catch(() => []);
  return findHeaderRowIndex(block, 'society id') || 1;
}

async function getPlannedHeaders() {
  const headerRow = await findPlannedHeaderRow();
  const rows = await readRange(`${SHEET_PLANNED_NAME}!A${headerRow}:ZZ${headerRow}`).catch(() => []);
  return rows[0] || null;
}

async function getPlannedRows() {
  const headerRow = await findPlannedHeaderRow();
  const rows = await readRange(`${SHEET_PLANNED_NAME}!A${headerRow + 1}:ZZ100000`).catch(() => []);
  return rows;
}

function plannedCellMatchesCode(cell, targetNormalized) {
  const raw = String(cell || '');
  if (!raw.trim() || !targetNormalized) return false;
  const tokens = raw.split(',');
  for (let tok of tokens) {
    tok = tok.trim();
    if (!tok) continue;
    if (normalizeCode(tok) === targetNormalized) return true;
    if (tok.indexOf('_') !== -1) {
      for (const p of tok.split('_')) {
        if (normalizeCode(p) === targetNormalized) return true;
      }
    }
  }
  return false;
}

function parsePlannedDateCell(cell) {
  const val = String(cell || '').trim();
  if (!val) return null;
  if (val.indexOf('/') !== -1) {
    const parts = val.split('/');
    if (parts.length === 3) {
      const month = parseInt(parts[0], 10) - 1;
      const day = parseInt(parts[1], 10);
      const year = parseInt(parts[2], 10);
      if (!isNaN(month) && !isNaN(day) && !isNaN(year)) return new Date(year, month, day);
    }
  }
  const viaDash = parseHeaderDate(val);
  if (viaDash) return viaDash;
  return null;
}

async function getPlannedForCode(code, todayMidnight) {
  try {
    const headers = await getPlannedHeaders();
    if (!headers) return [];
    const rows = await getPlannedRows();
    const target = normalizeCode(code);
    const idx = (n) => colIndex(headers, n);

    const idxSocietyName = idx(['Society Name']);
    const idxActivity = idx(['Type of Activity']);
    const idxDate = idx(['Date of Activation']);
    const idxType = idx(['Internal/External/FOC', 'Internal / External / FOC']);
    const idxProm = idx(['BDE prom code', 'BDE Prom Code', 'Prom Code', 'BDE Code']);
    const idxLat = idx(['Lat', 'Latitude']);
    const idxLong = idx(['Long', 'Lng', 'Longitude']);
    if (idxProm < 0) return [];

    const results = [];
    for (const row of rows) {
      if (!plannedCellMatchesCode(row[idxProm], target)) continue;
      const activationDate = idxDate >= 0 ? parsePlannedDateCell(row[idxDate]) : null;
      if (!activationDate) continue;
      if (activationDate.getTime() !== todayMidnight.getTime()) continue;

      const lat = idxLat >= 0 && row[idxLat] !== '' ? parseFloat(row[idxLat]) : NaN;
      const lng = idxLong >= 0 && row[idxLong] !== '' ? parseFloat(row[idxLong]) : NaN;
      const hasCoords = !isNaN(lat) && !isNaN(lng);

      results.push({
        societyName: idxSocietyName >= 0 ? String(row[idxSocietyName] || '') : '',
        activity: idxActivity >= 0 ? String(row[idxActivity] || '') : '',
        type: idxType >= 0 ? String(row[idxType] || '') : '',
        mapUrl: hasCoords ? `https://www.google.com/maps/search/?api=1&query=${lat},${lng}` : null
      });
    }
    return results;
  } catch (err) {
    return []; // Planned-sheet problems must never break the passbook
  }
}

// ------------------------------------------------------------------
// Status classification (mirrors classifyStatus_)
// ------------------------------------------------------------------
function classifyStatus(raw, colDate, today, now) {
  const v = String(raw || '').toUpperCase();
  if (v === 'P' || v === 'PRESENT') return { code: 'P', label: 'Present' };
  if (v === 'L' || v === 'LEAVE') return { code: 'L', label: 'Leave' };
  if (['HD', 'H', 'HALF DAY', 'HALFDAY', 'HALF-DAY'].includes(v)) return { code: 'HD', label: 'Half Day' };
  if (['WO', 'W.OFF', 'WEEK OFF', 'WEEKOFF'].includes(v)) return { code: 'WO', label: 'Week Off' };
  if (v === 'A' || v === 'ABSENT') return { code: 'A', label: 'Absent' };

  if (v === '' || v === '-') {
    const deadline = new Date(colDate.getTime() + PENDING_GRACE_HOURS * 3600 * 1000);
    const msLeft = deadline - now;
    if (msLeft > 0) {
      return {
        code: 'PENDING', label: 'Pending', deadlineIso: deadline.toISOString(),
        msLeft, hoursLeft: Math.floor(msLeft / 3600000), minutesLeft: Math.floor((msLeft % 3600000) / 60000)
      };
    }
    return { code: 'A', label: 'Absent' };
  }
  return { code: 'OTHER', label: raw };
}

// ------------------------------------------------------------------
// RESPONSE sheet (auto-create + append, mirrors ensureResponseSheet_/logResponse)
// ------------------------------------------------------------------
let _responseSheetChecked = false;
async function ensureResponseSheet() {
  if (_responseSheetChecked) return;
  const api = await getSheetsApi();
  const meta = await api.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = (meta.data.sheets || []).some(s => s.properties.title === SHEET_RESPONSE_NAME);
  if (!exists) {
    await api.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_RESPONSE_NAME } } }] }
    });
    await api.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_RESPONSE_NAME}!A1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[
          'Timestamp', 'Mail ID', 'Action / Button', 'Latitude', 'Longitude',
          'Location Accuracy (m)', 'Location Status', 'User Agent', 'Platform',
          'Language', 'Screen Resolution', 'Page / Screen'
        ]]
      }
    });
  }
  _responseSheetChecked = true;
}

async function logResponse(payload) {
  payload = payload || {};
  await ensureResponseSheet();
  await appendRow(SHEET_RESPONSE_NAME, [
    new Date().toISOString(),
    payload.mail || '',
    payload.action || '',
    payload.lat ?? '',
    payload.lng ?? '',
    payload.locAccuracy ?? '',
    payload.locStatus || '',
    payload.userAgent || '',
    payload.platform || '',
    payload.language || '',
    payload.screen || '',
    payload.page || ''
  ]);
}

module.exports = {
  TIMEZONE,
  normalizeMail, normalizeCode,
  findCodeForMail, findRowByCode, buildProfile,
  getDateColumns, getPlannedForCode, classifyStatus,
  logResponse
};
