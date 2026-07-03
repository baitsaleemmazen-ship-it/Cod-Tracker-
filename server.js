require('dotenv').config();
const express = require('express');
const multer = require('multer');
const session = require('express-session');
const { google } = require('googleapis');
const Anthropic = require('@anthropic-ai/sdk');
const fs = require('fs');
const path = require('path');
let cloudinary, XLSX;
try { cloudinary = require('cloudinary').v2; } catch(e) { console.error('cloudinary not installed'); }
try { XLSX = require('xlsx'); } catch(e) { console.error('xlsx not installed'); }

// ─── Cloudinary config ────────────────────────────────────────────────────────
if (cloudinary) {
  cloudinary.config({
    cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
    api_key: process.env.CLOUDINARY_API_KEY,
    api_secret: process.env.CLOUDINARY_API_SECRET
  });
}

async function uploadToCloudinary(b64, mediaType, folder, publicId) {
  if (!cloudinary) return null;
  try {
    const dataUri = `data:${mediaType};base64,${b64}`;
    const result = await cloudinary.uploader.upload(dataUri, {
      folder: `cod-receipts/${folder}`,
      public_id: publicId,
      overwrite: true,
      resource_type: 'image',
      invalidate: true,
      expires_at: Math.floor(Date.now() / 1000) + (30 * 24 * 60 * 60)
    });
    return result.secure_url;
  } catch(e) {
    console.error('Cloudinary upload error:', e.message);
    return null;
  }
}

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const uploadBoth = upload.fields([{ name: 'receipt', maxCount: 1 }, { name: 'talabat', maxCount: 1 }]);
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'cod-tracker-secret-2026',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 }
}));

// ─── Riders store (Google Sheet as source of truth) ──────────────────────────
let riders = [
  {id:"768546",name:"Waleed Hamad"},{id:"839888",name:"Umar Altaf"},{id:"879982",name:"Mohammed Imran"},
  {id:"885396",name:"Abdur Rahman"},{id:"896229",name:"Ghayoor Ahmed"},{id:"896228",name:"Imran Solman"},
  {id:"925662",name:"Yakub Hossain"},{id:"925731",name:"Abu Sufian"},{id:"938330",name:"Mohammed Shahid"},
  {id:"939163",name:"Anowar Hossain"},{id:"939171",name:"Md Alamgir Hossein"},{id:"1036871",name:"Moshiur Rahman"},
  {id:"1036793",name:"Sohid Islam"},{id:"1383612",name:"MD Mohin Uddin"},{id:"1423878",name:"Md Mahfujur"},
  {id:"1423882",name:"Muhammad Ilyas"},{id:"1397838",name:"Ali Raza"},{id:"2111340",name:"Nurul Islam"},
  {id:"2110059",name:"Tajul Islam"},{id:"2112576",name:"Muhammad Javed"},{id:"2116901",name:"Mosharaf Hossain"},
  {id:"2113747",name:"Md Kamal Hossain"},{id:"2110038",name:"Muhammad Arslan"},{id:"2114891",name:"Ismail Latif"},
  {id:"2129901",name:"Md Jamal Hussain"},{id:"2836733",name:"Naeem Shahzad"},{id:"3807922",name:"Kashif Iqbal"},
  {id:"3877440",name:"Mohammed Nadeem"},{id:"1383608",name:"Amer Shazad"},{id:"1417159",name:"Shazad Saleem"},
  {id:"1629806",name:"Mohammed Wajid"},{id:"925713",name:"Mohammed Ali"},{id:"4516509",name:"Omar Gaber"},
  {id:"4478518",name:"Ahmed Al Alwai"},{id:"4520599",name:"Badar Almukhaini"},{id:"4559077",name:"Sabir Albalushi"},
  {id:"4559994",name:"Sarmad Said"},{id:"4594386",name:"Sujit Dhirendra"},{id:"4564834",name:"Feysal Alaamri"},
  {id:"4572472",name:"Mohammed Shahid 2"},{id:"4607634",name:"Faysal Ahmed"},{id:"4602593",name:"Hamid Jadad"},
  {id:"4611922",name:"Salah Al Hajri"},{id:"1421555",name:"Shaikh Habib"}
];

// Load riders from Google Sheet (Riders tab) — overwrites default list if exists
async function loadRidersFromSheet() {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const existing = meta.data.sheets.map(s => s.properties.title);
    if (!existing.includes('Riders')) {
      // Create Riders tab and populate with default list
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: 'Riders' } } }] }
      });
      const rows = [['Name', 'ID'], ...riders.map(r => [r.name, r.id])];
      await sheets.spreadsheets.values.update({
        spreadsheetId, range: 'Riders!A1',
        valueInputOption: 'RAW',
        requestBody: { values: rows }
      });
      console.log('Created Riders tab with default list');
    } else {
      // Load from sheet
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Riders!A:B' });
      const rows = res.data.values || [];
      const loaded = rows.slice(1).filter(r => r[0] && r[1]).map(r => ({ name: r[0].trim(), id: r[1].trim() }));
      if (loaded.length > 0) { riders = loaded; console.log(`Loaded ${riders.length} riders from sheet`); }
    }
  } catch (e) { console.error('loadRidersFromSheet error:', e.message); }
}

async function saveRidersToSheet() {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const rows = [['Name', 'ID'], ...riders.map(r => [r.name, r.id])];
    await sheets.spreadsheets.values.update({
      spreadsheetId, range: 'Riders!A1',
      valueInputOption: 'RAW',
      requestBody: { values: rows }
    });
  } catch (e) { console.error('saveRidersToSheet error:', e.message); }
}

// ─── Submissions store ────────────────────────────────────────────────────────
const SUBS_FILE = path.join(__dirname, 'submissions.json');
let submissions = [];
if (fs.existsSync(SUBS_FILE)) {
  try { submissions = JSON.parse(fs.readFileSync(SUBS_FILE)); } catch(e) { submissions = []; }
}
function saveSubmissions() {
  try {
    // Strip any base64 image data before saving to keep file small
    const clean = submissions.map(s => {
      const { image_b64, image_type, talabat_b64, talabat_type, ...rest } = s;
      return rest;
    });
    fs.writeFileSync(SUBS_FILE, JSON.stringify(clean, null, 2));
  } catch(e) { console.error('saveSubmissions error:', e.message); }
}

// Load today's submissions from Google Sheet (Submissions tab) on startup
async function loadTodaySubmissionsFromSheet() {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const existing = meta.data.sheets.map(s => s.properties.title);
    if (!existing.includes('Submissions')) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: 'Submissions' } } }] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId, range: 'Submissions!A1',
        valueInputOption: 'RAW',
        requestBody: { values: [['id','rider_id','rider_name','amount','talabat_amount','bank','date','submitted_at','status','flags','is_late','account_name','beneficiary_name','needs_amount','bank_url','talabat_url','talabat_deliveries']] }
      });
      return;
    }
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Submissions!A:Q' });
    const rows = (res.data.values || []).slice(1);
    // Load ALL submissions from sheet (not just today)
    const fromSheet = rows
      .filter(r => r[0] && r[6]) // must have id and date
      .map(r => ({
        id: r[0], rider_id: r[1], rider_name: r[2],
        amount: r[3] ? parseFloat(r[3]) : null,
        talabat_amount: r[4] ? parseFloat(r[4]) : null,
        bank: r[5] || null, date: r[6], submitted_at: r[7], status: r[8],
        flags: r[9] ? (() => { try { return JSON.parse(r[9]); } catch(e) { return []; } })() : [],
        is_late: r[10] === 'true',
        account_name: r[11] || null, beneficiary_name: r[12] || null,
        needs_amount: r[13] === 'true',
        bank_url: r[14] || null,
        talabat_url: r[15] || null,
        talabat_deliveries: r[16] ? parseInt(r[16]) : null
      }));
    // Merge — keep existing (with images in memory) and add any missing from sheet
    const existingIds = new Set(submissions.map(s => s.id));
    let added = 0;
    fromSheet.forEach(s => { if (!existingIds.has(s.id)) { submissions.push(s); added++; } });
    // Also update bank_url/talabat_url for existing submissions that may be missing them
    fromSheet.forEach(s => {
      const existing = submissions.find(ex => ex.id === s.id);
      if (existing) {
        if (!existing.bank_url && s.bank_url) existing.bank_url = s.bank_url;
        if (!existing.talabat_url && s.talabat_url) existing.talabat_url = s.talabat_url;
        if (!existing.talabat_deliveries && s.talabat_deliveries) existing.talabat_deliveries = s.talabat_deliveries;
      }
    });
    console.log(`Loaded ${fromSheet.length} submissions from sheet (${added} new)`);
  } catch(e) { console.error('loadSubmissionsFromSheet error:', e.message); }
}

async function saveSubmissionToSheet(sub) {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    await sheets.spreadsheets.values.append({
      spreadsheetId, range: 'Submissions!A:Q', valueInputOption: 'RAW',
      requestBody: { values: [[
        sub.id, sub.rider_id, sub.rider_name,
        sub.amount || '', sub.talabat_amount || '', sub.bank || '',
        sub.date, sub.submitted_at, sub.status,
        JSON.stringify(sub.flags || []),
        sub.is_late ? 'true' : 'false',
        sub.account_name || '', sub.beneficiary_name || '',
        sub.needs_amount ? 'true' : 'false',
        sub.bank_url || '', sub.talabat_url || '',
        sub.talabat_deliveries || ''
      ]]}
    });
  } catch(e) { console.error('saveSubmissionToSheet error:', e.message); }
}

async function updateSubmissionInSheet(sub) {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Submissions!A:A' });
    const ids = (res.data.values || []).map(r => r[0]);
    const rowIndex = ids.indexOf(sub.id);
    if (rowIndex === -1) return;
    const rowNum = rowIndex + 1;
    await sheets.spreadsheets.values.update({
      spreadsheetId, range: `Submissions!A${rowNum}:Q${rowNum}`, valueInputOption: 'RAW',
      requestBody: { values: [[
        sub.id, sub.rider_id, sub.rider_name,
        sub.amount || '', sub.talabat_amount || '', sub.bank || '',
        sub.date, sub.submitted_at, sub.status,
        JSON.stringify(sub.flags || []),
        sub.is_late ? 'true' : 'false',
        sub.account_name || '', sub.beneficiary_name || '',
        sub.needs_amount ? 'true' : 'false',
        sub.bank_url || '', sub.talabat_url || '',
        sub.talabat_deliveries || ''
      ]]}
    });
  } catch(e) { console.error('updateSubmissionInSheet error:', e.message); }
}

// ─── In-memory image store ────────────────────────────────────────────────────
const imageStore = {};


function getTodayIds() {
  const today = new Date(new Date().getTime() + 4*60*60*1000).toISOString().slice(0, 10);
  return new Set(
    submissions
      .filter(s => s.date === today && s.status === 'approved')
      .map(s => s.rider_id)
  );
}

// ─── Google Sheets helper ─────────────────────────────────────────────────────
async function getSheetAuth() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  return { auth, sheets: google.sheets({ version: 'v4', auth }), spreadsheetId: process.env.SHEET_ID };
}

async function ensureDailySheet(date) {
  const parts = date.split('-'); // YYYY-MM-DD
  const sheetName = `${parts[2]}-${parts[1]}-${parts[0]}`; // DD-MM-YYYY
  const displayDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheets = meta.data.sheets;
    const existing = existingSheets.map(s => s.properties.title);

    if (!existing.includes(sheetName)) {
      // Create new sheet tab
      const addRes = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] }
      });
      const sheetId = addRes.data.replies[0].addSheet.properties.sheetId;

      // Write data: row1=company, row2=date, row3=headers, row4+=riders
      const headers = ['#', 'Rider Name', 'Rider ID', 'Orders', 'Talabat Collected (OMR)', 'Bank Amount (OMR)', 'Bank', 'Submitted At', 'Status'];
      const riderRows = riders.map((r, i) => [i + 1, r.name, r.id, '', '', '', '', '', 'Not Submitted']);
      const COLS = 9;
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A1`,
        valueInputOption: 'RAW',
        requestBody: { values: [
          ['Future Wave', '', '', '', '', '', '', '', ''],
          [displayDate, '', '', '', '', '', '', '', ''],
          ...([headers]),
          ...riderRows
        ]}
      });

      const totalRows = riders.length + 3;

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [
          // Merge company name row A1:H1
          { mergeCells: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: COLS }, mergeType: 'MERGE_ALL' } },
          // Merge date row A2:H2
          { mergeCells: { range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: COLS }, mergeType: 'MERGE_ALL' } },

          // Company name style — dark background, white bold large text, centered
          { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: COLS },
            cell: { userEnteredFormat: {
              backgroundColor: { red: 0.067, green: 0.31, blue: 0.165 },
              horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
              textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 16 }
            }}, fields: 'userEnteredFormat' }},

          // Date row style — light green background, dark text, centered
          { repeatCell: { range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: COLS },
            cell: { userEnteredFormat: {
              backgroundColor: { red: 0.851, green: 0.918, blue: 0.863 },
              horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
              textFormat: { foregroundColor: { red: 0.067, green: 0.31, blue: 0.165 }, bold: true, fontSize: 12 }
            }}, fields: 'userEnteredFormat' }},

          // Header row style — dark green background, white bold text, centered
          { repeatCell: { range: { sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: COLS },
            cell: { userEnteredFormat: {
              backgroundColor: { red: 0.067, green: 0.31, blue: 0.165 },
              horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
              textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 11 }
            }}, fields: 'userEnteredFormat' }},

          // All rider rows — default white background, centered text
          { repeatCell: { range: { sheetId, startRowIndex: 3, endRowIndex: totalRows, startColumnIndex: 0, endColumnIndex: COLS },
            cell: { userEnteredFormat: {
              backgroundColor: { red: 1, green: 1, blue: 1 },
              horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
              textFormat: { fontSize: 10 }
            }}, fields: 'userEnteredFormat' }},

          // Status column (H) — red for "Not Submitted"
          { repeatCell: { range: { sheetId, startRowIndex: 3, endRowIndex: totalRows, startColumnIndex: 7, endColumnIndex: 8 },
            cell: { userEnteredFormat: {
              backgroundColor: { red: 0.957, green: 0.8, blue: 0.8 },
              horizontalAlignment: 'CENTER',
              textFormat: { foregroundColor: { red: 0.6, green: 0.0, blue: 0.0 }, bold: true, fontSize: 10 }
            }}, fields: 'userEnteredFormat' }},

          // Alternating row colors for rider rows
          ...riders.map((_, i) => ({
            repeatCell: {
              range: { sheetId, startRowIndex: 3 + i, endRowIndex: 4 + i, startColumnIndex: 0, endColumnIndex: 7 },
              cell: { userEnteredFormat: {
                backgroundColor: i % 2 === 0
                  ? { red: 1, green: 1, blue: 1 }
                  : { red: 0.949, green: 0.949, blue: 0.949 }
              }}, fields: 'userEnteredFormat(backgroundColor)'
            }
          })),

          // Row heights
          { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 45 }, fields: 'pixelSize' } },
          { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } },
          { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } },
          { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 3, endIndex: totalRows }, properties: { pixelSize: 25 }, fields: 'pixelSize' } },

          // Column widths
          { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } },   // #
          { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 180 }, fields: 'pixelSize' } },  // Name
          { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 100 }, fields: 'pixelSize' } },  // ID
          { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 150 }, fields: 'pixelSize' } },  // Bank Amount
          { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 160 }, fields: 'pixelSize' } },  // Talabat
          { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },  // Bank
          { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },  // Time
          { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },  // Status

          // Borders for data area
          { updateBorders: { range: { sheetId, startRowIndex: 2, endRowIndex: totalRows, startColumnIndex: 0, endColumnIndex: COLS },
            top: { style: 'SOLID', width: 2, color: { red: 0.067, green: 0.31, blue: 0.165 } },
            bottom: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } },
            left: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } },
            right: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } },
            innerHorizontal: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } },
            innerVertical: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } }
          }}
        ]}
      });
    }
    return sheetName;
  } catch (e) {
    console.error('ensureDailySheet error:', e.message);
    return null;
  }
}

async function fillRiderRow(submission) {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const sheetName = await ensureDailySheet(submission.date);
    if (!sheetName) return false;

    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetObj = meta.data.sheets.find(s => s.properties.title === sheetName);
    const sheetId = sheetObj ? sheetObj.properties.sheetId : null;

    // Get all rows to find rider's row
    // Sheet structure: Row1=Company, Row2=Date, Row3=Headers, Row4+=Riders
    // Columns: A=#, B=Name, C=ID, D=Bank Amount, E=Talabat, F=Bank, G=Time, H=Status
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!C:C` });
    const idCol = res.data.values || [];

    const isLate = submission.is_late || false;
    const statusText = isLate ? 'Late' : 'Submitted';

    // Find rider row by ID in column C — data starts at row index 3 (row 4 in sheet)
    let rowIndex = -1;
    for (let i = 3; i < idCol.length; i++) {
      if (idCol[i] && idCol[i][0] && idCol[i][0].toString().trim() === submission.rider_id.toString().trim()) {
        rowIndex = i + 1; // 1-indexed for Sheets API
        break;
      }
    }

    if (rowIndex === -1) {
      // Rider not found — append at end
      await sheets.spreadsheets.values.append({
        spreadsheetId, range: `${sheetName}!A:I`, valueInputOption: 'RAW',
        requestBody: { values: [['', submission.rider_name, submission.rider_id, submission.talabat_deliveries || '', submission.talabat_amount || '', submission.amount || '', submission.bank || '', new Date(new Date(submission.submitted_at).getTime()+4*60*60*1000).toLocaleTimeString("en-US",{hour:"2-digit",minute:"2-digit",hour12:true}), statusText]] }
      });
    } else {
      // Fill rider's existing row — columns D to I (Bank Amount, Talabat, Orders, Bank, Time, Status)
      await sheets.spreadsheets.values.update({
        spreadsheetId, range: `${sheetName}!D${rowIndex}:I${rowIndex}`, valueInputOption: 'RAW',
        requestBody: { values: [[submission.talabat_deliveries || '', submission.talabat_amount || '', submission.amount || '', submission.bank || '', new Date(new Date(submission.submitted_at).getTime()+4*60*60*1000).toLocaleTimeString("en-US",{hour:"2-digit",minute:"2-digit",hour12:true}), statusText]] }
      });

      // Color the row — green if on time, yellow if late
      if (sheetId !== null) {
        const rowIdx = rowIndex - 1; // 0-indexed
        const bgColor = isLate
          ? { red: 1.0, green: 0.953, blue: 0.714 }   // yellow
          : { red: 0.714, green: 0.918, blue: 0.757 }; // green
        const textColor = isLate
          ? { red: 0.4, green: 0.3, blue: 0.0 }
          : { red: 0.067, green: 0.31, blue: 0.165 };
        const statusBg = isLate
          ? { red: 0.98, green: 0.85, blue: 0.45 }
          : { red: 0.2, green: 0.659, blue: 0.322 };

        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests: [
            // Color columns A-G (data columns)
            { repeatCell: { range: { sheetId, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 0, endColumnIndex: 7 },
              cell: { userEnteredFormat: { backgroundColor: bgColor, textFormat: { foregroundColor: textColor, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
              fields: 'userEnteredFormat' }},
            // Status column H
            { repeatCell: { range: { sheetId, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 7, endColumnIndex: 8 },
              cell: { userEnteredFormat: {
                backgroundColor: statusBg,
                horizontalAlignment: 'CENTER',
                textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 10 }
              }}, fields: 'userEnteredFormat' }}
          ]}
        });
      }
    }
    return true;
  } catch (e) {
    console.error('fillRiderRow error:', e.message);
    return false;
  }
}

// ─── Auth middleware ──────────────────────────────────────────────────────────
function requireAdmin(req, res, next) {
  if (req.session && req.session.admin) return next();
  res.redirect('/admin/login');
}

// ══════════════════════════════════════════════════════════════════════════════
// RIDER APP — public
// ══════════════════════════════════════════════════════════════════════════════
app.get('/', (req, res) => {
  res.send(`<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>Future Wave — Rider Portal</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f2f5;min-height:100vh;}
.header{background:#0d4d25;padding:16px 20px;display:flex;align-items:center;gap:10px;}
.header h1{color:#fff;font-size:17px;font-weight:600;}
.header .sub{color:rgba(255,255,255,0.7);font-size:12px;}
.tabs{display:flex;background:#fff;border-bottom:1px solid #eee;position:sticky;top:0;z-index:10;}
.tab{flex:1;padding:13px 8px;text-align:center;font-size:13px;font-weight:500;color:#888;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.2s;}
.tab.active{color:#0d4d25;border-bottom-color:#0d4d25;font-weight:600;}
.tab-icon{font-size:16px;display:block;margin-bottom:2px;}
.pane{display:none;padding:16px;max-width:480px;margin:0 auto;}
.pane.active{display:block;}
.card{background:#fff;border-radius:14px;padding:16px;margin-bottom:12px;box-shadow:0 1px 4px rgba(0,0,0,0.06);}
.card-title{font-size:14px;font-weight:600;margin-bottom:12px;color:#222;}
input[type=number],input[type=text]{width:100%;padding:12px;border:1.5px solid #ddd;border-radius:10px;font-size:16px;background:#fafafa;outline:none;}
input:focus{border-color:#0d4d25;}
.btn{width:100%;padding:13px;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;margin-top:10px;}
.btn-primary{background:#0d4d25;color:#fff;}
.btn-secondary{background:#e8f5e9;color:#0d4d25;}
.btn:disabled{background:#ccc;color:#fff;}
.btn-row{display:flex;gap:8px;margin-top:0;}
.btn-half{flex:1;padding:12px;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;border:none;}
.btn-cam{background:#0d4d25;color:#fff;}
.btn-gal{background:#fff;color:#0d4d25;border:2px solid #0d4d25;}
.preview{display:none;margin-top:8px;border-radius:10px;overflow:hidden;border:1px solid #eee;}
.preview img{width:100%;max-height:160px;object-fit:cover;display:block;}
.preview-name{font-size:11px;color:#1e7e34;padding:6px 10px;background:#f5f5f5;}
.msg{padding:12px;border-radius:10px;font-size:13px;text-align:center;margin-top:10px;}
.msg.ok{background:#e6f4ea;color:#1e7e34;}
.msg.err{background:#fce8e6;color:#c62828;}
.msg.warn{background:#fff3e0;color:#e65100;}
.spinner{text-align:center;padding:14px;color:#888;font-size:13px;display:none;}
/* Status styles */
.status-card{border-radius:14px;padding:20px;text-align:center;margin-top:8px;}
.status-card.approved{background:#e6f4ea;border:1px solid #b7dfbe;}
.status-card.flagged{background:#fff3e0;border:1px solid #ffe0b2;}
.status-card.rejected{background:#fce8e6;border:1px solid #f5c6c2;}
.status-card.none{background:#f5f5f5;border:1px solid #eee;}
.status-icon{font-size:44px;margin-bottom:8px;}
.status-label{font-size:18px;font-weight:700;margin-bottom:4px;}
.status-sub{font-size:13px;color:#666;line-height:1.5;}
.status-amount{font-size:26px;font-weight:700;color:#1e7e34;margin-top:10px;}
.status-time{font-size:11px;color:#aaa;margin-top:6px;}
/* Fuel styles */
.week-row{display:flex;justify-content:space-between;align-items:center;padding:12px 0;border-bottom:1px solid #f0f0f0;}
.week-row:last-child{border-bottom:none;}
.week-label{font-size:13px;color:#555;}
.week-amt{font-size:17px;font-weight:700;color:#0d4d25;}
.week-amt.zero{color:#aaa;font-size:14px;font-weight:400;}
.total-row{background:#0d4d25;border-radius:12px;padding:14px;display:flex;justify-content:space-between;align-items:center;margin-top:10px;}
.total-label{color:rgba(255,255,255,0.8);font-size:13px;}
.total-amt{color:#fff;font-size:22px;font-weight:700;}
.rider-name-badge{background:#e8f5e9;color:#0d4d25;border-radius:8px;padding:8px 12px;font-size:14px;font-weight:600;margin-bottom:12px;display:inline-block;}
</style></head><body>

<div class="header">
  <div>
    <h1>🚚 Future Wave</h1>
    <div class="sub">Rider Portal</div>
  </div>
</div>

<div class="tabs">
  <div class="tab active" onclick="switchTab('submit')"><span class="tab-icon">📦</span>Submit</div>
  <div class="tab" onclick="switchTab('status')"><span class="tab-icon">📋</span>Status</div>
  <div class="tab" onclick="switchTab('fuel')"><span class="tab-icon">⛽</span>Fuel</div>
</div>

<!-- SUBMIT TAB -->
<div class="pane active" id="pane-submit">
  <div class="card">
    <div class="card-title">Your ID</div>
    <input type="number" id="riderSel" placeholder="Enter your rider ID" inputmode="numeric">
  </div>
  <div class="card">
    <div class="card-title">🏦 Bank Receipt</div>
    <input type="file" id="bankInput" accept="image/*" onchange="previewFile(this,'bankPreview','bankImg')" style="display:none">
    <input type="file" id="bankCamera" accept="image/*" capture="environment" onchange="previewFile(this,'bankPreview','bankImg')" style="display:none">
    <div class="btn-row">
      <button type="button" class="btn-half btn-cam" onclick="document.getElementById('bankCamera').click()">📷 Camera</button>
      <button type="button" class="btn-half btn-gal" onclick="document.getElementById('bankInput').click()">🖼 Gallery</button>
    </div>
    <div class="preview" id="bankPreview"><img id="bankImg" alt=""><div class="preview-name" id="bankName"></div></div>
  </div>
  <div class="card">
    <div class="card-title">🛵 Talabat Screenshot</div>
    <input type="file" id="talabatInput" accept="image/*" onchange="previewFile(this,'talabatPreview','talabatImg')" style="display:none">
    <input type="file" id="talabatCamera" accept="image/*" capture="environment" onchange="previewFile(this,'talabatPreview','talabatImg')" style="display:none">
    <div class="btn-row">
      <button type="button" class="btn-half btn-cam" onclick="document.getElementById('talabatCamera').click()">📷 Camera</button>
      <button type="button" class="btn-half btn-gal" onclick="document.getElementById('talabatInput').click()">🖼 Gallery</button>
    </div>
    <div class="preview" id="talabatPreview"><img id="talabatImg" alt=""><div class="preview-name" id="talabatName"></div></div>
  </div>
  <button class="btn btn-primary" id="submitBtn" onclick="doSubmit()">Submit Receipt</button>
  <div class="spinner" id="submitSpinner">⏳ Reading your receipts, please wait...</div>
  <div id="submitMsg"></div>
</div>

<!-- STATUS TAB -->
<div class="pane" id="pane-status">
  <div class="card">
    <div class="card-title">Check today's submission</div>
    <input type="number" id="statusId" placeholder="Enter your rider ID" inputmode="numeric">
    <button class="btn btn-primary" onclick="doStatus()">Check status</button>
  </div>
  <div class="spinner" id="statusSpinner">Checking...</div>
  <div id="statusResult"></div>
</div>

<!-- FUEL TAB -->
<div class="pane" id="pane-fuel">
  <div class="card">
    <div class="card-title">Check your fuel allowance</div>
    <input type="number" id="fuelId" placeholder="Enter your rider ID" inputmode="numeric">
    <button class="btn btn-primary" onclick="doFuel()">Check fuel</button>
  </div>
  <div class="spinner" id="fuelSpinner">Loading...</div>
  <div id="fuelResult"></div>
</div>

<script>
function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('active', ['submit','status','fuel'][i]===name));
  document.querySelectorAll('.pane').forEach(p => p.classList.remove('active'));
  document.getElementById('pane-'+name).classList.add('active');
}

function previewFile(input, wrapId, imgId) {
  const file = input.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    document.getElementById(imgId).src = e.target.result;
    document.getElementById(wrapId).style.display = 'block';
    const nameEl = document.getElementById(wrapId.replace('Preview','Name'));
    if (nameEl) nameEl.textContent = '✓ ' + file.name;
  };
  reader.readAsDataURL(file);
}

async function doSubmit() {
  const riderId = document.getElementById('riderSel').value.trim();
  const bankFile = document.getElementById('bankInput').files[0] || document.getElementById('bankCamera').files[0];
  const talabatFile = document.getElementById('talabatInput').files[0] || document.getElementById('talabatCamera').files[0];
  const msgEl = document.getElementById('submitMsg');
  msgEl.innerHTML = '';
  if (!riderId) { msgEl.innerHTML = '<div class="msg err">Please enter your ID number.</div>'; return; }
  if (!bankFile) { msgEl.innerHTML = '<div class="msg err">Please upload your bank receipt.</div>'; return; }
  if (!talabatFile) { msgEl.innerHTML = '<div class="msg err">Please upload your Talabat screenshot.</div>'; return; }
  document.getElementById('submitBtn').disabled = true;
  document.getElementById('submitSpinner').style.display = 'block';
  try {
    const chk = await fetch('/check-submitted?id=' + encodeURIComponent(riderId));
    const chkData = await chk.json();
    if (chkData.submitted) {
      document.getElementById('submitSpinner').style.display = 'none';
      document.getElementById('submitBtn').disabled = false;
      msgEl.innerHTML = '<div class="msg warn">⚠️ You already submitted today. Check the Status tab to see your result.</div>';
      return;
    }
  } catch(e) {}
  const fd = new FormData();
  fd.append('rider_id', riderId);
  fd.append('receipt', bankFile);
  fd.append('talabat', talabatFile);
  try {
    const res = await fetch('/submit', { method: 'POST', body: fd });
    const data = await res.json();
    document.getElementById('submitSpinner').style.display = 'none';
    if (data.ok) {
      document.getElementById('submitBtn').style.display = 'none';
      document.querySelectorAll('.card').forEach(c => c.style.display = 'none');
      msgEl.innerHTML = '<div class="msg ok" style="padding:20px;font-size:15px;">✅ Submitted successfully!<br><span style="font-size:13px;color:#555;margin-top:4px;display:block;">Check the Status tab to track your submission.</span></div>';
    } else {
      msgEl.innerHTML = '<div class="msg err">❌ ' + (data.error || 'Error. Please try again.') + '</div>';
      document.getElementById('submitBtn').disabled = false;
    }
  } catch(e) {
    document.getElementById('submitSpinner').style.display = 'none';
    msgEl.innerHTML = '<div class="msg err">❌ Network error. Please try again.</div>';
    document.getElementById('submitBtn').disabled = false;
  }
}

async function doStatus() {
  const id = document.getElementById('statusId').value.trim();
  if (!id) return;
  document.getElementById('statusSpinner').style.display = 'block';
  document.getElementById('statusResult').innerHTML = '';
  const res = await fetch('/status/check?id=' + encodeURIComponent(id));
  const data = await res.json();
  document.getElementById('statusSpinner').style.display = 'none';
  const el = document.getElementById('statusResult');
  if (!data.found) {
    el.innerHTML = '<div class="status-card none"><div class="status-icon">📭</div><div class="status-label">No submission today</div><div class="status-sub">You have not submitted yet today.</div></div>';
    return;
  }
  const map = {
    approved: {icon:'✅', label:'Approved', sub:'Your submission has been approved.', cls:'approved'},
    flagged:  {icon:'⏳', label:'Under Review', sub:'Being reviewed by supervisor.', cls:'flagged'},
    rejected: {icon:'❌', label:'Rejected', sub:'Your submission was rejected.' + (data.reason ? '<br><b>Reason: '+data.reason+'</b>' : ''), cls:'rejected'}
  };
  const m = map[data.status] || {icon:'📄', label:data.status, sub:'', cls:'none'};
  el.innerHTML = '<div class="status-card '+m.cls+'"><div class="status-icon">'+m.icon+'</div><div class="status-label">'+m.label+'</div><div class="status-sub">'+m.sub+'</div>'+(data.amount?'<div class="status-amount">'+data.amount+' OMR</div>':'')+'<div class="status-time">Submitted at '+data.time+'</div></div>';
}

async function doFuel() {
  const id = document.getElementById('fuelId').value.trim();
  if (!id) return;
  document.getElementById('fuelSpinner').style.display = 'block';
  document.getElementById('fuelResult').innerHTML = '';
  const res = await fetch('/fuel/check?id=' + encodeURIComponent(id));
  const data = await res.json();
  document.getElementById('fuelSpinner').style.display = 'none';
  const el = document.getElementById('fuelResult');
  if (!data.ok) {
    el.innerHTML = '<div class="msg err">'+data.error+'</div>'; return;
  }
  const total = data.weeks.reduce((s,w) => s+(parseFloat(w.amount)||0), 0);
  const rows = data.weeks.map(w => '<div class="week-row"><div class="week-label">📅 '+w.week+'</div><div class="week-amt '+(parseFloat(w.amount)>0?'':'zero')+'">'+( parseFloat(w.amount)>0 ? w.amount+' OMR' : '—')+'</div></div>').join('');
  el.innerHTML = '<div class="card"><div class="rider-name-badge">👤 '+data.rider_name+'</div>'+rows+(data.weeks.length>1?'<div class="total-row"><div class="total-label">Total</div><div class="total-amt">'+total.toFixed(3)+' OMR</div></div>':'')+'</div>';
}

// Auto-fill status/fuel ID from submit tab
document.getElementById('riderSel').addEventListener('input', e => {
  document.getElementById('statusId').value = e.target.value;
  document.getElementById('fuelId').value = e.target.value;
});
</script>
</body></html>`);
});


// ─── Submit endpoint ──────────────────────────────────────────────────────────
app.post('/submit', uploadBoth, async (req, res) => {
  try {
    const { rider_id } = req.body;
    const rider = riders.find(r => r.id === rider_id);
    // If ID not found, create a temporary rider object and flag it
    const riderObj = rider || { id: rider_id, name: `Unknown ID: ${rider_id}` };
    if (!rider_id) return res.json({ ok: false, error: 'Please enter your ID number.' });
    const bankFile = req.files && req.files['receipt'] && req.files['receipt'][0];
    const talabatFile = req.files && req.files['talabat'] && req.files['talabat'][0];
    if (!bankFile) return res.json({ ok: false, error: 'No bank receipt uploaded.' });
    if (!talabatFile) return res.json({ ok: false, error: 'No Talabat screenshot uploaded.' });

    const today = new Date(new Date().getTime() + 4*60*60*1000).toISOString().slice(0, 10);
    const todayIds = getTodayIds();
    const isDuplicate = todayIds.has(rider_id);

    // Flag if ID not in rider list
    if (!rider) {
      flags.push(`ID ${rider_id} is not in the rider list`);
      status = 'flagged';
    }

    // AI analysis — bank receipt
    const bankB64 = bankFile.buffer.toString('base64');
    const bankMediaType = bankFile.mimetype || 'image/jpeg';
    let aiResult = {};
    try {
      const response = await anthropic.messages.create({
        model: 'claude-sonnet-4-5',
        max_tokens: 500,
        messages: [{
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: bankMediaType, data: bankB64 } },
            { type: 'text', text: `This is a bank transfer receipt from Oman, either in English or Arabic. It may have a decorative background pattern — ignore it and read only the text fields. Arabic numerals (٠١٢٣٤٥٦٧٨٩) must be converted to Western numerals (0123456789).

Extract these fields (check both English and Arabic labels):

- Amount: look for "Amount" or "المبلغ" or "الإجمالي" → extract number only, convert Arabic numerals (e.g. ١١٫٢٢٠ = 11.220)
- Sender name: "Debit Account Name" or "اسم الحساب المدين" or "اسم المرسل"
- Beneficiary name: "Beneficiary Name" or "اسم المستفيد" → who received the money
- Rider ID: in "Remarks" or "الملاحظات" or "البيان" → look for "id XXXXXXX" (6-7 digit number after "id "), convert Arabic numerals if needed
- Date: "Transaction Date" or "تاريخ العملية" or "التاريخ", convert Arabic numerals if needed
- Bank: "Beneficiary Bank" or "بنك المستفيد"
- Transaction status: look for "Transaction Completed", "Processed Successfully", "تمت العملية بنجاح", "ناجح", "مكتمل" = successful. "Failed", "فشل", "Pending", "معلق", "Rejected", "مرفوض" = NOT successful.

Return ONLY this JSON, no markdown, no explanation:
{"amount":<number in Western numerals or null>,"currency":"OMR","date":"<YYYY-MM-DD or null>","detected_id":"<6-7 digit id in Western numerals or not_found>","account_name":"<sender full name or null>","beneficiary_name":"<beneficiary full name or null>","bank_name":"<bank name or null>","is_legit_receipt":true,"transaction_successful":<true or false>}` }
          ]
        }]
      });
      const raw = response.content.map(c => c.text || '').join('').replace(/```json|```/g, '').trim();
      console.log('AI raw response:', raw);
      aiResult = JSON.parse(raw);
      console.log('AI parsed:', JSON.stringify(aiResult));
    } catch (e) {
      console.error('AI error:', e.message);
    }

    // AI analysis — Talabat screenshot
    const talabatB64 = talabatFile.buffer.toString('base64');
    const talabatMediaType = talabatFile.mimetype || 'image/jpeg';
    let talabatResult = {};
    try {
      const tRes = await anthropic.messages.create({
        model: 'claude-sonnet-4-5',
        max_tokens: 300,
        messages: [{
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: talabatMediaType, data: talabatB64 } },
            { type: 'text', text: `This is a Talabat delivery app screenshot showing a rider's delivery history. It may be in English or Arabic. Arabic numerals (٠١٢٣٤٥٦٧٨٩) must be converted to Western numerals (0123456789).

Extract:
- Collected cash amount — look for "Collected", "المحصّل", "النقد المحصّل", "إجمالي النقد", or any OMR/ر.ع amount shown as total cash collected. Convert Arabic numerals if needed (٥٥٫٩٠ = 55.90)
- Number of deliveries — "Deliveries", "التوصيلات", "الطلبات", or a count number
- Date shown — may be in Arabic format (١٧ مايو ٢٠٢٦ = 2026-05-17)

Return ONLY this JSON, no markdown:
{"collected_amount":<number in Western numerals e.g. 55.90 or null>,"deliveries":<number or null>,"date":"<YYYY-MM-DD or null>"}` }
          ]
        }]
      });
      const tRaw = tRes.content.map(c => c.text || '').join('').replace(/```json|```/g, '').trim();
      console.log('Talabat AI response:', tRaw);
      talabatResult = JSON.parse(tRaw);
    } catch (e) {
      console.error('Talabat AI error:', e.message);
    }

    // Fraud checks — only flag for real fraud, not AI reading issues
    const flags = [];
    let status = 'approved';

    // Transaction success check
    if (aiResult.transaction_successful === false) {
      flags.push('⚠️ Transaction NOT completed — receipt may show a failed or pending transfer');
      status = 'flagged';
    }

    // Beneficiary check — must be FUTURE WAVE TECHNOLOGIES
    if (aiResult.beneficiary_name) {
      const ben = aiResult.beneficiary_name.toUpperCase().trim();
      if (!ben.includes('FUTURE WAVE')) {
        flags.push(`Beneficiary is "${aiResult.beneficiary_name}" — not Future Wave Technologies`);
        status = 'flagged';
      }
    } else {
      flags.push('Beneficiary name not detected — please verify receipt');
      status = 'flagged';
    }

    // Name mismatch: account name on receipt doesn't match rider name
    if (aiResult.account_name) {
      const accountName = aiResult.account_name.toLowerCase().replace(/\s+/g, ' ').trim();
      const riderName = rider.name.toLowerCase().replace(/\s+/g, ' ').trim();
      // Check if any word of rider name appears in account name
      const riderWords = riderName.split(' ').filter(w => w.length > 2);
      const nameMatch = riderWords.some(word => accountName.includes(word));
      if (!nameMatch) {
        flags.push(`Account name on receipt "${aiResult.account_name}" doesn't match rider name "${rider.name}"`);
        status = 'flagged';
      }
    }

    // Real fraud: ID on receipt doesn't match the rider who submitted
    if (aiResult.detected_id && aiResult.detected_id !== 'not_found' && aiResult.detected_id !== rider_id) {
      flags.push(`ID on receipt (${aiResult.detected_id}) doesn't match rider ID (${rider_id})`);
      status = 'flagged';
    }

    // Real fraud: same rider already submitted today
    if (isDuplicate) {
      flags.push('This rider already submitted today — possible duplicate');
      status = 'flagged';
    }

    // Discrepancy check — bank amount vs Talabat collected differ by more than 1 OMR
    if (aiResult.amount && talabatResult.collected_amount) {
      const diff = Math.abs(aiResult.amount - talabatResult.collected_amount);
      if (diff > 1) {
        flags.push(`⚠️ Amount mismatch: Bank ${aiResult.amount} OMR vs Talabat ${talabatResult.collected_amount} OMR (difference: ${diff.toFixed(3)} OMR)`);
        status = 'flagged';
      }
    }

    // Duplicate receipt detection — check if same bank amount + same date already submitted today
    if (aiResult.amount) {
      const todayDuplicateReceipt = submissions.find(s =>
        s.date === today &&
        s.rider_id !== rider_id &&
        s.amount === aiResult.amount &&
        s.bank === (aiResult.bank_name || 'Bank')
      );
      if (todayDuplicateReceipt) {
        flags.push(`⚠️ Same amount (${aiResult.amount} OMR) already submitted by ${todayDuplicateReceipt.rider_name} today — possible duplicate receipt`);
        status = 'flagged';
      }
    }

    // Soft flag only (still approved): amount not detected — admin enters manually
    const needsAmount = !aiResult.amount;

    // Late submission check (after 8 PM GMT+4)
    const gmt4Hour = new Date(new Date().getTime() + 4*60*60*1000).getHours();
    const isLate = gmt4Hour >= 18;

    const submission = {
      id: Date.now().toString(),
      rider_id,
      rider_name: riderObj.name,
      amount: aiResult.amount || null,
      currency: aiResult.currency || 'OMR',
      bank: aiResult.bank_name || 'Bank',
      date: today,
      submitted_at: new Date().toISOString(),
      status,
      flags,
      needs_amount: needsAmount,
      talabat_amount: talabatResult.collected_amount || null,
      talabat_deliveries: talabatResult.deliveries || null,
      account_name: aiResult.account_name || null,
      beneficiary_name: aiResult.beneficiary_name || null,
      is_late: isLate,
      bank_url: null,
      talabat_url: null
    };

    submissions.push(submission);
    saveSubmissions();
    saveSubmissionToSheet(submission).catch(e => console.error('saveSubmissionToSheet error:', e.message));

    // Store in memory for instant display
    imageStore[submission.id] = { bankB64, bankMediaType, talabatB64, talabatMediaType };

    // Auto-write to sheet if approved and amount known
    if (status === 'approved' && aiResult.amount) {
      fillRiderRow(submission).catch(e => console.error('fillRiderRow error:', e.message));
    }

    // Respond to rider immediately — don't wait for Cloudinary
    res.json({ ok: true });

    // Upload to Cloudinary in background after responding
    if (cloudinary) {
      const riderSafeName = (riderObj.name || 'unknown').replace(/[^a-zA-Z0-9]/g, '_');
      Promise.all([
        uploadToCloudinary(bankB64, bankMediaType, today, `${submission.id}_${riderSafeName}_bank`),
        uploadToCloudinary(talabatB64, talabatMediaType, today, `${submission.id}_${riderSafeName}_talabat`)
      ]).then(([bankUrl, talabatUrl]) => {
        if (bankUrl) { submission.bank_url = bankUrl; }
        if (talabatUrl) { submission.talabat_url = talabatUrl; }
        if (bankUrl || talabatUrl) {
          saveSubmissions();
          updateSubmissionInSheet(submission).catch(e => console.error('updateSubmissionInSheet error:', e.message));
        }
      }).catch(e => console.error('Cloudinary background upload error:', e.message));
    }

    // WhatsApp alert in background
    const waMsg = status === 'flagged'
      ? `⚠️ *COD Alert — Flagged*\nRider: ${riderObj.name}\nAmount: ${aiResult.amount ? aiResult.amount + ' OMR' : 'Not detected'}\nReason: ${flags.join(', ')}`
      : `✅ *COD Submitted*\nRider: ${riderObj.name}\nAmount: ${aiResult.amount ? aiResult.amount + ' OMR' : 'Not detected'}${isLate ? '\n⏰ LATE submission' : ''}`;
    sendWhatsApp(waMsg);
  } catch (e) {
    console.error(e);
    res.json({ ok: false, error: 'Server error.' });
  }
});

// ─── Receipt image endpoint ───────────────────────────────────────────────────
app.get('/receipt/:id', requireAdmin, (req, res) => {
  const sub = submissions.find(s => s.id === req.params.id);
  if (sub && sub.bank_url) return res.redirect(sub.bank_url);
  const imgs = imageStore[req.params.id];
  if (imgs && imgs.bankB64) {
    res.setHeader('Content-Type', imgs.bankMediaType || 'image/jpeg');
    return res.send(Buffer.from(imgs.bankB64, 'base64'));
  }
  res.setHeader('Content-Type', 'image/svg+xml');
  res.send(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><rect width="200" height="200" fill="#f5f5f5" rx="12"/><text x="100" y="90" text-anchor="middle" font-size="40">📄</text><text x="100" y="130" text-anchor="middle" font-size="13" fill="#999">Image not available</text></svg>`);
});

app.get('/talabat/:id', requireAdmin, (req, res) => {
  const sub = submissions.find(s => s.id === req.params.id);
  if (sub && sub.talabat_url) return res.redirect(sub.talabat_url);
  const imgs = imageStore[req.params.id];
  if (imgs && imgs.talabatB64) {
    res.setHeader('Content-Type', imgs.talabatMediaType || 'image/jpeg');
    return res.send(Buffer.from(imgs.talabatB64, 'base64'));
  }
  res.setHeader('Content-Type', 'image/svg+xml');
  res.send(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><rect width="200" height="200" fill="#fff3e0" rx="12"/><text x="100" y="90" text-anchor="middle" font-size="40">🛵</text><text x="100" y="130" text-anchor="middle" font-size="13" fill="#999">Image not available</text></svg>`);
});


// ══════════════════════════════════════════════════════════════════════════════
app.get('/admin/login', (req, res) => {
  res.send(`<!DOCTYPE html><html><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin Login</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,sans-serif;background:#f5f5f5;display:flex;align-items:center;justify-content:center;min-height:100vh;}
  .card{background:#fff;padding:2rem;border-radius:16px;width:320px;box-shadow:0 2px 16px rgba(0,0,0,0.08);}
  h1{font-size:20px;margin-bottom:1.5rem;font-weight:600;}
  input{width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:8px;font-size:14px;margin-bottom:1rem;}
  button{width:100%;padding:11px;background:#1a73e8;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;}
  .err{color:#c62828;font-size:13px;margin-top:0.5rem;}
</style></head><body>
<div class="card">
  <h1>🔐 Admin Login</h1>
  <form method="POST" action="/admin/login">
    <input type="password" name="password" placeholder="Enter password" required autofocus>
    <button type="submit">Sign in</button>
    ${req.query.err ? '<p class="err">Wrong password. Try again.</p>' : ''}
  </form>
</div>
</body></html>`);
});

app.post('/admin/login', (req, res) => {
  if (req.body.password === process.env.ADMIN_PASSWORD) {
    req.session.admin = true;
    res.redirect('/admin');
  } else {
    res.redirect('/admin/login?err=1');
  }
});

app.get('/admin/logout', (req, res) => { req.session.destroy(); res.redirect('/admin/login'); });

// ─── WhatsApp notification via CallMeBot ─────────────────────────────────────
async function sendWhatsApp(message) {
  try {
    const phone = process.env.WA_PHONE;
    const apiKey = process.env.WA_APIKEY;
    if (!phone || !apiKey) return;
    const encoded = encodeURIComponent(message);
    const url = `https://api.callmebot.com/whatsapp.php?phone=${phone}&text=${encoded}&apikey=${apiKey}`;
    await fetch(url);
  } catch (e) { console.error('WhatsApp error:', e.message); }
}

// ─── 8 PM daily summary (GMT+4) ──────────────────────────────────────────────
function scheduleDailySummary() {
  const now = new Date();
  const gmt4 = new Date(now.getTime() + 4 * 60 * 60 * 1000);
  const next610pm = new Date(gmt4);
  next610pm.setHours(18, 10, 0, 0);
  if (gmt4 >= next610pm) next610pm.setDate(next610pm.getDate() + 1);
  const msUntil = next610pm - gmt4;
  setTimeout(async () => {
    const today = new Date(new Date().getTime() + 4*60*60*1000).toISOString().slice(0,10);
    const todaySubs = submissions.filter(s => s.date === today);
    const approved = todaySubs.filter(s => s.status === 'approved');
    const flagged = todaySubs.filter(s => s.status === 'flagged');
    const pending = riders.filter(r => !todaySubs.find(s => s.rider_id === r.id));
    const totalAmt = approved.reduce((s,e) => s+(parseFloat(e.amount)||0), 0);
    const msg = `📦 *COD Daily Summary - ${today}*\n\n✅ Submitted: ${approved.length}\n⚠️ Flagged: ${flagged.length}\n❌ Not submitted: ${pending.length}\n💰 Total COD: ${totalAmt.toFixed(3)} OMR\n\n${pending.length > 0 ? '❌ Missing:\n' + pending.slice(0,10).map(r=>`• ${r.name}`).join('\n') : '✅ All riders submitted!'}`;
    await sendWhatsApp(msg);
    scheduleDailySummary();
  }, msUntil);
}
scheduleDailySummary();

// ─── Weekly fuel sheet — every Sunday at 8 PM GMT+4 ──────────────────────────
async function generateFuelSheet() {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const now = new Date(new Date().getTime() + 4*60*60*1000);

    // Week: Monday to Sunday (current week ending today)
    const sunday = new Date(now);
    const monday = new Date(now);
    monday.setDate(sunday.getDate() - 6);

    const weekStart = monday.toISOString().slice(0,10);
    const weekEnd = sunday.toISOString().slice(0,10);
    const tabName = `Fuel ${weekStart} to ${weekEnd}`;

    // Read deliveries from Submissions sheet (column G=date, B=rider_id, C=rider_name, I=status, Q=deliveries)
    const subsRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Submissions!A:Q' });
    const subsRows = (subsRes.data.values || []).slice(1);

    // Filter approved submissions within the week range
    const weekRows = subsRows.filter(r => {
      const date = r[6]; // column G
      const status = r[8]; // column I
      return date && date >= weekStart && date <= weekEnd && status === 'approved';
    });

    // Sum deliveries per rider
    const riderDeliveries = {};
    weekRows.forEach(r => {
      const riderId = r[1]; // column B
      const riderName = r[2]; // column C
      const deliveries = parseInt(r[16]) || 0; // column Q
      if (!riderId) return;
      if (!riderDeliveries[riderId]) {
        riderDeliveries[riderId] = { name: riderName, id: riderId, deliveries: 0 };
      }
      riderDeliveries[riderId].deliveries += deliveries;
    });

    // Build rows for all riders
    const riderRows = riders.map((r, i) => {
      const data = riderDeliveries[r.id];
      const deliveries = data ? data.deliveries : 0;
      const fuel = deliveries === 0 ? 0 : deliveries >= 75 ? (deliveries > 75 ? 30 : 25) : 0;
      return [i + 1, r.name, r.id, deliveries, fuel];
    });

    // Create sheet tab
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const existing = meta.data.sheets.map(s => s.properties.title);
    if (existing.includes(tabName)) return; // already generated

    const addRes = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: tabName } } }] }
    });
    const sheetId = addRes.data.replies[0].addSheet.properties.sheetId;
    const COLS = 5;
    const totalRows = riders.length + 3;

    // Write data
    await sheets.spreadsheets.values.update({
      spreadsheetId, range: `${tabName}!A1`, valueInputOption: 'RAW',
      requestBody: { values: [
        ['Future Wave', '', '', '', ''],
        [`Fuel Allowance — ${weekStart} to ${weekEnd}`, '', '', '', ''],
        ['#', 'Rider Name', 'Rider ID', 'Total Orders', 'Fuel (OMR)'],
        ...riderRows
      ]}
    });

    // Formatting
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [
        // Merge title rows
        { mergeCells: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: COLS }, mergeType: 'MERGE_ALL' } },
        { mergeCells: { range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: COLS }, mergeType: 'MERGE_ALL' } },

        // Company name row — dark green
        { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: COLS },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.067, green: 0.31, blue: 0.165 }, horizontalAlignment: 'CENTER',
            textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 16 } } }, fields: 'userEnteredFormat' }},

        // Date row — light green
        { repeatCell: { range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: COLS },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.851, green: 0.918, blue: 0.863 }, horizontalAlignment: 'CENTER',
            textFormat: { foregroundColor: { red: 0.067, green: 0.31, blue: 0.165 }, bold: true, fontSize: 12 } } }, fields: 'userEnteredFormat' }},

        // Header row — dark green
        { repeatCell: { range: { sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: COLS },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.067, green: 0.31, blue: 0.165 }, horizontalAlignment: 'CENTER',
            textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 11 } } }, fields: 'userEnteredFormat' }},

        // Data rows — alternating white/grey
        ...riders.map((_, i) => ({
          repeatCell: {
            range: { sheetId, startRowIndex: 3 + i, endRowIndex: 4 + i, startColumnIndex: 0, endColumnIndex: COLS },
            cell: { userEnteredFormat: {
              backgroundColor: i % 2 === 0 ? { red: 1, green: 1, blue: 1 } : { red: 0.949, green: 0.949, blue: 0.949 },
              horizontalAlignment: 'CENTER', textFormat: { fontSize: 10 }
            }}, fields: 'userEnteredFormat'
          }
        })),

        // Fuel column (E) — highlight non-zero in green
        ...riderRows.map((row, i) => row[4] > 0 ? ({
          repeatCell: {
            range: { sheetId, startRowIndex: 3 + i, endRowIndex: 4 + i, startColumnIndex: 4, endColumnIndex: 5 },
            cell: { userEnteredFormat: {
              backgroundColor: { red: 0.714, green: 0.918, blue: 0.757 },
              textFormat: { bold: true, foregroundColor: { red: 0.067, green: 0.31, blue: 0.165 } }
            }}, fields: 'userEnteredFormat'
          }
        }) : null).filter(Boolean),

        // Row heights
        { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 45 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 3, endIndex: totalRows }, properties: { pixelSize: 25 }, fields: 'pixelSize' } },

        // Column widths
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 180 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 100 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },

        // Borders
        { updateBorders: { range: { sheetId, startRowIndex: 2, endRowIndex: totalRows, startColumnIndex: 0, endColumnIndex: COLS },
          top: { style: 'SOLID', width: 2, color: { red: 0.067, green: 0.31, blue: 0.165 } },
          bottom: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } },
          left: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } },
          right: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } },
          innerHorizontal: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } },
          innerVertical: { style: 'SOLID', width: 1, color: { red: 0.8, green: 0.8, blue: 0.8 } }
        }}
      ]}
    });

    console.log(`Fuel sheet generated: ${tabName}`);
  } catch(e) { console.error('generateFuelSheet error:', e.message); }
}

async function generatePerformanceSheet() {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const now = new Date(new Date().getTime() + 4*60*60*1000);
    const sunday = new Date(now);
    const monday = new Date(now);
    monday.setDate(sunday.getDate() - 6);
    const weekStart = monday.toISOString().slice(0,10);
    const weekEnd = sunday.toISOString().slice(0,10);
    const tabName = `Perf ${weekStart} to ${weekEnd}`;

    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    if (meta.data.sheets.map(s => s.properties.title).includes(tabName)) return;

    // Read from Submissions sheet
    const subsRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Submissions!A:Q' });
    const subsRows = (subsRes.data.values || []).slice(1);
    const weekRows = subsRows.filter(r => r[6] && r[6] >= weekStart && r[6] <= weekEnd);

    // Build per-rider stats
    const riderStats = {};
    riders.forEach(r => {
      riderStats[r.id] = { name: r.name, id: r.id, days: 0, orders: 0, submitted: false };
    });
    weekRows.forEach(r => {
      const id = r[1];
      if (!riderStats[id]) riderStats[id] = { name: r[2], id, days: 0, orders: 0 };
      riderStats[id].days++;
      riderStats[id].orders += (parseInt(r[16]) || 0);
      riderStats[id].submitted = true;
    });

    const COLS = 6;
    const riderRows = riders.map((r, i) => {
      const s = riderStats[r.id] || { days: 0, orders: 0, submitted: false };
      const flag = !s.submitted ? '🔴 Did not work' : s.orders < 75 ? '⚠️ Low orders' : '✅ Good';
      return [i + 1, r.name, r.id, s.days, s.orders, flag];
    });

    const addRes = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: tabName } } }] }
    });
    const sheetId = addRes.data.replies[0].addSheet.properties.sheetId;

    await sheets.spreadsheets.values.update({
      spreadsheetId, range: `${tabName}!A1`, valueInputOption: 'RAW',
      requestBody: { values: [
        ['Future Wave', '', '', '', '', ''],
        [`Performance — ${weekStart} to ${weekEnd}`, '', '', '', '', ''],
        ['#', 'Rider Name', 'Rider ID', 'Days Worked', 'Total Orders', 'Performance'],
        ...riderRows
      ]}
    });

    const totalRows = riders.length + 3;
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [
        { mergeCells: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: COLS }, mergeType: 'MERGE_ALL' } },
        { mergeCells: { range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: COLS }, mergeType: 'MERGE_ALL' } },
        { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: COLS },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.067, green: 0.31, blue: 0.165 }, horizontalAlignment: 'CENTER',
            textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 16 } } }, fields: 'userEnteredFormat' }},
        { repeatCell: { range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: COLS },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.851, green: 0.918, blue: 0.863 }, horizontalAlignment: 'CENTER',
            textFormat: { foregroundColor: { red: 0.067, green: 0.31, blue: 0.165 }, bold: true, fontSize: 12 } } }, fields: 'userEnteredFormat' }},
        { repeatCell: { range: { sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: COLS },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.067, green: 0.31, blue: 0.165 }, horizontalAlignment: 'CENTER',
            textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 11 } } }, fields: 'userEnteredFormat' }},
        // Color rows by performance
        ...riderRows.map((row, i) => {
          const flag = row[5];
          const bg = flag.includes('Did not work') ? { red: 0.957, green: 0.8, blue: 0.8 }
            : flag.includes('Low orders') ? { red: 1.0, green: 0.953, blue: 0.714 }
            : { red: 0.714, green: 0.918, blue: 0.757 };
          return { repeatCell: { range: { sheetId, startRowIndex: 3+i, endRowIndex: 4+i, startColumnIndex: 0, endColumnIndex: COLS },
            cell: { userEnteredFormat: { backgroundColor: bg, horizontalAlignment: 'CENTER', textFormat: { fontSize: 10 } } }, fields: 'userEnteredFormat' }};
        }),
        { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 45 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 1, endIndex: totalRows }, properties: { pixelSize: 25 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 180 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 150 }, fields: 'pixelSize' } },
      ]}
    });
    console.log(`Performance sheet generated: ${tabName}`);
  } catch(e) { console.error('generatePerformanceSheet error:', e.message); }
}

function scheduleWeeklyFuel() {
  const now = new Date();
  const gmt4 = new Date(now.getTime() + 4*60*60*1000);
  const next = new Date(gmt4);
  const dayOfWeek = gmt4.getDay();
  const daysUntilSunday = dayOfWeek === 0 ? 0 : 7 - dayOfWeek;
  next.setDate(gmt4.getDate() + daysUntilSunday);
  next.setHours(20, 0, 0, 0);
  if (gmt4 >= next) next.setDate(next.getDate() + 7);
  const msUntil = next - gmt4;
  setTimeout(async () => {
    await generateFuelSheet();
    await generatePerformanceSheet();
    scheduleWeeklyFuel();
  }, msUntil);
  console.log(`Fuel sheet scheduled for: ${next.toISOString()}`);
}
scheduleWeeklyFuel();

app.post('/admin/generate-fuel', requireAdmin, async (req, res) => {
  generateFuelSheet().catch(e => console.error('Fuel sheet error:', e.message));
  generatePerformanceSheet().catch(e => console.error('Performance sheet error:', e.message));
  res.redirect('/admin?fuel=generating');
});

// ══════════════════════════════════════════════════════════════════════════════
// ADMIN — dashboard

// ══════════════════════════════════════════════════════════════════════════════
app.get('/admin', requireAdmin, (req, res) => {
  const today = new Date(new Date().getTime() + 4*60*60*1000).toISOString().slice(0,10);
  const selectedDate = req.query.date || today;
  const isToday = selectedDate === today;
  const todaySubs = submissions.filter(s => s.date === selectedDate);
  const approved = todaySubs.filter(s => s.status === 'approved');
  const flagged = todaySubs.filter(s => s.status === 'flagged');
  const totalAmt = approved.reduce((s, e) => s + (parseFloat(e.amount) || 0), 0);
  const submittedIds = new Set(todaySubs.map(s => s.rider_id));
  const pendingRiders = isToday ? riders.filter(r => !submittedIds.has(r.id)) : [];

  const rows = todaySubs.sort((a, b) => b.submitted_at.localeCompare(a.submitted_at)).map(s => {
    const hasBank = s.bank_url || imageStore[s.id]?.bankB64;
    const hasTalabat = s.talabat_url || imageStore[s.id]?.talabatB64;
    return `
    <tr>
      <td>
        <a href="/receipt/${s.id}" target="_blank">
          ${hasBank
            ? `<img src="${s.bank_url || `/receipt/${s.id}`}" style="width:48px;height:48px;object-fit:cover;border-radius:8px;border:1px solid #eee;display:block;" alt="receipt" onerror="this.src='/receipt/${s.id}'">`
            : `<div style="width:48px;height:48px;background:#f5f5f5;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:20px;">📄</div>`}
        </a>
      </td>
      <td style="font-weight:500;font-size:13px;">${s.rider_name}${s.is_late ? ' <span style="background:#fff3e0;color:#e65100;font-size:10px;padding:1px 5px;border-radius:4px;">LATE</span>' : ''}</td>
      <td style="color:#888;font-size:12px;">${s.rider_id}</td>
      <td style="font-weight:600;font-size:13px;">${s.amount ? s.amount.toLocaleString() + ' OMR' : '<span style="color:#c62828;font-size:12px;">—</span>'}</td>
      <td style="font-weight:600;color:#e65100;font-size:13px;">${s.talabat_amount ? s.talabat_amount.toLocaleString() + ' OMR' : '—'}</td>
      <td style="font-size:12px;color:#888;">${new Date(new Date(s.submitted_at).getTime() + 4*60*60*1000).toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit',hour12:true})}</td>
      <td>
        <span class="badge ${s.status === 'approved' ? 'ok' : s.status === 'flagged' ? 'flagged' : 'rej'}">
          ${s.status === 'approved' ? '✓' : s.status === 'flagged' ? '⚠' : '✗'} ${s.status === 'approved' ? 'OK' : s.status === 'flagged' ? 'Flag' : 'Rej'}
        </span>
        ${s.status === 'flagged' || s.needs_amount ? `
          ${s.flags && s.flags.length ? `<div style="font-size:10px;color:#c62828;margin-top:3px;">${s.flags.join('<br>')}</div>` : ''}
          ${s.account_name ? `<div style="font-size:10px;color:#888;margin-top:2px;">👤 ${s.account_name}</div>` : ''}
          ${s.beneficiary_name ? `<div style="font-size:10px;color:#888;">→ ${s.beneficiary_name}</div>` : ''}
          <form method="POST" action="/admin/approve/${s.id}" style="margin:5px 0 3px;">
            <input name="manual_amount" type="number" step="0.01" placeholder="Amount (OMR)" value="${s.amount || ''}" style="width:100%;padding:4px 7px;border:1px solid #ddd;border-radius:6px;font-size:12px;margin-bottom:4px;" ${s.amount ? '' : 'required'}>
            <div style="display:flex;gap:4px;flex-wrap:wrap;">
              <button class="act-btn ok-btn">✓ Approve</button>
              ${hasBank ? `<a href="/receipt/${s.id}" target="_blank" class="act-btn" style="background:#e8f0fe;color:#1a73e8;text-decoration:none;">👁 Bank</a>` : ''}
              ${hasTalabat ? `<a href="/talabat/${s.id}" target="_blank" class="act-btn" style="background:#fff3e0;color:#e65100;text-decoration:none;">🛵</a>` : ''}
            </div>
          </form>
          <div style="display:flex;gap:4px;margin-top:3px;">
            ${s.status === 'flagged' ? `
            <form method="POST" action="/admin/reject/${s.id}" style="margin:3px 0 0;">
              <input name="reason" type="text" placeholder="Rejection reason..." style="width:100%;padding:3px 7px;border:1px solid #ddd;border-radius:6px;font-size:11px;margin-bottom:3px;">
              <button class="act-btn rej-btn">✗ Reject</button>
            </form>` : ''}
            <form method="POST" action="/admin/delete/${s.id}" onsubmit="return confirm('Delete?')" style="margin:0;"><button class="act-btn" style="background:#f5f5f5;color:#888;">🗑</button></form>
          </div>` :
          `<div style="margin-top:3px;display:flex;gap:4px;flex-wrap:wrap;align-items:center;">
            ${hasBank ? `<a href="/receipt/${s.id}" target="_blank" style="font-size:11px;color:#1a73e8;">👁 Bank</a>` : ''}
            ${hasTalabat ? `<a href="/talabat/${s.id}" target="_blank" style="font-size:11px;color:#e65100;">🛵 Talabat</a>` : ''}
            <a href="/admin/edit/${s.id}" style="font-size:11px;color:#888;">✏️ Edit</a>
            <form method="POST" action="/admin/delete/${s.id}" onsubmit="return confirm('Delete?')" style="margin:0;"><button class="act-btn" style="background:#f5f5f5;color:#888;font-size:11px;">🗑</button></form>
          </div>`}
      </td>
    </tr>`;
  }).join('');

  const pendingRows = pendingRiders.map(r => `
    <div style="display:flex;align-items:center;justify-content:space-between;padding:8px 0;border-bottom:1px solid #f5f5f5;">
      <span style="font-size:13px;">❌ ${r.name}</span>
      <span style="font-size:11px;color:#aaa;">${r.id}</span>
    </div>`).join('');

  res.send(`<!DOCTYPE html><html><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>COD Dashboard</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f5;color:#222;font-size:14px;}
  .topbar{background:#fff;padding:12px 16px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #eee;position:sticky;top:0;z-index:10;}
  .topbar h1{font-size:16px;font-weight:700;}
  .topbar-links{display:flex;gap:8px;}
  .topbar-links a{font-size:12px;color:#1a73e8;text-decoration:none;padding:5px 10px;border:1px solid #ddd;border-radius:8px;white-space:nowrap;}
  .content{padding:12px;max-width:900px;margin:0 auto;}
  .stats{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin-bottom:12px;}
  @media(min-width:600px){.stats{grid-template-columns:repeat(4,1fr);}}
  .stat{background:#fff;border-radius:12px;padding:12px;border:1px solid #eee;text-align:center;}
  .stat-val{font-size:24px;font-weight:700;}
  .stat-lbl{font-size:11px;color:#888;margin-top:2px;}
  .section{background:#fff;border-radius:12px;border:1px solid #eee;margin-bottom:12px;overflow:hidden;}
  .section-hdr{padding:10px 14px;font-size:13px;font-weight:600;border-bottom:1px solid #eee;display:flex;align-items:center;justify-content:space-between;}
  .pending-list{padding:0 14px;max-height:200px;overflow-y:auto;}
  table{width:100%;border-collapse:collapse;}
  th{text-align:left;font-size:11px;font-weight:600;color:#888;padding:8px 10px;border-bottom:1px solid #eee;text-transform:uppercase;white-space:nowrap;}
  td{padding:8px 10px;font-size:13px;border-bottom:1px solid #f5f5f5;vertical-align:top;}
  tr:last-child td{border-bottom:none;}
  .badge{font-size:11px;font-weight:600;padding:2px 7px;border-radius:999px;display:inline-block;white-space:nowrap;}
  .badge.ok{background:#e6f4ea;color:#1e7e34;}
  .badge.flagged{background:#fce8e6;color:#c62828;}
  .badge.rej{background:#f5f5f5;color:#888;}
  .act-btn{font-size:11px;padding:3px 8px;border-radius:6px;cursor:pointer;border:none;font-weight:500;display:inline-block;}
  .ok-btn{background:#e6f4ea;color:#1e7e34;}
  .rej-btn{background:#fce8e6;color:#c62828;}
  @media(max-width:500px){th:nth-child(3),td:nth-child(3),th:nth-child(5),td:nth-child(5){display:none;}}
</style>
<meta http-equiv="refresh" content="15">
<script>
let refreshTimer;
function resetRefresh() {
  clearTimeout(refreshTimer);
  refreshTimer = setTimeout(() => location.reload(), 15000);
}
document.addEventListener('DOMContentLoaded', () => {
  resetRefresh();
  document.querySelectorAll('input, select').forEach(el => {
    el.addEventListener('focus', () => clearTimeout(refreshTimer));
    el.addEventListener('blur', resetRefresh);
  });
});
</script>
</head><body>
<div class="topbar">
  <h1>📦 COD${isToday ? '' : ` — ${selectedDate}`}</h1>
  <div class="topbar-links">
    <form method="GET" action="/admin" style="margin:0;display:flex;gap:6px;align-items:center;">
      <input type="date" name="date" value="${selectedDate}" max="${today}" style="font-size:12px;padding:4px 8px;border:1px solid #ddd;border-radius:8px;">
      <button type="submit" style="font-size:12px;padding:5px 10px;background:#1a73e8;color:#fff;border:none;border-radius:8px;cursor:pointer;">Go</button>
    </form>
    <a href="/admin/riders">👥 Riders</a>
    <a href="/admin/export">⬇️ CSV</a>
    <a href="/admin/fuel-upload">⛽ Fuel</a>
    <a href="/admin/logout">🚪</a>
  </div>
</div>
<div class="content">
  <div class="stats">
    <div class="stat"><div class="stat-val">${todaySubs.length}</div><div class="stat-lbl">Submitted</div></div>
    <div class="stat"><div class="stat-val" style="color:#c62828;">${pendingRiders.length}</div><div class="stat-lbl">Not submitted</div></div>
    <div class="stat"><div class="stat-val" style="color:#e65100;">${flagged.length}</div><div class="stat-lbl">Flagged</div></div>
    <div class="stat"><div class="stat-val" style="font-size:16px;padding-top:6px;">${totalAmt.toFixed(3)}<br><span style="font-size:11px;color:#888;">OMR</span></div><div class="stat-lbl">Total COD</div></div>
  </div>

  ${pendingRiders.length > 0 ? `
  <div class="section">
    <div class="section-hdr"><span>❌ Not submitted yet (${pendingRiders.length})</span></div>
    <div class="pending-list">${pendingRows}</div>
  </div>` : `<div class="section"><div class="section-hdr" style="color:#1e7e34;">✅ All riders submitted today!</div></div>`}

  <div class="section">
    <div class="section-hdr">📋 Today's submissions</div>
    <div style="overflow-x:auto;">
    <table>
      <thead><tr><th>📄</th><th>Rider</th><th>ID</th><th>Bank</th><th>Talabat</th><th>Time</th><th>Status</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="7" style="text-align:center;padding:2rem;color:#aaa;">No submissions yet</td></tr>'}</tbody>
    </table>
    </div>
  </div>
</div>
</body></html>`);
});

// ─── Approve / Reject ─────────────────────────────────────────────────────────
app.post('/admin/approve/:id', requireAdmin, async (req, res) => {
  const sub = submissions.find(s => s.id === req.params.id);
  if (sub) {
    sub.status = 'approved';
    if (req.body.manual_amount) {
      sub.amount = parseFloat(req.body.manual_amount);
    }
    sub.bank = sub.bank || 'Bank';
    saveSubmissions();
    res.redirect('/admin');
    updateSubmissionInSheet(sub).catch(e => console.error('updateSubmissionInSheet error:', e.message));
    fillRiderRow(sub).catch(e => console.error('Sheet write error:', e.message));
  } else {
    res.redirect('/admin');
  }
});

// ─── Edit amount ──────────────────────────────────────────────────────────────
app.get('/admin/edit/:id', requireAdmin, (req, res) => {
  const sub = submissions.find(s => s.id === req.params.id);
  if (!sub) return res.redirect('/admin');
  res.send(`<!DOCTYPE html><html><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Edit Amount</title>
<style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:-apple-system,sans-serif;background:#f5f5f5;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:1rem;}.card{background:#fff;padding:1.5rem;border-radius:16px;width:100%;max-width:360px;}h2{font-size:16px;margin-bottom:1rem;}label{font-size:13px;color:#666;display:block;margin-bottom:4px;}input{width:100%;padding:10px;border:1px solid #ddd;border-radius:8px;font-size:14px;margin-bottom:1rem;}button{width:100%;padding:11px;background:#1a73e8;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;}.back{display:block;text-align:center;margin-top:0.75rem;font-size:13px;color:#1a73e8;text-decoration:none;}</style>
</head><body><div class="card">
  <h2>✏️ Edit — ${sub.rider_name}</h2>
  <form method="POST" action="/admin/edit/${sub.id}">
    <label>Bank Amount (OMR)</label>
    <input type="number" step="0.001" name="amount" value="${sub.amount || ''}" placeholder="e.g. 11.960" required>
    <label>Talabat Collected (OMR)</label>
    <input type="number" step="0.01" name="talabat_amount" value="${sub.talabat_amount || ''}" placeholder="e.g. 55.90">
    <button type="submit">Save changes</button>
  </form>
  <a href="/admin" class="back">← Back to dashboard</a>
</div></body></html>`);
});

app.post('/admin/edit/:id', requireAdmin, async (req, res) => {
  const sub = submissions.find(s => s.id === req.params.id);
  if (sub) {
    sub.amount = parseFloat(req.body.amount) || sub.amount;
    sub.talabat_amount = parseFloat(req.body.talabat_amount) || sub.talabat_amount;
    saveSubmissions();
    if (sub.status === 'approved') await fillRiderRow(sub);
  }
  res.redirect('/admin');
});


app.post('/admin/delete/:id', requireAdmin, (req, res) => {
  submissions = submissions.filter(s => s.id !== req.params.id);
  saveSubmissions();
  res.redirect('/admin');
});


app.post('/admin/reject/:id', requireAdmin, (req, res) => {
  const sub = submissions.find(s => s.id === req.params.id);
  if (sub) {
    sub.status = 'rejected';
    sub.rejection_reason = req.body.reason || '';
    saveSubmissions();
    updateSubmissionInSheet(sub).catch(e => console.error(e.message));
  }
  res.redirect('/admin');
});

// ══════════════════════════════════════════════════════════════════════════════
// ADMIN — manage riders
// ══════════════════════════════════════════════════════════════════════════════
app.get('/admin/riders', requireAdmin, (req, res) => {
  const rows = riders.map(r => `
    <tr>
      <td>${r.name}</td>
      <td style="color:#888;">${r.id}</td>
      <td>
        <form method="POST" action="/admin/riders/delete/${r.id}" style="margin:0;" onsubmit="return confirm('Remove ${r.name}?')">
          <button class="del-btn">Remove</button>
        </form>
      </td>
    </tr>`).join('');

  res.send(`<!DOCTYPE html><html><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Manage Riders</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,sans-serif;background:#f5f5f5;}
  .topbar{background:#fff;padding:0.875rem 1.5rem;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #eee;}
  .topbar h1{font-size:17px;font-weight:600;}
  .topbar a{font-size:13px;color:#1a73e8;text-decoration:none;}
  .content{padding:1.5rem;max-width:700px;margin:0 auto;}
  .add-form{background:#fff;border-radius:12px;padding:1.25rem;border:1px solid #eee;margin-bottom:1.5rem;display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;}
  .field{display:flex;flex-direction:column;gap:4px;flex:1;min-width:140px;}
  label{font-size:12px;font-weight:500;color:#666;}
  input{padding:9px 12px;border:1px solid #ddd;border-radius:8px;font-size:14px;}
  .add-btn{padding:9px 20px;background:#1a73e8;color:#fff;border:none;border-radius:8px;font-weight:600;cursor:pointer;font-size:14px;white-space:nowrap;}
  table{width:100%;background:#fff;border-radius:12px;border-collapse:collapse;border:1px solid #eee;}
  th{text-align:left;font-size:12px;font-weight:600;color:#888;padding:10px 12px;border-bottom:1px solid #eee;text-transform:uppercase;}
  td{padding:10px 12px;font-size:13px;border-bottom:1px solid #f5f5f5;}
  tr:last-child td{border-bottom:none;}
  .del-btn{font-size:12px;padding:4px 10px;background:#fce8e6;color:#c62828;border:none;border-radius:6px;cursor:pointer;}
  .count{font-size:13px;color:#888;margin-bottom:0.75rem;}
</style></head><body>
<div class="topbar"><h1>👥 Manage Riders</h1><a href="/admin">← Back to dashboard</a></div>
<div class="content">
  <form method="POST" action="/admin/riders/add" class="add-form">
    <div class="field"><label>Rider name</label><input name="name" placeholder="e.g. Mohammed Ali" required></div>
    <div class="field"><label>Rider ID</label><input name="id" placeholder="e.g. 123456" required></div>
    <button type="submit" class="add-btn">+ Add rider</button>
  </form>
  <p class="count">${riders.length} riders total</p>
  <table>
    <thead><tr><th>Name</th><th>ID</th><th></th></tr></thead>
    <tbody>${rows}</tbody>
  </table>
</div>
</body></html>`);
});

app.post('/admin/riders/add', requireAdmin, async (req, res) => {
  const { name, id } = req.body;
  if (name && id && !riders.find(r => r.id === id.trim())) {
    riders.push({ id: id.trim(), name: name.trim() });
    await saveRidersToSheet();
  }
  res.redirect('/admin/riders');
});

app.post('/admin/riders/delete/:id', requireAdmin, async (req, res) => {
  riders = riders.filter(r => r.id !== req.params.id);
  await saveRidersToSheet();
  res.redirect('/admin/riders');
});

// ─── Debug endpoint ───────────────────────────────────────────────────────────
app.get('/admin/debug', requireAdmin, (req, res) => {
  const today = new Date(new Date().getTime() + 4*60*60*1000).toISOString().slice(0,10);
  res.json({
    total_submissions: submissions.length,
    today,
    today_count: submissions.filter(s => s.date === today).length,
    dates_in_memory: [...new Set(submissions.map(s => s.date))].sort().reverse(),
    last_5: submissions.slice(-5).map(s => ({ id: s.id, date: s.date, rider: s.rider_name, status: s.status, has_bank_url: !!s.bank_url }))
  });
});


app.get('/admin/export', requireAdmin, (req, res) => {
  const today = new Date().toISOString().slice(0, 10);
  const rows = submissions.filter(s => s.date === today && s.status === 'approved');
  const csv = [
    ['Date', 'Rider Name', 'Rider ID', 'Amount (OMR)', 'Bank', 'Submitted At'],
    ...rows.map(r => [r.date, r.rider_name, r.rider_id, r.amount || '', r.bank || '', r.submitted_at])
  ].map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename=COD_${today}.csv`);
  res.send(csv);
});

// ─── Start ────────────────────────────────────────────────────────────────────
// ─── Fuel data storage (Google Sheet tab: FuelData) ──────────────────────────
async function saveFuelDataToSheet(weekLabel, rows) {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const existing = meta.data.sheets.map(s => s.properties.title);

    if (!existing.includes('FuelData')) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: 'FuelData' } } }] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId, range: 'FuelData!A1',
        valueInputOption: 'RAW',
        requestBody: { values: [['Week', 'Rider ID', 'Rider Name', 'Fuel Amount (OMR)', 'Uploaded At']] }
      });
    }

    // Remove existing entries for this week first
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'FuelData!A:A' });
    const allWeeks = (res.data.values || []).map(r => r[0]);
    const toDelete = [];
    allWeeks.forEach((w, i) => { if (w === weekLabel) toDelete.push(i + 1); });
    // Clear existing week rows (replace with empty)
    for (const rowNum of toDelete.reverse()) {
      await sheets.spreadsheets.values.update({
        spreadsheetId, range: `FuelData!A${rowNum}:E${rowNum}`,
        valueInputOption: 'RAW', requestBody: { values: [['', '', '', '', '']] }
      });
    }

    // Append new rows
    const uploadedAt = new Date(new Date().getTime() + 4*60*60*1000).toLocaleString('en-US');
    const newRows = rows.map(r => [weekLabel, r.id, r.name, r.amount, uploadedAt]);
    await sheets.spreadsheets.values.append({
      spreadsheetId, range: 'FuelData!A:E',
      valueInputOption: 'RAW',
      requestBody: { values: newRows }
    });
    return true;
  } catch(e) { console.error('saveFuelDataToSheet error:', e.message); return false; }
}

async function getFuelData(riderId) {
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'FuelData!A:E' });
    const rows = (res.data.values || []).slice(1);
    const riderRows = rows.filter(r => r[1] && r[1].toString().trim() === riderId.toString().trim() && r[0]);
    return riderRows.map(r => ({ week: r[0], rider_id: r[1], rider_name: r[2], amount: r[3] }));
  } catch(e) { console.error('getFuelData error:', e.message); return []; }
}

// ─── Admin fuel upload page ───────────────────────────────────────────────────
app.get('/admin/fuel-upload', requireAdmin, (req, res) => {
  res.send(`<!DOCTYPE html><html><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Upload Fuel Sheet</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,sans-serif;background:#f5f5f5;padding:1.5rem;}
  .card{background:#fff;border-radius:16px;padding:1.5rem;max-width:500px;margin:0 auto;box-shadow:0 2px 16px rgba(0,0,0,0.08);}
  h1{font-size:18px;font-weight:600;margin-bottom:4px;}
  .sub{font-size:13px;color:#888;margin-bottom:1.5rem;}
  label{font-size:13px;font-weight:500;color:#444;display:block;margin-bottom:4px;}
  input[type=text],input[type=file]{width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:10px;font-size:14px;margin-bottom:1rem;background:#fafafa;}
  button{width:100%;padding:12px;background:#e65100;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;}
  .back{display:block;text-align:center;margin-top:1rem;font-size:13px;color:#1a73e8;text-decoration:none;}
  .info{background:#e8f0fe;border-radius:10px;padding:12px;font-size:12px;color:#1a73e8;margin-bottom:1rem;line-height:1.6;}
  .msg{padding:12px;border-radius:10px;font-size:13px;margin-top:1rem;}
  .msg.ok{background:#e6f4ea;color:#1e7e34;}
  .msg.err{background:#fce8e6;color:#c62828;}
</style></head><body>
<div class="card">
  <h1>⛽ Upload Fuel Sheet</h1>
  <p class="sub">Upload the weekly Excel file to update rider fuel amounts</p>
  <div class="info">
    📋 Excel file must have these columns:<br>
    <strong>Column A:</strong> Rider ID<br>
    <strong>Column B:</strong> Rider Name<br>
    <strong>Column C:</strong> Fuel Amount (OMR)<br>
    First row = headers (will be skipped)
  </div>
  <form method="POST" action="/admin/fuel-upload" enctype="multipart/form-data">
    <label>Week label (e.g. 19-05 to 25-05)</label>
    <input type="text" name="week_label" placeholder="e.g. 19-05-2026 to 25-05-2026" required>
    <label>Excel file (.xlsx)</label>
    <input type="file" name="fuel_file" accept=".xlsx" required>
    <button type="submit">⬆️ Upload & Save</button>
  </form>
  ${req.query.success ? '<div class="msg ok">✅ Fuel data uploaded successfully!</div>' : ''}
  ${req.query.error ? '<div class="msg err">❌ Error: ' + req.query.error + '</div>' : ''}
  <a href="/admin" class="back">← Back to dashboard</a>
</div>
</body></html>`);
});

const uploadFuel = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });
app.post('/admin/fuel-upload', requireAdmin, uploadFuel.single('fuel_file'), async (req, res) => {
  try {
    const weekLabel = req.body.week_label;
    if (!req.file) return res.redirect('/admin/fuel-upload?error=No+file+uploaded');
    if (!weekLabel) return res.redirect('/admin/fuel-upload?error=No+week+label');

    if (!XLSX) return res.redirect('/admin/fuel-upload?error=xlsx+module+not+available');
    // Parse Excel
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Skip header row, extract ID, Name, Amount
    const rows = data.slice(1).filter(r => r[0] && r[2]).map(r => ({
      id: r[0].toString().trim(),
      name: (r[1] || '').toString().trim(),
      amount: parseFloat(r[2]) || 0
    }));

    if (rows.length === 0) return res.redirect('/admin/fuel-upload?error=No+valid+data+found+in+file');

    const saved = await saveFuelDataToSheet(weekLabel, rows);
    if (saved) res.redirect('/admin/fuel-upload?success=1');
    else res.redirect('/admin/fuel-upload?error=Failed+to+save+to+sheet');
  } catch(e) {
    console.error('Fuel upload error:', e.message);
    res.redirect('/admin/fuel-upload?error=' + encodeURIComponent(e.message));
  }
});


// ─── Rider status page ────────────────────────────────────────────────────────
app.get('/status', (req, res) => {
  res.send(`<!DOCTYPE html><html><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Submission Status</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:1rem;}
  .card{background:#fff;border-radius:16px;padding:1.5rem;width:100%;max-width:400px;box-shadow:0 2px 16px rgba(0,0,0,0.08);}
  h1{font-size:20px;font-weight:600;margin-bottom:4px;}
  .sub{font-size:13px;color:#888;margin-bottom:1.5rem;}
  input{width:100%;padding:12px;border:1px solid #ddd;border-radius:10px;font-size:16px;margin-bottom:1rem;background:#fafafa;}
  button{width:100%;padding:12px;background:#1a73e8;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;}
  .result{margin-top:1.25rem;border-radius:12px;padding:1.25rem;text-align:center;}
  .approved{background:#e6f4ea;border:1px solid #b7dfbe;}
  .flagged{background:#fff3e0;border:1px solid #ffe0b2;}
  .rejected{background:#fce8e6;border:1px solid #f5c6c2;}
  .none{background:#f5f5f5;border:1px solid #eee;}
  .icon{font-size:40px;margin-bottom:8px;}
  .status-text{font-size:16px;font-weight:600;margin-bottom:4px;}
  .status-sub{font-size:13px;color:#666;}
  .amount{font-size:24px;font-weight:700;margin-top:10px;color:#1e7e34;}
</style></head><body>
<div class="card">
  <h1>📋 Submission Status</h1>
  <p class="sub">Enter your ID to check today's submission</p>
  <input type="number" id="riderId" placeholder="Enter your rider ID" inputmode="numeric">
  <button onclick="checkStatus()">Check status</button>
  <div id="result"></div>
</div>
<script>
async function checkStatus() {
  const id = document.getElementById('riderId').value.trim();
  if (!id) return;
  document.getElementById('result').innerHTML = '<div style="text-align:center;padding:1rem;color:#888;font-size:13px;">Checking...</div>';
  const res = await fetch('/status/check?id=' + encodeURIComponent(id));
  const data = await res.json();
  const el = document.getElementById('result');
  if (!data.found) {
    el.innerHTML = '<div class="result none"><div class="icon">📭</div><div class="status-text">No submission today</div><div class="status-sub">You have not submitted yet today.</div></div>';
    return;
  }
  const icons = {approved:'✅',flagged:'⏳',rejected:'❌'};
  const texts = {approved:'Approved',flagged:'Under Review',rejected:'Rejected'};
  const subs = {approved:'Your submission has been approved.',flagged:'Your submission is being reviewed.',rejected:'Your submission was rejected.' + (data.reason ? '<br><b>Reason: '+data.reason+'</b>' : '')};
  el.innerHTML = '<div class="result '+data.status+'"><div class="icon">'+icons[data.status]+'</div><div class="status-text">'+texts[data.status]+'</div><div class="status-sub">'+subs[data.status]+'</div>'+(data.amount?'<div class="amount">'+data.amount+' OMR</div>':'')+'<div class="status-sub" style="margin-top:8px;font-size:11px;color:#aaa;">Submitted at '+data.time+'</div></div>';
}
document.getElementById('riderId').addEventListener('keypress', e => { if(e.key==='Enter') checkStatus(); });
</script></body></html>`);
});

app.get('/status/check', (req, res) => {
  const { id } = req.query;
  const today = new Date(new Date().getTime() + 4*60*60*1000).toISOString().slice(0,10);
  const sub = submissions.find(s => s.rider_id === id && s.date === today);
  if (!sub) return res.json({ found: false });
  return res.json({ found: true, status: sub.status, amount: sub.amount || null,
    time: new Date(new Date(sub.submitted_at).getTime() + 4*60*60*1000).toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit',hour12:true}),
    reason: sub.rejection_reason || null });
});

app.get('/check-submitted', (req, res) => {
  const { id } = req.query;
  const today = new Date(new Date().getTime() + 4*60*60*1000).toISOString().slice(0,10);
  const sub = submissions.find(s => s.rider_id === id && s.date === today);
  res.json({ submitted: !!sub, status: sub ? sub.status : null });
});

app.get('/fuel', (req, res) => {
  res.send(`<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fuel Allowance</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:1rem;}
  .card{background:#fff;border-radius:16px;padding:1.5rem;width:100%;max-width:420px;box-shadow:0 2px 16px rgba(0,0,0,0.08);}
  h1{font-size:20px;font-weight:600;margin-bottom:4px;}
  .sub{font-size:13px;color:#888;margin-bottom:1.5rem;}
  label{font-size:13px;font-weight:500;color:#444;display:block;margin-bottom:6px;}
  input{width:100%;padding:12px;border:1px solid #ddd;border-radius:10px;font-size:16px;margin-bottom:1rem;background:#fafafa;}
  button{width:100%;padding:12px;background:#1a73e8;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;}
  .msg{padding:12px;border-radius:10px;font-size:13px;text-align:center;margin-top:1rem;}
  .msg.err{background:#fce8e6;color:#c62828;}
  .spinner{display:none;text-align:center;padding:1rem;color:#888;font-size:13px;}
  .results{margin-top:1.25rem;}
  .rider-name{font-size:18px;font-weight:600;margin-bottom:1rem;color:#1a73e8;}
  .week-card{background:#f9f9f9;border:1px solid #eee;border-radius:12px;padding:1rem;margin-bottom:0.75rem;display:flex;justify-content:space-between;align-items:center;}
  .week-label{font-size:13px;color:#666;}
  .week-amount{font-size:20px;font-weight:700;color:#1e7e34;}
  .week-amount.zero{color:#aaa;font-size:16px;}
  .total-card{background:#1a73e8;border-radius:12px;padding:1rem;text-align:center;margin-top:1rem;}
  .total-label{font-size:13px;color:rgba(255,255,255,0.8);}
  .total-amount{font-size:28px;font-weight:700;color:#fff;margin-top:4px;}
</style></head><body>
<div class="card">
  <h1>⛽ Fuel Allowance</h1>
  <p class="sub">Enter your ID to check your fuel balance</p>
  <div id="formWrap">
    <label>Your ID number</label>
    <input type="number" id="riderIdInput" placeholder="Enter your rider ID" inputmode="numeric">
    <button onclick="checkFuel()">Check my fuel</button>
  </div>
  <div class="spinner" id="spinner">⏳ Loading...</div>
  <div id="msg"></div>
  <div id="results" class="results" style="display:none;"></div>
</div>
<script>
async function checkFuel() {
  const id = document.getElementById('riderIdInput').value.trim();
  if (!id) { showMsg('Please enter your ID number.'); return; }
  document.getElementById('spinner').style.display = 'block';
  document.getElementById('msg').innerHTML = '';
  document.getElementById('results').style.display = 'none';
  try {
    const res = await fetch('/fuel/check?id=' + encodeURIComponent(id));
    const data = await res.json();
    document.getElementById('spinner').style.display = 'none';
    if (!data.ok) { showMsg(data.error || 'No fuel data found for this ID.'); return; }
    renderResults(data);
  } catch(e) {
    document.getElementById('spinner').style.display = 'none';
    showMsg('Network error. Please try again.');
  }
}
function renderResults(data) {
  const total = data.weeks.reduce((s, w) => s + (parseFloat(w.amount) || 0), 0);
  const weeksHtml = data.weeks.map(w => \`
    <div class="week-card">
      <div><div class="week-label">📅 \${w.week}</div></div>
      <div class="\${parseFloat(w.amount) > 0 ? 'week-amount' : 'week-amount zero'}">\${parseFloat(w.amount) > 0 ? w.amount + ' OMR' : '—'}</div>
    </div>
  \`).join('');
  document.getElementById('results').innerHTML = \`
    <div class="rider-name">👤 \${data.rider_name}</div>
    \${weeksHtml}
    \${data.weeks.length > 1 ? \`<div class="total-card"><div class="total-label">Total fuel allowance</div><div class="total-amount">\${total.toFixed(3)} OMR</div></div>\` : ''}
  \`;
  document.getElementById('results').style.display = 'block';
}
function showMsg(text) {
  document.getElementById('msg').innerHTML = '<div class="msg err">' + text + '</div>';
}
document.getElementById('riderIdInput').addEventListener('keypress', e => { if (e.key === 'Enter') checkFuel(); });
</script>
</body></html>`);
});

app.get('/fuel/check', async (req, res) => {
  const { id } = req.query;
  if (!id) return res.json({ ok: false, error: 'No ID provided.' });
  const data = await getFuelData(id.trim());
  if (!data || data.length === 0) return res.json({ ok: false, error: 'No fuel records found for this ID. Please check with your supervisor.' });
  const riderName = data[0].rider_name || 'Rider';
  return res.json({ ok: true, rider_name: riderName, weeks: data.map(r => ({ week: r.week, amount: r.amount })) });
});


const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`COD Tracker running on port ${PORT}`);
  getSheetAuth().then(async () => {
    await loadRidersFromSheet();
    await loadTodaySubmissionsFromSheet();
  }).catch(e => console.error('Startup sheet error:', e.message));
});
