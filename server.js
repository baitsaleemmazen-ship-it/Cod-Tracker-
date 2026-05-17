require('dotenv').config();
const express = require('express');
const multer = require('multer');
const session = require('express-session');
const { google } = require('googleapis');
const Anthropic = require('@anthropic-ai/sdk');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'cod-tracker-secret-2026',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 }
}));

// ─── Riders store (in-memory, persisted to riders.json) ───────────────────────
const RIDERS_FILE = path.join(__dirname, 'riders.json');
let riders = [];
if (fs.existsSync(RIDERS_FILE)) {
  riders = JSON.parse(fs.readFileSync(RIDERS_FILE));
} else {
  riders = [
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
  fs.writeFileSync(RIDERS_FILE, JSON.stringify(riders, null, 2));
}
function saveRiders() { fs.writeFileSync(RIDERS_FILE, JSON.stringify(riders, null, 2)); }

// ─── Submissions store ────────────────────────────────────────────────────────
const SUBS_FILE = path.join(__dirname, 'submissions.json');
let submissions = [];
if (fs.existsSync(SUBS_FILE)) {
  submissions = JSON.parse(fs.readFileSync(SUBS_FILE));
}
function saveSubmissions() { fs.writeFileSync(SUBS_FILE, JSON.stringify(submissions, null, 2)); }

// ─── Today's submitted IDs (for duplicate detection) ─────────────────────────
function getTodayIds() {
  const today = new Date().toISOString().slice(0, 10);
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
  // date format: DD-MM-YYYY
  const parts = date.split('-'); // YYYY-MM-DD
  const sheetName = `${parts[2]}-${parts[1]}-${parts[0]}`; // DD-MM-YYYY
  try {
    const { sheets, spreadsheetId } = await getSheetAuth();

    // Get existing sheets
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const existing = meta.data.sheets.map(s => s.properties.title);

    if (!existing.includes(sheetName)) {
      // Create new sheet tab
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] }
      });

      // Pre-populate header + all riders in order
      const header = [['Rider Name', 'Rider ID', 'Amount (OMR)', 'Bank', 'Date', 'Time']];
      const riderRows = riders.map(r => [r.name, r.id, '', '', '', '']);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A1`,
        valueInputOption: 'RAW',
        requestBody: { values: [...header, ...riderRows] }
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

    // Get all rows to find rider's row
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:B`
    });
    const rows = res.data.values || [];

    // Find rider row by ID (column B)
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i] && rows[i][1] === submission.rider_id) {
        rowIndex = i + 1; // 1-indexed for Sheets API
        break;
      }
    }

    if (rowIndex === -1) {
      // Rider not in sheet — append at end
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: `${sheetName}!A:F`,
        valueInputOption: 'RAW',
        requestBody: { values: [[submission.rider_name, submission.rider_id, submission.amount || '', submission.bank || '', submission.date, new Date(submission.submitted_at).toLocaleTimeString()]] }
      });
    } else {
      // Fill in the rider's pre-existing row
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!C${rowIndex}:F${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: { values: [[submission.amount || '', submission.bank || '', submission.date, new Date(submission.submitted_at).toLocaleTimeString()]] }
      });
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
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>COD Receipt Submission</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:1rem;}
  .card{background:#fff;border-radius:16px;padding:1.5rem;width:100%;max-width:420px;box-shadow:0 2px 16px rgba(0,0,0,0.08);}
  h1{font-size:20px;font-weight:600;margin-bottom:4px;}
  .sub{font-size:13px;color:#888;margin-bottom:1.5rem;}
  label{font-size:13px;font-weight:500;color:#444;display:block;margin-bottom:4px;}
  select,input{width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:10px;font-size:14px;margin-bottom:1rem;background:#fafafa;}
  .upload-zone{border:2px dashed #ddd;border-radius:12px;padding:2rem;text-align:center;cursor:pointer;margin-bottom:1rem;transition:background 0.15s;}
  .upload-zone:hover{background:#f9f9f9;}
  .upload-zone input{display:none;}
  .preview{max-width:100%;max-height:180px;border-radius:8px;margin:0.5rem auto;display:block;}
  button{width:100%;padding:12px;background:#1a73e8;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;}
  button:disabled{background:#aaa;}
  .msg{padding:12px;border-radius:10px;font-size:13px;text-align:center;margin-top:1rem;}
  .msg.ok{background:#e6f4ea;color:#1e7e34;}
  .msg.err{background:#fce8e6;color:#c62828;}
  .spinner{display:none;text-align:center;padding:1rem;color:#888;font-size:13px;}
</style></head><body>
<div class="card">
  <h1>📦 COD Receipt</h1>
  <p class="sub">Submit your daily bank receipt below</p>
  <form id="form" enctype="multipart/form-data">
    <label>Your name</label>
    <select name="rider_id" required id="riderSel">
      <option value="">— select your name —</option>
      ${riders.map(r => `<option value="${r.id}">${r.name}</option>`).join('')}
    </select>
    <label>Bank receipt photo</label>
    <input type="file" id="imgInput" name="receipt" accept="image/*" onchange="previewFile(this)" style="display:none">
    <input type="file" id="imgCamera" name="receipt" accept="image/*" capture="environment" onchange="previewFile(this)" style="display:none">
    <div style="display:flex;gap:8px;margin-bottom:1rem;">
      <button type="button" onclick="document.getElementById('imgCamera').click()" style="flex:1;padding:14px;background:#1a73e8;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;">📷 Take Photo</button>
      <button type="button" onclick="document.getElementById('imgInput').click()" style="flex:1;padding:14px;background:#fff;color:#1a73e8;border:2px solid #1a73e8;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;">🖼 Gallery</button>
    </div>
    <div id="imgPreviewWrap" style="display:none;text-align:center;margin-bottom:1rem;">
      <img id="imgPreview" class="preview" style="max-width:100%;max-height:180px;border-radius:8px;" alt="preview">
      <p style="font-size:12px;color:#888;margin-top:4px;" id="imgName"></p>
    </div>
    <button type="submit" id="submitBtn">Submit Receipt</button>
  </form>
  <div class="spinner" id="spinner">⏳ Reading your receipt...</div>
  <div id="msg"></div>
</div>
<script>
function previewFile(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    document.getElementById('imgPreview').src = e.target.result;
    document.getElementById('imgPreview').style.display = 'block';
    document.getElementById('uploadHint').style.display = 'none';
  };
  reader.readAsDataURL(file);
}
document.getElementById('form').onsubmit = async e => {
  e.preventDefault();
  const riderId = document.getElementById('riderSel').value;
  const file = window._selectedFile || document.getElementById('imgInput').files[0] || document.getElementById('imgCamera').files[0];
  if (!riderId) { showMsg('Please select your name.', 'err'); return; }
  if (!file) { showMsg('Please take a photo or choose from gallery.', 'err'); return; }
  document.getElementById('submitBtn').disabled = true;
  document.getElementById('spinner').style.display = 'block';
  document.getElementById('msg').innerHTML = '';
  const fd = new FormData();
  fd.append('rider_id', riderId);
  fd.append('receipt', file);
  try {
    const res = await fetch('/submit', { method: 'POST', body: fd });
    const data = await res.json();
    document.getElementById('spinner').style.display = 'none';
    if (data.ok) {
      document.getElementById('form').style.display = 'none';
      showMsg('✅ Receipt submitted successfully! Thank you.', 'ok');
    } else {
      showMsg('❌ ' + (data.error || 'Error. Please try again.'), 'err');
      document.getElementById('submitBtn').disabled = false;
    }
  } catch(err) {
    document.getElementById('spinner').style.display = 'none';
    showMsg('❌ Network error. Please try again.', 'err');
    document.getElementById('submitBtn').disabled = false;
  }
};
function showMsg(text, type) {
  document.getElementById('msg').innerHTML = '<div class="msg ' + type + '">' + text + '</div>';
}
</script>
</body></html>`);
});

// ─── Submit endpoint ──────────────────────────────────────────────────────────
app.post('/submit', upload.single('receipt'), async (req, res) => {
  try {
    const { rider_id } = req.body;
    const rider = riders.find(r => r.id === rider_id);
    if (!rider) return res.json({ ok: false, error: 'Rider not found.' });
    if (!req.file) return res.json({ ok: false, error: 'No image uploaded.' });

    const today = new Date().toISOString().slice(0, 10);
    const todayIds = getTodayIds();

    // Check duplicate
    const isDuplicate = todayIds.has(rider_id);

    // AI analysis
    const b64 = req.file.buffer.toString('base64');
    const mediaType = req.file.mimetype || 'image/jpeg';

    let aiResult = {};
    try {
      const response = await anthropic.messages.create({
        model: 'claude-sonnet-4-5',
        max_tokens: 500,
        messages: [{
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: mediaType, data: b64 } },
            { type: 'text', text: `This is a Bank Muscat (or similar Omani bank) transfer receipt. It has a decorative swirl/wave background pattern — ignore the background and read only the text fields.

Extract these specific fields:
- "Amount:" field → shows like "OMR  11.220" or "OMR  45.500" → extract the number only (e.g. 11.220)
- "Remarks:" field → contains the rider ID like "Sal 16 may 2026 id 1397838" → extract the 7-digit number after "id "
- "Transaction Date and Time:" field → extract the date part
- "Beneficiary Bank:" field → extract bank name

Return ONLY this JSON, no markdown, no explanation:
{"amount":<number e.g. 11.220>,"currency":"OMR","date":"<YYYY-MM-DD>","detected_id":"<7-digit id from Remarks>","bank_name":"<bank name>","is_legit_receipt":true}` }
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

    // Fraud checks — only flag for real fraud, not AI reading issues
    const flags = [];
    let status = 'approved';

    // Real fraud: ID on receipt doesn't match the rider who submitted
    if (aiResult.detected_id && aiResult.detected_id !== rider_id && riderMap && !Object.values(riderMap || {}).includes(aiResult.detected_id)) {
      flags.push(`ID on receipt (${aiResult.detected_id}) doesn't match rider ID (${rider_id})`);
      status = 'flagged';
    }

    // Real fraud: same rider already submitted today
    if (isDuplicate) {
      flags.push('This rider already submitted today — possible duplicate');
      status = 'flagged';
    }

    // Soft flag only (still approved): amount not detected — admin enters manually
    const needsAmount = !aiResult.amount;

    const submission = {
      id: Date.now().toString(),
      rider_id,
      rider_name: rider.name,
      amount: aiResult.amount || null,
      currency: aiResult.currency || 'OMR',
      bank: aiResult.bank_name || 'Bank',
      date: today,
      submitted_at: new Date().toISOString(),
      status,
      flags,
      needs_amount: needsAmount,
      detected_id: aiResult.detected_id || null,
      image_b64: b64,
      image_type: mediaType
    };

    submissions.push(submission);
    saveSubmissions();

    // Auto-write to sheet if approved and amount known
    if (status === 'approved' && aiResult.amount) {
      await fillRiderRow(submission);
    }

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.json({ ok: false, error: 'Server error.' });
  }
});

// ─── Receipt image endpoint ───────────────────────────────────────────────────
app.get('/receipt/:id', requireAdmin, (req, res) => {
  const sub = submissions.find(s => s.id === req.params.id);
  if (!sub || !sub.image_b64) return res.status(404).send('Not found');
  const buf = Buffer.from(sub.image_b64, 'base64');
  res.setHeader('Content-Type', sub.image_type || 'image/jpeg');
  res.send(buf);
});

// ══════════════════════════════════════════════════════════════════════════════
// ADMIN — login
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

// ══════════════════════════════════════════════════════════════════════════════
// ADMIN — dashboard
// ══════════════════════════════════════════════════════════════════════════════
app.get('/admin', requireAdmin, (req, res) => {
  const today = new Date().toISOString().slice(0, 10);
  const todaySubs = submissions.filter(s => s.date === today);
  const approved = todaySubs.filter(s => s.status === 'approved');
  const flagged = todaySubs.filter(s => s.status === 'flagged');
  const totalAmt = approved.reduce((s, e) => s + (parseFloat(e.amount) || 0), 0);

  const rows = todaySubs.sort((a, b) => b.submitted_at.localeCompare(a.submitted_at)).map(s => `
    <tr>
      <td>
        ${s.image_b64 ? `<a href="/receipt/${s.id}" target="_blank"><img src="/receipt/${s.id}" style="width:56px;height:56px;object-fit:cover;border-radius:8px;border:1px solid #eee;display:block;" alt="receipt"></a>` : '<div style="width:56px;height:56px;background:#f5f5f5;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;">📄</div>'}
      </td>
      <td style="font-weight:500;">${s.rider_name}</td>
      <td style="color:#888;">${s.rider_id}</td>
      <td style="font-weight:600;">${s.amount ? s.amount.toLocaleString() + ' OMR' : '<span style="color:#c62828;">Not detected</span>'}</td>
      <td>${new Date(s.submitted_at).toLocaleTimeString()}</td>
      <td>
        <span class="badge ${s.status === 'approved' ? 'ok' : s.status === 'flagged' ? 'flagged' : 'rej'}">
          ${s.status === 'approved' ? '✓ Approved' : s.status === 'flagged' ? '⚠ Flagged' : '✗ Rejected'}
        </span>
          ${s.status === 'flagged' || s.needs_amount ? `
          ${s.flags.length ? `<div style="font-size:11px;color:#c62828;margin-top:4px;">${s.flags.join('<br>')}</div>` : ''}
          ${s.needs_amount && s.status !== 'flagged' ? `<div style="font-size:11px;color:#e65100;margin-top:4px;">⚠ Amount not detected — enter manually</div>` : ''}
          <form method="POST" action="/admin/approve/${s.id}" style="margin:6px 0 4px;">
            <input name="manual_amount" type="number" step="0.01" placeholder="Enter amount (OMR)" value="${s.amount || ''}" style="width:100%;padding:5px 8px;border:1px solid #ddd;border-radius:6px;font-size:12px;margin-bottom:4px;" ${s.amount ? '' : 'required'}>
            <div style="display:flex;gap:6px;">
              <button class="act-btn ok-btn" type="submit">✓ Approve</button>
              <a href="/receipt/${s.id}" target="_blank" class="act-btn" style="background:#e8f0fe;color:#1a73e8;text-decoration:none;padding:4px 10px;border-radius:6px;font-size:12px;">👁 View</a>
            </div>
          </form>
          ${s.status === 'flagged' ? `<form method="POST" action="/admin/reject/${s.id}" style="margin:0 0 4px;"><button class="act-btn rej-btn">✗ Reject</button></form>` : ''}
          <form method="POST" action="/admin/delete/${s.id}" onsubmit="return confirm('Delete this submission?')" style="margin:0;"><button class="act-btn" style="background:#f5f5f5;color:#888;font-size:11px;">🗑 Delete</button></form>` :
          `<div style="margin-top:4px;display:flex;gap:6px;align-items:center;">
            ${s.image_b64 ? `<a href="/receipt/${s.id}" target="_blank" style="font-size:11px;color:#1a73e8;">View receipt</a>` : ''}
            <form method="POST" action="/admin/delete/${s.id}" onsubmit="return confirm('Delete this submission?')" style="margin:0;"><button class="act-btn" style="background:#f5f5f5;color:#888;font-size:11px;">🗑 Delete</button></form>
          </div>`}
      </td>
    </tr>`).join('');

  res.send(`<!DOCTYPE html><html><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>COD Admin Dashboard</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,sans-serif;background:#f5f5f5;color:#222;}
  .topbar{background:#fff;padding:0.875rem 1.5rem;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #eee;position:sticky;top:0;z-index:10;}
  .topbar h1{font-size:17px;font-weight:600;}
  .topbar-links{display:flex;gap:12px;align-items:center;}
  .topbar-links a{font-size:13px;color:#1a73e8;text-decoration:none;padding:6px 12px;border:1px solid #ddd;border-radius:8px;}
  .content{padding:1.5rem;max-width:1000px;margin:0 auto;}
  .stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:1.5rem;}
  .stat{background:#fff;border-radius:12px;padding:1rem 1.25rem;border:1px solid #eee;}
  .stat-val{font-size:26px;font-weight:600;}
  .stat-lbl{font-size:12px;color:#888;margin-top:2px;}
  table{width:100%;background:#fff;border-radius:12px;border-collapse:collapse;overflow:hidden;border:1px solid #eee;}
  th{text-align:left;font-size:12px;font-weight:600;color:#888;padding:10px 12px;border-bottom:1px solid #eee;text-transform:uppercase;letter-spacing:0.03em;}
  td{padding:10px 12px;font-size:13px;border-bottom:1px solid #f5f5f5;vertical-align:top;}
  tr:last-child td{border-bottom:none;}
  .badge{font-size:11px;font-weight:600;padding:3px 10px;border-radius:999px;display:inline-block;}
  .badge.ok{background:#e6f4ea;color:#1e7e34;}
  .badge.flagged{background:#fce8e6;color:#c62828;}
  .badge.rej{background:#f5f5f5;color:#888;}
  .act-btn{font-size:12px;padding:4px 10px;border-radius:6px;cursor:pointer;border:none;font-weight:500;}
  .ok-btn{background:#e6f4ea;color:#1e7e34;}
  .rej-btn{background:#fce8e6;color:#c62828;}
  .empty{text-align:center;padding:3rem;color:#aaa;font-size:14px;}
  @media(max-width:600px){.stats{grid-template-columns:repeat(2,1fr);}th:nth-child(4),td:nth-child(4){display:none;}}
</style>
<meta http-equiv="refresh" content="30">
</head><body>
<div class="topbar">
  <h1>📦 COD Dashboard</h1>
  <div class="topbar-links">
    <a href="/admin/riders">Manage riders</a>
    <a href="/admin/export">Export CSV</a>
    <a href="/admin/logout">Logout</a>
  </div>
</div>
<div class="content">
  <div class="stats">
    <div class="stat"><div class="stat-val">${todaySubs.length}</div><div class="stat-lbl">Submissions today</div></div>
    <div class="stat"><div class="stat-val" style="color:#1e7e34;">${approved.length}</div><div class="stat-lbl">Approved</div></div>
    <div class="stat"><div class="stat-val" style="color:#c62828;">${flagged.length}</div><div class="stat-lbl">Flagged</div></div>
    <div class="stat"><div class="stat-val" style="font-size:18px;padding-top:4px;">${totalAmt.toLocaleString(undefined,{maximumFractionDigits:2})} OMR</div><div class="stat-lbl">Total COD approved</div></div>
  </div>
  <table>
    <thead><tr><th>Receipt</th><th>Rider</th><th>ID</th><th>Amount</th><th>Time</th><th>Status</th></tr></thead>
    <tbody>${rows || '<tr><td colspan="6" class="empty">No submissions yet today.</td></tr>'}</tbody>
  </table>
</div>
</body></html>`);
});

// ─── Approve / Reject ─────────────────────────────────────────────────────────
app.post('/admin/approve/:id', requireAdmin, async (req, res) => {
  const sub = submissions.find(s => s.id === req.params.id);
  if (sub) {
    sub.status = 'approved';
    // Use manually entered amount if AI didn't detect it
    if (req.body.manual_amount && !sub.amount) {
      sub.amount = parseFloat(req.body.manual_amount);
    } else if (req.body.manual_amount) {
      sub.amount = parseFloat(req.body.manual_amount);
    }
    sub.bank = sub.bank || 'Bank';
    saveSubmissions();
    await fillRiderRow(sub);
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
  if (sub) { sub.status = 'rejected'; saveSubmissions(); }
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

app.post('/admin/riders/add', requireAdmin, (req, res) => {
  const { name, id } = req.body;
  if (name && id && !riders.find(r => r.id === id.trim())) {
    riders.push({ id: id.trim(), name: name.trim() });
    saveRiders();
  }
  res.redirect('/admin/riders');
});

app.post('/admin/riders/delete/:id', requireAdmin, (req, res) => {
  riders = riders.filter(r => r.id !== req.params.id);
  saveRiders();
  res.redirect('/admin/riders');
});

// ─── Export CSV ───────────────────────────────────────────────────────────────
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
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`COD Tracker running on port ${PORT}`));
