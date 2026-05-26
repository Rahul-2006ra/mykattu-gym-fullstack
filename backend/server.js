// =====================================================================
// MYKATTU GYM — Backend Server
// =====================================================================
// Secrets are loaded from .env file — never hardcoded here.
// This file is safe to share / push to GitHub.
// =====================================================================

require('dotenv').config();   // loads .env file automatically

const express    = require('express');
const cors       = require('cors');
const nodemailer = require('nodemailer');
const ExcelJS    = require('exceljs');
const path       = require('path');
const fs         = require('fs');
const https      = require('https');

const app  = express();
const PORT = process.env.PORT || 3001;

// ── SECRETS (loaded from .env — never hardcoded) ──────────────────────
const OWNER_EMAIL       = process.env.OWNER_EMAIL;
const GMAIL_USER        = process.env.GMAIL_USER;
const GMAIL_APP_PASS    = process.env.GMAIL_APP_PASS;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const ANTHROPIC_MODEL   = process.env.ANTHROPIC_MODEL || 'claude-haiku-4-5';
// ─────────────────────────────────────────────────────────────────────

const EXCEL_FILE = path.join(__dirname, 'mykattu_leads.xlsx');
const EMAIL_TIMEOUT_MS = Number(process.env.EMAIL_TIMEOUT_MS || 12000);
const AI_TIMEOUT_MS    = Number(process.env.AI_TIMEOUT_MS || 20000);

// ── WARN if running on Render free tier (no persistent disk) ──────────
if (process.env.RENDER) {
  console.warn('⚠️  Running on Render — Excel file will be lost on redeploy!');
  console.warn('   → Consider adding a Render Disk or exporting leads via email only.');
}

// ── CORS — allows your GitHub Pages frontend to talk to this backend ──
app.use(cors({
  origin: [
    'https://rahul-2006ra.github.io',   // your live GitHub Pages site
    'http://localhost:3000',             // for local development
    'http://localhost:5500',             // for VS Code Live Server
    'http://127.0.0.1:5500',            // for VS Code Live Server (alternate)
  ],
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
}));
// ─────────────────────────────────────────────────────────────────────

app.use(express.json());
app.use(express.static(__dirname));

// ── STARTUP ENV VALIDATION ────────────────────────────────────────────
const missingVars = [];
if (!OWNER_EMAIL)    missingVars.push('OWNER_EMAIL');
if (!GMAIL_USER)     missingVars.push('GMAIL_USER');
if (!GMAIL_APP_PASS) missingVars.push('GMAIL_APP_PASS');
if (missingVars.length) {
  console.error('❌ Missing environment variables:', missingVars.join(', '));
  console.error('   → Set these in Render Dashboard → Environment tab');
}

// ── NODEMAILER SETUP ──────────────────────────────────────────────────
const emailConfigured = !!(OWNER_EMAIL && GMAIL_USER && GMAIL_APP_PASS);
const transporter = emailConfigured ? nodemailer.createTransport({
  host:   'smtp.gmail.com',
  port:   465,
  secure: true,
  connectionTimeout: EMAIL_TIMEOUT_MS,
  greetingTimeout:   EMAIL_TIMEOUT_MS,
  socketTimeout:     EMAIL_TIMEOUT_MS,
  auth: {
    user: GMAIL_USER,
    pass: GMAIL_APP_PASS,
  },
}) : null;

if (transporter) {
transporter.verify((error) => {
  if (error) {
    console.error('❌ Email setup error:', error.message);
    console.error('   → Check GMAIL_USER and GMAIL_APP_PASS in Render Environment tab');
  } else {
    console.log('✅ Email ready — Gmail connected');
  }
});
}

// ── EXCEL HELPER ──────────────────────────────────────────────────────
const HEADERS = [
  'S.No', 'Date', 'Time', 'First Name', 'Last Name',
  'Phone', 'Email', 'Fitness Goal', 'Preferred Plan', 'Message', 'Source'
];

function withTimeout(promise, ms, label) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

function getLatestUserMessage(messages = []) {
  const lastUser = [...messages].reverse().find(m => m && m.role === 'user');
  return String(lastUser?.content || '').toLowerCase();
}

function localCoachReply(messages) {
  const msg = getLatestUserMessage(messages);

  if (msg.includes('time') || msg.includes('timing') || msg.includes('open') || msg.includes('hour')) {
    return "Bro, MYKATTU GYM is open Mon-Fri 5 AM-11 PM, Sat 6 AM-10 PM, and Sun 7 AM-8 PM. Come for a free trial and we will map your plan, champ!";
  }
  if (msg.includes('price') || msg.includes('plan') || msg.includes('membership') || msg.includes('fee')) {
    return "Boss, our plans are Warrior Rs.999/mo, Elite Rs.1999/mo, Legend Rs.3499/mo, and Annual Rs.29999/yr. Elite is the best value if you want AI coaching, scans, classes, and nutrition support.";
  }
  if (msg.includes('address') || msg.includes('location') || msg.includes('where')) {
    return "Champ, we are at State Highway 39, opp. St. Anne's High School, T.B Cross, Hesaraghatta, Bengaluru 560088. Call or WhatsApp +91 98765 43210 and we will guide you.";
  }
  if (msg.includes('fat') || msg.includes('weight loss') || msg.includes('lose weight')) {
    return "Bro, start with 3 days strength training, 2 days cardio, 8-10k steps daily, and a small calorie deficit with high protein. Book a free trial and we will build your exact fat-loss plan.";
  }
  if (msg.includes('muscle') || msg.includes('bulk') || msg.includes('strength')) {
    return "Boss, focus on progressive overload, compound lifts, 1.6-2.2g protein per kg body weight, and proper sleep. Our Elite plan is great for muscle gain with AI tracking and trainer support.";
  }
  if (msg.includes('diet') || msg.includes('protein') || msg.includes('meal')) {
    return "Champ, keep every meal protein-first: eggs, chicken, paneer, dal, curd, or whey, then add rice/roti and vegetables based on your goal. Tell me your weight and goal and I will suggest a simple split.";
  }

  return "Bro, I can help with workouts, diet, fat loss, muscle gain, supplements, timings, pricing, and free trial booking. Ask me your goal and I will guide you like a coach.";
}

async function appendToExcel(lead) {
  const workbook = new ExcelJS.Workbook();
  let isNewFile = !fs.existsSync(EXCEL_FILE);

  if (!isNewFile) {
    try {
      await workbook.xlsx.readFile(EXCEL_FILE);
    } catch (e) {
      console.warn('⚠️  Could not read Excel file, creating fresh one:', e.message);
      isNewFile = true;
    }
  }

  let sheet = workbook.getWorksheet('Leads');
  if (!sheet) {
    sheet = workbook.addWorksheet('Leads');
    const hr = sheet.addRow(HEADERS);
    hr.eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0099FF' } };
      cell.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.border    = { bottom: { style: 'thin', color: { argb: 'FF0077CC' } } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    hr.height = 22;
    [6, 12, 10, 14, 14, 16, 28, 20, 20, 36, 14].forEach((w, i) => {
      sheet.getColumn(i + 1).width = w;
    });
  }

  // FIX: rowCount includes the header row (row 1), so data rows start at rowCount.
  // S.No = rowCount - 1 gives correct 1-based serial number for leads only.
  const dataRowCount = sheet.rowCount; // header is row 1, first lead will be row 2
  const sn = dataRowCount; // S.No = existing rows (header + previous leads), so next lead = dataRowCount

  const now     = new Date();
  const dateStr = now.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  const timeStr = now.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', hour12: true });

  const dr = sheet.addRow([
    sn, dateStr, timeStr,
    lead.fname   || '',
    lead.lname   || '',
    lead.phone   || '',
    lead.email   || '',
    lead.goal    || '',
    lead.plan    || '',
    lead.message || '',
    lead.source  || 'Website Form',
  ]);

  const isEven = (sn % 2 === 0);
  dr.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFF0F8FF' : 'FFFFFFFF' } };
    cell.border    = { bottom: { style: 'hair', color: { argb: 'FFE2E8F0' } } };
    cell.alignment = { vertical: 'middle', wrapText: true };
  });
  dr.height = 18;

  await workbook.xlsx.writeFile(EXCEL_FILE);
  return sn;
}

// ── OWNER ALERT EMAIL ─────────────────────────────────────────────────
async function sendOwnerEmail(lead, sn) {
  if (!transporter || !OWNER_EMAIL) return { skipped: true };
  await transporter.sendMail({
    from:    `"MYKATTU GYM" <${GMAIL_USER}>`,
    to:      OWNER_EMAIL,
    subject: `🔥 New Lead #${sn}: ${lead.fname} ${lead.lname} — ${lead.plan || 'Interested'}`,
    html: `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:linear-gradient(135deg,#0099ff,#0077cc);padding:24px 32px;border-radius:12px 12px 0 0">
        <h1 style="color:#fff;margin:0;font-size:20px">🏋️ New Lead — MYKATTU GYM</h1>
        <p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px">Lead #${sn} • ${new Date().toLocaleString('en-IN')}</p>
      </div>
      <div style="background:#fff;padding:24px 32px;border:1px solid #e2e8f0">
        <table style="width:100%;border-collapse:collapse;font-size:14px">
          <tr style="background:#eff6ff"><td style="padding:10px 14px;font-weight:bold;color:#0099ff;width:35%">Name</td><td style="padding:10px 14px">${lead.fname} ${lead.lname}</td></tr>
          <tr><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Phone</td><td style="padding:10px 14px">${lead.phone}</td></tr>
          <tr style="background:#eff6ff"><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Email</td><td style="padding:10px 14px">${lead.email || '—'}</td></tr>
          <tr><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Goal</td><td style="padding:10px 14px">${lead.goal || '—'}</td></tr>
          <tr style="background:#eff6ff"><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Plan</td><td style="padding:10px 14px">${lead.plan || '—'}</td></tr>
          <tr><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Message</td><td style="padding:10px 14px">${lead.message || '—'}</td></tr>
          <tr style="background:#eff6ff"><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Source</td><td style="padding:10px 14px">${lead.source || 'Website Form'}</td></tr>
        </table>
      </div>
      <div style="background:#eff6ff;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center">
        <p style="margin:0;font-size:13px;color:#334155">⚡ Contact this lead within <strong>1 hour</strong> for best results. Lead saved to Excel.</p>
      </div>
    </div>`,
  });
}

// ── USER THANK-YOU EMAIL ──────────────────────────────────────────────
async function sendUserEmail(lead) {
  if (!transporter) return { skipped: true };
  if (!lead.email) return;
  await transporter.sendMail({
    from:    `"MYKATTU GYM" <${GMAIL_USER}>`,
    to:      lead.email,
    subject: `🏋️ Thanks for your interest in MYKATTU GYM — We'll contact you within 12 hrs!`,
    html: `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:linear-gradient(135deg,#0099ff,#0077cc);padding:36px 32px;text-align:center;border-radius:12px 12px 0 0">
        <div style="font-size:48px">🏋️</div>
        <h1 style="color:#fff;margin:8px 0 0;font-size:24px">Welcome to MYKATTU!</h1>
        <p style="color:rgba(255,255,255,0.85);margin:6px 0 0">Bengaluru's #1 AI-Powered Gym</p>
      </div>
      <div style="background:#fff;padding:32px;border:1px solid #e2e8f0">
        <p style="font-size:16px;color:#0f172a">Hi <strong>${lead.fname}</strong>,</p>
        <p style="font-size:15px;color:#334155;line-height:1.7">
          Thank you for showing interest in MYKATTU GYM! 🙌<br><br>
          Our team will connect with you within <strong>12 hours</strong> to discuss your fitness goals and kick-start your transformation.
        </p>
        <div style="background:#eff6ff;border-radius:10px;padding:20px 24px;margin:20px 0">
          <p style="font-size:13px;font-weight:bold;color:#0099ff;margin:0 0 10px;text-transform:uppercase;letter-spacing:1px">Your Details</p>
          <p style="font-size:14px;color:#334155;margin:4px 0">📋 <strong>Goal:</strong> ${lead.goal || 'General Fitness'}</p>
          <p style="font-size:14px;color:#334155;margin:4px 0">💎 <strong>Plan:</strong> ${lead.plan || 'To be discussed'}</p>
        </div>
        <p style="font-size:14px;color:#334155;line-height:1.8">
          📍 State Highway 39, opp. St. Anne's High School, T.B Cross, Hesaraghatta, Bengaluru 560088<br>
          📞 +91 98765 43210 (WhatsApp available)<br>
          🕐 Mon–Fri: 5 AM–11 PM | Sat: 6 AM–10 PM | Sun: 7 AM–8 PM
        </p>
        <div style="text-align:center;margin-top:24px">
          <a href="https://wa.me/919876543210" style="display:inline-block;background:linear-gradient(135deg,#0099ff,#0077cc);color:#fff;text-decoration:none;padding:14px 32px;border-radius:10px;font-weight:bold">💬 Chat on WhatsApp</a>
        </div>
      </div>
      <div style="background:#eff6ff;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center">
        <p style="margin:0;font-size:13px;color:#334155">Follow us: <strong>@MYKATTUGym</strong> on Instagram & YouTube</p>
        <p style="margin:4px 0 0;font-size:12px;color:#94a3b8">© MYKATTU GYM, Hesaraghatta, Bengaluru</p>
      </div>
    </div>`,
  });
}

// ── API: SUBMIT LEAD FORM ─────────────────────────────────────────────
app.post('/api/lead', async (req, res) => {
  try {
    const lead = req.body;
    if (!lead.fname || !lead.phone) {
      return res.status(400).json({ success: false, message: 'Name and phone are required.' });
    }

    // Save to Excel (non-fatal — don't fail the whole request if Excel fails on Render)
    let sn = '?';
    try {
      sn = await appendToExcel(lead);
      console.log(`✅ Lead #${sn} saved to Excel: ${lead.fname} ${lead.lname} (${lead.phone})`);
    } catch (excelErr) {
      console.error('⚠️  Excel save failed (continuing anyway):', excelErr.message);
    }

    // Send emails with a hard timeout so the website never hangs on SMTP.
    const [ownerR, userR] = await Promise.allSettled([
      withTimeout(sendOwnerEmail(lead, sn), EMAIL_TIMEOUT_MS, 'Owner email'),
      withTimeout(sendUserEmail(lead), EMAIL_TIMEOUT_MS, 'User email'),
    ]);

    if (ownerR.status === 'rejected') console.error('❌ Owner email failed:', ownerR.reason?.message);
    else console.log('✅ Owner email sent to', OWNER_EMAIL);

    if (userR.status === 'rejected') console.error('❌ User email failed:', userR.reason?.message);
    else if (lead.email) console.log('✅ User email sent to', lead.email);

    // If Excel failed and both emails failed, return an error so the frontend knows.
    // Otherwise the lead is captured and the website should show success.
    if (sn === '?' && ownerR.status === 'rejected' && userR.status === 'rejected') {
      return res.status(500).json({
        success: false,
        message: 'Could not send emails. Please check Gmail credentials in Render Environment tab.',
      });
    }

    res.json({
      success: true,
      message: emailConfigured ? 'Lead saved!' : 'Lead saved locally. Email is not configured.',
      serialNo: sn,
      emailSent: (ownerR.status === 'fulfilled' && !ownerR.value?.skipped) ||
                 (userR.status === 'fulfilled' && !userR.value?.skipped),
    });

  } catch (err) {
    console.error('❌ Lead route error:', err);
    res.status(500).json({ success: false, message: 'Server error. Please try again.' });
  }
});

// ── API: CLAUDE AI CHAT PROXY ─────────────────────────────────────────
app.post('/api/chat', (req, res) => {
  try {
    const { messages } = req.body;
    if (!messages || !Array.isArray(messages)) {
      return res.status(400).json({ error: 'messages array required' });
    }

    if (!ANTHROPIC_API_KEY || ANTHROPIC_API_KEY === 'YOUR_ANTHROPIC_API_KEY') {
      return res.json({ reply: localCoachReply(messages), fallback: true });
    }

    const bodyData = JSON.stringify({
      model:      ANTHROPIC_MODEL,
      max_tokens: 500,
      system: "You are MYKATTU AI Coach for MYKATTU GYM in Hesaraghatta, Bengaluru. " +
              "Personality: friendly energetic gym coach, call user 'bro'/'boss'/'champ', keep replies short (2-5 sentences), use fitness emojis. " +
              "You are expert in: workouts, fat loss, muscle building, diet plans, macros, protein, supplements (whey/creatine/BCAA), " +
              "training programs (PPL/full body/5x5), HIIT, injury prevention, recovery, motivation. " +
              "MYKATTU plans: Silver ₹999/mo (basic), Gold ₹1999/mo (all equipment+classes), Platinum ₹3499/mo (PT+diet+all access), Annual ₹29999/yr. " +
              "Address: State Highway 39, opp. St. Anne's High School, T.B Cross, Hesaraghatta, Bengaluru 560088. " +
              "Phone: +91 98765 43210. Timings: Mon-Fri 5AM-11PM, Sat 6AM-10PM, Sun 7AM-8PM. " +
              "If user wants to join, suggest booking a free trial. Answer all fitness questions confidently.",
      messages: messages,
    });

    const options = {
      hostname: 'api.anthropic.com',
      path:     '/v1/messages',
      method:   'POST',
      headers: {
        'Content-Type':      'application/json',
        'x-api-key':         ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'Content-Length':    Buffer.byteLength(bodyData),
      },
    };

    const apiReq = https.request(options, (apiRes) => {
      let data = '';
      apiRes.on('data', chunk => { data += chunk; });
      apiRes.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.content && parsed.content[0]?.text) {
            res.json({ reply: parsed.content[0].text });
          } else {
            console.error('Anthropic response:', data);
            return res.json({ reply: localCoachReply(messages), fallback: true });
          }
        } catch (e) {
          console.error('Anthropic parse error:', e.message, data);
          return res.json({ reply: localCoachReply(messages), fallback: true });
        }
      });
    });

    apiReq.on('error', (e) => {
      console.error('Anthropic request error:', e.message);
      if (!res.headersSent) return res.json({ reply: localCoachReply(messages), fallback: true });
    });

    apiReq.setTimeout(AI_TIMEOUT_MS, () => {
      apiReq.destroy(new Error(`Anthropic request timed out after ${AI_TIMEOUT_MS}ms`));
    });

    apiReq.write(bodyData);
    apiReq.end();

  } catch (err) {
    console.error('Chat route error:', err);
    return res.json({ reply: localCoachReply(req.body?.messages), fallback: true });
  }
});

// ── HEALTH CHECK ──────────────────────────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    time: new Date().toISOString(),
    emailConfigured,
    ownerEmail: OWNER_EMAIL || 'NOT SET',
  });
});

// ── START SERVER ──────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n✅ MYKATTU Backend running → http://localhost:${PORT}`);
  console.log(`📊 Excel file: ${EXCEL_FILE}`);
  console.log(`📧 Owner email: ${OWNER_EMAIL || '⚠️  NOT SET'}`);
  console.log(`📬 Gmail user: ${GMAIL_USER || '⚠️  NOT SET'}`);
  console.log(`🔑 Gmail pass: ${GMAIL_APP_PASS ? '✅ Set' : '⚠️  NOT SET'}`);
  console.log(`🤖 AI Chat: ${ANTHROPIC_API_KEY && ANTHROPIC_API_KEY !== 'YOUR_ANTHROPIC_API_KEY' ? 'Claude ready ✨' : '⚠️  Add ANTHROPIC_API_KEY to .env'}\n`);
});
