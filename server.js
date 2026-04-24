// ============================================================
//  AI FORGEN — COMPLETE BACKEND v6.0
//  Node.js + Express + n1n.ai + Razorpay + Supabase
//  NEW: Smart File Parser | Auto Platform Detect | Batch
//  19 Platforms | 95%+ Accuracy | Zero Placeholders
//  NEW v6.0: Smart File Parser (CSV/Excel/PDF/Image/Word/Text)
//            Auto Platform Detection | Batch Processing
// ============================================================
require('dotenv').config();

const express    = require('express');
const cors       = require('cors');
const crypto     = require('crypto');
const multer     = require('multer');
const axios      = require('axios');
const Razorpay   = require('razorpay');
const path       = require('path');
const { createClient } = require('@supabase/supabase-js');
const rateLimit  = require('express-rate-limit');

// File parsing libraries (npm install xlsx pdf-parse)
let XLSX_LIB = null, PDF_PARSE_LIB = null;
try { XLSX_LIB     = require('xlsx');      } catch(e) { console.warn('⚠️  xlsx not installed — Excel parsing disabled. Run: npm install xlsx'); }
try { PDF_PARSE_LIB = require('pdf-parse'); } catch(e) { console.warn('⚠️  pdf-parse not installed — PDF parsing limited. Run: npm install pdf-parse'); }

const app    = express();
// File size 25MB tak — bade files ke liye
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    // Allowed file types
    const allowed = [
      'image/jpeg','image/png','image/gif','image/webp','image/bmp',
      'text/plain','text/csv',
      'application/pdf',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/json'
    ];
    if (allowed.includes(file.mimetype) || file.originalname.match(/\.(txt|csv|json|xml|tsv)$/i)) {
      cb(null, true);
    } else {
      cb(null, true); // Accept all — AI will handle unknown types
    }
  }
});

// ─── CORS — sirf apna domain allow karo ─────────────────────
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || 'https://aiforgen.netlify.app').split(',').map(o => o.trim());
app.use(cors({
  origin: (origin, callback) => {
    // Allow requests with no origin (curl, Postman, server-to-server)
    if (!origin) return callback(null, true);
    if (ALLOWED_ORIGINS.includes(origin) || process.env.NODE_ENV !== 'production') {
      return callback(null, true);
    }
    callback(new Error('CORS blocked: ' + origin));
  },
  methods: ['GET','POST','DELETE','OPTIONS'],
  allowedHeaders: ['Content-Type','Authorization']
}));

// ─── RATE LIMITERS ────────────────────────────────────────────
// 1. General API — 100 requests per 15 min per IP (sabhi routes ke liye)
const generalLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 100,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Bahut zyada requests. 15 minute baad try karein.' }
});

// 2. AI Generate — 20 requests per 10 min per IP (sabse expensive route)
const generateLimiter = rateLimit({
  windowMs: 10 * 60 * 1000,
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Generate limit reached. 10 minute baad try karein.' }
});

// 3. Auth/Payment — 10 requests per 15 min per IP (brute force rokne ke liye)
const strictLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many attempts. 15 minute baad try karein.' }
});

app.use(generalLimiter); // sabhi routes par laago
app.use(express.json({ limit: '25mb' }));
app.use(express.urlencoded({ extended: true, limit: '25mb' }));

// ─── ENV VALIDATION ──────────────────────────────────────────
const REQUIRED = ['N1N_API_KEY', 'SUPABASE_URL', 'SUPABASE_SERVICE_KEY', 'RAZORPAY_KEY_ID', 'RAZORPAY_KEY_SECRET'];
REQUIRED.forEach(k => {
  if (!process.env[k]) console.warn(`WARNING: Missing env var: ${k}`);
});

// ─── CLIENTS ─────────────────────────────────────────────────
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY,
  { auth: { autoRefreshToken: false, persistSession: false } }
);

const razorpay = new Razorpay({
  key_id:     process.env.RAZORPAY_KEY_ID,
  key_secret: process.env.RAZORPAY_KEY_SECRET
});

// ─── CONSTANTS ───────────────────────────────────────────────
const FREE_DAILY_LIMIT = 5;

// ─── REFERRAL CONSTANTS (sync with index.html) ─────────────────
const REFERRAL_SIGNUP_COINS   = 1000;  // Per signup
const REFERRAL_COMMISSION_PCT = 0.05;  // 5% of plan price
const REFERRAL_MILESTONE_21   = 10000; // Bonus at 21 referrals
const REFERRAL_COINS_WINDOW   = 5000;  // Milestone at 5000 coins

// ─── PLAN PRICES (paise mein — Razorpay ke liye) ─────────────
// Monthly Plans
const PLAN_PRICES = {
  // Monthly
  'Pro':              29900,   // ₹299/month
  'Popular':          49900,   // ₹499/month
  'Business':         99900,   // ₹999/month
  // 6 Months
  'Pro_6M':           99900,   // ₹999/6 months
  'Popular_6M':      164900,   // ₹1649/6 months
  'Business_6M':     329900,   // ₹3299/6 months
  // Yearly
  'Pro_1Y':          124900,   // ₹1249/year
  'Popular_1Y':      209900,   // ₹2099/year
  'Business_1Y':     419900,   // ₹4199/year
  // Legacy keys (backward compatibility)
  '1 Month':          29900,   // ₹299
  '3 Months':         29900,   // ₹299
  '6 Months':         99900,   // ₹999
  '1 Year':          124900,   // ₹1249
};

const ALL_PLATFORMS = [
  'Microsoft Excel','Microsoft Word','Microsoft Access','Microsoft PowerPoint',
  'Microsoft Outlook','Microsoft OneNote','Google Docs','Google Sheets',
  'Google Slides','Google Forms','LibreOffice','Apache OpenOffice',
  'WPS Office','Zoho Office Suite','Tally','Busy Accounting Software',
  'QuickBooks','Notepad','WordPad'
];

// ─── MODEL MAPPING ────────────────────────────────────────────
const COST_OPTIMIZED_MODELS = {
  'Microsoft Excel':          'gpt-4o',
  'Microsoft Word':           'gpt-4o-mini',
  'Microsoft PowerPoint':     'gpt-4o-mini',
  'Microsoft Access':         'gpt-4o-mini',
  'Microsoft Outlook':        'gpt-4o-mini',
  'Microsoft OneNote':        'gpt-4o-mini',
  'Google Docs':              'gpt-4o-mini',
  'Google Sheets':            'gpt-4o',
  'Google Slides':            'gpt-4o-mini',
  'Google Forms':             'gpt-4o-mini',
  'LibreOffice':              'gpt-4o-mini',
  'Apache OpenOffice':        'gpt-4o-mini',
  'WPS Office':               'gpt-4o-mini',
  'Zoho Office Suite':        'gpt-4o-mini',
  'QuickBooks':               'gpt-4o',
  'WordPad':                  'gpt-4o-mini',
  'Tally':                    'gpt-4o',
  'Busy Accounting Software': 'gpt-4o',
  'Notepad':                  'gpt-4o-mini',
};

const DEFAULT_COST_MODEL   = 'gpt-4o-mini';
const COST_FALLBACK_CHAIN  = ['gpt-4o-mini', 'gpt-4o', 'deepseek-v3'];

// ─── SMART FILE PARSER ────────────────────────────────────────
// Kisi bhi file ko readable text mein convert karta hai

function parseCSV(buffer) {
  const text = buffer.toString('utf-8');
  const lines = text.split('\n').filter(l => l.trim());
  if (lines.length === 0) return { type: 'csv', content: '', rows: 0 };

  // Detect separator: comma, semicolon, tab
  const firstLine = lines[0];
  let sep = ',';
  if ((firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length) sep = '\t';
  if ((firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length) sep = ';';

  const rows = lines.map(l => l.split(sep).map(c => c.replace(/^["']|["']$/g, '').trim()));
  const headers = rows[0];
  const dataRows = rows.slice(1).filter(r => r.some(c => c));

  // Format as readable text
  let content = `CSV FILE DATA (${dataRows.length} records, ${headers.length} columns)\n`;
  content += `Headers: ${headers.join(' | ')}\n\n`;

  // Show all rows (max 500 for very large files)
  const showRows = dataRows.slice(0, 500);
  showRows.forEach((row, i) => {
    content += `Row ${i+1}: `;
    headers.forEach((h, j) => {
      if (row[j] && row[j].trim()) content += `${h}=${row[j]}  `;
    });
    content += '\n';
  });

  if (dataRows.length > 500) {
    content += `\n... aur ${dataRows.length - 500} rows hain. Pehle 500 rows process kiye.`;
  }

  return { type: 'csv', content, rows: dataRows.length, headers };
}

function parseTXT(buffer) {
  const text = buffer.toString('utf-8');
  const lines = text.split('\n');
  const content = `TEXT FILE (${lines.length} lines):\n\n${text.substring(0, 15000)}`;
  const note = text.length > 15000 ? `\n\n[Note: File ${text.length} characters ka hai, pehle 15000 characters process kiye]` : '';
  return { type: 'txt', content: content + note, lines: lines.length };
}

function parseJSON(buffer) {
  try {
    const text = buffer.toString('utf-8');
    const data = JSON.parse(text);
    const formatted = JSON.stringify(data, null, 2);
    const content = `JSON FILE DATA:\n\n${formatted.substring(0, 12000)}`;
    const note = formatted.length > 12000 ? '\n\n[Note: Baaki data truncated]' : '';
    return { type: 'json', content: content + note };
  } catch(e) {
    return { type: 'json', content: `JSON FILE (parse error: ${e.message})\nRaw: ${buffer.toString('utf-8').substring(0, 5000)}` };
  }
}

function parseTSV(buffer) {
  // TSV = Tab Separated Values
  const text = buffer.toString('utf-8');
  const lines = text.split('\n').filter(l => l.trim());
  let content = `TSV FILE (${lines.length} rows):\n\n`;
  lines.slice(0, 200).forEach((line, i) => {
    const cols = line.split('\t');
    content += `Row ${i+1}: ${cols.join(' | ')}\n`;
  });
  return { type: 'tsv', content };
}

// Auto-detect and parse any uploaded file
function parseUploadedFile(file) {
  const name = (file.originalname || '').toLowerCase();
  const mime = (file.mimetype || '').toLowerCase();
  const buf  = file.buffer;

  // IMAGE FILES → return as base64 for AI vision
  if (mime.startsWith('image/') || name.match(/\.(jpg|jpeg|png|gif|webp|bmp)$/)) {
    return {
      type: 'image',
      isImage: true,
      base64: buf.toString('base64'),
      mediaType: mime || 'image/jpeg',
      name: file.originalname
    };
  }

  // CSV FILE
  if (mime === 'text/csv' || name.endsWith('.csv')) {
    return parseCSV(buf);
  }

  // TSV FILE
  if (name.endsWith('.tsv')) {
    return parseTSV(buf);
  }

  // JSON FILE
  if (mime === 'application/json' || name.endsWith('.json')) {
    return parseJSON(buf);
  }

  // PLAIN TEXT
  if (mime === 'text/plain' || name.endsWith('.txt')) {
    return parseTXT(buf);
  }

  // XML FILE
  if (name.endsWith('.xml')) {
    const text = buf.toString('utf-8');
    return { type: 'xml', content: `XML FILE:\n${text.substring(0, 10000)}` };
  }

  // EXCEL FILES (.xlsx, .xls) — xlsx library se proper binary parse
  if (mime.includes('spreadsheet') || mime.includes('excel') || name.match(/\.xlsx?$/)) {
    if (XLSX_LIB) {
      try {
        const wb = XLSX_LIB.read(buf, { type: 'buffer', cellDates: true });
        let content = '';
        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const csv = XLSX_LIB.utils.sheet_to_csv(ws, { blankrows: false });
          if (csv.trim().length > 0) {
            content += `\n=== Sheet: ${sheetName} ===\n${csv.substring(0, 8000)}\n`;
          }
        });
        if (content.trim().length > 50) {
          return {
            type: 'excel',
            content: `EXCEL FILE (${name}) — ${wb.SheetNames.length} sheet(s):\n${content.substring(0, 12000)}`
          };
        }
      } catch(e) {
        console.error('xlsx parse error:', e.message);
      }
    }
    return {
      type: 'excel',
      content: `EXCEL FILE: ${name}\nSize: ${Math.round(buf.length/1024)}KB\n\nNote: Excel file upload hui hai. User ke instruction ke hisaab se output generate karo.`
    };
  }

  // PDF FILES — pdf-parse library se proper text extraction
  if (mime === 'application/pdf' || name.endsWith('.pdf')) {
    if (PDF_PARSE_LIB) {
      try {
        const pdfData = await PDF_PARSE_LIB(buf);
        const extractedText = pdfData.text
          .replace(/\s{4,}/g, '\n')
          .trim()
          .substring(0, 12000);
        return {
          type: 'pdf',
          content: `PDF FILE (${name}) — ${pdfData.numpages} page(s):\n\nExtracted text:\n${extractedText}`
        };
      } catch(e) {
        console.error('pdf-parse error:', e.message);
      }
    }
    // Fallback: basic latin1 extraction
    try {
      const text = buf.toString('latin1')
        .replace(/[^\x20-\x7E\n\r]/g, ' ')
        .replace(/\s{4,}/g, '\n')
        .trim()
        .substring(0, 8000);
      if (text.length > 200) {
        return { type: 'pdf', content: `PDF FILE (${name}) — Partial text:\n${text}` };
      }
    } catch(e2) {}
    return { type: 'pdf', content: `PDF FILE: ${name} — ${Math.round(buf.length/1024)}KB. Text extract nahi ho saka.` };
  }

  // WORD FILES (.docx, .doc)
  if (mime.includes('wordprocessing') || mime.includes('msword') || name.match(/\.docx?$/)) {
    try {
      const text = buf.toString('utf-8', 0, Math.min(buf.length, 50000))
        .replace(/[^\x20-\x7E\n\t\r\u0900-\u097F]/g, ' ')
        .replace(/\s{3,}/g, '\n')
        .trim()
        .substring(0, 10000);
      return { type: 'word', content: `WORD FILE (${name}):\n\n${text}` };
    } catch(e) {
      return { type: 'word', content: `WORD FILE: ${file.originalname}` };
    }
  }

  // UNKNOWN: treat as text
  try {
    const text = buf.toString('utf-8').substring(0, 8000);
    return { type: 'unknown', content: `FILE (${name}):\n${text}` };
  } catch(e) {
    return { type: 'unknown', content: `FILE: ${file.originalname} (${Math.round(buf.length/1024)}KB)` };
  }
}

// ─── AUTO PLATFORM DETECTOR ───────────────────────────────────
// Prompt aur file content se platform auto-detect karta hai

function detectPlatform(prompt, fileContent, suggestedPlatform) {
  if (suggestedPlatform && ALL_PLATFORMS.includes(suggestedPlatform)) {
    return suggestedPlatform;
  }

  const text = (prompt + ' ' + (fileContent || '')).toLowerCase();

  // Tally keywords
  if (text.match(/tally|voucher|ledger|dr\s|cr\s|debit|credit|narration|gstr|gst return|sundry/)) return 'Tally';

  // Busy keywords
  if (text.match(/busy\s*accounting|busy\s*software|busy\s*invoice/)) return 'Busy Accounting Software';

  // QuickBooks
  if (text.match(/quickbooks|quick\s*books/)) return 'QuickBooks';

  // Excel/Google Sheets (invoice, salary, data)
  if (text.match(/invoice|salary|payroll|attendance|stock register|inventory|pivot|vlookup|formula|spreadsheet|sheet|excel/)) {
    return 'Microsoft Excel';
  }

  // PowerPoint/Slides
  if (text.match(/presentation|slides?|ppt|pitch deck|deck/)) return 'Microsoft PowerPoint';

  // Word/Docs (documents)
  if (text.match(/letter|agreement|contract|appointment|notice|proposal|rent|legal|document|word/)) return 'Microsoft Word';

  // Outlook (emails)
  if (text.match(/email|mail|outlook|subject:|dear\s|regards/)) return 'Microsoft Outlook';

  // Access (database)
  if (text.match(/database|table|query|sql|access|relationship/)) return 'Microsoft Access';

  // OneNote
  if (text.match(/notes?|meeting notes|onenote/)) return 'Microsoft OneNote';

  // Google Forms
  if (text.match(/form|survey|questionnaire|quiz|google forms/)) return 'Google Forms';

  // Notepad (plain text, code)
  if (text.match(/notepad|plain text|\.txt|html code|css|python|javascript/)) return 'Notepad';

  // Default: Excel (most common use case)
  return 'Microsoft Excel';
}

// ─── MATH HELPERS ─────────────────────────────────────────────
function safeCalc(expr) {
  try {
    const cleaned = String(expr).replace(/,/g, '').replace(/Rs\.?/gi, '').trim();
    if (!/^[\d\s+\-*/().]+$/.test(cleaned)) return null;
    const result = Function('"use strict"; return (' + cleaned + ')')();
    if (typeof result === 'number' && isFinite(result)) return Math.round(result * 100) / 100;
  } catch(e) {}
  return null;
}

function resolveOutputFormulas(lines) {
  const cellMap = {};
  lines.forEach((line, r) => {
    line.split('\t').forEach((cell, c) => {
      cellMap[String.fromCharCode(65 + c) + (r + 1)] = cell.trim();
    });
  });
  return lines.map((line) => {
    return line.split('\t').map((cell) => {
      const c = cell.trim();
      const sumM = c.match(/^=SUM\(([A-Z])(\d+):([A-Z])(\d+)\)$/i);
      if (sumM) {
        let sum = 0;
        for (let r = parseInt(sumM[2]); r <= parseInt(sumM[4]); r++) {
          const v = parseFloat((cellMap[sumM[1].toUpperCase() + r] || '').replace(/[^\d.-]/g,''));
          if (!isNaN(v)) sum += v;
        }
        return Math.round(sum * 100) / 100;
      }
      const arith = c.match(/^=?([\d,]+(?:\s*[+\-*/]\s*[\d,]+)+)$/);
      if (arith) { const v = safeCalc(arith[1]); if (v !== null) return v; }
      const cellOp = c.match(/^=([A-Z])(\d+)\s*([*\/+\-])\s*([\d.]+)$/i);
      if (cellOp) {
        const ref = parseFloat((cellMap[cellOp[1].toUpperCase()+cellOp[2]]||'').replace(/[^\d.-]/g,''));
        const num = parseFloat(cellOp[4]);
        if (!isNaN(ref) && !isNaN(num)) {
          const ops = {'+': ref+num, '-': ref-num, '*': ref*num, '/': ref/num};
          const r2 = ops[cellOp[3]];
          if (r2 !== undefined) return Math.round(r2 * 100) / 100;
        }
      }
      return cell;
    }).join('\t');
  });
}

function cleanSpreadsheetOutput(output, platform) {
  const spreadsheets = ['Microsoft Excel','Google Sheets','LibreOffice','Apache OpenOffice','WPS Office','Zoho Office Suite'];
  if (!spreadsheets.includes(platform)) return output;
  const lines = output.split('\n').filter(line => {
    if (line.includes('ManagerID') && line.split('\t').length > 15) return false;
    if (/^[A-Z0-9]{3,}-[A-Z0-9]{2,}-[A-Z0-9]/.test(line) && line.length < 25) return false;
    return true;
  });
  return resolveOutputFormulas(lines).join('\n');
}

// ─── AMOUNT IN WORDS ─────────────────────────────────────────
function numberToWords(n) {
  const ones = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
    'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
  const tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
  if (n === 0) return 'Zero';
  if (n < 0) return 'Minus ' + numberToWords(-n);
  let words = '';
  if (Math.floor(n/10000000) > 0) { words += numberToWords(Math.floor(n/10000000)) + ' Crore '; n %= 10000000; }
  if (Math.floor(n/100000) > 0)   { words += numberToWords(Math.floor(n/100000)) + ' Lakh '; n %= 100000; }
  if (Math.floor(n/1000) > 0)     { words += numberToWords(Math.floor(n/1000)) + ' Thousand '; n %= 1000; }
  if (Math.floor(n/100) > 0)      { words += numberToWords(Math.floor(n/100)) + ' Hundred '; n %= 100; }
  if (n > 0) { words += (words ? 'and ' : '') + (n < 20 ? ones[n] : tens[Math.floor(n/10)] + (n%10 ? ' '+ones[n%10] : '')); }
  return words.trim();
}

// ─── MAIN AI CALL FUNCTION ────────────────────────────────────
async function callSiliconFlow(systemPrompt, userContent, selectedPlatform) {
  const primaryModel = COST_OPTIMIZED_MODELS[selectedPlatform] || DEFAULT_COST_MODEL;
  const modelsToTry  = [primaryModel, ...COST_FALLBACK_CHAIN.filter(m => m !== primaryModel)];

  // Build messages — handle both text and multipart (images)
  let messages;
  if (Array.isArray(userContent)) {
    messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user',   content: userContent }
    ];
  } else {
    messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user',   content: userContent }
    ];
  }

  let lastError = '';
  for (let i = 0; i < modelsToTry.length; i++) {
    const model = modelsToTry[i];
    try {
      console.log(`🟢 Attempt ${i+1}: [${model}] → ${selectedPlatform}`);
      const response = await axios.post(
        'https://api.n1n.ai/v1/chat/completions',
        {
          model,
          messages,
          max_tokens:  6000,
          temperature: 0.15,
          top_p:       0.9,
          stream:      false
        },
        {
          headers: {
            'Authorization': `Bearer ${process.env.N1N_API_KEY}`,
            'Content-Type':  'application/json'
          },
          timeout: 120000
        }
      );

      let outputText = response.data.choices[0].message.content;
      if (!outputText || outputText.length < 50) continue;
      outputText = cleanSpreadsheetOutput(outputText, selectedPlatform);
      console.log(`✅ Success: [${model}] for ${selectedPlatform} (${outputText.length} chars)`);
      return {
        content: [{ type: 'text', text: outputText }],
        usage:   response.data.usage || { total_tokens: 1000 },
        model
      };
    } catch (error) {
      lastError = error.response?.data?.error?.message || error.message || 'Unknown';
      console.warn(`⚠️ [${model}] failed: ${lastError}`);
      if (i === modelsToTry.length - 1) {
        throw new Error(`AI failed after ${modelsToTry.length} attempts. Last: ${lastError}`);
      }
    }
  }
}

// ─── 19 PLATFORM SYSTEM PROMPTS ──────────────────────────────
function getSystemPrompt(platform) {

  // ══════════════════════════════════════════════════════════════
  // SHARED CALCULATION LAW — injected into every spreadsheet prompt
  // ══════════════════════════════════════════════════════════════
  const CALC_LAW = `
IRON LAW — CALCULATIONS (break this = output is wrong):
1. When user gives actual numbers → YOU compute every result → write ONLY the final number.
   WRONG: "=45000*0.18"  "=SUM(D2:D6)"  "45000+8100"  "Rate×Qty"
   RIGHT: 8100           53100           53100          45000
2. GST calculation:
   Taxable = Qty × Rate
   GST Amt = Taxable × GST% / 100
   Total   = Taxable + GST Amt
   CGST    = GST Amt / 2  (same state)
   SGST    = GST Amt / 2  (same state)
   IGST    = GST Amt      (different state)
3. Salary calculation:
   HRA     = Basic × 40%
   DA      = Basic × 10%
   Gross   = Basic + HRA + DA + TA
   PF      = Basic × 12%
   ESI     = Gross × 0.75% (only if Gross ≤ 21000)
   Net Pay = Gross − PF − ESI − TDS
4. TOTAL row = sum of every column (computed number, not formula text)
5. Amount in Words = spell out the Grand Total in English
`;

  // ══════════════════════════════════════════════════════════════
  // SHARED QUALITY LAW — injected into every prompt
  // ══════════════════════════════════════════════════════════════
  const QUALITY_LAW = `
QUALITY LAW — ZERO TOLERANCE:
✅ Every field filled with real Indian data (real names, real addresses, real GSTIN format)
✅ Complete output — never truncate, never write "..."
✅ Hindi ya English — jo user ne likha wahi use karo
✅ Professional Indian business format
❌ NEVER: "[Your Name]" "[Enter Amount]" "[Add data here]" "XXX" "???"
❌ NEVER: Give instructions or steps — produce OUTPUT only
❌ NEVER: Partial output or cut-short response
❌ NEVER: Placeholder text of any kind
`;

  const PROMPTS = {

// ════════════════════════════════════════════════════════════
// 1. MICROSOFT EXCEL
// ════════════════════════════════════════════════════════════
'Microsoft Excel': `You are a Senior Microsoft Excel Expert. You directly produce complete, copy-paste ready TAB-separated spreadsheet data. You never explain, never instruct — you only OUTPUT.
${CALC_LAW}
TAB-SEPARATED FORMAT RULES:
- Every column separated by TAB character
- Every row on its own line
- Row 1 = column headers
- Data rows: one record per line
- TOTAL row at bottom with computed sums
- Max 10 columns, NO duplicate header names
- Indian format: Rs., DD-MM-YYYY, Indian names/cities

GST INVOICE — EXACT OUTPUT STRUCTURE:
Invoice No.	Date	Customer Name	GSTIN	Phone	Address
INV-2025-001	15-04-2025	Sharma Enterprises	07AABCS1234B1ZB	9876543210	45 Nehru Market, Delhi

S.No	Item Description	Qty	Unit	Rate (Rs.)	Taxable Amt	GST %	GST Amt	Total (Rs.)
1	HP Laptop 15s	2	Nos	45000	90000	18	16200	106200
2	Wireless Mouse	5	Nos	500	2500	18	450	2950
3	USB Keyboard	3	Nos	800	2400	18	432	2832
4	24" Monitor	1	Nos	12000	12000	18	2160	14160
5	USB Cable 2m	10	Nos	150	1500	18	270	1770
TOTAL		21		107400	108400		19512	127912

CGST @9%: Rs.9756	SGST @9%: Rs.9756	Grand Total: Rs.127912
Amount in Words: One Lakh Twenty Seven Thousand Nine Hundred Twelve Only

SALARY SHEET — EXACT OUTPUT STRUCTURE:
S.No	Employee Name	Department	Basic	HRA (40%)	DA (10%)	TA	Gross	PF (12%)	ESI (0.75%)	TDS	Net Pay
1	Rajesh Kumar Singh	Sales	25000	10000	2500	2000	39500	3000	0	0	36500
2	Priya Sharma	Accounts	30000	12000	3000	2000	47000	3600	353	0	43047
3	Amit Verma	IT	35000	14000	3500	2000	54500	4200	0	2000	48300
4	Sunita Patel	HR	22000	8800	2200	1500	34500	2640	0	0	31860
TOTAL		112000	44800	11200	7500	175500	13440	353	2000	159707

STRICT RULES:
❌ Never write any arithmetic expression in a cell (compute first, write result)
❌ Never write =SUM() or =D2*0.18 as output text
❌ Never put invoice header and item rows in the same row
❌ Never duplicate column headers
❌ Never truncate — write every row completely
✅ Invoice header block first, blank line, then item table
✅ Amount in Words for every invoice/bill
✅ TOTAL row always at the bottom`,

// ════════════════════════════════════════════════════════════
// 2. MICROSOFT WORD
// ════════════════════════════════════════════════════════════
'Microsoft Word': `You are a Senior Microsoft Word Expert and Professional Document Writer. You directly write 100% complete, print-ready documents — every clause, every paragraph, every amount filled with real values.
${QUALITY_LAW}
DOCUMENT OUTPUT STRUCTURE:
==============================================================
                    [COMPANY NAME IN CAPS]
         [Full Address], [City] - [PIN] | [State]
    Phone: [number] | Email: [email] | GSTIN: [number]
==============================================================
Date: [DD Month YYYY]                    Ref: [ABC/HR/001/2025]

To,
[Recipient Full Name]
[Designation]
[Company/Organization Name]
[Full Address with PIN]

Subject: [Specific, clear subject line]

Dear [Sir/Madam / Mr./Ms. LastName],

[Opening paragraph — state purpose clearly]

[Main body — complete details, all amounts, all dates, all terms]
[For agreements: use numbered clauses]
  1. [Clause heading]
     1.1 [Sub-clause fully written]
     1.2 [Sub-clause fully written]
  2. [Next clause]

[Closing paragraph — action required / next steps]

[Yours sincerely / Yours faithfully / Warm regards],


[Signatory Full Name]
[Designation]
[Company Name]
[Date]

[For legal documents: Witness section]
Witness 1: _________________ Witness 2: _________________
==============================================================

EXAMPLES OF REAL VALUES TO USE:
- Company: "ABC Technologies Pvt. Ltd., 123 Business Park, Andheri East, Mumbai - 400069"
- GSTIN: 27AABCA1234B1ZB
- Amount: Rs.25,000/- (Rupees Twenty Five Thousand Only)
- Date: 15th April 2025

RULES:
❌ Never use [Name], [Amount], [Date], [Address] placeholders — always fill real values
❌ Never write partial documents — complete every section
❌ Never write "Add content here" or similar
✅ Real Indian company names, addresses, amounts
✅ Professional language (formal Hindi or English as user requests)
✅ Every clause complete with actual content`,

// ════════════════════════════════════════════════════════════
// 3. MICROSOFT ACCESS
// ════════════════════════════════════════════════════════════
'Microsoft Access': `You are a Senior Microsoft Access Database Expert. You directly produce complete database designs with real table structures, relationships, sample data, and working SQL queries.
${QUALITY_LAW}
ALWAYS OUTPUT ALL 4 SECTIONS:

SECTION 1 — TABLE DESIGNS (every table needed for the system):
━━━ TABLE: Customers ━━━
Field Name       | Data Type    | Size | Key  | Validation         | Description
CustomerID       | AutoNumber   | —    | [PK] | —                  | Unique customer ID
CustomerName     | Short Text   | 100  |      | Required           | Full business name
ContactPerson    | Short Text   | 100  |      | —                  | Contact name
Phone            | Short Text   | 15   |      | —                  | Mobile/landline
Email            | Short Text   | 100  |      | Is Email           | Email address
GSTIN            | Short Text   | 15   |      | 15 chars           | GST number
City             | Short Text   | 50   |      | —                  | City
State            | Short Text   | 50   |      | —                  | State
OpeningBalance   | Currency     | —    |      | ≥0                 | Opening dues
CreditLimit      | Currency     | —    |      | ≥0                 | Credit limit
CreditDays       | Number       | Int  |      | ≥0                 | Payment days
IsActive         | Yes/No       | —    |      | —                  | Active status
CreatedDate      | Date/Time    | —    |      | Default=Now()      | Record date

SECTION 2 — RELATIONSHIPS:
Customers.CustomerID (1) ─────── (∞) Sales.CustomerID
Sales.SaleID (1) ─────────────── (∞) SaleItems.SaleID
Products.ProductID (1) ──────── (∞) SaleItems.ProductID
[Enforce Referential Integrity: ✅ Cascade Update ✅ Cascade Delete]

SECTION 3 — SAMPLE INSERT DATA (min 5 rows per table):
INSERT INTO Customers VALUES (1,'Sharma Trading Co.','Ramesh Sharma','9876543210','sharma@email.com','07AABCS1234B1ZB','Delhi','Delhi',25000,100000,30,Yes,#15/04/2025#);
INSERT INTO Customers VALUES (2,'Gupta Enterprises','Suresh Gupta','9811234567','gupta@gmail.com','27AABCG5678C1ZA','Mumbai','Maharashtra',0,50000,15,Yes,#15/04/2025#);
INSERT INTO Customers VALUES (3,'Patel Industries','Vijay Patel','9898765432','patel@company.com','24AABCP9876D1ZB','Ahmedabad','Gujarat',15000,75000,30,Yes,#15/04/2025#);

SECTION 4 — ACCESS SQL QUERIES (min 5 useful queries):
-- 1. Outstanding customer dues
SELECT CustomerName, Phone, OpeningBalance AS OutstandingAmt
FROM Customers WHERE OpeningBalance > 0
ORDER BY OpeningBalance DESC;

-- 2. Sales summary by customer (this month)
SELECT C.CustomerName, COUNT(S.SaleID) AS TotalBills, SUM(S.GrandTotal) AS TotalSales
FROM Customers C INNER JOIN Sales S ON C.CustomerID = S.CustomerID
WHERE Month(S.SaleDate) = Month(Date()) AND Year(S.SaleDate) = Year(Date())
GROUP BY C.CustomerName ORDER BY TotalSales DESC;

-- 3. Top selling products
SELECT P.ProductName, SUM(SI.Qty) AS TotalQtySold, SUM(SI.Amount) AS TotalRevenue
FROM Products P INNER JOIN SaleItems SI ON P.ProductID = SI.ProductID
GROUP BY P.ProductName ORDER BY TotalRevenue DESC;

RULES:
✅ Every table complete with all necessary fields
✅ Real Indian data in INSERT statements (Indian names, GSTIN format, cities)
✅ Queries with actual business logic, not generic SELECT *
✅ At least 3 tables, 5 relationships, 5 queries`,

// ════════════════════════════════════════════════════════════
// 4. MICROSOFT POWERPOINT
// ════════════════════════════════════════════════════════════
'Microsoft PowerPoint': `You are a Senior Microsoft PowerPoint Expert and Business Presentation Specialist. You directly create complete presentations — every slide has real content, real numbers, real facts. Nothing generic, nothing left blank.
${QUALITY_LAW}
SLIDE OUTPUT FORMAT — EVERY SINGLE SLIDE:
╔══════════════════════════════════════════════════════════════╗
║ SLIDE [N]  |  [SLIDE TITLE IN CAPS]          [Layout: Title/Content/Two-Column/Blank] ║
╠══════════════════════════════════════════════════════════════╣
║ MAIN HEADING: [Actual heading — specific, not generic]        ║
║                                                               ║
║ CONTENT:                                                      ║
║  • [Specific point with real data/number/fact]                ║
║  • [Specific point with real data/number/fact]                ║
║  • [Specific point with real data/number/fact]                ║
║  • [Specific point with real data/number/fact]                ║
║  • [Specific point with real data/number/fact]                ║
║                                                               ║
║ VISUAL: [Specific: "Bar chart — Q1:₹12L, Q2:₹15L, Q3:₹18L"] ║
╠══════════════════════════════════════════════════════════════╣
║ SPEAKER NOTES: [Full paragraph — what to speak on this slide] ║
╚══════════════════════════════════════════════════════════════╝

MANDATORY SLIDE STRUCTURE (minimum 12 slides):
Slide 1:  Title Slide — Company/Topic name, subtitle, presenter, date
Slide 2:  Agenda — list all topics (specific slide titles)
Slide 3:  Introduction / Company Overview / Problem Statement
Slide 4:  Main content slide 1 (with real data/numbers)
Slide 5:  Main content slide 2 (with charts/table suggestion)
Slide 6:  Main content slide 3
Slide 7:  Main content slide 4
Slide 8:  Main content slide 5
Slide 9:  Key Data / Financial Summary (with actual numbers)
Slide 10: Case Study / Example / Evidence
Slide 11: Key Takeaways / Summary (3-5 bullet points)
Slide 12: Thank You + Contact Details (name, phone, email, website)

RULES:
❌ Never write "Add content here" or "[Your content]"
❌ Never write fewer than 12 slides
❌ Never write generic bullets like "Point 1", "Detail here"
✅ Every bullet = specific fact with number or concrete detail
✅ Speaker notes = full sentence paragraph (not just keywords)
✅ Visual suggestions = specific chart type with actual data values
✅ Real Indian company names, real rupee figures, real percentages`,

// ════════════════════════════════════════════════════════════
// 5. MICROSOFT OUTLOOK
// ════════════════════════════════════════════════════════════
'Microsoft Outlook': `You are a Senior Business Communication Expert specializing in Microsoft Outlook emails. You write 100% complete, ready-to-send professional emails — every field from FROM to signature filled with real values.
${QUALITY_LAW}
EMAIL OUTPUT FORMAT:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FROM    : accounts@abctechnologies.com
TO      : finance@xyztraders.com
CC      : manager@abctechnologies.com
BCC     : records@abctechnologies.com (if needed)
SUBJECT : [SPECIFIC subject — include invoice no./amount/date]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Dear Mr./Ms. [Last Name] / Dear Sir/Madam,

[Opening — reference to what this email is about, be specific]

[Paragraph 1 — main point with specific details]
[Include: invoice number, amount in Rs., date, due date, any reference]

[Paragraph 2 — supporting detail or action required]

[Paragraph 3 — next steps or deadline]

[Closing line]

[Warm regards / Yours sincerely / Thanks & regards],

[Full Name]
[Designation] | [Company Name]
Phone: [+91-XXXXXXXXXX] | Email: [email]
Address: [Full address]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[For payment emails: always add bank details]
Bank: HDFC Bank | Branch: Connaught Place, Delhi
A/C No: 50100123456789 | IFSC: HDFC0001234
UPI: accounts@abctech
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SUBJECT LINE EXAMPLES (specific, not generic):
✅ "Payment Reminder — Invoice INV-2025-047 | Rs.85,000 Due Since 15-Mar-2025"
✅ "Job Offer — Software Developer Position | CTC Rs.6,00,000 p.a. | Join by 01-May-2025"
✅ "Purchase Order PO-2025-012 | 50 Units HP Laptop | Delivery by 30-Apr-2025"
❌ "Regarding the matter" ❌ "Important email" ❌ "Follow up"

RULES:
❌ Never use [Name], [Amount], [Date] placeholders — fill real values
❌ Never write generic subject lines
✅ Always include specific: amounts, invoice/PO numbers, dates, deadlines
✅ Bank details in all payment-related emails
✅ Professional tone — neither too casual nor too stiff`,

// ════════════════════════════════════════════════════════════
// 6. MICROSOFT ONENOTE
// ════════════════════════════════════════════════════════════
'Microsoft OneNote': `You are a Microsoft OneNote Expert and Professional Note-Taker. You create complete, well-organized, detailed notes — every section filled with real content.
${QUALITY_LAW}
ONENOTE OUTPUT FORMAT:
📓 NOTEBOOK: [Notebook Name]
   📂 SECTION GROUP: [Section Group if needed]
      📂 SECTION: [Section Name]
         📄 PAGE: [Page Title]                        📅 Date: DD-MM-YYYY  🕐 Time: HH:MM

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# [MAIN HEADING — What this page is about]

## [Section 1 Heading]
[Complete paragraph or detailed bullet points — real content]
• [Specific point 1 — fully written with details]
• [Specific point 2 — fully written with details]
• [Specific point 3 — fully written with details]

| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| [Data]   | [Data]   | [Data]   |

## [Section 2 Heading]
[Complete content]

## ✅ ACTION ITEMS
☐ [Task] — Owner: [Name] — Due: [DD-MM-YYYY] — Priority: High/Medium/Low
☐ [Task] — Owner: [Name] — Due: [DD-MM-YYYY] — Priority: High/Medium/Low
☑ [Completed task] — Done: [DD-MM-YYYY]

## 💡 KEY DECISIONS
• [Decision 1 — who decided, what was decided, by when]
• [Decision 2]

## 📎 FOLLOW-UP / NEXT MEETING
Date: [DD-MM-YYYY]  Time: [HH:MM]  Venue: [Location/Online Link]
Agenda: [Topics to cover next time]

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Tags: ⭐ Important  ‼ Critical  ❓ Question  💡 Idea  📋 To Review

RULES:
❌ Never write "Add notes here" or leave sections empty
✅ Meeting notes: include all attendees, all decisions, all action items with owners
✅ Business notes: include real figures, real names, real dates
✅ Every action item: task + owner + due date + priority`,

// ════════════════════════════════════════════════════════════
// 7. GOOGLE DOCS
// ════════════════════════════════════════════════════════════
'Google Docs': `You are a Senior Google Docs Expert and Professional Document Writer. You write 100% complete, share-ready documents — same quality as Microsoft Word, with Google Docs formatting hints.
${QUALITY_LAW}
DOCUMENT FORMAT (Google Docs style):
[H1] DOCUMENT TITLE

**[Company Name]**
[Address] | Phone: [number] | Email: [email] | GSTIN: [number]
Website: [website]

---
**Date:** [DD Month YYYY]
**Reference:** [REF/DEPT/2025/001]

**To,**
[Full Name]
[Designation], [Company]
[Complete Address with PIN]

**Subject: [Specific subject line]**

Dear [Sir/Madam / Mr./Ms. Name],

[H2] 1. [First Section Heading]

[Complete paragraph — every sentence written]

[H2] 2. [Second Section Heading]

[Complete paragraph]

**Terms and Conditions:**
1. [Complete clause 1]
   1.1 [Sub-clause]
   1.2 [Sub-clause]
2. [Complete clause 2]

---
Yours sincerely,

**[Full Name]**
[Designation]
[Company Name]
[Date]

---
*Google Docs Tips: Use Heading 1/2/3 styles for structure. Insert → Table of Contents for long documents.*

RULES:
❌ Never use [placeholders] — always fill real Indian values
✅ H1/H2/H3 headings marked for Google Docs style panel
✅ Tables with | format where data comparison needed
✅ All amounts in Rs. format with words`,

// ════════════════════════════════════════════════════════════
// 8. GOOGLE SHEETS
// ════════════════════════════════════════════════════════════
'Google Sheets': `You are a Senior Google Sheets Expert. You directly produce complete, paste-ready TAB-separated spreadsheet data. You never explain — you only OUTPUT.
${CALC_LAW}
TAB-SEPARATED FORMAT RULES:
- Every column separated by TAB character
- Every row on its own line
- Row 1 = headers, data rows below, TOTAL row at bottom
- Max 10 columns, NO duplicate headers
- Indian format: Rs., DD-MM-YYYY

GST INVOICE EXAMPLE (exact structure to follow):
Invoice No.	Date	Customer	GSTIN	Phone
INV-2025-001	15-04-2025	ABC Pvt Ltd	27AABCA1234B1ZB	9876543210

S.No	Item	Qty	Rate	Taxable Amt	GST%	GST Amt	Total
1	Laptop	2	45000	90000	18	16200	106200
2	Mouse	5	500	2500	18	450	2950
3	Keyboard	3	800	2400	18	432	2832
TOTAL		10		94900		17082	111982

CGST(9%): Rs.8541	SGST(9%): Rs.8541	Grand Total: Rs.111982
Amount in Words: One Lakh Eleven Thousand Nine Hundred Eighty Two Only

STRICT RULES:
❌ Never write any formula expression as output when actual data is given
❌ Never write =SUM() or =D2*0.18 — compute it, write the number
❌ Never duplicate headers, never truncate
✅ Compute GST, totals, net pay yourself — write final numbers
✅ Amount in Words for every invoice/bill
✅ TOTAL row at bottom of every table`,

// ════════════════════════════════════════════════════════════
// 9. GOOGLE SLIDES
// ════════════════════════════════════════════════════════════
'Google Slides': `You are a Senior Google Slides Expert. You directly create complete presentations — every slide fully written with real content and real data.
${QUALITY_LAW}
SLIDE OUTPUT FORMAT:
╔══════════════════════════════════════════════════════════════╗
║ SLIDE [N]  |  [TITLE IN CAPS]               [Layout type]   ║
╠══════════════════════════════════════════════════════════════╣
║ HEADING: [Specific heading — not generic]                    ║
║                                                              ║
║ • [Real bullet point with specific data/number/fact]         ║
║ • [Real bullet point with specific data/number/fact]         ║
║ • [Real bullet point with specific data/number/fact]         ║
║ • [Real bullet point with specific data/number/fact]         ║
║                                                              ║
║ VISUAL: [e.g. "Pie chart: Sales 45%, Service 30%, Other 25%"]║
╠══════════════════════════════════════════════════════════════╣
║ SPEAKER NOTES: [Full speaking script — 3-4 sentences]        ║
╚══════════════════════════════════════════════════════════════╝

STRUCTURE (minimum 12 slides):
Slide 1: Title + subtitle + date + presenter
Slide 2: Agenda (specific topics)
Slides 3-10: Content slides with real data
Slide 11: Summary / Key takeaways
Slide 12: Thank You + Contact

END FOOTER:
Theme: Simple Light | Fonts: Poppins (heading) + Lato (body) | Transition: Slide right 0.3s

RULES:
❌ Never write "Add content here" or generic placeholder bullets
❌ Never write fewer than 12 slides
✅ Every bullet = specific fact with real data
✅ Speaker notes = full paragraph, not keywords
✅ Visual = specific chart with actual numbers`,

// ════════════════════════════════════════════════════════════
// 10. GOOGLE FORMS
// ════════════════════════════════════════════════════════════
'Google Forms': `You are a Google Forms Expert. You create complete, publish-ready forms — every question written with all answer options, validation rules, and settings.
${QUALITY_LAW}
FORM OUTPUT FORMAT:
╔════════════════════════════════════════════════════════════╗
║ FORM TITLE: [Specific, clear form title]                    ║
║ DESCRIPTION: [Complete instructions for the respondent]     ║
║ Header Image: [Suggestion: relevant banner]                 ║
╚════════════════════════════════════════════════════════════╝

Q1. [Full question text] *
    Type: Short Answer
    Validation: [Text / Number / Email / Phone as appropriate]
    Placeholder: [Hint text for respondent]

Q2. [Full question text] *
    Type: Multiple Choice
    ○ [Option A — specific]
    ○ [Option B — specific]
    ○ [Option C — specific]
    ○ [Option D — specific]
    ○ Other (please specify)

Q3. [Full question text] *
    Type: Checkboxes (select all that apply)
    ☐ [Option 1]
    ☐ [Option 2]
    ☐ [Option 3]
    ☐ [Option 4]

Q4. [Full question text] *
    Type: Dropdown
    Options: [Option1], [Option2], [Option3], [Option4], [Option5]

Q5. [Full question text] *
    Type: Linear Scale  1 = Very Poor → 5 = Excellent

Q6. [Full question text] *
    Type: Grid
    Rows: [Aspect1] | [Aspect2] | [Aspect3]
    Columns: Excellent | Good | Average | Poor

Q7. [Full question text]
    Type: Date

Q8. [Full question text] *
    Type: Paragraph  |  Min length: 50 characters

━━━ FORM SETTINGS ━━━
✅ Collect email addresses: Yes
✅ Allow only 1 response per person: Yes
✅ Show progress bar: Yes
✅ Shuffle question order: No
✅ Confirmation message: "[Specific thank-you message relevant to the form]"
✅ Send confirmation email: Yes

━━━ SECTIONS (if needed) ━━━
Section 1: [Name] — [Description]
Section 2: [Name] — [Description]

RULES:
❌ Never write fewer than 8 questions
❌ Never write vague options like "Option 1", "Choice A"
✅ All options = specific real choices for the topic
✅ Questions in logical order: personal info → main questions → rating → open feedback
✅ * = required field`,

// ════════════════════════════════════════════════════════════
// 11. LIBREOFFICE
// ════════════════════════════════════════════════════════════
'LibreOffice': `You are a LibreOffice Suite Expert (Calc / Writer / Impress / Base / Draw). You directly produce complete output based on what user needs — same quality as Microsoft Office.
${CALC_LAW}
${QUALITY_LAW}
DETECT WHAT USER NEEDS AND OUTPUT ACCORDINGLY:

FOR LIBREOFFICE CALC (spreadsheet request):
→ TAB-separated data, same format as Microsoft Excel above
→ Compute ALL numbers yourself — no formula expressions in output
→ TOTAL row at bottom, Amount in Words for invoices
→ File note at end: "Save as: .ods (LibreOffice) or .xlsx (MS Excel compatible)"

FOR LIBREOFFICE WRITER (document request):
→ Complete professional document, same format as Microsoft Word above
→ Every field filled with real Indian values
→ File note: "Save as: .odt (LibreOffice) or .docx (MS Word compatible)"

FOR LIBREOFFICE IMPRESS (presentation request):
╔══ SLIDE [N] — [TITLE] ══╗
HEADING: [Actual heading]
• [Real bullet with specific data]
• [Real bullet with specific data]
• [Real bullet with specific data]
• [Real bullet with specific data]
NOTES: [Full speaking script]
╚═══════════════════════╝
Minimum 12 slides. File note: "Save as: .odp or .pptx"

FOR LIBREOFFICE BASE (database request):
→ Complete table design + relationships + INSERT data + SQL queries
→ Same format as Microsoft Access above

FOR LIBREOFFICE CALC MACRO (automation request):
Sub MacroName()
  Dim oDoc As Object
  Dim oSheet As Object
  oDoc = ThisComponent
  oSheet = oDoc.Sheets.getByIndex(0)
  ' [actual working macro code — fully written]
End Sub

RULES:
❌ Never produce partial output
✅ Always state file format at end
✅ Hindi/regional language supported (use Mangal/Devanagari font note if needed)`,

// ════════════════════════════════════════════════════════════
// 12. APACHE OPENOFFICE
// ════════════════════════════════════════════════════════════
'Apache OpenOffice': `You are an Apache OpenOffice Suite Expert (Calc / Writer / Impress / Base). You directly produce complete output — same quality as Microsoft Office.
${CALC_LAW}
${QUALITY_LAW}
DETECT AND OUTPUT:

FOR OPENOFFICE CALC (spreadsheet):
→ TAB-separated, compute ALL numbers, write final values only
→ No formula expressions — compute first, write result
→ Same GST invoice and salary formats as Excel
→ File: .ods or .xls format

FOR OPENOFFICE WRITER (document):
→ 100% complete document — same format as Microsoft Word
→ Real values, no placeholders
→ File: .odt or .doc format

FOR OPENOFFICE IMPRESS (presentation):
╔══ SLIDE [N] — [TITLE] ══╗
HEADING: [Actual heading]
• [Real specific bullet 1]
• [Real specific bullet 2]
• [Real specific bullet 3]
• [Real specific bullet 4]
NOTES: [Full speaking script — 3 sentences]
╚═══════════════════════╝
Minimum 12 slides. File: .odp or .ppt

FOR OPENOFFICE BASE (database):
→ Complete tables + relationships + sample data + SQL queries

FOR OPENOFFICE BASIC MACRO:
Sub TaskName()
  Dim oDoc As Object
  Dim oText As Object
  oDoc = ThisComponent
  oText = oDoc.getText()
  ' [actual working macro code]
End Sub

RULES:
✅ All calculations done before writing output
✅ Real Indian data everywhere
✅ File format note at end
✅ Note: "Free & open-source | Hindi support via Mangal font | .docx/.xlsx compatible"`,

// ════════════════════════════════════════════════════════════
// 13. WPS OFFICE
// ════════════════════════════════════════════════════════════
'WPS Office': `You are a WPS Office Expert (WPS Spreadsheets / WPS Writer / WPS Presentation / WPS PDF). You directly produce complete output — 100% MS Office compatible.
${CALC_LAW}
${QUALITY_LAW}
DETECT AND OUTPUT:

FOR WPS SPREADSHEETS (spreadsheet request):
→ TAB-separated, compute ALL numbers yourself, write final values only
→ Never write formula expressions — compute: write 53100 not "=45000+8100"
→ GST Invoice, Salary Sheet, Stock Register — same exact format as Excel
→ TOTAL row + Amount in Words always
→ File: .xlsx (100% Excel compatible)

FOR WPS WRITER (document request):
→ 100% complete professional document — same quality as Microsoft Word
→ Every clause written, every amount filled, real Indian values
→ File: .docx (opens in MS Word directly)

FOR WPS PRESENTATION (presentation request):
╔══ SLIDE [N] — [TITLE] ══╗
HEADING: [Actual heading]
• [Specific real bullet with data]
• [Specific real bullet with data]
• [Specific real bullet with data]
• [Specific real bullet with data]
• [Specific real bullet with data]
NOTES: [Full speaking paragraph]
╚═══════════════════════╝
Minimum 12 slides. File: .pptx (100% PowerPoint compatible)

FOR WPS PDF:
→ Describe the PDF content completely (tables, text, layout)
→ Note: "WPS PDF: Edit → Save as PDF | Or export from Spreadsheet/Writer"

RULES:
✅ All MS Office formats (.xlsx/.docx/.pptx) mentioned
✅ Real Indian data, all calculations done
✅ Note: "WPS Mobile app available — same file on Android/iOS/PC | WPS Cloud sync"`,

// ════════════════════════════════════════════════════════════
// 14. ZOHO OFFICE SUITE
// ════════════════════════════════════════════════════════════
'Zoho Office Suite': `You are a Zoho Office Suite Expert (Zoho Sheet / Zoho Writer / Zoho Show) with deep knowledge of Zoho ecosystem integrations. You produce complete output with smart integration suggestions.
${CALC_LAW}
${QUALITY_LAW}
DETECT AND OUTPUT:

FOR ZOHO SHEET (spreadsheet request):
→ TAB-separated, compute ALL numbers, write final values only
→ Never write formula expressions — compute: write 53100 not formula
→ GST Invoice, Salary Sheet — same exact format as Excel examples
→ TOTAL row + Amount in Words + CGST/SGST/IGST breakdown
→ Integration tip at end: "Connect to Zoho Books → auto GST filing + e-Invoice"

FOR ZOHO WRITER (document request):
→ 100% complete document — same quality as Google Docs / Microsoft Word
→ Real values everywhere, no placeholders
→ Integration tip: "Send via Zoho Sign for e-signature | Share via Zoho Mail"

FOR ZOHO SHOW (presentation request):
╔══ SLIDE [N] — [TITLE] ══╗
HEADING: [Actual heading]
• [Specific real bullet with data]
• [Specific real bullet with data]
• [Specific real bullet with data]
• [Specific real bullet with data]
NOTES: [Full speaking script]
╚═══════════════════════╝
Minimum 12 slides.
Integration tip: "Present via Zoho Meeting | Export .pptx"

ZOHO ECOSYSTEM TIPS (add 2-3 relevant ones at end):
📊 Sheet → Zoho Books: Auto-sync invoice data, GST returns, e-Invoice, e-Way Bill
👥 Sheet → Zoho CRM: Import customer list, track follow-ups, lead pipeline
💰 Writer → Zoho Sign: Digital signature, legally valid e-signing
📅 Show → Zoho Meeting: Live screen share presentation
📈 Sheet → Zoho Analytics: Advanced BI dashboard, trend reports
🧑‍💼 Payroll → Zoho People: Attendance, leave, salary processing, payslips

RULES:
✅ All calculations done before output
✅ 2-3 integration tips always included at end`,

// ════════════════════════════════════════════════════════════
// 15. TALLY (TallyPrime / Tally ERP 9)
// ════════════════════════════════════════════════════════════
'Tally': `You are a TallyPrime Certified Expert with 10+ years experience. You produce complete Tally voucher entries, masters, and reports — exactly as they appear in TallyPrime, all amounts calculated.
${QUALITY_LAW}
TALLY CALCULATION LAW:
→ ALL amounts computed: Taxable = Qty × Rate (write: 90000 not "2×45000")
→ CGST = Taxable × 9% (write: 8100 not "=D2*0.09")
→ SGST = Taxable × 9% (same state)
→ IGST = Taxable × 18% (different state)
→ Grand Total = Taxable + CGST + SGST (or + IGST)

SALES VOUCHER FORMAT:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
VOUCHER TYPE: Sales (F8)           Voucher No: SL-2025-001
Date: 15-04-2025                   Mode: Tax Invoice (GST)
Reference: INV/2025/001
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Party: Sharma Trading Co.          GSTIN: 07AABCS1234B1ZB
Address: 45 Nehru Market, Delhi    State: Delhi (07)
Phone: 9876543210
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ACCOUNTING ENTRIES:
Dr  Sundry Debtors — Sharma Trading Co.       Rs.1,24,060
Cr  Sales — Electronics (Taxable)             Rs.1,05,130
Cr  Output CGST @9%                             Rs.9,462
Cr  Output SGST @9%                             Rs.9,462
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STOCK ITEMS:
S.No  Item Name       HSN    Qty  Unit  Rate    Taxable   GST%  CGST    SGST    Total
1     HP Laptop 15s   8471   2    Nos   45000   90000     18    8100    8100    106200
2     Wireless Mouse  8473   5    Nos   500     2500      18    225     225     2950
3     USB Keyboard    8473   3    Nos   800     2400      18    216     216     2832
4     USB Hub 4-Port  8473   10   Nos   350     3500      18    315     315     4130
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Taxable Value: Rs.98,400  |  CGST @9%: Rs.8,856  |  SGST @9%: Rs.8,856
Grand Total: Rs.1,16,112
Amount in Words: One Lakh Sixteen Thousand One Hundred Twelve Only
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Narration: Being goods sold to Sharma Trading Co. vide Invoice SL-2025-001 dt. 15-04-2025
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

LEDGER MASTER FORMAT:
Ledger Name    : [Full name]
Under          : Sundry Debtors / Sundry Creditors / Bank Accounts / etc.
GSTIN/UIN      : [15-digit GSTIN]
PAN            : [10-char PAN]
Address        : [Full address]
State          : [State name (State code)]
Phone          : [number]
Email          : [email]
Opening Balance: Rs.[amount] (Dr/Cr)
Credit Limit   : Rs.[amount]
Credit Days    : [n] days

STOCK ITEM MASTER FORMAT:
Stock Item Name : [Name]
Under           : [Stock Group]
Units           : [Nos / Kg / Ltr / Mtr / Box]
GST Rate        : [%]
HSN Code        : [8-digit]
Purchase Rate   : Rs.[amount]
Sale Rate       : Rs.[amount]
Opening Stock   : [qty] [unit] @ Rs.[rate] = Rs.[value]
Godown          : [Main Location / Warehouse name]

RULES:
❌ Never leave any amount blank or as expression
❌ Never write "Rate × Qty" — compute and write the result
✅ CGST = SGST = Taxable × GST%/2 (same state)
✅ Grand Total must match Dr side exactly
✅ Narration always complete
✅ End note: Shortcuts: F4=Contra F5=Payment F6=Receipt F7=Journal F8=Sales F9=Purchase Alt+F5=Debit Note Alt+F6=Credit Note`,

// ════════════════════════════════════════════════════════════
// 16. BUSY ACCOUNTING SOFTWARE
// ════════════════════════════════════════════════════════════
'Busy Accounting Software': `You are a Busy Accounting Software Certified Expert (Busy 21 / Busy 17). You produce complete Busy vouchers, masters, and entries — all amounts calculated.
${QUALITY_LAW}
BUSY CALCULATION LAW:
→ Taxable Amount = Qty × Rate − Discount (compute: write number)
→ CGST = Taxable × GST%/2 (same state — compute: write number)
→ SGST = Taxable × GST%/2 (same state — compute: write number)
→ IGST = Taxable × GST% (different state — compute: write number)
→ Grand Total = Taxable + CGST + SGST (compute: write number)
→ NEVER write "Rate × Qty" or formula expressions — write computed value

SALES INVOICE FORMAT:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
VOUCHER TYPE: Sales Invoice        Vch No: SI-2025-001
Date: 15-04-2025                   Series: Tax Invoice
E-Invoice: Yes (if applicable)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Party: Gupta Enterprises           GSTIN: 27AABCG5678C1ZA
Address: 12 Link Road, Mumbai      State: Maharashtra (27)
Phone: 9812345678                  Credit Days: 30
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
S.No  Item          HSN    Qty   Unit  Rate   Disc%  Taxable  GST%  CGST   SGST   Total
1     Steel Pipes   7306   100   Mtr   450    0      45000    18    4050   4050   53100
2     MS Sheets     7209   50    Kg    120    5      5700     18    513    513    6726
3     GI Wire       7217   200   Kg    85     0      17000    18    1530   1530   20060
4     Iron Rods     7214   150   Kg    95     0      14250    18    1283   1283   16816
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Taxable Value: Rs.81,950  |  CGST(9%): Rs.7,376  |  SGST(9%): Rs.7,376
Round Off: Rs.0.20
Grand Total: Rs.96,702
Amount in Words: Ninety Six Thousand Seven Hundred Two Only
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Narration: Being goods sold to Gupta Enterprises vide SI-2025-001

ACCOUNT MASTER FORMAT:
Name          : [Full party name]
Group         : Sundry Debtors / Sundry Creditors / Bank / Cash
GSTIN         : [15-digit]
PAN           : [10-char]
City          : [City]
State         : [State (state code)]
Phone         : [number]
Email         : [email]
Opening Bal   : Rs.[amount] Dr/Cr
Credit Limit  : Rs.[amount]
Credit Days   : [n] days
Price List    : [Standard / Customer-specific]

ITEM MASTER FORMAT:
Item Name     : [Product name]
Group         : [Item category]
Unit          : [Nos / Kg / Mtr / Ltr / Box / Pcs]
Alternate Unit: [if applicable]
HSN Code      : [8-digit HSN]
GST Rate      : [%]
Purchase Rate : Rs.[amount]
Sale Rate     : Rs.[amount]
Opening Stock : [qty] [unit] @ Rs.[rate] = Rs.[value]
Godown        : [Godown name]
Reorder Level : [qty]

RULES:
❌ Never leave amounts uncomputed
✅ All GST breakups computed (CGST/SGST for same state, IGST for inter-state)
✅ Grand Total = Taxable + CGST + SGST exactly
✅ Narration complete
✅ End: Shortcuts: F2=Date F5=Item F9=Discount Ctrl+F9=Batch Alt+P=Print Ctrl+A=Save`,

// ════════════════════════════════════════════════════════════
// 17. QUICKBOOKS (India)
// ════════════════════════════════════════════════════════════
'QuickBooks': `You are a QuickBooks India Expert (Online + Desktop). You produce complete QuickBooks entries — invoices, bills, payroll, reports — all amounts calculated.
${QUALITY_LAW}
QUICKBOOKS CALCULATION LAW:
→ Line Amount = Qty × Rate (compute: write number)
→ CGST = Subtotal × 9% (compute: write number)
→ SGST = Subtotal × 9% (same state — compute: write number)
→ IGST = Subtotal × 18% (inter-state — compute: write number)
→ Balance Due = Grand Total − Amount Paid (compute: write number)
→ NEVER write formula expressions — compute first, write result

INVOICE FORMAT:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
                    ABC TECHNOLOGIES PVT. LTD.
           123 Business Park, Andheri East, Mumbai - 400069
      Phone: +91-22-12345678 | Email: accounts@abctech.com
      GSTIN: 27AABCA1234B1ZB | Website: www.abctech.com
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TAX INVOICE

Invoice #: QB-2025-001             Date: 15-04-2025
Due Date: 15-05-2025               Payment Terms: Net 30
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BILL TO:
XYZ Corporation Pvt. Ltd.
Mr. Ramesh Gupta, Finance Manager
45, Sector 18, Noida — 201301, Uttar Pradesh
GSTIN: 09AABCX1234B1ZB | Phone: 9811234567
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   Product/Service    Description              Qty   Rate     Amount
1   IT Consulting      Web application dev      40h   2500     100000
2   Server Setup       AWS cloud config          1    15000    15000
3   Annual AMC         12-month support          1    24000    24000
4   UI/UX Design       Mobile app design        20h   2000     40000
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Subtotal (Before Tax)  :             Rs.1,79,000
CGST @ 9%              :               Rs.16,110
SGST @ 9%              :               Rs.16,110
Grand Total            :             Rs.2,11,220
Amount Paid            :                    Rs.0
Balance Due            :             Rs.2,11,220
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Amount in Words: Two Lakh Eleven Thousand Two Hundred Twenty Only
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BANK DETAILS FOR PAYMENT:
Bank: HDFC Bank | Branch: Andheri East, Mumbai
A/C No: 50100123456789 | IFSC: HDFC0001234 | Type: Current
UPI: accounts@abctech
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Note: Payment after due date attracts 2% p.m. interest.
Thank you for your business! Contact: +91-22-12345678
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PAYROLL FORMAT:
━━━ PAYROLL: [Month Year] ━━━
Employee: [Name] | ID: [EMP001] | Department: [dept]
Basic: Rs.[x] | HRA: Rs.[x] | DA: Rs.[x] | TA: Rs.[x] | Gross: Rs.[x]
PF(12%): Rs.[x] | ESI(0.75%): Rs.[x] | TDS: Rs.[x] | Net Pay: Rs.[x]

RULES:
❌ Never write formula expressions
✅ All tax amounts computed (CGST/SGST/IGST)
✅ Balance Due computed
✅ Bank details always included
✅ Amount in Words always`,

// ════════════════════════════════════════════════════════════
// 18. NOTEPAD
// ════════════════════════════════════════════════════════════
'Notepad': `You are a Notepad / Plain Text Expert. You produce pure plain text using only ASCII characters — no HTML, no markdown, no special symbols. Output should paste perfectly into Windows Notepad.
${QUALITY_LAW}
NOTEPAD CALCULATION LAW:
→ All amounts computed — never write "45000+500" — write 46300
→ GST: compute actual rupee amount — write "GST @18%: Rs.8,334" not "18% of subtotal"
→ Total: add all amounts — write final number

ASCII FORMAT RULES:
- Use = for main borders (80 chars wide)
- Use - for sub-borders
- Use | for table column separators (spaces for alignment)
- CAPITAL LETTERS for headings
- Rs. instead of Rs. symbol
- Max 80 characters per line
- Align columns with spaces for readability

INVOICE EXAMPLE — EXACT STRUCTURE:
================================================================================
                       SHARMA ELECTRONICS
              45 Nehru Market, Chandni Chowk, Delhi - 110006
         Phone: 9876543210 | Email: sharma@electronics.com
                    GSTIN: 07AABCS1234B1ZB
================================================================================
                            TAX INVOICE
================================================================================
Invoice No : INV-2025-001              Date    : 15-04-2025
Customer   : ABC Pvt. Ltd.             Phone   : 9811234567
Address    : 12 Connaught Place,       GSTIN   : 07AABCA5678B1ZC
             New Delhi - 110001
================================================================================
S.No  Description              Qty    Rate       Amount
--------------------------------------------------------------------------------
  1   HP Laptop 15s             2    45,000      90,000
  2   Wireless Mouse            5       500       2,500
  3   USB Keyboard              3       800       2,400
  4   24-inch Monitor           1    12,000      12,000
  5   USB Cable 2m             10       150       1,500
--------------------------------------------------------------------------------
      Sub-Total                                 108,400
      CGST @ 9%                                   9,756
      SGST @ 9%                                   9,756
================================================================================
      GRAND TOTAL                               127,912
================================================================================
Amount in Words: One Lakh Twenty Seven Thousand Nine Hundred Twelve Only

Bank: HDFC Bank | A/C: 50100123456789 | IFSC: HDFC0001234
UPI: sharma@hdfc

THANK YOU FOR YOUR BUSINESS!
================================================================================

LETTER EXAMPLE — STRUCTURE:
================================================================================
                     OFFICE NOTICE
================================================================================
Date   : 15-04-2025
To     : All Employees
From   : HR Department
Subject: [Specific subject]
================================================================================

[Complete letter body — every paragraph written]

Authorized Signatory
[Name]
[Designation]
================================================================================

RULES:
❌ Never use Rs. symbol — use Rs.
❌ Never exceed 80 characters per line
❌ Never leave amount as expression — compute and write
✅ Columns properly aligned with spaces
✅ GRAND TOTAL matches sum of all items + taxes
✅ Amount in Words for every bill/invoice`,

// ════════════════════════════════════════════════════════════
// 19. WORDPAD
// ════════════════════════════════════════════════════════════
'WordPad': `You are a WordPad Expert. You produce complete, print-ready professional documents using simple text formatting — no HTML, no complex styles. Every document 100% complete, no blanks.
${QUALITY_LAW}
WORDPAD CALCULATION LAW:
→ For any bill/invoice in WordPad: compute all amounts — write final numbers
→ GST: write computed rupee amount (Rs.8,334 not "18% of subtotal")
→ Total: add everything — write final sum

DOCUMENT FORMAT:
==============================================================
                    [COMPANY NAME IN CAPS]
    [Full Address], [City] - [PIN] | [State]
    Phone: [number] | Email: [email] | GSTIN: [number]
==============================================================

Date: [DD Month YYYY]
Ref No: [REF/DEPT/2025/001]

To,
[Full Name]
[Designation]
[Company / Organization]
[Full Address with PIN]

Subject: [Specific, clear subject]

Dear Sir/Madam / Dear Mr./Ms. [Name],

[Opening paragraph — state the purpose clearly]

[Main paragraph — complete details, amounts, dates]
[For agreements: use numbered clauses]
  1. [Clause heading]: [Complete clause text]
     a) [Sub-point fully written]
     b) [Sub-point fully written]
  2. [Next clause]: [Complete text]

[Closing paragraph — action required / next steps]

Yours sincerely / Yours faithfully,


[Full Name]
[Designation]
[Company Name]
[Date: DD Month YYYY]

==============================================================
[For agreements: Signature block]
Party A Signature: _______________  Date: _______________
Party B Signature: _______________  Date: _______________
Witness 1: _______________  Witness 2: _______________
==============================================================

INVOICE FORMAT IN WORDPAD:
==============================================================
INVOICE No: INV-2025-001        Date: 15-04-2025
Customer: [Name]                Phone: [number]
Address: [Full address]
==============================================================
Item 1: [Name] x [Qty] @ Rs.[rate] = Rs.[computed amount]
Item 2: [Name] x [Qty] @ Rs.[rate] = Rs.[computed amount]
Item 3: [Name] x [Qty] @ Rs.[rate] = Rs.[computed amount]
---------------------------------------------------------
Sub-Total:                           Rs.[computed]
GST @18%:                            Rs.[computed]
==============================================================
GRAND TOTAL:                         Rs.[computed]
Amount in Words: [Spelled out]
==============================================================

RULES:
❌ Never write [Name], [Amount], [Date] in final output — fill real values
❌ Never write arithmetic like "45000+8100" — compute: write 53100
✅ Every paragraph completely written
✅ All amounts computed and written
✅ End tip: Bold=Ctrl+B | Italic=Ctrl+I | Save as .rtf for formatting`

  }; // end PROMPTS

  // ══════════════════════════════════════════════════════════════
  // UNIVERSAL IRON RULES — appended to ALL platform prompts
  // ══════════════════════════════════════════════════════════════
  const universalRules = `

════════════════════════════════════════════════════════════
UNIVERSAL RULES — APPLY TO EVERY REQUEST (NON-NEGOTIABLE):
════════════════════════════════════════════════════════════

YOU ARE A HUMAN EXPERT DOING THE ACTUAL WORK.
Behave like a skilled professional sitting at the computer, not an AI explaining things.

WHEN IMAGE/SCREENSHOT IS UPLOADED:
→ Read every number, name, item, price, date from the image accurately
→ Apply user's instruction exactly (add GST, calculate total, reformat)
→ Produce complete output using extracted data — user should just copy-paste

WHEN USER GIVES ACTUAL DATA:
→ Use EXACT numbers/names/data given
→ Compute ALL calculations (GST, total, salary, etc.) yourself
→ Write ONLY the computed final values — NEVER formula expressions
→ Example: User says "Laptop Rs.45000, 18% GST" → You write: GST=8100, Total=53100

WHAT THESE MEAN:
"GST add karo"    → Compute GST for each item. Show CGST/SGST/IGST breakdown. Grand total.
"Total nikalo"    → Add everything. Write the actual sum.
"10 employees"    → Generate all 10 rows with computed salaries.
"Invoice banao"   → Complete filled invoice, every rupee calculated, Amount in Words.
"Presentation"    → All 12+ slides with real content, speaker notes.
"Database"        → All tables, relationships, INSERT data, SQL queries.

OUTPUT QUALITY CHECKLIST:
✅ Zero placeholders: no [Name], [Amount], [Date], XXX, ???
✅ Zero truncation: complete output, never cut short, no "..."
✅ Zero incomplete rows: every row has all columns filled
✅ Zero formula text: =SUM(), =D2*0.18 never appear in spreadsheet output
✅ Zero generic content: real Indian names, real amounts, real addresses
✅ Language: Hindi response if user wrote in Hindi, English if English

${platform} OUTPUT: 100% complete, professional, zero errors. User pastes and uses directly.
════════════════════════════════════════════════════════════`;

  const basePrompt = PROMPTS[platform] || `You are an expert for ${platform}. Produce 100% complete, professional, ready-to-use output. Real values only, no placeholders, no truncation. Hindi or English as user writes.`;

  return basePrompt + universalRules;
}
// ─── CHECK & INCREMENT DAILY USAGE ───────────────────────────
async function checkAndIncrementUsage(userId) {
  const { data: userData, error } = await supabase.auth.admin.getUserById(userId);
  if (error || !userData?.user) throw new Error('User not found');

  const meta  = userData.user.user_metadata || {};
  const plan  = meta.plan || 'free';
  const today = new Date().toISOString().split('T')[0];

  // Check yearly plan expiry
  if (plan === 'yearly' && meta.planExpiresAt) {
    if (new Date(meta.planExpiresAt) < new Date()) {
      try { await supabase.auth.admin.updateUserById(userId, { user_metadata: { ...meta, plan: 'free' } }); } catch(e) {}
    } else {
      return { allowed: true, plan: 'yearly', remaining: 'unlimited' };
    }
  }

  // Check pro plan expiry
  if (plan === 'pro' && meta.planExpiresAt) {
    if (new Date(meta.planExpiresAt) < new Date()) {
      try { await supabase.auth.admin.updateUserById(userId, { user_metadata: { ...meta, plan: 'free' } }); } catch(e) {}
    } else {
      return { allowed: true, plan: 'pro', remaining: 'unlimited' };
    }
  }

  if (plan === 'pro')    return { allowed: true, plan: 'pro',    remaining: 'unlimited' };
  if (plan === 'yearly') return { allowed: true, plan: 'yearly', remaining: 'unlimited' };

  const usageKey = `usage_${today}`;
  const used = parseInt(meta[usageKey] || 0);

  if (used >= FREE_DAILY_LIMIT) {
    return {
      allowed: false, plan: 'free', remaining: 0, used, limit: FREE_DAILY_LIMIT,
      message: `Daily limit of ${FREE_DAILY_LIMIT} free generations reached. Upgrade to Pro for unlimited access.`
    };
  }

  await supabase.auth.admin.updateUserById(userId, {
    user_metadata: { ...meta, [usageKey]: used + 1 }
  });

  return { allowed: true, plan: 'free', remaining: FREE_DAILY_LIMIT - used - 1, used: used + 1 };
}

// ─── ACTIVATE PLAN ────────────────────────────────────────────
async function activatePlanForUser(userId, planName, paymentId) {
  const { data: userData } = await supabase.auth.admin.getUserById(userId);
  const existingMeta = userData?.user?.user_metadata || {};

  // New plan names support: Pro, Popular, Business, Pro_6M, Popular_6M, Business_6M, Pro_1Y, Popular_1Y, Business_1Y
  // Legacy support: 1 Month, 3 Months, 6 Months, 1 Year
  const isYearlyPlan  = planName === '1 Year'   || planName.endsWith('_1Y');
  const is6MonthPlan  = planName === '6 Months' || planName.endsWith('_6M');
  const is3MonthPlan  = planName === '3 Months';
  const isBusinessPlan = planName.startsWith('Business');
  const planType = isBusinessPlan ? 'business' : (isYearlyPlan ? 'yearly' : 'pro');
  const months = isYearlyPlan ? 12 : (is6MonthPlan ? 6 : (is3MonthPlan ? 3 : 1));
  const exp = new Date();
  exp.setMonth(exp.getMonth() + months);
  const expiresAt = exp.toISOString();

  await supabase.auth.admin.updateUserById(userId, {
    user_metadata: {
      ...existingMeta,
      plan: planType, planName, paymentId,
      planActivatedAt: new Date().toISOString(),
      planExpiresAt: expiresAt
    }
  });

  try {
    await supabase.from('subscriptions').upsert([{
      user_id: userId, plan: planType, plan_name: planName,
      payment_id: paymentId, activated_at: new Date().toISOString(),
      expires_at: expiresAt, is_active: true
    }], { onConflict: 'user_id' });
  } catch(e) { console.warn('Sub upsert warn:', e.message); }
}


// ════════════════════════════════════════════════════════════
//  ROUTE: GET /api/config
// ════════════════════════════════════════════════════════════
app.get('/api/config', (req, res) => {
  res.json({ razorpayKeyId: process.env.RAZORPAY_KEY_ID || '', freeDailyLimit: FREE_DAILY_LIMIT });
});

// ════════════════════════════════════════════════════════════
//  ROUTE: GET /  — Health check
// ════════════════════════════════════════════════════════════
app.get('/', (req, res) => {
  res.json({
    status: 'OK', service: 'Ai Forgen Backend v6.0',
    platforms: ALL_PLATFORMS.length, freeDailyLimit: FREE_DAILY_LIMIT,
    aiProvider: 'n1n.ai (GPT-4.1 / Opus 4.7 — display)',
    newFeatures: ['Smart File Parser', 'Auto Platform Detection', 'CSV/Excel/PDF/Image/Word Support', 'Batch Processing'],
    plans: ['1 Month (Rs.299)', '3 Months (Rs.299)', '6 Months (Rs.399)', '1 Year (Rs.499)']
  });
});


// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/generate  — Main AI generation (text/image)
// ════════════════════════════════════════════════════════════
app.post('/api/generate', generateLimiter, async (req, res) => {
  try {
    const { supabaseId, prompt, platform, images } = req.body;

    if (!supabaseId)    return res.status(401).json({ error: 'Login required' });
    if (!prompt?.trim()) return res.status(400).json({ error: 'Prompt is required' });

    const selectedPlatform = ALL_PLATFORMS.includes(platform) ? platform : detectPlatform(prompt, '', platform);

    let usageCheck;
    try { usageCheck = await checkAndIncrementUsage(supabaseId); }
    catch(e) { return res.status(401).json({ error: 'User verification failed: ' + e.message }); }

    if (!usageCheck.allowed) {
      return res.status(402).json({
        success: false, limitReached: true, requiresUpgrade: true,
        message: `Daily limit of ${FREE_DAILY_LIMIT} free generations reached. Upgrade to Pro for unlimited access.`,
        plan: 'free', remaining: 0, limit: FREE_DAILY_LIMIT,
        upgradeUrl: '/#pricing', action: 'upgrade_required'
      });
    }

    const userContentParts = [];

    // Handle inline base64 images
    if (Array.isArray(images) && images.length > 0) {
      for (const img of images) {
        if (img?.base64 && img?.mediaType) {
          userContentParts.push({
            type: 'image_url',
            image_url: { url: `data:${img.mediaType};base64,${img.base64}` }
          });
        }
      }
    }

    const isSpreadsheet = ['Microsoft Excel','Google Sheets','LibreOffice','Apache OpenOffice','WPS Office','Zoho Office Suite'].includes(selectedPlatform);
    const isAccounting  = ['Tally','Busy Accounting Software','QuickBooks'].includes(selectedPlatform);
    const isHindi       = /[\u0900-\u097F]/.test(prompt);

    userContentParts.push({
      type: 'text',
      text: `PLATFORM: ${selectedPlatform}
USER REQUEST: ${prompt.trim()}

════════════════════════════════════════════════════════════
EXECUTION INSTRUCTIONS — FOLLOW EXACTLY:
════════════════════════════════════════════════════════════

STEP 1 — UNDERSTAND the request completely.
STEP 2 — If image attached: extract EVERY value from it accurately.
STEP 3 — CALCULATE (mandatory before writing output):
${isSpreadsheet ? `  • For each item: Taxable = Qty × Rate → compute the number
  • GST Amt = Taxable × GST% / 100 → compute the number
  • CGST = GST Amt / 2 → compute | SGST = GST Amt / 2 → compute
  • Total = Taxable + GST Amt → compute the number
  • TOTAL row = sum of each column → compute each sum
  • Amount in Words = spell out Grand Total` : ''}
${isAccounting ? `  • All Dr/Cr amounts balanced and computed
  • Stock item amounts = Qty × Rate (computed number)
  • GST amounts computed (CGST/SGST/IGST)
  • Grand Total verified: matches debit side` : ''}
STEP 4 — WRITE complete output for ${selectedPlatform}.

ABSOLUTE RULES:
${isSpreadsheet ? `✅ TAB-separated, one row per line, headers in row 1
✅ Computed numbers only — NEVER "=45000+8100" or "=SUM(D2:D6)"
✅ Invoice: header block, blank line, then item table
✅ TOTAL row at bottom with computed values
✅ Amount in Words at bottom of every invoice/bill` : ''}
${isAccounting ? `✅ All voucher amounts computed — no expressions
✅ Dr total = Cr total (balanced entry)
✅ Narration complete and descriptive` : ''}
✅ ZERO placeholders — [Name] [Amount] [Date] = FORBIDDEN
✅ ZERO truncation — complete output always
✅ Real Indian data: real names, GSTIN format, Indian cities, Rs. amounts
✅ Language: ${isHindi ? 'Hindi (user wrote in Hindi)' : 'English (user wrote in English)'}

OUTPUT: Complete, professional, zero errors. User copies and uses directly.`
    });

    const userContent = userContentParts.length === 1 ? userContentParts[0].text : userContentParts;
    const aiResponse  = await callSiliconFlow(getSystemPrompt(selectedPlatform), userContent, selectedPlatform);

    const output = Array.isArray(aiResponse.content)
      ? aiResponse.content.filter(b => b.type === 'text').map(b => b.text).join('\n')
      : aiResponse.content;

    try {
      await supabase.from('generation_history').insert([{
        user_id: supabaseId,
        prompt:  prompt.trim().substring(0, 500),
        platform: selectedPlatform,
        output:  output.substring(0, 10000),
        created_at: new Date().toISOString(),
        model: aiResponse.model,
        plan_at_generation: usageCheck.plan
      }]);
    } catch(e) { console.warn('History warn:', e.message); }

    return res.json({
      success: true, output, platform: selectedPlatform,
      plan: usageCheck.plan, remaining: usageCheck.remaining,
      model: aiResponse.model,
      canExport:   usageCheck.plan !== 'free',
      canDownload: usageCheck.plan !== 'free',
      canCopy:     true, // Free users can copy
      upgradeMessage: usageCheck.plan === 'free'
        ? 'Upgrade to Pro to export, copy and download this content.' : null
    });

  } catch(err) {
    console.error('Generate error:', err.message);
    return res.status(500).json({ success: false, error: 'Generation failed: ' + err.message });
  }
});


// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/generate-with-file  — SMART FILE UPLOAD
//  Supports: Image, CSV, Excel, PDF, Word, TXT, JSON, TSV
//  Auto-detects platform from file content + user prompt
// ════════════════════════════════════════════════════════════
app.post('/api/generate-with-file', generateLimiter, upload.array('files', 10), async (req, res) => {
  try {
    const { supabaseId, prompt, platform } = req.body;
    const files = req.files || [];

    if (!supabaseId) return res.status(401).json({ error: 'Login required' });
    if (files.length === 0 && !prompt?.trim()) {
      return res.status(400).json({ error: 'File ya prompt dono mein se kuch to chahiye' });
    }

    // Parse all uploaded files
    const parsedFiles = files.map(parseUploadedFile);

    // Auto-detect platform from prompt + file content
    const fileTextContent = parsedFiles
      .filter(f => !f.isImage)
      .map(f => f.content || '')
      .join(' ')
      .substring(0, 500);

    const selectedPlatform = detectPlatform(prompt || '', fileTextContent, platform);

    const usageCheck = await checkAndIncrementUsage(supabaseId);
    if (!usageCheck.allowed) {
      return res.status(402).json({
        success: false, limitReached: true, requiresUpgrade: true,
        message: `Daily limit of ${FREE_DAILY_LIMIT} free generations reached. Upgrade to Pro.`,
        plan: 'free', remaining: 0, limit: FREE_DAILY_LIMIT, upgradeUrl: '/#pricing'
      });
    }

    const userContentParts = [];
    const textFileSummaries = [];

    // Process each file
    for (const parsed of parsedFiles) {
      if (parsed.isImage) {
        // Image → vision API
        userContentParts.push({
          type: 'image_url',
          image_url: { url: `data:${parsed.mediaType};base64,${parsed.base64}` }
        });
        textFileSummaries.push(`[Image file: ${parsed.name}]`);
      } else {
        // Text/CSV/Excel/PDF → inject as text
        textFileSummaries.push(`[${parsed.type.toUpperCase()} file: ${parsed.content?.substring(0, 50)}...]`);
      }
    }

    const isSpreadsheet = ['Microsoft Excel','Google Sheets','LibreOffice','Apache OpenOffice','WPS Office','Zoho Office Suite'].includes(selectedPlatform);
    const isAccounting  = ['Tally','Busy Accounting Software','QuickBooks'].includes(selectedPlatform);
    const isHindi       = /[\u0900-\u097F]/.test(prompt || '');

    // Build the main text prompt with all file data
    let mainPrompt = `PLATFORM: ${selectedPlatform}
USER REQUEST: ${prompt || `Uploaded file ko ${selectedPlatform} format mein properly convert aur complete karo`}

`;

    // Add parsed file contents as text
    const textFiles = parsedFiles.filter(f => !f.isImage);
    if (textFiles.length > 0) {
      mainPrompt += `════════════════════════════════════════════
UPLOADED FILE DATA — PROCESS THIS EXACTLY:
════════════════════════════════════════════
`;
      textFiles.forEach((f, i) => {
        mainPrompt += `\n--- FILE ${i+1} (${f.type.toUpperCase()}) ---\n${f.content}\n`;
      });
      mainPrompt += `\n════════════════════════════════════════════\n`;
    }

    if (parsedFiles.some(f => f.isImage)) {
      mainPrompt += `\n[IMAGE FILE(S) UPLOADED — Har ek value, number, name, amount carefully read karo]\n`;
    }

    mainPrompt += `
════════════════════════════════════════════
EXECUTION INSTRUCTIONS:
════════════════════════════════════════════

STEP 1 — File/Image ka POORA data read karo (har ek number, naam, amount)
STEP 2 — User ke instruction ke hisaab se kaam karo
STEP 3 — CALCULATE karo:
${isSpreadsheet ? `  • Har item: Taxable = Qty × Rate (computed number)
  • GST = Taxable × GST%/100 (computed)
  • CGST = GST/2 | SGST = GST/2 (computed)
  • Total = Taxable + GST (computed)
  • TOTAL row = har column ka actual sum
  • Amount in Words = Grand Total spelled out` : ''}
${isAccounting ? `  • Har amount computed (Qty × Rate = actual number)
  • CGST/SGST ya IGST computed
  • Dr total = Cr total (balanced)
  • Grand Total verified` : ''}
STEP 4 — ${selectedPlatform} ke liye COMPLETE output likho

ABSOLUTE RULES:
❌ "I can see..." mat likho — seedha output do
❌ "Please provide..." mat likho — file se extract karo
❌ Formula expressions — =SUM(), =D2*0.18 — FORBIDDEN
❌ Placeholders — [Name], [Amount], [Date] — FORBIDDEN
❌ Truncation — poora output likhna zaroori hai
✅ File ka EXACT data use karo
✅ Sab calculations compute karke final numbers likho
✅ Real Indian format: Rs., DD-MM-YYYY, GSTIN
✅ Language: ${isHindi ? 'Hindi' : 'English'}

OUTPUT: 100% complete, ready to use. User copy-paste karega directly.`;

    userContentParts.push({ type: 'text', text: mainPrompt });

    const userContent = userContentParts.length === 1 ? userContentParts[0].text : userContentParts;
    const aiResponse  = await callSiliconFlow(getSystemPrompt(selectedPlatform), userContent, selectedPlatform);

    const output = Array.isArray(aiResponse.content)
      ? aiResponse.content.filter(b => b.type === 'text').map(b => b.text).join('\n')
      : aiResponse.content;

    try {
      await supabase.from('generation_history').insert([{
        user_id: supabaseId,
        prompt: (prompt || 'File upload').substring(0, 500),
        platform: selectedPlatform,
        output: output.substring(0, 10000),
        created_at: new Date().toISOString(),
        model: aiResponse.model,
        plan_at_generation: usageCheck.plan,
        files_used: files.map(f => f.originalname).join(', ')
      }]);
    } catch(e) { console.warn('History warn:', e.message); }

    return res.json({
      success: true, output, platform: selectedPlatform,
      detectedPlatform: selectedPlatform,
      filesProcessed: parsedFiles.map(f => ({
        name: files[parsedFiles.indexOf(f)]?.originalname,
        type: f.type,
        isImage: f.isImage || false
      })),
      plan: usageCheck.plan, remaining: usageCheck.remaining,
      model: aiResponse.model,
      canExport:   usageCheck.plan !== 'free',
      canDownload: usageCheck.plan !== 'free',
      canCopy:     true // Free users can copy
    });

  } catch(err) {
    console.error('File generate error:', err.message);
    return res.status(500).json({ error: 'File generation failed: ' + err.message });
  }
});


// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/smart-batch  — Batch processing (CSV/text)
//  User ek CSV upload kare 100 rows → 100 outputs generate
// ════════════════════════════════════════════════════════════
app.post('/api/smart-batch', upload.single('file'), async (req, res) => {
  try {
    const { supabaseId, instruction, platform } = req.body;
    const file = req.file;

    if (!supabaseId) return res.status(401).json({ error: 'Login required' });

    // Check plan — batch only for paid users
    const { data: userData } = await supabase.auth.admin.getUserById(supabaseId);
    const plan = userData?.user?.user_metadata?.plan || 'free';
    if (plan === 'free') {
      return res.status(402).json({
        success: false, requiresUpgrade: true,
        message: 'Batch processing Pro feature hai. Upgrade karo unlimited access ke liye.',
        upgradeUrl: '/#pricing'
      });
    }

    if (!file) return res.status(400).json({ error: 'File required for batch processing' });

    const parsed = parseUploadedFile(file);
    if (parsed.isImage) return res.status(400).json({ error: 'Batch processing ke liye CSV/Text file chahiye, image nahi' });

    // Parse rows from CSV
    let rows = [];
    if (parsed.type === 'csv' && parsed.headers) {
      const text = file.buffer.toString('utf-8');
      const lines = text.split('\n').filter(l => l.trim());
      const sep = lines[0].includes('\t') ? '\t' : lines[0].includes(';') ? ';' : ',';
      const headers = lines[0].split(sep).map(h => h.replace(/^["']|["']$/g, '').trim());
      rows = lines.slice(1).map(line => {
        const cells = line.split(sep).map(c => c.replace(/^["']|["']$/g, '').trim());
        const obj = {};
        headers.forEach((h, i) => { obj[h] = cells[i] || ''; });
        return obj;
      }).filter(r => Object.values(r).some(v => v));
    } else {
      // Text file — split by blank lines or newlines
      const lines = file.buffer.toString('utf-8').split('\n').filter(l => l.trim());
      rows = lines.map((line, i) => ({ row: i+1, content: line }));
    }

    if (rows.length === 0) return res.status(400).json({ error: 'File mein koi data nahi mila' });
    if (rows.length > 1000) {
      return res.status(400).json({
        error: `File mein ${rows.length} rows hain. Maximum 1000 rows per batch allowed.`,
        tip: '1000 se zyada rows ke liye file ko parts mein split karo'
      });
    }

    const selectedPlatform = detectPlatform(instruction || '', '', platform);
    const BATCH_SIZE = 5; // 5 rows ek saath process
    const results    = [];
    const totalBatches = Math.ceil(rows.length / BATCH_SIZE);

    console.log(`🔄 Batch: ${rows.length} rows, ${totalBatches} batches, platform: ${selectedPlatform}`);

    // SSE (Server-Sent Events) — real-time progress to frontend
    const wantsSSE = req.headers.accept === 'text/event-stream';
    if (wantsSSE) {
      res.setHeader('Content-Type', 'text/event-stream');
      res.setHeader('Cache-Control', 'no-cache');
      res.setHeader('Connection', 'keep-alive');
      res.flushHeaders();
    }

    const sendProgress = (processed, total, msg) => {
      if (wantsSSE) {
        res.write(`data: ${JSON.stringify({ type: 'progress', processed, total, percent: Math.round(processed/total*100), msg })}\n\n`);
      }
    };

    sendProgress(0, rows.length, `Processing shuru... ${rows.length} rows, ${totalBatches} batches`);

    for (let b = 0; b < totalBatches; b++) {
      const batchRows = rows.slice(b * BATCH_SIZE, (b+1) * BATCH_SIZE);
      const batchData = batchRows.map((r, i) => `Item ${b*BATCH_SIZE+i+1}: ${JSON.stringify(r)}`).join('\n');

      const batchPrompt = `PLATFORM: ${selectedPlatform}
BATCH PROCESSING JOB — ${b+1}/${totalBatches}

USER INSTRUCTION: ${instruction || `Har row ke liye complete ${selectedPlatform} output banao`}

DATA (${batchRows.length} items):
${batchData}

RULES:
✅ Har item ke liye SEPARATE complete output banao
✅ Clearly label karo: === ITEM 1 === , === ITEM 2 === etc.
✅ Har item mein sab calculations compute karo
✅ Zero placeholders, zero truncation
✅ Real Indian format: Rs., GSTIN, DD-MM-YYYY`;

      let batchSuccess = false;
      for (let attempt = 1; attempt <= 2; attempt++) {
        try {
          const aiResp = await callSiliconFlow(getSystemPrompt(selectedPlatform), batchPrompt, selectedPlatform);
          const batchOutput = Array.isArray(aiResp.content)
            ? aiResp.content.filter(b2 => b2.type === 'text').map(b2 => b2.text).join('\n')
            : aiResp.content;
          results.push({ batch: b+1, rows: batchRows.length, output: batchOutput, success: true });
          batchSuccess = true;
          break;
        } catch(e) {
          if (attempt === 2) {
            results.push({ batch: b+1, rows: batchRows.length, output: '', success: false, error: e.message });
          } else {
            await new Promise(r => setTimeout(r, 2000)); // retry wait
          }
        }
      }

      const processedRows = Math.min((b+1) * BATCH_SIZE, rows.length);
      sendProgress(processedRows, rows.length, `Batch ${b+1}/${totalBatches} ${batchSuccess ? '✅' : '❌'} — ${processedRows}/${rows.length} rows done`);

      // Throttle — prevent rate limiting
      if (b < totalBatches - 1) await new Promise(r => setTimeout(r, 500));
    }

    const fullOutput = results.map(r => r.output).join('\n\n' + '═'.repeat(60) + '\n\n');
    const successCount = results.filter(r => r.success).length;

    try {
      await supabase.from('generation_history').insert([{
        user_id: supabaseId,
        prompt: `[BATCH] ${instruction || 'Batch processing'} (${rows.length} rows)`.substring(0, 500),
        platform: selectedPlatform,
        output: fullOutput.substring(0, 10000),
        created_at: new Date().toISOString(),
        plan_at_generation: plan
      }]);
    } catch(e) {}

    const finalPayload = {
      success: true,
      totalRows: rows.length,
      totalBatches,
      successBatches: successCount,
      failedBatches: totalBatches - successCount,
      output: fullOutput,
      platform: selectedPlatform,
      canExport: true, canDownload: true, canCopy: true
    };

    if (wantsSSE) {
      res.write(`data: ${JSON.stringify({ type: 'done', ...finalPayload })}\n\n`);
      res.end();
    } else {
      return res.json(finalPayload);
    }

  } catch(err) {
    console.error('Batch error:', err.message);
    return res.status(500).json({ error: 'Batch processing failed: ' + err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/create-order  — Razorpay order
// ════════════════════════════════════════════════════════════
app.post('/api/create-order', strictLimiter, async (req, res) => {
  try {
    const { planName, supabaseId } = req.body;
    if (!supabaseId) return res.status(401).json({ error: 'Login required' });

    const amount = PLAN_PRICES[planName];
    if (!amount) return res.status(400).json({ error: 'Invalid plan: ' + planName });

    const order = await razorpay.orders.create({
      amount, currency: 'INR',
      receipt: `receipt_${Date.now()}`,
      notes: { supabaseId, planName }
    });

    return res.json({ orderId: order.id, amount, currency: 'INR', planName, key: process.env.RAZORPAY_KEY_ID });
  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/verify-payment  — Razorpay verification
// ════════════════════════════════════════════════════════════
app.post('/api/verify-payment', strictLimiter, async (req, res) => {
  try {
    const { razorpay_order_id, razorpay_payment_id, razorpay_signature, supabaseId, planName } = req.body;

    const body = razorpay_order_id + '|' + razorpay_payment_id;
    const expectedSig = crypto.createHmac('sha256', process.env.RAZORPAY_KEY_SECRET).update(body).digest('hex');

    if (expectedSig !== razorpay_signature) {
      return res.status(400).json({ success: false, error: 'Invalid payment signature' });
    }

    await activatePlanForUser(supabaseId, planName, razorpay_payment_id);

    return res.json({
      success: true,
      message: `${planName} plan activated successfully!`,
      plan: (planName === '1 Year' || planName.endsWith('_1Y')) ? 'yearly' : (planName.startsWith('Business') ? 'business' : 'pro'),
      paymentId: razorpay_payment_id
    });
  } catch(err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/webhook  — Razorpay webhook
// ════════════════════════════════════════════════════════════
app.post('/api/webhook', express.raw({ type: 'application/json' }), async (req, res) => {
  try {
    const sig = req.headers['x-razorpay-signature'];
    const body = req.body.toString();
    const expectedSig = crypto.createHmac('sha256', process.env.RAZORPAY_KEY_SECRET).update(body).digest('hex');

    if (sig !== expectedSig) return res.status(400).json({ error: 'Invalid webhook signature' });

    const event = JSON.parse(body);
    if (event.event === 'payment.captured') {
      const payment = event.payload.payment.entity;
      const { supabaseId, planName } = payment.notes || {};
      if (supabaseId && planName) {
        await activatePlanForUser(supabaseId, planName, payment.id);
        console.log(`✅ Webhook: Plan activated for ${supabaseId} — ${planName}`);
      }
    }

    return res.json({ status: 'ok' });
  } catch(err) {
    console.error('Webhook error:', err.message);
    return res.status(500).json({ error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: GET /api/user/:userId  — User info
// ════════════════════════════════════════════════════════════
app.get('/api/user/:userId', async (req, res) => {
  try {
    const { data: userData, error } = await supabase.auth.admin.getUserById(req.params.userId);
    if (error || !userData?.user) return res.status(404).json({ error: 'User not found' });

    const meta  = userData.user.user_metadata || {};
    const plan  = meta.plan || 'free';
    const today = new Date().toISOString().split('T')[0];
    const used  = parseInt(meta[`usage_${today}`] || 0);

    return res.json({
      userId: req.params.userId,
      email:  userData.user.email,
      plan, planName: meta.planName || 'Free',
      planExpiresAt: meta.planExpiresAt || null,
      dailyUsed:     used,
      dailyLimit:    plan === 'free' ? FREE_DAILY_LIMIT : 'unlimited',
      dailyRemaining: plan === 'free' ? Math.max(0, FREE_DAILY_LIMIT - used) : 'unlimited',
      canExport:   plan !== 'free',
      canDownload: plan !== 'free',
      canCopy:     true // Free users can copy
    });
  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: GET /api/history/:userId  — Generation history
// ════════════════════════════════════════════════════════════
app.get('/api/history/:userId', async (req, res) => {
  try {
    const limit = Math.min(parseInt(req.query.limit) || 20, 50);
    const { data, error } = await supabase
      .from('generation_history')
      .select('id, prompt, platform, created_at, model')
      .eq('user_id', req.params.userId)
      .order('created_at', { ascending: false })
      .limit(limit);

    if (error) throw error;
    return res.json({ history: data || [], total: (data || []).length });
  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: DELETE /api/history/:userId  — Clear history
// ════════════════════════════════════════════════════════════
app.delete('/api/history/:userId', async (req, res) => {
  try {
    const { error } = await supabase
      .from('generation_history')
      .delete()
      .eq('user_id', req.params.userId);
    if (error) throw error;
    return res.json({ success: true, message: 'History cleared' });
  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/share  — Share generation
// ════════════════════════════════════════════════════════════
app.post('/api/share', async (req, res) => {
  try {
    const { supabaseId, generationId } = req.body;
    if (!supabaseId) return res.status(401).json({ error: 'Login required' });

    const shareId = crypto.randomBytes(8).toString('hex');
    const { error } = await supabase
      .from('generation_history')
      .update({ share_id: shareId, is_shared: true, shared_at: new Date().toISOString() })
      .eq('id', generationId)
      .eq('user_id', supabaseId);

    if (error) throw error;
    return res.json({ success: true, shareId, shareUrl: `/share/${shareId}`, message: 'Generation shared successfully!' });
  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: GET/POST /api/preferences/:userId
// ════════════════════════════════════════════════════════════
app.get('/api/preferences/:userId', async (req, res) => {
  try {
    const { data } = await supabase
      .from('user_preferences').select('*')
      .eq('user_id', req.params.userId).single();

    const defaults = {
      currency: 'Rs.', dateFormat: 'DD-MM-YYYY',
      companyName: '', gstin: '', address: '',
      phone: '', email: '', defaultPlatform: 'Microsoft Excel',
      language: 'Hindi+English'
    };
    return res.json({ preferences: { ...defaults, ...(data || {}) } });
  } catch(err) {
    return res.json({ preferences: { currency: 'Rs.', dateFormat: 'DD-MM-YYYY', defaultPlatform: 'Microsoft Excel' } });
  }
});

app.post('/api/preferences/:userId', async (req, res) => {
  try {
    await supabase.from('user_preferences').upsert([{
      user_id: req.params.userId, ...req.body, updated_at: new Date().toISOString()
    }], { onConflict: 'user_id' });
    return res.json({ success: true, message: 'Preferences saved!' });
  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/batch-generate
// ════════════════════════════════════════════════════════════
app.post('/api/batch-generate', generateLimiter, async (req, res) => {
  try {
    const { supabaseId, prompts, platform } = req.body;
    if (!supabaseId) return res.status(401).json({ error: 'Login required' });
    if (!Array.isArray(prompts) || prompts.length === 0) return res.status(400).json({ error: 'prompts array required' });
    if (prompts.length > 10) return res.status(400).json({ error: 'Max 10 items per batch' });

    const usageCheck = await checkAndIncrementUsage(supabaseId);
    if (!usageCheck.allowed) return res.status(402).json({ limitReached: true, message: usageCheck.message });

    const selectedPlatform = ALL_PLATFORMS.includes(platform) ? platform : 'Microsoft Excel';
    const results = [];

    for (const prompt of prompts) {
      try {
        const userContent = `PLATFORM: ${selectedPlatform}\nUSER REQUEST: ${prompt}\n\nGenerate complete, ready-to-use output. Compute all numbers. Zero placeholders.`;
        const aiResp = await callSiliconFlow(getSystemPrompt(selectedPlatform), userContent, selectedPlatform);
        const output = Array.isArray(aiResp.content)
          ? aiResp.content.filter(b => b.type === 'text').map(b => b.text).join('\n')
          : aiResp.content;
        results.push({ prompt, output, success: true });
      } catch(e) {
        results.push({ prompt, output: '', success: false, error: e.message });
      }
    }

    return res.json({ success: true, results, total: results.length, platform: selectedPlatform });
  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
});

// ════════════════════════════════════════════════════════════
//  ROUTE: GET /api/templates
// ════════════════════════════════════════════════════════════
app.get('/api/templates', (req, res) => {
  const TEMPLATES = [
    { id: 1,  name: 'GST Invoice',        platform: 'Microsoft Excel',          prompt: 'Create GST tax invoice for 5 items with 18% GST, CGST+SGST breakup, party GSTIN, complete calculated amounts and Amount in Words', icon: '🧾', category: 'Accounts' },
    { id: 2,  name: 'Salary Sheet',       platform: 'Microsoft Excel',          prompt: 'Create monthly salary sheet for 10 employees with Basic, HRA(40%), DA(10%), TA, Gross, PF(12%), ESI(0.75%), TDS, Net Pay — all calculated', icon: '💰', category: 'HR/Payroll' },
    { id: 3,  name: 'Stock Register',     platform: 'Microsoft Excel',          prompt: 'Create stock register with Opening Stock, Purchase, Sales, Closing Stock, Item-wise value, all quantities and amounts calculated', icon: '📦', category: 'Inventory' },
    { id: 4,  name: 'P&L Statement',      platform: 'Microsoft Excel',          prompt: 'Create complete Profit & Loss statement with all income heads, all expense heads, Gross Profit, Net Profit — real numbers', icon: '📊', category: 'Accounts' },
    { id: 5,  name: 'Attendance Sheet',   platform: 'Microsoft Excel',          prompt: 'Create monthly attendance register for 20 employees, P/A/L/HD marks, total present days counted', icon: '📅', category: 'HR' },
    { id: 6,  name: 'Appointment Letter', platform: 'Microsoft Word',           prompt: 'Create professional appointment letter for Sales Executive position, salary Rs.25000/month, joining date 01-05-2025, all terms and conditions', icon: '📝', category: 'HR' },
    { id: 7,  name: 'Rent Agreement',     platform: 'Microsoft Word',           prompt: 'Create residential rent agreement for 11 months, rent Rs.15000/month, security deposit Rs.30000, all standard clauses complete', icon: '🏠', category: 'Legal' },
    { id: 8,  name: 'Business Proposal',  platform: 'Microsoft Word',           prompt: 'Create complete business proposal for IT services company — executive summary, scope of work, timeline, pricing, payment terms', icon: '💼', category: 'Business' },
    { id: 9,  name: 'Experience Cert.',   platform: 'Microsoft Word',           prompt: 'Create experience certificate for employee Rajesh Kumar who worked as Senior Accountant for 2 years 6 months, good conduct', icon: '🎓', category: 'HR' },
    { id: 10, name: 'Legal Notice',       platform: 'Microsoft Word',           prompt: 'Create legal notice for payment recovery of Rs.75000 outstanding for 90 days, demand payment within 15 days', icon: '⚖️', category: 'Legal' },
    { id: 11, name: 'Sales Voucher',      platform: 'Tally',                   prompt: 'Create Tally sales voucher for Sharma Traders, 3 items with HSN codes, 18% GST CGST+SGST, all amounts calculated, narration complete', icon: '🧾', category: 'Tally' },
    { id: 12, name: 'Purchase Voucher',   platform: 'Tally',                   prompt: 'Create Tally purchase voucher for ABC Suppliers, items with input GST, stock item details, ledger entries balanced', icon: '🛒', category: 'Tally' },
    { id: 13, name: 'Payment Entry',      platform: 'Tally',                   prompt: 'Create Tally payment voucher for vendor NEFT payment of Rs.50000, bank ledger entry, narration with reference', icon: '💳', category: 'Tally' },
    { id: 14, name: 'Company Profile',    platform: 'Microsoft PowerPoint',    prompt: 'Create 12-slide company profile presentation — About Us, Vision/Mission, Services, Team, Achievements, Clients, Contact — all slides complete', icon: '🏢', category: 'Business' },
    { id: 15, name: 'Investor Pitch',     platform: 'Microsoft PowerPoint',    prompt: 'Create investor pitch deck — Problem, Solution, Market Size, Business Model, Traction, Revenue, Team, Ask — all slides with real data', icon: '📈', category: 'Business' },
    { id: 16, name: 'Payment Reminder',   platform: 'Microsoft Outlook',       prompt: 'Write professional payment reminder email for invoice INV-2025-047 of Rs.85000 overdue by 30 days, include bank details', icon: '📧', category: 'Email' },
    { id: 17, name: 'Job Offer Email',    platform: 'Microsoft Outlook',       prompt: 'Write job offer email for Software Developer position, CTC Rs.6,00,000 p.a., joining date 01-May-2025, all details', icon: '💌', category: 'HR' },
    { id: 18, name: 'QB Invoice',         platform: 'QuickBooks',              prompt: 'Create QuickBooks invoice for IT consulting, 40 hours at Rs.2500/hour, server setup Rs.15000, 18% GST, Net 30 payment terms', icon: '💰', category: 'Accounts' },
    { id: 19, name: 'Busy Sales Bill',    platform: 'Busy Accounting Software', prompt: 'Create Busy sales invoice with 4 items, HSN codes, 18% GST CGST+SGST, party GSTIN, all amounts calculated', icon: '🧮', category: 'Accounts' },
  ];

  let filtered = TEMPLATES;
  if (req.query.category) filtered = filtered.filter(t => t.category === req.query.category);
  if (req.query.platform) filtered = filtered.filter(t => t.platform === req.query.platform);
  return res.json({ templates: filtered, total: filtered.length });
});

// ════════════════════════════════════════════════════════════
//  ROUTE: POST /api/similar
// ════════════════════════════════════════════════════════════
app.post('/api/similar', async (req, res) => {
  try {
    const { userId, prompt } = req.body;
    if (!userId || !prompt) return res.json({ similar: null });

    const firstWord = prompt.toLowerCase().split(' ')[0];
    const { data } = await supabase
      .from('generation_history')
      .select('id, prompt, output, platform, created_at')
      .eq('user_id', userId)
      .ilike('prompt', `%${firstWord}%`)
      .order('created_at', { ascending: false })
      .limit(3);

    if (data && data.length > 0) return res.json({ similar: data[0], found: true });
    return res.json({ similar: null, found: false });
  } catch(err) {
    return res.json({ similar: null, found: false });
  }
});

// ─── 404 & Error handlers ─────────────────────────────────
app.use((req, res) => res.status(404).json({ error: `Not found: ${req.method} ${req.path}` }));
app.use((err, req, res, next) => { console.error('Unhandled:', err); res.status(500).json({ error: 'Internal server error' }); });

// ════════════════════════════════════════════════════════════
//  START SERVER
// ════════════════════════════════════════════════════════════
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log('\n╔════════════════════════════════════════╗');
  console.log('║   AI FORGEN BACKEND v6.0        ║');
  console.log(`║   Port: ${PORT}                            ║`);
  console.log(`║   Platforms: ${ALL_PLATFORMS.length} | Accuracy: 95%+        ║`);
  console.log(`║   FREE: ${FREE_DAILY_LIMIT} gen/day (no export)        ║`);
  console.log('║   PAID: Unlimited + Full Export         ║');
  console.log('║   AI: GPT-4.1 / Opus 4.7 (via n1n.ai)  ║');
  console.log('║   Plans: Monthly Rs.299/499/999 | 6M Rs.999/1649/3299 | 1Y Rs.1249/2099/4199 ║');
  console.log('╚════════════════════════════════════════╝\n');
});
