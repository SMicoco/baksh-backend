// Baksh Consulting Limited - backend (cloned from the proven iCoCo backend)
// Express + SQLite + multer (CV uploads) + JWT admin + Resend email + CSV/XLSX export + OpenAI
require('dotenv').config();
const express   = require('express');
const helmet    = require('helmet');
const cors      = require('cors');
const rateLimit = require('express-rate-limit');
const jwt       = require('jsonwebtoken');
const Database  = require('better-sqlite3');
const { Resend } = require('resend');
const multer    = require('multer');
const ExcelJS   = require('exceljs');
const path      = require('path');
const fs        = require('fs');
const crypto    = require('crypto');

const PORT           = process.env.PORT || 3000;
const JWT_SECRET     = process.env.JWT_SECRET || process.env.ADMIN_SECRET || crypto.randomBytes(32).toString('hex');
const ADMIN_USERNAME = process.env.ADMIN_USERNAME || 'admin';
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'changeme';
const RESEND_API_KEY = process.env.RESEND_API_KEY;
const TO_EMAIL       = process.env.TO_EMAIL   || process.env.NOTIFY_EMAIL || 'sm@icocoassociates.co.uk';
const FROM_EMAIL     = process.env.FROM_EMAIL || 'onboarding@resend.dev';
const PUBLIC_URL     = process.env.PUBLIC_URL || 'https://baksh-consulting.onrender.com';
const DATA_DIR       = process.env.DATA_DIR   || path.join(__dirname, 'data');
const DB_PATH        = process.env.DB_PATH    || path.join(DATA_DIR, 'baksh.db');
const CV_DIR         = path.join(DATA_DIR, 'cvs');
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const OPENAI_MODEL   = process.env.OPENAI_MODEL || 'gpt-4o-mini';

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(CV_DIR,   { recursive: true });

// ---- DB ----
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.exec(`
  CREATE TABLE IF NOT EXISTS submissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    form_type TEXT NOT NULL DEFAULT 'contact',
    name TEXT NOT NULL, email TEXT NOT NULL, phone TEXT,
    organisation TEXT, sector TEXT, role TEXT, right_to_work TEXT,
    linkedin TEXT, experience TEXT, availability TEXT, source TEXT, location TEXT,
    cv_filename TEXT, cv_original_name TEXT, cv_size INTEGER,
    message TEXT, ip TEXT, user_agent TEXT,
    notes TEXT DEFAULT '', status TEXT DEFAULT 'new',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  CREATE INDEX IF NOT EXISTS idx_subs_created ON submissions(created_at DESC);
  CREATE INDEX IF NOT EXISTS idx_subs_type ON submissions(form_type);
`);
const insertSub = db.prepare(`INSERT INTO submissions
  (form_type,name,email,phone,organisation,sector,role,right_to_work,linkedin,experience,availability,source,location,
   cv_filename,cv_original_name,cv_size,message,ip,user_agent)
  VALUES (@form_type,@name,@email,@phone,@organisation,@sector,@role,@right_to_work,@linkedin,@experience,@availability,@source,@location,
   @cv_filename,@cv_original_name,@cv_size,@message,@ip,@user_agent)`);
const listSubs   = db.prepare(`SELECT * FROM submissions ORDER BY created_at DESC LIMIT 1000`);
const listByType = db.prepare(`SELECT * FROM submissions WHERE form_type = ? ORDER BY created_at DESC LIMIT 1000`);
const getSub     = db.prepare(`SELECT * FROM submissions WHERE id = ?`);
const updNotes   = db.prepare(`UPDATE submissions SET notes = ? WHERE id = ?`);
const updStatus  = db.prepare(`UPDATE submissions SET status = ? WHERE id = ?`);
const delSub     = db.prepare(`DELETE FROM submissions WHERE id = ?`);
const stats      = db.prepare(`SELECT
  COUNT(*) AS total,
  SUM(CASE WHEN status='new' THEN 1 ELSE 0 END) AS new_count,
  SUM(CASE WHEN status='replied' THEN 1 ELSE 0 END) AS replied,
  SUM(CASE WHEN form_type='contact' THEN 1 ELSE 0 END) AS contacts,
  SUM(CASE WHEN form_type='career' THEN 1 ELSE 0 END) AS careers,
  SUM(CASE WHEN form_type='ai' THEN 1 ELSE 0 END) AS ai,
  SUM(CASE WHEN form_type='crypto' THEN 1 ELSE 0 END) AS crypto,
  SUM(CASE WHEN form_type='robotics' THEN 1 ELSE 0 END) AS robotics
FROM submissions`);

// ---- multer ----
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, CV_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase().slice(0, 10);
    cb(null, 'cv-' + Date.now() + '-' + crypto.randomBytes(6).toString('hex') + ext);
  }
});
const upload = multer({ storage, limits: { fileSize: 10 * 1024 * 1024 } });

// ---- email ----
const resend = RESEND_API_KEY ? new Resend(RESEND_API_KEY) : null;
function esc(s){return String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
async function notifyEmail(sub) {
  if (!resend) { console.warn('[email] skipped (no api key)'); return; }
  const labels = { contact:'New contact enquiry', career:'New career application', enquiry:'New project enquiry' };
  const subjectLabel = labels[sub.form_type] || 'New website submission';
  const dateStr = new Date().toLocaleString('en-GB');
  const lines = [];
  if (sub.organisation) lines.push('<p><strong>Organisation:</strong> ' + esc(sub.organisation) + '</p>');
  if (sub.sector)       lines.push('<p><strong>Area / topic:</strong> ' + esc(sub.sector) + '</p>');
  if (sub.role)         lines.push('<p><strong>Role:</strong> ' + esc(sub.role) + '</p>');
  if (sub.location)     lines.push('<p><strong>Location:</strong> ' + esc(sub.location) + '</p>');
  if (sub.right_to_work)lines.push('<p><strong>Right to work:</strong> ' + esc(sub.right_to_work) + '</p>');
  if (sub.linkedin)     lines.push('<p><strong>LinkedIn / portfolio:</strong> ' + esc(sub.linkedin) + '</p>');
  if (sub.availability) lines.push('<p><strong>Availability:</strong> ' + esc(sub.availability) + '</p>');
  if (sub.cv_original_name) lines.push('<p><strong>CV:</strong> ' + esc(sub.cv_original_name) + ' (' + Math.round((sub.cv_size||0)/1024) + ' KB)</p>');
  const html =
    '<div style="font-family:Inter,Arial,sans-serif;color:#1a1a1a;max-width:600px;margin:0 auto">' +
    '<h2 style="color:#1b3a78;margin-bottom:6px">' + subjectLabel + '</h2>' +
    '<p style="color:#666;font-size:13px;margin-top:0">Baksh Consulting Limited &middot; ' + dateStr + '</p>' +
    '<p><strong>Name:</strong> ' + esc(sub.name) + '</p>' +
    '<p><strong>Email:</strong> <a href="mailto:' + esc(sub.email) + '">' + esc(sub.email) + '</a></p>' +
    (sub.phone ? '<p><strong>Phone:</strong> ' + esc(sub.phone) + '</p>' : '') +
    lines.join('') +
    (sub.message ? '<p><strong>Message:</strong></p><div style="background:#f3f4f6;padding:12px 14px;border-radius:8px;white-space:pre-wrap">' + esc(sub.message) + '</div>' : '') +
    '<hr style="margin:24px 0;border:none;border-top:1px solid #e5e7eb">' +
    '<p style="font-size:12px;color:#666">View in admin: <a href="' + PUBLIC_URL + '/admin">' + PUBLIC_URL + '/admin</a></p>' +
    '</div>';
  try {
    await resend.emails.send({ from: 'Baksh Website <' + FROM_EMAIL + '>', to: [TO_EMAIL], replyTo: sub.email, subject: subjectLabel + ' - ' + sub.name, html });
    console.log('[email] sent for submission', sub.id);
  } catch (e) { console.error('[email] failed:', e.message); }
}

// ---- OpenAI helper ----
async function callOpenAI(messages, maxtok, json) {
  if (!OPENAI_API_KEY) return null;
  const payload = { model: OPENAI_MODEL, max_tokens: maxtok || 700, messages };
  if (json) payload.response_format = { type: 'json_object' };
  const r = await fetch('https://api.openai.com/v1/chat/completions', { method:'POST', headers:{ 'content-type':'application/json', authorization:'Bearer '+OPENAI_API_KEY }, body: JSON.stringify(payload) });
  const j = await r.json();
  if (j && j.error) throw new Error(j.error.message || 'api-error');
  return ((j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '').trim();
}

// ---- app ----
const app = express();
app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true, limit: '2mb' }));
app.set('trust proxy', 1);
const apiLimiter   = rateLimit({ windowMs: 60000, max: 40, standardHeaders: true });
const loginLimiter = rateLimit({ windowMs: 15 * 60000, max: 10, standardHeaders: true });

function requireAdmin(req, res, next) {
  const tok = (req.headers.authorization || '').replace(/^Bearer\s+/i, '') || req.query.token;
  if (!tok) return res.status(401).json({ error: 'No token' });
  try { req.admin = jwt.verify(tok, JWT_SECRET); next(); } catch { return res.status(401).json({ error: 'Bad token' }); }
}

function buildSub(b, req, formType, file) {
  const name = (b.name || ((b.first||'') + ' ' + (b.last||''))).trim();
  return {
    form_type: formType,
    name: name.slice(0, 200),
    email: String(b.email || '').slice(0, 200),
    phone: String(b.phone || '').slice(0, 80),
    organisation: String(b.org || b.organisation || '').slice(0, 200),
    sector: String(b.sector || b.topic || '').slice(0, 160),
    role: String(b.role || '').slice(0, 200),
    right_to_work: String(b.right_to_work || b.rtw || '').slice(0, 200),
    linkedin: String(b.linkedin || b.url || '').slice(0, 400),
    experience: String(b.experience || '').slice(0, 4000),
    availability: String(b.availability || '').slice(0, 200),
    source: String(b.source || '').slice(0, 200),
    location: String(b.location || '').slice(0, 200),
    cv_filename: file ? file.filename : null,
    cv_original_name: file ? file.originalname : null,
    cv_size: file ? file.size : null,
    message: String(b.message || b.comment || '').slice(0, 6000),
    ip: req.ip || '',
    user_agent: (req.headers['user-agent'] || '').slice(0, 400)
  };
}
function saveAndNotify(sub, res) {
  if (!sub.name || !sub.email) return res.status(400).json({ error: 'Missing fields' });
  const r = insertSub.run(sub); sub.id = r.lastInsertRowid;
  notifyEmail(sub);
  res.json({ ok: true, id: sub.id });
}

// ---- public form endpoints ----
// Unified endpoint used by every Baksh form (contact, careers, expertise enquiry)
app.post('/api/submit', apiLimiter, upload.single('cv'), (req, res) => {
  try {
    if (req.body && req.body['bot-field']) return res.json({ ok: true });
    const t = String((req.body && req.body.type) || 'contact');
    const m = { career:'career', ai:'ai', crypto:'crypto', robotics:'robotics', contact:'contact' };
    let formType = m[t];
    if (!formType) { const tp = String((req.body && (req.body.topic||req.body.sector))||'').toLowerCase();
      formType = /crypto/.test(tp) ? 'crypto' : /robot/.test(tp) ? 'robotics' : /(\bai\b|artificial|intelligence)/.test(tp) ? 'ai' : 'contact'; }
    saveAndNotify(buildSub(req.body || {}, req, formType, req.file), res);
  } catch (e) { console.error(e); res.status(500).json({ error: 'Server error' }); }
});
// iCoCo-style endpoints kept for compatibility
app.post('/api/contact', apiLimiter, (req, res) => { try { saveAndNotify(buildSub(req.body, req, 'contact', null), res); } catch (e) { res.status(500).json({ error: 'Server error' }); } });
app.post('/api/careers/vacancy', apiLimiter, upload.single('cv'), (req, res) => { try { saveAndNotify(buildSub(req.body, req, 'career', req.file), res); } catch (e) { res.status(500).json({ error: 'Server error' }); } });

// ---- AI (OpenAI) ----
app.post('/api/ask', apiLimiter, async (req, res) => {
  const q = ((req.body && req.body.question) || '').toString().slice(0, 500).trim();
  if (!q) return res.status(400).json({ error: 'No question' });
  if (!OPENAI_API_KEY) return res.json({ answer: "The assistant isn't configured yet.", sources: [] });
  const sys = "You are the AI assistant for Baksh Consulting Limited, led by Professor Nadeem Baksh, advising on artificial intelligence, cryptocurrency and robotics. Answer concisely in 2-4 short paragraphs. General information only.";
  try { const a = await callOpenAI([{ role:'system', content:sys }, { role:'user', content:q }], 700, false); res.json({ answer: a || "I couldn't answer just now.", sources: [] }); }
  catch (e) { res.json({ answer: "The assistant ran into an issue. Please use the contact page.", sources: [] }); }
});
app.post('/api/ai', apiLimiter, async (req, res) => {
  const d = req.body || {};
  if (!OPENAI_API_KEY) return res.json({ error: 'no-key' });
  const BASE = "You are a senior consultant at Baksh Consulting Limited (AI, cryptocurrency, robotics). Concise, professional, specific. General guidance only.";
  const jobj = t => { const m = t && t.match(/\{[\s\S]*\}/); return m ? JSON.parse(m[0]) : null; };
  try {
    if (d.task === 'strategy') { const t = await callOpenAI([{role:'system',content:BASE+' Return ONLY JSON {"summary":"..","opportunities":[{"title":"..","desc":".."}],"roadmap":{"now":[".."],"next":[".."],"later":[".."]},"risks":[".."],"quickwins":[".."]}'},{role:'user',content:'AI strategy for: '+(d.org||'an organisation')+', objective '+(d.objective||'improvement')+', stage '+(d.stage||'exploring')}],1100,true); return res.json(jobj(t)||{summary:t}); }
    if (d.task === 'usecases') { const t = await callOpenAI([{role:'system',content:BASE+' Return ONLY JSON {"items":[{"title":"..","desc":"..","impact":"High|Medium","effort":"Low|Medium|High"}]} (6 items).'},{role:'user',content:'6 use cases of AI/crypto/robotics for the '+(d.sector||'general')+' sector.'}],900,true); return res.json(jobj(t)||{items:[]}); }
    if (d.task === 'summarise') { const t = await callOpenAI([{role:'system',content:BASE+' Return ONLY JSON {"summary":"..","points":["..",".."]}'},{role:'user',content:(d.text||'').slice(0,6000)}],700,true); return res.json(jobj(t)||{summary:t,points:[]}); }
    if (d.task === 'explain') { const t = await callOpenAI([{role:'system',content:BASE+' Explain in plain language, 2-3 sentences. Return only the text.'},{role:'user',content:(d.text||'').slice(0,600)}],400,false); return res.json({text:t}); }
    if (d.task === 'proposal') { const t = await callOpenAI([{role:'system',content:BASE+" Four labelled paragraphs: 'Understanding:', 'Approach:', 'What to watch:', 'First step:'. ~180 words. Indicative only."},{role:'user',content:'Challenge: '+(d.challenge||'').slice(0,900)}],700,false); return res.json({text:t}); }
    if (d.task === 'assess') { const t = await callOpenAI([{role:'system',content:BASE+' Return ONLY JSON {"summary":"..","strengths":["..",".."],"gaps":["..",".."],"actions":["..","..",".."]}'},{role:'user',content:(d.sector||'AI')+' readiness, score '+(d.score||0)+'/100. Answers: '+JSON.stringify(d.answers||{})}],800,true); return res.json(jobj(t)||{summary:t}); }
    res.status(400).json({ error: 'unknown task' });
  } catch (e) { res.json({ error: 'ai-failed' }); }
});
const FEEDS = [ {cat:'AI',source:'AI News',url:'https://www.artificialintelligence-news.com/feed/'}, {cat:'Crypto',source:'Cointelegraph',url:'https://cointelegraph.com/rss'}, {cat:'Robotics',source:'The Robot Report',url:'https://www.therobotreport.com/feed/'} ];
function cleanX(s){return String(s).replace(/<!\[CDATA\[/g,'').replace(/\]\]>/g,'').replace(/<[^>]+>/g,'').replace(/&amp;/g,'&').replace(/&#0?39;/g,"'").replace(/&quot;/g,'"').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/\s+/g,' ').trim();}
function parseFeed(xml,cat,source){const out=[];let atom=false;let b=xml.split(/<item[\s>]/i);if(b.length<2){b=xml.split(/<entry[\s>]/i);atom=true;}for(let i=1;i<b.length&&out.length<8;i++){const t=b[i].match(/<title[^>]*>([\s\S]*?)<\/title>/i);const title=t?cleanX(t[1]):'';let link='';if(atom){const l=b[i].match(/<link[^>]*href="([^"]+)"/i);link=l?l[1]:'';}else{const l=b[i].match(/<link[^>]*>([\s\S]*?)<\/link>/i);link=l?cleanX(l[1]):'';}if(title&&link)out.push({category:cat,title,link,source});}return out;}
let NEWS={t:0,d:null};
app.get('/api/news', async (req, res) => {
  try { if (NEWS.d && Date.now()-NEWS.t < 1800000) return res.json(NEWS.d);
    let items=[]; await Promise.all(FEEDS.map(async f=>{try{const r=await fetch(f.url,{headers:{'User-Agent':'BakshNewsBot/1.0'}});const x=await r.text();parseFeed(x,f.cat,f.source).forEach(it=>items.push(it));}catch(e){}}));
    const out={updated:new Date().toISOString(),items:items.slice(0,12)}; NEWS={t:Date.now(),d:out}; res.json(out);
  } catch (e) { res.json({ updated:new Date().toISOString(), items: [] }); }
});

// ---- admin ----
app.post('/api/admin/login', loginLimiter, (req, res) => {
  const { username, password } = req.body || {};
  if ((username === ADMIN_USERNAME || !username) && password === ADMIN_PASSWORD) {
    return res.json({ token: jwt.sign({ u: ADMIN_USERNAME }, JWT_SECRET, { expiresIn: '12h' }) });
  }
  res.status(401).json({ error: 'Invalid credentials' });
});
app.get('/api/admin/stats', requireAdmin, (req, res) => res.json(stats.get()));
app.get('/api/admin/submissions', requireAdmin, (req, res) => { const type = req.query.type; res.json({ submissions: type ? listByType.all(type) : listSubs.all() }); });
app.put('/api/admin/sub/:id/notes', requireAdmin, (req, res) => { updNotes.run(String(req.body.notes||'').slice(0,4000), Number(req.params.id)); res.json({ ok: true }); });
app.put('/api/admin/sub/:id/status', requireAdmin, (req, res) => { updStatus.run(String(req.body.status||'new').slice(0,30), Number(req.params.id)); res.json({ ok: true }); });
app.delete('/api/admin/sub/:id', requireAdmin, (req, res) => { const row = getSub.get(Number(req.params.id)); if (row && row.cv_filename) { try { fs.unlinkSync(path.join(CV_DIR, row.cv_filename)); } catch (e) {} } delSub.run(Number(req.params.id)); res.json({ ok: true }); });
app.get('/api/admin/cv/:id', requireAdmin, (req, res) => { const row = getSub.get(Number(req.params.id)); if (!row || !row.cv_filename) return res.status(404).send('No CV'); const file = path.join(CV_DIR, row.cv_filename); if (!fs.existsSync(file)) return res.status(404).send('CV missing'); res.download(file, row.cv_original_name || row.cv_filename); });
app.get('/api/admin/export.csv', requireAdmin, (req, res) => {
  const rows = listSubs.all();
  const headers = ['id','created_at','form_type','status','name','email','phone','organisation','sector','role','location','linkedin','message','notes','cv_original_name'];
  const csv = [headers.join(',')].concat(rows.map(r => headers.map(h => '"' + String(r[h]||'').replace(/"/g,'""').replace(/\n/g,' ') + '"').join(','))).join('\n');
  res.set('Content-Type','text/csv').set('Content-Disposition','attachment; filename="baksh-submissions.csv"').send('﻿'+csv);
});
app.get('/api/admin/export.xlsx', requireAdmin, async (req, res) => {
  const rows = listSubs.all();
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Submissions');
  ws.columns = [ {header:'ID',key:'id',width:6},{header:'Date',key:'created_at',width:18},{header:'Form',key:'form_type',width:12},{header:'Status',key:'status',width:10},{header:'Name',key:'name',width:22},{header:'Email',key:'email',width:28},{header:'Phone',key:'phone',width:16},{header:'Organisation',key:'organisation',width:22},{header:'Area/Topic',key:'sector',width:16},{header:'Role',key:'role',width:18},{header:'Location',key:'location',width:18},{header:'LinkedIn',key:'linkedin',width:28},{header:'Message',key:'message',width:50},{header:'Notes',key:'notes',width:30},{header:'CV',key:'cv_original_name',width:28} ];
  rows.forEach(r => ws.addRow(r)); ws.getRow(1).font = { bold: true };
  res.set('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.set('Content-Disposition','attachment; filename="baksh-submissions.xlsx"');
  await wb.xlsx.write(res); res.end();
});

// ---- static site + admin ----
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'admin.html')));
app.use(express.static(__dirname, { extensions: ['html'] }));

app.listen(PORT, () => {
  console.log('Baksh backend listening on :' + PORT + '  data: ' + DATA_DIR);
  if (!RESEND_API_KEY) console.warn('[warn] RESEND_API_KEY not set: emails disabled');
});
