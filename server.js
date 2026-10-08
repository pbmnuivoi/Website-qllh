
const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, 'school-app');
const DB_DIR = path.join(ROOT, 'db');
const DB_FILE = process.env.DB_FILE || path.join(DB_DIR, 'CSDL_EXCEL_TONG.xlsx');
const SEED_DB_FILE = process.env.SEED_DB_FILE || path.join(DB_DIR, 'CSDL_EXCEL_TONG.xlsx');
fs.mkdirSync(path.dirname(DB_FILE), { recursive: true });

app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true, limit: '20mb' }));
app.use((req,res,next) => {
  const origin = req.headers.origin;
  if (origin) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Session-Id');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  }
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});
app.use((req,res,next) => {
  if (req.path.startsWith('/api/') || /\\.(html|js|css|json)$/.test(req.path)) {
    res.setHeader('Cache-Control','no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma','no-cache');
    res.setHeader('Expires','0');
  }
  next();
});
app.use(express.static(PUBLIC_DIR, { etag: false, maxAge: 0 }));

const DEFAULT_USERS = [
  { username: 'Admin', password: '@Zz123456', role: 'admin', name: 'Quản trị viên' }
];
const DEFAULT_CLASSES = [
  '1A','1B','1C','1D','1E','2A','2B','2C','2D','2E',
  '3A','3B','3C','3D','3E','4A','4B','4C','4D','4E',
  '5A','5B','5C','5D','5E'
];
const DEFAULT_SUBJECTS = ['Tin học', 'Công nghệ'];
const DEFAULT_SETTINGS = { schoolName:'Trường Tiểu học Núi Voi', ubndName:'UBND phường Chi Lăng' };

let db = {
  users: DEFAULT_USERS,
  classes: DEFAULT_CLASSES,
  subjects: DEFAULT_SUBJECTS,
  students: {},
  scores: {},
  learningEvaluations: {},
  learningCards: {},
  settings: { ...DEFAULT_SETTINGS }
};
let sessions = new Map();
let writeChain = Promise.resolve();

function clone(v) { return JSON.parse(JSON.stringify(v)); }

function cleanStudentName(name, id) {
  if (!name) return 'Chưa có tên';
  let n = String(name).trim();
  if (id && n.startsWith(id)) n = n.slice(id.length).trim();
  n = n.replace(/^\d{8,}\s+/, '').trim();
  return n || 'Chưa có tên';
}
function normalizeStudents(list) {
  return (Array.isArray(list) ? list : []).map((st, i) => {
    const id = String(st.id || st.ma || '').trim();
    return { id, name: cleanStudentName(st.name || st.hoTen || '', id), gender: String(st.gender || st.gioiTinh || st['Giới tính'] || '').trim(), stt: st.stt || i + 1 };
  }).filter(s => s.id);
}
function sortClasses(a,b) {
  const na = parseInt(a), nb = parseInt(b);
  return na !== nb ? na - nb : String(a).localeCompare(String(b));
}

function ensureScoreShape() {
  const subjects = db.subjects || DEFAULT_SUBJECTS;
  db.scores = db.scores || {};
  Object.values(db.students || {}).flat().forEach(st => {
    if (!db.scores[st.id]) db.scores[st.id] = {};
    if (typeof db.scores[st.id].points === 'number') {
      const old = db.scores[st.id];
      db.scores[st.id] = {};
      subjects.forEach(sub => {
        db.scores[st.id][sub] = {
          points: sub === subjects[0] ? old.points : 10,
          history: sub === subjects[0] ? (old.history || []) : []
        };
      });
    }
    subjects.forEach(sub => {
      if (!db.scores[st.id][sub] || typeof db.scores[st.id][sub].points !== 'number') {
        db.scores[st.id][sub] = { points: 10, history: [] };
      }
    });
  });
}

function normalizeDb(input) {
  const d = input || {};
  db.users = Array.isArray(d.users) && d.users.length ? d.users : DEFAULT_USERS;
  db.classes = Array.isArray(d.classes) && d.classes.length ? d.classes : DEFAULT_CLASSES;
  db.classes = [...new Set(db.classes.map(x => String(x).trim()).filter(Boolean))].sort(sortClasses);
  db.subjects = Array.isArray(d.subjects) && d.subjects.length ? d.subjects.map(String) : DEFAULT_SUBJECTS;
  db.students = (d.students && !Array.isArray(d.students)) ? d.students : {};
  Object.keys(db.students).forEach(cls => db.students[cls] = normalizeStudents(db.students[cls]));
  db.scores = (d.scores && typeof d.scores === 'object' && !Array.isArray(d.scores)) ? d.scores : {};
  db.learningEvaluations = (d.learningEvaluations && typeof d.learningEvaluations === 'object' && !Array.isArray(d.learningEvaluations)) ? d.learningEvaluations : {};
  db.learningCards = (d.learningCards && typeof d.learningCards === 'object' && !Array.isArray(d.learningCards)) ? d.learningCards : {};
  db.settings = { ...DEFAULT_SETTINGS, ...(d.settings && typeof d.settings === 'object' ? d.settings : {}) };
  ensureScoreShape();
  return db;
}

function readWorkbook() {
  if (!fs.existsSync(DB_FILE)) return null;
  const wb = XLSX.readFile(DB_FILE);
  const out = { users: [], classes: [], subjects: [], students: {}, scores: {}, learningEvaluations: {}, learningCards: {}, settings: { ...DEFAULT_SETTINGS } };

  function rows(sheet) {
    if (!wb.Sheets[sheet]) return [];
    return XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: '' });
  }
  rows('USERS').forEach(r => out.users.push({
    username: String(r.username || ''), password: String(r.password || ''),
    role: String(r.role || 'admin'), name: String(r.name || '')
  }));
  rows('CLASSES').forEach(r => { if (r.classCode) out.classes.push(String(r.classCode)); });
  rows('SUBJECTS').forEach(r => { if (r.subject) out.subjects.push(String(r.subject)); });
  rows('STUDENTS').forEach(r => {
    const cls = String(r.classCode || '').trim();
    if (!cls) return;
    if (!out.students[cls]) out.students[cls] = [];
    out.students[cls].push({
      id: String(r.id || '').trim(),
      name: String(r.name || '').trim(),
      stt: r.stt || out.students[cls].length + 1,
      gender: String(r.gender || r.gioiTinh || r['Giới tính'] || '').trim()
    });
  });
  rows('SETTINGS').forEach(r => {
    const key=String(r.key || '').trim();
    if (key) out.settings[key]=String(r.value || '');
  });
  rows('LEARNING_EVALUATIONS').forEach(r => {
    const sid=String(r.studentId || '').trim(), subject=String(r.subject || '').trim();
    if (!sid || !subject) return;
    if (!out.learningEvaluations[sid]) out.learningEvaluations[sid] = {};
    let htt=[], ht=[];
    try { htt = r.htt ? JSON.parse(String(r.htt)) : []; } catch (_) {}
    try { ht = r.ht ? JSON.parse(String(r.ht)) : []; } catch (_) {}
    out.learningEvaluations[sid][subject] = {
      htt: Array.isArray(htt) ? htt.slice(0,6) : [],
      ht: Array.isArray(ht) ? ht.slice(0,6) : []
    };
  });
  rows('LEARNING_CARDS').forEach(r => {
    const sid=String(r.studentId || '').trim(), subject=String(r.subject || '').trim(), card=String(r.cardType || '').trim();
    if (!sid || !subject || !card) return;
    if (!out.learningCards[sid]) out.learningCards[sid] = {};
    if (!out.learningCards[sid][subject]) out.learningCards[sid][subject] = {};
    out.learningCards[sid][subject][card] = Number(r.quantity || 0);
  });
  rows('SCORES').forEach(r => {
    const sid = String(r.studentId || '').trim();
    const subject = String(r.subject || '').trim();
    if (!sid || !subject) return;
    if (!out.scores[sid]) out.scores[sid] = {};
    let history = [];
    try { history = r.history ? JSON.parse(String(r.history)) : []; } catch (_) {}
    out.scores[sid][subject] = {
      points: Number(r.points ?? 10),
      history: Array.isArray(history) ? history : []
    };
  });
  return normalizeDb(out);
}

function writeWorkbook(nextDb) {
  const data = normalizeDb(clone(nextDb));
  const wb = XLSX.utils.book_new();

  const users = data.users.map(u => ({
    username: u.username, password: u.password, role: u.role, name: u.name || ''
  }));
  const classes = data.classes.map(classCode => ({ classCode }));
  const subjects = data.subjects.map(subject => ({ subject }));
  const students = [];
  Object.keys(data.students).sort(sortClasses).forEach(classCode => {
    (data.students[classCode] || []).forEach(st => students.push({
      classCode, stt: st.stt, id: st.id, name: st.name, gender: st.gender || ''
    }));
  });
  const settings = Object.keys(data.settings || {}).map(key => ({ key, value:String(data.settings[key] ?? '') }));
  const learningEvaluations = [];
  Object.keys(data.learningEvaluations || {}).forEach(studentId => {
    Object.keys(data.learningEvaluations[studentId] || {}).forEach(subject => {
      const e = data.learningEvaluations[studentId][subject] || {};
      learningEvaluations.push({ studentId, subject, htt: JSON.stringify(Array.isArray(e.htt) ? e.htt.slice(0,6) : []), ht: JSON.stringify(Array.isArray(e.ht) ? e.ht.slice(0,6) : []) });
    });
  });
  const learningCards = [];
  Object.keys(data.learningCards || {}).forEach(studentId => {
    Object.keys(data.learningCards[studentId] || {}).forEach(subject => {
      Object.keys(data.learningCards[studentId][subject] || {}).forEach(cardType => {
        const quantity = Number(data.learningCards[studentId][subject][cardType] || 0);
        if (quantity > 0) learningCards.push({ studentId, subject, cardType, quantity });
      });
    });
  });
  const scores = [];
  Object.keys(data.scores).forEach(studentId => {
    const item = data.scores[studentId] || {};
    Object.keys(item).forEach(subject => {
      if (subject === '_classCode') return;
      const e = item[subject] || {};
      scores.push({
        studentId, subject,
        points: Number(e.points ?? 10),
        history: JSON.stringify(Array.isArray(e.history) ? e.history : [])
      });
    });
  });

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(users), 'USERS');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(classes), 'CLASSES');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(subjects), 'SUBJECTS');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(students), 'STUDENTS');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(scores), 'SCORES');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(learningEvaluations), 'LEARNING_EVALUATIONS');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(learningCards), 'LEARNING_CARDS');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settings), 'SETTINGS');

  if (!fs.existsSync(DB_DIR)) fs.mkdirSync(DB_DIR, { recursive: true });
  const tmp = path.join(path.dirname(DB_FILE), 'CSDL_EXCEL_TONG.tmp.xlsx');
  XLSX.writeFile(wb, tmp);
  if (fs.existsSync(DB_FILE)) fs.unlinkSync(DB_FILE);
  fs.renameSync(tmp, DB_FILE);
  db = data;
}

function persist(nextDb) {
  writeChain = writeChain.then(() => writeWorkbook(nextDb));
  return writeChain;
}

function getCookie(req, name) {
  const raw = String(req.headers.cookie || '');
  const m = raw.split(';').map(x => x.trim()).find(x => x.startsWith(name + '='));
  return m ? decodeURIComponent(m.slice(name.length + 1)) : null;
}
function getSession(req) {
  const headerSid = String(req.headers['x-session-id'] || '').trim();
  const sid = headerSid || getCookie(req, 'school_sid');
  if (!sid) return null;
  return sessions.get(String(sid)) || null;
}
function getSessionId(req) {
  return String(req.headers['x-session-id'] || getCookie(req, 'school_sid') || '').trim() || null;
}
function requireAdmin(req,res) {
  const s = getSession(req);
  if (!s || s.role !== 'admin') {
    res.status(401).json({ ok:false, msg:'Vui lòng đăng nhập Admin.' });
    return null;
  }
  return s;
}

function seedFromJsonFiles() {
  const files = fs.existsSync(path.join(PUBLIC_DIR, 'data'))
    ? fs.readdirSync(path.join(PUBLIC_DIR, 'data')).filter(f => /^students_.*\.json$/i.test(f)) : [];
  files.forEach(file => {
    try {
      const x = JSON.parse(fs.readFileSync(path.join(PUBLIC_DIR, 'data', file), 'utf8'));
      if (x.classCode && Array.isArray(x.students)) {
        db.students[x.classCode] = normalizeStudents(x.students);
        if (!db.classes.includes(x.classCode)) db.classes.push(x.classCode);
      }
    } catch (e) { console.warn('Không đọc được', file, e.message); }
  });
  db.classes.sort(sortClasses);
  ensureScoreShape();
}

function initialize() {
  if (!fs.existsSync(DB_FILE) && SEED_DB_FILE !== DB_FILE && fs.existsSync(SEED_DB_FILE)) {
    fs.mkdirSync(path.dirname(DB_FILE), { recursive: true });
    fs.copyFileSync(SEED_DB_FILE, DB_FILE);
  }
  const existing = readWorkbook();
  if (existing && existing.users && existing.users.length) {
    db = existing;
  } else {
    seedFromJsonFiles();
    writeWorkbook(db);
  }
}
initialize();

// ===== API =====
app.get('/api/bootstrap', (req,res) => {
  try {
    const latest = readWorkbook();
    if (latest) db = latest;
  } catch (e) {
    console.error('Không đọc được CSDL Excel:', e);
    return res.status(500).json({ok:false,msg:'Không thể đọc CSDL_EXCEL_TONG.xlsx: '+e.message});
  }
  const session = getSession(req);
  if (!session) return res.json({ ok:true, requireLogin:true, session:null });
  res.json({ ok:true, data:clone(db), session:clone(session), sessionToken:getSessionId(req) });
});

app.get('/api/session', (req,res) => {
  const session = getSession(req);
  res.json({ ok:true, session: session ? clone(session) : null });
});

app.post('/api/login', (req,res) => {
  const username = String(req.body.username || '').trim();
  const password = String(req.body.password || '');
  const user = db.users.find(u => u.username === username && u.password === password);
  if (!user) return res.status(401).json({ ok:false, msg:'Sai tài khoản hoặc mật khẩu' });
  const sid = crypto.randomBytes(24).toString('hex');
  const session = { username:user.username, name:user.name || user.username, role:user.role, loginAt:Date.now() };
  sessions.set(sid, session);
  res.setHeader('Set-Cookie', `school_sid=${encodeURIComponent(sid)}; HttpOnly; Path=/; SameSite=Lax`);
  res.json({ ok:true, user:clone(session), sessionToken:sid });
});

app.post('/api/logout', (req,res) => {
  const sid = String(req.headers['x-session-id'] || getCookie(req, 'school_sid') || '').trim();
  if (sid) sessions.delete(String(sid));
  res.setHeader('Set-Cookie', 'school_sid=; HttpOnly; Path=/; Max-Age=0; SameSite=Lax');
  res.json({ ok:true });
});

app.post('/api/save', async (req,res) => {
  if (!requireAdmin(req,res)) return;
  try {
    const next = normalizeDb(req.body.data);
    await persist(next);
    res.json({ ok:true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok:false, msg:'Không thể ghi CSDL Excel: ' + e.message });
  }
});

app.get('/api/export', (req,res) => {
  if (!requireAdmin(req,res)) return;
  if (!fs.existsSync(DB_FILE)) return res.status(404).json({ok:false,msg:'Chưa có file CSDL Excel.'});
  res.download(DB_FILE, 'CSDL_EXCEL_TONG.xlsx');
});

app.get('/api/health', (req,res) => res.json({ ok:true, database:'CSDL_EXCEL_TONG.xlsx' }));

app.get('*', (req,res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'index.html'));
});

app.listen(PORT, () => console.log(`School app running at http://localhost:${PORT}`));
