require('dotenv').config();

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { Pool } = require('pg');
const XLSX = require('xlsx');

const PORT = Number(process.env.PORT || 8080);
const HOST = process.env.HOST || '0.0.0.0';
const IS_PROD = process.env.NODE_ENV === 'production';
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, 'data');
const BACKUP_DIR = path.join(ROOT, 'backups');
fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(BACKUP_DIR, { recursive: true });

if (!process.env.DATABASE_URL) {
  throw new Error('Missing required environment variable: DATABASE_URL');
}

const db = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: IS_PROD ? { rejectUnauthorized: false } : false,
  max: Number(process.env.DB_POOL_MAX || 10),
  idleTimeoutMillis: 30000,
  connectionTimeoutMillis: 10000
});

db.on('error', (err) => {
  console.error('Unexpected PostgreSQL pool error:', err.message);
});

const ROLES = ['Administrator'];
const WRITE_ROLES = new Set(['Administrator']);
const PAYMENT_ROLES = new Set(['Administrator']);
const ADMIN_ROLES = new Set(['Administrator']);
let excelImportInProgress = false;

function now() {
  return new Date().toISOString();
}

function uid(p) {
  return p + '-' + Date.now().toString(36).toUpperCase() + '-' + crypto.randomBytes(3).toString('hex').toUpperCase();
}

function blankData() {
  return {
    customers: [],
    loans: [],
    schedules: [],
    payments: [],
    blacklist: [],
    notifications: [],
    deletedRecords: [],
    expiredCustomers: [],
    pendingQueue: [],
    settings: {
      appName: 'Loan Management',
      currency: 'INR',
      defaultInterest: 2,
      defaultPenalty: 0,
      reminderDays: [7, 3, 1, 0],
      logoData: '',
      logoEnabled: true
    }
  };
}

function normalizeData(d) {
  const b = blankData();
  const x = d && typeof d === 'object' ? d : {};
  const legacyQueue = Array.isArray(x.settings?.pendingQueue) ? x.settings.pendingQueue : [];
  const schedules = Array.isArray(x.schedules) ? x.schedules : [];
  const rawQueue = Array.isArray(x.pendingQueue) ? x.pendingQueue : [...legacyQueue];
  const validPendingIds = new Set(
    schedules.filter(s => s && s.pendingAddedAt).map(s => String(s.id))
  );

  return {
    customers: Array.isArray(x.customers) ? x.customers : [],
    loans: Array.isArray(x.loans) ? x.loans : [],
    schedules,
    payments: Array.isArray(x.payments) ? x.payments : [],
    blacklist: Array.isArray(x.blacklist) ? x.blacklist : [],
    notifications: Array.isArray(x.notifications) ? x.notifications : [],
    deletedRecords: Array.isArray(x.deletedRecords) ? x.deletedRecords : [],
    expiredCustomers: Array.isArray(x.expiredCustomers) ? x.expiredCustomers : [],
    pendingQueue: rawQueue.map(String).filter(id => validPendingIds.has(id)),
    settings: { ...b.settings, ...(x.settings || {}) }
  };
}

function validDate(v) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v || ''))) return false;
  const d = new Date(v + 'T00:00:00');
  return !Number.isNaN(d.getTime());
}

function mobileOk(v, required = true) {
  const x = String(v || '').trim();
  return (!required && !x) || /^[6-9]\d{9}$/.test(x);
}

function positive(v) {
  return Number.isFinite(Number(v)) && Number(v) > 0;
}

function integrity(data) {
  const e = [], cs = new Set(), ls = new Set(), kh = new Set();
  const customerMobileCounts = new Map();
  const loanMap = new Map();
  const scheduleMap = new Map();
  const principalByLoan = new Map();

  for (const c of data.customers) {
    const id = String(c.id);
    if (cs.has(id)) e.push('Duplicate customer ID: ' + id);
    cs.add(id);
    if (!String(c.firstName || '').trim()) e.push('Customer ' + id + ' has no first name');
    if (!mobileOk(c.mobile, true)) e.push('Customer ' + id + ' has invalid mobile');
    if (!String(c.city || '').trim()) e.push('Customer ' + id + ' has no city');
    if (!String(c.district || '').trim()) e.push('Customer ' + id + ' has no district');
    const m = String(c.mobile || '');
    customerMobileCounts.set(m, (customerMobileCounts.get(m) || 0) + 1);
  }
  for (const [m, count] of customerMobileCounts) {
    if (count > 1) e.push('Duplicate customer mobile: ' + m);
  }

  for (const l of data.loans) {
    const id = String(l.id);
    if (ls.has(id)) e.push('Duplicate loan ID: ' + id);
    ls.add(id);
    loanMap.set(id, l);
    if (!cs.has(String(l.customerId))) e.push('Loan ' + id + ' references missing customer');
    if (!positive(l.amount)) e.push('Loan ' + id + ' has invalid amount');
    if (!validDate(l.startDate)) e.push('Loan ' + id + ' has invalid start date');
    const k = String(l.khataNo || l.legacyKhataNo || '').trim().toLowerCase();
    if (!k) e.push('Loan ' + id + ' missing Khata No');
    else if (kh.has(k)) e.push('Duplicate Khata No: ' + k);
    else kh.add(k);
  }

  for (const sc of data.schedules) {
    const id = String(sc.id);
    scheduleMap.set(id, sc);
    if (!ls.has(String(sc.loanId))) e.push('Schedule ' + id + ' references missing loan');
    if (!validDate(sc.dueDate)) e.push('Schedule ' + id + ' has invalid due date');
  }

  // Validate payment totals and references in linear time. The old
  // implementation repeatedly filtered/scanned all payments for every payment,
  // which became O(n²) with 29k+ payment records and made mutations take many
  // seconds. Accumulate principal by loan in one pass instead.
  for (const p of data.payments) {
    const paymentId = String(p.id);
    const loanId = String(p.loanId);
    const loan = loanMap.get(loanId);
    if (!loan) e.push('Payment ' + paymentId + ' references missing loan');
    if (!validDate(p.date)) e.push('Payment ' + paymentId + ' has invalid date');
    if (loan && validDate(loan.startDate) && validDate(p.date) && p.date < loan.startDate) {
      e.push('Payment ' + paymentId + ' is before loan start date');
    }

    const principal = Number(p.principal || 0);
    const interest = Number(p.interest || 0);
    const penalty = Number(p.penalty || 0);
    const t = principal + interest + penalty;
    if (![principal, interest, penalty, t].every(Number.isFinite) || principal < 0 || interest < 0 || penalty < 0) {
      e.push('Payment ' + paymentId + ' has invalid amounts');
    }
    if (Math.abs(t - Number(p.total || 0)) > 0.01) e.push('Payment ' + paymentId + ' total mismatch');

    if (loan) principalByLoan.set(loanId, (principalByLoan.get(loanId) || 0) + principal);

    if (p.scheduleId) {
      const sc = scheduleMap.get(String(p.scheduleId));
      if (!sc) e.push('Payment ' + paymentId + ' references missing schedule');
      else if (String(sc.loanId) !== loanId) e.push('Payment ' + paymentId + ' schedule does not belong to loan');
    }
  }

  for (const [loanId, principal] of principalByLoan) {
    const loan = loanMap.get(loanId);
    if (loan && principal > Number(loan.amount) + 0.01) {
      e.push('Payments for loan ' + loanId + ' exceed loan principal');
    }
  }

  for (const x of data.expiredCustomers || []) {
    if (!cs.has(String(x.customerId))) e.push('Expired record ' + x.id + ' references missing customer');
  }

  return e;
}
function diff(before, after, key) {
  const b = new Map((before[key] || []).map(x => [String(x.id), x]));
  const a = new Map((after[key] || []).map(x => [String(x.id), x]));
  const out = [];

  for (const [id, x] of a) {
    if (!b.has(id)) out.push(['CREATE', id]);
    else if (JSON.stringify(b.get(id)) !== JSON.stringify(x)) out.push(['UPDATE', id]);
  }

  for (const id of b.keys()) {
    if (!a.has(id)) out.push(['DELETE', id]);
  }

  return out;
}

function changes(before, after) {
  return ['customers', 'loans', 'schedules', 'payments', 'blacklist', 'notifications', 'deletedRecords', 'expiredCustomers', 'pendingQueue']
    .flatMap(k => diff(before, after, k).map(([action, id]) => ({ action, type: k, id })));
}

// PostgreSQL database schema.
async function initDb() {
  await db.query(`
    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      username TEXT UNIQUE NOT NULL,
      mobile TEXT,
      role TEXT NOT NULL,
      password_hash TEXT NOT NULL,
      active BOOLEAN NOT NULL DEFAULT TRUE,
      created_at TIMESTAMPTZ NOT NULL
    )
  `);

  await db.query(`
    CREATE TABLE IF NOT EXISTS sessions (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      token_hash TEXT UNIQUE NOT NULL,
      expires_at TIMESTAMPTZ NOT NULL,
      created_at TIMESTAMPTZ NOT NULL
    )
  `);

  await db.query(`
    CREATE TABLE IF NOT EXISTS user_data (
      user_id TEXT PRIMARY KEY,
      data_json JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL
    )
  `);

  await db.query(`
    CREATE TABLE IF NOT EXISTS app_backups (
      id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
      file_name TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL,
      size_bytes BIGINT
    )
  `);

  await db.query(`
    CREATE TABLE IF NOT EXISTS app_meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    )
  `);

  await db.query('CREATE INDEX IF NOT EXISTS idx_sessions_token_hash ON sessions(token_hash)');
  await db.query('CREATE INDEX IF NOT EXISTS idx_sessions_user_id ON sessions(user_id)');
  await db.query('CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)');

  const metaResult = await db.query(
    'SELECT value FROM app_meta WHERE key = $1',
    ['shared_data_v1']
  );

  const sharedResult = await db.query(
    'SELECT data_json FROM user_data WHERE user_id = $1',
    [SHARED_DATA_ID]
  );

  if (!metaResult.rows.length) {
    const rows = await db.query(
      'SELECT data_json FROM user_data WHERE user_id <> $1',
      [SHARED_DATA_ID]
    );

    const merged = mergeDataSets(rows.rows.map(r => {
      try {
        return typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json;
      } catch {
        return blankData();
      }
    }));

    await db.query(
      `INSERT INTO user_data(user_id, data_json, updated_at)
       VALUES ($1, $2::jsonb, $3)
       ON CONFLICT (user_id) DO UPDATE
       SET data_json = EXCLUDED.data_json,
           updated_at = EXCLUDED.updated_at`,
      [SHARED_DATA_ID, JSON.stringify(merged), now()]
    );

    await db.query(
      `INSERT INTO app_meta(key, value)
       VALUES ($1, $2)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
      ['shared_data_v1', now()]
    );
  } else if (!sharedResult.rows.length) {
    const d = blankData();
    await db.query(
      `INSERT INTO user_data(user_id, data_json, updated_at)
       VALUES ($1, $2::jsonb, $3)`,
      [SHARED_DATA_ID, JSON.stringify(d), now()]
    );
  }

  console.log('PostgreSQL schema initialized successfully.');
}

const SHARED_DATA_ID = '__SHARED__';

function mergeDataSets(items) {
  const out = blankData();
  const maps = {
    customers: new Map(),
    loans: new Map(),
    schedules: new Map(),
    payments: new Map(),
    blacklist: new Map(),
    notifications: new Map(),
    deletedRecords: new Map(),
    expiredCustomers: new Map()
  };
  const pendingIds = new Set();

  for (const raw of items) {
    const d = normalizeData(raw);
    for (const k of Object.keys(maps)) {
      for (const x of d[k] || []) {
        if (x && x.id != null) maps[k].set(String(x.id), x);
      }
    }

    for (const id of d.pendingQueue || []) pendingIds.add(String(id));
    if (!out.settings.logoData && d.settings.logoData) out.settings.logoData = d.settings.logoData;
    if (d.settings.appName) {
      out.settings.appName = d.settings.appName === 'Kissan King Assistance' ? 'Loan Management' : d.settings.appName;
    }
    if (Number.isFinite(Number(d.settings.defaultInterest))) out.settings.defaultInterest = Number(d.settings.defaultInterest);
    if (Number.isFinite(Number(d.settings.defaultPenalty))) out.settings.defaultPenalty = Number(d.settings.defaultPenalty);
    if (Array.isArray(d.settings.reminderDays) && d.settings.reminderDays.length) out.settings.reminderDays = [...d.settings.reminderDays];
    if (d.settings.logoEnabled === false) out.settings.logoEnabled = false;
  }

  for (const k of Object.keys(maps)) out[k] = [...maps[k].values()];
  out.pendingQueue = [...pendingIds];
  return normalizeData(out);
}

async function userByUsername(username) {
  const result = await db.query(
    'SELECT * FROM users WHERE username = $1 AND active = TRUE',
    [username]
  );
  return result.rows[0] || null;
}

async function userById(id) {
  const result = await db.query(
    'SELECT id, name, username, mobile, role, active FROM users WHERE id = $1',
    [id]
  );
  return result.rows[0] || null;
}

async function dataByShared() {
  const result = await db.query(
    'SELECT data_json FROM user_data WHERE user_id = $1',
    [SHARED_DATA_ID]
  );
  return result.rows[0] || null;
}

async function countUsers() {
  const result = await db.query('SELECT COUNT(*)::int AS c FROM users');
  return Number(result.rows[0].c);
}

async function userData() {
  const r = await dataByShared();

  if (!r) {
    const d = blankData();
    await db.query(
      `INSERT INTO user_data(user_id, data_json, updated_at)
       VALUES ($1, $2::jsonb, $3)
       ON CONFLICT (user_id) DO NOTHING`,
      [SHARED_DATA_ID, JSON.stringify(d), now()]
    );
    return d;
  }

  const value = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json;
  return normalizeData(value);
}

async function saveData(data) {
  const normalized = normalizeData(data);
  await db.query(
    `INSERT INTO user_data(user_id, data_json, updated_at)
     VALUES ($1, $2::jsonb, $3)
     ON CONFLICT (user_id) DO UPDATE
     SET data_json = EXCLUDED.data_json,
         updated_at = EXCLUDED.updated_at`,
    [SHARED_DATA_ID, JSON.stringify(normalized), now()]
  );
}

function hashPassword(password) {
  const salt = crypto.randomBytes(16);
  const derived = crypto.scryptSync(password, salt, 64, { N: 16384, r: 8, p: 1 });
  return `scrypt$16384$8$1$${salt.toString('base64url')}$${derived.toString('base64url')}`;
}

function verifyPassword(password, stored) {
  try {
    const [alg, n, r, p, saltB64, hashB64] = String(stored).split('$');
    if (alg !== 'scrypt') return false;
    const salt = Buffer.from(saltB64, 'base64url');
    const hash = Buffer.from(hashB64, 'base64url');
    const derived = crypto.scryptSync(password, salt, hash.length, {
      N: Number(n),
      r: Number(r),
      p: Number(p)
    });
    return crypto.timingSafeEqual(hash, derived);
  } catch {
    return false;
  }
}

function cookies(req) {
  const out = {};
  for (const part of String(req.headers.cookie || '').split(';')) {
    const i = part.indexOf('=');
    if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim());
  }
  return out;
}

async function sessionUser(req) {
  const token = cookies(req).kk_session;
  if (!token) return null;

  const h = crypto.createHash('sha256').update(token).digest('hex');
  const result = await db.query(
    `SELECT u.id, u.name, u.username, u.mobile, u.role, u.active, s.expires_at
     FROM sessions s
     JOIN users u ON u.id = s.user_id
     WHERE s.token_hash = $1`,
    [h]
  );

  const r = result.rows[0];
  if (!r || !r.active || new Date(r.expires_at) <= new Date()) {
    if (r) await db.query('DELETE FROM sessions WHERE token_hash = $1', [h]);
    return null;
  }

  return r;
}

async function setSession(res, userId, remember) {
  const raw = crypto.randomBytes(48).toString('base64url');
  const hash = crypto.createHash('sha256').update(raw).digest('hex');
  const days = remember ? 30 : 1;
  const exp = new Date(Date.now() + days * 86400000).toISOString();

  await db.query(
    `INSERT INTO sessions(id, user_id, token_hash, expires_at, created_at)
     VALUES ($1, $2, $3, $4, $5)`,
    [uid('SES'), userId, hash, exp, now()]
  );

  res.setHeader(
    'Set-Cookie',
    `kk_session=${encodeURIComponent(raw)}; HttpOnly; Path=/; SameSite=None; Max-Age=${days * 86400}${IS_PROD ? '; Secure' : ''}`
  );
}

async function clearSession(req, res) {
  const token = cookies(req).kk_session;
  if (token) {
    const h = crypto.createHash('sha256').update(token).digest('hex');
    await db.query('DELETE FROM sessions WHERE token_hash = $1', [h]);
  }
  res.setHeader(
    'Set-Cookie',
    `kk_session=; HttpOnly; Path=/; SameSite=None; Max-Age=0${IS_PROD ? '; Secure' : ''}`
  );
}

const attempts = new Map();

function rateLimited(req) {
  const ip = req.socket.remoteAddress || 'unknown';
  const t = Date.now();
  const x = attempts.get(ip) || { start: t, count: 0 };
  if (t - x.start > 15 * 60 * 1000) {
    x.start = t;
    x.count = 0;
  }
  x.count++;
  attempts.set(ip, x);
  return x.count > 30;
}

function securityHeaders(res) {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('Referrer-Policy', 'same-origin');
  res.setHeader('Permissions-Policy', 'geolocation=(),camera=(),microphone=()');
  if (IS_PROD) res.setHeader('Strict-Transport-Security', 'max-age=31536000; includeSubDomains');
}

const ALLOWED_ORIGINS = new Set([
  'http://localhost:5500',
  'http://127.0.0.1:5500',
  "https://kissan-king-assistance-frontend.onrender.com"
]);

function corsHeaders(req, res) {
  const origin = req.headers.origin;

  if (origin && ALLOWED_ORIGINS.has(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Vary', 'Origin');
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
    res.setHeader(
      'Access-Control-Allow-Headers',
      'Content-Type, X-Requested-With'
    );
  }
}

function send(res, status, body, headers = {}) {
  securityHeaders(res);
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  for (const [k, v] of Object.entries(headers)) res.setHeader(k, v);
  res.end(JSON.stringify(body));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let b = '';
    req.on('data', c => {
      b += c;
      if (b.length > 10 * 1024 * 1024) req.destroy();
    });
    req.on('end', () => {
      try {
        resolve(b ? JSON.parse(b) : {});
      } catch {
        reject(new Error('Invalid JSON'));
      }
    });
    req.on('error', reject);
  });
}

function pathParts(url) {
  return new URL(url, 'http://localhost').pathname.split('/').filter(Boolean);
}

function excelText(v) {
  return v == null ? '' : String(v).trim();
}

function excelNumber(v) {
  if (v == null || v === '') return 0;
  const n = Number(String(v).replace(/[,₹\s]/g, ''));
  return Number.isFinite(n) ? n : NaN;
}

function excelDate(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
  }
  const x = excelText(v);
  if (/^\d{4}-\d{2}-\d{2}$/.test(x)) return x;

  let m = x.match(/^(\d{1,2})[-\/]([A-Za-z]{3,9})[-\/](\d{4})$/);
  if (m) {
    const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    const mi = months.indexOf(m[2].slice(0, 3).toLowerCase());
    if (mi >= 0) return `${m[3]}-${String(mi + 1).padStart(2, '0')}-${String(Number(m[1])).padStart(2, '0')}`;
  }

  m = x.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return `${m[3]}-${String(Number(m[2])).padStart(2, '0')}-${String(Number(m[1])).padStart(2, '0')}`;

  const n = Number(v);
  if (Number.isFinite(n) && n > 20000 && n < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }

  const d = new Date(x);
  if (!Number.isNaN(d.getTime())) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  return '';
}

function excelRows(sheet) {
  return sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true }) : [];
}

function pick(row, names) {
  for (const n of names) {
    if (Object.prototype.hasOwnProperty.call(row, n) && excelText(row[n]) !== '') return row[n];
  }

  const keys = Object.keys(row);
  const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  for (const n of names) {
    const nn = norm(n);
    const k = keys.find(x => norm(x) === nn);
    if (k != null && excelText(row[k]) !== '') return row[k];
  }
  return '';
}

async function importExcelPayload(base64, mode = 'add') {
  const buf = Buffer.from(String(base64 || ''), 'base64');
  if (!buf.length) throw new Error('Empty Excel file.');

  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const sheets = {};
  wb.SheetNames.forEach(n => sheets[n.toLowerCase()] = excelRows(wb.Sheets[n]));

  const rowsReviews = sheets.reviews || [];
  const rowsCustomers = sheets.customers || [];
  const rowsLoans = sheets.loans || [];
  const rowsPayments = sheets.payments || [];

  if (!rowsReviews.length && !rowsCustomers.length && !rowsLoans.length && !rowsPayments.length) {
    throw new Error('No supported data sheet found. Use Reviews, Customers, Loans or Payments.');
  }

  const target = mode === 'replace' ? blankData() : await userData();
  const issues = [];
  const stats = { customers: 0, loans: 0, payments: 0, schedules: 0, blacklist: 0, expiredCustomers: 0, skipped: 0 };
  const customersById = new Map(target.customers.map(c => [String(c.id), c]));
  const customersByMobile = new Map(target.customers.map(c => [String(c.mobile || ''), c]));
  const loansById = new Map(target.loans.map(l => [String(l.id), l]));
  const loansByKhata = new Map(target.loans.map(l => [String(l.khataNo || l.legacyKhataNo || '').toLowerCase(), l]));
  const paymentsById = new Map(target.payments.map(p => [String(p.id), p]));
  const newLoans = [];
  const importedPayments = [];

  const addCustomer = (r, rowNo) => {
    let id = excelText(pick(r, ['Customer Id', 'Customer ID', 'CustomerId']));
    if (!id) id = `KK-IMP-${String(customersById.size + 1).padStart(6, '0')}`;

    const mobile = excelText(pick(r, ['Personal Mobile', 'Mobile', 'Phone']));
    let c = customersById.get(id) || (!mobile ? '' : customersByMobile.get(mobile));
    const name = excelText(pick(r, ['Name', 'Customer Name'])) || 'Unnamed Customer';
    const parts = name.split(/\s+/).filter(Boolean);
    const next = {
      id,
      firstName: parts[0] || 'Unnamed',
      middleName: parts.length > 2 ? parts.slice(1, -1).join(' ') : '',
      lastName: parts.length > 1 ? parts[parts.length - 1] : '',
      mobile,
      city: excelText(pick(r, ['City'])),
      district: excelText(pick(r, ['District'])),
      address: excelText(pick(r, ['Address'])),
      homeNumber: '',
      pincode: '',
      guarantorName: excelText(pick(r, ['Guarantor Name'])),
      guarantorMobile: excelText(pick(r, ['Guarantor Mobile'])),
      guarantorAddress: excelText(pick(r, ['Guarantor Address'])),
      reference: excelText(pick(r, ['Customer refrence', 'Customer reference', 'Reference'])),
      createdAt: new Date().toISOString()
    };

    if (!next.city) issues.push(`Customer row ${rowNo}: City is required.`);
    if (!next.district) issues.push(`Customer row ${rowNo}: District is required.`);
    if (!/^\d{10}$/.test(mobile)) issues.push(`Customer row ${rowNo}: Mobile must be 10 digits.`);

    if (c) {
      Object.assign(c, next);
      return c;
    }

    target.customers.push(next);
    customersById.set(id, next);
    if (mobile) customersByMobile.set(mobile, next);
    stats.customers++;
    return next;
  };

  const customerRows = rowsCustomers.length ? rowsCustomers : rowsReviews;
  customerRows.forEach((r, i) => addCustomer(r, i + 2));

  const customerFor = (r, rowNo) => {
    let id = excelText(pick(r, ['Customer Id', 'Customer ID', 'CustomerId']));
    let c = id ? customersById.get(id) : null;
    if (!c) {
      const m = excelText(pick(r, ['Personal Mobile', 'Mobile', 'Phone']));
      c = customersByMobile.get(m);
    }
    if (!c) c = addCustomer(r, rowNo);
    return c;
  };

  const loanRows = rowsLoans.length ? rowsLoans : rowsReviews;
  const loanKeys = new Set();

  loanRows.forEach((r, i) => {
    const rowNo = i + 2;
    const c = customerFor(r, rowNo);
    if (!c) return;

    let id = excelText(pick(r, ['Loan Id', 'Loan ID', 'LoanId']));
    const khata = excelText(pick(r, ['Khata No', 'Khata', 'KhataNo']));
    if (!id) id = `KK-LN-IMP-${String(loansById.size + 1).padStart(5, '0')}`;
    if (!khata) {
      issues.push(`Loan row ${rowNo}: Khata No is required.`);
      return;
    }

    const key = id + '|' + khata.toLowerCase();
    if (loanKeys.has(key)) {
      stats.skipped++;
      return;
    }
    loanKeys.add(key);

    let loan = loansById.get(id) || loansByKhata.get(khata.toLowerCase());
    const startDate = excelDate(pick(r, ['Loan Date', 'Start Date', 'Loan Apply Date']));
    const amount = excelNumber(pick(r, ['Loan Amount', 'Amount']));
    const rate = excelNumber(pick(r, ['Loan Interest', 'Interest', 'Interest Rate']));
    const durationRaw = excelNumber(pick(r, ['Duration', 'Duration (months)', 'Months']));
    const duration = Number.isInteger(durationRaw) && durationRaw > 0 ? durationRaw : 12;
    const method = excelText(pick(r, ['Interest Method', 'Method'])) || 'Reducing Balance';
    const emi = excelText(pick(r, ['EMI Option', 'EMI'])) || 'YES';

    const next = {
      id,
      customerId: c.id,
      khataNo: khata,
      loanType: excelText(pick(r, ['Loan Type', 'Type'])) || 'Personal',
      loanAgainst: excelText(pick(r, ['Loan Against', 'Against'])) || 'Personal',
      emiOption: String(emi).toUpperCase() === 'NO' ? 'NO' : 'YES',
      amount,
      interestRate: Number.isFinite(rate) ? rate : 0,
      startDate,
      duration,
      dueDay: startDate ? Number(startDate.slice(8, 10)) : 1,
      method: method === 'Flat Monthly' ? 'Flat Monthly' : 'Reducing Balance',
      penalty: Math.max(0, excelNumber(pick(r, ['Penalty', 'Penalty per overdue EMI'])) || 0),
      notes: excelText(pick(r, ['Notes', 'Remarks'])),
      createdAt: new Date().toISOString()
    };

    if (!validDate(startDate)) issues.push(`Loan row ${rowNo}: invalid loan date.`);
    if (!positive(amount)) issues.push(`Loan row ${rowNo}: loan amount must be greater than zero.`);
    if (khata && loansByKhata.has(khata.toLowerCase()) && !loan) {
      issues.push(`Loan row ${rowNo}: Khata No already exists.`);
    }

    if (loan) {
      Object.assign(loan, next);
    } else {
      target.loans.push(next);
      loansById.set(id, next);
      loansByKhata.set(khata.toLowerCase(), next);
      newLoans.push(next);
      stats.loans++;
    }
  });

  // Legacy Reviews often contains repeated rows for the same loan.
  const paymentRows = rowsPayments.length ? rowsPayments : rowsReviews;
  paymentRows.forEach((r, i) => {
    const rowNo = i + 2;
    let loanId = excelText(pick(r, ['Loan Id', 'Loan ID', 'LoanId']));
    let loan = loanId ? loansById.get(loanId) : null;
    if (!loan) {
      const kh = excelText(pick(r, ['Khata No', 'Khata', 'KhataNo']));
      loan = loansByKhata.get(kh.toLowerCase());
      loanId = loan?.id || '';
    }
    if (!loan) {
      issues.push(`Payment row ${rowNo}: loan not found.`);
      return;
    }

    let id = excelText(pick(r, ['Payment Id', 'Payment ID', 'PaymentId']));
    if (!id) id = `PAY-IMP-${String(paymentsById.size + 1).padStart(6, '0')}`;
    if (paymentsById.has(id)) {
      stats.skipped++;
      return;
    }

    const pd = excelDate(pick(r, ['Payment Date', 'Date']));
    const principal = excelNumber(pick(r, ['Muddal Jama', 'Principal', 'Principal Paid'])) || 0;
    const interest = excelNumber(pick(r, ['Interest Jama', 'Interest', 'Interest Paid'])) || 0;
    const penalty = excelNumber(pick(r, ['Penalty'])) || 0;

    if (!validDate(pd)) issues.push(`Payment row ${rowNo}: invalid payment date.`);
    if (principal < 0 || interest < 0 || penalty < 0) issues.push(`Payment row ${rowNo}: payment amounts cannot be negative.`);
    if (validDate(pd) && validDate(loan.startDate) && pd < loan.startDate) issues.push(`Payment row ${rowNo}: payment date is before loan date.`);

    const p = {
      id,
      loanId: loan.id,
      scheduleId: '',
      date: pd,
      principal,
      interest,
      penalty,
      total: Number((principal + interest + penalty).toFixed(2)),
      mode: excelText(pick(r, ['Payment Mode', 'Mode'])) || 'Cash',
      notes: excelText(pick(r, ['Remarks', 'Notes'])),
      createdAt: new Date().toISOString()
    };

    target.payments.push(p);
    paymentsById.set(id, p);
    importedPayments.push(p);
    stats.payments++;
  });

  const bl = sheets.blacklist_test || sheets.blacklist || [];
  bl.forEach((r, i) => {
    const id = excelText(pick(r, ['Customer Id', 'Customer ID']));
    if (id && customersById.has(id) && !target.blacklist.some(x => String(x.customerId) === id)) {
      target.blacklist.push({
        id: `BL-IMP-${i + 1}-${Date.now()}`,
        customerId: id,
        reason: excelText(pick(r, ['Reason'])) || 'Imported',
        date: excelDate(pick(r, ['Date'])) || now().slice(0, 10),
        notes: excelText(pick(r, ['Notes']))
      });
      stats.blacklist++;
    }
  });

  const ex = sheets.expired_test || sheets.expired || [];
  ex.forEach((r, i) => {
    const id = excelText(pick(r, ['Customer Id', 'Customer ID']));
    if (id && customersById.has(id) && !target.expiredCustomers.some(x => String(x.customerId) === id)) {
      target.expiredCustomers.push({
        id: `EX-IMP-${i + 1}-${Date.now()}`,
        customerId: id,
        status: 'DECEASED',
        date: excelDate(pick(r, ['Date of Death', 'Date'])) || now().slice(0, 10),
        notes: excelText(pick(r, ['Notes']))
      });
      stats.expiredCustomers++;
    }
  });

  // Generate schedules for new loans and attach ONLY newly imported payments.
  // Use a loan/date index instead of scanning every schedule for every payment.
  const schedulesByLoan = new Map();
  for (const loan of newLoans) {
    if (!validDate(loan.startDate) || !Number.isInteger(loan.duration) || loan.duration < 1) continue;
    let outstanding = Number(loan.amount) || 0;
    for (let i = 1; i <= loan.duration; i++) {
      const principal = loan.emiOption === 'YES'
        ? (i === loan.duration
          ? Number(outstanding.toFixed(2))
          : Number((Number(loan.amount) / loan.duration).toFixed(2)))
        : 0;
      const interest = Number((Math.max(0, outstanding) * Number(loan.interestRate || 0) / 100).toFixed(2));
      const due = monthlyDateServer(loan.startDate, i - 1);
      const schedule = {
        id: `SCH-IMP-${loan.id}-${i}`,
        loanId: loan.id,
        customerId: loan.customerId,
        installment: i,
        dueDate: due,
        principal,
        interest,
        emi: Number((principal + interest).toFixed(2)),
        paid: 0,
        penalty: 0,
        status: 'UPCOMING'
      };
      target.schedules.push(schedule);
      if (!schedulesByLoan.has(String(loan.id))) schedulesByLoan.set(String(loan.id), []);
      schedulesByLoan.get(String(loan.id)).push(schedule);
      outstanding = Math.max(0, outstanding - principal);
      stats.schedules++;
    }
  }

  for (const p of importedPayments) {
    if (p.scheduleId) continue;
    const candidates = schedulesByLoan.get(String(p.loanId)) || [];
    let best = candidates.find(s => s.dueDate === p.date);
    if (!best && candidates.length) {
      let bestDiff = Infinity;
      for (const s of candidates) {
        const d = Math.abs(Date.parse(s.dueDate + 'T00:00:00') - Date.parse(String(p.date) + 'T00:00:00'));
        if (Number.isFinite(d) && d < bestDiff) { bestDiff = d; best = s; }
      }
    }
    if (best) {
      p.scheduleId = best.id;
      best.paid = Number(best.paid || 0) + Number(p.principal || 0) + Number(p.interest || 0);
      best.status = best.paid + 0.005 >= Number(best.emi || 0) ? 'PAID' : 'PARTIAL';
    }
  }

  const integrityErrors = integrity(normalizeData(target));
  issues.push(...integrityErrors.slice(0, 50));
  return {
    data: normalizeData(target),
    issues: [...new Set(issues)].slice(0, 100),
    stats,
    sheets: wb.SheetNames
  };
}

function monthlyDateServer(start, n) {
  const d = new Date(String(start) + 'T00:00:00');
  if (Number.isNaN(d.getTime())) return '';
  const total = d.getFullYear() * 12 + d.getMonth() + Number(n || 0);
  const y = Math.floor(total / 12);
  const m = ((total % 12) + 12) % 12;
  const last = new Date(y, m + 1, 0).getDate();
  return `${y}-${String(m + 1).padStart(2, '0')}-${String(Math.min(d.getDate(), last)).padStart(2, '0')}`;
}


function isoDateFromQuery(value) {
  const v = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : new Date().toISOString().slice(0, 10);
}

function dashboardCustomerName(c) {
  return [c?.firstName, c?.middleName, c?.lastName].filter(Boolean).join(' ') || String(c?.name || c?.id || '');
}

function buildDashboardData(data, asOf, range = '6m') {
  const today = isoDateFromQuery(asOf);
  const customers = Array.isArray(data.customers) ? data.customers : [];
  const loans = Array.isArray(data.loans) ? data.loans : [];
  const schedules = Array.isArray(data.schedules) ? data.schedules : [];
  const payments = Array.isArray(data.payments) ? data.payments : [];
  const expiredIds = new Set((data.expiredCustomers || []).map(x => String(x.customerId)));
  const customerById = new Map(customers.map(c => [String(c.id), c]));
  const loanById = new Map(loans.map(l => [String(l.id), l]));
  const paymentsByLoan = new Map();
  const paymentsBySchedule = new Map();
  const paymentsByLoanDate = new Map();
  const paidPrincipalByLoan = new Map();

  for (const p of payments) {
    const loanId = String(p.loanId || '');
    if (!paymentsByLoan.has(loanId)) paymentsByLoan.set(loanId, []);
    paymentsByLoan.get(loanId).push(p);
    const sid = String(p.scheduleId || '').trim();
    if (sid) {
      if (!paymentsBySchedule.has(sid)) paymentsBySchedule.set(sid, []);
      paymentsBySchedule.get(sid).push(p);
    }
    const dateKey = loanId + '|' + String(p.date || '');
    if (!paymentsByLoanDate.has(dateKey)) paymentsByLoanDate.set(dateKey, []);
    paymentsByLoanDate.get(dateKey).push(p);
    paidPrincipalByLoan.set(loanId, (paidPrincipalByLoan.get(loanId) || 0) + Number(p.principal || 0));
  }

  const isExpired = id => expiredIds.has(String(id));
  const activeCustomer = id => !isExpired(id);
  const loanOutstanding = l => Math.max(0, Number(l?.amount || 0) - (paidPrincipalByLoan.get(String(l?.id)) || 0));
  const schedulePayments = s => {
    const sid = String(s?.id || '');
    const direct = sid ? (paymentsBySchedule.get(sid) || []) : [];
    if (direct.length) return direct;
    return paymentsByLoanDate.get(String(s?.loanId || '') + '|' + String(s?.dueDate || '')) || [];
  };
  const effectivePaid = s => {
    const total = schedulePayments(s).reduce((a, p) => a + Number(p.principal || 0) + Number(p.interest || 0), 0);
    return Math.max(Number(s?.paid || 0), total);
  };
  const effectiveDue = s => {
    const installment = Math.max(0, Number(s?.emi || 0) - effectivePaid(s));
    const penaltyPaid = schedulePayments(s).reduce((a, p) => a + Number(p.penalty || 0), 0);
    const unpaidPenalty = Math.max(0, Number(s?.penalty || 0) - penaltyPaid);
    return Number((installment + unpaidPenalty).toFixed(2));
  };
  const daysBetween = (a, b) => {
    const x = new Date(String(a) + 'T00:00:00');
    const y = new Date(String(b) + 'T00:00:00');
    return Number.isNaN(x.getTime()) || Number.isNaN(y.getTime()) ? 0 : Math.round((y - x) / 86400000);
  };

  const activeCustomers = customers.filter(c => activeCustomer(c.id));
  const activeLoans = loans.filter(l => activeCustomer(l.customerId));
  const completedLoans = activeLoans.filter(l => loanOutstanding(l) <= 0.005);
  const lent = activeLoans.reduce((a, l) => a + Number(l.amount || 0), 0);
  const remaining = activeLoans.reduce((a, l) => a + loanOutstanding(l), 0);
  const allTimeInterest = payments.reduce((a, p) => a + Number(p.interest || 0), 0);

  const todaysPayments = payments.filter(p => String(p.date || '') === today && loanById.has(String(p.loanId)));
  const activeSchedulesToday = schedules.filter(s => activeCustomer(s.customerId) && String(s.dueDate || '') === today);
  const todayScheduleIds = new Set(activeSchedulesToday.map(s => String(s.id)));
  const todayLoanIds = new Set(activeSchedulesToday.map(s => String(s.loanId)));
  const dueTodayPayments = todaysPayments.filter(p => p.scheduleId ? todayScheduleIds.has(String(p.scheduleId)) : todayLoanIds.has(String(p.loanId)));
  const overduePaymentsToday = todaysPayments.filter(p => !dueTodayPayments.includes(p));

  const grouped = new Map();
  for (const s of activeSchedulesToday) {
    const key = String(s.loanId);
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(s);
  }

  const dueTodayRows = [];
  for (const [loanId, ss] of grouped) {
    const loan = loanById.get(loanId);
    if (!loan) continue;
    const scheduledAmount = ss.reduce((sum, s) => sum + Number(s.emi || 0) + Number(s.penalty || 0), 0);
    const due = ss.reduce((sum, s) => sum + effectiveDue(s), 0);
    const relevantPayments = todaysPayments.filter(p => String(p.loanId) === loanId && (ss.some(x => String(x.id) === String(p.scheduleId || '')) || (!p.scheduleId && ss.some(x => String(x.dueDate) === String(p.date)))));
    const paidOnDate = relevantPayments.reduce((sum, p) => sum + Number(p.total || 0), 0);
    const fullyPaid = ss.every(s => effectiveDue(s) <= 0.005);
    const remainingDue = fullyPaid ? 0 : Math.max(0, due - paidOnDate);
    dueTodayRows.push({ loanId, loan, customer: customerById.get(String(loan.customerId)) || null, schedules: ss, due: remainingDue, grossDue: Math.max(scheduledAmount, due), paidOnDate, fullyPaid });
  }

  const expected = dueTodayRows.reduce((a, r) => a + Number(r.grossDue || 0), 0);
  const collected = todaysPayments.reduce((a, p) => a + Number(p.total || 0), 0);
  const principalToday = todaysPayments.reduce((a, p) => a + Number(p.principal || 0), 0);
  const interestToday = todaysPayments.reduce((a, p) => a + Number(p.interest || 0), 0);
  const penaltyToday = todaysPayments.reduce((a, p) => a + Number(p.penalty || 0), 0);
  const dueTodayCollected = dueTodayPayments.reduce((a, p) => a + Number(p.total || 0), 0);
  const overdueCollectedToday = overduePaymentsToday.reduce((a, p) => a + Number(p.total || 0), 0);
  const overdue = schedules.filter(s => activeCustomer(s.customerId) && String(s.dueDate || '') < today && effectiveDue(s) > 0.005);
  const overdueAmount = overdue.reduce((a, s) => a + effectiveDue(s), 0);

  const deceasedLoans = loans.filter(l => isExpired(l.customerId));
  const deceasedOutstanding = deceasedLoans.reduce((a, l) => a + loanOutstanding(l), 0);
  const deceasedOverdue = schedules.filter(s => isExpired(s.customerId) && String(s.dueDate || '') < today && effectiveDue(s) > 0.005);
  const deceasedOverdueAmount = deceasedOverdue.reduce((a, s) => a + effectiveDue(s), 0);

  const pendingIds = new Set((data.pendingQueue || []).map(String));
  const pendingOverdue = schedules.filter(s => activeCustomer(s.customerId) && pendingIds.has(String(s.id)) && s.pendingAddedAt && loanOutstanding(loanById.get(String(s.loanId))) > 0.005 && String(loanById.get(String(s.loanId))?.status || 'ACTIVE').toUpperCase() !== 'CLOSED' && String(s.dueDate || '') < today && effectiveDue(s) > 0.005);
  const topMap = new Map();
  for (const s of pendingOverdue) {
    const loan = loanById.get(String(s.loanId));
    const customer = loan ? customerById.get(String(loan.customerId)) : null;
    if (!loan || !customer) continue;
    const key = String(customer.id);
    const item = topMap.get(key) || { customer, amount: 0, days: 0, count: 0 };
    item.amount += effectiveDue(s);
    item.days = Math.max(item.days, Math.max(0, daysBetween(s.dueDate, today)));
    item.count += 1;
    topMap.set(key, item);
  }
  const topOverdue = [...topMap.values()].sort((a,b) => b.amount - a.amount).slice(0, 5);

  const risk = { '1–7 Days': 0, '8–30 Days': 0, '31–60 Days': 0, '60+ Days': 0 };
  for (const s of overdue) {
    const days = Math.max(1, daysBetween(s.dueDate, today));
    if (days <= 7) risk['1–7 Days']++;
    else if (days <= 30) risk['8–30 Days']++;
    else if (days <= 60) risk['31–60 Days']++;
    else risk['60+ Days']++;
  }

  const activity = [];
  const seen = new Set();
  const addActivity = (e, key) => {
    if (e.date !== today || seen.has(key)) return;
    seen.add(key); activity.push({ ...e, sortTime: e.createdAt || `${e.date}T00:00:00` });
  };
  for (const p of todaysPayments) {
    const l = loanById.get(String(p.loanId));
    const c = l ? customerById.get(String(l.customerId)) : null;
    addActivity({ date: p.date, createdAt: p.createdAt || p.activityCreatedAt, type: 'payment', title: 'Payment received', detail: `${dashboardCustomerName(c)} · ${l?.id || ''}`, amount: Number(p.total || 0) }, `payment:${p.id}`);
  }
  for (const l of loans) {
    const c = customerById.get(String(l.customerId)); const created = l.createdAt || l.activityCreatedAt || '';
    addActivity({ date: String(created).slice(0,10), createdAt: created, type: 'loan', title: 'New loan created', detail: `${dashboardCustomerName(c)} · ${l.id}`, amount: Number(l.amount || 0) }, `loan:${l.id}`);
  }
  for (const c of customers) {
    const created = c.createdAt || c.activityCreatedAt || '';
    addActivity({ date: String(created).slice(0,10), createdAt: created, type: 'customer', title: 'Customer registered', detail: dashboardCustomerName(c), amount: null }, `customer:${c.id}`);
  }
  for (const b of data.blacklist || []) {
    const c = customerById.get(String(b.customerId));
    addActivity({ date: b.date, createdAt: b.createdAt || b.date, type: 'risk', title: 'Customer blacklisted', detail: dashboardCustomerName(c) || String(b.customerId), amount: null }, `blacklist:${b.id || b.customerId}`);
  }
  for (const x of data.expiredCustomers || []) {
    const c = customerById.get(String(x.customerId));
    addActivity({ date: x.date, createdAt: x.createdAt || x.date, type: 'risk', title: 'Customer marked expired/deceased', detail: dashboardCustomerName(c) || String(x.customerId), amount: null }, `expired:${x.id || x.customerId}`);
  }
  activity.sort((a,b) => (Date.parse(b.sortTime || '') || 0) - (Date.parse(a.sortTime || '') || 0));

  const validPayments = payments.filter(p => loanById.has(String(p.loanId)) && /^\d{4}-\d{2}-\d{2}$/.test(String(p.date || '')));
  const trend = [];
  const now = new Date(today + 'T12:00:00');
  const monthKey = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
  const dayKey = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  const sumPayments = (from, to) => validPayments.reduce((sum,p) => { const d = new Date(String(p.date) + 'T00:00:00'); return d >= from && d <= to ? sum + Number(p.total || 0) : sum; }, 0);
  if (range === '7d') {
    for (let i=6;i>=0;i--) { const d=new Date(now); d.setDate(d.getDate()-i); const key=dayKey(d); trend.push({ key, label:d.toLocaleString('en-IN',{day:'2-digit',month:'short'}), total:sumPayments(new Date(key+'T00:00:00'),new Date(key+'T23:59:59')) }); }
  } else if (range === '30d') {
    for (let i=5;i>=0;i--) { const end=new Date(now); end.setDate(end.getDate()-i*5); const key=dayKey(end); const from=new Date(key+'T00:00:00'); const to=new Date(from); to.setDate(to.getDate()+4); trend.push({ key, label:end.toLocaleString('en-IN',{day:'2-digit',month:'short'}), total:sumPayments(from,to) }); }
  } else {
    const count = range === '1y' ? 12 : 6;
    for (let i=count-1;i>=0;i--) { const d=new Date(now.getFullYear(),now.getMonth()-i,1); const key=monthKey(d); const from=new Date(d.getFullYear(),d.getMonth(),1); const to=new Date(d.getFullYear(),d.getMonth()+1,0,23,59,59); trend.push({ key, label:d.toLocaleString('en-IN',{month:'short'}), total:sumPayments(from,to) }); }
  }

  return {
    asOf: today,
    counts: { customers: activeCustomers.length, loans: activeLoans.length, activeLoans: activeLoans.filter(l => loanOutstanding(l) > 0).length, completedLoans: completedLoans.length },
    portfolio: { lent, outstanding: remaining, interest: allTimeInterest },
    collection: { expected, collected, dueTodayCollected, interestToday, principalToday, penaltyToday, overdueCollectedToday, overdueAmount, collectionPct: expected > 0 ? (dueTodayCollected / expected) * 100 : 0 },
    dueTodayRows: dueTodayRows.map(r => ({ loanId:r.loanId, loan:r.loan, customer:r.customer, schedules:r.schedules, due:r.due, grossDue:r.grossDue, paidOnDate:r.paidOnDate, fullyPaid:r.fullyPaid })),
    topOverdue,
    risk,
    activity: activity.slice(0,8),
    specialCases: { loans: deceasedLoans.length, outstanding: deceasedOutstanding, overdue: deceasedOverdueAmount },
    trend
  };
}

async function api(req, res) {
  const parts = pathParts(req.url);
  const method = req.method || 'GET';

  if (method === 'GET' && parts[1] === 'health') {
    try {
      await db.query('SELECT 1');
      return send(res, 200, { ok: true, service: 'loan-management', database: 'PostgreSQL', authentication: 'server' });
    } catch (e) {
      return send(res, 503, { ok: false, service: 'loan-management', database: 'PostgreSQL', error: 'Database unavailable' });
    }
  }

  if (parts[0] !== 'api') return false;

  if (method === 'GET' && parts[1] === 'public' && parts[2] === 'branding') {
    const r = await db.query(`
      SELECT COALESCE(data_json->'settings', '{}'::jsonb) AS settings
      FROM user_data WHERE user_id = $1`, [SHARED_DATA_ID]);
    const settings = r.rows[0]?.settings || {};
    return send(res, 200, {
      branding: {
        appName: String(settings.appName || 'Loan Management'),
        logoData: String(settings.logoData || ''),
        logoEnabled: settings.logoEnabled !== false
      }
    });
  }

  if (method === 'GET' && parts[1] === 'dashboard') {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const d = await userData();
    const url = new URL(req.url, 'http://localhost');
    const asOf = isoDateFromQuery(url.searchParams.get('date'));
    const range = String(url.searchParams.get('range') || '6m');
    return send(res, 200, {
      dashboard: buildDashboardData(d, asOf, ['7d','30d','6m','1y'].includes(range) ? range : '6m'),
      user: u
    });
  }

  // Paginated customer read API. This is the first step toward removing
  // large client-side customer scans. The current JSONB storage model is
  // intentionally preserved for Phase 1/2; only the response is paginated.
  // Paginated payment history API. The current JSONB storage model is preserved;
  // only the requested payment rows are returned to the browser. Phase 3 can
  // move this contract to a normalized PostgreSQL payments table.
  if (method === 'GET' && parts[1] === 'payments' && parts[2]) {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const paymentId = decodeURIComponent(parts[2]);
    const d = await userData(u.userId);
    const payments = Array.isArray(d.payments) ? d.payments : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const schedules = Array.isArray(d.schedules) ? d.schedules : [];
    const payment = payments.find(p => String(p?.id) === String(paymentId));
    if (!payment) return send(res, 404, { error: 'Payment not found' });
    const loan = loans.find(l => String(l?.id) === String(payment.loanId)) || null;
    const customer = loan ? customers.find(c => String(c?.id) === String(loan.customerId)) || null : null;
    const schedule = payment.scheduleId
      ? schedules.find(s => String(s?.id) === String(payment.scheduleId)) || null
      : null;
    return send(res, 200, { payment, loan, customer, schedule, user: u });
  }

  if (method === 'GET' && parts[1] === 'payments' && !parts[2]) {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const url = new URL(req.url, 'http://localhost');
    const rawPage = Number.parseInt(url.searchParams.get('page') || '1', 10);
    const rawLimit = Number.parseInt(url.searchParams.get('limit') || '50', 10);
    const page = Number.isFinite(rawPage) ? Math.max(1, rawPage) : 1;
    const limit = Number.isFinite(rawLimit) ? Math.min(100, Math.max(1, rawLimit)) : 50;
    const search = String(url.searchParams.get('search') || '').trim().toLowerCase().slice(0, 100);
    const from = isoDateFromQuery(url.searchParams.get('from'));
    const to = isoDateFromQuery(url.searchParams.get('to'));
    const mode = String(url.searchParams.get('mode') || '').trim().toLowerCase().slice(0, 50);
    const customerId = String(url.searchParams.get('customerId') || '').trim();
    const loanId = String(url.searchParams.get('loanId') || '').trim();
    const d = await userData(u.userId);
    const payments = Array.isArray(d.payments) ? d.payments : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const loanById = new Map(loans.map(l => [String(l.id), l]));
    const customerById = new Map(customers.map(c => [String(c.id), c]));
    const rows = [];
    for (const payment of payments) {
      const lid = String(payment?.loanId || '');
      const loan = loanById.get(lid) || null;
      const customer = loan ? customerById.get(String(loan.customerId)) || null : null;
      const cid = String(loan?.customerId || '');
      if (customerId && cid !== customerId) continue;
      if (loanId && lid !== loanId) continue;
      const date = String(payment?.date || '');
      if (from && date < from) continue;
      if (to && date > to) continue;
      if (mode && String(payment?.mode || '').toLowerCase() !== mode) continue;
      const name = [customer?.firstName, customer?.middleName, customer?.lastName].filter(Boolean).join(' ') || String(customer?.name || '');
      const haystack = [payment.id, payment.date, payment.loanId, loan?.khataNo, loan?.legacyKhataNo, loan?.id, cid, name, customer?.mobile, customer?.reference].join(' ').toLowerCase();
      if (search && !haystack.includes(search)) continue;
      rows.push({ ...payment, loan, customer });
    }
    rows.sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')) || String(b.id || '').localeCompare(String(a.id || '')));
    const total = rows.length;
    const totalCollection = rows.reduce((sum, p) => sum + Number(p.total || 0), 0);
    const principal = rows.reduce((sum, p) => sum + Number(p.principal || 0), 0);
    const interest = rows.reduce((sum, p) => sum + Number(p.interest || 0), 0);
    const penalty = rows.reduce((sum, p) => sum + Number(p.penalty || 0), 0);
    const loanMap = new Map();
    for (const row of rows) if (row.loan) loanMap.set(String(row.loan.id), row.loan);
    const loanAmount = [...loanMap.values()].reduce((sum, l) => sum + Number(l.amount || 0), 0);
    const principalPaidByLoan = new Map();
    for (const p of payments) {
      const lid = String(p?.loanId || '');
      if (!lid) continue;
      principalPaidByLoan.set(lid, (principalPaidByLoan.get(lid) || 0) + Number(p?.principal || 0));
    }
    const remaining = [...loanMap.values()].reduce((sum, l) => sum + Math.max(0, Number(l.amount || 0) - (principalPaidByLoan.get(String(l.id)) || 0)), 0);
    const totalPages = Math.max(1, Math.ceil(total / limit));
    const safePage = Math.min(page, totalPages);
    const start = (safePage - 1) * limit;
    const data = rows.slice(start, start + limit);
    return send(res, 200, {
      payments: data,
      summary: { transactions: total, totalCollection, principal, interest, penalty, loanAmount, remaining, matchingLoans: loanMap.size },
      pagination: { page: safePage, limit, total, totalPages, hasNext: safePage < totalPages, hasPrevious: safePage > 1 },
      user: u
    });
  }

  // Targeted collections read APIs. These return only the collection rows needed
  // by the current page instead of forcing the browser to load the complete JSONB
  // database. Phase 3 can move the same contracts to normalized SQL tables.
  if (method === 'GET' && parts[1] === 'collections' && parts[2] === 'today') {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const url = new URL(req.url, 'http://localhost');
    const selected = isoDateFromQuery(url.searchParams.get('date'));
    const search = String(url.searchParams.get('search') || '').trim().toLowerCase().slice(0, 100);
    const d = await userData(u.userId);
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const schedules = Array.isArray(d.schedules) ? d.schedules : [];
    const payments = Array.isArray(d.payments) ? d.payments : [];
    const expired = new Set((Array.isArray(d.expiredCustomers) ? d.expiredCustomers : []).map(x => String(x?.customerId || x?.id || '')));
    const pendingIds = new Set((Array.isArray(d.pendingQueue) ? d.pendingQueue : []).map(String));
    const customerById = new Map(customers.map(c => [String(c.id), c]));
    const loanById = new Map(loans.map(l => [String(l.id), l]));
    const loanCustomer = l => l ? customerById.get(String(l.customerId)) || null : null;
    const scheduleTotals = new Map();
    for (const p of payments) {
      const sid = String(p?.scheduleId || '').trim();
      if (!sid) continue;
      const x = scheduleTotals.get(sid) || { principalInterest: 0, penalty: 0 };
      x.principalInterest += Number(p?.principal || 0) + Number(p?.interest || 0);
      x.penalty += Number(p?.penalty || 0);
      scheduleTotals.set(sid, x);
    }
    const scheduleDue = s => {
      const t = scheduleTotals.get(String(s?.id)) || { principalInterest: 0, penalty: 0 };
      const installment = Math.max(0, Number(s?.emi || 0) - Math.max(Number(s?.paid || 0), t.principalInterest));
      const unpaidPenalty = Math.max(0, Number(s?.penalty || 0) - t.penalty);
      return Number((installment + unpaidPenalty).toFixed(2));
    };
    const byLoan = new Map();
    for (const s of schedules) {
      if (String(s?.dueDate || '') !== selected || s?.manualPending) continue;
      const l = loanById.get(String(s?.loanId || ''));
      if (!l || expired.has(String(l.customerId))) continue;
      const due = scheduleDue(s);
      const sid = String(s.id);
      const t = scheduleTotals.get(sid) || { principalInterest: 0, penalty: 0 };
      const status = due <= 0.005 ? 'PAID' : String(s.dueDate) < selected ? 'OVERDUE' : String(s.dueDate) === selected ? 'DUE TODAY' : 'UPCOMING';
      const row = byLoan.get(String(l.id)) || { loan: l, customer: loanCustomer(l), schedules: [], due: 0, grossDue: 0, paidOnDate: 0, fullyPaidInstallment: true, hasPaidSchedule: false };
      row.schedules.push(s);
      row.due += due;
      row.grossDue += Number(s?.emi || 0) + Number(s?.penalty || 0);
      row.fullyPaidInstallment = row.fullyPaidInstallment && due <= 0.005;
      row.hasPaidSchedule = row.hasPaidSchedule || status === 'PAID';
      byLoan.set(String(l.id), row);
    }
    const paymentBySchedule = new Map();
    const paymentByLoan = new Map();
    const principalByLoan = new Map();
    for (const p of payments) {
      const plid = String(p?.loanId || '');
      principalByLoan.set(plid, (principalByLoan.get(plid) || 0) + Number(p?.principal || 0));
      if (String(p?.date || '') !== selected) continue;
      const lid = String(p?.loanId || '');
      const sid = String(p?.scheduleId || '').trim();
      if (sid) paymentBySchedule.set(sid, (paymentBySchedule.get(sid) || 0) + Number(p?.total || 0));
      paymentByLoan.set(lid, (paymentByLoan.get(lid) || 0) + Number(p?.total || 0));
    }
    const rows = [];
    for (const row of byLoan.values()) {
      const l = row.loan, c = row.customer;
      row.loan = { ...l, paidPrincipal: principalByLoan.get(String(l.id)) || 0 };
      row.paidOnDate = row.schedules.reduce((sum, s) => sum + (paymentBySchedule.get(String(s.id)) || 0), 0);
      const fullyPaidToday = row.fullyPaidInstallment || (row.paidOnDate > 0 && row.grossDue > 0 && row.paidOnDate >= row.grossDue - 0.005);
      const visible = !row.schedules.some(s => pendingIds.has(String(s.id))) || row.fullyPaidInstallment;
      if (!visible) continue;
      const haystack = [l?.legacyKhataNo, l?.khataNo, l?.id, c?.id, c?.mobile, c?.reference, c?.firstName, c?.middleName, c?.lastName, c?.city, c?.district].filter(Boolean).join(' ').toLowerCase();
      if (search && !haystack.includes(search)) continue;
      const unpaid = row.schedules.find(s => scheduleDue(s) > 0.005) || row.schedules[0];
      rows.push({ ...row, s: unpaid, due: row.fullyPaidInstallment ? 0 : Number(Math.max(0, row.due - row.paidOnDate).toFixed(2)), fullyPaidToday });
    }
    rows.sort((a,b) => String(a.loan?.id || '').localeCompare(String(b.loan?.id || '')));
    const allocated = payments.filter(p => String(p?.date || '') === selected && (!p.scheduleId || rows.some(r => r.schedules.some(s => String(s.id) === String(p.scheduleId)))));
    return send(res, 200, {
      date: selected,
      rows,
      summary: {
        entries: rows.length,
        expected: rows.reduce((sum,r) => sum + Number(r.grossDue || 0), 0),
        collected: allocated.reduce((sum,p) => sum + Number(p.total || 0), 0),
        interest: allocated.reduce((sum,p) => sum + Number(p.interest || 0), 0),
        pending: rows.reduce((sum,r) => sum + Number(r.due || 0), 0)
      },
      user: u
    });
  }

  if (method === 'GET' && parts[1] === 'collections' && parts[2] === 'pending') {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const url = new URL(req.url, 'http://localhost');
    const rawPage = Number.parseInt(url.searchParams.get('page') || '1', 10);
    const rawLimit = Number.parseInt(url.searchParams.get('limit') || '50', 10);
    const page = Number.isFinite(rawPage) ? Math.max(1, rawPage) : 1;
    const limit = Number.isFinite(rawLimit) ? Math.min(100, Math.max(1, rawLimit)) : 50;
    const search = String(url.searchParams.get('search') || '').trim().toLowerCase().slice(0, 100);
    const asOf = isoDateFromQuery(url.searchParams.get('date'));
    const d = await userData(u.userId);
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const schedules = Array.isArray(d.schedules) ? d.schedules : [];
    const payments = Array.isArray(d.payments) ? d.payments : [];
    const expired = new Set((Array.isArray(d.expiredCustomers) ? d.expiredCustomers : []).map(x => String(x?.customerId || x?.id || '')));
    const pendingIds = new Set((Array.isArray(d.pendingQueue) ? d.pendingQueue : []).map(String));
    const customerById = new Map(customers.map(c => [String(c.id), c]));
    const loanById = new Map(loans.map(l => [String(l.id), l]));
    const totals = new Map();
    for (const p of payments) {
      const sid = String(p?.scheduleId || '').trim();
      if (!sid) continue;
      const x = totals.get(sid) || { principalInterest: 0, penalty: 0 };
      x.principalInterest += Number(p?.principal || 0) + Number(p?.interest || 0);
      x.penalty += Number(p?.penalty || 0);
      totals.set(sid, x);
    }
    const due = s => {
      const t = totals.get(String(s?.id)) || { principalInterest: 0, penalty: 0 };
      return Number((Math.max(0, Number(s?.emi || 0) - Math.max(Number(s?.paid || 0), t.principalInterest)) + Math.max(0, Number(s?.penalty || 0) - t.penalty)).toFixed(2));
    };
    const rows = [];
    for (const s of schedules) {
      if (!pendingIds.has(String(s?.id)) || !s?.pendingAddedAt) continue;
      const l = loanById.get(String(s?.loanId || ''));
      if (!l || expired.has(String(l.customerId))) continue;
      const c = customerById.get(String(l.customerId)) || null;
      const pending = due(s);
      if (pending <= 0.005 || String(l.status || '').toUpperCase() === 'CLOSED') continue;
      const status = String(s.dueDate) < asOf ? 'OVERDUE' : String(s.dueDate) === asOf ? 'DUE TODAY' : 'UPCOMING';
      const khata = l?.legacyKhataNo || l?.khataNo || l?.id || '-';
      const haystack = [c?.firstName,c?.middleName,c?.lastName,c?.mobile,c?.id,khata,l?.id,status].filter(Boolean).join(' ').toLowerCase();
      if (search && !haystack.includes(search)) continue;
      rows.push({ schedule:s, loan:l, customer:c, pending, status, daysLate: Math.max(0, Math.floor((new Date(asOf+'T00:00:00') - new Date(String(s.dueDate)+'T00:00:00')) / 86400000)) });
    }
    rows.sort((a,b) => String(b.schedule?.dueDate || '').localeCompare(String(a.schedule?.dueDate || '')));
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / limit));
    const safePage = Math.min(page, totalPages);
    const start = (safePage - 1) * limit;
    const data = rows.slice(start, start + limit);
    const overdue = rows.filter(r => r.status === 'OVERDUE');
    const dueToday = rows.filter(r => r.status === 'DUE TODAY');
    return send(res, 200, {
      rows: data,
      summary: { totalPending: rows.reduce((sum,r) => sum + r.pending, 0), overdueLoans: new Set(overdue.map(r => r.loan.id)).size, overdueEmis: overdue.length, dueToday: dueToday.reduce((sum,r) => sum + r.pending, 0) },
      pagination: { page:safePage, limit, total, totalPages, hasNext:safePage<totalPages, hasPrevious:safePage>1 },
      user:u
    });
  }

  if (method === 'GET' && parts[1] === 'customers' && parts[2]) {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const customerId = decodeURIComponent(parts[2]);
    const x = await userData(u.userId);
    const customers = Array.isArray(x.customers) ? x.customers : [];
    const loans = Array.isArray(x.loans) ? x.loans : [];
    const payments = Array.isArray(x.payments) ? x.payments : [];
    const schedules = Array.isArray(x.schedules) ? x.schedules : [];
    const blacklistRows = Array.isArray(x.blacklist) ? x.blacklist : [];
    const expiredCustomers = Array.isArray(x.expiredCustomers) ? x.expiredCustomers : [];
    const customer = customers.find(c => String(c?.id) === String(customerId));
    if (!customer) return send(res, 404, { error: 'Customer not found' });

    const customerLoans = loans.filter(l => String(l?.customerId) === String(customerId));
    const loanIds = new Set(customerLoans.map(l => String(l?.id)));
    const customerPayments = payments.filter(p => loanIds.has(String(p?.loanId)));
    const customerSchedules = schedules.filter(s => loanIds.has(String(s?.loanId)));
    const blacklist = blacklistRows.find(b => String(b?.customerId) === String(customerId)) || null;
    const expired = expiredCustomers.find(c => String(c?.id || c?.customerId) === String(customerId)) || null;

    return send(res, 200, {
      customer,
      loans: customerLoans,
      payments: customerPayments,
      schedules: customerSchedules,
      blacklist,
      expired,
      summary: {
        loanCount: customerLoans.length,
        totalLoan: customerLoans.reduce((sum, l) => sum + Number(l?.amount || 0), 0),
        paymentCount: customerPayments.length,
        scheduleCount: customerSchedules.length
      },
      user: u
    });
  }

  // Paginated loan read API. The JSONB storage model is preserved for now;
  // only the response sent to the browser is paginated. Phase 3 can move the
  // same contract to normalized PostgreSQL tables for true SQL pagination.
  if (method === 'GET' && parts[1] === 'loans' && parts[2]) {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const loanId = decodeURIComponent(parts[2]);
    const d = await userData(u.userId);
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const payments = Array.isArray(d.payments) ? d.payments : [];
    const schedules = Array.isArray(d.schedules) ? d.schedules : [];
    const loan = loans.find(l => String(l?.id) === String(loanId));
    if (!loan) return send(res, 404, { error: 'Loan not found' });
    const customer = customers.find(c => String(c?.id) === String(loan.customerId)) || null;
    const loanPayments = payments.filter(p => String(p?.loanId) === String(loanId));
    const loanSchedules = schedules.filter(s => String(s?.loanId) === String(loanId));
    const paidPrincipal = loanPayments.reduce((sum, p) => sum + Number(p?.principal || 0), 0);
    const totalPaid = loanPayments.reduce((sum, p) => sum + Number(p?.total || 0), 0);
    const outstanding = Math.max(0, Number(loan.amount || 0) - paidPrincipal);
    const today = isoDateFromQuery(new URL(req.url, 'http://localhost').searchParams.get('date'));
    const next = loanSchedules
      .filter(s => Number(s?.emi || 0) - Number(s?.paid || 0) > 0.005 && String(s?.dueDate || '') >= today)
      .sort((a, b) => String(a?.dueDate || '').localeCompare(String(b?.dueDate || '')))[0] || null;
    const overdue = loanSchedules.some(s => String(s?.dueDate || '') < today && String(s?.status || '').toUpperCase() !== 'PAID' && (Number(s?.emi || 0) - Number(s?.paid || 0) > 0.005));
    const status = outstanding <= 0.005 ? 'CLOSED' : String(loan.status || '').toUpperCase() === 'DEAD' ? 'DEAD' : overdue ? 'OVERDUE' : 'ACTIVE';
    return send(res, 200, {
      loan,
      customer,
      payments: loanPayments,
      schedules: loanSchedules,
      summary: { paidPrincipal, totalPaid, outstanding, status, nextDue: next?.dueDate || null },
      user: u
    });
  }

  if (method === 'GET' && parts[1] === 'loans' && !parts[2]) {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const url = new URL(req.url, 'http://localhost');
    const rawPage = Number.parseInt(url.searchParams.get('page') || '1', 10);
    const rawLimit = Number.parseInt(url.searchParams.get('limit') || '50', 10);
    const page = Number.isFinite(rawPage) ? Math.max(1, rawPage) : 1;
    const limit = Number.isFinite(rawLimit) ? Math.min(100, Math.max(1, rawLimit)) : 50;
    const search = String(url.searchParams.get('search') || '').trim().toLowerCase().slice(0, 100);
    const statusFilter = String(url.searchParams.get('status') || '').trim().toUpperCase();
    const sort = String(url.searchParams.get('sort') || 'startDate');
    const order = String(url.searchParams.get('order') || 'desc').toLowerCase() === 'asc' ? 'asc' : 'desc';
    const d = await userData(u.userId);
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const payments = Array.isArray(d.payments) ? d.payments : [];
    const schedules = Array.isArray(d.schedules) ? d.schedules : [];
    const expiredIds = new Set((d.expiredCustomers || []).map(x => String(x.customerId || x.id || '')));
    const customerById = new Map(customers.map(c => [String(c.id), c]));
    const paidPrincipalByLoan = new Map();
    for (const payment of payments) {
      const lid = String(payment?.loanId || '');
      paidPrincipalByLoan.set(lid, (paidPrincipalByLoan.get(lid) || 0) + Number(payment?.principal || 0));
    }
    const schedulesByLoan = new Map();
    for (const schedule of schedules) {
      const lid = String(schedule?.loanId || '');
      if (!schedulesByLoan.has(lid)) schedulesByLoan.set(lid, []);
      schedulesByLoan.get(lid).push(schedule);
    }
    const today = isoDateFromQuery(url.searchParams.get('date'));
    const rows = [];
    for (const loan of loans) {
      const customerId = String(loan?.customerId || '');
      if (!customerId || expiredIds.has(customerId)) continue;
      const customer = customerById.get(customerId) || {};
      const paid = paidPrincipalByLoan.get(String(loan.id)) || 0;
      const remaining = Math.max(0, Number(loan.amount || 0) - paid);
      const loanSchedules = schedulesByLoan.get(String(loan.id)) || [];
      const overdue = remaining > 0.005 && loanSchedules.some(s => String(s?.dueDate || '') < today && String(s?.status || '').toUpperCase() !== 'PAID' && (Number(s?.emi || 0) - Number(s?.paid || 0) > 0.005));
      const status = remaining <= 0.005 ? 'COMPLETED' : String(loan.status || '').toUpperCase() === 'DEAD' ? 'DEAD' : overdue ? 'OVERDUE' : 'ACTIVE';
      const name = [customer.firstName, customer.middleName, customer.lastName].filter(Boolean).join(' ');
      const haystack = [loan.id, loan.khataNo, loan.legacyKhataNo, name, customer.mobile, customer.id, customer.city, customer.district, loan.loanType, loan.loanAgainst, status].join(' ').toLowerCase();
      if (search && !haystack.includes(search)) continue;
      if (statusFilter && status !== statusFilter) continue;
      rows.push({ loan, customer, remaining, status });
    }
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / limit));
    const safePage = Math.min(page, totalPages);
    const valueForSort = row => {
      if (sort === 'amount') return Number(row.loan.amount || 0);
      if (sort === 'remaining') return Number(row.remaining || 0);
      if (sort === 'customer') return [row.customer.firstName, row.customer.middleName, row.customer.lastName].filter(Boolean).join(' ').toLowerCase();
      if (sort === 'status') return row.status;
      return String(row.loan.startDate || '');
    };
    rows.sort((a, b) => {
      const av = valueForSort(a), bv = valueForSort(b);
      const cmp = typeof av === 'number' && typeof bv === 'number' ? av - bv : String(av).localeCompare(String(bv));
      return order === 'asc' ? cmp : -cmp;
    });
    const start = (safePage - 1) * limit;
    const data = rows.slice(start, start + limit).map(r => ({
      ...r.loan,
      customer: r.customer,
      remaining: r.remaining,
      computedStatus: r.status
    }));
    return send(res, 200, {
      loans: data,
      pagination: { page: safePage, limit, total, totalPages, hasNext: safePage < totalPages, hasPrevious: safePage > 1 },
      user: u
    });
  }


  // Paginated overdue-installment list. This is separate from the explicit
  // Pending queue because Overdue Loans historically includes every overdue
  // unpaid schedule, whether or not it was manually added to Pending.
  if (method === 'GET' && parts[1] === 'overdue' && !parts[2]) {
    const u=await sessionUser(req); if(!u)return send(res,401,{error:'Authentication required'});
    const url=new URL(req.url,'http://localhost');
    const page=Math.max(1,Number.parseInt(url.searchParams.get('page')||'1',10)||1);
    const limit=Math.min(100,Math.max(1,Number.parseInt(url.searchParams.get('limit')||'50',10)||50));
    const search=String(url.searchParams.get('search')||'').trim().toLowerCase().slice(0,100);
    const date=isoDateFromQuery(url.searchParams.get('date'));
    const d=await userData(u.userId),customers=Array.isArray(d.customers)?d.customers:[],loans=Array.isArray(d.loans)?d.loans:[],schedules=Array.isArray(d.schedules)?d.schedules:[],payments=Array.isArray(d.payments)?d.payments:[];
    const expiredIds=new Set((d.expiredCustomers||[]).map(x=>String(x.customerId||x.id||''))),byId=new Map(customers.map(c=>[String(c.id),c])),loanById=new Map(loans.map(l=>[String(l.id),l])),paidBySchedule=new Map(),paidByLoan=new Map();
    for(const pay of payments){const sid=String(pay.scheduleId||'');const lid=String(pay.loanId||'');const total=Number(pay.total||0);if(sid)paidBySchedule.set(sid,(paidBySchedule.get(sid)||0)+total);if(lid)paidByLoan.set(lid,(paidByLoan.get(lid)||0)+Number(pay.principal||0));}
    const rows=[];
    for(const sch of schedules){const loan=loanById.get(String(sch.loanId));if(!loan||expiredIds.has(String(loan.customerId)))continue;const customer=byId.get(String(loan.customerId))||{};const scheduled=Math.max(0,Number(sch.emi||0));const paid=Math.max(Number(sch.paid||0),paidBySchedule.get(String(sch.id))||0);const pending=Math.max(0,scheduled-paid);if(!String(sch.dueDate||'')||String(sch.dueDate)>=date||pending<=0.005)continue;const hay=[loan.id,loan.khataNo,loan.legacyKhataNo,loan.customerId,customer.id,customer.firstName,customer.middleName,customer.lastName,customer.name,customer.mobile].join(' ').toLowerCase();if(search&&!hay.includes(search))continue;rows.push({schedule:sch,loan,customer,pending,daysLate:Math.max(0,Math.floor((new Date(date+'T00:00:00')-new Date(String(sch.dueDate)+'T00:00:00'))/86400000)),status:'OVERDUE'});}
    rows.sort((a,b)=>String(a.schedule.dueDate).localeCompare(String(b.schedule.dueDate)));
    const total=rows.length,totalPages=Math.max(1,Math.ceil(total/limit)),safePage=Math.min(page,totalPages),start=(safePage-1)*limit;
    return send(res,200,{rows:rows.slice(start,start+limit),pagination:{page:safePage,limit,total,totalPages,hasNext:safePage<totalPages,hasPrevious:safePage>1},user:u});
  }

  // Paginated blacklist list. Uses the existing JSONB store for now; Phase 3
  // can move this contract to a normalized PostgreSQL table.
  if (method === 'GET' && parts[1] === 'blacklist' && !parts[2]) {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const url = new URL(req.url, 'http://localhost');
    const page = Math.max(1, Number.parseInt(url.searchParams.get('page') || '1', 10) || 1);
    const limit = Math.min(100, Math.max(1, Number.parseInt(url.searchParams.get('limit') || '50', 10) || 50));
    const search = String(url.searchParams.get('search') || '').trim().toLowerCase().slice(0, 100);
    const d = await userData(u.userId);
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const blacklist = Array.isArray(d.blacklist) ? d.blacklist : [];
    const customerById = new Map(customers.map(c => [String(c.id), c]));
    const loansByCustomer = new Map();
    for (const l of loans) {
      const cid = String(l.customerId || '');
      if (!loansByCustomer.has(cid)) loansByCustomer.set(cid, []);
      loansByCustomer.get(cid).push(l);
    }
    const rows = blacklist.map(b => {
      const customer = customerById.get(String(b.customerId));
      if (!customer) return null;
      const customerLoans = loansByCustomer.get(String(customer.id)) || [];
      const hay = [customer.id, customer.firstName, customer.middleName, customer.lastName, customer.name, customer.mobile, customer.city, b.reason, b.notes].join(' ').toLowerCase();
      if (search && !hay.includes(search)) return null;
      return { blacklist: b, customer, loanCount: customerLoans.length, outstanding: customerLoans.reduce((sum, l) => sum + Math.max(0, Number(l.amount || 0)), 0) };
    }).filter(Boolean);
    rows.sort((a,b) => String(a.customer.firstName || a.customer.name || a.customer.id).localeCompare(String(b.customer.firstName || b.customer.name || b.customer.id), undefined, {sensitivity:'base'}));
    const total = rows.length, totalPages = Math.max(1, Math.ceil(total / limit)), safePage = Math.min(page, totalPages), start=(safePage-1)*limit;
    return send(res, 200, { rows: rows.slice(start,start+limit), pagination:{page:safePage,limit,total,totalPages,hasNext:safePage<totalPages,hasPrevious:safePage>1}, user:u });
  }

  // Paginated expired/deceased customer list.
  if (method === 'GET' && parts[1] === 'expired-people' && !parts[2]) {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });
    const url = new URL(req.url, 'http://localhost');
    const page = Math.max(1, Number.parseInt(url.searchParams.get('page') || '1', 10) || 1);
    const limit = Math.min(100, Math.max(1, Number.parseInt(url.searchParams.get('limit') || '50', 10) || 50));
    const search = String(url.searchParams.get('search') || '').trim().toLowerCase().slice(0, 100);
    const d = await userData(u.userId);
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const expired = Array.isArray(d.expiredCustomers) ? d.expiredCustomers : [];
    const customerById = new Map(customers.map(c => [String(c.id), c]));
    const loansByCustomer = new Map();
    for (const l of loans) { const cid=String(l.customerId||''); if(!loansByCustomer.has(cid)) loansByCustomer.set(cid,[]); loansByCustomer.get(cid).push(l); }
    const rows = expired.map(ex => {
      const cid=String(ex.customerId || ex.id || ''), customer=customerById.get(cid); if(!customer)return null;
      const hay=[customer.id,customer.firstName,customer.middleName,customer.lastName,customer.name,customer.mobile,customer.city].join(' ').toLowerCase();
      if(search && !hay.includes(search))return null;
      const customerLoans=loansByCustomer.get(cid)||[];
      return { expired:ex, customer, loanCount:customerLoans.length, outstanding:customerLoans.reduce((sum,l)=>sum+Math.max(0,Number(l.amount||0)),0) };
    }).filter(Boolean);
    rows.sort((a,b)=>String(b.expired.date||'').localeCompare(String(a.expired.date||'')));
    const total=rows.length,totalPages=Math.max(1,Math.ceil(total/limit)),safePage=Math.min(page,totalPages),start=(safePage-1)*limit;
    return send(res,200,{rows:rows.slice(start,start+limit),pagination:{page:safePage,limit,total,totalPages,hasNext:safePage<totalPages,hasPrevious:safePage>1},user:u});
  }

  // Targeted schedule search. Returns only matching loans; selecting one can
  // reuse the existing /api/loans/:id detail endpoint.
  if (method === 'GET' && parts[1] === 'schedule' && parts[2] === 'search') {
    const u = await sessionUser(req);
    if (!u) return send(res,401,{error:'Authentication required'});
    const url=new URL(req.url,'http://localhost');
    const q=String(url.searchParams.get('q')||'').trim().toLowerCase().slice(0,100);
    if(!q) return send(res,200,{loans:[],user:u});
    const d=await userData(u.userId), customers=Array.isArray(d.customers)?d.customers:[], loans=Array.isArray(d.loans)?d.loans:[];
    const expiredIds=new Set((d.expiredCustomers||[]).map(x=>String(x.customerId||x.id||'')));
    const byId=new Map(customers.map(c=>[String(c.id),c]));
    const rows=loans.filter(l=>!expiredIds.has(String(l.customerId))).map(l=>{const c=byId.get(String(l.customerId))||{};const name=[c.firstName,c.middleName,c.lastName].filter(Boolean).join(' ')||c.name||'';return {loan:l,customer:c,name};}).filter(r=>[r.loan.id,r.loan.customerId,r.loan.khataNo,r.loan.legacyKhataNo,r.name,r.customer.mobile].join(' ').toLowerCase().includes(q)).slice(0,20);
    return send(res,200,{loans:rows,user:u});
  }

  // Server-side report/analytics aggregation. The browser receives only the
  // selected year's 12 monthly rows and KPI values instead of all payments.
  if (method === 'GET' && (parts[1] === 'reports' || parts[1] === 'analytics') && !parts[2]) {
    const u=await sessionUser(req); if(!u)return send(res,401,{error:'Authentication required'});
    const url=new URL(req.url,'http://localhost');
    const requested=Number.parseInt(url.searchParams.get('year')||'',10);
    const d=await userData(u.userId);
    const customers=Array.isArray(d.customers)?d.customers:[], loans=Array.isArray(d.loans)?d.loans:[], payments=Array.isArray(d.payments)?d.payments:[];
    const validPayments=payments.filter(p=>/^\d{4}-\d{2}-\d{2}$/.test(String(p.date||'')) && loans.some(l=>String(l.id)===String(p.loanId)));
    const years=new Set();
    const addYear=v=>{const m=String(v||'').match(/^(\d{4})-/);if(m)years.add(Number(m[1]));};
    customers.forEach(c=>addYear(c.createdAt||c.activityCreatedAt));
    loans.forEach(l=>addYear(l.startDate||l.loanDate));
    validPayments.forEach(p=>addYear(p.date));
    const yearsSorted=[...years].sort((a,b)=>b-a); if(!yearsSorted.length)yearsSorted.push(new Date().getFullYear());
    const year=yearsSorted.includes(requested)?requested:yearsSorted[0];
    const monthRows=Array.from({length:12},(_,i)=>{const key=`${year}-${String(i+1).padStart(2,'0')}`;const ls=loans.filter(l=>String(l.startDate||l.loanDate||'').startsWith(key));const ps=validPayments.filter(p=>String(p.date).startsWith(key));return {month:i+1,label:new Date(year,i,1).toLocaleString('en-IN',{month:'short'}),investment:ls.reduce((a,l)=>a+Number(l.amount||0),0),principal:ps.reduce((a,p)=>a+Number(p.principal||0),0),interest:ps.reduce((a,p)=>a+Number(p.interest||0),0),penalty:ps.reduce((a,p)=>a+Number(p.penalty||0),0),total:ps.reduce((a,p)=>a+Number(p.total||0),0),paymentCount:ps.length};});
    const yearPayments=validPayments.filter(p=>String(p.date).startsWith(String(year)));
    const paidPrincipalByLoan=new Map(),paymentsByLoan=new Map(); for(const pay of validPayments){const lid=String(pay.loanId);paidPrincipalByLoan.set(lid,(paidPrincipalByLoan.get(lid)||0)+Number(pay.principal||0));if(!paymentsByLoan.has(lid))paymentsByLoan.set(lid,[]);paymentsByLoan.get(lid).push(pay);}
    for(const list of paymentsByLoan.values()) list.sort((a,b)=>String(a.date).localeCompare(String(b.date)));
    const completedLoans=loans.filter(l=>{const target=Math.max(0,Number(l.amount||0));if(target<=0)return false;let sum=0,last=null;for(const pay of (paymentsByLoan.get(String(l.id))||[])){sum+=Number(pay.principal||0);if(sum>=target-0.005){last=pay.date;break;}}const completion=l.completedAt||last;return completion&&String(completion).startsWith(String(year));}).length;
    const newCustomers=customers.filter(c=>String(c.createdAt||c.activityCreatedAt||'').startsWith(String(year))).length;
    const newLoans=loans.filter(l=>String(l.startDate||l.loanDate||'').startsWith(String(year))).length;
    const summary={totalPrincipal:yearPayments.reduce((a,p)=>a+Number(p.principal||0),0),totalInterest:yearPayments.reduce((a,p)=>a+Number(p.interest||0),0),totalPenalty:yearPayments.reduce((a,p)=>a+Number(p.penalty||0),0),totalCollection:yearPayments.reduce((a,p)=>a+Number(p.total||0),0),paymentCount:yearPayments.length,investment:monthRows.reduce((a,m)=>a+m.investment,0)};
    if(parts[1]==='reports') return send(res,200,{year,years:yearsSorted,rows:monthRows,summary,user:u});
    return send(res,200,{year,years:yearsSorted,summary:{investment:summary.investment,interest:summary.totalInterest,penalty:summary.totalPenalty},months:monthRows.map(m=>({month:m.month,label:m.label,investment:m.investment,interest:m.interest,penalty:m.penalty})),activity:{newCustomers,newLoans,completedLoans},user:u});
  }

  // Global search returns a small result set rather than filtering the full
  // database in the browser.
  if (method === 'GET' && parts[1] === 'search' && !parts[2]) {
    const u=await sessionUser(req); if(!u)return send(res,401,{error:'Authentication required'});
    const url=new URL(req.url,'http://localhost'); const q=String(url.searchParams.get('q')||'').trim().toLowerCase().slice(0,100);
    if(!q)return send(res,200,{customers:[],loans:[],user:u});
    const d=await userData(u.userId),customers=Array.isArray(d.customers)?d.customers:[],loans=Array.isArray(d.loans)?d.loans:[];
    const byId=new Map(customers.map(c=>[String(c.id),c]));
    const active=customers.filter(c=>!(d.expiredCustomers||[]).some(x=>String(x.customerId||x.id)===String(c.id)));
    const cs=active.filter(c=>[c.id,c.firstName,c.middleName,c.lastName,c.name,c.mobile,c.city].join(' ').toLowerCase().includes(q)).slice(0,8);
    const ls=loans.filter(l=>{const c=byId.get(String(l.customerId))||{};const name=[c.firstName,c.middleName,c.lastName].filter(Boolean).join(' ')||c.name||'';return [l.id,l.khataNo,l.legacyKhataNo,l.customerId,name,c.mobile].join(' ').toLowerCase().includes(q)}).slice(0,8).map(l=>({loan:l,customer:byId.get(String(l.customerId))||null}));
    return send(res,200,{customers:cs,loans:ls,user:u});
  }

  if (method === 'GET' && parts[1] === 'customers' && !parts[2]) {
    const u = await sessionUser(req);
    if (!u) return send(res, 401, { error: 'Authentication required' });

    const url = new URL(req.url, 'http://localhost');
    const rawPage = Number.parseInt(url.searchParams.get('page') || '1', 10);
    const rawLimit = Number.parseInt(url.searchParams.get('limit') || '50', 10);
    const page = Number.isFinite(rawPage) ? Math.max(1, rawPage) : 1;
    const limit = Number.isFinite(rawLimit) ? Math.min(100, Math.max(1, rawLimit)) : 50;
    const search = String(url.searchParams.get('search') || '').trim().toLowerCase().slice(0, 100);
    const sort = String(url.searchParams.get('sort') || 'name');
    const order = String(url.searchParams.get('order') || 'asc').toLowerCase() === 'desc' ? 'desc' : 'asc';

    const allowedSorts = new Set(['id', 'name', 'mobile', 'city', 'district', 'createdAt', 'loanCount', 'totalLoan', 'remaining', 'status']);
    const sortKey = allowedSorts.has(sort) ? sort : 'name';
    const d = await userData();
    const customers = Array.isArray(d.customers) ? d.customers : [];
    const loans = Array.isArray(d.loans) ? d.loans : [];
    const payments = Array.isArray(d.payments) ? d.payments : [];
    const expiredIds = new Set((d.expiredCustomers || []).map(x => String(x.customerId)));
    const blacklistedIds = new Set((d.blacklist || []).map(x => String(x.customerId)));

    // Build loan and principal summaries in one pass each, so each customer
    // is not repeatedly scanning all loans/payments.
    const loanCountByCustomer = new Map();
    const totalLoanByCustomer = new Map();
    const loanIdsByCustomer = new Map();
    const loanCustomerById = new Map();
    for (const loan of loans) {
      const cid = String(loan.customerId || '');
      const lid = String(loan.id || '');
      if (!cid) continue;
      loanCountByCustomer.set(cid, (loanCountByCustomer.get(cid) || 0) + 1);
      totalLoanByCustomer.set(cid, (totalLoanByCustomer.get(cid) || 0) + Number(loan.amount || 0));
      if (!loanIdsByCustomer.has(cid)) loanIdsByCustomer.set(cid, new Set());
      if (lid) loanIdsByCustomer.get(cid).add(lid);
      if (lid) loanCustomerById.set(lid, cid);
    }

    const paidPrincipalByLoan = new Map();
    for (const payment of payments) {
      const lid = String(payment.loanId || '');
      if (!lid) continue;
      paidPrincipalByLoan.set(lid, (paidPrincipalByLoan.get(lid) || 0) + Number(payment.principal || 0));
    }

    const remainingByCustomer = new Map();
    for (const loan of loans) {
      const cid = String(loan.customerId || '');
      if (!cid) continue;
      const remaining = Math.max(0, Number(loan.amount || 0) - (paidPrincipalByLoan.get(String(loan.id || '')) || 0));
      remainingByCustomer.set(cid, (remainingByCustomer.get(cid) || 0) + remaining);
    }

    const customerRows = [];
    for (const customer of customers) {
      const cid = String(customer.id || '');
      if (!cid || expiredIds.has(cid)) continue;

      const name = [customer.firstName, customer.middleName, customer.lastName].filter(Boolean).join(' ') || String(customer.name || cid);
      const city = String(customer.city || customer.village || '');
      const mobile = String(customer.mobile || '');
      const district = String(customer.district || '');
      const reference = String(customer.reference || customer.customerReference || '');
      const loanCount = loanCountByCustomer.get(cid) || 0;
      const totalLoan = totalLoanByCustomer.get(cid) || 0;
      const remaining = remainingByCustomer.get(cid) || 0;
      const status = blacklistedIds.has(cid) ? 'BLACKLISTED' : (loanCount > 0 && remaining <= 0.005 ? 'COMPLETED' : 'ACTIVE');
      const haystack = `${cid} ${name} ${mobile} ${city} ${district} ${reference} ${customer.address || ''} ${status}`.toLowerCase();

      if (search && !haystack.includes(search)) continue;

      customerRows.push({
        ...customer,
        city,
        loanCount,
        totalLoan,
        remaining,
        status
      });
    }

    const compare = (a, b) => {
      let av = a[sortKey];
      let bv = b[sortKey];
      if (sortKey === 'name') { av = [a.firstName, a.middleName, a.lastName].filter(Boolean).join(' ') || a.name || a.id; bv = [b.firstName, b.middleName, b.lastName].filter(Boolean).join(' ') || b.name || b.id; }
      if (typeof av === 'number' && typeof bv === 'number') return av - bv;
      return String(av ?? '').localeCompare(String(bv ?? ''), undefined, { numeric: true, sensitivity: 'base' });
    };
    customerRows.sort((a, b) => {
      const result = compare(a, b);
      return order === 'desc' ? -result : result;
    });

    const total = customerRows.length;
    const totalPages = Math.max(1, Math.ceil(total / limit));
    const safePage = Math.min(page, totalPages);
    const start = (safePage - 1) * limit;
    const data = customerRows.slice(start, start + limit);

    return send(res, 200, {
      customers: data,
      pagination: { page: safePage, limit, total, totalPages, hasNext: safePage < totalPages, hasPrevious: safePage > 1 },
      user: u
    });
  }

  if (parts[1] === 'auth') {
    if (method === 'POST' && parts[2] === 'register') {
      const requester = await sessionUser(req);
      const existing = await countUsers();
      if (existing > 0 && (!requester || requester.role !== 'Administrator')) {
        return send(res, 403, { error: 'Only the first account can self-register. An Administrator must create additional users.' });
      }

      const b = await readBody(req);
      const name = String(b.name || '').trim().slice(0, 100);
      const username = String(b.username || '').trim().toLowerCase();
      const mobile = String(b.mobile || '').trim();
      const role = String(b.role || '').trim();
      const password = String(b.password || '');

      if (existing === 0 && role !== 'Administrator') {
        return send(res, 400, { error: 'The first account must be Administrator' });
      }

      if (!name || !/^[a-z0-9._-]{3,30}$/.test(username) || (mobile && !mobileOk(mobile)) ||
        !ROLES.includes(role) || password.length < 8 || !/[A-Za-z]/.test(password) || !/[0-9]/.test(password)) {
        return send(res, 400, { error: 'Invalid registration details' });
      }

      if (await userByUsername(username)) return send(res, 409, { error: 'Username already exists' });

      if (mobile) {
        const mobileResult = await db.query('SELECT id FROM users WHERE mobile = $1', [mobile]);
        if (mobileResult.rows.length) return send(res, 409, { error: 'Mobile number is already registered' });
      }

      const id = uid('USR');
      await db.query(
        `INSERT INTO users(id, name, username, mobile, role, password_hash, active, created_at)
         VALUES ($1, $2, $3, $4, $5, $6, TRUE, $7)`,
        [id, name, username, mobile || null, role, hashPassword(password), now()]
      );

      await db.query(
        `INSERT INTO user_data(user_id, data_json, updated_at)
         VALUES ($1, $2::jsonb, $3)
         ON CONFLICT (user_id) DO NOTHING`,
        [id, JSON.stringify(blankData()), now()]
      );

      return send(res, 201, { user: { id, name, username, mobile, role } });
    }

    if (method === 'POST' && parts[2] === 'login') {
      if (rateLimited(req)) {
        return send(res, 429, { error: 'Too many login attempts. Please try again later.' }, { 'Retry-After': '900' });
      }

      const b = await readBody(req);
      const username = String(b.username || '').trim().toLowerCase();
      const password = String(b.password || '');
      const u = await userByUsername(username);

      if (!u || !verifyPassword(password, u.password_hash)) {
        return send(res, 401, { error: 'Invalid username or password' });
      }

      await setSession(res, u.id, !!b.remember);
      return send(res, 200, { user: { id: u.id, name: u.name, username: u.username, mobile: u.mobile, role: u.role } });
    }

    if (method === 'GET' && parts[2] === 'me') {
      const u = await sessionUser(req);
      if (!u) return send(res, 401, { error: 'Authentication required' });
      return send(res, 200, { user: u });
    }

    if (method === 'POST' && parts[2] === 'logout') {
      await clearSession(req, res);
      return send(res, 200, { ok: true });
    }

    if (method === 'POST' && parts[2] === 'reset-password') {
      if (rateLimited(req)) {
        return send(res, 429, { error: 'Too many attempts. Please try again later.' }, { 'Retry-After': '900' });
      }

      const b = await readBody(req);
      const u = await userByUsername(String(b.username || '').trim().toLowerCase());
      const p = String(b.password || '');
      const m = String(b.mobile || '').trim();

      if (!u || u.mobile !== m || p.length < 8 || !/[A-Za-z]/.test(p) || !/[0-9]/.test(p)) {
        return send(res, 400, { error: 'Username, mobile or password is invalid' });
      }

      await db.query('UPDATE users SET password_hash = $1 WHERE id = $2', [hashPassword(p), u.id]);
      return send(res, 200, { ok: true });
    }
  }

  const u = await sessionUser(req);
  if (!u) return send(res, 401, { error: 'Authentication required' });

  // Administrator-only business data reset. Keeps administrator accounts and application settings.
  if (method === 'POST' && parts[1] === 'admin' && parts[2] === 'clear-data' && !parts[3]) {
    if (!ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Administrator permission required' });

    const client = await db.connect();
    try {
      await client.query('BEGIN');
      const result = await client.query(
        'SELECT data_json FROM user_data WHERE user_id = $1 FOR UPDATE',
        [SHARED_DATA_ID]
      );
      const before = result.rows[0]
        ? normalizeData(typeof result.rows[0].data_json === 'string' ? JSON.parse(result.rows[0].data_json) : result.rows[0].data_json)
        : blankData();

      const counts = {
        customers: before.customers.length,
        loans: before.loans.length,
        schedules: before.schedules.length,
        payments: before.payments.length,
        blacklist: before.blacklist.length,
        notifications: before.notifications.length,
        deletedRecords: before.deletedRecords.length,
        expiredCustomers: before.expiredCustomers.length,
        pendingQueue: before.pendingQueue.length
      };

      const cleared = blankData();
      // Preserve application configuration/logo while clearing business records.
      cleared.settings = { ...before.settings };

      await client.query(
        `INSERT INTO user_data(user_id, data_json, updated_at)
         VALUES ($1, $2::jsonb, $3)
         ON CONFLICT (user_id) DO UPDATE
         SET data_json = EXCLUDED.data_json, updated_at = EXCLUDED.updated_at`,
        [SHARED_DATA_ID, JSON.stringify(cleared), now()]
      );

      await client.query('COMMIT');
      return send(res, 200, { ok: true, cleared: counts, preserved: ['settings', 'administrator accounts'] });
    } catch (e) {
      try { await client.query('ROLLBACK'); } catch {}
      console.error('Clear data failed:', e);
      return send(res, 500, { error: e.message || 'Could not clear application data' });
    } finally {
      client.release();
    }
  }

  if (method === 'POST' && parts[1] === 'import-excel') {
    if (!ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Administrator permission required' });

    let b;
    try {
      b = await readBody(req);
    } catch (e) {
      return send(res, 400, { error: e.message || 'Invalid import request' });
    }

    const mode = b.mode === 'replace' ? 'replace' : 'add';
    if (!b.preview && excelImportInProgress) {
      return send(res, 409, { error: 'Another Excel import is already in progress. Please wait for it to finish before starting another import.' });
    }

    if (b.preview) {
      try {
        const result = await importExcelPayload(b.fileBase64, mode);
        return send(res, 200, { preview: true, issues: result.issues, stats: result.stats, sheets: result.sheets });
      } catch (e) {
        return send(res, 400, { error: e.message || 'Excel preview failed' });
      }
    }

    excelImportInProgress = true;
    try {
      const result = await importExcelPayload(b.fileBase64, mode);

      if (result.issues.length) {
        return send(res, 400, {
          error: 'Import validation failed. Preview and fix the listed issues before importing.',
          details: result.issues.slice(0, 20),
          stats: result.stats
        });
      }

      if (mode === 'replace') await createBackup();
      await saveData(result.data);
      return send(res, 200, { ok: true, stats: result.stats });
    } catch (e) {
      console.error('Excel import failed:', e.message);
      return send(res, 400, { error: e.message || 'Excel import failed' });
    } finally {
      excelImportInProgress = false;
    }
  }

  // Targeted customer create API. The browser sends only the new customer
  // instead of downloading and PUTing the entire shared JSON document.
  if (method === 'POST' && parts[1] === 'customers' && !parts[2]) {
    if (!WRITE_ROLES.has(u.role)) return send(res, 403, { error: 'You do not have write permission' });

    let b;
    try {
      b = await readBody(req);
    } catch (e) {
      return send(res, 400, { error: e.message || 'Invalid request body' });
    }

    const incoming = b?.customer && typeof b.customer === 'object' ? { ...b.customer } : null;
    if (!incoming) return send(res, 400, { error: 'Customer data is required' });

    const firstName = String(incoming.firstName || '').trim();
    const mobile = String(incoming.mobile || '').trim();
    const city = String(incoming.city || incoming.village || '').trim();
    const district = String(incoming.district || '').trim();
    const guarantorMobile = String(incoming.guarantorMobile || '').trim();
    const pincode = String(incoming.pincode || '').trim();

    if (!firstName) return send(res, 400, { error: 'Customer name is required' });
    if (!mobileOk(mobile, true)) return send(res, 400, { error: 'Enter a valid 10-digit customer mobile number' });
    if (!city) return send(res, 400, { error: 'City is required' });
    if (!district) return send(res, 400, { error: 'District is required' });
    if (pincode && !/^\d{6}$/.test(pincode)) return send(res, 400, { error: 'Pincode must be 6 digits' });
    if (guarantorMobile && !mobileOk(guarantorMobile, false)) return send(res, 400, { error: 'Enter a valid guarantor mobile number' });

    const d = await userData();
    const customers = Array.isArray(d.customers) ? d.customers : [];
    if (customers.some(c => String(c?.mobile || '').trim() === mobile)) {
      return send(res, 409, { error: 'A customer with this mobile number already exists' });
    }

    let maxId = 0;
    for (const c of customers) {
      const m = String(c?.id || '').match(/^KK-(\d+)$/i);
      if (m) maxId = Math.max(maxId, Number(m[1]) || 0);
    }
    const id = `KK-${String(maxId + 1).padStart(6, '0')}`;
    const createdAt = now();
    const customer = {
      ...incoming,
      id,
      state: 'Maharashtra',
      taluka: String(incoming.taluka || ''),
      ownerId: u.id,
      createdAt,
      activityCreatedAt: createdAt
    };

    d.customers.push(customer);
    try {
      await saveData(d);
    } catch (e) {
      console.error('Customer create failed:', e);
      return send(res, 500, { error: e.message || 'Could not save customer' });
    }

    return send(res, 201, { ok: true, customer, user: u });
  }

  // Targeted mutation API. The browser sends only changed records instead of
  // PUTing the entire shared JSONB document. Multiple related changes are applied
  // atomically in one request, while the existing JSONB storage remains intact.
  if (method === 'POST' && parts[1] === 'mutations' && !parts[2]) {
    const b = await readBody(req);
    const operations = Array.isArray(b?.operations) ? b.operations.slice(0, 5000) : [];
    if (!operations.length) return send(res, 400, { error: 'At least one mutation is required' });

    const allowedTypes = new Set(['customers','loans','schedules','payments','blacklist','notifications','deletedRecords','expiredCustomers','pendingQueue','settings']);
    for (const op of operations) {
      if (!op || !allowedTypes.has(String(op.type))) return send(res, 400, { error: 'Invalid mutation type' });
      if (!['create','update','delete','replace'].includes(String(op.action))) return send(res, 400, { error: 'Invalid mutation action' });
    }

    const hasDelete = operations.some(op => op.action === 'delete');
    const hasPayment = operations.some(op => op.type === 'payments');
    const hasWrite = operations.some(op => ['customers','loans','schedules','blacklist','expiredCustomers','pendingQueue'].includes(op.type));
    const hasDeletedHistory = operations.some(op => op.type === 'deletedRecords');
    const hasSettings = operations.some(op => op.type === 'settings');
    if ((hasDelete || hasDeletedHistory) && !ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Only Administrators can delete records or modify audit history' });
    if (hasPayment && !PAYMENT_ROLES.has(u.role)) return send(res, 403, { error: 'You do not have permission to manage payments' });
    if (hasWrite && !WRITE_ROLES.has(u.role)) return send(res, 403, { error: 'You do not have write permission' });
    if (hasSettings && !ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Only Administrators can change settings' });

    const client = await db.connect();
    try {
      await client.query('BEGIN');
      const result = await client.query('SELECT data_json FROM user_data WHERE user_id = $1 FOR UPDATE', [SHARED_DATA_ID]);
      let data;
      if (result.rows[0]) {
        const value = typeof result.rows[0].data_json === 'string' ? JSON.parse(result.rows[0].data_json) : result.rows[0].data_json;
        data = normalizeData(value);
      } else {
        data = blankData();
      }

      const arrays = new Set(['customers','loans','schedules','payments','blacklist','notifications','deletedRecords','expiredCustomers']);
      for (const op of operations) {
        if (op.type === 'settings' && op.action === 'replace') {
          data.settings = { ...blankData().settings, ...(op.record || {}) };
          continue;
        }
        if (op.type === 'pendingQueue' && op.action === 'replace') {
          data.pendingQueue = Array.isArray(op.records) ? op.records.map(String) : [];
          continue;
        }
        if (!arrays.has(op.type)) continue;
        const list = data[op.type];
        const id = String(op.id ?? op.record?.id ?? '');
        const index = list.findIndex(x => String(x?.id) === id);
        if (op.action === 'create') {
          if (index >= 0) { await client.query('ROLLBACK'); return send(res, 409, { error: `${op.type} record already exists: ${id}` }); }
          if (!op.record || !id) { await client.query('ROLLBACK'); return send(res, 400, { error: `Invalid ${op.type} create operation` }); }
          list.push(op.record);
        } else if (op.action === 'update') {
          if (index < 0) { await client.query('ROLLBACK'); return send(res, 404, { error: `${op.type} record not found: ${id}` }); }
          if (!op.record) { await client.query('ROLLBACK'); return send(res, 400, { error: `Invalid ${op.type} update operation` }); }
          list[index] = op.record;
        } else if (op.action === 'delete') {
          if (index >= 0) list.splice(index, 1);
        }
      }

      const errors = integrity(data);
      if (errors.length) {
        await client.query('ROLLBACK');
        return send(res, 400, { error: 'Data validation failed', details: errors.slice(0, 20) });
      }

      const normalized = normalizeData(data);
      await client.query(
        `INSERT INTO user_data(user_id, data_json, updated_at)
         VALUES ($1, $2::jsonb, $3)
         ON CONFLICT (user_id) DO UPDATE SET data_json = EXCLUDED.data_json, updated_at = EXCLUDED.updated_at`,
        [SHARED_DATA_ID, JSON.stringify(normalized), now()]
      );
      await client.query('COMMIT');
      return send(res, 200, { ok: true, applied: operations.length, user: u });
    } catch (e) {
      try { await client.query('ROLLBACK'); } catch {}
      console.error('Mutation API failed:', e);
      return send(res, 500, { error: e.message || 'Could not apply changes' });
    } finally {
      client.release();
    }
  }

  if (method === 'GET' && parts[1] === 'db') {
    return send(res, 200, { data: await userData(), user: u });
  }

  if (method === 'PUT' && parts[1] === 'db') {
    const b = await readBody(req);
    const incoming = normalizeData(b.data);
    const errors = integrity(incoming);
    if (errors.length) return send(res, 400, { error: 'Data validation failed', details: errors.slice(0, 20) });

    const before = await userData();
    const ch = changes(before, incoming);
    const deletes = ch.filter(x => x.action === 'DELETE');

    if (deletes.length && !ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Only Administrators can delete records' });
    if (ch.some(x => x.type === 'deletedRecords') && !ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Only Administrators can modify deleted-record history' });
    if (ch.some(x => x.type === 'payments') && !PAYMENT_ROLES.has(u.role)) return send(res, 403, { error: 'You do not have permission to manage payments' });
    if (ch.some(x => ['customers', 'loans', 'schedules', 'blacklist', 'expiredCustomers', 'pendingQueue'].includes(x.type)) && !WRITE_ROLES.has(u.role)) {
      return send(res, 403, { error: 'You do not have write permission' });
    }
    if (JSON.stringify(before.settings) !== JSON.stringify(incoming.settings) && !new Set(['Administrator']).has(u.role)) {
      return send(res, 403, { error: 'Only Administrators can change settings' });
    }

    await saveData(incoming);
    return send(res, 200, { ok: true, data: incoming });
  }

  if (method === 'GET' && parts[1] === 'backup' && parts[2] === 'summary') {
    if (!ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Administrator permission required' });
    const data = await userData();
    return send(res, 200, {
      customers: Array.isArray(data.customers) ? data.customers.length : 0,
      loans: Array.isArray(data.loans) ? data.loans.length : 0,
      payments: Array.isArray(data.payments) ? data.payments.length : 0,
      expiredCustomers: Array.isArray(data.expiredCustomers) ? data.expiredCustomers.length : 0
    });
  }

  if (method === 'GET' && parts[1] === 'payments' && parts[2] === 'export') {
    const data = await userData();
    const customers = new Map((Array.isArray(data.customers) ? data.customers : []).map(c => [String(c.id), c]));
    const loans = new Map((Array.isArray(data.loans) ? data.loans : []).map(l => [String(l.id), l]));
    const escCsv = value => `"${String(value ?? '').replaceAll('"', '""')}"`;
    const header = ['Payment ID','Date','Customer','Loan ID','Principal','Interest','Penalty','Total','Mode','Notes'];
    const rows = (Array.isArray(data.payments) ? data.payments : []).map(p => {
      const loan = loans.get(String(p.loanId));
      const customer = loan ? customers.get(String(loan.customerId)) : null;
      return [p.id,p.date,customerName(customer || {}),p.loanId,p.principal,p.interest,p.penalty,p.total,p.mode,p.notes];
    });
    const csv = [header, ...rows].map(row => row.map(escCsv).join(',')).join('\n');
    res.writeHead(200, {
      'Content-Type': 'text/csv; charset=utf-8',
      'Content-Disposition': `attachment; filename="payments-${todayISO()}.csv"`,
      'Cache-Control': 'no-store'
    });
    return res.end(csv);
  }

  // Dedicated JSON backup restore endpoint. This replaces the shared JSONB document
  // directly instead of routing the backup through the browser mutation-diff path.
  // That avoids comparing/stringifying every record twice for large backups.
  if (method === 'POST' && parts[1] === 'backup' && parts[2] === 'restore' && !parts[3]) {
    if (!ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Administrator permission required' });

    let b;
    try {
      b = await readBody(req);
    } catch (e) {
      return send(res, 400, { error: e.message || 'Invalid backup request' });
    }

    const incoming = b?.sharedData && typeof b.sharedData === 'object' ? b.sharedData : b;
    if (!incoming || !Array.isArray(incoming.customers) || !Array.isArray(incoming.loans) ||
        !Array.isArray(incoming.schedules) || !Array.isArray(incoming.payments) ||
        !Array.isArray(incoming.blacklist)) {
      return send(res, 400, { error: 'Invalid or incompatible backup file' });
    }

    const normalized = normalizeData(incoming);
    const errors = integrity(normalized);
    if (errors.length) {
      return send(res, 400, {
        error: 'Backup validation failed',
        details: errors.slice(0, 20)
      });
    }

    try {
      await saveData(normalized);
      return send(res, 200, {
        ok: true,
        restored: {
          customers: normalized.customers.length,
          loans: normalized.loans.length,
          schedules: normalized.schedules.length,
          payments: normalized.payments.length,
          blacklist: normalized.blacklist.length,
          expiredCustomers: normalized.expiredCustomers.length
        },
        user: u
      });
    } catch (e) {
      console.error('Backup restore failed:', e);
      return send(res, 500, { error: e.message || 'Could not restore backup' });
    }
  }

  if (method === 'GET' && parts[1] === 'backup') {
    if (!ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Administrator permission required' });
    return send(res, 200, { version: 2, createdAt: now(), sharedData: await userData() });
  }

  return send(res, 404, { error: 'Not found' });
}

function contentType(file) {
  return {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
    '.json': 'application/json'
  }[path.extname(file).toLowerCase()] || 'application/octet-stream';
}

async function createBackup() {
  try {
    const usersResult = await db.query(
      'SELECT id, name, username, mobile, role, active, created_at FROM users ORDER BY created_at'
    );
    const payload = {
      version: 2,
      createdAt: now(),
      users: usersResult.rows,
      sharedData: await userData()
    };

    const file = `backup-${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
    const fullPath = path.join(BACKUP_DIR, file);
    fs.writeFileSync(fullPath, JSON.stringify(payload));

    const size = fs.statSync(fullPath).size;
    await db.query(
      `INSERT INTO app_backups(file_name, created_at, size_bytes)
       VALUES ($1, $2, $3)`,
      [file, payload.createdAt, size]
    );

    const files = fs.readdirSync(BACKUP_DIR).filter(f => f.endsWith('.json')).sort();
    while (files.length > 14) {
      try {
        fs.unlinkSync(path.join(BACKUP_DIR, files.shift()));
      } catch {
        // Ignore individual cleanup failures.
      }
    }

    console.log('Automatic backup:', file);
    return file;
  } catch (e) {
    console.error('Backup failed:', e.message);
    return null;
  }
}

const server = http.createServer(async (req, res) => {
  corsHeaders(req, res);

  if (req.method === 'OPTIONS') {
    res.statusCode = 204;
    res.end();
    return;
  }
  try {
    if (String(req.url).startsWith('/api/')) {
      await api(req, res);
      return;
    }

    let urlPath = new URL(req.url, 'http://localhost').pathname;
    if (urlPath === '/' || urlPath === '') urlPath = '/index.html';

    const safe = path.normalize(urlPath).replace(/^([.][.][\\/])+/, '');
    const file = path.join(ROOT, safe);

    if (!file.startsWith(ROOT)) return send(res, 403, { error: 'Forbidden' });

    fs.stat(file, (err, st) => {
      if (err || !st.isFile()) {
        securityHeaders(res);
        res.statusCode = 404;
        return res.end('Not found');
      }
      securityHeaders(res);
      res.setHeader('Content-Type', contentType(file));
      fs.createReadStream(file).pipe(res);
    });
  } catch (e) {
    console.error(e);
    if (!res.headersSent) send(res, 500, { error: 'Server error' });
  }
});

async function start() {
  try {
    await db.query('SELECT NOW()');
    console.log('PostgreSQL connection successful.');
    await initDb();

    server.listen(PORT, HOST, () => {
      console.log(`Loan Management server running at http://${HOST}:${PORT}`);
    });
  } catch (e) {
    console.error('Failed to start application:', e.message);
    await db.end().catch(() => { });
    process.exit(1);
  }
}

// Automatic backup check: once per minute, around 02:00 server local time.
setInterval(() => {
  const d = new Date();
  if (d.getHours() === 2 && d.getMinutes() < 5) {
    createBackup().catch(err => console.error('Scheduled backup error:', err.message));
  }
}, 60000);

process.on('SIGTERM', async () => {
  console.log('SIGTERM received. Closing server...');
  server.close(async () => {
    await db.end();
    process.exit(0);
  });
});

process.on('SIGINT', async () => {
  console.log('SIGINT received. Closing server...');
  server.close(async () => {
    await db.end();
    process.exit(0);
  });
});

start();
