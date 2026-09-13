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

  for (const c of data.customers) {
    const id = String(c.id);
    if (cs.has(id)) e.push('Duplicate customer ID: ' + id);
    cs.add(id);
    if (!String(c.firstName || '').trim()) e.push('Customer ' + id + ' has no first name');
    if (!mobileOk(c.mobile, true)) e.push('Customer ' + id + ' has invalid mobile');
    if (!String(c.city || '').trim()) {
      e.push('Customer ' + id + ' has no city');
    }
    if (!String(c.district || '').trim()) e.push('Customer ' + id + ' has no district');
    const m = String(c.mobile || '');
    if ([...data.customers].filter(x => String(x.mobile || '') === m).length > 1) {
      e.push('Duplicate customer mobile: ' + m);
    }
  }

  for (const l of data.loans) {
    const id = String(l.id);
    if (ls.has(id)) e.push('Duplicate loan ID: ' + id);
    ls.add(id);
    if (!cs.has(String(l.customerId))) e.push('Loan ' + id + ' references missing customer');
    if (!positive(l.amount)) e.push('Loan ' + id + ' has invalid amount');
    if (!validDate(l.startDate)) e.push('Loan ' + id + ' has invalid start date');
    const k = String(l.khataNo || l.legacyKhataNo || '').trim().toLowerCase();
    if (!k) e.push('Loan ' + id + ' missing Khata No');
    else if (kh.has(k)) e.push('Duplicate Khata No: ' + k);
    else kh.add(k);
  }

  for (const s of data.schedules) {
    if (!ls.has(String(s.loanId))) e.push('Schedule ' + s.id + ' references missing loan');
    if (!validDate(s.dueDate)) e.push('Schedule ' + s.id + ' has invalid due date');
  }

  for (const p of data.payments) {
    if (!ls.has(String(p.loanId))) e.push('Payment ' + p.id + ' references missing loan');
    if (!validDate(p.date)) e.push('Payment ' + p.id + ' has invalid date');

    const loan = data.loans.find(x => String(x.id) === String(p.loanId));
    if (loan && validDate(loan.startDate) && validDate(p.date) && p.date < loan.startDate) {
      e.push('Payment ' + p.id + ' is before loan start date');
    }

    const principal = Number(p.principal || 0);
    const interest = Number(p.interest || 0);
    const penalty = Number(p.penalty || 0);
    const t = principal + interest + penalty;

    if (![principal, interest, penalty, t].every(Number.isFinite) || principal < 0 || interest < 0 || penalty < 0) {
      e.push('Payment ' + p.id + ' has invalid amounts');
    }

    if (Math.abs(t - Number(p.total || 0)) > 0.01) e.push('Payment ' + p.id + ' total mismatch');

    if (loan) {
      const other = data.payments
        .filter(x => String(x.loanId) === String(loan.id) && String(x.id) !== String(p.id))
        .reduce((sum, x) => sum + Number(x.principal || 0), 0);
      if (other + principal > Number(loan.amount) + 0.01) {
        e.push('Payment ' + p.id + ' exceeds loan principal');
      }
    }

    if (p.scheduleId) {
      const sc = data.schedules.find(x => String(x.id) === String(p.scheduleId));
      if (!sc) e.push('Payment ' + p.id + ' references missing schedule');
      else if (String(sc.loanId) !== String(p.loanId)) e.push('Payment ' + p.id + ' schedule does not belong to loan');
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
    `kk_session=${encodeURIComponent(raw)}; HttpOnly; Path=/; SameSite=${IS_PROD ? 'None' : 'Lax'}; Max-Age=${days * 86400}${IS_PROD ? '; Secure' : ''}`
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
    `kk_session=; HttpOnly; Path=/; SameSite=Lax; Max-Age=0${IS_PROD ? '; Secure' : ''}`
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
  'http://127.0.0.1:5500'
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

  // Generate schedules for new loans and attach imported payments to matching due dates.
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
      target.schedules.push({
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
      });
      outstanding = Math.max(0, outstanding - principal);
      stats.schedules++;
    }
  }

  for (const p of target.payments) {
    if (p.scheduleId) continue;
    const candidates = target.schedules.filter(s => String(s.loanId) === String(p.loanId));
    let best = candidates.find(s => s.dueDate === p.date);
    if (!best) {
      const ts = candidates
        .map(s => ({ s, d: Math.abs(new Date(s.dueDate) - new Date(p.date)) }))
        .sort((a, b) => a.d - b.d)[0];
      best = ts?.s;
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
    const d = await userData();
    return send(res, 200, {
      branding: {
        appName: String(d.settings?.appName || 'Loan Management'),
        logoData: String(d.settings?.logoData || ''),
        logoEnabled: d.settings?.logoEnabled !== false
      }
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

  if (method === 'POST' && parts[1] === 'import-excel') {
    if (!ADMIN_ROLES.has(u.role)) return send(res, 403, { error: 'Administrator permission required' });

    const b = await readBody(req);
    try {
      const mode = b.mode === 'replace' ? 'replace' : 'add';
      const result = await importExcelPayload(b.fileBase64, mode);

      if (b.preview) {
        return send(res, 200, { preview: true, issues: result.issues, stats: result.stats, sheets: result.sheets });
      }

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
      return send(res, 400, { error: e.message || 'Excel import failed' });
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
