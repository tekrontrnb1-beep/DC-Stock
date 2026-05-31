/* ============================================================
   Tekron — Нөөцийн самбар (Stock / Inventory Dashboard)
   Excel/CSV импорт · ABC/XYZ · хэлтэс ба барааны үлдэгдэл · түүх
   Өгөгдөл localStorage-д хадгалагдана; "Нийтлэх" товчоор stock-data.json
   татаж repo-д байрлуулбал та болон захирал хоёр ижил мэдээллийг харна.
   ============================================================ */
'use strict';

/* ---------- Storage ---------- */
const STORE_KEY = 'tekron-stock-v1';
const PUBLISHED_URL = 'data/stock-data.json';

let DB = emptyDB();
function emptyDB() {
  return { products: [], balances: [], sales: [], orders: [], updatedAt: null };
}

/* ---------- UI state ---------- */
const S = {
  selDate: null,
  dept: '', cls: '', status: '', status2: '', search: '', lowOnly: false,
  sortKey: 'value', sortDir: -1,
  histScope: '__all__', histMetric: 'balance', histRange: 30,
  matrixMetric: 'value', // 'value' | 'coverSales' | 'coverOrders'
  source: 'empty', // 'local' | 'published' | 'sample' | 'empty'
};

/* ---------- Indexes (rebuilt on data change) ---------- */
let IX = null;

/* ============================================================
   Helpers
   ============================================================ */
const $ = (id) => document.getElementById(id);
const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; };
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function fmtInt(n) {
  if (n == null || isNaN(n)) return '—';
  const neg = n < 0;
  const s = Math.abs(Math.round(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, "'");
  return (neg ? '-' : '') + s;
}
function money(n) {
  if (n == null || isNaN(n)) return '—';
  return fmtInt(n) + '₮';
}
function moneyShort(n) {
  if (n == null || isNaN(n)) return '—';
  const a = Math.abs(n);
  if (a >= 1e9) return (n / 1e9).toFixed(1).replace(/\.0$/, '') + ' тэрбум₮';
  if (a >= 1e6) return (n / 1e6).toFixed(1).replace(/\.0$/, '') + ' сая₮';
  if (a >= 1e3) return (n / 1e3).toFixed(0) + 'к₮';
  return fmtInt(n) + '₮';
}

const WD = ['Ня', 'Да', 'Мя', 'Лх', 'Пү', 'Ба', 'Бя'];
function ymd(d) {
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
function fmtDate(s) {
  if (!s) return '—';
  const [y, m, d] = s.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  return `${m}/${d} (${WD[dt.getDay()]})`;
}
function fmtDateFull(s) {
  if (!s) return '—';
  const [y, m, d] = s.split('-').map(Number);
  return `${y} оны ${m}-р сарын ${d}`;
}
/* add n days to a 'YYYY-MM-DD' string */
function addDays(s, n) {
  const [y, m, d] = s.split('-').map(Number);
  const dt = new Date(y, m - 1, d); dt.setDate(dt.getDate() + n);
  return ymd(dt);
}
/* number of days in the retail (Sales.csv) period from meta.retailPeriod "MM.DD–MM.DD" */
function retailDays() {
  const p = DB.meta && DB.meta.retailPeriod;
  if (!p) return 30;
  const m = p.match(/(\d{1,2})\.(\d{1,2}).*?(\d{1,2})\.(\d{1,2})/);
  if (!m) return 30;
  const y = +(S.selDate || '2026').slice(0, 4);
  const diff = Math.round((new Date(y, +m[3] - 1, +m[4]) - new Date(y, +m[1] - 1, +m[2])) / 86400000) + 1;
  return diff > 0 ? diff : 30;
}

/* Parse many date formats / Excel serials → 'YYYY-MM-DD' */
function toYmd(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return ymd(v);
  if (typeof v === 'number') { // Excel serial day count
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d) ? null : ymd(d);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return `${m[1]}-${String(+m[2]).padStart(2, '0')}-${String(+m[3]).padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) return `${m[3]}-${String(+m[2]).padStart(2, '0')}-${String(+m[1]).padStart(2, '0')}`;
  const d = new Date(s);
  return isNaN(d) ? null : ymd(d);
}
function num(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

/* ============================================================
   Column mapping — tolerant to Mongolian / English headers
   ============================================================ */
const ALIASES = {
  code: ['code', 'sku', 'код', 'барааны код', 'бар код', 'бар코д', 'дугаар', 'id', 'item code', 'артикул', 'item'],
  name: ['name', 'нэр', 'бараа', 'барааны нэр', 'тайлбар', 'product', 'description', 'барааны нэршил'],
  department: ['department', 'dept', 'хэлтэс', 'бүлэг', 'group', 'category', 'категори', 'чиглэл', 'төрөл', 'групп'],
  abc: ['abc', 'abc ангилал', 'abc class', 'abc-ангилал'],
  xyz: ['xyz', 'xyz ангилал', 'xyz class', 'xyz-ангилал'],
  cls: ['class', 'ангилал', '9 ангилал', 'abcxyz', 'abc/xyz', 'abc xyz', 'класс', 'angilal', '9-н ангилал'],
  status: ['status', 'төлөв', 'байдал', 'state', 'барааны төлөв'],
  price: ['price', 'үнэ', 'нэгж үнэ', 'нэгжийн үнэ', 'зарах үнэ', 'cost', 'өртөг', 'unit price'],
  unit: ['unit', 'нэгж', 'хэмжих нэгж', 'хэмжигдэхүүн', 'measure', 'uom'],
  minStock: ['min', 'minstock', 'min stock', 'доод хязгаар', 'нөөцийн доод', 'аюулгүй нөөц', 'safety stock', 'reorder', 'доод үлдэгдэл'],
  date: ['date', 'огноо', 'өдөр', 'тайлант огноо', 'тайлан огноо', 'он сар'],
  qty: ['qty', 'quantity', 'үлдэгдэл', 'тоо', 'тоо хэмжээ', 'balance', 'stock', 'on hand', 'onhand', 'remaining', 'хэмжээ', 'эцсийн үлдэгдэл'],
  salesQty: ['борлуулалт', 'зарагдсан', 'sold', 'sales', 'sales qty', 'борлуулсан', 'тоо', 'qty', 'хэмжээ', 'тоо хэмжээ'],
  amount: ['amount', 'дүн', 'орлого', 'revenue', 'борлуулалтын дүн', 'нийт дүн', 'үнийн дүн'],
  branch: ['branch', 'салбар', 'салбарын нэр', 'store', 'дэлгүүр', 'цэг', 'захиалагч'],
  orderQty: ['захиалга', 'захиалсан', 'order', 'ordered', 'order qty', 'захиалгын тоо', 'тоо', 'qty', 'хэмжээ', 'тоо хэмжээ'],
};
function normKey(k) { return String(k).trim().toLowerCase().replace(/\s+/g, ' '); }
function pick(row, keys, aliasName) {
  const aliases = ALIASES[aliasName];
  for (const k of keys) {
    for (const a of aliases) { if (k === a) return row[a]; }
  }
  return undefined;
}
/* row: object with normalized keys */
function normRow(raw) {
  const o = {}; for (const k in raw) o[normKey(k)] = raw[k]; return o;
}

function mapProduct(r) {
  const keys = Object.keys(r);
  let abc = String(pick(r, keys, 'abc') ?? '').trim().toUpperCase();
  let xyz = String(pick(r, keys, 'xyz') ?? '').trim().toUpperCase();
  let cls = String(pick(r, keys, 'cls') ?? '').trim().toUpperCase().replace(/[^AXBYCZ]/g, '');
  if ((!abc || !xyz) && cls.length >= 2) { abc = abc || cls[0]; xyz = xyz || cls[1]; }
  if (abc && xyz) cls = abc + xyz;
  cls = (abc + xyz).match(/^[ABC][XYZ]$/) ? abc + xyz : '';
  const code = String(pick(r, keys, 'code') ?? '').trim();
  if (!code) return null;
  return {
    code,
    name: String(pick(r, keys, 'name') ?? code).trim(),
    department: String(pick(r, keys, 'department') ?? 'Бусад').trim() || 'Бусад',
    abc, xyz, cls,
    status: String(pick(r, keys, 'status') ?? '').trim(),
    price: num(pick(r, keys, 'price')),
    unit: String(pick(r, keys, 'unit') ?? 'ш').trim() || 'ш',
    minStock: num(pick(r, keys, 'minStock')),
  };
}
function mapBalance(r) {
  const keys = Object.keys(r);
  const code = String(pick(r, keys, 'code') ?? '').trim();
  const date = toYmd(pick(r, keys, 'date'));
  if (!code || !date) return null;
  return { code, date, qty: num(pick(r, keys, 'qty')) };
}
function mapSale(r) {
  const keys = Object.keys(r);
  const code = String(pick(r, keys, 'code') ?? '').trim();
  const date = toYmd(pick(r, keys, 'date'));
  if (!code || !date) return null;
  return { code, date, qty: num(pick(r, keys, 'salesQty')), amount: num(pick(r, keys, 'amount')), branch: String(pick(r, keys, 'branch') ?? '').trim() };
}
function mapOrder(r) {
  const keys = Object.keys(r);
  const code = String(pick(r, keys, 'code') ?? '').trim();
  const date = toYmd(pick(r, keys, 'date'));
  if (!code || !date) return null;
  return { code, date, qty: num(pick(r, keys, 'orderQty')), branch: String(pick(r, keys, 'branch') ?? '').trim(), status: String(pick(r, keys, 'status') ?? '').trim() };
}

/* ============================================================
   File reading (CSV native + XLSX via SheetJS)
   ============================================================ */
function detectDelim(line) {
  const counts = { ',': 0, ';': 0, '\t': 0 };
  let inQ = false;
  for (const ch of line) {
    if (ch === '"') inQ = !inQ;
    else if (!inQ && counts[ch] != null) counts[ch]++;
  }
  return Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0] || ',';
}
function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const firstLine = text.slice(0, text.indexOf('\n') >= 0 ? text.indexOf('\n') : text.length);
  const delim = detectDelim(firstLine);
  const rows = []; let row = []; let field = ''; let inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
      else field += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === delim) { row.push(field); field = ''; }
      else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
      else if (c === '\r') { /* skip */ }
      else field += c;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  if (!rows.length) return [];
  const header = rows.shift().map(h => h.trim());
  return rows.filter(r => r.some(c => String(c).trim() !== '')).map(r => {
    const o = {}; header.forEach((h, i) => { o[h] = r[i] !== undefined ? r[i] : ''; }); return o;
  });
}
function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    const isCsv = /\.csv$/i.test(file.name) || file.type === 'text/csv';
    reader.onerror = () => reject(new Error('Файл уншихад алдаа гарлаа'));
    if (isCsv) {
      reader.onload = () => resolve(parseCSV(String(reader.result)));
      reader.readAsText(file, 'utf-8');
    } else {
      reader.onload = () => {
        try {
          if (typeof XLSX === 'undefined') throw new Error('XLSX сан ачаалагдаагүй (интернэт шалгана уу)');
          const wb = XLSX.read(new Uint8Array(reader.result), { type: 'array', cellDates: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json(ws, { defval: '', raw: true }));
        } catch (e) { reject(e); }
      };
      reader.readAsArrayBuffer(file);
    }
  });
}

/* Merge helpers (dedupe) */
function mergeBy(arr, incoming, keyFn) {
  const map = new Map(arr.map(x => [keyFn(x), x]));
  for (const x of incoming) map.set(keyFn(x), x);
  return [...map.values()];
}

async function importFile(file, type) {
  const raw = await readFile(file);
  const rows = raw.map(normRow);
  let added = 0;
  if (type === 'products') {
    const mapped = rows.map(mapProduct).filter(Boolean);
    DB.products = mergeBy(DB.products, mapped, x => x.code);
    added = mapped.length;
  } else if (type === 'balances') {
    const mapped = rows.map(mapBalance).filter(Boolean);
    DB.balances = mergeBy(DB.balances, mapped, x => x.code + '|' + x.date);
    added = mapped.length;
  } else if (type === 'sales') {
    const mapped = rows.map(mapSale).filter(Boolean);
    DB.sales = mergeBy(DB.sales, mapped, x => x.code + '|' + x.date + '|' + x.branch);
    added = mapped.length;
  } else if (type === 'orders') {
    const mapped = rows.map(mapOrder).filter(Boolean);
    DB.orders = mergeBy(DB.orders, mapped, x => x.code + '|' + x.date + '|' + x.branch);
    added = mapped.length;
  }
  DB.updatedAt = ymd(new Date());
  return added;
}

/* ============================================================
   Persistence
   ============================================================ */
function saveLocal() {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(DB)); S.source = 'local'; }
  catch (e) { toast('⚠️ Хадгалах багтаамж хүрсэнгүй. Илүү бага өгөгдөл оруулна уу.'); }
}
function loadLocal() {
  try { const raw = localStorage.getItem(STORE_KEY); return raw ? JSON.parse(raw) : null; }
  catch (e) { return null; }
}
const PUBLISHED_ENC = 'data/stock-data.enc';
async function fetchPublished() {
  // 1) plaintext (local working copy — gitignored, only present on the admin's machine)
  try {
    const res = await fetch(PUBLISHED_URL, { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      if (data && Array.isArray(data.products)) return { kind: 'plain', db: Object.assign(emptyDB(), data) };
    }
  } catch (e) { /* ignore */ }
  // 2) encrypted (the published artifact)
  try {
    const res = await fetch(PUBLISHED_ENC, { cache: 'no-store' });
    if (res.ok) return { kind: 'enc', buf: await res.arrayBuffer() };
  } catch (e) { /* ignore */ }
  return null;
}

/* ---------- Passcode-protected (encrypted) data ----------
   File layout: 'DCS1' | iter(uint32 BE) | salt(16) | iv(12) | ciphertext+tag
   AES-256-GCM, key = PBKDF2-SHA256(passcode, salt, iter). Matches build_data.py. */
async function decryptData(buf, passcode) {
  const u8 = new Uint8Array(buf);
  if (u8.length < 36 || String.fromCharCode(u8[0], u8[1], u8[2], u8[3]) !== 'DCS1') throw new Error('format');
  const iter = new DataView(buf).getUint32(4, false);
  const salt = u8.slice(8, 24), iv = u8.slice(24, 36), ct = u8.slice(36);
  const km = await crypto.subtle.importKey('raw', new TextEncoder().encode(passcode), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' },
    km, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
  const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, ct);
  return JSON.parse(new TextDecoder().decode(plain));
}
function askPasscode() {
  return new Promise(resolve => {
    const inp = $('gate-input'), btn = $('gate-btn');
    const go = () => { btn.onclick = null; inp.onkeydown = null; resolve(inp.value); };
    btn.onclick = go;
    inp.onkeydown = (e) => { if (e.key === 'Enter') go(); };
    setTimeout(() => inp.focus(), 50);
  });
}
async function loadEncrypted(buf) {
  const cached = sessionStorage.getItem('dcstock-pass');
  if (cached) { try { return Object.assign(emptyDB(), await decryptData(buf, cached)); } catch (e) { sessionStorage.removeItem('dcstock-pass'); } }
  $('gate').hidden = false; document.body.style.overflow = 'hidden';
  while (true) {
    const pass = await askPasscode();
    const btn = $('gate-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Тайлж байна…';
    try {
      const db = Object.assign(emptyDB(), await decryptData(buf, pass));
      sessionStorage.setItem('dcstock-pass', pass);
      $('gate').hidden = true; document.body.style.overflow = '';
      return db;
    } catch (e) {
      sessionStorage.removeItem('dcstock-pass');
      $('gate-err').hidden = false; $('gate-err').textContent = '🔒 Нууц код буруу байна. Дахин оролдоно уу.';
      $('gate-input').value = '';
    } finally {
      btn.disabled = false; btn.textContent = '🔓 Нээх';
    }
  }
}

/* ============================================================
   Indexing & computations
   ============================================================ */
function buildIndexes() {
  const productByCode = new Map(DB.products.map(p => [p.code, p]));
  // balances per code, sorted by date (for as-of lookups & sparklines)
  const balByCode = new Map();
  for (const b of DB.balances) {
    if (!balByCode.has(b.code)) balByCode.set(b.code, []);
    balByCode.get(b.code).push(b);
  }
  for (const arr of balByCode.values()) arr.sort((a, b) => a.date < b.date ? -1 : 1);

  // all dates present in data
  const dateSet = new Set();
  DB.balances.forEach(b => dateSet.add(b.date));
  DB.sales.forEach(s => dateSet.add(s.date));
  DB.orders.forEach(o => dateSet.add(o.date));
  const dates = [...dateSet].sort();

  // sales / orders qty per code per date
  const salesByCodeDate = new Map(); // code|date -> {qty, amount}
  for (const s of DB.sales) {
    const k = s.code + '|' + s.date;
    const cur = salesByCodeDate.get(k) || { qty: 0, amount: 0 };
    cur.qty += s.qty; cur.amount += s.amount; salesByCodeDate.set(k, cur);
  }
  const ordersByCodeDate = new Map();
  for (const o of DB.orders) {
    const k = o.code + '|' + o.date;
    const cur = ordersByCodeDate.get(k) || { qty: 0, amount: 0 };
    cur.qty += o.qty; cur.amount += (o.amount || 0); ordersByCodeDate.set(k, cur);
  }

  const departments = [...new Set(DB.products.map(p => p.department))].sort((a, b) => a.localeCompare(b, 'mn'));
  const classes = ['AX', 'AY', 'AZ', 'BX', 'BY', 'BZ', 'CX', 'CY', 'CZ'];
  const statuses = [...new Set(DB.products.map(p => p.status).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'mn'));
  const statuses2 = [...new Set(DB.products.map(p => p.status2).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'mn'));
  const salesDates = [...new Set(DB.sales.map(s => s.date))].sort();
  const orderDates = [...new Set(DB.orders.map(o => o.date))].sort();

  IX = { productByCode, balByCode, dates, salesByCodeDate, ordersByCodeDate, departments, classes, statuses, statuses2, salesDates, orderDates };
}

/* latest date in a sorted list that is on or before `d` */
function latestWith(dates, d) {
  let r = null;
  for (const x of dates) { if (x <= d) r = x; else break; }
  return r;
}

/* balance of a product as-of a date (carry last known forward) */
function balanceAsOf(code, date) {
  const arr = IX.balByCode.get(code);
  if (!arr || !arr.length) return null;
  let val = null;
  for (const b of arr) { if (b.date <= date) val = b.qty; else break; }
  return val;
}
function salesOn(code, date) { return (IX.salesByCodeDate.get(code + '|' + date) || { qty: 0, amount: 0 }); }
function ordersOn(code, date) { return (IX.ordersByCodeDate.get(code + '|' + date) || { qty: 0, amount: 0 }); }
function prevDate(date) {
  const i = IX.dates.indexOf(date);
  return i > 0 ? IX.dates[i - 1] : null;
}

/* products passing dept/status/search filters (NOT class) */
function baseFiltered() {
  const q = S.search.trim().toLowerCase();
  return DB.products.filter(p => {
    if (S.dept && p.department !== S.dept) return false;
    if (S.status && p.status !== S.status) return false;
    if (S.status2 && p.status2 !== S.status2) return false;
    if (q && !(p.name.toLowerCase().includes(q) || p.code.toLowerCase().includes(q))) return false;
    return true;
  });
}
function filteredProducts() {
  return baseFiltered().filter(p => !S.cls || p.cls === S.cls);
}

/* enrich a product with computed values on selected date */
function enrich(p) {
  const d = S.selDate;
  const pd = prevDate(d);
  const qty = balanceAsOf(p.code, d);
  const qtyPrev = pd ? balanceAsOf(p.code, pd) : null;
  const delta = (qty != null && qtyPrev != null) ? qty - qtyPrev : null;
  const value = qty != null ? qty * p.price : null;
  const sales = salesOn(p.code, d).qty;
  const orders = ordersOn(p.code, d).qty;
  const low = qty != null && (qty <= 0 || (p.minStock > 0 && qty <= p.minStock));
  // Өдрийн дундаж зарцуулалт — сонгосон огноо хүртэлх сүүлийн ~сарын дата-аар (огноогоор хувирна)
  const salesRate = RANGE.sales.get(p.code) || 0;   // борлуулалт/өдөр
  const ordersRate = RANGE.orders.get(p.code) || 0; // захиалга/өдөр
  // Нөөц хоног (эргэц) = өнөөдрийн үлдэгдэл / өдрийн дундаж зарцуулалт
  const coverSales = (qty != null && salesRate > 0) ? qty / salesRate : null;
  const coverOrders = (qty != null && ordersRate > 0) ? qty / ordersRate : null;
  return { ...p, qty, qtyPrev, delta, value, sales, orders, low, salesRate, ordersRate, coverSales, coverOrders };
}
function fmtDays(n) {
  if (n == null) return '<span style="color:var(--text-3)">—</span>';
  if (n >= 999) return '999+';
  return (n < 10 ? n.toFixed(1) : Math.round(n).toString());
}

/* Өдрийн дундаж зарцуулалтын хурд — сонгосон огноо хүртэлх сүүлийн 30 хоногийн дата-аар (бараа тус бүрээр) */
let RANGE = { sales: new Map(), orders: new Map(), sDates: [], oDates: [] };
function computeRangeRates() {
  if (!IX || !S.selDate) { RANGE = { sales: new Map(), orders: new Map(), sDates: [], oDates: [] }; return; }
  // Тухайн сарын 1-нээс сонгосон өдөр хүртэл (month-to-date). Жнь 05-05 → 05-01..05-05.
  const monthStart = S.selDate.slice(0, 7) + '-01';
  const inRange = (d) => d >= monthStart && d <= S.selDate;
  const sDates = IX.salesDates.filter(inRange);
  const oDates = IX.orderDates.filter(inRange);
  const nS = sDates.length || 1, nO = oDates.length || 1;
  const sSum = new Map(), oSum = new Map();
  for (const s of DB.sales) if (inRange(s.date)) sSum.set(s.code, (sSum.get(s.code) || 0) + s.qty);
  for (const o of DB.orders) if (inRange(o.date)) oSum.set(o.code, (oSum.get(o.code) || 0) + o.qty);
  const sR = new Map(), oR = new Map();
  for (const [c, v] of sSum) sR.set(c, v / nS);
  for (const [c, v] of oSum) oR.set(c, v / nO);
  RANGE = { sales: sR, orders: oR, sDates, oDates };
}

/* ============================================================
   Sample data generator (demo)
   ============================================================ */
function mulberry32(a) { return function () { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
function buildSample() {
  const rnd = mulberry32(20260531);
  const ri = (lo, hi) => Math.floor(lo + rnd() * (hi - lo + 1));
  const DAYS = 45;
  const branches = ['Төв салбар', 'Зүүн салбар', 'Баруун салбар', 'Онлайн'];
  const defs = [
    ['Хүнс', [['Гурил 1кг', 2800], ['Цагаан будаа 5кг', 18500], ['Ургамлын тос 1л', 6400], ['Чихэр 1кг', 4200], ['Давс 1кг', 1500], ['Гоймон 400г', 2100], ['Сүү 1л', 3300], ['Талх', 2500]]],
    ['Ундаа', [['Кока-Кола 0.5л', 2500], ['Ус 1.5л', 1800], ['Жүүс 1л', 4500], ['Эрчим хүчний ундаа', 5500], ['Цай 100г', 6800]]],
    ['Ахуйн бараа', [['Гар саван', 3900], ['Нунтаг саван 3кг', 19500], ['Цэвэрлэгээний бодис', 8700], ['Шүдний оо', 5600], ['Жорлонгийн цаас 8ш', 12500]]],
    ['Цахилгаан', [['LED чийдэн 9W', 7800], ['Батарей AA 4ш', 6500], ['Утас цэнэглэгч', 22000], ['Чихэвч', 38000], ['Адаптер 220V', 14500]]],
    ['Гоо сайхан', [['Шампунь 400мл', 12800], ['Арьсны тос', 24500], ['Нүүрний маск', 4500], ['Үнэртэн жижиг', 45000]]],
  ];
  const statuses = ['Идэвхтэй', 'Идэвхтэй', 'Идэвхтэй', 'Идэвхтэй', 'Шинэ', 'Дуусч буй', 'Зогссон'];
  const xyzPool = ['X', 'X', 'Y', 'Y', 'Z'];
  const products = [], balances = [], sales = [], orders = [];
  let idx = 0;
  for (const [dept, items] of defs) {
    for (const [nm, price] of items) {
      idx++;
      const code = 'P' + String(idx).padStart(3, '0');
      // ABC by price tier + a bit of randomness
      const abc = price >= 18000 ? 'A' : price >= 5000 ? (rnd() < .6 ? 'B' : 'A') : (rnd() < .6 ? 'C' : 'B');
      const xyz = xyzPool[ri(0, xyzPool.length - 1)];
      const status = statuses[ri(0, statuses.length - 1)];
      const baseSales = abc === 'A' ? ri(28, 55) : abc === 'B' ? ri(12, 26) : ri(3, 11);
      const cv = xyz === 'X' ? 0.18 : xyz === 'Y' ? 0.45 : 0.95;
      const minStock = Math.round(baseSales * 3);
      const status2 = (status === 'Зогссон' || status === 'Дуусч буй') ? 'Non Active' : 'Active';
      products.push({ code, name: nm, department: dept, abc, xyz, cls: abc + xyz, status, status2, price, sellPrice: Math.round(price / 0.72), unit: 'ш', minStock, avgSales: baseSales, avgOrders: Math.max(1, Math.round(baseSales * 0.4)) });

      const today = new Date();
      let stock = Math.round(baseSales * ri(8, 16));
      const newFrom = status === 'Шинэ' ? ri(15, 30) : 0;
      let stopped = status === 'Зогссон';
      for (let i = DAYS - 1; i >= 0; i--) {
        const d = new Date(today); d.setDate(today.getDate() - i);
        const date = ymd(d);
        const dayIndex = DAYS - 1 - i;
        const active = dayIndex >= newFrom && !(stopped && dayIndex > DAYS - 8);
        let sold = 0;
        if (active) {
          const wkBoost = (d.getDay() === 5 || d.getDay() === 6) ? 1.3 : 1;
          sold = Math.max(0, Math.round((baseSales * wkBoost) * (1 + (rnd() * 2 - 1) * cv)));
          sold = Math.min(sold, stock);
        }
        stock -= sold;
        // replenish when low — but let "Дуусч буй" items drift low in the last 10 days
        const noRestock = status === 'Дуусч буй' && i < 10;
        if (active && stock <= minStock && !noRestock) { stock += Math.round(baseSales * ri(9, 16)); }
        balances.push({ code, date, qty: stock });
        if (sold > 0) sales.push({ code, date, qty: sold, amount: sold * price, branch: branches[ri(0, branches.length - 1)] });
        // branch orders a few times a week for important items
        if (active && rnd() < (abc === 'A' ? .5 : abc === 'B' ? .3 : .15)) {
          const b = branches[ri(0, branches.length - 1)];
          orders.push({ code, date, qty: Math.max(1, Math.round(baseSales * (0.2 + rnd() * 0.5))), branch: b, status: rnd() < .7 ? 'Биелсэн' : 'Хүлээгдэж буй' });
        }
      }
    }
  }
  return { products, balances, sales, orders, updatedAt: ymd(new Date()) };
}

/* ============================================================
   Rendering
   ============================================================ */
let histChart = null, pdChart = null;

function render() {
  const hasData = DB.products.length > 0;
  $('content').querySelectorAll('.card, .kpis, .grid-2').forEach(n => n.hidden = !hasData);
  $('empty-state').hidden = hasData;
  $('toolbar').style.display = hasData ? '' : 'none';
  updateStatusBar();
  if (!hasData) return;
  computeRangeRates();
  renderFilters();
  renderKPIs();
  renderMatrix();
  renderDepartments();
  renderHistScope();
  renderHistory();
  renderTable();
}

function updateStatusBar() {
  const srcLabel = { local: 'Локал ноорог', published: 'Нийтлэгдсэн', sample: '🧪 Жишээ горим (туршилт)', empty: '' }[S.source] || '';
  if (!DB.products.length) { $('data-status').textContent = 'Өгөгдөл алга'; $('as-of').textContent = ''; return; }
  const histDays = IX ? IX.dates.length : 0;
  $('data-status').innerHTML = `${esc(srcLabel)} · ${fmtInt(DB.products.length)} бараа · ${histDays} өдрийн түүх` +
    (DB.updatedAt ? ` · шинэчилсэн ${esc(DB.updatedAt)}` : '');
  $('as-of').textContent = 'Байдлаар: ' + fmtDateFull(S.selDate);
}

function fillSelect(sel, values, current, allLabel) {
  const cur = current;
  sel.innerHTML = (allLabel != null ? `<option value="">${allLabel}</option>` : '') +
    values.map(v => `<option value="${esc(v.value)}"${v.value === cur ? ' selected' : ''}>${esc(v.label)}</option>`).join('');
}
function renderFilters() {
  // date
  const dateOpts = [...IX.dates].reverse().map(d => ({ value: d, label: fmtDateFull(d) + (d === IX.dates[IX.dates.length - 1] ? ' — хамгийн сүүлийн' : '') }));
  fillSelect($('f-date'), dateOpts, S.selDate, null);
  fillSelect($('f-dept'), IX.departments.map(d => ({ value: d, label: d })), S.dept, 'Бүх Department');
  fillSelect($('f-class'), IX.classes.map(c => ({ value: c, label: c })), S.cls, 'Бүх ангилал');
  fillSelect($('f-status'), IX.statuses.map(s => ({ value: s, label: s })), S.status, 'Бүх төлөв');
  fillSelect($('f-status2'), IX.statuses2.map(s => ({ value: s, label: s })), S.status2, 'Бүх төлөв');
  if ($('f-search').value !== S.search) $('f-search').value = S.search;
}

function renderKPIs() {
  const list = filteredProducts().map(enrich);
  const withQty = list.filter(p => p.qty != null);
  const totalQty = withQty.reduce((s, p) => s + p.qty, 0);
  const totalVal = withQty.reduce((s, p) => s + (p.value || 0), 0);
  const totalValPrev = withQty.reduce((s, p) => s + ((p.qtyPrev != null ? p.qtyPrev : p.qty) * p.price), 0);
  const valDelta = totalVal - totalValPrev;
  const totalAvgSales = withQty.reduce((s, p) => s + (p.salesRate || 0), 0);
  const avgCover = totalAvgSales > 0 ? totalQty / totalAvgSales : null;
  const low = list.filter(p => p.low).length;
  const zero = list.filter(p => p.qty === 0).length;
  // Өдрийн дундаж — шүүсэн бараа, сонгосон огноо хүртэлх сүүлийн ~сарын дата-аар (RANGE; огноо/филтерээр хувирна)
  const avgRetailDaily = list.reduce((s, p) => s + (p.salesRate || 0), 0);
  const avgOrdersDaily = list.reduce((s, p) => s + (p.ordersRate || 0), 0);
  const rangeLabel = (ds) => ds.length ? `${fmtDate(ds[0])}–${fmtDate(ds[ds.length - 1])} · ${ds.length} өдөр` : 'мэдээлэл алга';
  const salesRangeLabel = rangeLabel(RANGE.sDates);
  const ordRangeLabel = rangeLabel(RANGE.oDates);

  const deltaHtml = (d) => {
    if (!d) return '<span style="color:var(--text-3)">өөрчлөлтгүй</span>';
    const cls = d > 0 ? 'up' : 'down';
    return `<span class="${cls}">${d > 0 ? '▲' : '▼'} ${moneyShort(Math.abs(d))}</span> өчигдрөөс`;
  };

  const cards = [
    { cls: 'violet', label: '📦 Нийт бараа', value: fmtInt(list.length), sub: `${IX.departments.length} Department` },
    { cls: '', label: '🧮 Өнөөдрийн нийт үлдэгдэл', value: fmtInt(totalQty) + ' ш', sub: `${fmtInt(withQty.length)} нэр төрөл` },
    { cls: 'green', label: '💰 Нөөцийн үнэ (өртөг)', value: moneyShort(totalVal), sub: deltaHtml(valDelta) },
    { cls: 'violet', label: '⏳ Дундаж нөөц хоног', value: avgCover == null ? '—' : fmtDays(avgCover) + ' хоног', sub: 'борлуулалтаар (жигнэсэн)' },
    { cls: low ? 'red' : 'green', key: 'low', label: '⚠️ Анхаарах бараа', value: fmtInt(low), sub: `${fmtInt(zero)} нь дууссан · харах →` },
    { cls: 'amber', label: '🛒 Өдрийн дундаж борлуулалт', value: fmtInt(avgRetailDaily) + ' ш/өдөр', sub: salesRangeLabel },
    { cls: '', label: '📋 Өдрийн дундаж захиалга', value: fmtInt(avgOrdersDaily) + ' ш/өдөр', sub: ordRangeLabel },
  ];
  $('kpis').innerHTML = cards.map(c => `
    <div class="kpi ${c.cls}${c.key === 'low' ? ' clickable' + (S.lowOnly ? ' active' : '') : ''}"${c.key ? ` data-kpi="${c.key}"` : ''}>
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-sub">${c.sub}</div>
    </div>`).join('');
  const lowCard = $('kpis').querySelector('[data-kpi="low"]');
  if (lowCard) lowCard.onclick = () => {
    S.lowOnly = !S.lowOnly;
    render();
  };
}

function renderMatrix() {
  const metric = S.matrixMetric || 'value';
  const list = baseFiltered().map(enrich);
  const cells = {};
  IX.classes.forEach(c => cells[c] = { count: 0, value: 0, qty: 0, aS: 0, aO: 0 });
  let unclassified = 0;
  list.forEach(p => {
    const cell = cells[p.cls];
    if (!cell) { unclassified++; return; }
    cell.count++; cell.value += (p.value || 0);
    if (p.qty != null) { cell.qty += p.qty; cell.aS += (p.salesRate || 0); cell.aO += (p.ordersRate || 0); }
  });
  const metricOf = (cell) => metric === 'value' ? cell.value
    : metric === 'coverSales' ? (cell.aS > 0 ? cell.qty / cell.aS : null)
    : (cell.aO > 0 ? cell.qty / cell.aO : null);
  const fmtMetric = (cell) => {
    if (metric === 'value') return moneyShort(cell.value);
    const v = metricOf(cell); return v == null ? '—' : fmtDays(v) + ' хоног';
  };
  let maxM = 0;
  IX.classes.forEach(c => { const v = metricOf(cells[c]); if (v != null && v > maxM) maxM = v; });

  const cols = ['X', 'Y', 'Z'], rows = ['A', 'B', 'C'];
  const colDesc = { X: 'тогтвортой', Y: 'дунд', Z: 'тогтворгүй' };
  const rowDesc = { A: 'өндөр үнэ', B: 'дунд', C: 'бага' };
  let html = '<div class="mx-corner"></div>';
  cols.forEach(c => html += `<div class="mx-col-h">${c}<small>${colDesc[c]}</small></div>`);
  rows.forEach(r => {
    html += `<div class="mx-row-h">${r}<small>${rowDesc[r]}</small></div>`;
    cols.forEach(c => {
      const code = r + c, cell = cells[code];
      const mv = metricOf(cell);
      const intensity = maxM ? (mv || 0) / maxM : 0;
      const bg = `rgba(37,99,235,${(0.06 + intensity * 0.32).toFixed(3)})`;
      const active = S.cls === code ? ' active' : '';
      html += `<div class="mx-cell${active}" data-cls="${code}" style="background:${bg}">
        <div class="mx-code">${code}</div>
        <div class="mx-count">${cell.count}<small> бараа</small></div>
        <div class="mx-val">${fmtMetric(cell)}</div>
      </div>`;
    });
  });
  $('matrix').innerHTML = html;
  const hint = metric === 'value' ? '' : 'Нөөц хоног = үлдэгдэл ÷ өдрийн дундаж (жигнэсэн)';
  $('matrix-hint').textContent = (unclassified ? `${unclassified} бараа ангилаагүй · ` : '') + hint;
  $('matrix').querySelectorAll('.mx-cell').forEach(c => c.onclick = () => {
    S.cls = (S.cls === c.dataset.cls) ? '' : c.dataset.cls;
    render();
  });
}

function sparkline(values, w = 90, h = 30) {
  if (!values.length) return '';
  const min = Math.min(...values), max = Math.max(...values), range = max - min || 1;
  const step = values.length > 1 ? w / (values.length - 1) : 0;
  const pts = values.map((v, i) => `${(i * step).toFixed(1)},${(h - 3 - ((v - min) / range) * (h - 6)).toFixed(1)}`).join(' ');
  const last = values[values.length - 1], first = values[0];
  const color = last >= first ? 'var(--green)' : 'var(--red)';
  return `<svg class="dept-spark" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">
    <polyline points="${pts}" fill="none" stroke="${color}" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round"/>
  </svg>`;
}

function renderDepartments() {
  const d = S.selDate, pd = prevDate(d);
  const last7 = IX.dates.slice(-7);
  const byDept = {};
  baseFiltered().forEach(p => {
    if (S.cls && p.cls !== S.cls) return;
    if (!byDept[p.department]) byDept[p.department] = { qty: 0, prev: 0, value: 0, codes: [] };
    const q = balanceAsOf(p.code, d); const qp = pd ? balanceAsOf(p.code, pd) : null;
    if (q != null) { byDept[p.department].qty += q; byDept[p.department].value += q * p.price; }
    if (qp != null) byDept[p.department].prev += qp;
    byDept[p.department].codes.push(p.code);
  });
  const rows = Object.entries(byDept).sort((a, b) => b[1].value - a[1].value);
  if (!rows.length) { $('dept-list').innerHTML = '<div class="hint" style="padding:20px;text-align:center">Өгөгдөл алга</div>'; return; }
  $('dept-hint').textContent = fmtDate(d) + ' байдлаар';
  $('dept-list').innerHTML = rows.map(([name, v]) => {
    const series = last7.map(dt => v.codes.reduce((s, c) => s + (balanceAsOf(c, dt) || 0), 0));
    const delta = v.qty - v.prev;
    const dcls = delta > 0 ? 'delta-up' : delta < 0 ? 'delta-down' : 'delta-flat';
    const dtxt = delta === 0 ? '—' : `${delta > 0 ? '▲' : '▼'} ${fmtInt(Math.abs(delta))}`;
    return `<div class="dept-row">
      <div class="dept-name">${esc(name)}<small>${v.codes.length} бараа · ${moneyShort(v.value)}</small></div>
      ${sparkline(series)}
      <div class="dept-qty">${fmtInt(v.qty)}<small> ш</small></div>
      <div class="dept-delta ${dcls}">${dtxt}</div>
    </div>`;
  }).join('');
}

function renderHistScope() {
  const opts = [{ value: '__all__', label: 'Нийт (бүх бараа)' }]
    .concat(IX.departments.map(d => ({ value: 'dept:' + d, label: 'Department: ' + d })));
  // keep current selection if still valid
  if (!opts.some(o => o.value === S.histScope)) S.histScope = '__all__';
  fillSelect($('hist-scope'), opts, S.histScope, null);
  $('hist-range').value = String(S.histRange);
  $('hist-metric').querySelectorAll('.seg-btn').forEach(b => b.classList.toggle('active', b.dataset.v === S.histMetric));
}

function histCodes() {
  if (S.histScope.startsWith('dept:')) {
    const dep = S.histScope.slice(5);
    return DB.products.filter(p => p.department === dep && (!S.cls || p.cls === S.cls)).map(p => p.code);
  }
  return baseFiltered().filter(p => !S.cls || p.cls === S.cls).map(p => p.code);
}
function renderHistory() {
  let dates = IX.dates.slice();
  if (S.histRange > 0) dates = dates.slice(-S.histRange);
  const codes = new Set(histCodes());
  const series = dates.map(dt => {
    if (S.histMetric === 'balance') {
      let s = 0; codes.forEach(c => { const q = balanceAsOf(c, dt); if (q != null) s += q; }); return s;
    }
    if (S.histMetric === 'sales') {
      let s = 0; DB.sales.forEach(x => { if (x.date === dt && codes.has(x.code)) s += x.qty; }); return s;
    }
    let s = 0; DB.orders.forEach(x => { if (x.date === dt && codes.has(x.code)) s += x.qty; }); return s;
  });
  const labels = dates.map(fmtDate);
  const metricLabel = { balance: 'Үлдэгдэл (ш)', sales: 'Борлуулалт (ш)', orders: 'Захиалга (ш)' }[S.histMetric];
  const color = { balance: '#2563eb', sales: '#d97706', orders: '#6c5ce7' }[S.histMetric];
  drawLine($('hist-chart'), labels, [{ label: metricLabel, data: series, color, fill: S.histMetric === 'balance' }], h => h);
}

function drawLine(canvas, labels, datasets, fmt) {
  const cfg = {
    type: 'line',
    data: {
      labels,
      datasets: datasets.map(d => ({
        label: d.label, data: d.data, borderColor: d.color,
        backgroundColor: d.fill ? d.color + '22' : 'transparent', fill: !!d.fill,
        tension: 0.3, borderWidth: 2.2, pointRadius: 0, pointHoverRadius: 4, pointHoverBackgroundColor: d.color,
      })),
    },
    options: {
      responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: datasets.length > 1, labels: { boxWidth: 12, font: { size: 12 } } },
        tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${fmtInt(c.parsed.y)}` } },
      },
      scales: {
        x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 10, font: { size: 11 }, color: '#8a93a6' } },
        y: { beginAtZero: true, grid: { color: '#eef0f5' }, ticks: { callback: (v) => fmtInt(v), font: { size: 11 }, color: '#8a93a6' } },
      },
    },
  };
  if (canvas === $('hist-chart')) { if (histChart) histChart.destroy(); histChart = new Chart(canvas, cfg); }
  else { if (pdChart) pdChart.destroy(); pdChart = new Chart(canvas, cfg); }
}

/* ---------- status chip ---------- */
function statusChip(status) {
  if (!status) return '<span class="chip st-none">—</span>';
  const s = status.toLowerCase();
  let cls = 'st-none';
  if (/устсан|устгагд|deleted/.test(s)) cls = 'st-stop';
  else if (/non[\s-]*active|идэвхгүй/.test(s)) cls = 'st-warn';
  else if (/түр зогс|агуулах дээр|дуусч|бага|low|анхаар|хүрэлц/.test(s)) cls = 'st-warn';
  else if (/идэвх|active|хэвийн|байгаа|энгийн/.test(s)) cls = 'st-active';
  else if (/шинэ|new/.test(s)) cls = 'st-new';
  else if (/зогс|stop|discont|болиул|хасагд/.test(s)) cls = 'st-stop';
  return `<span class="chip ${cls}">${esc(status)}</span>`;
}
function clsChip(cls) {
  if (!cls) return '<span class="chip cls-none">—</span>';
  return `<span class="chip cls-chip cls-${cls[0]}">${esc(cls)}</span>`;
}

function renderTable() {
  let list = filteredProducts().map(enrich);
  if (S.lowOnly) list = list.filter(p => p.low);
  const k = S.sortKey, dir = S.sortDir;
  list.sort((a, b) => {
    let va = a[k], vb = b[k];
    if (k === 'name' || k === 'department' || k === 'status' || k === 'status2' || k === 'cls' || k === 'code') {
      va = String(va || ''); vb = String(vb || ''); return va.localeCompare(vb, 'mn') * dir;
    }
    va = va == null ? -Infinity : va; vb = vb == null ? -Infinity : vb;
    return (va - vb) * dir;
  });
  $('prod-count').textContent = fmtInt(list.length);
  document.querySelectorAll('#prod-table th').forEach(th => {
    th.classList.toggle('sorted', th.dataset.sort === k);
    th.classList.toggle('desc', th.dataset.sort === k && dir < 0);
  });
  const tbody = $('prod-tbody');
  if (!list.length) { tbody.innerHTML = `<tr><td colspan="11" style="text-align:center;color:var(--text-3);padding:30px">Шүүлтүүрт тохирох бараа алга</td></tr>`; return; }
  const TABLE_CAP = 300;
  const shown = list.slice(0, TABLE_CAP);
  const capNote = list.length > TABLE_CAP
    ? `<tr><td colspan="11" style="text-align:center;color:var(--text-3);padding:14px">${fmtInt(list.length)}-аас эхний ${TABLE_CAP}-г харууллаа · хайлт/шүүлтүүрээр нарийсгана уу</td></tr>`
    : '';
  tbody.innerHTML = shown.map(p => {
    const qtyTxt = p.qty == null ? '<span style="color:var(--text-3)">—</span>'
      : `<span class="${p.qty === 0 ? 'qty-zero' : p.low ? 'qty-low' : ''}">${p.low ? '<span class="warn-dot"></span>' : ''}${fmtInt(p.qty)}</span>`;
    let deltaTxt = '—', dcls = '';
    if (p.delta != null && p.delta !== 0) { dcls = p.delta > 0 ? 'pos' : 'neg'; deltaTxt = `${p.delta > 0 ? '+' : ''}${fmtInt(p.delta)}`; }
    return `<tr data-code="${esc(p.code)}">
      <td class="code-cell">${esc(p.code)}</td>
      <td class="name-cell">${esc(p.name)}</td>
      <td>${esc(p.department)}</td>
      <td class="center">${clsChip(p.cls)}</td>
      <td>${statusChip(p.status)}</td>
      <td>${statusChip(p.status2)}</td>
      <td class="num num-cell">${qtyTxt}</td>
      <td class="num delta-cell ${dcls}">${deltaTxt}</td>
      <td class="num num-cell">${p.value == null ? '—' : moneyShort(p.value)}</td>
      <td class="num num-cell"><span class="${p.coverSales != null && p.coverSales < 3 ? 'qty-low' : ''}">${fmtDays(p.coverSales)}</span></td>
      <td class="num num-cell"><span class="${p.coverOrders != null && p.coverOrders < 3 ? 'qty-low' : ''}">${fmtDays(p.coverOrders)}</span></td>
    </tr>`;
  }).join('') + capNote;
  tbody.querySelectorAll('tr[data-code]').forEach(tr => tr.onclick = () => openProduct(tr.dataset.code));
}

/* ============================================================
   Product detail modal
   ============================================================ */
function openProduct(code) {
  const p = IX.productByCode.get(code); if (!p) return;
  const e = enrich(p);
  const dates = IX.dates.slice();
  $('modal').classList.add('wide');
  $('modal-title').textContent = `${p.name} (${p.code})`;
  $('modal-body').innerHTML = `
    <div class="pd-head">${clsChip(p.cls)} <span class="chip cls-none">${esc(p.department)}</span>
      <span class="pd-stlbl">Нийлүүл:</span>${statusChip(p.status)}
      <span class="pd-stlbl">Захих:</span>${statusChip(p.status2)}</div>
    <div class="pd-meta">
      <div class="m"><div class="l">Өнөөдрийн үлдэгдэл</div><div class="v">${e.qty == null ? '—' : fmtInt(e.qty) + ' ' + esc(p.unit)}</div></div>
      <div class="m"><div class="l">Үнийн дүн (өртөг)</div><div class="v">${e.value == null ? '—' : moneyShort(e.value)}</div></div>
      <div class="m"><div class="l">Нөөц хоног (борлуулалт)</div><div class="v">${fmtDays(e.coverSales)}</div></div>
      <div class="m"><div class="l">Нөөц хоног (зах.)</div><div class="v">${fmtDays(e.coverOrders)}</div></div>
      <div class="m"><div class="l">Өдрийн дундаж борлуулалт</div><div class="v">${e.salesRate ? e.salesRate.toFixed(1) : '—'}</div></div>
      <div class="m"><div class="l">Өдрийн дундаж захиалга</div><div class="v">${e.ordersRate ? e.ordersRate.toFixed(1) : '—'}</div></div>
      <div class="m"><div class="l">Өртөг үнэ</div><div class="v">${money(p.price)}</div></div>
      <div class="m"><div class="l">Худалдах үнэ</div><div class="v">${p.sellPrice ? money(p.sellPrice) : '—'}</div></div>
      ${p.salesPeriodQty ? `<div class="m"><div class="l">Жижиглэн борл. (нийт)</div><div class="v">${fmtInt(p.salesPeriodQty)} ш</div></div>` : ''}
    </div>
    <div class="pd-chart"><canvas id="pd-canvas"></canvas></div>
    <div class="hint" style="text-align:center">Сүүлийн ${dates.length} өдрийн үлдэгдэл · борлуулалт · захиалга</div>`;
  showOverlay();
  const balSeries = dates.map(dt => { const q = balanceAsOf(code, dt); return q == null ? null : q; });
  const salesSeries = dates.map(dt => salesOn(code, dt).qty);
  const orderSeries = dates.map(dt => ordersOn(code, dt).qty);
  drawLine($('pd-canvas'), dates.map(fmtDate), [
    { label: 'Үлдэгдэл', data: balSeries, color: '#2563eb', fill: true },
    { label: 'Борлуулалт', data: salesSeries, color: '#16a34a', fill: false },
    { label: 'Захиалга', data: orderSeries, color: '#d97706', fill: false },
  ], h => h);
}

/* ============================================================
   Import modal
   ============================================================ */
const IMPORT_DEFS = [
  { type: 'products', title: '📦 Бараа (ABC/XYZ + төлөв)', cols: 'код, нэр, хэлтэс, abc, xyz, төлөв, үнэ, доод хязгаар', tmpl: 'код,нэр,хэлтэс,abc,xyz,төлөв,үнэ,доод хязгаар\nP001,Гурил 1кг,Хүнс,A,X,Идэвхтэй,2800,90\nP002,Шампунь,Гоо сайхан,B,Y,Шинэ,12800,30' },
  { type: 'balances', title: '🧮 Өдрийн үлдэгдэл', cols: 'огноо, код, үлдэгдэл', tmpl: 'огноо,код,үлдэгдэл\n2026-05-31,P001,420\n2026-05-31,P002,85' },
  { type: 'sales', title: '🛒 Борлуулалт', cols: 'огноо, код, тоо, дүн, салбар', tmpl: 'огноо,код,тоо,дүн,салбар\n2026-05-31,P001,45,126000,Төв салбар' },
  { type: 'orders', title: '📋 Салбарын захиалга', cols: 'огноо, код, салбар, тоо, төлөв', tmpl: 'огноо,код,салбар,тоо,төлөв\n2026-05-31,P001,Зүүн салбар,30,Хүлээгдэж буй' },
];
function openImport() {
  $('modal').classList.remove('wide');
  $('modal-title').textContent = 'Excel / CSV импорт';
  const counts = { products: DB.products.length, balances: DB.balances.length, sales: DB.sales.length, orders: DB.orders.length };
  $('modal-body').innerHTML = `
    <div class="import-grid">
      ${IMPORT_DEFS.map(d => `
        <div class="drop" data-type="${d.type}">
          <h4>${d.title}</h4>
          <div class="cols">Багана: <code>${d.cols.split(', ').join('</code> <code>')}</code></div>
          <div class="drop-row">
            <label class="file-btn">Файл сонгох<input type="file" accept=".csv,.xlsx,.xls" data-type="${d.type}"></label>
            <button class="tmpl" data-tmpl="${d.type}">Загвар татах</button>
            <span class="stat" data-stat="${d.type}">${counts[d.type] ? counts[d.type] + ' мөр орсон' : ''}</span>
          </div>
        </div>`).join('')}
    </div>
    <div class="import-note">
      💡 <b>Зөвлөмж:</b> Excel-ийн эхний мөр нь баганы нэр байх ёстой. Баганы нэрийг монгол эсвэл англиар бичсэн ч таних бөгөөд эрэмбэ хамаагүй.
      <code>код</code> нь бараа, үлдэгдэл, борлуулалт, захиалгын файлуудыг холбоно. Огноо: <code>2026-05-31</code> хэлбэрээр.
      ABC, XYZ-ийг тус тусад нь эсвэл нэг <code>ангилал</code> баганад (жнь <code>AX</code>) өгч болно.
    </div>
    <div class="import-foot">
      <button class="btn danger sm" id="imp-clear">🗑 Бүх өгөгдөл устгах</button>
      <button class="btn" id="imp-done">Болсон</button>
    </div>`;
  showOverlay();

  $('modal-body').querySelectorAll('input[type=file]').forEach(inp => {
    inp.onchange = async () => { if (inp.files[0]) await handleImport(inp.files[0], inp.dataset.type); inp.value = ''; };
  });
  $('modal-body').querySelectorAll('.tmpl').forEach(b => b.onclick = () => {
    const def = IMPORT_DEFS.find(d => d.type === b.dataset.tmpl);
    downloadFile(def.type + '-zagvar.csv', '﻿' + def.tmpl, 'text/csv');
  });
  // drag & drop
  $('modal-body').querySelectorAll('.drop').forEach(zone => {
    zone.ondragover = (e) => { e.preventDefault(); zone.classList.add('over'); };
    zone.ondragleave = () => zone.classList.remove('over');
    zone.ondrop = async (e) => { e.preventDefault(); zone.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) await handleImport(f, zone.dataset.type); };
  });
  $('imp-clear').onclick = () => {
    if (!confirm('Бүх бараа, үлдэгдэл, борлуулалт, захиалгын өгөгдлийг устгах уу?')) return;
    DB = emptyDB(); localStorage.removeItem(STORE_KEY); S.source = 'empty'; S.selDate = null;
    buildIndexes(); closeOverlay(); render(); toast('Өгөгдөл устгагдлаа');
  };
  $('imp-done').onclick = () => closeOverlay();
}
async function handleImport(file, type) {
  const stat = $('modal-body').querySelector(`[data-stat="${type}"]`);
  if (stat) { stat.className = 'stat'; stat.textContent = 'Уншиж байна…'; }
  try {
    const added = await importFile(file, type);
    saveLocal();
    buildIndexes();
    // default selected date = latest
    if (IX.dates.length) S.selDate = IX.dates[IX.dates.length - 1];
    if (stat) { stat.className = 'stat ok'; stat.textContent = `✓ ${added} мөр орлоо`; }
    render();
    toast(`✓ ${file.name}: ${added} мөр импортлогдлоо`);
  } catch (e) {
    if (stat) { stat.className = 'stat err'; stat.textContent = '✕ ' + (e.message || 'алдаа'); }
    toast('⚠️ ' + (e.message || 'Импорт амжилтгүй'));
  }
}

/* ============================================================
   Publish modal
   ============================================================ */
function openPublish() {
  if (!DB.products.length) { toast('Эхлээд өгөгдөл оруулна уу'); return; }
  $('modal').classList.remove('wide');
  $('modal-title').textContent = 'Өгөгдлийг нийтлэх';
  $('modal-body').innerHTML = `
    <p style="color:var(--text-2);margin-bottom:16px">Та болон захирал хоёр <b>ижил мэдээллийг</b> харахын тулд өгөгдлөө дараах байдлаар нийтэлнэ:</p>
    <div class="pub-step"><div class="n">1</div><div class="t"><b>stock-data.json</b> файлыг татаж аваарай.</div></div>
    <div class="pub-step"><div class="n">2</div><div class="t">Татсан файлыг repo-гийн <code>dashboard/data/</code> фолдерт хийнэ (хуучныг дарж бичнэ).</div></div>
    <div class="pub-step"><div class="n">3</div><div class="t"><code>git add . &amp;&amp; git commit -m "data" &amp;&amp; git push</code> хийнэ.</div></div>
    <div class="pub-step"><div class="n">4</div><div class="t">Хэдэн минутын дараа энэ хаягийг та хоёр автоматаар шинэ өгөгдөлтэйгээр харна.</div></div>
    <div class="warn-box">⚠️ <b>Анхаар:</b> GitHub Pages хаяг нийтэд нээлттэй тул нийтэлсэн өгөгдлийг линктэй хэн ч харж болзошгүй. Хувийн байлгах шаардлагатай бол нэвтрэлт/нууц код нэмэх боломжтой — надад хэлээрэй.</div>
    <div style="display:flex;gap:10px;margin-top:18px;justify-content:flex-end">
      <button class="btn ghost" id="pub-cancel">Болих</button>
      <button class="btn" id="pub-dl">⬇️ stock-data.json татах</button>
    </div>`;
  showOverlay();
  $('pub-cancel').onclick = closeOverlay;
  $('pub-dl').onclick = () => {
    DB.updatedAt = ymd(new Date());
    downloadFile('stock-data.json', JSON.stringify(DB), 'application/json');
    toast('✓ Татагдлаа. dashboard/data/ дотор хийгээд push хийнэ үү.');
  };
}

function downloadFile(name, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = el('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 100);
}

/* ============================================================
   Overlay / toast
   ============================================================ */
function showOverlay() { $('overlay').hidden = false; document.body.style.overflow = 'hidden'; }
function closeOverlay() {
  $('overlay').hidden = true; document.body.style.overflow = '';
  if (pdChart) { pdChart.destroy(); pdChart = null; }
  $('modal').classList.remove('wide');
}
let toastTimer = null;
function toast(msg) {
  const t = $('toast'); t.hidden = false; t.textContent = msg;
  requestAnimationFrame(() => t.classList.add('show'));
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.hidden = true, 250); }, 3200);
}

/* ============================================================
   Events
   ============================================================ */
function wire() {
  $('f-date').onchange = (e) => { S.selDate = e.target.value; render(); };
  $('f-dept').onchange = (e) => { S.dept = e.target.value; render(); };
  $('f-class').onchange = (e) => { S.cls = e.target.value; render(); };
  $('f-status').onchange = (e) => { S.status = e.target.value; render(); };
  $('f-status2').onchange = (e) => { S.status2 = e.target.value; render(); };
  let searchTimer = null;
  $('f-search').oninput = (e) => { S.search = e.target.value; clearTimeout(searchTimer); searchTimer = setTimeout(render, 180); };
  $('f-reset').onclick = () => { S.dept = S.cls = S.status = S.status2 = S.search = ''; S.lowOnly = false; if (IX.dates.length) S.selDate = IX.dates[IX.dates.length - 1]; render(); };

  $('hist-scope').onchange = (e) => { S.histScope = e.target.value; renderHistory(); };
  $('hist-range').onchange = (e) => { S.histRange = +e.target.value; renderHistory(); };
  $('hist-metric').querySelectorAll('.seg-btn').forEach(b => b.onclick = () => { S.histMetric = b.dataset.v; renderHistScope(); renderHistory(); });
  $('matrix-metric').querySelectorAll('.seg-btn').forEach(b => b.onclick = () => {
    S.matrixMetric = b.dataset.v;
    $('matrix-metric').querySelectorAll('.seg-btn').forEach(x => x.classList.toggle('active', x.dataset.v === b.dataset.v));
    renderMatrix();
  });

  document.querySelectorAll('#prod-table th[data-sort]').forEach(th => th.onclick = () => {
    const k = th.dataset.sort;
    if (S.sortKey === k) S.sortDir *= -1; else { S.sortKey = k; S.sortDir = (k === 'name' || k === 'code' || k === 'department' || k === 'status' || k === 'status2' || k === 'cls') ? 1 : -1; }
    renderTable();
  });

  $('btn-import').onclick = openImport;
  $('btn-publish').onclick = openPublish;
  $('es-import').onclick = openImport;
  $('es-sample').onclick = () => {
    DB = buildSample(); saveLocal(); S.source = 'sample';
    buildIndexes(); S.selDate = IX.dates[IX.dates.length - 1];
    render(); toast('✨ Жишээ өгөгдөл ачаалагдлаа');
  };

  $('modal-close').onclick = closeOverlay;
  $('overlay').onclick = (e) => { if (e.target === $('overlay')) closeOverlay(); };
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('overlay').hidden) closeOverlay(); });
}

/* ============================================================
   Init
   ============================================================ */
async function init() {
  wire();
  // Demo mode: open with ?demo to explore the dashboard on sample data (does not touch real data)
  if (/(\?|&)demo\b/.test(location.search) || location.hash.includes('demo')) {
    DB = buildSample(); S.source = 'sample';
    buildIndexes(); S.selDate = IX.dates[IX.dates.length - 1];
    render(); return;
  }
  const local = loadLocal();
  const pub = await fetchPublished();
  let chosen = null;
  if (local && Array.isArray(local.products) && local.products.length) { chosen = local; S.source = 'local'; }
  // Plaintext published wins when at least as new as the local draft (so re-publishing always shows).
  if (pub && pub.kind === 'plain' && pub.db.products.length && (!chosen || (pub.db.updatedAt || '') >= ((local && local.updatedAt) || ''))) {
    chosen = pub.db; S.source = 'published';
  }
  // Encrypted published: decrypt (passcode gate) when there is no local draft to show.
  if (!chosen && pub && pub.kind === 'enc') {
    const db = await loadEncrypted(pub.buf);
    if (db && db.products && db.products.length) { chosen = db; S.source = 'published'; }
  }
  if (chosen) DB = Object.assign(emptyDB(), chosen);
  buildIndexes();
  if (IX.dates.length) S.selDate = IX.dates[IX.dates.length - 1];
  render();
}
init();
