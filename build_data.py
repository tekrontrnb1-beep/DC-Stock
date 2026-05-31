# -*- coding: utf-8 -*-
"""
DC-Stock — ETL
Reads the raw export files in ./import and writes ./data/stock-data.json (local,
gitignored) + ./data/stock-data.enc (encrypted, published) in the dashboard schema.

Re-run whenever you refresh the files in ./import:
    python build_data.py     (or double-click refresh.bat)

Source files expected in ./import (filenames are matched by pattern):
  - 상품마스터*.xlsx ........... product master: нэр, Department, Өртөг үнэ, Худалдах үнэ
  - 9001 master*.xlsx .......... V9001 (DC) goods registry: GoodsCd + reOrdAppDt (төлөв)
  - ABCXYZ*.xlsx ............... ABC / XYZ / class per barcode ('raw' sheet)
  - Үлдэгдэл MM.DD.xlsx ........ daily stock balance (date from file name; year below)
  - Store order*.csv ........... daily store orders + outbound (qty + amount)
  - Sales*.csv ................. retail sales total for the period (qty per barcode)

Universe = ТОЛЬКО products that have a stock balance (үлдэгдэлтэй).
Inventory value = qty * Өртөг үнэ (cost).
"""
import os, re, sys, csv, glob, json, struct
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
IMP = os.path.join(BASE, 'import')
OUT = os.path.join(BASE, 'data', 'stock-data.json')
ENC = os.path.join(BASE, 'data', 'stock-data.enc')
YEAR = 2026                 # year for the Үлдэгдэл MM.DD file names
UPDATED = '%d-05-31' % YEAR # stamp shown in the dashboard

def find(pat):
    # snapshot files (master/9001/ABCXYZ): pick the LATEST by file name (newest date in name)
    g = sorted(glob.glob(os.path.join(IMP, pat)))
    return g[-1] if g else None

def find_all(*pats):
    out = []
    for p in pats:
        out += glob.glob(os.path.join(IMP, p))
    return sorted(set(out))

def bc(v):
    if v is None: return ''
    if isinstance(v, float): return str(int(v)) if v.is_integer() else str(v)
    s = str(v).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def numf(v):
    if v is None or v == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(',', '').replace("'", '').strip())
    except: return 0.0

def col_index(hdr, *names):
    H = {str(h).strip(): i for i, h in enumerate(hdr) if h is not None}
    for n in names:
        if n in H: return H[n]
    # case-insensitive fallback
    Hl = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h is not None}
    for n in names:
        if n.lower() in Hl: return Hl[n.lower()]
    return None

# ---------------- OLD MASTER (name / department / cost / sell price) ----------------
print('· Reading product master …')
master = {}
mp = find('상품마스터*') or find('*master*Улаанбаатар*') or find('*마스터*')
wb = openpyxl.load_workbook(mp, read_only=True, data_only=True); ws = wb.worksheets[0]
it = ws.iter_rows(values_only=True); hdr = list(next(it))
ci = dict(bar=col_index(hdr, 'Бар код'), name=col_index(hdr, 'Барааны нэр'),
          dep=col_index(hdr, 'Department нэр'), cat=col_index(hdr, 'Category нэр'),
          cost=col_index(hdr, 'Өртөг үнэ'), sell=col_index(hdr, 'Худалдах үнэ'))
for row in it:
    code = bc(row[ci['bar']]) if ci['bar'] is not None else ''
    if not code: continue
    master[code] = {
        'name': str(row[ci['name']]).strip() if ci['name'] is not None and row[ci['name']] is not None else '',
        'department': str(row[ci['dep']]).strip() if ci['dep'] is not None and row[ci['dep']] is not None else '',
        'category': str(row[ci['cat']]).strip() if ci['cat'] is not None and row[ci['cat']] is not None else '',
        'cost': numf(row[ci['cost']]) if ci['cost'] is not None else 0.0,
        'sell': numf(row[ci['sell']]) if ci['sell'] is not None else 0.0,
    }
wb.close(); print('  master items:', len(master))

# ---------------- 9001 MASTER (status = reOrdAppDt) ----------------
print('· Reading 9001 status master …')
status9001 = {}
sp = find('9001 master*') or find('*9001*')
if sp:
    wb = openpyxl.load_workbook(sp, read_only=True, data_only=True); ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True); hdr = list(next(it))
    cbar = col_index(hdr, 'GoodsCd', 'Бар код', 'Баркод')
    c1 = col_index(hdr, 'RegionGoodsStat')   # Нийлүүлэгчийн төлөв (Active/Non Active/Deleted)
    c2 = col_index(hdr, 'TradeOrdStatCd')    # Салбарын агуулах руу захих төлөв
    for row in it:
        code = bc(row[cbar]) if cbar is not None else ''
        if not code: continue
        s1 = str(row[c1]).strip() if c1 is not None and row[c1] is not None else ''
        s2 = str(row[c2]).strip() if c2 is not None and row[c2] is not None else ''
        status9001[code] = (s1, s2)
    wb.close()
print('  9001 status items:', len(status9001))

# ---------------- ABC / XYZ ----------------
print('· Reading ABC/XYZ …')
abcx = {}
ap = find('ABCXYZ*') or find('*ABC*')
wb = openpyxl.load_workbook(ap, read_only=True, data_only=True)
ws = wb['raw'] if 'raw' in wb.sheetnames else wb.worksheets[0]
cols = None
for row in ws.iter_rows(values_only=True):
    vals = [str(c).strip() if c is not None else '' for c in row]
    if cols is None:
        if 'Баркод' in vals and 'class' in vals:
            cols = {'bar': vals.index('Баркод'), 'cls': vals.index('class'),
                    'abc': vals.index('ABC') if 'ABC' in vals else None,
                    'xyz': vals.index('XYZ') if 'XYZ' in vals else None}
        continue
    bar = bc(row[cols['bar']])
    if not bar: continue
    abc = (str(row[cols['abc']]).strip().upper() if cols['abc'] is not None and row[cols['abc']] is not None else '')
    xyz = (str(row[cols['xyz']]).strip().upper() if cols['xyz'] is not None and row[cols['xyz']] is not None else '')
    cls = (str(row[cols['cls']]).strip().upper() if row[cols['cls']] is not None else '')
    if not re.match(r'^[ABC][XYZ]$', cls):
        cls = (abc + xyz) if re.match(r'^[ABC][XYZ]$', abc + xyz) else ''
    abcx[bar] = {'abc': abc, 'xyz': xyz, 'cls': cls}
wb.close(); print('  abc/xyz items:', len(abcx))

# ---------------- BALANCES (daily) ----------------
print('· Reading daily balances …')
balances = []; avg_out = {}; bal_codes = set(); bal_name = {}
for path in sorted(glob.glob(os.path.join(IMP, 'Үлдэгдэл*.xls*'))):
    m = re.search(r'(\d{1,2})[.\-_](\d{1,2})', os.path.basename(path))
    if not m: print('  ! skip (no date):', os.path.basename(path)); continue
    date = '%d-%02d-%02d' % (YEAR, int(m.group(1)), int(m.group(2)))
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.worksheets[0]
    it = ws.iter_rows(values_only=True); h = list(next(it))
    cbar = col_index(h, 'Бар код', 'No.', 'Баркод'); cqty = col_index(h, 'Үлдэгдэл', 'Inventory')
    cavg = col_index(h, 'Өдрийн дундаж гаралт'); cnm = col_index(h, 'Барааны нэр')
    if cbar is None or cqty is None: print('  ! skip (cols):', os.path.basename(path)); wb.close(); continue
    for row in it:
        code = bc(row[cbar])
        if not code: continue
        balances.append({'date': date, 'code': code, 'qty': int(round(numf(row[cqty])))})
        bal_codes.add(code)
        if cavg is not None: avg_out[code] = numf(row[cavg])
        if cnm is not None and code not in bal_name and row[cnm] is not None: bal_name[code] = str(row[cnm]).strip()
    wb.close()
bal_dates = sorted(set(b['date'] for b in balances))
print('  balance rows:', len(balances), ' dates:', len(bal_dates), bal_dates[:1], '→', bal_dates[-1:])

# ---------------- ORDERS + OUTBOUND (daily) ----------------
print('· Reading store orders / outbound …')
orders, sales = [], []
# Олон файл уншина — шинэ хугацааны файлуудыг нэг нэгээр нэмж болно (ORDER_DATE-ээр нэгтгэнэ).
order_files = find_all('Store order*.csv') or find_all('*outbound*.csv', '*order*.csv')
agg = {}; code_oq = {}; code_bq = {}
for op in order_files:
    with open(op, encoding='utf-8-sig', newline='') as f:
        r = csv.DictReader(f)
        for d in r:
            date = (d.get('ORDER_DATE') or '').strip()[:10]
            code = bc(d.get('ITEM_BAR_CODE'))
            if not date or not code: continue
            a = agg.get((date, code))
            if not a: a = {'oq': 0.0, 'oa': 0.0, 'bq': 0.0, 'ba': 0.0}; agg[(date, code)] = a
            oq = numf(d.get('TOTAL_ORDER_QTY')); bq = numf(d.get('CALCULATED_OUTBOUND_QTY'))
            a['oq'] += oq; a['oa'] += numf(d.get('TOTAL_ORDERED_AMOUNT'))
            a['bq'] += bq; a['ba'] += numf(d.get('TOTAL_OUTBOUND_AMOUNT'))
            code_oq[code] = code_oq.get(code, 0.0) + oq
            code_bq[code] = code_bq.get(code, 0.0) + bq
print('  order files:', [os.path.basename(f) for f in order_files])
order_codes = set()
for (date, code), a in agg.items():
    order_codes.add(code)
    if a['oq']: orders.append({'date': date, 'code': code, 'qty': int(round(a['oq'])), 'amount': int(round(a['oa']))})
    if a['bq']: sales.append({'date': date, 'code': code, 'qty': int(round(a['bq'])), 'amount': int(round(a['ba']))})
ord_dates = sorted(set(o['date'] for o in orders))
n_order_days = max(1, len(ord_dates))
print('  order rows:', len(orders), ' outbound rows:', len(sales), ' order days:', n_order_days, ord_dates[:1], '→', ord_dates[-1:])

# ---------------- RETAIL SALES total (qty per barcode) — олон файл нэмж болно ----------------
retail = {}
sales_files = find_all('Sales*.csv')
for rp in sales_files:
    with open(rp, encoding='utf-8-sig', newline='') as f:
        r = csv.reader(f); next(r, None)
        for row in r:
            if len(row) < 3: continue
            retail[bc(row[0])] = retail.get(bc(row[0]), 0.0) + numf(row[2])
retail_period = ''
if sales_files:
    pm = re.search(r'(\d{1,2}\.\d{1,2})\s*-\s*(\d{1,2}\.\d{1,2})', ' '.join(os.path.basename(f) for f in sales_files))
    if pm: retail_period = pm.group(1) + '–' + pm.group(2)
    print('· retail sales files:', [os.path.basename(f) for f in sales_files], 'items:', len(retail), 'period:', retail_period)

# ---------------- BUILD PRODUCTS (universe = balance-having AND registered in V9001) ----------------
# Үлдэгдэл файл олон агуулахын бараа агуулдаг тул V9001-д (9001 master) бүртгэлтэйг нь л үлдээнэ.
universe = set(bal_codes)
if status9001:
    before = len(universe)
    universe &= set(status9001.keys())
    print('  universe: %d үлдэгдэлтэй → %d (V9001-д бүртгэлтэй); %d хасагдсан (өөр агуулах)' % (before, len(universe), before - len(universe)))
mm = ma = ms = 0
products = []
for code in sorted(universe):
    m = master.get(code, {}); ax = abcx.get(code, {})
    if code in master: mm += 1
    if code in abcx: ma += 1
    if code in status9001: ms += 1
    avgS = round(avg_out.get(code, 0.0), 3)              # avg daily outbound (борлуулалт)
    avgO = round(code_oq.get(code, 0.0) / n_order_days, 3)  # avg daily branch order (захиалга)
    products.append({
        'code': code,
        'name': m.get('name') or bal_name.get(code) or code,
        'department': m.get('department') or 'Бусад',
        'category': m.get('category') or '',
        'abc': ax.get('abc', ''), 'xyz': ax.get('xyz', ''), 'cls': ax.get('cls', ''),
        'status': status9001.get(code, ('', ''))[0],     # Нийлүүлэгчийн төлөв (RegionGoodsStat)
        'status2': status9001.get(code, ('', ''))[1],    # Захих төлөв (TradeOrdStatCd)
        'price': round(m.get('cost', 0.0), 2),           # COST — used for inventory value
        'sellPrice': round(m.get('sell', 0.0), 2),
        'unit': 'ш',
        'minStock': int(round(avgS * 3)) if avgS else 0,
        'avgSales': avgS, 'avgOrders': avgO,
        'salesPeriodQty': int(round(retail.get(code, 0))) if code in retail else 0,
    })

balances = [b for b in balances if b['code'] in universe]
sales = [s for s in sales if s['code'] in universe]
orders = [o for o in orders if o['code'] in universe]

print('\n=== SUMMARY ===')
print('  products (үлдэгдэлтэй):', len(products))
print('   - matched master   :', mm, '(%.0f%%)' % (100 * mm / max(1, len(products))))
print('   - matched abc/xyz  :', ma, '(%.0f%%)' % (100 * ma / max(1, len(products))))
print('   - matched 9001 stat:', ms, '(%.0f%%)' % (100 * ms / max(1, len(products))))
print('  balances:', len(balances), ' sales(outbound):', len(sales), ' orders:', len(orders))

DB = {
    'products': products, 'balances': balances, 'sales': [], 'orders': orders,
    'updatedAt': UPDATED,
    'meta': {'source': 'DC-Stock ETL', 'value': 'cost (Өртөг үнэ)',
             'sales': 'retail (Sales.csv, period total per item)', 'universe': 'V9001 + balance',
             'retailPeriod': retail_period, 'balanceDates': bal_dates},
}

# ---------------- WRITE plaintext + encrypted ----------------
os.makedirs(os.path.dirname(OUT), exist_ok=True)
plaintext = json.dumps(DB, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
with open(OUT, 'wb') as f: f.write(plaintext)
print('\n✓ wrote', OUT, '(%.2f MB) — plaintext, local only (gitignored)' % (len(plaintext) / 1e6))

ITER = 200000
def encrypt_to(path, data, password):
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt = os.urandom(16); iv = os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(password.encode('utf-8'))
    ct = AESGCM(key).encrypt(iv, data, None)
    with open(path, 'wb') as f:
        f.write(b'DCS1'); f.write(struct.pack('>I', ITER)); f.write(salt); f.write(iv); f.write(ct)

pw_path = os.path.join(IMP, 'passcode.txt')
password = open(pw_path, encoding='utf-8').read().strip() if os.path.exists(pw_path) else os.environ.get('DC_STOCK_PASS', '').strip()
if password:
    encrypt_to(ENC, plaintext, password)
    print('🔒 wrote', ENC, '(%.2f MB) — encrypted, commit THIS for publishing' % (os.path.getsize(ENC) / 1e6))
else:
    if os.path.exists(ENC): os.remove(ENC)
    print('… no passcode → skipped .enc. Put your passcode in import/passcode.txt to publish.')
