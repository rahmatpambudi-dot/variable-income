"""
extract_data.py
Ekstrak data OT dari Google Sheets (Overtime Nasional 2026)
dan data Insentif dari Google Sheets (Insentif 2026),
join per NIK per bulan, update HTML dashboard.

Dijalankan oleh GitHub Actions. Butuh env vars:
  GDRIVE_CREDENTIALS : JSON service account key (dari GitHub Secrets)
  SHEET_ID_OT        : ID spreadsheet Overtime Nasional 2026
  SHEET_ID_INSENTIF  : ID spreadsheet Insentif 2026
"""

import os, json, re, time
from collections import defaultdict
from datetime import datetime, timezone, timedelta

import gspread
from google.oauth2.service_account import Credentials

# ── Config ────────────────────────────────────────────────────────────────────

OT_SHEET_TAB = 'Raw Data'

# NIK yang difilter (dummy/test accounts)
DUMMY_NIKS = {'999999', '0', ''}
DUMMY_NAME_KEYWORDS = ['DUMMY', 'TEST', 'DUMMY CUSTOMER', 'DUMMY CUSTOME']

# Mapping Location Name (OT) → site key untuk join insentif
# Hanya 3 site ini yang punya insentif
INSENTIF_SITE_MAPPING = {
    'JBBK': {
        'loc_contains': ['JABABEKA', 'JBBK'],
        'bu': ['HCI'],
    },
    'CKP': {
        'loc_contains': ['CIKUPA'],
        'bu': ['HCI'],
    },
    'SDA': {
        'loc_contains': ['SIDOARJO', 'SURABAYA', 'JUANDA'],
        'bu': [],  # semua BU
    },
}

# Tab insentif per site
INSENTIF_SITES = [
    'JBBK', 'CKP', 'SDA',
    'Hub Bogor', 'Hub Tangerang', 'Hub Utara', 'Hub Bandung',
    'Hub Yogya', 'Hub Semarang', 'Hub Lampung', 'Hub Palembang', 'Hub Kediri'
]

MONTH_ORDER = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December'
]

MONTH_ID = ['','Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des']

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
]

HTML_PATH = 'dashboard_variable_income.html'

# ── Auth ──────────────────────────────────────────────────────────────────────

def get_gc():
    creds_json = os.environ['GDRIVE_CREDENTIALS']
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# ── Helpers ───────────────────────────────────────────────────────────────────

def col_idx(headers, name):
    for i, h in enumerate(headers):
        if str(h).strip().lower() == name.lower():
            return i
    return -1

def to_num(v):
    if v in (None, '', 'None'): return 0.0
    try:
        return float(str(v).replace(',', '').replace(' ', ''))
    except:
        return 0.0

def normalize_name(name):
    n = str(name).strip().upper()
    n = re.sub(r'[.\s]+$', '', n)
    n = re.sub(r'\s+', ' ', n)
    return n

def is_dummy(nik, name):
    if str(nik).strip() in DUMMY_NIKS:
        return True
    name_up = str(name).strip().upper()
    for kw in DUMMY_NAME_KEYWORDS:
        if kw in name_up:
            return True
    return False

def insentif_site_for_location(loc_name, bu):
    """Return insentif site key (JBBK/CKP/SDA) jika cocok, else None"""
    loc_upper = loc_name.upper()
    bu_upper = bu.upper()
    for site, cfg in INSENTIF_SITE_MAPPING.items():
        loc_match = any(k in loc_upper for k in cfg['loc_contains'])
        bu_match = (len(cfg['bu']) == 0) or (bu_upper in [b.upper() for b in cfg['bu']])
        if loc_match and bu_match:
            return site
    return None

def clean_location_name(loc):
    """Bersihkan nama lokasi untuk dijadikan site label di dashboard"""
    loc = str(loc).strip()
    # Hapus prefix DC / prefix umum
    loc = re.sub(r'^DC\s+', '', loc, flags=re.IGNORECASE).strip()
    return loc if loc else 'UNKNOWN'

# ── Extract OT ────────────────────────────────────────────────────────────────

def extract_ot(wb_ot):
    print('\n📥 Extracting OT data...')
    ws = wb_ot.worksheet(OT_SHEET_TAB)
    print('  Fetching all rows...')
    all_rows = ws.get_all_values()
    if not all_rows:
        print('  [ERROR] Sheet kosong!')
        return {}

    # Header ada di row ke-2 (index 1), row ke-1 adalah judul/kosong
    headers = all_rows[1]
    total_cols = len(headers)
    print(f'  {total_cols} kolom, {len(all_rows)-1} baris data')

    # Header berulang — cari dari KANAN untuk kolom summary, dari KIRI untuk identifier
    def col_first(name):
        for i, h in enumerate(headers):
            if str(h).strip().lower() == name.strip().lower():
                return i
        return -1

    def col_last(name):
        for i in range(len(headers)-1, -1, -1):
            if str(headers[i]).strip().lower() == name.strip().lower():
                return i
        return -1

    def col_contains_last(keyword):
        for i in range(len(headers)-1, -1, -1):
            if keyword.lower() in str(headers[i]).strip().lower():
                return i
        return -1

    ci = {
        'nik'      : col_first('Employee ID'),
        'name'     : col_first('Employee Name'),
        'month'    : col_last('Month'),
        'hours'    : col_contains_last('OT Hour Paid'),
        'idr'      : col_contains_last('OT (IDR)'),
        'location' : col_last('Location Name'),
        'bu'       : col_last('BU'),
        'site_cat' : col_last('Site Category'),
    }
    # Verifikasi nama header aktual
    for k, v in ci.items():
        tag = repr(headers[v]) if v >= 0 else 'NOT FOUND'
        print(f'  [{k}] col {v} = {tag}')

    ot_data = {}
    skipped_dummy = 0
    skipped_nomonth = 0

    for row in all_rows[2:]:
        def g(c): return row[c].strip() if 0 <= c < len(row) else ''

        nik      = g(ci['nik'])
        name     = g(ci['name'])
        month    = g(ci['month'])
        location = g(ci['location'])
        bu       = g(ci['bu'])
        site_cat = g(ci['site_cat'])
        hours    = to_num(g(ci['hours']))
        idr      = to_num(g(ci['idr']))

        # Filter dummy
        if is_dummy(nik, name):
            skipped_dummy += 1
            continue

        # Normalize month — bisa angka (1-12) atau nama English
        MONTH_NUM_MAP = {
            '1':'January','2':'February','3':'March','4':'April',
            '5':'May','6':'June','7':'July','8':'August',
            '9':'September','10':'October','11':'November','12':'December'
        }
        if month not in MONTH_ORDER:
            month = MONTH_NUM_MAP.get(str(month).strip(), '')
        if not month:
            skipped_nomonth += 1
            continue

        if not nik:
            continue

        # Tentukan insentif site (JBBK/CKP/SDA) atau None
        ins_site = insentif_site_for_location(location, bu)

        # Site label untuk display: pakai insentif site key kalau ada, else location name
        display_site = ins_site if ins_site else location

        if nik not in ot_data:
            ot_data[nik] = {
                'name'      : name,
                'name_norm' : normalize_name(name),
                'site'      : display_site,
                'ins_site'  : ins_site,  # None jika bukan JBBK/CKP/SDA
                'location'  : location,
                'bu'        : bu,
                'site_cat'  : site_cat,
                'months'    : defaultdict(lambda: {'hours': 0.0, 'idr': 0.0}),
            }

        ot_data[nik]['months'][month]['hours'] += hours
        ot_data[nik]['months'][month]['idr']   += idr

    print(f'  ✅ OT: {len(ot_data)} drivers')
    print(f'     Skipped dummy: {skipped_dummy}, skip no-month: {skipped_nomonth}')
    return ot_data

# ── Extract Insentif ──────────────────────────────────────────────────────────

def extract_insentif(wb_ins):
    print('\n📥 Extracting Insentif data...')
    ins_data = {}

    for i, site in enumerate(INSENTIF_SITES):
        try:
            ws = wb_ins.worksheet(site)
            all_rows = ws.get_all_values()
            if not all_rows:
                print(f'  [SKIP] {site} kosong')
                continue

            headers = all_rows[0]
            ci = {
                'nik'   : col_idx(headers, 'NIK1'),
                'name'  : col_idx(headers, 'driver'),
                'month' : col_idx(headers, 'Month Rev'),
                'ins'   : col_idx(headers, 'Insentif per MPP'),
            }

            if ci['nik'] < 0 or ci['ins'] < 0:
                print(f'  [WARN] {site}: kolom tidak ditemukan')
                continue

            count = 0
            for row in all_rows[2:]:
                def g(c): return row[c].strip() if 0 <= c < len(row) else ''

                nik   = g(ci['nik'])
                name  = g(ci['name'])
                month = g(ci['month'])
                ins   = to_num(g(ci['ins']))

                if is_dummy(nik, name):
                    continue

                if not nik or month not in MONTH_ORDER or ins <= 0:
                    continue

                if nik not in ins_data:
                    ins_data[nik] = {
                        'name'  : name,
                        'site'  : site,
                        'months': defaultdict(float),
                    }

                ins_data[nik]['months'][month] += ins
                count += 1

            print(f'  ✅ {site}: {count} rows')
            if i < len(INSENTIF_SITES) - 1:
                time.sleep(5)

        except gspread.exceptions.WorksheetNotFound:
            print(f'  [MISS] Sheet "{site}" tidak ditemukan')
        except Exception as e:
            print(f'  [ERROR] {site}: {e}')

    print(f'  Total insentif drivers: {len(ins_data)}')
    return ins_data

# ── Detect Months ─────────────────────────────────────────────────────────────

def detect_months(ot_data, ins_data):
    month_set = set()
    for d in ot_data.values():
        month_set.update(d['months'].keys())
    for d in ins_data.values():
        month_set.update(d['months'].keys())
    months = sorted(month_set, key=lambda m: MONTH_ORDER.index(m))
    print(f'\n📅 Months: {months}')
    return months

# ── Join & Build ──────────────────────────────────────────────────────────────

def build_driver_data(ot_data, ins_data, months):
    all_niks = set(ot_data.keys()) | set(ins_data.keys())
    drivers = []

    for nik in all_niks:
        ot  = ot_data.get(nik)
        ins = ins_data.get(nik)

        if ot:
            name     = ot['name']
            site     = ot['site']
            location = ot['location']
            bu       = ot['bu']
            site_cat = ot['site_cat']
        else:
            name     = ins['name'] if ins else nik
            site     = ins['site'] if ins else '-'
            location = '-'
            bu       = '-'
            site_cat = '-'

        monthly = {}
        total_ot_hours = 0.0
        total_ot_idr   = 0.0
        total_ins      = 0.0

        for m in months:
            ot_h   = ot['months'][m]['hours'] if ot and m in ot['months'] else 0.0
            ot_idr = ot['months'][m]['idr']   if ot and m in ot['months'] else 0.0
            ins_m  = ins['months'][m]          if ins and m in ins['months'] else 0.0

            monthly[m] = {
                'ot_hours': round(ot_h, 2),
                'ot_idr'  : round(ot_idr),
                'insentif': round(ins_m),
                'total'   : round(ot_idr + ins_m),
            }
            total_ot_hours += ot_h
            total_ot_idr   += ot_idr
            total_ins      += ins_m

        drivers.append({
            'nik'           : nik,
            'name'          : name,
            'site'          : site,
            'location'      : location,
            'bu'            : bu,
            'site_cat'      : site_cat,
            'has_ot'        : ot is not None,
            'has_ins'       : ins is not None,
            'monthly'       : monthly,
            'total_ot_hours': round(total_ot_hours, 2),
            'total_ot_idr'  : round(total_ot_idr),
            'total_ins'     : round(total_ins),
            'grand_total'   : round(total_ot_idr + total_ins),
        })

    drivers.sort(key=lambda x: -x['grand_total'])

    ot_only  = sum(1 for d in drivers if d['has_ot'] and not d['has_ins'])
    ins_only = sum(1 for d in drivers if not d['has_ot'] and d['has_ins'])
    both     = sum(1 for d in drivers if d['has_ot'] and d['has_ins'])
    print(f'\n✅ Total drivers: {len(drivers)}')
    print(f'   Both: {both}, OT only: {ot_only}, Ins only: {ins_only}')
    return drivers

# ── HTML Update ───────────────────────────────────────────────────────────────

def jd(obj):
    return json.dumps(obj, separators=(',', ':'), ensure_ascii=False)

def update_html(drivers, months, last_data_date):
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    html = re.sub(r'Update: [\d\w\s]+2026', f'Update: {last_data_date}', html)
    html = re.sub(r'const MONTHS=\[[^\]]*\]', f'const MONTHS={jd(months)}', html)
    html = re.sub(r'const DRIVERS=\[.*?\];', f'const DRIVERS={jd(drivers)};', html, flags=re.DOTALL)

    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    wib = timezone(timedelta(hours=7))
    now = datetime.now(wib).strftime('%d %b %Y %H:%M WIB')
    print(f'\n✅ HTML updated [{now}]')

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print('=== Variable Income Dashboard — Auto Update ===\n')
    gc = get_gc()

    wb_ot  = gc.open_by_key(os.environ['SHEET_ID_OT'])
    wb_ins = gc.open_by_key(os.environ['SHEET_ID_INSENTIF'])

    ot_data  = extract_ot(wb_ot)
    ins_data = extract_insentif(wb_ins)
    months   = detect_months(ot_data, ins_data)
    drivers  = build_driver_data(ot_data, ins_data, months)

    wib = timezone(timedelta(hours=7))
    today = datetime.now(wib)
    last_month = months[-1] if months else today.strftime('%B')
    last_month_idx = MONTH_ORDER.index(last_month) + 1
    last_data_date = f"{today.day} {MONTH_ID[last_month_idx]} {today.year}"

    update_html(drivers, months, last_data_date)

if __name__ == '__main__':
    main()
