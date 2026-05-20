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

# Sheet tab name untuk RAW DATA OT
OT_SHEET_TAB = 'RAW DATA'

# Mapping site insentif → filter kondisi di OT (Location Name contains + BU filter)
# Format: insentif_site → {'loc_contains': [...], 'bu': [...]}
SITE_MAPPING = {
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
    """Normalize nama untuk matching: uppercase, trim, hapus titik/karakter aneh di akhir"""
    n = str(name).strip().upper()
    n = re.sub(r'[.\s]+$', '', n)  # hapus titik/spasi di akhir
    n = re.sub(r'\s+', ' ', n)     # normalisasi spasi ganda
    return n

def site_for_location(loc_name, bu):
    """Tentukan site insentif dari Location Name dan BU"""
    loc_upper = loc_name.upper()
    bu_upper = bu.upper()
    for site, cfg in SITE_MAPPING.items():
        loc_match = any(k in loc_upper for k in cfg['loc_contains'])
        bu_match = (len(cfg['bu']) == 0) or (bu_upper in [b.upper() for b in cfg['bu']])
        if loc_match and bu_match:
            return site
    return None

# ── Extract OT ────────────────────────────────────────────────────────────────

def extract_ot(wb_ot):
    """
    Baca sheet RAW DATA OT.
    Return: dict NIK → {
        'name': str,
        'site': str,  # JBBK/CKP/SDA
        'location': str,
        'bu': str,
        'months': { 'January': {'hours': float, 'idr': float}, ... }
    }
    """
    print('\n📥 Extracting OT data...')
    ws = wb_ot.worksheet(OT_SHEET_TAB)

    # Fetch in batches to avoid timeout on 50k rows
    print('  Fetching all rows (may take a while)...')
    all_rows = ws.get_all_values()
    if not all_rows:
        print('  [ERROR] Sheet kosong!')
        return {}

    headers = all_rows[0]
    print(f'  Headers ({len(headers)} cols): {headers[:5]}...{headers[-5:]}')

    ci = {
        'nik'      : col_idx(headers, 'Employee ID'),
        'name'     : col_idx(headers, 'Employee Name'),
        'month'    : col_idx(headers, 'Month'),
        'hours'    : col_idx(headers, 'Total OT Hour Paid'),
        'idr'      : col_idx(headers, 'OT (IDR)'),
        'location' : col_idx(headers, 'Location Name'),
        'bu'       : col_idx(headers, 'BU'),
        'site_cat' : col_idx(headers, 'Site Category'),
    }
    print(f'  Column indices: {ci}')

    ot_data = {}
    skipped = 0

    for row in all_rows[1:]:
        def g(c): return row[c].strip() if 0 <= c < len(row) else ''

        nik      = g(ci['nik'])
        name     = g(ci['name'])
        month    = g(ci['month'])
        location = g(ci['location'])
        bu       = g(ci['bu'])
        hours    = to_num(g(ci['hours']))
        idr      = to_num(g(ci['idr']))

        if not nik or not name or month not in MONTH_ORDER:
            skipped += 1
            continue

        site = site_for_location(location, bu)
        if not site:
            skipped += 1
            continue

        if nik not in ot_data:
            ot_data[nik] = {
                'name'    : normalize_name(name),
                'name_raw': name,
                'site'    : site,
                'location': location,
                'bu'      : bu,
                'months'  : defaultdict(lambda: {'hours': 0.0, 'idr': 0.0}),
            }

        ot_data[nik]['months'][month]['hours'] += hours
        ot_data[nik]['months'][month]['idr']   += idr

    print(f'  ✅ OT: {len(ot_data)} drivers ditemukan, {skipped} rows skipped')
    return ot_data

# ── Extract Insentif ──────────────────────────────────────────────────────────

def extract_insentif(wb_ins):
    """
    Baca semua tab insentif.
    Return: dict NIK → {
        'name': str,
        'site': str,
        'months': { 'January': float, ... }
    }
    """
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
                print(f'  [WARN] {site}: kolom NIK1 atau Insentif per MPP tidak ditemukan')
                continue

            count = 0
            for row in all_rows[1:]:
                def g(c): return row[c].strip() if 0 <= c < len(row) else ''

                nik   = g(ci['nik'])
                name  = g(ci['name'])
                month = g(ci['month'])
                ins   = to_num(g(ci['ins']))

                if not nik or month not in MONTH_ORDER or ins <= 0:
                    continue

                if nik not in ins_data:
                    ins_data[nik] = {
                        'name'   : normalize_name(name),
                        'site'   : site,
                        'months' : defaultdict(float),
                    }

                ins_data[nik]['months'][month] += ins
                count += 1

            print(f'  ✅ {site}: {count} rows')
            if i < len(INSENTIF_SITES) - 1:
                time.sleep(5)

        except gspread.exceptions.WorksheetNotFound:
            print(f'  [MISS] {site}')
        except Exception as e:
            print(f'  [ERROR] {site}: {e}')

    print(f'  Total insentif drivers: {len(ins_data)}')
    return ins_data

# ── Join & Build ──────────────────────────────────────────────────────────────

def build_driver_data(ot_data, ins_data, months):
    """
    Join OT + Insentif per NIK.
    Return list of driver dicts, sorted by total IDR desc.
    """
    all_niks = set(ot_data.keys()) | set(ins_data.keys())
    drivers = []

    for nik in all_niks:
        ot  = ot_data.get(nik)
        ins = ins_data.get(nik)

        # Tentukan nama & site
        if ot:
            name = ot['name_raw']
            site = ot['site']
            location = ot['location']
            bu = ot['bu']
        else:
            name = ins['name'] if ins else nik
            site = ins['site'] if ins else '-'
            location = '-'
            bu = '-'

        # Build monthly breakdown
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
            'nik'          : nik,
            'name'         : name,
            'site'         : site,
            'location'     : location,
            'bu'           : bu,
            'has_ot'       : ot is not None,
            'has_ins'      : ins is not None,
            'monthly'      : monthly,
            'total_ot_hours': round(total_ot_hours, 2),
            'total_ot_idr' : round(total_ot_idr),
            'total_ins'    : round(total_ins),
            'grand_total'  : round(total_ot_idr + total_ins),
        })

    drivers.sort(key=lambda x: -x['grand_total'])
    print(f'\n✅ Total drivers (join): {len(drivers)}')
    print(f'   - OT only  : {sum(1 for d in drivers if d["has_ot"] and not d["has_ins"])}')
    print(f'   - Ins only : {sum(1 for d in drivers if not d["has_ot"] and d["has_ins"])}')
    print(f'   - Both     : {sum(1 for d in drivers if d["has_ot"] and d["has_ins"])}')
    return drivers

# ── Detect Months ─────────────────────────────────────────────────────────────

def detect_months(ot_data, ins_data):
    month_set = set()
    for d in ot_data.values():
        month_set.update(d['months'].keys())
    for d in ins_data.values():
        month_set.update(d['months'].keys())
    months = sorted(month_set, key=lambda m: MONTH_ORDER.index(m))
    print(f'\n📅 Months detected: {months}')
    return months

# ── HTML Update ───────────────────────────────────────────────────────────────

def jd(obj):
    return json.dumps(obj, separators=(',', ':'), ensure_ascii=False)

def update_html(drivers, months, last_data_date):
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    html = re.sub(r'Update: \d+ \w+ \d{4}', f'Update: {last_data_date}', html)
    html = re.sub(r'const MONTHS=\[[^\]]*\]', f'const MONTHS={jd(months)}', html)
    html = re.sub(r'const DRIVERS=\[.*?\];', f'const DRIVERS={jd(drivers)};', html, flags=re.DOTALL)

    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    wib = timezone(timedelta(hours=7))
    now = datetime.now(wib).strftime('%d %b %Y %H:%M WIB')
    print(f'\n✅ HTML updated: {HTML_PATH} [{now}]')

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

    # Last data date
    wib = timezone(timedelta(hours=7))
    today = datetime.now(wib)
    last_month = months[-1] if months else today.strftime('%B')
    last_month_idx = MONTH_ORDER.index(last_month) + 1
    last_data_date = f"{today.day} {MONTH_ID[last_month_idx]} {today.year}"

    print('\n✏️  Updating HTML...')
    update_html(drivers, months, last_data_date)

if __name__ == '__main__':
    main()
