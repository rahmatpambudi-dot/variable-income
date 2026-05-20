"""
extract_data.py
Base data = MPP dari insentif (per site tab)
Join OT by NIK
Output: per driver → NIK, Nama, Site, OT, Insentif, Total
"""

import os, json, re, time
from collections import defaultdict
from datetime import datetime, timezone, timedelta

import gspread
from google.oauth2.service_account import Credentials

# ── Config ────────────────────────────────────────────────────────────────────

OT_SHEET_TAB = 'Raw Data'

DUMMY_NIKS = {'999999', '0', ''}
DUMMY_NAME_KEYWORDS = ['DUMMY', 'TEST']

INSENTIF_SITES = [
    'JBBK', 'CKP', 'SDA',
    'Hub Bogor', 'Hub Tangerang', 'Hub Utara', 'Hub Bandung',
    'Hub Yogya', 'Hub Semarang', 'Hub Lampung', 'Hub Palembang', 'Hub Kediri'
]

MONTH_ORDER = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December'
]
MONTH_NUM_MAP = {
    '1':'January','2':'February','3':'March','4':'April',
    '5':'May','6':'June','7':'July','8':'August',
    '9':'September','10':'October','11':'November','12':'December'
}
MONTH_ID = ['','Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des']

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
]

HTML_PATH = 'dashboard_variable_income.html'

# ── Auth ──────────────────────────────────────────────────────────────────────

def get_gc():
    creds_dict = json.loads(os.environ['GDRIVE_CREDENTIALS'])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# ── Helpers ───────────────────────────────────────────────────────────────────

def to_num(v):
    if v in (None, '', 'None'): return 0.0
    try:
        return float(str(v).replace(',', '').replace(' ', ''))
    except:
        return 0.0

def is_dummy(nik, name):
    if str(nik).strip() in DUMMY_NIKS: return True
    name_up = str(name).strip().upper()
    return any(kw in name_up for kw in DUMMY_NAME_KEYWORDS)

def normalize_month(m):
    m = str(m).strip()
    if m in MONTH_ORDER: return m
    return MONTH_NUM_MAP.get(m, '')

# ── Extract Insentif (BASE MPP) ───────────────────────────────────────────────

def extract_insentif(wb_ins):
    """
    Base data = semua driver dari sheet insentif per site.
    Return: dict NIK → {name, site, months: {month: insentif_idr}}
    """
    print('\n📥 Extracting Insentif (MPP base)...')
    mpp = {}

    for i, site in enumerate(INSENTIF_SITES):
        try:
            ws = wb_ins.worksheet(site)
            all_rows = ws.get_all_values()
            if len(all_rows) < 2:
                print(f'  [SKIP] {site} kosong')
                continue

            headers = all_rows[0]
            def ci(name):
                for j, h in enumerate(headers):
                    if str(h).strip().lower() == name.lower(): return j
                return -1

            c_nik   = ci('NIK1')
            c_name  = ci('driver')
            c_month = ci('Month Rev')
            c_ins   = ci('Insentif per MPP')

            if c_nik < 0:
                print(f'  [WARN] {site}: kolom NIK1 tidak ditemukan, headers: {headers[:5]}')
                continue

            count = 0
            for row in all_rows[1:]:
                def g(c): return row[c].strip() if 0 <= c < len(row) else ''
                nik   = g(c_nik)
                name  = g(c_name)
                month = normalize_month(g(c_month))
                ins   = to_num(g(c_ins))

                if is_dummy(nik, name) or not nik or not month:
                    continue

                if nik not in mpp:
                    mpp[nik] = {
                        'name'  : name,
                        'site'  : site,
                        'months': defaultdict(float),
                    }

                mpp[nik]['months'][month] += ins
                count += 1

            print(f'  ✅ {site}: {count} rows, {len([k for k,v in mpp.items() if v["site"]==site])} drivers')
            if i < len(INSENTIF_SITES) - 1:
                time.sleep(3)

        except gspread.exceptions.WorksheetNotFound:
            print(f'  [MISS] Sheet "{site}" tidak ditemukan')
        except Exception as e:
            print(f'  [ERROR] {site}: {e}')

    print(f'  Total MPP drivers: {len(mpp)}')
    return mpp

# ── Extract OT ────────────────────────────────────────────────────────────────

def extract_ot(wb_ot):
    """
    Baca RAW DATA OT.
    Return: dict NIK → {name, site, site_cat, location, bu, months: {month: {hours, idr}}}
    """
    print('\n📥 Extracting OT data...')
    ws = wb_ot.worksheet(OT_SHEET_TAB)
    print('  Fetching all rows...')
    all_rows = ws.get_all_values()
    if len(all_rows) < 2:
        print('  [ERROR] Sheet kosong!')
        return {}

    # Header di row ke-2 (index 1)
    headers = all_rows[1]
    print(f'  {len(headers)} kolom, {len(all_rows)-2} baris data')

    def col_first(name):
        for j, h in enumerate(headers):
            if str(h).strip().lower() == name.strip().lower(): return j
        return -1

    def col_last(name):
        for j in range(len(headers)-1, -1, -1):
            if str(headers[j]).strip().lower() == name.strip().lower(): return j
        return -1

    def col_contains_last(keyword):
        for j in range(len(headers)-1, -1, -1):
            if keyword.lower() in str(headers[j]).strip().lower(): return j
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
    print(f'  Column indices: {ci}')

    ot = {}
    skipped = 0

    for row in all_rows[2:]:
        def g(c): return row[c].strip() if 0 <= c < len(row) else ''

        nik      = g(ci['nik'])
        name     = g(ci['name'])
        location = g(ci['location'])
        bu       = g(ci['bu'])
        site_cat = g(ci['site_cat'])
        hours    = to_num(g(ci['hours']))
        idr      = to_num(g(ci['idr']))
        month    = normalize_month(g(ci['month']))

        if is_dummy(nik, name) or not nik or not month:
            skipped += 1
            continue

        if nik not in ot:
            ot[nik] = {
                'name'    : name,
                'site'    : location,   # display per location name
                'site_cat': site_cat,
                'location': location,
                'bu'      : bu,
                'months'  : defaultdict(lambda: {'hours': 0.0, 'idr': 0.0}),
            }

        ot[nik]['months'][month]['hours'] += hours
        ot[nik]['months'][month]['idr']   += idr

    print(f'  ✅ OT: {len(ot)} drivers, {skipped} rows skipped')
    return ot

# ── Detect Months ─────────────────────────────────────────────────────────────

def detect_months(mpp, ot):
    month_set = set()
    for d in mpp.values(): month_set.update(d['months'].keys())
    for d in ot.values():  month_set.update(d['months'].keys())
    months = sorted(month_set, key=lambda m: MONTH_ORDER.index(m))
    print(f'\n📅 Months: {months}')
    return months

# ── Build Driver Data ─────────────────────────────────────────────────────────

def build_driver_data(mpp, ot, months):
    """
    Base = MPP (insentif). Join OT by NIK.
    Driver yang ada OT tapi tidak di MPP tetap masuk (OT only).
    """
    all_niks = set(mpp.keys()) | set(ot.keys())
    drivers = []

    for nik in all_niks:
        m = mpp.get(nik)
        o = ot.get(nik)

        # Nama & site dari MPP kalau ada, else dari OT
        if m:
            name = m['name']
            site = m['site']
        else:
            name = o['name'] if o else nik
            site = o['site_cat'] if o else '-'

        # Site category dari OT kalau ada
        site_cat = o['site_cat'] if o else ''
        location = o['location'] if o else ''
        bu       = o['bu'] if o else ''

        monthly = {}
        total_ot_hours = 0.0
        total_ot_idr   = 0.0
        total_ins      = 0.0

        for month in months:
            ot_h   = o['months'][month]['hours'] if o and month in o['months'] else 0.0
            ot_idr = o['months'][month]['idr']   if o and month in o['months'] else 0.0
            ins    = m['months'][month]           if m and month in m['months'] else 0.0

            monthly[month] = {
                'ot_hours': round(ot_h, 2),
                'ot_idr'  : round(ot_idr),
                'insentif': round(ins),
                'total'   : round(ot_idr + ins),
            }
            total_ot_hours += ot_h
            total_ot_idr   += ot_idr
            total_ins      += ins

        drivers.append({
            'nik'           : nik,
            'name'          : name,
            'site'          : site,
            'site_cat'      : site_cat,
            'location'      : location,
            'bu'            : bu,
            'has_ot'        : o is not None,
            'has_ins'       : m is not None,
            'monthly'       : monthly,
            'total_ot_hours': round(total_ot_hours, 2),
            'total_ot_idr'  : round(total_ot_idr),
            'total_ins'     : round(total_ins),
            'grand_total'   : round(total_ot_idr + total_ins),
        })

    drivers.sort(key=lambda x: -x['grand_total'])

    both    = sum(1 for d in drivers if d['has_ot'] and d['has_ins'])
    ot_only = sum(1 for d in drivers if d['has_ot'] and not d['has_ins'])
    ins_only= sum(1 for d in drivers if not d['has_ot'] and d['has_ins'])
    print(f'\n✅ Total drivers: {len(drivers)} (Both: {both}, OT only: {ot_only}, Ins only: {ins_only})')
    return drivers

# ── HTML Update ───────────────────────────────────────────────────────────────

def jd(obj):
    return json.dumps(obj, separators=(',', ':'), ensure_ascii=False)

def update_html(drivers, months, last_data_date):
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    html = re.sub(r'Update: [^<"]+', f'Update: {last_data_date}', html)
    html = re.sub(r'const MONTHS=\[[^\]]*\]', f'const MONTHS={jd(months)}', html)
    html = re.sub(r'const DRIVERS=\[.*?\];', f'const DRIVERS={jd(drivers)};', html, flags=re.DOTALL)

    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    wib = timezone(timedelta(hours=7))
    now = datetime.now(wib).strftime('%d %b %Y %H:%M WIB')
    print(f'✅ HTML updated [{now}]')

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print('=== Variable Income Dashboard — Auto Update ===\n')
    gc = get_gc()

    wb_ot  = gc.open_by_key(os.environ['SHEET_ID_OT'])
    wb_ins = gc.open_by_key(os.environ['SHEET_ID_INSENTIF'])

    mpp     = extract_insentif(wb_ins)
    ot      = extract_ot(wb_ot)
    months  = detect_months(mpp, ot)
    drivers = build_driver_data(mpp, ot, months)

    wib = timezone(timedelta(hours=7))
    today = datetime.now(wib)
    last_month = months[-1] if months else today.strftime('%B')
    last_month_idx = MONTH_ORDER.index(last_month) + 1
    last_data_date = f"{today.day} {MONTH_ID[last_month_idx]} {today.year}"

    update_html(drivers, months, last_data_date)

if __name__ == '__main__':
    main()
