"""
extract_data.py - Variable Income Dashboard
Base = OT. Join insentif by NIK (akumulasi semua tab).
Location → display site via hardcode mapping.
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
MONTH_NUM_MAP = {str(i+1): m for i, m in enumerate(MONTH_ORDER)}
MONTH_ID = ['','Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des']

# Mapping Location Name → display site (hardcode)
LOCATION_TO_SITE = [
    (['JABABEKA'],                       'Jababeka'),
    (['CIKUPA'],                         'Cikupa'),
    (['SIDOARJO', 'SURABAYA', 'JUANDA'], 'Sidoarjo'),
]

# Mapping insentif tab → display site (fallback untuk ins only)
INS_SITE_DISPLAY = {
    'JBBK'          : 'Jababeka',
    'CKP'           : 'Cikupa',
    'SDA'           : 'Sidoarjo',
    'Hub Bogor'     : 'Hub Bogor',
    'Hub Tangerang' : 'Hub Tangerang',
    'Hub Utara'     : 'Hub Utara',
    'Hub Bandung'   : 'Hub Bandung',
    'Hub Yogya'     : 'Hub Yogya',
    'Hub Semarang'  : 'Hub Semarang',
    'Hub Lampung'   : 'Hub Lampung',
    'Hub Palembang' : 'Hub Palembang',
    'Hub Kediri'    : 'Hub Kediri',
}

# Site category untuk insentif-only drivers
INS_SITE_CAT = {
    'JBBK'          : 'NDC',
    'CKP'           : 'NDC',
    'SDA'           : 'NDC',
    'Hub Bogor'     : 'HUB',
    'Hub Tangerang' : 'HUB',
    'Hub Utara'     : 'HUB',
    'Hub Bandung'   : 'HUB',
    'Hub Yogya'     : 'HUB',
    'Hub Semarang'  : 'HUB',
    'Hub Lampung'   : 'HUB',
    'Hub Palembang' : 'HUB',
    'Hub Kediri'    : 'HUB',
}

# Site Category override
SITE_CAT_OVERRIDE = {
    # HUB — yang salah di Raw Data
    'DC BALI - DENPASAR'            : 'HUB',
    'DC HANKAM RAYA'                : 'HUB',
    'DC BALIKPAPAN'                 : 'HUB',
    'DC ALAM SUTERA'                : 'HUB',
    'DC HUB YOGYAKARTA'             : 'HUB',
    'DC HUB SEMARANG'               : 'HUB',
    'DC HUB LP BANJARMASIN'         : 'HUB',
    'DC HUB CIKARANG GLC 7'         : 'HUB',
    'KENDARI -DC HUB KENDARI'       : 'HUB',
    'DC HUB TASIKMALAYA'            : 'HUB',
    'DC HUB RYACUDU LAMPUNG'        : 'HUB',
    'DC HUB DUMAI - BUKIT DATUK'    : 'HUB',
    'DC HUB PEKANBARU'              : 'HUB',
    'DC HUB BANYUWANGI'             : 'HUB',
    'DC HUB JEMBER'                 : 'HUB',
    'DC HUB SINGKAWANG'             : 'HUB',
    'DC HUB SUDIRMAN - PURWOKERTO'  : 'HUB',
    'DC GARUT'                      : 'HUB',
    'DC PEMATANG SIANTAR'           : 'HUB',
    'DC TEGAL'                      : 'HUB',
    'DC DAMAR PADANG'               : 'HUB',
    'DC BADUNG - BALI'              : 'HUB',
    'DC HUB JAMBI'                  : 'HUB',
    'DC HUB PALANGKARAYA'           : 'HUB',
    'DC HUB BENGKULU'               : 'HUB',
    # RDC
    'DC TALLO MAKASSAR'             : 'RDC',
    'DC TALLO MAKASSAR (AHI)'       : 'RDC',
    'DC TANJUNG MORAWA MEDAN'       : 'RDC',
    'DC TANJUNG MORAWA MEDAN (KLS)' : 'RDC',
    'DC TANJUNG MORAWA (AHI)'       : 'RDC',
    # Lainnya
    'DC CIKANDE'                    : 'Lainnya',
    'DC CIKANDE 2'                  : 'Lainnya',
    'DC CIKANDE - SERANG KM 41'     : 'Lainnya',
}

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
    try: return float(str(v).replace(',','').replace(' ',''))
    except: return 0.0

def is_dummy(nik, name=''):
    if str(nik).strip() in DUMMY_NIKS: return True
    return any(kw in str(name).strip().upper() for kw in DUMMY_NAME_KEYWORDS)

def normalize_month(m):
    m = str(m).strip()
    if m in MONTH_ORDER: return m
    return MONTH_NUM_MAP.get(m, '')

def col_first(headers, name):
    for i,h in enumerate(headers):
        if str(h).strip().lower()==name.strip().lower(): return i
    return -1

def col_last(headers, name):
    for i in range(len(headers)-1,-1,-1):
        if str(headers[i]).strip().lower()==name.strip().lower(): return i
    return -1

def col_contains_last(headers, keyword):
    for i in range(len(headers)-1,-1,-1):
        if keyword.lower() in str(headers[i]).strip().lower(): return i
    return -1

def get_display_site(location):
    """Map location name ke display site label"""
    loc_upper = str(location).strip().upper()
    for keywords, label in LOCATION_TO_SITE:
        if any(k in loc_upper for k in keywords):
            return label
    return location  # fallback ke nama asli

def get_site_cat(location, raw_site_cat):
    """Site category dengan override"""
    loc_upper = str(location).strip().upper()
    # Cek explicit override dulu
    for key, cat in SITE_CAT_OVERRIDE.items():
        if key.upper() in loc_upper:
            return cat
    # Rule: nama mengandung HUB tapi bukan NDC/RDC → paksa HUB
    if 'HUB' in loc_upper and raw_site_cat not in ('NDC', 'RDC'):
        return 'HUB'
    return raw_site_cat if raw_site_cat else 'Lainnya'

# ── Extract Insentif ──────────────────────────────────────────────────────────

def extract_insentif(wb_ins):
    """
    Baca semua tab insentif.
    NIK bisa ada di multiple tab (perbantuan) → akumulasi semua.
    Return: dict NIK → {name, months: {month: idr}}
    """
    print('\n📥 Extracting Insentif...')
    nik_to_ins = {}

    for i, site in enumerate(INSENTIF_SITES):
        try:
            ws = wb_ins.worksheet(site)
            all_rows = ws.get_all_values()
            if len(all_rows) < 2:
                print(f'  [SKIP] {site} kosong')
                continue

            headers = all_rows[0]
            c_nik   = col_first(headers, 'NIK1')
            c_name  = col_first(headers, 'driver')
            c_month = col_first(headers, 'Month Rev')
            c_ins   = col_first(headers, 'Insentif per MPP')

            if c_nik < 0:
                print(f'  [WARN] {site}: NIK1 tidak ditemukan')
                continue

            count = 0
            for row in all_rows[1:]:
                def g(c): return row[c].strip() if 0<=c<len(row) else ''
                nik   = g(c_nik)
                name  = g(c_name)
                month = normalize_month(g(c_month))
                ins   = to_num(g(c_ins))

                if is_dummy(nik, name) or not nik or not month or ins <= 0:
                    continue

                if nik not in nik_to_ins:
                    nik_to_ins[nik] = {'name': name, 'ins_site': site, 'months': defaultdict(float)}

                # Akumulasi — handle perbantuan (NIK di multiple tab)
                nik_to_ins[nik]['months'][month] += ins
                count += 1

            print(f'  ✅ {site}: {count} rows')
            if i < len(INSENTIF_SITES)-1: time.sleep(3)

        except gspread.exceptions.WorksheetNotFound:
            print(f'  [MISS] "{site}"')
        except Exception as e:
            print(f'  [ERROR] {site}: {e}')

    print(f'  Total insentif NIKs: {len(nik_to_ins)}')
    return nik_to_ins

# ── Extract OT ────────────────────────────────────────────────────────────────

def extract_ot(wb_ot):
    print('\n📥 Extracting OT...')
    ws = wb_ot.worksheet(OT_SHEET_TAB)
    print('  Fetching all rows...')
    all_rows = ws.get_all_values()
    if len(all_rows) < 3:
        print('  [ERROR] Sheet kosong!')
        return {}

    headers = all_rows[1]  # header di row ke-2
    print(f'  {len(headers)} kolom, {len(all_rows)-2} baris data')

    ci = {
        'nik'      : col_first(headers, 'Employee ID'),
        'name'     : col_first(headers, 'Employee Name'),
        'month'    : col_last(headers, 'Month'),
        'hours'    : col_contains_last(headers, 'OT Hour Paid'),
        'idr'      : col_contains_last(headers, 'OT (IDR)'),
        'location' : col_last(headers, 'Location Name'),
        'bu'       : col_last(headers, 'BU'),
        'site_cat' : col_last(headers, 'Site Category'),
    }
    print(f'  Column indices: {ci}')

    ot = {}
    skipped = 0

    for row in all_rows[2:]:
        def g(c): return row[c].strip() if 0<=c<len(row) else ''

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
                'name'        : name,
                'location'    : location,
                'display_site': get_display_site(location),
                'site_cat'    : get_site_cat(location, site_cat),
                'bu'          : bu,
                'months'      : defaultdict(lambda: {'hours':0.0,'idr':0.0}),
            }

        ot[nik]['months'][month]['hours'] += hours
        ot[nik]['months'][month]['idr']   += idr

    print(f'  ✅ OT: {len(ot)} drivers, {skipped} rows skipped')
    return ot

# ── Detect Months ─────────────────────────────────────────────────────────────

def detect_months(nik_to_ins, ot):
    month_set = set()
    for d in nik_to_ins.values(): month_set.update(d['months'].keys())
    for d in ot.values():         month_set.update(d['months'].keys())
    months = sorted(month_set, key=lambda m: MONTH_ORDER.index(m))
    print(f'\n📅 Months: {months}')
    return months

# ── Build Driver Data ─────────────────────────────────────────────────────────

def build_driver_data(nik_to_ins, ot, months):
    all_niks = set(ot.keys()) | set(nik_to_ins.keys())
    drivers = []

    for nik in all_niks:
        o   = ot.get(nik)
        ins = nik_to_ins.get(nik)

        if o:
            name     = o['name']
            site     = o['display_site']
            site_cat = o['site_cat']
            location = o['location']
            bu       = o['bu']
        else:
            # Ins only — pakai mapping dari tab insentif
            ins_tab  = ins['ins_site'] if ins else ''
            name     = ins['name'] if ins else nik
            site     = INS_SITE_DISPLAY.get(ins_tab, ins_tab or '-')
            site_cat = INS_SITE_CAT.get(ins_tab, 'HUB')
            location = site
            bu       = ''

        monthly = {}
        total_ot_hours = 0.0
        total_ot_idr   = 0.0
        total_ins      = 0.0

        for month in months:
            ot_h   = o['months'][month]['hours'] if o and month in o['months'] else 0.0
            ot_idr = o['months'][month]['idr']   if o and month in o['months'] else 0.0
            ins_m  = ins['months'][month]         if ins and month in ins['months'] else 0.0

            monthly[month] = {
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
            'site_cat'      : site_cat,
            'location'      : location,
            'bu'            : bu,
            'has_ot'        : o is not None,
            'has_ins'       : ins is not None,
            'monthly'       : monthly,
            'total_ot_hours': round(total_ot_hours, 2),
            'total_ot_idr'  : round(total_ot_idr),
            'total_ins'     : round(total_ins),
            'grand_total'   : round(total_ot_idr + total_ins),
        })

    drivers.sort(key=lambda x: -x['grand_total'])

    both     = sum(1 for d in drivers if d['has_ot'] and d['has_ins'])
    ot_only  = sum(1 for d in drivers if d['has_ot'] and not d['has_ins'])
    ins_only = sum(1 for d in drivers if not d['has_ot'] and d['has_ins'])
    print(f'\n✅ Total: {len(drivers)} drivers (Both:{both}, OT only:{ot_only}, Ins only:{ins_only})')
    return drivers

# ── HTML Update ───────────────────────────────────────────────────────────────

def jd(obj):
    return json.dumps(obj, separators=(',',':'), ensure_ascii=False)

def update_html(drivers, months, last_data_date):
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()
    html = re.sub(r'Update: [^<"]+', f'Update: {last_data_date}', html)
    html = re.sub(r'const MONTHS=\[[^\]]*\]', f'const MONTHS={jd(months)}', html)
    html = re.sub(r'const DRIVERS=\[.*?\];', f'const DRIVERS={jd(drivers)};', html, flags=re.DOTALL)
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    wib = timezone(timedelta(hours=7))
    print(f'✅ HTML updated [{datetime.now(wib).strftime("%d %b %Y %H:%M WIB")}]')

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print('=== Variable Income Dashboard — Auto Update ===\n')
    gc     = get_gc()
    wb_ot  = gc.open_by_key(os.environ['SHEET_ID_OT'])
    wb_ins = gc.open_by_key(os.environ['SHEET_ID_INSENTIF'])

    nik_to_ins = extract_insentif(wb_ins)
    ot         = extract_ot(wb_ot)
    months     = detect_months(nik_to_ins, ot)
    drivers    = build_driver_data(nik_to_ins, ot, months)

    wib   = timezone(timedelta(hours=7))
    today = datetime.now(wib)
    last_month      = months[-1] if months else today.strftime('%B')
    last_month_idx  = MONTH_ORDER.index(last_month) + 1
    last_data_date  = f"{today.day} {MONTH_ID[last_month_idx]} {today.year}"

    update_html(drivers, months, last_data_date)

if __name__ == '__main__':
    main()
