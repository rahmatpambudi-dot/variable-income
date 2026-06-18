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

# OT Classification
# ── OT Category Classification ───────────────────────────────────────────────
# Urutan penting: cek dari yang paling spesifik ke umum

STORE_BRANDS   = ['AZKO','CHATIME','TGI','ASTORE','ASTOR']
STORE_KW       = ['STORE','TOKO','DK B','DK-B','ZYLLEM','RACKING',
                  'SOFT OPENING','GRAND OPENING','OPENING STORE',
                  'PAMERAN','TOYS','CLOSINGAN','BAZAAR','BAZZAR',
                  'BACK UP OPERASIONAL STORE','RITTASE STORE']
CUSTOMER_KW    = ['DO CUSTOMER','DO KE CUSTOMER','ANTAR DO','KIRIM DO',
                  'MENGIRIM DO','MENGIRIMKAN DO','DELIVERY TO CUST',
                  'DELIVERY TO CUSTOMER','DELIV TO CUST',
                  'CUST ','CUSTOMER','TITIK','DELIVERI BARANG',
                  'KIRIM BARANG CUST','KIRIM BARANG CUSTOMER',
                  'LEMBUR KIRIM','PENGIRIMAN SESUAI','DROP POINT',
                  'OD ','CILEGON','PARCEL','TARIK BARANG',
                  'OVER CAPACITY','TOTAL DEMAN','SUPPORT KIRIM',
                  'PROJEK CUSTOMER','DO CIKUPA']
AREA_KW        = ['KARAWANG','TANGGERANG','TANGERANG','CIKARANG','DEPOK',
                  'CIBITUNG','SATELIT','JAKUT','JAKSEL','JAKTIM','JAKBAR',
                  'BEKASI','PURWAKARTA','BOGOR','PIK','PLUIT','BAYWALK',
                  'SOHO','GREENVILE','CIMANGGIS']

def classify_ot(paid_type, description):
    pt   = str(paid_type).strip().upper()
    desc = str(description).strip().upper()

    # 1. Holiday
    if 'HOLIDAY' in pt or any(x in desc for x in ['HOVT','LEBARAN','HARI RAYA','HOLIDAY OT','HOLIDAY OVERTIME']):
        return 'Holiday'

    # 2. Langsiran
    if any(x in desc for x in ['LANGSIRAN','LANSIRAN','LANGSIR ']):
        return 'Langsiran'

    # 3. Bongkar/Muat (pure bongkar tanpa kirim)
    if any(x in desc for x in ['BONGKAR CONTAINER','RECEIVED BARANG','BONGKARAN']) or 'BONGKAR CONTAINER' in pt:
        return 'Bongkar/Muat'

    # 4. Stock Opname
    if any(x in desc for x in ['STOCK OPNAME','OPNAME','STOCK OPNAM']) or 'STOCK OPNAME' in pt:
        return 'Stock Opname'

    # 5. Physical Check
    if any(x in desc for x in ['PHYSICAL CHECK','PENGECEKAN ASSET','CEK ASSET',
                                 'PYISICAL CHEK','PHYICAL CHECK','PHYSICAL CHEK',' PC ']):
        return 'Physical Check'

    # 6. KIR
    if desc.strip() == 'KIR' or desc.startswith('KIR '):
        return 'KIR'

    # 7. Pengiriman Luar Kota
    if any(x in desc for x in ['KELUAR KOTA','LUAR KOTA']) or 'KELUAR KOTA' in pt:
        return 'Pengiriman Luar Kota'

    # 8. Perbantuan
    if 'PERBANTUAN' in desc:
        return 'Perbantuan'

    # 9. Opening Store
    if any(x in desc for x in ['SOFT OPENING','GRAND OPENING','OPENING STORE']):
        return 'Opening Store'

    # 10. Project
    if any(x in desc for x in ['PROJECT','PROYEK','PROJEK','PROJECK','MOBIL LISTRIK','MARMER']):
        return 'Project'

    # 13. Parkir
    if any(x in desc for x in ['PARKIR','CORPORATE','PARKING']):
        return 'Parkir'

    # 14. Delivery Mix (ada unsur store DAN customer, atau lembur tim HUB)
    has_store    = any(x in desc for x in STORE_BRANDS + STORE_KW)
    has_customer = any(x in desc for x in CUSTOMER_KW + AREA_KW)
    if has_store and has_customer:
        return 'Delivery Mix'
    if any(x in desc for x in ['OT TIM DC HUB','OVERTIME TIM','OVERTIME TIM DC HUB',
                                 'LEMBUR DRIVER DC HUB','LEMBUR DRIVER & ASST DRIVER DC HUB',
                                 'LEMBURAN DRIVER DC HUB','PERIODE CUT OFF','CUT OFF PERIODE',
                                 'PERIODE 16 J','PERIODE 12 ']):
        return 'Delivery Mix'

    # 15. Delivery Store
    if has_store:
        return 'Delivery Store'

    # 16. Ritase
    if any(x in desc for x in ['RITASE','RTIASE','RITAASE','RIT 2','RIT2']):
        return 'Ritase'

    # 17. Delivery Customer
    if has_customer:
        return 'Delivery Customer'

    # 18. Ritase typo
    if any(x in desc for x in ['RIT2','RIT 2','RTIASE','RITAASE','2 RIT','3 RIT','4 RIT']):
        return 'Ritase'

    # 19. Physical Check singkatan
    if desc.strip() == 'PC':
        return 'Physical Check'

    # 20. Support Pengantaran
    if any(x in desc for x in ['ANTAR JEMPUT KARYAWAN','ANTRVJEMPUT','ANNTAR JEMPUT',
                                 'ANTR JEMPUT KARYAWAN','JEMPUT KARYAWAN',
                                 'DEPO + BANDARA','TGR + BANDARA','BANDARA',
                                 'BCA + ','HO + RUKO','PT KALDEN',
                                 'ANTAR PAK ','PROCUREMENT','VISIT KE VENDOR',
                                 'JALUR KE VENDOR','AMBIL BARANG DI VENDOR',
                                 'ANTAR JEMPUT','JEMPUT YANG LEMBUR']):
        return 'Support Pengantaran'

    # 21. Delivery Customer lainnya
    if any(x in desc for x in ['ACC PAK MARCO','OVERTIME ACC',
                                 'LOADING','MUAT DAN KIRIM','BONGKAR MUAT DAN KIRIM',
                                 'BONGKAR, MUAT','BONGKAR MUTA','DELIVERY','DELIV',
                                 'MENGIRIM','PENGIRIMAN','PENGANTARAN','KIRIM']):
        return 'Delivery Customer'

    # 22. Bongkar/Muat typo
    if any(x in desc for x in ['BONGKARAN','BONGAKARAN']):
        return 'Bongkar/Muat'

    # 23. Tidak ada deskripsi
    if desc == '':
        return 'Tidak Ada Deskripsi'

    return 'Lainnya'

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

def normalize_nik(nik):
    n = str(nik).strip()
    try:
        return str(int(n))  # strip leading zeros
    except:
        return n.upper()  # keep E-prefix NIKs as-is

def is_dummy(nik, name=''):
    if str(nik).strip() in DUMMY_NIKS: return True
    return any(kw in str(name).strip().upper() for kw in DUMMY_NAME_KEYWORDS)

def normalize_date(val, expected_month_num=None):
    """Parse OT Date ke format YYYY-MM-DD. Return '' jika gagal.
    expected_month_num: int 1-12, dipakai untuk resolve ambiguitas M/D vs D/M."""
    v = str(val).strip()
    if not v or v in ('None', '-', ''):
        return ''
    # YYYY-MM-DD
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', v)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    # DD/MM/YYYY or MM/DD/YYYY — ambiguous kalau keduanya <=12
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', v)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), m.group(3)
        if a > 12:  # pasti D/M/YYYY
            d, mo = a, b
        elif b > 12:  # pasti M/D/YYYY
            d, mo = b, a
        elif expected_month_num:
            # Gunakan Month Rev untuk resolve ambiguitas
            if b == expected_month_num:
                d, mo = a, b  # M/D/YYYY: b adalah bulan
            elif a == expected_month_num:
                d, mo = b, a  # D/M/YYYY: a adalah bulan (swap)
            else:
                d, mo = a, b  # fallback: asumsikan M/D/YYYY (format Google Sheets)
        else:
            d, mo = a, b  # fallback: asumsikan M/D/YYYY (format Google Sheets default)
        return f"{y}-{mo:02d}-{d:02d}"
    # DD-MM-YYYY
    m = re.match(r'^(\d{1,2})-(\d{1,2})-(\d{4})$', v)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    # Coba parse via datetime
    for fmt in ('%d/%m/%Y','%m/%d/%Y','%Y/%m/%d','%d-%m-%Y','%B %d, %Y','%b %d, %Y'):
        try:
            return datetime.strptime(v, fmt).strftime('%Y-%m-%d')
        except:
            pass
    return ''

def normalize_month(m):
    m = str(m).strip()
    if m in MONTH_ORDER: return m
    m_stripped = m.lstrip('0') or '0'
    result = MONTH_NUM_MAP.get(m_stripped, MONTH_NUM_MAP.get(m, ''))
    if result: return result
    match = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', m)
    if match:
        return MONTH_NUM_MAP.get(match.group(1).lstrip('0') or '0', '')
    match = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', m)
    if match:
        return MONTH_NUM_MAP.get(match.group(2).lstrip('0') or '0', '')
    match = re.match(r'^(\d{1,2})-(\d{1,2})-(\d{4})$', m)
    if match:
        return MONTH_NUM_MAP.get(match.group(2).lstrip('0') or '0', '')
    return ''

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
    loc_upper = str(location).strip().upper()
    for keywords, label in LOCATION_TO_SITE:
        if any(k in loc_upper for k in keywords):
            return label
    return location

def get_site_cat(location, raw_site_cat):
    loc_upper = str(location).strip().upper()
    for key, cat in SITE_CAT_OVERRIDE.items():
        if key.upper() in loc_upper:
            return cat
    if 'HUB' in loc_upper and raw_site_cat not in ('NDC', 'RDC'):
        return 'HUB'
    return raw_site_cat if raw_site_cat else 'Lainnya'

# ── Extract Insentif ──────────────────────────────────────────────────────────

def extract_insentif(wb_ins):
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
            c_nik1   = col_first(headers, 'NIK1')
            c_name1  = col_first(headers, 'driver')
            c_nik2   = col_first(headers, 'nik2')
            c_name2  = col_first(headers, 'kenek1')
            c_month  = col_first(headers, 'Month Rev')
            c_ins    = col_first(headers, 'Insentif per MPP')
            c_tgl    = col_first(headers, 'Tanggal')   # ← NEW kolom A

            if c_nik1 < 0:
                print(f'  [WARN] {site}: NIK1 tidak ditemukan')
                continue

            count = 0
            date_ok = 0
            for row in all_rows[1:]:
                def g(c): return str(row[c]).strip() if 0<=c<len(row) else ''
                month   = normalize_month(g(c_month))
                ins     = to_num(g(c_ins))
                expected_m_num2 = MONTH_ORDER.index(month) + 1 if month else None
                tgl_raw = normalize_date(g(c_tgl), expected_m_num2) if c_tgl >= 0 else ''  # ← raw parse

                # Validate tgl: bulan harus match Month Rev
                # Kalau tidak match atau kosong → skip (jangan fallback ke -01)
                tgl = ''
                if month and tgl_raw:
                    expected_month_num = MONTH_ORDER.index(month) + 1
                    try:
                        parsed_month_num = int(tgl_raw[5:7])
                        if parsed_month_num == expected_month_num:
                            tgl = tgl_raw  # valid
                        # else: tidak match → skip
                    except:
                        tgl = ''
                # elif month and not tgl_raw: skip — tidak ada tanggal → tidak masuk dates

                if not month or ins <= 0:
                    continue
                if tgl:
                    date_ok += 1

                for c_nik, c_name in [(c_nik1, c_name1), (c_nik2, c_name2)]:
                    nik  = g(c_nik)
                    name = g(c_name)
                    if is_dummy(nik, name) or not nik:
                        continue
                    nik = normalize_nik(nik)
                    if not nik or nik in ('999999', '0'):
                        continue

                    if nik not in nik_to_ins:
                        nik_to_ins[nik] = {
                            'name'    : name,
                            'ins_site': site,
                            'months'  : defaultdict(float),
                            'dates'   : defaultdict(float),  # ← NEW
                        }

                    nik_to_ins[nik]['months'][month] += ins
                    if tgl:
                        nik_to_ins[nik]['dates'][tgl] += ins  # ← NEW
                    count += 1

            print(f'  ✅ {site}: {count} rows, {date_ok} with date')
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
        'nik'       : col_first(headers, 'Employee ID'),
        'name'      : col_first(headers, 'Employee Name'),
        'month'     : col_last(headers, 'Month'),
        'ot_date'   : col_first(headers, 'OT Date'),   # ← NEW
        'hours'     : col_contains_last(headers, 'OT Hour Paid'),
        'idr'       : col_contains_last(headers, 'OT (IDR)'),
        'location'  : col_last(headers, 'Location Name'),
        'bu'        : col_last(headers, 'BU'),
        'site_cat'  : col_last(headers, 'Site Category'),
        'paid_type' : col_first(headers, 'Paid Type'),
        'desc'      : col_first(headers, 'Description'),
    }
    print(f'  Column indices: {ci}')
    print(f'  OT Date col: {ci["ot_date"]} {"✅" if ci["ot_date"] >= 0 else "❌ NOT FOUND"}')

    ot = {}
    skipped = 0
    date_ok = 0
    date_fail = 0

    for row in all_rows[2:]:
        def g(c): return str(row[c]).strip() if 0<=c<len(row) else ''

        nik      = g(ci['nik'])
        name     = g(ci['name'])
        location  = g(ci['location'])
        bu        = g(ci['bu'])
        # Cikupa hanya punya 1 BU (HCI) — force override
        if 'CIKUPA' in g(ci['location']).upper():
            bu = 'HCI'
        site_cat  = g(ci['site_cat'])
        hours     = to_num(g(ci['hours']))
        idr       = to_num(g(ci['idr']))
        month     = normalize_month(g(ci['month']))
        paid_type = g(ci['paid_type'])
        desc      = g(ci['desc'])
        expected_m_num = MONTH_ORDER.index(month) + 1 if month else None
        ot_date_raw = normalize_date(g(ci['ot_date']), expected_m_num) if ci['ot_date'] >= 0 else ''

        # Validate OT Date: bulan harus match Month Rev
        # Kalau tidak match atau kosong → skip (jangan fallback ke -01)
        ot_date = ''
        if month and ot_date_raw:
            expected_m = MONTH_ORDER.index(month) + 1
            try:
                parsed_m = int(ot_date_raw[5:7])
                if parsed_m == expected_m:
                    ot_date = ot_date_raw  # valid, pakai apa adanya
                # else: tanggal tidak match bulan → skip, biarkan ot_date = ''
            except:
                ot_date = ''
        # elif month and not ot_date_raw: skip — tidak ada tanggal → tidak masuk dates

        if is_dummy(nik, name) or not nik or not month:
            skipped += 1
            continue
        nik = normalize_nik(nik)

        if ot_date: date_ok += 1
        else: date_fail += 1

        ot_cat = classify_ot(paid_type, desc)

        if nik not in ot:
            ot[nik] = {
                'name'            : name,
                'location'        : location,
                'display_site'    : get_display_site(location),
                'site_cat'        : get_site_cat(location, site_cat),
                'bu'              : bu,
                'months'          : defaultdict(lambda: {'hours':0.0,'idr':0.0}),
                'dates'           : defaultdict(lambda: {'hours':0.0,'idr':0.0}),  # ← NEW
                'ot_cats_total'   : defaultdict(lambda: {'hours':0.0,'idr':0.0}),
                'ot_cats_monthly' : defaultdict(lambda: defaultdict(lambda: {'hours':0.0,'idr':0.0})),
            }

        ot[nik]['months'][month]['hours'] += hours
        ot[nik]['months'][month]['idr']   += idr
        ot[nik]['ot_cats_total'][ot_cat]['hours'] += hours
        ot[nik]['ot_cats_total'][ot_cat]['idr']   += idr
        ot[nik]['ot_cats_monthly'][month][ot_cat]['hours'] += hours
        ot[nik]['ot_cats_monthly'][month][ot_cat]['idr']   += idr

        # Simpan per tanggal ← NEW
        if ot_date:
            ot[nik]['dates'][ot_date]['hours'] += hours
            ot[nik]['dates'][ot_date]['idr']   += idr

    print(f'  ✅ OT: {len(ot)} drivers, {skipped} rows skipped')
    print(f'  OT Date parsed: {date_ok} ok, {date_fail} failed')
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
            ins_tab  = ins['ins_site'] if ins else ''
            name     = ins['name'] if ins else nik
            site     = INS_SITE_DISPLAY.get(ins_tab, ins_tab or '-')
            site_cat = INS_SITE_CAT.get(ins_tab, 'HUB')
            location = site
            bu       = 'HCI' if ins_tab == 'CKP' else ''

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

        # OT per tanggal
        dates = {}
        if o and o.get('dates'):
            for dt, vals in o['dates'].items():
                dates[dt] = {
                    'hours': round(vals['hours'], 2),
                    'idr'  : round(vals['idr']),
                }

        # Insentif per tanggal ← NEW
        ins_dates = {}
        if ins and ins.get('dates'):
            for dt, val in ins['dates'].items():
                ins_dates[dt] = round(val)

        # OT categories breakdown
        ot_cats = {}
        ot_cats_monthly = {}
        if o and 'ot_cats_total' in o:
            for cat, vals in o['ot_cats_total'].items():
                ot_cats[cat] = {
                    'hours': round(vals['hours'], 2),
                    'idr'  : round(vals['idr']),
                    'pct'  : round(vals['idr'] / total_ot_idr * 100, 1) if total_ot_idr > 0 else 0,
                }
        if o and 'ot_cats_monthly' in o:
            for month_key, cats in o['ot_cats_monthly'].items():
                ot_cats_monthly[month_key] = {}
                month_ot = sum(v['idr'] for v in cats.values()) or 1
                for cat, vals in cats.items():
                    ot_cats_monthly[month_key][cat] = {
                        'hours': round(vals['hours'], 2),
                        'idr'  : round(vals['idr']),
                        'pct'  : round(vals['idr'] / month_ot * 100, 1),
                    }

        drivers.append({
            'nik'             : nik,
            'name'            : name,
            'site'            : site,
            'site_cat'        : site_cat,
            'location'        : location,
            'bu'              : bu,
            'has_ot'          : o is not None,
            'has_ins'         : ins is not None,
            'monthly'         : monthly,
            'dates'           : dates,      # {YYYY-MM-DD: {hours, idr}}
            'ins_dates'       : ins_dates,  # {YYYY-MM-DD: idr}
            'ot_cats'         : ot_cats,
            'ot_cats_monthly' : ot_cats_monthly,
            'total_ot_hours'  : round(total_ot_hours, 2),
            'total_ot_idr'    : round(total_ot_idr),
            'total_ins'       : round(total_ins),
            'grand_total'     : round(total_ot_idr + total_ins),
        })

    drivers.sort(key=lambda x: -x['grand_total'])

    both     = sum(1 for d in drivers if d['has_ot'] and d['has_ins'])
    ot_only  = sum(1 for d in drivers if d['has_ot'] and not d['has_ins'])
    ins_only = sum(1 for d in drivers if not d['has_ot'] and d['has_ins'])

    from collections import defaultdict as _dd
    cat_totals = _dd(lambda: {'hours':0.0,'idr':0.0})
    for d in drivers:
        for cat, vals in d.get('ot_cats',{}).items():
            cat_totals[cat]['hours'] += vals['hours']
            cat_totals[cat]['idr']   += vals['idr']
    total_ot_all = sum(v['idr'] for v in cat_totals.values()) or 1
    print(f'\n📊 OT Classification:')
    for cat, vals in sorted(cat_totals.items(), key=lambda x: -x[1]['idr']):
        pct = vals['idr']/total_ot_all*100
        print(f'  {cat}: Rp {vals["idr"]:,.0f} ({pct:.1f}%) — {vals["hours"]:.0f} jam')

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
