"""
data_processor.py  –  Invesmate Analytics Dashboard
Parses the three Excel uploads and returns JSON-ready data.
"""
import pandas as pd
import numpy as np
import re
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _n(val):
    try:
        v = float(val)
        return 0 if (np.isnan(v) or np.isinf(v)) else v
    except:
        return 0

def _s(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ''
    return str(val).strip()

def _d(val):
    try:
        if pd.isna(val): return ''
        if isinstance(val, (datetime, pd.Timestamp)):
            return val.strftime('%Y-%m-%d')
        s = str(val).strip()
        if '/' in s:
            try:
                # Handle merged cells like '18/8/2023  19/8/2023' - take first
                s = re.split(r'\s{2,}', s.strip())[0]
                return pd.to_datetime(s, dayfirst=True).strftime('%Y-%m-%d')
            except: pass
        return s[:10] if len(s) >= 10 else ''
    except:
        return ''

def _col(df, *keywords, exact=False, exclude=None):
    """Find first column matching any keyword (case-insensitive)."""
    excl = [e.lower() for e in (exclude or [])]
    cols = [str(c) for c in df.columns]
    for kw in keywords:
        kw_l = kw.lower()
        for c in cols:
            c_l = c.lower()
            if any(e in c_l for e in excl):
                continue
            if exact:
                if c_l == kw_l:
                    return c
            else:
                if kw_l in c_l:
                    return c
    return None

TRAINER_MAP = {
    'rohitava majumdar':           'Rohitava Majumder',
    'rohitav majumder':            'Rohitava Majumder',
    'rohitava majumder**':         'Rohitava Majumder',
    'debargha  saha':              'Debargho Saha',
    'debargha saha':               'Debargho Saha',
    'debargho\u00a0saha':          'Debargho Saha',
    'pratim kumer chakraborty':    'Pratim Kumar Chakraborty',
    'hironmoy laheri':             'Hironmoy Lahiri',
    'hironmoy lahiri\u00a0':       'Hironmoy Lahiri',
    'sandipan das':                'Sandipan Kumar Das',
    'kunal saha (special advanced class)':    'Kunal Saha',
    'sayan sarker(special advanced class)':   'Sayan Sarker',
    'rohitava majumdar':           'Rohitava Majumder',
}

def norm_trainer(name):
    parts = [p.strip() for p in re.split(r',|&|\n', str(name)) if p.strip()]
    out = []
    for p in parts:
        p = re.sub(r'\s*\(Special Advanced Class\)\s*', '', p, flags=re.I).strip()
        p = re.sub(r'\s+', ' ', p)
        out.append(TRAINER_MAP.get(p.lower(), p))
    return ', '.join(dict.fromkeys(out))   # dedupe, preserve order

# ─────────────────────────────────────────────────────────────────────────────
# FILE 1  →  BCMB + INSIGNIA  (Free_Class_Lead_Report.xlsx)
# ─────────────────────────────────────────────────────────────────────────────
def parse_webinar_file(file_obj):
    xl     = pd.ExcelFile(file_obj)
    sheets = xl.sheet_names

    # ── Pick the right sheet names ───────────────────────────────────────────
    SKIP = {'log','hitting','call','re-target','retarget','backup','rough',
            'comparison','summary','offline','forx','fund','hindi','invesmeet',
            'simplify','monitoring','lead wise','joining','percentage',
            'day to day','sheet1','8_45','sunday','tuesday','friday'}

    def pick_sheet(keyword):
        candidates = []
        for s in sheets:
            sl = s.lower()
            if keyword not in sl:
                continue
            if any(skip in sl for skip in SKIP):
                continue
            candidates.append(s)
        # Prefer shortest name (avoids 'BCMB_Webinar Hitting Report' etc.)
        candidates.sort(key=len)
        return candidates[0] if candidates else None

    bcmb_sheet = pick_sheet('bcmb')
    insg_sheet = pick_sheet('insg') or pick_sheet('insignia')

    bcmb_records = _parse_bcmb(xl, bcmb_sheet)
    insg_records = _parse_insg(xl, insg_sheet)
    return bcmb_records, insg_records


def _parse_bcmb(xl, sheet_name):
    if not sheet_name:
        return []
    df = xl.parse(sheet_name, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    # Column resolution  – be precise to avoid percentage / re-targeted cols
    c_trainer  = _col(df, 'trainer', exact=True) or _col(df, 'trainer', exclude=['re-target'])
    c_type     = _col(df, 'type', exact=True) or _col(df, 'location', exact=True)
    c_date     = _col(df, 'date', exact=True) or _col(df, 'date', exclude=['web','hitting'])
    c_targeted = _col(df, 'targeted', exact=True) or _col(df, 'targeted',
                      exclude=['to','%','re-','retarget','dialed','visited','regist','over','seat','new','old'])
    c_reg      = _col(df, 'registered', exact=True) or _col(df, 'registered', exclude=['%','to'])
    c_over30   = _col(df, 'over 30 min', exact=True) or _col(df, 'over 30', exclude=['%','to'])
    c_sb       = _col(df, 'seat booked', exact=True) or _col(df, 'seat booked', exclude=['%','to','amount'])
    c_joined   = _col(df, 'total joined', exact=True) or _col(df, 'joined', exclude=['%','re-','new','old','semi'])
    c_rev      = _col(df, 'seat booking amount') or _col(df, 'course amount')

    records = []
    for _, row in df.iterrows():
        date_val   = _d(row.get(c_date, '')) if c_date else ''
        targeted   = int(_n(row.get(c_targeted, 0))) if c_targeted else 0
        if not date_val or targeted < 1:
            continue

        trainer  = norm_trainer(_s(row.get(c_trainer,  'Unknown')) if c_trainer  else 'Unknown')
        type_raw = _s(row.get(c_type, 'Live'))       if c_type     else 'Live'
        reg      = int(_n(row.get(c_reg,    0)))     if c_reg      else 0
        over30   = int(_n(row.get(c_over30, 0)))     if c_over30   else 0
        sb       = int(_n(row.get(c_sb,     0)))     if c_sb       else 0
        joined   = int(_n(row.get(c_joined, 0)))     if c_joined   else 0
        revenue  = int(_n(row.get(c_rev,    0)))     if c_rev      else sb * 5632

        t = type_raw.upper()
        if 'REC' in t:                         wtype = 'Rec'
        elif 'BACKUP' in t or 'BACK' in t:     wtype = 'Backup'
        elif 'PRACTICE' in t:                  wtype = 'Practice'
        elif 'CANCEL' in t:                    wtype = 'Cancel'
        elif 'ZOOM' in t:                      wtype = 'Live\n(ZOOM)'
        else:                                   wtype = 'Live'

        if revenue == 0 and sb > 0:
            revenue = sb * 5632

        ym = date_val[:7]
        records.append({
            'date': date_val, 'yearMonth': ym,
            'trainer': trainer, 'course': 'BCMB',
            'type': wtype, 'mode': 'Online',
            'targeted': targeted, 'registered': reg,
            'over30': over30, 'seatBooked': sb,
            'joined': joined, 'revenue': revenue,
            'expenses': 0, 'surplus': revenue,
        })
    return sorted(records, key=lambda r: r['date'])


def _parse_insg(xl, sheet_name):
    if not sheet_name:
        return []
    df = xl.parse(sheet_name, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    c_trainer  = _col(df, 'trainer', exact=True)
    c_type     = _col(df, 'type', exact=True)
    c_date     = _col(df, 'date', exact=True) or _col(df, 'date', exclude=['web','hitting','hidden'])
    # INSG has 'Targated' typo
    c_targeted = (_col(df, 'targated', exact=True) or
                  _col(df, 'targeted', exact=True) or
                  _col(df, 'targated', exclude=['%','to']))
    c_reg      = _col(df, 'registered', exact=True) or _col(df, 'registered', exclude=['%','to'])
    c_over30   = _col(df, 'over 30', exact=True) or _col(df, 'over 30', exclude=['%','to'])
    c_sb       = _col(df, 'seat booked', exact=True) or _col(df, 'seat booked', exclude=['%','to'])
    c_joined   = _col(df, 'unique viewer') or _col(df, 'total joined') or _col(df, 'joined', exclude=['%'])

    records = []
    for _, row in df.iterrows():
        date_val = _d(row.get(c_date, '')) if c_date else ''
        targeted = int(_n(row.get(c_targeted, 0))) if c_targeted else 0
        if not date_val or targeted < 1:
            continue

        trainer = norm_trainer(_s(row.get(c_trainer, 'Unknown')) if c_trainer else 'Unknown')
        type_raw = _s(row.get(c_type, 'Live')) if c_type else 'Live'
        reg    = int(_n(row.get(c_reg,    0))) if c_reg    else 0
        over30 = int(_n(row.get(c_over30, 0))) if c_over30 else 0
        sb     = int(_n(row.get(c_sb,     0))) if c_sb     else 0
        joined = int(_n(row.get(c_joined, 0))) if c_joined else 0
        revenue = sb * 8999

        wtype = 'Rec' if 'REC' in type_raw.upper() else 'Live'
        ym    = date_val[:7]

        records.append({
            'date': date_val, 'yearMonth': ym,
            'trainer': trainer, 'course': 'INSIGNIA',
            'type': wtype, 'mode': 'Online',
            'targeted': targeted, 'registered': reg,
            'over30': over30, 'seatBooked': sb,
            'joined': joined, 'revenue': revenue,
            'expenses': 0, 'surplus': revenue,
        })
    return sorted(records, key=lambda r: r['date'])

# ─────────────────────────────────────────────────────────────────────────────
# FILE 2  →  Offline Seminar Report
# ─────────────────────────────────────────────────────────────────────────────
def parse_seminar_file(file_obj):
    xl = pd.ExcelFile(file_obj)
    df = xl.parse('Offline Report', header=1)
    df = df[pd.to_numeric(df['Sr No'], errors='coerce').notna()].copy()

    def n(col):
        return pd.to_numeric(df[col], errors='coerce').fillna(0)

    dates = pd.to_datetime(df['Seminar Date'], errors='coerce')
    records = []
    for i in range(len(df)):
        d = dates.iloc[i]
        if pd.isna(d):
            continue
        exp   = float(n('Actual Expenses').iloc[i])
        arev  = float(n('Actual Revenue(W/O GST)\nAttendees').iloc[i])
        trev  = float(n('Total Revenue\n(W/O GST)\nAttendees').iloc[i])
        surp  = float(n('Surplus or Deficit').iloc[i])
        erev  = float(n('Expected Revenue').iloc[i])
        sur2exp = float(n('Surplus to Expense').iloc[i]) if exp > 0 else 0

        records.append({
            'date':             str(d.date()),
            'month':            d.strftime('%Y-%m'),
            'location':         _s(df['Location'].iloc[i]).upper(),
            'trainer':          norm_trainer(_s(df['Trainer'].iloc[i]).replace('\n', ', ')),
            'targeted':         int(n('Targeted\n').iloc[i]),
            'attended':         int(n('Total\nAttended').iloc[i]),
            'sb_seminar':       int(n('Total\nSeat\nBooked\n(in Seminar)').iloc[i]),
            'sb_morning':       int(n('Morning').iloc[i]),
            'sb_evening':       int(n('Evening').iloc[i]),
            'sb_non_webinar':   int(n('Non\nWebinar').iloc[i]),
            'expenses':         round(exp,  2),
            'expected_revenue': round(erev, 2),
            'actual_revenue':   round(arev, 2),
            'total_revenue':    round(trev, 2),
            'surplus':          round(surp, 2),
            'att_rate':         round(float(n('Targeted to Attended (%)').iloc[i]) * 100, 1),
            'sb_rate':          round(float(n('Attended to Seat Booked (%)').iloc[i]) * 100, 1),
            'roi':              round(sur2exp * 100, 1),
        })
    return sorted(records, key=lambda r: r['date'])

# ─────────────────────────────────────────────────────────────────────────────
# FILE 3  →  Offline Attendee Details  (multi-sheet)
# ─────────────────────────────────────────────────────────────────────────────
LOC_MAP = {
    'bankura':'BANKURA','bongaon':'BONGAON','arambagh':'ARAMBAGH',
    'dh':'DIAMOND HARBOUR','bishnupur':'BISHNUPUR','bagnan':'BAGNAN',
    'midnapore':'MIDNAPORE','midnapure':'MIDNAPORE','chandannagar':'CHANDANNAGAR',
    'beharampur':'BEHARAMPUR','behrampur':'BEHARAMPUR',
    'krishnanagar':'KRISHNANAGAR','malda':'MALDA','raiganj':'RAIGANJ',
    'purulia':'PURULIA','burdwan':'BURDWAN','bandel':'BANDEL',
    'durgapur':'DURGAPUR','cooachbehar':'COOCHBEHAR','siliguri':'SILIGURI',
    'basirhat':'BASIRHAT','bolpur':'BOLPUR','kakdwip':'KAKDWIP',
    'balurghat':'BALURGHAT','ghatal':'GHATAL',
    'bankura_7_12':'BANKURA','chakdha':'CHAKDHA','jhargram':'JHARGRAM',
    'katwa':'KATWA','kathi':'KANTHI','bongaon_28_12':'BONGAON',
    'chandrakona':'CHANDRAKONA','bethuadahari':'BETHUADAHARI',
    'haldia':'HALDIA','sonarpur':'SONARPUR','ambika kalna':'AMBIKA KALNA',
    'alipurduar':'ALIPURDUAR','arambagh_1stfeb':'ARAMBAGH',
    'asansol':'ASANSOL','jalpaiguri':'JALPAIGURI',
    'rampurhat':'RAMPURHAT','adra':'ADRA','nabadwip':'NABADWIP',
    'coochbehar_re':'COOCHBEHAR','silliguri_re':'SILIGURI',
    'basirhat_re':'BASIRHAT','bolpur_re':'BOLPUR','kakdwip_re':'KAKDWIP',
    'durgapur_re':'DURGAPUR','bandel_re':'BANDEL','burdwan_re':'BURDWAN',
    'purulia_re':'PURULIA','raiganj_re':'RAIGANJ','malda_re':'MALDA',
    'krishnanagar_re':'KRISHNANAGAR','behrampur_re':'BEHARAMPUR',
    'chandanagar_re':'CHANDANNAGAR','midnapure_re':'MIDNAPORE',
    'bagnan_re':'BAGNAN','bishnupur_re':'BISHNUPUR',
    'dh_re':'DIAMOND HARBOUR','arambag_re':'ARAMBAGH','bongaon_re':'BONGAON',
    'midnapur':'MIDNAPORE',
}

SKIP_SHEETS_ATT = {
    'conversion 2','con','sheet1','summary','count','marking','line graph',
    'simplify bcmb','simplify_hitting','simplify insignia',
    'sales month wise summary','bcmb_webinar hitting report',
    'insignia webinar hitting report','offline seminar hitting report',
    'retargeted webinar hitting repo','webinar monitoring log',
    'lead wise log bcmb','sunday new lead percentage',
    'tuesday new lead percentage','friday new lead percentage',
    'day to day bcmb joining report','joining percentage',
    'bcmb log re-target','bcmb re-target','bcmb backup','invesmeet',
    'comparison','day to day bcmb ','forx','fund','hindi','rough sheet',
    '8_45','bcmb log','bcmb call analysis','insg','insignia log','bcmb',
    'offline','insglog','bcmblog',
}

def _classify_course(svc_name, svc_code):
    sn = str(svc_name).lower()
    sc = str(svc_code).lower()
    if 'insignia' in sn or 'global capital' in sn or 'ins10' in sc: return 'INSIGNIA'
    if 'equity' in sn and 'strategy' in sn:                          return 'Equity Strategy'
    if ('future' in sn or 'option' in sn) and 'f&o' in sn:          return 'F&O'
    if 'intraday' in sn or 'swing' in sn:                            return 'Intraday/Swing'
    if 'commodity' in sn:                                             return 'Commodity'
    return 'BCMB'  # default (power of trading, offline combo, etc.)


def parse_attendee_file(file_obj):
    xl = pd.ExcelFile(file_obj)
    all_rows = []

    for sheet in xl.sheet_names:
        if sheet.lower() in SKIP_SHEETS_ATT:
            continue
        try:
            df = xl.parse(sheet)
            if df.empty: continue
            # Normalise column names
            df.columns = [str(c).strip().lower().replace(' ','_') for c in df.columns]
            if 'student_name' not in df.columns and 'student_invid' not in df.columns:
                continue

            # Map sheet name → location
            key = re.sub(r'[\s]+', '_', sheet.lower()).strip('_')
            location = LOC_MAP.get(sheet.lower(), LOC_MAP.get(key, sheet.upper()))

            for col in ['payment_received','total_amount','total_due',
                        'total_additional_charges','total_gst']:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            df['status']   = df['status'].fillna('Unknown').astype(str).str.strip() \
                             if 'status' in df.columns else 'Unknown'
            df['location'] = location
            df['course_type'] = df.apply(
                lambda r: _classify_course(
                    r.get('service_name',''), r.get('service_code','')), axis=1)
            all_rows.append(df)
        except Exception:
            continue

    if not all_rows:
        return {}, {}, {}, {}

    att = pd.concat(all_rows, ignore_index=True)

    pr  = att['payment_received'] if 'payment_received' in att.columns else pd.Series(dtype=float)
    ta  = att['total_amount']     if 'total_amount'     in att.columns else pd.Series(dtype=float)
    td  = att['total_due']        if 'total_due'        in att.columns else pd.Series(dtype=float)
    sid = att['student_invid']    if 'student_invid'    in att.columns else pd.Series(range(len(att)))

    def safe_sum(s): return round(float(s.sum()), 2) if len(s) else 0
    def safe_mean(s): p = s[s > 0]; return round(float(p.mean()), 2) if len(p) else 0

    summary = {
        'total_students':    int(sid.nunique()),
        'total_records':     int(len(att)),
        'total_revenue':     safe_sum(pr),
        'total_amount':      safe_sum(ta),
        'total_due':         safe_sum(td),
        'active_students':   int((att['status'] == 'Active').sum()),
        'inactive_students': int((att['status'] == 'Inactive').sum()),
        'closed_students':   int((att['status'] == 'Closed').sum()),
        'avg_payment':       safe_mean(pr),
    }

    ct_stats = {}
    for ct, g in att.groupby('course_type'):
        gpr = g['payment_received'] if 'payment_received' in g.columns else pd.Series(dtype=float)
        gta = g['total_amount']     if 'total_amount'     in g.columns else pd.Series(dtype=float)
        gtd = g['total_due']        if 'total_due'        in g.columns else pd.Series(dtype=float)
        ct_stats[ct] = {
            'count':      int(len(g)),
            'revenue':    safe_sum(gpr),
            'active':     int((g['status'] == 'Active').sum()),
            'avg_amount': round(float(gta.mean()), 2) if len(gta) else 0,
            'total_due':  safe_sum(gtd),
        }

    sr_stats = {}
    srn_col = next((c for c in att.columns if 'sales_rep' in c), None)
    if srn_col:
        for rep, g in att[att[srn_col].notna()].groupby(srn_col):
            r = str(rep).strip()
            if not r or r == 'nan': continue
            gpr = g['payment_received'] if 'payment_received' in g.columns else pd.Series(dtype=float)
            sr_stats[r] = {
                'deals':    int(len(g)),
                'revenue':  safe_sum(gpr),
                'active':   int((g['status'] == 'Active').sum()),
                'avg_deal': safe_mean(gpr),
            }
    sr_stats = dict(sorted(sr_stats.items(), key=lambda x: -x[1]['revenue'])[:25])

    loc_stats = {}
    for loc, g in att.groupby('location'):
        gpr = g['payment_received'] if 'payment_received' in g.columns else pd.Series(dtype=float)
        gtd = g['total_due']        if 'total_due'        in g.columns else pd.Series(dtype=float)
        gid = g['student_invid']    if 'student_invid'    in g.columns else pd.Series(range(len(g)))
        loc_stats[loc] = {
            'students':  int(gid.nunique()),
            'revenue':   safe_sum(gpr),
            'active':    int((g['status'] == 'Active').sum()),
            'total_due': safe_sum(gtd),
        }
    loc_stats = dict(sorted(loc_stats.items(), key=lambda x: -x[1]['revenue'])[:40])

    return summary, ct_stats, sr_stats, loc_stats

# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
def process_all(webinar_file, seminar_file, attendee_file):
    errors = []

    try:
        bcmb, insg = parse_webinar_file(webinar_file)
    except Exception as e:
        errors.append(f'Webinar file error: {e}')
        bcmb, insg = [], []

    try:
        seminar = parse_seminar_file(seminar_file)
    except Exception as e:
        errors.append(f'Seminar file error: {e}')
        seminar = []

    try:
        att_summary, ct_stats, sr_stats, loc_stats = parse_attendee_file(attendee_file)
    except Exception as e:
        errors.append(f'Attendee file error: {e}')
        att_summary, ct_stats, sr_stats, loc_stats = {}, {}, {}, {}

    # OFFLINE_DATA: bridging format for ALL_DATA compatibility
    offline_rows = [{
        'date': s['date'], 'yearMonth': s['month'],
        'trainer': s['trainer'], 'location': s['location'],
        'course': 'OFFLINE', 'type': 'Offline', 'mode': 'Offline',
        'targeted': s['targeted'], 'registered': s['attended'],
        'over30': s['attended'], 'seatBooked': s['sb_seminar'],
        'joined': s['sb_seminar'],
        'revenue': s['actual_revenue'], 'expenses': s['expenses'],
        'surplus': s['surplus'],
    } for s in seminar]

    return {
        'bcmb':        bcmb,
        'insg':        insg,
        'offline':     offline_rows,
        'seminar':     seminar,
        'att_summary': att_summary,
        'ct_stats':    ct_stats,
        'sr_stats':    sr_stats,
        'loc_stats':   loc_stats,
        'errors':      errors,
        'stats': {
            'bcmb_count':    len(bcmb),
            'insg_count':    len(insg),
            'seminar_count': len(seminar),
            'locations':     len(set(s['location'] for s in seminar)),
        },
    }
