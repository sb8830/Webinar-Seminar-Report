import pandas as pd
import json

PRICING = {
    "BCMB": 5632,
    "INSIGNIA": 8999
}

# ---------- HELPERS ----------
def _n(x):
    try:
        return float(x)
    except:
        return 0

def _s(x):
    return str(x).strip() if pd.notna(x) else ""

# ---------- WEBINAR ----------
def parse_webinar(file):
    df = pd.read_excel(file)

    rows = []
    for _, r in df.iterrows():
        course = _s(r.get("Course", "BCMB"))
        sb = _n(r.get("Seat Booked", 0))

        revenue = sb * PRICING.get(course, 0)

        rows.append({
            "trainer": _s(r.get("Trainer")),
            "course": course,
            "revenue": revenue,
            "students": sb
        })

    return rows

# ---------- SEMINAR ----------
def parse_seminar(file):
    df = pd.read_excel(file, sheet_name=0)

    rows = []
    for _, r in df.iterrows():
        rows.append({
            "location": _s(r.get("Location")),
            "revenue": _n(r.get("Revenue")),
            "students": _n(r.get("Attended"))
        })

    return rows

# ---------- ATTENDEE ----------
def parse_attendee(file):
    df = pd.read_excel(file)

    total_students = len(df)
    total_revenue = df["Fees Paid"].sum()

    return {
        "summary": {
            "students": int(total_students),
            "revenue": float(total_revenue)
        }
    }

# ---------- MAIN ----------
def process_all(w, s, a):
    errors = []

    try:
        webinar = parse_webinar(w)
    except Exception as e:
        webinar = []
        errors.append(f"Webinar error: {e}")

    try:
        seminar = parse_seminar(s)
    except Exception as e:
        seminar = []
        errors.append(f"Seminar error: {e}")

    try:
        attendee = parse_attendee(a)
    except Exception as e:
        attendee = {"summary": {}}
        errors.append(f"Attendee error: {e}")

    return {
        "webinar": webinar,
        "seminar": seminar,
        "attendee": attendee,
        "errors": errors
    }

# ---------- BUILD JS ----------
def build_js_data(data):
    return f"""
    const DATA = {json.dumps(data)};
    """
