"""
app.py  –  Invesmate Analytics Dashboard
Streamlit app: upload 3 Excel files → get the full dashboard
"""
import streamlit as st
import json
import os
from pathlib import Path
from data_processor import process_all

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Invesmate Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Hide Streamlit chrome for cleaner look
st.markdown("""
<style>
  #MainMenu, footer, header { visibility: hidden; }
  .stApp { background: #060910; }
  .block-container { padding: 0 !important; max-width: 100% !important; }
  section[data-testid="stSidebar"] { display: none; }
  div[data-testid="stToolbar"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# LOAD TEMPLATE  (robust path resolution for Streamlit Cloud)
# ─────────────────────────────────────────────────────────────────────────────
from pathlib import Path

def _find_template():
    here = Path(__file__).resolve().parent
    candidates = [
        here / 'dashboard_template.html',
        Path(os.getcwd()) / 'dashboard_template.html',
        Path('/mount/src') / here.name / 'dashboard_template.html',
        Path('/mount/src/webinar-seminar-report/dashboard_template.html'),
    ]
    for p in candidates:
        if p.exists():
            return p
    # Last resort: walk up
    for p in Path(os.getcwd()).rglob('dashboard_template.html'):
        return p
    return None

_tmpl = _find_template()
if _tmpl is None:
    st.error("dashboard_template.html not found. Make sure it is committed to your repo.")
    st.stop()

with open(_tmpl, 'r', encoding='utf-8') as f:
    DASHBOARD_TEMPLATE = f.read()

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
if 'dashboard_html' not in st.session_state:
    st.session_state.dashboard_html = None
if 'processing' not in st.session_state:
    st.session_state.processing = False

# ─────────────────────────────────────────────────────────────────────────────
# UPLOAD PAGE
# ─────────────────────────────────────────────────────────────────────────────
def show_upload_page():
    st.markdown("""
    <div style="
        min-height:100vh;
        background:linear-gradient(135deg,#060910 0%,#0c1018 50%,#111520 100%);
        display:flex;flex-direction:column;align-items:center;justify-content:center;
        padding:40px 20px;font-family:'DM Sans',sans-serif;
    ">
    <link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">
    </div>
    """, unsafe_allow_html=True)

    # Header
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="text-align:center;margin-bottom:40px">
          <div style="width:60px;height:60px;background:linear-gradient(135deg,#4f8ef7,#b44fe7);
               border-radius:16px;display:inline-flex;align-items:center;justify-content:center;
               font-size:30px;margin-bottom:20px;box-shadow:0 0 30px rgba(79,142,247,.4)">📊</div>
          <h1 style="font-family:'Syne',sans-serif;font-size:32px;font-weight:800;
               color:#eceef5;margin:0;letter-spacing:-1px">Invesmate Analytics Hub</h1>
          <p style="color:#4a5068;font-size:14px;margin:8px 0 0;letter-spacing:.5px;text-transform:uppercase">
            Upload your data files to generate the full dashboard
          </p>
        </div>
        """, unsafe_allow_html=True)

        # Info box
        st.markdown("""
        <div style="background:rgba(79,142,247,.06);border:1px solid rgba(79,142,247,.15);
             border-radius:12px;padding:16px 20px;margin-bottom:30px">
          <p style="color:#8a90aa;font-size:13px;margin:0;line-height:1.7">
            <strong style="color:#4f8ef7">Required files (3):</strong><br>
            🔵 <strong style="color:#eceef5">Free_Class_Lead_Report</strong> — BCMB & INSIGNIA webinar data (BCMB + INSG sheets)<br>
            🟠 <strong style="color:#eceef5">Offline_Seminar_Report</strong> — Seminar operations & financials (Offline Report sheet)<br>
            🟣 <strong style="color:#eceef5">Offline_Indepth_Details_Attendees</strong> — Student enrollment data (multi-location sheets)
          </p>
        </div>
        """, unsafe_allow_html=True)

    # Upload widgets
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div style="background:#0c1018;border:1px solid rgba(255,255,255,.06);border-radius:12px;padding:16px;margin-bottom:12px">
          <div style="font-size:22px;margin-bottom:8px">🔵</div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#eceef5;margin-bottom:4px">
            Free Class Lead Report
          </div>
          <div style="font-size:11px;color:#4a5068">Contains BCMB & INSIGNIA webinar performance data</div>
        </div>
        """, unsafe_allow_html=True)
        webinar_file = st.file_uploader(
            "Upload Free_Class_Lead_Report",
            type=['xlsx', 'xls'],
            key='webinar_file',
            label_visibility='collapsed'
        )

    with col2:
        st.markdown("""
        <div style="background:#0c1018;border:1px solid rgba(255,255,255,.06);border-radius:12px;padding:16px;margin-bottom:12px">
          <div style="font-size:22px;margin-bottom:8px">🟠</div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#eceef5;margin-bottom:4px">
            Offline Seminar Report
          </div>
          <div style="font-size:11px;color:#4a5068">Seminar financials: revenue, expenses, attendance, SB</div>
        </div>
        """, unsafe_allow_html=True)
        seminar_file = st.file_uploader(
            "Upload Offline_Seminar_Report",
            type=['xlsx', 'xls'],
            key='seminar_file',
            label_visibility='collapsed'
        )

    with col3:
        st.markdown("""
        <div style="background:#0c1018;border:1px solid rgba(255,255,255,.06);border-radius:12px;padding:16px;margin-bottom:12px">
          <div style="font-size:22px;margin-bottom:8px">🟣</div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#eceef5;margin-bottom:4px">
            Attendee Details
          </div>
          <div style="font-size:11px;color:#4a5068">Student-level enrollment, payments, sales rep data</div>
        </div>
        """, unsafe_allow_html=True)
        attendee_file = st.file_uploader(
            "Upload Offline_Indepth_Details",
            type=['xlsx', 'xls'],
            key='attendee_file',
            label_visibility='collapsed'
        )

    # Generate button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        all_uploaded = webinar_file and seminar_file and attendee_file
        st.markdown("<br>", unsafe_allow_html=True)

        if all_uploaded:
            if st.button("🚀 Generate Dashboard", use_container_width=True, type="primary"):
                with st.spinner("Processing data files…"):
                    try:
                        data = process_all(webinar_file, seminar_file, attendee_file)

                        if data['errors']:
                            for e in data['errors']:
                                st.warning(f"⚠️ {e}")

                        # Build JS data injection
                        js_data = build_js_data(data)

                        # Inject into template
                        dashboard_html = DASHBOARD_TEMPLATE.replace(
                            '{{DATA_PLACEHOLDER}}', js_data
                        )

                        st.session_state.dashboard_html = dashboard_html

                        stats = data['stats']
                        st.success(
                            f"✅ Processed: {stats['bcmb_count']} BCMB · "
                            f"{stats['insg_count']} INSIGNIA · "
                            f"{stats['seminar_count']} seminars across "
                            f"{stats['locations']} locations"
                        )
                        st.rerun()

                    except Exception as e:
                        st.error(f"❌ Error processing files: {e}")
                        import traceback
                        st.code(traceback.format_exc())
        else:
            missing = []
            if not webinar_file:  missing.append("Free Class Lead Report")
            if not seminar_file:  missing.append("Offline Seminar Report")
            if not attendee_file: missing.append("Attendee Details")
            st.markdown(f"""
            <div style="text-align:center;padding:12px;background:rgba(255,255,255,.03);
                 border-radius:8px;color:#4a5068;font-size:13px">
              Waiting for: {' · '.join(missing)}
            </div>
            """, unsafe_allow_html=True)


def build_js_data(data):
    """Build the JS data constants string to inject into template."""
    def j(obj):
        return json.dumps(obj, ensure_ascii=False, default=str)

    lines = [
        f"const BCMB_DATA = {j(data['bcmb'])};",
        f"const INSG_DATA = {j(data['insg'])};",
        f"const OFFLINE_DATA = {j(data['offline'])};",
        "const ALL_DATA = [...BCMB_DATA.map(r=>({...r,course:'BCMB'})), ...INSG_DATA.map(r=>({...r,course:'INSIGNIA'})), ...OFFLINE_DATA.map(r=>({...r,course:'OFFLINE'}))];",
        f"const SEMINAR_DATA = {j(data['seminar'])};",
        f"const ATTENDEE_SUMMARY = {j(data['att_summary'])};",
        f"const SALES_REP_STATS = {j(data['sr_stats'])};",
        f"const COURSE_TYPE_STATS = {j(data['ct_stats'])};",
        f"const LOCATION_STATS_ATT = {j(data['loc_stats'])};",
    ]
    return '\n'.join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# DASHBOARD PAGE
# ─────────────────────────────────────────────────────────────────────────────
def show_dashboard():
    # Reset button in top-right
    col1, col2 = st.columns([6, 1])
    with col2:
        if st.button("← Upload New Files", key="reset_btn"):
            st.session_state.dashboard_html = None
            st.rerun()

    # Render dashboard as full-page iframe
    st.components.v1.html(
        st.session_state.dashboard_html,
        height=900,
        scrolling=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.dashboard_html:
    show_dashboard()
else:
    show_upload_page()
