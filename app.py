import streamlit as st
import pandas as pd
from data_processor import process_all, build_js_data

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Analytics Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------- LOAD TEMPLATE ----------------
def load_template():
    try:
        with open("dashboard_template.html", "r", encoding="utf-8") as f:
            return f.read()
    except:
        st.error("❌ dashboard_template.html not found in repository")
        st.stop()

DASHBOARD_TEMPLATE = load_template()

# ---------------- CACHE PROCESS ----------------
@st.cache_data
def process_cached(w, s, a):
    return process_all(w, s, a)

# ---------------- FILE VALIDATION ----------------
def validate_file(file, name):
    if file is None:
        st.warning(f"⚠️ {name} file missing")
        return False
    if file.size == 0:
        st.error(f"❌ {name} file is empty")
        return False
    return True

# ---------------- SESSION STATE ----------------
if "dashboard_html" not in st.session_state:
    st.session_state.dashboard_html = None

# ---------------- UPLOAD PAGE ----------------
def show_upload_page():
    st.title("📊 Upload Files to Generate Dashboard")

    webinar_file = st.file_uploader("Upload Webinar File", type=["xlsx"])
    seminar_file = st.file_uploader("Upload Seminar File", type=["xlsx"])
    attendee_file = st.file_uploader("Upload Attendee File", type=["xlsx"])

    if st.button("🚀 Generate Dashboard"):

        if not all([
            validate_file(webinar_file, "Webinar"),
            validate_file(seminar_file, "Seminar"),
            validate_file(attendee_file, "Attendee")
        ]):
            st.stop()

        with st.spinner("Processing files... ⏳"):
            data = process_cached(webinar_file, seminar_file, attendee_file)

        # Show errors
        if data["errors"]:
            with st.expander("⚠️ Data Issues Found"):
                for err in data["errors"]:
                    st.write(f"- {err}")

        js_data = build_js_data(data)
        dashboard_html = DASHBOARD_TEMPLATE.replace("{{DATA_PLACEHOLDER}}", js_data)

        st.session_state.dashboard_html = dashboard_html
        st.rerun()

# ---------------- DASHBOARD VIEW ----------------
def show_dashboard():
    st.button("🔄 Upload New Files", on_click=reset_app)

    st.download_button(
        label="📥 Download Dashboard",
        data=st.session_state.dashboard_html,
        file_name="dashboard.html",
        mime="text/html"
    )

    st.components.v1.html(st.session_state.dashboard_html, height=900, scrolling=True)

def reset_app():
    st.session_state.dashboard_html = None
    st.rerun()

# ---------------- ROUTER ----------------
if st.session_state.dashboard_html is None:
    show_upload_page()
else:
    show_dashboard()
