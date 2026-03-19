# Invesmate Analytics Dashboard

A Streamlit app that generates a full BI dashboard from 3 Excel uploads.

## 🚀 Quick Deploy to Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set **Main file path** to `app.py`
5. Deploy!

## 📁 Project Structure

```
invesmate_app/
├── app.py                    # Streamlit app (upload UI + dashboard rendering)
├── data_processor.py         # Excel parsing & data transformation
├── dashboard_template.html   # Dashboard HTML/CSS/JS (data injected at runtime)
├── requirements.txt          # Python dependencies
└── README.md
```

## 📊 Required Upload Files

| File | Purpose | Key Sheets |
|------|---------|------------|
| `Free_Class_Lead_Report.xlsx` | BCMB & INSIGNIA webinar performance | `BCMB`, `INSG` (or `INSIGNIA`) |
| `Offline_Seminar_Report.xlsx` | Seminar operations & financials | `Offline Report` |
| `Offline_Indepth_Details_Attendees.xlsx` | Student enrollment & payments | Multiple location sheets |

## 🔧 Local Development

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📈 Dashboard Tabs

- **📋 Executive** — KPIs, revenue trend, course mix, top trainers
- **⚖️ Course Compare** — BCMB vs INSIGNIA side-by-side
- **🔽 Funnel** — Targeted → Registered → Over30 → Seat Booked conversion
- **🎓 Trainers** — Per-trainer performance, drilldown, leaderboard
- **📈 Trends** — Time-series: growth, registration rate, day-of-week
- **🎥 Webinars** — Full webinar log with all filters
- **🏢 Offline** — Seminar ops + attendee intelligence + sales leaderboard
- **💡 Insights** — AI-generated strategic recommendations

## 🎛️ Filters

All filters apply globally across all tabs:
- **Course** pill: All / BCMB / INSIGNIA / Offline
- **Year** dropdown
- **Date range** From–To
- **Trainer** dropdown (auto-populated from data)
- **Mode** toggle: All / Online / Offline
- **Type** toggle: All / Live / Recorded

## 🏗️ How It Works

1. User uploads 3 Excel files via Streamlit UI
2. `data_processor.py` parses each file into structured Python dicts
3. Data is JSON-serialised and injected into `dashboard_template.html`
4. The complete HTML is rendered as a full-page component in Streamlit
