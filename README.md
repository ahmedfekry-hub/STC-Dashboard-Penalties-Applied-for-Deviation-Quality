# STC Quality Executive Dashboard

Board-ready Streamlit dashboard for `Deviation.xlsx`.

## Files
- `app.py` - Streamlit dashboard and PDF export logic
- `Deviation.xlsx` - the only data file required
- `requirements.txt` - Python dependencies
- `README.md` - deployment guide

## What is included
- Executive KPI cards
- Penalty Applied vs No Penalty vs Expected Penalties Waived donut chart
- Top Work Orders by deviation count
- Civil / Fiber / Safety classification
- Executive Work Order Summary: one row per Work Order
- Short deviation names, e.g. `DAMAGE TO PROPERTY` instead of the long original question
- Board-ready PDF export

## PDF notes
The exported PDF intentionally excludes:
- Service Affecting
- Status

The PDF table includes:
- Work Order
- District
- Civil Foreman
- Inspector
- Total Deviations
- Penalty Applied
- No Penalty
- Expected Penalties Waived
- Civil / Fiber / Safety counts
- Top 3 Deviations

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud / GitHub deployment
Upload all files in this folder to your GitHub repository root, then deploy the repository from Streamlit Cloud.
