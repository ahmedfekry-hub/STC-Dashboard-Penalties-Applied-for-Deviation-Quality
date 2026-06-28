# STC Deviation Quality Dashboard — PDF Export Version

This package updates the previous STC Deviation dashboard to match the requested export approach.

## Key updates
- Removed Risk Score completely.
- Added an **Export Executive PDF Report** button inside the dashboard.
- WO ranking is based on:
  - Total deviation count
  - WO share %
  - Penalty count and penalty %
  - Service-affecting count and %
  - Number of deviation types
- PDF report includes:
  - Executive KPI cards
  - Top WO chart
  - Most common deviation chart
  - WO ranking table
  - District heatmap
  - Deviation nature table
  - Management action tracker

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Source
- `Deviation.xlsx` is the source data file.
