# STC Quality Executive Dashboard - Board Ready

## Files included
- `app.py`
- `requirements.txt`
- `README.md`
- `Deviation.xlsx`

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Dashboard features
- One source file only: `Deviation.xlsx`
- Executive cards:
  - Total Deviations
  - Unique WOs
  - Penalty Applied
  - No Penalty Applied
  - Expected Penalties Waived
  - Service Affecting
- Board Work Order Executive Summary table:
  - Work Order
  - District
  - Civil Foreman
  - Inspector
  - Total Deviations
  - Penalty Applied
  - No Penalty Applied
  - Expected Penalties Waived
  - Service Affecting
  - Civil / Fiber / Safety / Other
  - % of Total
  - Top Deviation Nature
- Detailed deviation nature by WO
- Board-ready PDF export

## Important logic
`Expected Penalties` means penalties that should have been applied, but were waived/cancelled.
