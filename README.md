# STC Deviation Board-ready Dashboard - Fixed PDF Export

## Contents
- `app.py`
- `requirements.txt`
- `README.md`
- `Deviation.xlsx`

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Updates included
- Fixed PDF export `HexColor` issue.
- PDF report includes a professional donut chart with outside percentage callouts/connectors.
- Main PDF table includes: Work Order, District, Civil Foreman, Inspector, Total Deviations, Penalty Applied, No Penalty Applied, Expected Penalties Waived, Service Affecting, Civil, Fiber, Safety, Other, % of Total, and Top Deviation Nature.
- `Expected Penalties Waived` means penalties that should have been applied but were cancelled/not applied.
- Uses only one source data file: `Deviation.xlsx`.
