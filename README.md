# STC Quality Executive Dashboard

GitHub-ready Streamlit app using one data file only: `Deviation.xlsx`.

## Files
- `app.py`
- `requirements.txt`
- `README.md`
- `Deviation.xlsx`

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Main Output
The dashboard and exported PDF include an Executive Work Order Summary table with:

- WorkOrderNum
- Total Deviation
- Penalty applied
- No penalty applied
- Expected Penalty Waived
- Civil Deviation
- Civil Penalty applied
- Fiber Deviation
- Fiber Penalty applied
- Safety Deviation
- Safety Penalty applied
- % OF Total
- Civil Foreman
- Inspector

`Service Affecting` and `Status` are intentionally excluded from the exported PDF.

## Expected Penalty Waived
Column `Expected Penalties` means the penalty was expected to be applied, but it was waived/cancelled.
