[README.md](https://github.com/user-attachments/files/25427511/README.md)
# Statsig Order Form Generator

This app generates a branded Statsig-style order-form PDF using a guided 3-step flow.

Input paths:
- `Q/A`: manual form entry
- `Upload Document`: extract values from `.txt`, `.pdf`, or `.docx`, then edit

## Flow
- Step 1: Input Source + Customer Information (required fields + continue button)
- Step 2: Terms (start date, term months, computed end date, picklists + continue button)
- Step 3: Product Selection (dynamic services table + PDF generation)

## Quick start

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open the URL shown by Streamlit (usually `http://localhost:8501`).

## Notes
- Upload extraction is heuristic-based and works best when labels are present (`Billing Email:`, `Start Date:`, etc).
- Generated totals are calculated from `Annual Service Fee` values.
