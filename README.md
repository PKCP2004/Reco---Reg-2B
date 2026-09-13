# GST Reconciliation Studio

A clean, reconciliation-only Windows/Streamlit application.

## What this package contains

- `app.py` — professional reconciliation interface
- `reconciliation.py` — reconciliation engine
- `sample_data/Sample_GSTR_2B.xlsx` — ready-to-test sample
- `sample_data/Sample_Inward_Register.xlsx` — ready-to-test sample
- `requirements.txt` — only required Python packages
- `run_windows.bat` — Windows launcher
- `README.md` — setup guide

## Removed

PDF-to-Excel extraction, OCR, PDF libraries, duplicate ZIPs, `.venv`, `.vscode`,
`__pycache__`, `.pyc`, and other unnecessary generated files.

## Windows

1. Extract the ZIP.
2. Double-click `run_windows.bat`.
3. The browser opens Streamlit.
4. Use the included sample files to test.
5. Replace them with your actual GSTR-2B and purchase register.

## Matching

The application standardizes common GST column names and performs:
- Exact GSTIN + invoice-number matching
- Tax-value comparison with tolerance
- GSTIN mismatch detection
- Likely invoice-number error detection
- 2B-only identification
- Register-only identification
- Excel report export
