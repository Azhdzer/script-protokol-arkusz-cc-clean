# Script protokol - arkusz CC (clean)

This repository contains a clean version of the automation scripts for generating:

- Excel worksheet copies (etapy 1-6)
- Word calibration certificates (etap 7)

## Included

- `generuj_arkusze.py`
- `_diagnose.py`
- `_diagnose2.py`
- `_verify.py`
- `requirements.txt`

## Not included

Private templates, source spreadsheets, and formula workbooks are intentionally excluded.

Examples of excluded files:

- `.xls`, `.xlsx`, `.xlsm`, `.xlsb`
- `.doc`, `.docx`

## Setup

1. Create and activate a Python virtual environment.
2. Install dependencies:

   pip install -r requirements.txt

3. Open `generuj_arkusze.py` and set your local file names/paths in the configuration section.
4. Run:

   python generuj_arkusze.py

## Notes

- The script automates Excel via xlwings/COM.
- For linked formulas, required linked workbooks must be available in your local environment.
