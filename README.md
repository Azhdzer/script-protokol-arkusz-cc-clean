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

## Run Modes

In the configuration section of `generuj_arkusze.py` you can control stages independently:

- `GENERUJ_EXCEL = True/False`
- `GENERUJ_WORD = True/False`

Behavior:

- `True/True`  -> full workflow (Excel + Word)
- `True/False` -> only Excel stages
- `False/True` -> Word generation from already existing Excel copies
- `False/False` -> no processing

## Protocol Variants

The script supports both classic `protokol CC` and `protokol CC-04` naming variants.

### E/F source columns on Strona 3

- Classic protocol: first copy uses `Q/R` (rows `20-24`), next copies move right by 2 columns.
- `CC-04` protocol: first copy uses `S/T` (rows `20-24`), next copies move right by 2 columns.

### F24 source for sheet Wyniki

F24 value is read from merged cells in row 17 on `Strona 3`:

- Classic protocol: `Q:R17`, then `S:T17`, `U:V17`, ...
- `CC-04` protocol: `S:T17`, then `U:V17`, `W:X17`, ...

### CC-04 type mapping (row 14)

For `CC-04`, type is read from merged cells in row 14 (`S:T14`, then right by 2 columns).
Detected tags: `LG`, `LD`, `PD`, `PG`.

For `CC-04`, data filling in working sheets is also adjusted:

- `D15:D19` comes from column `O` (rows per block, e.g. `O20:O24`).
- `C15:C19` column depends on the detected type:
   - `LG` -> column `K`
   - `PG` -> column `L`
   - `LD` -> column `M`
   - `PD` -> column `N`

Per working sheet (excluding `Wyniki`) the script fills:

- `LG` -> `K11=Pt100-09`, `K12=1586A-02`, `K13=101`, `K17=CC-04-L`
- `LD` -> `K11=Pt100-01`, `K12=1586A-02`, `K13=105`, `K17=CC-04-L`
- `PD` -> `K11=Pt100-18`, `K12=1586A-02`, `K13=107`, `K17=CC-04-P`
- `PG` -> `K11=Pt100-13`, `K12=1586A-02`, `K13=103`, `K17=CC-04-P`
