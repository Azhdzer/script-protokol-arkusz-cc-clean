#!/usr/bin/env python3
"""
Universal Excel/CSV parser for temperature & humidity data loggers.

Supported formats:
  • Rotronic HW4  (.xls text disguised as Excel, [#D]Measurements)
  • ALMEMO        (.xls Excel with repeated metadata blocks)
  • ElogVis PL   (.xls/.xlsx Polish logger, Data/Godzina header)
  • HOBO          (.csv with title row, LGR S/N columns)
  • Aranet        (.csv multi-sensor, Temperature (°C) / Humidity (%) columns)
  • Comet / TFA  (.xlsx single-column packed data: NNNN YYYY-MM-DD HH:MM:SS val1 val2)
  • Astozi        (.csv, Temperatura_... / Wilgotnosc_... columns)
  • Generic tabular CSV/Excel (keyword-based column detection, 20+ languages)

Output: wyniki/<source_name>_wynik.xlsx
Columns: Czas | Temperatura [°C] | Wilgotność [%RH]  (humidity only when available)

Usage:
  python analizuj_excele.py                        # processes excel_do_analizy/
  python analizuj_excele.py my_folder/             # custom input folder
  python analizuj_excele.py file.xlsx              # single file
  python analizuj_excele.py --debug                # show full tracebacks
"""

import os, re, sys, io, warnings, traceback
from pathlib import Path
import pandas as pd
from datetime import datetime

# Force UTF-8 output so special characters print correctly on Windows
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
if hasattr(sys.stderr, 'buffer'):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', line_buffering=True)

warnings.filterwarnings('ignore', category=UserWarning)
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=DeprecationWarning)

# ─── CONSTANTS ────────────────────────────────────────────────────────────────

DEFAULT_INPUT  = Path("excel_do_analizy")
DEFAULT_OUTPUT = Path("wyniki")
SUPPORTED_EXT  = {'.csv', '.xls', '.xlsx', '.txt'}
DEBUG = '--debug' in sys.argv

# Multi-language keyword patterns for column detection
TEMP_KW = re.compile(
    r'temp'                   # EN/DE/PL: temp(erature/eratur/eratura)
    r'|[°º]\s*c'             # °C unit in header
    r'|\[.{0,5}[Cc]\]'       # [°C], [ÂC], [*C] etc.
    r'|celsius'
    r'|wärme|ciepł'
    r'|[Tt]emperatur'        # DE: Temperatur
    r'|m\d+:.*°c'            # ALMEMO channel e.g. M04: °C
    r'|ntc',                  # NTC sensor
    re.I | re.UNICODE
)
HUM_KW = re.compile(
    r'wilgotn|wilg\b'        # PL: wilgotność / wilg wzgl
    r'|humid'                 # EN: humidity
    r'|feuchte|feucht'        # DE: Feuchtigkeit
    r'|\brh\b'                # RH (relative humidity)
    r'|%rh|%ww'              # units
    r'|%\s*h\b'              # %H, % H
    r'|hcrh'                  # ALMEMO humidity channel name
    r'|m\d+:.*%h',           # ALMEMO channel e.g. M14: %H
    re.I | re.UNICODE
)
TIME_KW = re.compile(
    r'\bczas\b'              # PL: czas (time)
    r'|data\s*/\s*godzina'   # PL: Data/Godzina (date/time combined)
    r'|\bdata\b'             # PL: data (date)
    r'|\bgodzina\b'          # PL: godzina (hour/time)
    r'|\btime\b'             # EN
    r'|\bdate\b'             # EN
    r'|datum'                 # DE
    r'|zeit\b'               # DE: Zeit
    r'|timestamp'
    r'|data\s+czas',         # PL: Data Czas
    re.I | re.UNICODE
)

# ─── UTILITIES ────────────────────────────────────────────────────────────────

def clean_num(val):
    """Convert to float. Handles comma decimals and '--' sentinel values."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return None if pd.isna(val) else float(val)
    s = str(val).strip().replace(',', '.')
    if re.fullmatch(r'[-–—]+\.?[-–—]*', s):
        return None
    try:
        return float(s)
    except:
        return None

def parse_dt(val):
    """Parse datetime from various string formats or native pandas/datetime types."""
    if val is None:
        return None
    if isinstance(val, pd.Timestamp):
        return val if not pd.isna(val) else None
    if isinstance(val, datetime):
        return pd.Timestamp(val)
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'nat', 'none', ''):
        return None
    # Try specific formats first (faster, more reliable)
    for fmt in (
        '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M',
        '%d.%m.%Y %H:%M:%S', '%d.%m.%Y %H:%M',
        '%d-%m-%Y %H:%M:%S', '%d-%m-%Y %H:%M',
        '%d.%m.%y %H:%M:%S', '%d.%m.%y %H:%M',
        '%m/%d/%y %H:%M:%S', '%m/%d/%Y %H:%M:%S',
        '%Y-%m-%dT%H:%M:%S',
        # Date-only fallbacks
        '%Y-%m-%d', '%d.%m.%Y', '%d.%m.%y',
    ):
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True)
    except Exception:
        return None

def find_cols(columns, pattern):
    """Return column names matching keyword pattern."""
    return [c for c in columns if pattern.search(str(c))]

def col_has_time_component(series):
    """Return True if a column already has complete datetime values (not date-only)."""
    for val in series.dropna().head(10):
        if isinstance(val, (pd.Timestamp, datetime)):
            return True
        if re.search(r'\d{2}:\d{2}', str(val)):
            return True
    return False

def detect_dayfirst(series):
    """
    Detect whether dot/slash-separated dates are DD.MM.YY or MM.DD.YY.
    Scans the full column for any component > 12 (impossible month → must be day).
    Returns True = day-first (European/PL), False = month-first (US/HOBO).
    """
    pat = re.compile(r'^(\d{1,2})[./](\d{1,2})[./]')
    first_gt12 = False
    second_gt12 = False
    for val in series.dropna():
        m = pat.match(str(val).strip())
        if not m:
            continue
        first, second = int(m.group(1)), int(m.group(2))
        if first > 12:
            first_gt12 = True
        if second > 12:
            second_gt12 = True
        if first_gt12 or second_gt12:
            break  # Early exit once we have enough evidence
    if first_gt12:
        return True   # DD.MM — first part is day
    if second_gt12:
        return False  # MM.DD — second part is day
    return True  # Default: European day-first

def build_times(df, t_cols):
    """
    Build a datetime Series from detected time columns.
    Merges separate date + time columns only when the first column
    contains date-only values (not already complete datetimes).
    """
    if not t_cols:
        return pd.Series([None] * len(df))
    if len(t_cols) >= 2 and not col_has_time_component(df[t_cols[0]]):
        return pd.to_datetime(
            df[t_cols[0]].astype(str) + ' ' + df[t_cols[1]].astype(str),
            dayfirst=True, errors='coerce'
        )
    return df[t_cols[0]].apply(parse_dt)

def read_csv_robust(filepath):
    """Try multiple separator + encoding combinations, return best DataFrame."""
    best = None
    for enc in ('utf-8-sig', 'utf-8', 'cp1250', 'latin-1'):
        for sep in (';', ',', '\t', '|'):
            try:
                df = pd.read_csv(filepath, sep=sep, encoding=enc, on_bad_lines='skip')
                if len(df.columns) >= 2 and (best is None or len(df.columns) > len(best.columns)):
                    best = df
            except Exception:
                pass
        if best is not None and len(best.columns) >= 3:
            break
    return best

def save_result(times, temps, hums, source_name, output_dir, suffix=''):
    """Standardize and save output Excel file."""
    data = {'Czas': times, 'Temperatura [°C]': temps}
    has_hum = bool(hums) and any(h is not None for h in hums)
    if has_hum:
        data['Wilgotność [%RH]'] = hums

    df = pd.DataFrame(data)
    df['Czas'] = pd.to_datetime(df['Czas'], errors='coerce')
    df['Temperatura [°C]'] = pd.to_numeric(df['Temperatura [°C]'], errors='coerce')
    if has_hum:
        df['Wilgotność [%RH]'] = pd.to_numeric(df['Wilgotność [%RH]'], errors='coerce')

    df = df.dropna(subset=['Czas', 'Temperatura [°C]'])
    df = df.drop_duplicates(subset=['Czas']).sort_values('Czas').reset_index(drop=True)

    if df.empty:
        print(f"    ⚠  Brak prawidłowych danych — plik pominięty")
        return df

    stem = Path(source_name).stem
    out_path = Path(output_dir) / f"{stem}{suffix}_wynik.xlsx"
    with pd.ExcelWriter(out_path, engine='openpyxl') as w:
        df.to_excel(w, index=False)

    cols_desc = 'Czas + Temp' + (' + Wilg' if has_hum else ' (brak wilgotności)')
    print(f"    ✓  {len(df)} pomiarów  [{cols_desc}]  →  {out_path.name}")
    return df

# ─── FORMAT DETECTION ─────────────────────────────────────────────────────────

def sniff_format(filepath):
    """Detect file format from first 1 KB of raw content."""
    filepath = Path(filepath)
    ext = filepath.suffix.lower()

    try:
        with open(filepath, 'rb') as f:
            raw = f.read(1024)
    except Exception:
        raw = b''

    text = raw.decode('latin-1', errors='replace')

    if b'[#I]Device Name' in raw or b'[#D]Measurements' in raw:
        return 'rotronic_hw4'

    if b'ALMEMO' in raw:
        return 'almemo'

    if ext == '.txt':
        # Comet/TFA text export (header: Nazwa:, Numer seryjny:, Start:, ...)
        if b'Numer seryjny' in raw or b'Nazwa:' in raw:
            return 'comet_txt'
        # Termostat xTHERM (log przez COM): naglowek zaczyna sie od 'czas odbioru'
        # i zawiera 2 czujniki temperatury (wewn. + zewn.) oraz wilgotnosc.
        if b'czas odbioru' in raw:
            return 'xtherm_com'
        return 'txt_generic'

    if ext == '.csv':
        # HOBO data logger
        if (b'LGR S/N' in raw or b'GMT+' in raw) and b'"Nr"' in raw:
            return 'hobo'
        # Aranet multi-sensor cloud logger
        if b'Aranet' in raw:
            return 'aranet'
        return 'csv_generic'

    if ext in ('.xls', '.xlsx'):
        return 'excel_inspect'

    return 'unknown'

# ─── FORMAT-SPECIFIC PARSERS ──────────────────────────────────────────────────

def parse_rotronic_hw4(filepath, output_dir):
    """Rotronic HW4 - INI-style text masquerading as .xls."""
    content = None
    for enc in ('utf-8', 'latin-1', 'cp1250'):
        try:
            with open(filepath, 'r', encoding=enc) as f:
                content = f.read()
            break
        except Exception:
            pass
    if content is None:
        raise ValueError("Cannot decode file")

    # Column headers from [#H] section
    m = re.search(r'\[#H\]Column Headers\r?\n(.+?)(?:\r?\n)', content)
    headers = [h.strip() for h in m.group(1).split('\t')] if m else []

    # Measurements from [#D] section
    m = re.search(r'\[#D\]Measurements\r?\n(.*)', content, re.DOTALL)
    if not m:
        raise ValueError("No [#D]Measurements section")

    rows = []
    for line in m.group(1).strip().split('\n'):
        parts = [p.strip() for p in line.strip().split('\t')]
        if len(parts) >= 3:
            rows.append(parts)
    if not rows:
        raise ValueError("No data rows in [#D] section")

    n_cols = max(len(r) for r in rows)
    while len(headers) < n_cols:
        headers.append(f'col{len(headers)}')
    df = pd.DataFrame(rows, columns=headers[:n_cols])

    # Combine Date + Time into single datetime
    if 'Date' in df.columns and 'Time' in df.columns:
        times = pd.to_datetime(df['Date'] + ' ' + df['Time'], dayfirst=True, errors='coerce')
    else:
        t_cols = find_cols(df.columns, TIME_KW)
        times = df[t_cols[0]].apply(parse_dt) if t_cols else pd.Series([None] * len(df))

    t_cols = find_cols(df.columns, TEMP_KW)
    h_cols = find_cols(df.columns, HUM_KW)

    if not t_cols:
        raise ValueError("No temperature column detected")

    temps = df[t_cols[0]].apply(clean_num)
    hums  = df[h_cols[0]].apply(clean_num).tolist() if h_cols else None

    save_result(times.tolist(), temps.tolist(), hums, filepath.name, output_dir)


def parse_almemo(filepath, output_dir):
    """ALMEMO - Excel with repeated metadata blocks; filter rows by datetime in col 0."""
    df_raw = pd.read_excel(filepath, header=None, engine='xlrd')

    # Find column mapping from DATE:/TIME: header rows
    header_rows = df_raw[df_raw.iloc[:, 0].astype(str).str.contains('DATE:', na=False)].index
    if len(header_rows) == 0:
        raise ValueError("No DATE: header found")

    hdr = df_raw.iloc[header_rows[0]]
    # Default: col2=Temp, col3=Humidity (ALMEMO standard layout)
    temp_idx = next(
        (i for i, v in enumerate(hdr) if pd.notna(v) and TEMP_KW.search(str(v))), 2
    )
    hum_idx = next(
        (i for i, v in enumerate(hdr) if pd.notna(v) and HUM_KW.search(str(v))), 3
    )

    times, temps, hums = [], [], []
    seen = set()

    for _, row in df_raw.iterrows():
        val0 = row.iloc[0]
        # Skip metadata/header rows
        if pd.isna(val0) or str(val0).strip().upper().startswith('DATE'):
            continue
        dt = parse_dt(val0)
        if dt is None:
            continue
        t = clean_num(row.iloc[temp_idx]) if len(row) > temp_idx else None
        if t is None:
            continue
        key = dt.isoformat()
        if key in seen:
            continue
        seen.add(key)
        h = clean_num(row.iloc[hum_idx]) if len(row) > hum_idx else None
        times.append(dt)
        temps.append(t)
        hums.append(h)

    save_result(times, temps, hums, filepath.name, output_dir)


def _parse_elogvis_sheet(df_raw, filepath, output_dir, sheet_tag=''):
    """Core ElogVis parser given a raw DataFrame."""
    header_idx = None
    for i, row in df_raw.iterrows():
        row_str = ' '.join(str(v) for v in row if pd.notna(v))
        if 'Data/Godzina' in row_str:
            header_idx = i
            break
        # Fallback: row with 'id' + data/time keyword + temp/hum keyword
        if (re.search(r'\bid\b', row_str, re.I)
                and TIME_KW.search(row_str)
                and (TEMP_KW.search(row_str) or HUM_KW.search(row_str))):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("No Data/Godzina header found")

    hdr  = df_raw.iloc[header_idx]
    data = df_raw.iloc[header_idx + 1:].reset_index(drop=True)

    # Col 0 = row id, Col 1 = datetime
    times = data.iloc[:, 1].apply(parse_dt)

    temp_idx = next(
        (i for i, v in enumerate(hdr) if pd.notna(v) and TEMP_KW.search(str(v))), 2
    )
    hum_idx = next(
        (i for i, v in enumerate(hdr) if pd.notna(v) and HUM_KW.search(str(v))), None
    )

    temps = data.iloc[:, temp_idx].apply(clean_num)
    hums  = data.iloc[:, hum_idx].apply(clean_num).tolist() if hum_idx is not None else None

    save_result(times.tolist(), temps.tolist(), hums, filepath.name, output_dir, sheet_tag)


def _parse_comet_tfa_sheet(df_raw, filepath, output_dir, sheet_tag=''):
    """Core Comet/TFA parser given a raw DataFrame (all data in col 0 as packed string)."""
    pat = re.compile(
        r'^\d+\s+(\d{4}[-/]\d{2}[-/]\d{2})\s+(\d{2}:\d{2}:\d{2})\s+'
        r'([\d.,]+)\s+([\d.,]+)'
    )
    times, temps, hums = [], [], []
    for _, row in df_raw.iterrows():
        s = str(row.iloc[0]).strip()
        m = pat.match(s)
        if not m:
            continue
        dt = parse_dt(f"{m.group(1)} {m.group(2)}")
        t  = clean_num(m.group(3))
        h  = clean_num(m.group(4))
        if dt and t is not None:
            times.append(dt)
            temps.append(t)
            hums.append(h)

    if not times:
        raise ValueError("No Comet/TFA pattern rows found")

    save_result(times, temps, hums, filepath.name, output_dir, sheet_tag)


def parse_hobo(filepath, output_dir):
    """HOBO CSV - skip title row, detect columns by keyword."""
    df = None
    for enc in ('utf-8-sig', 'utf-8', 'cp1250', 'latin-1'):
        try:
            candidate = pd.read_csv(filepath, encoding=enc, skiprows=1, on_bad_lines='skip')
            if len(candidate.columns) >= 2:
                df = candidate
                break
        except Exception:
            pass
    if df is None:
        raise ValueError("Cannot read HOBO file")

    t_cols    = find_cols(df.columns, TIME_KW)
    temp_cols = find_cols(df.columns, TEMP_KW)
    hum_cols  = find_cols(df.columns, HUM_KW)

    if not t_cols or not temp_cols:
        raise ValueError(f"Could not identify time/temp columns. Found: {list(df.columns)}")

    dayfirst = detect_dayfirst(df[t_cols[0]].dropna())
    times = pd.to_datetime(df[t_cols[0]], dayfirst=dayfirst, errors='coerce')
    temps = df[temp_cols[0]].apply(clean_num)
    hums  = df[hum_cols[0]].apply(clean_num).tolist() if hum_cols else None

    save_result(times.tolist(), temps.tolist(), hums, filepath.name, output_dir)


def parse_aranet(filepath, output_dir):
    """
    Aranet multi-sensor CSV - one output file per sensor.
    Columns: Time | SENSOR_ID ... Temperature (°C) [SENSOR_ID] | ... Humidity (%) [SENSOR_ID]
    """
    df = None
    for enc in ('utf-8-sig', 'utf-8', 'latin-1'):
        try:
            df = pd.read_csv(filepath, encoding=enc, on_bad_lines='skip')
            break
        except Exception:
            pass
    if df is None:
        raise ValueError("Cannot read Aranet file")

    time_col  = df.columns[0]
    temp_cols = find_cols(df.columns, TEMP_KW)
    hum_cols  = find_cols(df.columns, HUM_KW)

    if not temp_cols:
        raise ValueError("No temperature columns found in Aranet file")

    # Extract unique sensor IDs from column bracket suffixes: [...Temperature... [SENSOR_ID]]
    sensor_ids = []
    for col in temp_cols:
        m = re.search(r'\[([^\]]+)\]$', col)
        if m:
            sid = m.group(1)
            if sid not in sensor_ids:
                sensor_ids.append(sid)

    if not sensor_ids:
        # No sensor IDs in brackets — treat all temp cols as one combined sensor
        sensor_ids = [None]

    for sid in sensor_ids:
        if sid:
            tc_list = [c for c in temp_cols if c.endswith(f'[{sid}]')]
            hc_list = [c for c in hum_cols  if c.endswith(f'[{sid}]')]
            tag = f'_{sid}'
        else:
            tc_list = temp_cols
            hc_list = hum_cols
            tag = ''

        if not tc_list:
            continue

        times, temps, hums = [], [], []
        for _, row in df.iterrows():
            dt = parse_dt(row[time_col])
            if dt is None:
                continue
            t = clean_num(row[tc_list[0]])
            if t is None:
                continue
            h = clean_num(row[hc_list[0]]) if hc_list else None
            times.append(dt)
            temps.append(t)
            hums.append(h)

        if not times:
            print(f"    ⚠  Sensor {sid}: brak danych")
            continue

        save_result(times, temps, hums, filepath.name, output_dir, tag)


def parse_csv_generic(filepath, output_dir):
    """Generic CSV - smart separator/encoding detection + keyword column mapping."""
    df = read_csv_robust(filepath)
    if df is None:
        raise ValueError("Cannot parse CSV file with any separator/encoding")

    t_cols    = find_cols(df.columns, TIME_KW)
    temp_cols = find_cols(df.columns, TEMP_KW)
    hum_cols  = find_cols(df.columns, HUM_KW)

    if not t_cols:
        raise ValueError(
            f"No time column detected. Available columns: {list(df.columns)}"
        )
    times = build_times(df, t_cols)

    # Fallback for temperature/humidity: use numeric columns after the time columns
    if not temp_cols:
        last_t_pos = max(df.columns.get_loc(c) for c in t_cols) if t_cols else -1
        numeric_cols = [
            c for c in df.columns
            if df.columns.get_loc(c) > last_t_pos
            and pd.to_numeric(df[c], errors='coerce').notna().sum() > len(df) * 0.4
        ]
        if not numeric_cols:
            raise ValueError(
                f"No temperature column detected. Columns: {list(df.columns)}"
            )
        temp_cols = [numeric_cols[0]]
        if len(numeric_cols) >= 2:
            hum_cols = [numeric_cols[1]]
        print(f"    ⚠  Brak słów kluczowych temp/wilg — użyto kolumn pozycyjnych: "
              f"T={temp_cols}, RH={hum_cols or '—'}")

    temps = df[temp_cols[0]].apply(clean_num)
    hums  = df[hum_cols[0]].apply(clean_num).tolist() if hum_cols else None

    save_result(times.tolist(), temps.tolist(), hums, filepath.name, output_dir)


def parse_excel_inspect(filepath, output_dir):
    """Excel file: auto-detect sub-format per sheet."""
    ext    = filepath.suffix.lower()
    engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'

    try:
        xl = pd.ExcelFile(filepath, engine=engine)
    except Exception as e:
        raise ValueError(f"Cannot open Excel: {e}")

    multi = len(xl.sheet_names) > 1

    for sheet in xl.sheet_names:
        tag = f'_{sheet}' if multi else ''
        print(f"    sheet: '{sheet}'")

        try:
            df_raw = pd.read_excel(filepath, header=None, engine=engine, sheet_name=sheet)
        except Exception as e:
            print(f"      ✗ Nie można odczytać arkusza: {e}")
            continue

        full_text = df_raw.to_string()

        # ── ALMEMO (Excel with repeated metadata blocks) ─────────────────────
        has_almemo = 'ALMEMO' in full_text
        has_date_hdr = df_raw.iloc[:, 0].astype(str).str.contains('DATE:', na=False).any()
        if has_almemo and has_date_hdr:
            print(f"      → ALMEMO format")
            try:
                parse_almemo(filepath, output_dir)
                return  # One ALMEMO file = one output, no more sheets to process
            except Exception as e:
                print(f"      ✗ ALMEMO parse error: {e}")
                if DEBUG:
                    traceback.print_exc()

        # ── ElogVis (Polish logger) ──────────────────────────────────────────
        if 'Data/Godzina' in full_text:
            print(f"      → ElogVis PL format")
            try:
                _parse_elogvis_sheet(df_raw, filepath, output_dir, tag)
                continue
            except Exception as e:
                print(f"      ✗ ElogVis parse error: {e}")
                if DEBUG:
                    traceback.print_exc()

        # ── Comet / TFA (single-column packed strings) ──────────────────────
        comet_re = re.compile(r'^\d{3,4}\s+\d{4}[-/]\d{2}[-/]\d{2}\s+\d{2}:\d{2}:\d{2}')
        sample   = df_raw.iloc[:, 0].dropna().astype(str)
        if sample.str.match(comet_re).sum() > 3:
            print(f"      → Comet/TFA packed format")
            try:
                _parse_comet_tfa_sheet(df_raw, filepath, output_dir, tag)
                continue
            except Exception as e:
                print(f"      ✗ Comet/TFA parse error: {e}")
                if DEBUG:
                    traceback.print_exc()

        # ── Generic tabular Excel — search for a proper header row ──────────
        header_idx = None
        for i, row in df_raw.iterrows():
            row_str = ' '.join(str(v) for v in row if pd.notna(v))
            if TIME_KW.search(row_str) and (TEMP_KW.search(row_str) or HUM_KW.search(row_str)):
                header_idx = i
                break

        if header_idx is not None:
            print(f"      → Generic tabular Excel (header @ row {header_idx})")
            try:
                df = pd.read_excel(
                    filepath, header=int(header_idx), engine=engine, sheet_name=sheet
                )
                t_cols    = find_cols(df.columns, TIME_KW)
                temp_cols = find_cols(df.columns, TEMP_KW)
                hum_cols  = find_cols(df.columns, HUM_KW)

                if not t_cols:
                    print(f"      ✗ Brak kolumny czasu w arkuszu '{sheet}'")
                    continue

                times = build_times(df, t_cols)

                if not temp_cols:
                    # Positional fallback
                    last_t_pos = max(df.columns.get_loc(c) for c in t_cols) if t_cols else -1
                    numeric_cols = [
                        c for c in df.columns
                        if df.columns.get_loc(c) > last_t_pos
                        and pd.to_numeric(df[c], errors='coerce').notna().sum() > len(df) * 0.4
                    ]
                    if not numeric_cols:
                        print(f"      ✗ Brak kolumny temperatury w arkuszu '{sheet}'")
                        continue
                    temp_cols = [numeric_cols[0]]
                    if len(numeric_cols) >= 2:
                        hum_cols = [numeric_cols[1]]
                    print(f"      ⚠  Kolumny pozycyjne: T={temp_cols}, RH={hum_cols or '—'}")

                temps = df[temp_cols[0]].apply(clean_num)
                hums  = df[hum_cols[0]].apply(clean_num).tolist() if hum_cols else None
                save_result(times.tolist(), temps.tolist(), hums, filepath.name, output_dir, tag)
                continue
            except Exception as e:
                print(f"      ✗ Generic Excel parse error: {e}")
                if DEBUG:
                    traceback.print_exc()

        print(f"      ✗ Nierozpoznany format arkusza '{sheet}' — pomijam")

def parse_comet_txt(filepath, output_dir):
    """
    Comet/TFA text export (.txt) — two sub-formats after a metadata header block:

    Format A  (tab-separated):
      1\\t20.02.2026\\t11:16:00\\t24.08\\t44.14
      2\\t20.02.2026\\t11:17:00\\t24.17\\t41.81

    Format B  (space-padded):
      0001      20.02.2026    11:15:15           23.91     42.63
      0002      20.02.2026    11:16:15           24.08     39.40

    Both share the same regex — flexible whitespace between fields.
    Date is always DD.MM.YYYY (European), time HH:MM:SS.
    """
    content = None
    for enc in ('utf-8', 'utf-8-sig', 'cp1250', 'latin-1'):
        try:
            with open(filepath, 'r', encoding=enc) as f:
                content = f.read()
            break
        except Exception:
            pass
    if content is None:
        raise ValueError("Cannot decode .txt file")

    # Flexible pattern — handles both date formats:
    #   DD.MM.YYYY  (e.g. 20.02.2026)
    #   YYYY-MM-DD  (e.g. 2026-06-24)
    pat = re.compile(
        r'^\d+[\s\t]+(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})[\s\t]+(\d{2}:\d{2}:\d{2})[\s\t]+'
        r'([\d.,]+)(?:[\s\t]+([\d.,]+))?',
        re.MULTILINE
    )

    times, temps, hums = [], [], []
    for m in pat.finditer(content):
        dt = parse_dt(f"{m.group(1)} {m.group(2)}")
        t  = clean_num(m.group(3))
        h  = clean_num(m.group(4)) if m.group(4) else None
        if dt and t is not None:
            times.append(dt)
            temps.append(t)
            hums.append(h)

    if not times:
        raise ValueError("No data rows found in .txt file")

    save_result(times, temps, hums, filepath.name, output_dir)


def _best_numeric_col(df, cols):
    """
    Spośród podanych kolumn zwróć (nazwa, liczba_wartości) tej z największą liczbą
    prawidłowych wartości liczbowych. Komórki 'NAN'/puste się nie liczą.
    Pozwala wybrać kolumnę z realnymi danymi, gdy część kandydatów jest pusta
    (np. EMC_currTemp=NAN → wybierze EMC_analogTemp).
    """
    best, best_count = None, 0
    for c in cols:
        cnt = int(pd.to_numeric(
            df[c].map(lambda v: str(v).replace(',', '.')), errors='coerce'
        ).notna().sum())
        if cnt > best_count:
            best, best_count = c, cnt
    return best, best_count


def parse_txt_generic(filepath, output_dir):
    """
    Generyczny .txt z separatorem (np. log termostatu xTHERM przez COM:
    nagłówek 'czas odbioru;...;EMC_currTemp;EMC_analogTemp;EMC_humidity;...').

    Wykrywa kolumny czasu/temperatury/wilgotności po słowach kluczowych, a wśród
    kandydatów temp/wilg wybiera tę z REALNYMI danymi (pomija puste/NAN — np. gdy
    EMC_currTemp=NAN, użyje EMC_analogTemp; gdy wilgotność same NAN — pomija ją).
    """
    df = read_csv_robust(filepath)
    if df is None:
        raise ValueError("Cannot parse .txt file with any separator/encoding")

    t_cols    = find_cols(df.columns, TIME_KW)
    temp_cols = find_cols(df.columns, TEMP_KW)
    hum_cols  = find_cols(df.columns, HUM_KW)

    if not t_cols:
        raise ValueError(f"No time column detected. Columns: {list(df.columns)}")
    times = build_times(df, t_cols)

    # Temperatura: kolumna z najlepszym pokryciem danych spośród dopasowanych słowem.
    temp_col, tcount = _best_numeric_col(df, temp_cols)
    if temp_col is None or tcount == 0:
        # Fallback pozycyjny: pierwsza kolumna liczbowa po kolumnach czasu.
        last_t_pos = max(df.columns.get_loc(c) for c in t_cols)
        numeric_cols = [
            c for c in df.columns
            if df.columns.get_loc(c) > last_t_pos
            and pd.to_numeric(df[c], errors='coerce').notna().sum() > len(df) * 0.4
        ]
        if not numeric_cols:
            raise ValueError(
                f"No temperature column with data. Columns: {list(df.columns)}"
            )
        temp_col = numeric_cols[0]
        print(f"    ⚠  Brak kolumny temp ze słowem kluczowym — użyto pozycyjnej: {temp_col}")

    # Wilgotność: tylko jeśli któraś dopasowana kolumna ma realne dane.
    hum_col, hcount = _best_numeric_col(df, hum_cols)
    if hcount == 0:
        hum_col = None

    temps = df[temp_col].apply(clean_num)
    hums  = df[hum_col].apply(clean_num).tolist() if hum_col else None

    save_result(times.tolist(), temps.tolist(), hums, filepath.name, output_dir)


def _find_col(columns, *names):
    """Znajdź kolumnę po dokładnej nazwie (lub zawieraniu), bez względu na wielkość liter."""
    norm = [(str(c).strip().lower(), c) for c in columns]
    for name in names:                       # najpierw dokladne dopasowanie
        nl = name.lower()
        for low, orig in norm:
            if low == nl:
                return orig
    for name in names:                       # potem zawieranie
        nl = name.lower()
        for low, orig in norm:
            if nl in low:
                return orig
    return None


def _save_xtherm(times, t_wewn, t_zewn, hums, source_name, output_dir, suffix=''):
    """
    Zapis wyniku termostatu xTHERM: dwie kolumny temperatury (wewn. + zewn.)
    plus opcjonalna wilgotność. Obie kolumny temperatury zostaja zawsze (cecha
    tego typu); wilgotnosc tylko gdy ma dane.
    """
    L_INT, L_EXT, L_HUM = 'Temperatura wewn. [°C]', 'Temperatura zewn. [°C]', 'Wilgotność [%RH]'
    n = len(times)
    def _col(s):
        return list(s) if s is not None else [None] * n

    df = pd.DataFrame({'Czas': list(times), L_INT: _col(t_wewn), L_EXT: _col(t_zewn), L_HUM: _col(hums)})
    df['Czas'] = pd.to_datetime(df['Czas'], errors='coerce')
    for c in (L_INT, L_EXT, L_HUM):
        df[c] = pd.to_numeric(df[c], errors='coerce')

    # Wymagamy czasu i przynajmniej jednej temperatury w wierszu.
    df = df.dropna(subset=['Czas'])
    df = df[df[[L_INT, L_EXT]].notna().any(axis=1)]
    # Wilgotnosc usuwamy tylko gdy calkowicie pusta (zachowujemy 2 kolumny temp.).
    if df[L_HUM].notna().sum() == 0:
        df = df.drop(columns=[L_HUM])

    df = df.drop_duplicates(subset=['Czas']).sort_values('Czas').reset_index(drop=True)
    if df.empty:
        print(f"    ⚠  Brak prawidłowych danych — plik pominięty")
        return df

    stem = Path(source_name).stem
    out_path = Path(output_dir) / f"{stem}{suffix}_wynik.xlsx"
    with pd.ExcelWriter(out_path, engine='openpyxl') as w:
        df.to_excel(w, index=False)

    cols_desc = ' + '.join(c for c in df.columns if c != 'Czas')
    print(f"    ✓  {len(df)} pomiarów  [Czas + {cols_desc}]  →  {out_path.name}")
    return df


def parse_xtherm_com(filepath, output_dir):
    """
    Termostat xTHERM (log przez COM, naglowek 'czas odbioru;...').
    Czujnik WEWNĘTRZNY: temperatura = EMC_currTemp, wilgotnosc = EMC_humidity.
    Czujnik ZEWNĘTRZNY: temperatura = EMC_analogTemp.
    Zapis: Czas | Temperatura wewn. | Temperatura zewn. | Wilgotność.
    Gdy plik nie ma tych kolumn — przechodzi do zwyklego parse_txt_generic.
    """
    df = read_csv_robust(filepath)
    if df is None:
        raise ValueError("Cannot parse .txt file with any separator/encoding")

    t_cols  = find_cols(df.columns, TIME_KW)
    col_int = _find_col(df.columns, 'EMC_currTemp')
    col_ext = _find_col(df.columns, 'EMC_analogTemp')
    col_hum = _find_col(df.columns, 'EMC_humidity')

    if not t_cols or (col_int is None and col_ext is None):
        # Nie wyglada jak xTHERM — uzyj zwyklego parsera generycznego.
        return parse_txt_generic(filepath, output_dir)

    times  = build_times(df, t_cols)
    t_wewn = df[col_int].apply(clean_num) if col_int else None
    t_zewn = df[col_ext].apply(clean_num) if col_ext else None
    hums   = df[col_hum].apply(clean_num) if col_hum else None

    _save_xtherm(times, t_wewn, t_zewn, hums, filepath.name, output_dir)


# ─── DISPATCH TABLE ───────────────────────────────────────────────────────────

PARSERS = {
    'rotronic_hw4':  parse_rotronic_hw4,
    'almemo':        parse_almemo,
    'hobo':          parse_hobo,
    'aranet':        parse_aranet,
    'csv_generic':   parse_csv_generic,
    'excel_inspect': parse_excel_inspect,
    'comet_txt':     parse_comet_txt,
    'xtherm_com':    parse_xtherm_com,
    'txt_generic':   parse_txt_generic,
}

# ─── ENTRY POINT ─────────────────────────────────────────────────────────────

def process_file(filepath, output_dir):
    filepath = Path(filepath)
    fmt      = sniff_format(filepath)
    print(f"\n  [{filepath.name}]")
    print(f"    Format: {fmt}")

    parser = PARSERS.get(fmt)
    if parser is None:
        print(f"    ✗ Format nieobsługiwany: '{fmt}'")
        return

    try:
        parser(filepath, output_dir)
    except Exception as e:
        print(f"    ✗ Błąd: {e}")
        if DEBUG:
            traceback.print_exc()


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]

    if args:
        target = Path(args[0])
        if target.is_file():
            output_dir = DEFAULT_OUTPUT
            output_dir.mkdir(exist_ok=True)
            print(f"Przetwarzam pojedynczy plik: {target}")
            print(f"Wyniki → '{output_dir}/'")
            process_file(target, output_dir)
            return
        if target.is_dir():
            input_dir = target
        else:
            print(f"Błąd: '{target}' nie istnieje")
            sys.exit(1)
    else:
        input_dir = DEFAULT_INPUT

    output_dir = DEFAULT_OUTPUT
    output_dir.mkdir(exist_ok=True)

    files = sorted(
        f for f in input_dir.iterdir()
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXT
    )

    print("=" * 60)
    print(f"  Folder wejściowy : {input_dir}/")
    print(f"  Folder wyjściowy : {output_dir}/")
    print(f"  Znaleziono plików: {len(files)}")
    print("=" * 60)

    for f in files:
        process_file(f, output_dir)

    # Summary
    out_files = list(output_dir.glob('*_wynik.xlsx'))
    print("\n" + "=" * 60)
    print(f"  Gotowe! Zapisano {len(out_files)} pliku(ów) wynikowego(ych).")
    print(f"  Sprawdź folder: {output_dir}/")
    print("=" * 60)


if __name__ == '__main__':
    main()
