# -*- coding: utf-8 -*-
"""
generuj_obserwacje.py

Tworzy arkusz obserwacji CC na podstawie pliku TXT z danymi pomiarowymi.

1. Nazwa pliku TXT  → numer pomiaru (po dacie i czasie) → nazwa kopii szablonu
2. Kopia szablonu: xxx → numer pomiaru
3. Parsowanie pliku TXT:
       A8  – czujnik wzorcowy (tekst po ':')
       A15+ – dane pomiarowe (sep: Tab / Średnik / Przecinek), kolumny A:L
4. Dane A15:L(ostatni) → kopia szablonu A2:L(ostatni)
5. J1  – podmiana wzorca "Pt100-XX" na czujnik z pliku
6. N92 – aktualna data
7. O92 – podpis
8. Analiza stabilności:
       - szuka okna po 1.5h stabilnym B+C (rozgrzewka)
       - filtruje wiersze wg kryteriów K (punkt rosy) i L (temperatura)
       - wybiera 5-minutowe okno z najmniejszym rozrzutem K
       - koloruje okno na zielono (#E2EFDA)
       - wybiera 5 wierszy reprezentacyjnych (1/min, blisko granicy minutowej)
         i oznacza je ciemniejszą zielenią (#A9D08E) + pogrubienie
"""

import os
import re
import bisect
import shutil
import datetime
import statistics
from copy import copy as _copy_obj
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# =============================================================================
# KONFIGURACJA  ← edytuj tutaj przed uruchomieniem
# =============================================================================

FOLDER           = r"C:\Users\artisom.azhdzer\Desktop\Script protokoł - arkusz CC"
TXT_FILENAME     = "2026-06-29 06.46_00001.txt"
TEMPLATE         = "xxx_LA_TH_2026 - obserwacje CC.xlsx"
CC04_TEMPLATE    = "szablon_LA_TH_2026 - obserwacje.xlsx"
PROTOKOL_CC_TEMPLATE   = "xxx_LA_TH_2026 - protokół CC.xlsx"
PROTOKOL_CC04_TEMPLATE = "xxx_LA_TH_2026 - protokół CC-04.xlsx"
PODPIS           = "Artsiom Azhdzer"
STABILIZACJA_MIN  = datetime.timedelta(hours=1, minutes=45)   # czas rozgrzewki przed analizą
SUSZENIE_T_ZAKRES = (21.0, 27.0)   # zakres T [°C] charakterystyczny dla suszenia
SUSZENIE_RH_MAX   = 50.0           # max RH [%] — ponizej tej wartosci punkt moze byc suszeniem

# Folder z zunifikowanymi plikami wynikow (z analizuj_excele.py)
WYNIKI_FOLDER = os.path.join(FOLDER, "wyniki")

# Kolumna startowa dla danych srodowiskowych z wynikow w Strona 3
#   CC:   Q = 17  (temperatura=Q, wilgotnosc=R)
#   CC04: S = 19  (temperatura=S, wilgotnosc=T)
WYNIKI_START_COL_CC   = 17   # Q
WYNIKI_START_COL_CC04 = 19   # S

# Maksymalna tolerancja przy dopasowaniu timestampow [minuty]
WYNIKI_TOLERANCJA_MIN = 30

# Filtr zgodnosci NASTAWY z ODCZYTEM komory:
#   B = Tzadana  vs  D = Todczytana   (temperatura)
#   C = RHzadana vs  E = RHodczytana  (wilgotnosc)
# Jesli odczyt komory rozni sie od nastawy o wiecej niz prog [%], segment (punkt)
# jest pomijany — komora nie osiagnela zadanych warunkow.
FILTR_NASTAWA_ODCZYT = True
MAX_ROZNICA_PROCENT  = 10.0   # dozwolona wzgledna roznica |nastawa-odczyt|/nastawa * 100

# =============================================================================

FILL_LIGHT = PatternFill(fill_type='solid', fgColor='E2EFDA')  # jasna zieleń
FILL_DARK  = PatternFill(fill_type='solid', fgColor='A9D08E')  # ciemna zieleń


def parse_measurement_id(filename: str) -> str:
    """Zwraca część nazwy pliku po 'YYYY-MM-DD HH.MM ' (np. '119_141_147_149_151_164')."""
    base = os.path.splitext(filename)[0]
    m = re.search(r'\d{4}-\d{2}-\d{2} \d{2}\.\d{2}[ _](.+)', base)
    if not m:
        raise ValueError(f"Nierozpoznany format nazwy pliku: {filename!r}")
    return m.group(1)


def open_txt(path: str):
    """Otwiera plik TXT próbując kolejnych kodowań."""
    for enc in ('cp1250', 'utf-8-sig', 'utf-8', 'latin-1'):
        try:
            with open(path, encoding=enc) as f:
                return f.readlines()
        except UnicodeDecodeError:
            continue
    with open(path, encoding='utf-8', errors='replace') as f:
        return f.readlines()


def parse_txt(path: str):
    """
    Zwraca (sensor_name, rows):
      sensor_name – tekst po ':' z linii 8 (indeks 7)
      rows        – lista list[str] z linii 15+ (indeks 14+), pierwsze 12 kolumn (A:L)
    """
    lines = open_txt(path)

    # Linia 8 (indeks 7): "Czujnik wzorcowy: Pt100-11"
    if len(lines) > 7:
        sensor_line = lines[7].strip()
        sensor_name = sensor_line.split(':', 1)[1].strip() if ':' in sensor_line else sensor_line
    else:
        sensor_name = ''

    # Linia 15 i niżej (indeks 14+): dane pomiarowe
    rows = []
    for line in lines[14:]:
        line = line.strip()
        if not line:
            continue
        parts = re.split(r'[;\t,]', line)
        row = parts[:12]
        while len(row) < 12:
            row.append('')
        rows.append(row)

    return sensor_name, rows


def detect_file_type(lines) -> str:
    """Zwraca 'CC04' jesli linia 5 zawiera 'CC-04', inaczej 'CC'."""
    if len(lines) > 4 and 'CC-04' in lines[4]:
        return 'CC04'
    return 'CC'


def parse_txt_cc04(path: str):
    """
    Format CC-04:
      Linie 8-11 (ind. 7-10) : 4 czujniki, np. 'Czujnik wzorcowy: Pt100-09; Wejscie...'
      Linia 17   (ind. 16)   : naglowek kolumn
      Linia 18+  (ind. 17+)  : dane pomiarowe, kolumny A:U (21 kol.)
    Zwraca (sensor_names: list[str], rows: list[list[str]]).
    """
    lines = open_txt(path)

    sensor_names = []
    for line_idx in range(7, 11):
        if len(lines) > line_idx:
            m = re.search(r':\s*(Pt100-\d+)', lines[line_idx])
            sensor_names.append(m.group(1) if m else '')
    while len(sensor_names) < 4:
        sensor_names.append('')

    rows = []
    for line in lines[17:]:
        line = line.strip()
        if not line:
            continue
        parts = re.split(r'[;\t,]', line)
        row = parts[:21]
        while len(row) < 21:
            row.append('')
        rows.append(row)

    return sensor_names, rows


_ILLEGAL_CHARS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

def to_value(s: str):
    """Konwertuje string na odpowiedni typ Pythona (dla zapisu do Excela).
    Usuwa znaki kontrolne niedozwolone przez openpyxl."""
    s = _ILLEGAL_CHARS_RE.sub('', str(s).strip())
    if not s:
        return None
    try:
        return datetime.datetime.strptime(s, '%Y-%m-%d %H:%M:%S')
    except ValueError:
        pass
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s.replace(',', '.'))
    except ValueError:
        pass
    return s


# =============================================================================
# ANALIZA I KOLOROWANIE
# =============================================================================

def _s_to_float(s):
    """Konwertuje string na float; None jeśli niemożliwe (np. 'brak')."""
    try:
        return float(str(s).strip().replace(',', '.'))
    except (ValueError, AttributeError):
        return None


_DT_FORMATS = ('%Y-%m-%d %H:%M:%S', '%d.%m.%Y %H:%M:%S', '%d.%m.%Y %H:%M')

def _s_to_dt(s):
    """Parsuje string daty na datetime; obsługuje kilka formatów i podwójne spacje."""
    text = re.sub(r'\s+', ' ', str(s).strip())
    for fmt in _DT_FORMATS:
        try:
            return datetime.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def get_k_threshold(temp, rh):
    """
    Zwraca maksymalne dozwolone K (roztdp) dla zadanych T i RH.
      0.2  — warunki dobrze kontrolowane (srodkowa wilgotnosc w typowym zakresie T),
      0.45 — wartosc MAKSYMALNA / domyslna dla wszystkich pozostalych przypadkow
             (skrajna wilgotnosc, wysoka/nietypowa temperatura itd.),
      None — tylko gdy brak danych T lub RH.

    Zaostrzony prog 0.2 obowiazuje dla:
      ~10 C (5-15)  : 40 < RH < 75
      ~23 C (15-30) : 30 <= RH <= 75
      ~35 C (30-45) : 40 <= RH <= 75
    """
    if temp is None or rh is None:
        return None

    if 5 <= temp <= 15:          # ~10 C
        if 40 < rh < 75:
            return 0.2
    elif 15 < temp <= 30:        # ~23 C
        if 30 <= rh <= 75:
            return 0.2
    elif 30 < temp <= 45:        # ~35 C
        if 40 <= rh <= 75:
            return 0.2

    return 0.45   # domyslnie — jesli nic innego nie pasuje (wartosc maksymalna)


def _prepare_data(rows):
    """
    Przetwarza wiersze str na listę krotek (dt, temp, rh, k, l, d, e).
    Indeks na liście = Excel_row - 2.
    d, e = odczyt komory (Todczytana, RHodczytana) — do filtru zgodnosci.
    """
    result = []
    for row in rows:
        result.append((
            _s_to_dt(row[0]),       # A – czas
            _s_to_float(row[1]),    # B – Tzadana
            _s_to_float(row[2]),    # C – RHzadana
            _s_to_float(row[10]),   # K – roztdp(15min)
            _s_to_float(row[11]),   # L – roztempCh001
            _s_to_float(row[3]),    # D – Todczytana (odczyt komory)
            _s_to_float(row[4]),    # E – RHodczytana (odczyt komory)
        ))
    return result


def _prepare_data_cc04(rows):
    """
    CC-04: krotka (dt, temp, rh, k_dew, l1, l2, l3, l4, d, e)
      A=0  czas,  B=1 Tzadana,  C=2 RHzadana,
      D=3  Todczytana,  E=4 RHodczytana  (odczyt komory — filtr zgodnosci),
      Q=16 roztdp (punkt rosy),
      R=17 roztemp Ch1,  S=18 Ch2,  T=19 Ch3,  U=20 Ch4
    """
    result = []
    for row in rows:
        result.append((
            _s_to_dt(row[0]),
            _s_to_float(row[1]),
            _s_to_float(row[2]),
            _s_to_float(row[16]),   # Q – roztdp
            _s_to_float(row[17]),   # R – roztemp Ch1
            _s_to_float(row[18]),   # S – roztemp Ch2
            _s_to_float(row[19]),   # T – roztemp Ch3
            _s_to_float(row[20]),   # U – roztemp Ch4
            _s_to_float(row[3]),    # D – Todczytana (odczyt komory)
            _s_to_float(row[4]),    # E – RHodczytana (odczyt komory)
        ))
    return result


def _find_all_analysis_windows(data, min_stable):
    """
    Dla każdego segmentu stabilnego B+C trwającego >= min_stable szuka okna analizy:
      - start: pierwszy wiersz po upływie min_stable od początku segmentu
      - end  : pierwszy wiersz gdzie B lub C zmienia wartość (wyłącznie)

    Zwraca listę krotek (start_idx, end_idx).
    """
    windows = []
    n = len(data)
    i = 0

    while i < n:
        seg_b, seg_c, seg_dt0 = data[i][1], data[i][2], data[i][0]

        if seg_b is None or seg_c is None or seg_dt0 is None:
            i += 1
            continue

        # Koniec bieżącego segmentu stabilnego
        j = i + 1
        while j < n:
            b, c = data[j][1], data[j][2]
            if b != seg_b or c != seg_c:
                break
            j += 1
        # Segment stabilny: data[i : j]

        seg_end_dt = data[j - 1][0] or seg_dt0
        if seg_end_dt - seg_dt0 >= min_stable:
            # Segment wystarczająco długi — szukamy startu po czasie stabilizacji
            for ki in range(i, j):
                if data[ki][0] and data[ki][0] - seg_dt0 >= min_stable:
                    windows.append((ki, j))
                    break

        i = j  # następny segment

    return windows


def _find_best_minute_reps(data, valid_indices, start_idx, end_idx, score_fn):
    """
    Bezposrednia optymalizacja po wyniku score_fn.

    score_fn(i) -> float|None: funkcja zwracajaca metric do minimalizacji dla wiersza i.
    Przyklady:
      CC normalny:     lambda i: data[i][3]            (K – roztdp)
      CC temp-only:    lambda i: data[i][4]            (L – roztemp)
      CC-04 normalny:  lambda i: data[i][3]            (Q – roztdp)
      CC-04 temp-only: lambda i: max(data[i][4:8])    (max 4xL)

    1. Dla kazdej pelnej minuty szuka najblizszego waznego wiersza (max 30s od :00).
    2. Wybiera 5 kolejnych minut z minimalna srednia score_fn — przy rownym wyniku
       preferuje pozniejsze okno (porownanie <=).

    Zwraca liste 5 indeksow (0-based do `data`) lub None.
    """
    MAX_SEC = 30
    ONE_MIN = datetime.timedelta(minutes=1)

    if not valid_indices:
        return None

    dt_start = data[start_idx][0]
    dt_end   = data[end_idx - 1][0]
    if dt_start is None or dt_end is None:
        return None

    # Posortowane wazne wiersze: (datetime, score, data_idx)
    vrows = sorted(
        [(data[i][0], score_fn(i), i) for i in valid_indices
         if data[i][0] is not None and score_fn(i) is not None],
        key=lambda x: x[0]
    )
    vtimes = [vr[0] for vr in vrows]

    # Dla kazdej pelnej minuty szukamy najblizszego waznego wiersza
    cur = dt_start.replace(second=0, microsecond=0)
    minute_reps = {}   # minute_dt -> (score, data_idx)

    while cur <= dt_end + ONE_MIN:
        lo = bisect.bisect_left(vtimes,  cur - datetime.timedelta(seconds=MAX_SEC))
        hi = bisect.bisect_right(vtimes, cur + datetime.timedelta(seconds=MAX_SEC))
        if lo < hi:
            best = min(vrows[lo:hi], key=lambda x: abs((x[0] - cur).total_seconds()))
            minute_reps[cur] = (best[1], best[2])
        cur += ONE_MIN

    # Wybieramy 5 kolejnych minut z minimalna srednia score
    all_minutes = sorted(minute_reps.keys())
    if len(all_minutes) < 5:
        return None

    best_mean = float('inf')
    best_5    = None

    for pos in range(len(all_minutes) - 4):
        group = [all_minutes[pos + j] for j in range(5)]
        if any(group[j + 1] - group[j] != ONE_MIN for j in range(4)):
            continue
        vals = [minute_reps[m][0] for m in group]
        mean_val = sum(vals) / 5
        if mean_val <= best_mean:
            best_mean = mean_val
            best_5    = [minute_reps[m][1] for m in group]

    return best_5


def _apply_fill(ws, excel_row, fill, n_cols=12):
    for col in range(1, n_cols + 1):
        ws.cell(row=excel_row, column=col).fill = fill


def _apply_bold(ws, excel_row, n_cols=12):
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=excel_row, column=col)
        f = cell.font
        cell.font = Font(
            name=f.name, size=f.size, italic=f.italic,
            underline=f.underline, strike=f.strike, color=f.color,
            bold=True,
        )


def _rozne_procent(nastawa, odczyt, prog_proc):
    """
    True, gdy odczyt komory rozni sie od nastawy o wiecej niz prog_proc [%].
    Brak ktorejkolwiek wartosci -> False (nie filtrujemy). Nastawa ~0 -> porownanie
    bezwzgledne (prog traktowany jako wartosc), by uniknac dzielenia przez zero.
    """
    if nastawa is None or odczyt is None:
        return False
    baza = abs(nastawa)
    if baza < 1e-9:
        return abs(odczyt) > prog_proc
    return abs(nastawa - odczyt) / baza * 100.0 > prog_proc


def _process_segment(ws, data, start_idx, end_idx, seg_num, file_type='CC'):
    """Analizuje jeden segment i koloruje go. Zwraca liste 5 indeksow (0-based) lub None."""
    t0 = data[start_idx][0]
    t1 = data[end_idx - 1][0]
    b  = data[start_idx][1]
    c  = data[start_idx][2]
    print(f"  Segment {seg_num}: Excel row {2+start_idx}–{2+end_idx-1}  "
          f"T={b} RH={c}  ({t0} – {t1})")

    is_cc04   = (file_type == 'CC04')
    temp_only = (c is not None and c == 0.0)

    # Filtr zgodnosci nastawy z odczytem komory (B/D temp, C/E wilgotnosc).
    # Bierzemy medianowy odczyt z calego segmentu i porownujemy z nastawa.
    # Gdy rozbieznosc > MAX_ROZNICA_PROCENT — komora nie osiagnela warunkow, pomijamy.
    if FILTR_NASTAWA_ODCZYT:
        d_idx, e_idx = (8, 9) if is_cc04 else (5, 6)
        d_vals = [data[i][d_idx] for i in range(start_idx, end_idx) if data[i][d_idx] is not None]
        e_vals = [data[i][e_idx] for i in range(start_idx, end_idx) if data[i][e_idx] is not None]
        d_med  = statistics.median(d_vals) if d_vals else None
        e_med  = statistics.median(e_vals) if e_vals else None
        odrzuc_t  = _rozne_procent(b, d_med, MAX_ROZNICA_PROCENT)
        odrzuc_rh = (not temp_only) and _rozne_procent(c, e_med, MAX_ROZNICA_PROCENT)
        if odrzuc_t or odrzuc_rh:
            rh_txt = f"{c}->{e_med}" if not temp_only else "—"
            print(f"    [ODRZUCONO >{MAX_ROZNICA_PROCENT:g}%] nastawa vs odczyt komory: "
                  f"T {b}->{d_med}, RH {rh_txt} — pomijam segment.")
            return None

    # Wiersze spełniające kryteria
    valid_indices = []
    for i in range(start_idx, end_idx):
        row = data[i]
        dt, temp, rh, k = row[0], row[1], row[2], row[3]
        ls = row[4:8] if is_cc04 else (row[4],)   # 4 lub 1 spread temperatury

        if None in (dt, temp, rh):
            continue
        if any(l is None for l in ls):
            continue
        if any(l > 0.1 for l in ls):
            continue

        if not temp_only:
            if k is None:
                continue
            threshold = get_k_threshold(temp, rh)
            if threshold is None or k > threshold:
                continue

        valid_indices.append(i)

    n_l = '4L' if is_cc04 else 'L'
    crit_label = n_l if temp_only else f"K+{n_l}"
    print(f"    Wiersze spelniajace kryteria {crit_label}: {len(valid_indices)}")

    if not valid_indices:
        print(f"    brak wierszy spelniajacych kryteria {crit_label} — pomijam.")
        return None

    # Score function dla optymalizacji 5 reprezentantow
    if temp_only:
        if is_cc04:
            def score_fn(i):
                vals = [v for v in data[i][4:8] if v is not None]
                return max(vals) if vals else None
        else:
            score_fn = lambda i: data[i][4]
    else:
        score_fn = lambda i: data[i][3]   # K (roztdp) dla CC i CC-04

    rep_indices = _find_best_minute_reps(data, valid_indices, start_idx, end_idx, score_fn)
    if rep_indices is None:
        print(f"    brak 5 kolejnych minut z waznymi kryteriami — pomijam.")
        return None

    rep_times = [data[i][0] for i in rep_indices]
    print(f"    Wiersze reprezentacyjne: {[str(t) for t in rep_times]}")
    if not temp_only:
        rep_k = [round(data[i][3], 4) for i in rep_indices]
        print(f"    K (roztdp)  : {rep_k}  srednia={round(sum(rep_k)/5, 4)}")
    if is_cc04:
        for off, lbl in enumerate(['R', 'S', 'T', 'U']):
            vals = [round(data[i][4 + off], 4) for i in rep_indices
                    if data[i][4 + off] is not None]
            if vals:
                print(f"    L{lbl} (roztemp): {vals}  srednia={round(sum(vals)/len(vals),4)}")
    else:
        rep_l = [round(data[i][4], 4) for i in rep_indices]
        print(f"    L (roztemp) : {rep_l}  srednia={round(sum(rep_l)/5, 4)}")

    # Zielony obszar: od pelnej minuty pierwszego rep do +5 minut
    green_start = rep_times[0].replace(second=0, microsecond=0)
    green_end   = green_start + datetime.timedelta(minutes=5)

    n_cols = 21 if is_cc04 else 12

    green_indices = [
        i for i in range(start_idx, end_idx)
        if data[i][0] is not None and green_start <= data[i][0] <= green_end
    ]
    for i in green_indices:
        _apply_fill(ws, 2 + i, FILL_LIGHT, n_cols)
    if green_indices:
        print(f"    Wiersze jasno-zielone: {len(green_indices)} "
              f"(Excel {2+green_indices[0]}–{2+green_indices[-1]})")

    for i in rep_indices:
        _apply_fill(ws, 2 + i, FILL_DARK, n_cols)
        _apply_bold(ws, 2 + i, n_cols)

    return rep_indices


def analyze_and_highlight(ws, rows, file_type='CC'):
    """
    Dla każdego segmentu stabilnego B+C (>= STABILIZACJA_MIN):
      - koloruje najlepsze 5-minutowe okno na #E2EFDA
      - oznacza 5 wierszy reprezentacyjnych (#A9D08E + bold)
    """
    data = _prepare_data_cc04(rows) if file_type == 'CC04' else _prepare_data(rows)

    windows = _find_all_analysis_windows(data, STABILIZACJA_MIN)
    if not windows:
        print(f"  ANALIZA: brak segmentow stabilnych >= {STABILIZACJA_MIN} — pomijam kolorowanie.")
        return

    # Wykrywanie gisterezy:
    # Jesli (T, RH) pojawia sie powторnie ORAZ miedzy pierwszym a obecnym wystаpieniem
    # byl segment z tym samym T ale innym RH → powrót z gisterezy.
    # Segment NASTEPUJACY po takim powrocie jest zawsze suszeniem lub punktem tylko-temp.
    # Ta sama temperatura moze sie roznic o maks. ~1 stopien (poprawka czujnika),
    # dlatego T porownujemy z tolerancja < 1.0 C przy wykrywaniu gisterezy.
    # Segment C=0 (tylko-temp) po gisterezie NIE jest pomijany — mierzymy temperature.
    def _t_key(t_val, t_map):
        """Zwraca klucz z t_map najblizszy t_val (tolerancja < 1.0 C), lub t_val jesli brak."""
        for k in t_map:
            if abs(k - t_val) < 1.0:
                return k
        return t_val

    last_rh_for_t: dict = {}   # t_key -> ostatnio widziana RH dla tego T
    seen_tr:       set  = set() # (t_key, RH) widziane co najmniej raz

    post_histereza: set = set()
    for i, (si, _) in enumerate(windows):
        b = data[si][1]
        c = data[si][2]
        t  = round(b, 1) if b is not None else None
        rh = round(c, 1) if c is not None else None
        if t is not None and rh is not None:
            # Segment tylko-temperatura (RH=0): po nim komora wraca do normy,
            # nastepny segment jest suszeniem/powrotem — tak samo jak po gisterezie.
            if rh == 0:
                if i + 1 < len(windows):
                    post_histereza.add(i + 1)
                print(f"  [TYLKO-TEMP] T~{t} RH=0 (seg {i+1}), "
                      f"nastepny seg {i+2} jest po-suszeniowy")

            tk = _t_key(t, last_rh_for_t)   # kanoniczny klucz T (tolerancja < 1 C)
            prev_rh = last_rh_for_t.get(tk)
            if (tk, rh) in seen_tr and prev_rh is not None and prev_rh != rh:
                # Powrot do (T_group, RH) po segmencie z innym RH → gistereza
                if i + 1 < len(windows):
                    post_histereza.add(i + 1)
                print(f"  [HISTEREZA] T~{tk}: ...-> {prev_rh} -> {rh} (seg {i+1}), "
                      f"nastepny seg {i+2} jest po-histerezowy")
            seen_tr.add((tk, rh))
            last_rh_for_t[tk] = rh

    # Filtrowanie suszenia:
    # Pomijamy jezeli: T w SUSZENIE_T_ZAKRES, RH w (0, SUSZENIE_RH_MAX]
    # ORAZ (ta sama para (T,RH) juz wczesniej wystapila LUB segment jest po-histerezowy).
    t_min, t_max = SUSZENIE_T_ZAKRES
    seen_bc: set = set()
    filtered = []

    for idx, (start_idx, end_idx) in enumerate(windows):
        b = data[start_idx][1]
        c = data[start_idx][2]
        bc = (round(b, 1) if b is not None else None,
              round(c, 1) if c is not None else None)

        is_suszenie = (
            b is not None and c is not None
            and c > 0                      # RH=0 = brak sterowania wilgotnoscia, nie suszenie
            and t_min <= b <= t_max
            and c <= SUSZENIE_RH_MAX
            and (bc in seen_bc             # powtorzony segment
                 or idx in post_histereza) # lub pierwszy po gisterezie
        )

        if is_suszenie:
            t0 = data[start_idx][0]
            reason = 'po-histereza' if idx in post_histereza and bc not in seen_bc else 'powtorzenie'
            print(f"  [SUSZENIE/{reason}] Pomijam segment T={b} RH={c}  (start: {t0})")
        else:
            filtered.append((start_idx, end_idx))

        seen_bc.add(bc)

    print(f"  Segmentow lacznie: {len(windows)},  po odfiltrowaniu suszenia: {len(filtered)}")

    segments = []
    found = 0
    for num, (start_idx, end_idx) in enumerate(filtered, 1):
        rep_idx = _process_segment(ws, data, start_idx, end_idx, num, file_type)
        if rep_idx is not None:
            found += 1
            segments.append(rep_idx)
    print(f"  Pokolorowano segmentow: {found}/{len(filtered)}")
    return segments


# =============================================================================
# DANE ŚRODOWISKOWE Z WYNIKÓW (wyniki/*.xlsx → Strona 3)
# =============================================================================

def _zaladuj_wyniki_xlsx(path):
    """
    Wczytuje zunifikowany plik wynikow. Kolumny wykrywane po NAGLOWKU (wiersz 1):
      - Czas         : kolumna 1,
      - Temperatura  : pierwsza kolumna z 'Temperatura' (dla xTHERM = wewnetrzna),
      - Temperatura2 : druga kolumna z 'Temperatura' (xTHERM = zewnetrzna; inaczej None),
      - Wilgotnosc   : kolumna z 'Wilgotn'/'%rh'.
    Gdy naglowkow brak — fallback pozycyjny (kol. 2 = temp, kol. 3 = RH),
    zgodny ze starym 3-kolumnowym formatem.
    Zwraca liste krotek (dt, temp, rh, row_1based, temp2).
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active

        header = [str(c.value).lower() if c.value is not None else '' for c in ws[1]]
        temp_cols = [i for i, h in enumerate(header) if 'temperatura' in h]
        temp_idx  = temp_cols[0] if temp_cols else 1
        temp2_idx = temp_cols[1] if len(temp_cols) > 1 else None
        rh_idx    = next((i for i, h in enumerate(header) if 'wilgot' in h or '%rh' in h), 2)

        result = []
        for r_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or row[0] is None:
                continue
            dt_val   = row[0]
            temp_val = row[temp_idx] if len(row) > temp_idx else None
            rh_val   = row[rh_idx]   if len(row) > rh_idx else None
            temp2_val = row[temp2_idx] if (temp2_idx is not None and len(row) > temp2_idx) else None

            if isinstance(dt_val, datetime.datetime):
                dt = dt_val
            else:
                dt = _s_to_dt(str(dt_val))
            if dt is None:
                continue

            def _to_f(v):
                try:
                    return float(str(v).replace(',', '.')) if v is not None else None
                except (ValueError, TypeError):
                    return None

            result.append((dt, _to_f(temp_val), _to_f(rh_val), r_idx, _to_f(temp2_val)))
        wb.close()
        return result
    except Exception as exc:
        print(f"  [WYNIKI] Blad odczytu '{os.path.basename(path)}': {exc}")
        return []


def _znajdz_5_wierszy(wyniki_rows, target_times):
    """
    Dla listy 5 docelowych dat/czasow zwraca liste 5 najblizszych wpisow
    z wyniki_rows [(dt, temp, rh, row_idx, temp2)].
    Zwraca liste 5 krotek lub None jesli wyniki_rows jest puste.
    """
    if not wyniki_rows:
        return None
    matched = []
    for t_target in target_times:
        best = min(wyniki_rows, key=lambda e: abs((e[0] - t_target).total_seconds()))
        matched.append(best)
    return matched


def _oznacz_wyniki_xlsx(path, row_indices):
    """Koloruje podane wiersze w pliku wynikow na ciemno-zielono + bold."""
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        n_cols = ws.max_column
        for r_idx in row_indices:
            for col in range(1, n_cols + 1):
                cell = ws.cell(row=r_idx, column=col)
                cell.fill = FILL_DARK
                f = cell.font
                cell.font = Font(
                    name=f.name, size=f.size, italic=f.italic,
                    underline=f.underline, strike=f.strike, color=f.color,
                    bold=True,
                )
        wb.save(path)
        print(f"    [WYNIKI] Oznaczono {len(row_indices)} wierszy: {os.path.basename(path)}")
    except Exception as exc:
        print(f"    [WYNIKI] Blad oznaczania '{os.path.basename(path)}': {exc}")


def _wypelnij_wyniki_srodowiskowe(proto_ws, rep_groups, rows_obs, obs_type):
    """
    Wpisuje dane srodowiskowe (Temperatura+Wilgotnosc) z plikow WYNIKI_FOLDER do
    Strona 3.

    UKLAD: kazdy przyrzad (logger) = osobna PARA kolumn idaca w PRAWO —
      CC   : Q/R, S/T, U/V, ...      (start_col = 17)
      CC04 : S/T, U/V, W/X, Y/Z, AA/AB (start_col = 19, max 5 par)
    Wiersze (bloki po 5) = kolejne punkty pomiarowe.

    Pliki przypisywane do przyrzadow po KOLEJNOSCI NAZW; brane sa tylko te, ktore
    pasuja czasowo do punktow (stare/niepasujace pliki pomijamy automatycznie).

    PRZYPADEK SPECJALNY (rzadki) — plik z 4 kolumnami (xTHERM: Czas | Temp wewn |
    Temp zewn | Wilgotnosc) liczy sie jako DWA przyrzady:
      - para N   : temperatura WEWNETRZNA + wilgotnosc,
      - para N+1 : temperatura ZEWNETRZNA, BEZ wilgotnosci (kolumna RH pusta).
    """
    if not os.path.exists(WYNIKI_FOLDER):
        print(f"  [WYNIKI] Brak folderu '{WYNIKI_FOLDER}' — pomijam dane srodowiskowe.")
        return

    pliki = sorted(
        f for f in os.listdir(WYNIKI_FOLDER)
        if f.lower().endswith('.xlsx')
    )
    if not pliki:
        print(f"  [WYNIKI] Brak plikow .xlsx w '{WYNIKI_FOLDER}' — pomijam.")
        return

    print(f"  [WYNIKI] Wczytuje {len(pliki)} plik(ow) wynikow...")
    baza = {}
    for fname in pliki:
        dane = _zaladuj_wyniki_xlsx(os.path.join(WYNIKI_FOLDER, fname))
        if dane:
            baza[fname] = dane
    if not baza:
        print(f"  [WYNIKI] Zadne pliki wynikow nie zawieraja danych — pomijam.")
        return

    start_col = WYNIKI_START_COL_CC04 if obs_type == 'CC04' else WYNIKI_START_COL_CC
    BLOCK_START_ROW = 20
    BLOCK_SIZE      = 5
    tol_s = datetime.timedelta(minutes=WYNIKI_TOLERANCJA_MIN).total_seconds()

    # Liczba par kolumn przyrzadow — wykrywana z naglowka szablonu (wiersz 10 ma
    # numery 1,2,3,... co 2 kolumny: S=1, U=2, W=3, ...). Dzieki temu nie zależymy
    # od stalej — szablon CC04 ma ich 10 (S/T..AK/AL), dalej sa kolumny 'rozrzut'.
    MAX_PRZYRZADY = 0
    _c = start_col
    while _c <= proto_ws.max_column:
        _v = proto_ws.cell(row=10, column=_c).value
        if isinstance(_v, (int, float)) and int(_v) == MAX_PRZYRZADY + 1:
            MAX_PRZYRZADY += 1
            _c += 2
        else:
            break
    if MAX_PRZYRZADY == 0:
        MAX_PRZYRZADY = 5   # fallback, gdy nie udalo sie odczytac naglowka
    print(f"  [WYNIKI] Szablon ma {MAX_PRZYRZADY} par kolumn przyrzadow (od "
          f"{get_column_letter(start_col)}).")

    # Docelowe czasy dla kazdego punktu (None gdy brak timestampow)
    punkty = []
    for rep_indices in rep_groups:
        tts = [_s_to_dt(rows_obs[i][0]) for i in rep_indices]
        punkty.append(None if any(t is None for t in tts) else tts)

    oznaczenia_per_plik = {}   # fname -> set(row_idx)
    dev = 0                    # numer przyrzadu (0-based) -> para kolumn w prawo

    def _wpisz(col, matched_per_punkt, val_idx, z_wilgotnoscia):
        """Wpisuje jedna kolumne (temp z val_idx) + opc. wilgotnosc do bloków punktow."""
        for punkt_idx, matched in matched_per_punkt.items():
            r0 = BLOCK_START_ROW + punkt_idx * BLOCK_SIZE
            for off, krotka in enumerate(matched):
                er = r0 + off
                temp_v = krotka[val_idx]
                rh_v   = krotka[2]
                row_idx = krotka[3]
                if temp_v is not None:
                    proto_ws.cell(row=er, column=col).value = round(temp_v, 4)
                if z_wilgotnoscia and rh_v is not None:
                    proto_ws.cell(row=er, column=col + 1).value = round(rh_v, 4)
                oznaczenia_per_plik.setdefault(fname, set()).add(row_idx)

    for fname in pliki:
        if fname not in baza:
            continue
        dane = baza[fname]

        # Dopasuj plik do punktow: 5 najblizszych wierszy w tolerancji czasowej
        dopasowania = {}   # punkt_idx -> lista 5 krotek
        for punkt_idx, tts in enumerate(punkty):
            if tts is None:
                continue
            matched = _znajdz_5_wierszy(dane, tts)
            if matched is None:
                continue
            avg = sum(abs((tts[i] - matched[i][0]).total_seconds()) for i in range(5)) / 5
            if avg <= tol_s:
                dopasowania[punkt_idx] = matched

        if not dopasowania:
            continue   # plik nie pasuje czasowo (np. inne/stare pomiary) — pomijamy

        ma_zewn = any(row[4] is not None for row in dane)   # temp2 -> xTHERM = 2 przyrzady

        if dev >= MAX_PRZYRZADY:
            print(f"  [WYNIKI] '{fname}' — brak wolnej kolumny przyrzadu "
                  f"(max {MAX_PRZYRZADY}) — pomijam.")
            continue

        # Przyrzad: temperatura (wewnetrzna) + wilgotnosc
        col = start_col + dev * 2
        punkty_txt = ', '.join(str(p + 1) for p in sorted(dopasowania))
        print(f"  [WYNIKI] Przyrzad {dev+1} ({get_column_letter(col)}/{get_column_letter(col+1)})"
              f" <- '{fname}'  (punkty: {punkty_txt})")
        _wpisz(col, dopasowania, val_idx=1, z_wilgotnoscia=True)
        dev += 1

        # Przyrzad 2 (tylko xTHERM): temperatura zewnetrzna, BEZ wilgotnosci
        if ma_zewn:
            if dev >= MAX_PRZYRZADY:
                print(f"  [WYNIKI] '{fname}' (zewn.) — brak wolnej kolumny — pomijam czesc zewn.")
            else:
                col = start_col + dev * 2
                print(f"  [WYNIKI] Przyrzad {dev+1} ({get_column_letter(col)}, temp zewn., bez RH)"
                      f" <- '{fname}'")
                _wpisz(col, dopasowania, val_idx=4, z_wilgotnoscia=False)
                dev += 1

    if dev == 0:
        print(f"  [WYNIKI] Zaden plik nie pasowal czasowo do punktow — brak danych srodowiskowych.")

    # Oznacz wiersze w plikach wynikow
    for fname, row_set in oznaczenia_per_plik.items():
        _oznacz_wyniki_xlsx(os.path.join(WYNIKI_FOLDER, fname), sorted(row_set))


# =============================================================================
# GENEROWANIE PROTOKOŁU
# =============================================================================

def _delete_rows_via_excel(filepath, sheet_name, row_from, row_count):
    """
    Usuwa row_count wierszy od row_from przez Excel COM (przesuwa reszte w gore —
    odpowiednik zaznaczenia wierszy + Delete / Przesuniecie w gore).
    Zwraca True przy sukcesie, False jesli COM niedostepny.
    """
    try:
        import win32com.client
    except ImportError:
        return False

    xl = wb = None
    try:
        xl = win32com.client.Dispatch('Excel.Application')
        xl.Visible        = False
        xl.DisplayAlerts  = False
        xl.ScreenUpdating = False

        wb = xl.Workbooks.Open(os.path.abspath(filepath))
        ws = None
        for idx in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(idx).Name == sheet_name:
                ws = wb.Sheets(idx)
                break
        if ws is None:
            return False

        addr = f'{row_from}:{row_from + row_count - 1}'
        ws.Rows(addr).Delete(Shift=-4162)   # xlShiftUp

        wb.Save()
        return True

    except Exception as exc:
        print(f'  [Excel COM] Blad usuwania wierszy: {exc}')
        return False
    finally:
        if wb is not None:
            try: wb.Close(False)
            except: pass
        if xl is not None:
            try: xl.Quit()
            except: pass


def _insert_blocks_via_excel(filepath, sheet_name, block_start, block_size, n_extra):
    """
    Otwiera plik w Excelu przez COM i wykonuje n_extra razy operacje
    'Kopiuj wiersze block_start..block_start+block_size-1 → Wstaw skopiowane komorki'
    (odpowiednik Ctrl+C / Wstaw z przesuniciem w dol).
    Zapisuje i zamyka plik.  Zwraca True przy sukcesie, False jesli COM niedostepny.
    """
    try:
        import win32com.client
    except ImportError:
        return False

    xl = wb = None
    try:
        xl = win32com.client.Dispatch('Excel.Application')
        xl.Visible         = False
        xl.DisplayAlerts   = False
        xl.ScreenUpdating  = False

        wb  = xl.Workbooks.Open(os.path.abspath(filepath))
        ws  = None
        for idx in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(idx).Name == sheet_name:
                ws = wb.Sheets(idx)
                break
        if ws is None:
            return False

        addr = f'{block_start}:{block_start + block_size - 1}'
        for _ in range(n_extra):
            ws.Rows(addr).Copy()
            ws.Rows(addr).Insert(Shift=-4121)   # xlShiftDown
            xl.CutCopyMode = False

        wb.Save()
        return True

    except Exception as exc:
        print(f'  [Excel COM] Blad wstawiania blokow: {exc}')
        return False
    finally:
        if wb is not None:
            try: wb.Close(False)
            except: pass
        if xl is not None:
            try: xl.Quit()
            except: pass


def _round_hm(dt):
    """Zaokrągla datetime do pelnej minuty i zwraca string HH:MM."""
    if dt is None:
        return None
    if dt.second >= 30:
        dt = dt + datetime.timedelta(minutes=1)
    return dt.strftime('%H:%M')


def generuj_protokol(rep_groups, rows, measurement_id, obs_type, sensor_names=None):
    """
    Tworzy plik protokolu na podstawie reprezentacyjnych wierszy z obserwacji.

    rep_groups : lista list[5 int] – indeksy (0-based do `rows`) dla kazdego punktu
    rows       : surowe wiersze danych z pliku TXT (list of list[str])
    obs_type   : 'CC' lub 'CC04'
    """
    TEMPLATE_BLOCKS = 6     # liczba blokow (punktow) w szablonie
    BLOCK_START_ROW = 20    # pierwszy wiersz Excel pierwszego bloku
    BLOCK_SIZE      = 5     # wierszy na jeden punkt

    if obs_type == 'CC04':
        tmpl_name   = PROTOKOL_CC04_TEMPLATE
        src_indices = [5, 6, 7, 8, 9]       # F,G,H,I,J (0-based w rows[])
        dst_cols    = [11, 12, 13, 14, 15]  # K,L,M,N,O (1-based w Excel)
    else:
        tmpl_name   = PROTOKOL_CC_TEMPLATE
        src_indices = [5, 6]                # F,G
        dst_cols    = [12, 13]              # L,M

    tmpl_path = os.path.join(FOLDER, tmpl_name)
    out_name  = tmpl_name.replace('xxx', measurement_id, 1)
    out_path  = os.path.join(FOLDER, out_name)

    if not os.path.exists(tmpl_path):
        print(f"\n  PROTOKOL: Brak szablonu '{tmpl_name}' — pomijam generowanie protokolu.")
        return

    N = len(rep_groups)
    if N == 0:
        print("  PROTOKOL: Brak punktow pomiarowych — pomijam.")
        return

    shutil.copy2(tmpl_path, out_path)
    print(f"\nSkopiowano szablon protokolu: {tmpl_name}")

    # ── Dopasowanie liczby bloków przez Excel COM (przed zaladowaniem openpyxl) ─
    if N > TEMPLATE_BLOCKS:
        extra = N - TEMPLATE_BLOCKS
        print(f"  Wstawiam {extra} dodatkowych blokow przez Excel COM...")
        ok = _insert_blocks_via_excel(out_path, 'Strona 3', BLOCK_START_ROW, BLOCK_SIZE, extra)
        if ok:
            print(f"  Wstawiono {extra} blokow.")
        else:
            print(f"  UWAGA: Excel COM niedostepny (pip install pywin32). "
                  f"Ograniezam do {TEMPLATE_BLOCKS} punktow.")
            N          = TEMPLATE_BLOCKS
            rep_groups = rep_groups[:N]

    elif N < TEMPLATE_BLOCKS:
        extra     = TEMPLATE_BLOCKS - N
        del_from  = BLOCK_START_ROW + N * BLOCK_SIZE
        del_count = extra * BLOCK_SIZE
        print(f"  Usuwam {extra} nadmiarowych blokow przez Excel COM...")
        ok = _delete_rows_via_excel(out_path, 'Strona 3', del_from, del_count)
        if ok:
            print(f"  Usunieto {extra} blokow.")
        else:
            print(f"  UWAGA: Excel COM niedostepny — nadmiarowe bloki pozostana w pliku.")

    # ── Ładujemy plik (już z prawidłową liczbą bloków) z openpyxl ────────────
    proto_wb = openpyxl.load_workbook(out_path)

    if 'Strona 3' not in proto_wb.sheetnames:
        print(f"  PROTOKOL: Brak arkusza 'Strona 3' w '{tmpl_name}' — pomijam.")
        proto_wb.close()
        return
    ws3 = proto_wb['Strona 3']

    # ── Zapis danych dla każdego punktu ───────────────────────────────────────
    FILL_GREY = PatternFill(fill_type='solid', fgColor='BFBFBF')

    for punkt_idx, rep_indices in enumerate(rep_groups):
        r0 = BLOCK_START_ROW + punkt_idx * BLOCK_SIZE

        first_i   = rep_indices[0]
        last_i    = rep_indices[-1]
        row_first = rows[first_i]

        # Dane pomiarowe (F,G[,H,I,J] z obserwacji) → kolumny docelowe protokolu
        # Jesli wartosc to "brak" — komorke zostawiamy pusta i zaznaczamy szarym
        for row_off, data_i in enumerate(rep_indices):
            r = rows[data_i]
            for src_ci, dst_col in zip(src_indices, dst_cols):
                raw  = r[src_ci] if src_ci < len(r) else ''
                cell = ws3.cell(row=r0 + row_off, column=dst_col)
                if str(raw).strip().lower() == 'brak':
                    cell.value = None
                    cell.fill  = FILL_GREY
                else:
                    cell.value = to_value(raw) if raw else None

        # D i E z pierwszego wiersza punktu → I(r0) i J(r0)
        ws3.cell(row=r0, column=9).value  = (to_value(row_first[3])
                                              if len(row_first) > 3 else None)
        ws3.cell(row=r0, column=10).value = (to_value(row_first[4])
                                              if len(row_first) > 4 else None)

        # B (Tzadana) i C (RHzadana) z pierwszego wiersza → B(r0) i C(r0)
        rh_val  = _s_to_float(row_first[2]) if len(row_first) > 2 else None
        rh_zero = (rh_val is not None and rh_val == 0.0)

        ws3.cell(row=r0, column=2).value = (to_value(row_first[1])
                                             if len(row_first) > 1 else None)
        c_cell = ws3.cell(row=r0, column=3)
        if rh_zero:
            c_cell.value = '-'
            c_cell.fill  = FILL_GREY
        else:
            c_cell.value = rh_val

        # Jesli RH=0 (tryb tylko temperatura): kolumna O na szaro + "-" w kolumnach bocznych
        if rh_zero:
            for row_off in range(BLOCK_SIZE):
                ws3.cell(row=r0 + row_off, column=15).fill = FILL_GREY  # O
            if obs_type == 'CC04':
                ws3.cell(row=r0, column=17).value = '-'   # Q
                ws3.cell(row=r0, column=18).value = '-'   # R
            else:
                ws3.cell(row=r0, column=15).value = '-'   # O  (CC)
                ws3.cell(row=r0, column=16).value = '-'   # P  (CC)

        # Data i czasy z kolumny A obserwacji
        start_dt = _s_to_dt(row_first[0])
        end_dt   = _s_to_dt(rows[last_i][0])

        if start_dt:
            ws3.cell(row=r0 + 1, column=5).value = start_dt.strftime('%d.%m.%Y')
            ws3.cell(row=r0 + 2, column=5).value = _round_hm(start_dt)
        if end_dt:
            ws3.cell(row=r0 + 3, column=5).value = _round_hm(end_dt)

    # ── Podmiana Pt100-XX w naglowkach na Stronie 3 ──────────────────────────
    def _subst(cell, sname):
        """Podmienia Pt100-XX w wartosci komorki lub wpisuje nazwe jesli brak wzorca."""
        if not sname:
            return
        if cell.value and re.search(r'Pt100-\d+', str(cell.value)):
            cell.value = re.sub(r'Pt100-\d+', sname, str(cell.value))
        else:
            cell.value = sname

    if sensor_names:
        if obs_type == 'CC04':
            # K9, L9, M9, N9 → czujniki 1-4
            for col_off, sname in enumerate(sensor_names[:4]):
                _subst(ws3.cell(row=9, column=11 + col_off), sname)
        else:
            # CC: L9 → jeden czujnik
            _subst(ws3.cell(row=9, column=12), sensor_names[0])

    # ── Strona 1: F10 = "1-N" oraz czujniki wzorcowe ─────────────────────────
    if 'Strona 1' in proto_wb.sheetnames:
        ws1 = proto_wb['Strona 1']
        ws1.cell(row=10, column=6).value = f"1-{N}"

        if sensor_names:
            if obs_type == 'CC04':
                # G:H36-39 → czujniki 1-4  (komorki scalone G:H, piszemy do G=7)
                for row_off, sname in enumerate(sensor_names[:4]):
                    _subst(ws1.cell(row=36 + row_off, column=7), sname)
            else:
                # CC: H:I43 → jeden czujnik  (piszemy do H=8)
                _subst(ws1.cell(row=43, column=8), sensor_names[0])

    # ── Dane srodowiskowe z wynikow (wyniki/*.xlsx → Q/S kolumny Strona 3) ──────
    print("\n  Szukam danych srodowiskowych w wynikach...")
    _wypelnij_wyniki_srodowiskowe(ws3, rep_groups, rows, obs_type)

    proto_wb.save(out_path)
    print(f"Zapisano protokol: {out_name}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    measurement_id = parse_measurement_id(TXT_FILENAME)
    txt_path = os.path.join(FOLDER, TXT_FILENAME)

    if not os.path.exists(txt_path):
        raise FileNotFoundError(f"Nie znaleziono pliku TXT: {txt_path}")

    raw_lines = open_txt(txt_path)
    file_type = detect_file_type(raw_lines)

    print(f"Numer pomiaru : {measurement_id}")
    print(f"Plik wejsciowy: {TXT_FILENAME}")
    print(f"Typ pliku     : {file_type}")

    today        = datetime.date.today()
    sensor_names = None   # wypelniane w galezi CC04

    if file_type == 'CC04':
        # ── CC-04 ──────────────────────────────────────────────────────────────
        template_path = os.path.join(FOLDER, CC04_TEMPLATE)
        output_name   = CC04_TEMPLATE.replace('szablon', measurement_id, 1)
        output_path   = os.path.join(FOLDER, output_name)

        print(f"Szablon       : {CC04_TEMPLATE}")
        print(f"Plik wyjsciowy: {output_name}")

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Nie znaleziono szablonu: {template_path}")
        shutil.copy2(template_path, output_path)
        print("Skopiowano szablon.")

        sensor_names, rows = parse_txt_cc04(txt_path)
        print(f"Czujniki wzorcowe: {sensor_names}")
        print(f"Wierszy danych   : {len(rows)}")

        wb = openpyxl.load_workbook(output_path)
        ws = wb.active

        # A2:U(ostatni) – dane pomiarowe (21 kolumn)
        for r_i, row in enumerate(rows):
            for c_i, val in enumerate(row):
                ws.cell(row=2 + r_i, column=1 + c_i).value = to_value(val)

        # M1, N1, O1, P1 (kol. 13-16) – 4 czujniki wzorcowe
        for col, sname in zip(range(13, 17), sensor_names):
            cell = ws.cell(row=1, column=col)
            if cell.value and isinstance(cell.value, str):
                cell.value = re.sub(r'Pt100-\d+', sname, cell.value)
            elif sname:
                cell.value = 'Temperatura ' + sname
            print(f"  kol {col} (rząd 1) -> {cell.value!r}")

        # Y92, Y93 (kol. 25) – aktualna data
        ws.cell(row=92, column=25).value = today
        ws.cell(row=93, column=25).value = today

        print("\nAnaliza stabilnosci...")
        rep_groups = analyze_and_highlight(ws, rows, file_type='CC04')

    else:
        # ── CC (oryginalny) ────────────────────────────────────────────────────
        template_path = os.path.join(FOLDER, TEMPLATE)
        output_name   = TEMPLATE.replace('xxx', measurement_id, 1)
        output_path   = os.path.join(FOLDER, output_name)

        print(f"Szablon       : {TEMPLATE}")
        print(f"Plik wyjsciowy: {output_name}")

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Nie znaleziono szablonu: {template_path}")
        shutil.copy2(template_path, output_path)
        print("Skopiowano szablon.")

        sensor_name, rows = parse_txt(txt_path)
        sensor_names = [sensor_name]   # ujednolicamy z formatem CC04
        print(f"Czujnik wzorcowy: {sensor_name}")
        print(f"Wierszy danych  : {len(rows)}")

        wb = openpyxl.load_workbook(output_path)
        ws = wb.active

        # A2:L(ostatni) – dane pomiarowe
        for r_i, row in enumerate(rows):
            for c_i, val in enumerate(row):
                ws.cell(row=2 + r_i, column=1 + c_i).value = to_value(val)

        # J1 – aktualizacja czujnika wzorcowego
        j1 = ws.cell(row=1, column=10)
        if j1.value and isinstance(j1.value, str):
            new_val = re.sub(r'Pt100-\d+', sensor_name, j1.value)
            if new_val == j1.value and sensor_name:
                print(f"  UWAGA: Nie znaleziono wzorca 'Pt100-XX' w J1: {j1.value!r}")
            j1.value = new_val
        else:
            j1.value = sensor_name
        print(f"J1 -> {j1.value!r}")

        # N92 – data,  O92 – podpis
        ws.cell(row=92, column=14).value = today
        ws.cell(row=92, column=15).value = PODPIS

        print("\nAnaliza stabilnosci...")
        rep_groups = analyze_and_highlight(ws, rows, file_type='CC')

    wb.save(output_path)
    print(f"\nZapisano: {output_name}")

    if rep_groups:
        generuj_protokol(rep_groups, rows, measurement_id, file_type,
                         sensor_names=sensor_names)

    print("Gotowe!")


if __name__ == '__main__':
    main()
