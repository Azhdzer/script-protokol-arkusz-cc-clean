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
from openpyxl.styles import PatternFill, Font, Border
from openpyxl.utils import get_column_letter

import pz_dane   # wspolny modul: dane przyrzadow z PZ + Zestawienie

# =============================================================================
# KONFIGURACJA  ← edytuj tutaj przed uruchomieniem
# =============================================================================

FOLDER           = os.environ.get("CC_FOLDER") or \
                   r"C:\Users\artisom.azhdzer\Desktop\Script protokoł - arkusz CC"
TXT_FILENAME     = "2026-07-23 12.23_001"
# Wiele plikow TXT jednego (przerwanego) pomiaru — zostana rozparsowane i
# sklejone w jeden ciag. Pusta lista => uzywany jest pojedynczy TXT_FILENAME.
# Panel GUI moze podac liste przez zmienna OBS_TXT_FILES (rozdzielona ';').
TXT_FILENAMES    = ["",""]
TEMPLATE         = "xxx_LA_TH_2026 - obserwacje CC.xlsx"
CC04_TEMPLATE    = "szablon_LA_TH_2026 - obserwacje.xlsx"
PROTOKOL_CC_TEMPLATE   = "xxx_LA_TH_2026 - protokół CC.xlsx"
PROTOKOL_CC04_TEMPLATE = "xxx_LA_TH_2026 - protokół CC-04.xlsx"
PODPIS           = "Artsiom Azhdzer"
STABILIZACJA_MIN  = datetime.timedelta(hours=1, minutes=45)   # rozgrzewka: gdy odczyty nie wejda w widelki, okno od tego czasu od poczatku punktu

# Dobor okna analizy: liczymy od momentu, gdy ODCZYTY komory wejda w widelki wokol nastaw:
#   • temperatura: |Todczytana - Tzadana| <= PROG_WEJSCIA_TEMP  (°C, bezwzglednie),
#   • wilgotnosc  : |RHodczytana - RHzadana| w granicach PROG_WEJSCIA_RH_PROC %  (wzglednie),
# i odliczamy STABILIZACJA_PO_RH (2h). Okno = od tego czasu do konca punktu (tam 5 reprezentantow).
# WAZNE: gdy punkt trzymany jest ~2h, samo 2h "zjadloby" caly punkt — dlatego start
# jest cofany tak, by na koncu punktu zostal zawsze ogon pomiarowy MIN_OKNO_ANALIZY.
# Punkt tylko-temperatura: liczy sie samo wejscie temperatury w widelki.
STABILIZACJA_PO_RH    = datetime.timedelta(hours=2)
PROG_WEJSCIA_TEMP     = 0.4    # +-0.4 °C (bezwzglednie) od Tzadana — temperatura nie moze "uciekac"
PROG_WEJSCIA_RH_PROC  = 3.0    # +-3% (wzglednie, jak MAX_ROZNICA_PROCENT) od RHzadana
MIN_OKNO_ANALIZY      = datetime.timedelta(minutes=15)   # gwarantowany ogon pomiarowy na koncu punktu (nigdy 0)
SUSZENIE_T_ZAKRES = (21.0, 27.0)   # zakres T [°C] charakterystyczny dla suszenia
SUSZENIE_RH_MAX   = 50.0           # max RH [%] — ponizej tej wartosci punkt moze byc suszeniem

# Folder z zunifikowanymi plikami wynikow (z analizuj_excele.py)
WYNIKI_FOLDER = os.path.join(FOLDER, "wyniki")

# Dane przyrzadow: PDFy "Potwierdzenie zamowienia" (PZ/) + Zestawienie rozdzielczosci.
PZ_FOLDER        = os.environ.get("CC_PZ_FOLDER") or os.path.join(FOLDER, "PZ")
ZESTAWIENIE_PLIK = os.path.join(FOLDER, "Zestawienie wzorcowanych przyrządów.xlsx")

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

# --- Nadpisania z panelu GUI (app_gui.py) przez zmienne srodowiskowe ---
if os.environ.get("OBS_FILTR") is not None:
    FILTR_NASTAWA_ODCZYT = os.environ["OBS_FILTR"].strip().lower() in ("1", "true", "tak", "yes", "on")
try:
    if os.environ.get("OBS_PROG"):
        MAX_ROZNICA_PROCENT = float(os.environ["OBS_PROG"])
except ValueError:
    pass
try:
    if os.environ.get("OBS_TOL"):
        WYNIKI_TOLERANCJA_MIN = int(float(os.environ["OBS_TOL"]))
except ValueError:
    pass

# =============================================================================

FILL_LIGHT = PatternFill(fill_type='solid', fgColor='E2EFDA')  # jasna zieleń
FILL_DARK  = PatternFill(fill_type='solid', fgColor='A9D08E')  # ciemna zieleń
# Punkt, ktory NIE przeszedl kryterium (wybrany awaryjnie „najblizej po czasie"):
FILL_WARN_LIGHT = PatternFill(fill_type='solid', fgColor='FCE4D6')  # jasny pomaranczowy (blok)
FILL_WARN_DARK  = PatternFill(fill_type='solid', fgColor='F4B183')  # pomaranczowy (reprezentanci)


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


# --- CC-04: kanaly czujnikow wzorcowych ------------------------------------
# Multimetr rejestruje np. 8 kanalow (Ch101..Ch108). W OBSERWACJI pokazujemy WSZYSTKIE
# (pelna struktura pliku, kolejnosc jak w pliku), a do ANALIZY i PROTOKOLU bierzemy tylko
# GLOWNE (typowo „co drugi": 101,103,105,107). Numery Pt100 (09/13/01/18) czytane sa
# z pliku PO NUMERZE KANALU — nie wpisujemy ich recznie.
CC04_KANALY_WSZYSTKIE = [101, 102, 103, 104, 105, 106, 107, 108]   # kolejnosc jak w pliku
CC04_KANALY_GLOWNE    = [101, 103, 105, 107]                       # -> analiza + protokol
CC04_KANALY_ZAPASOWE  = [c for c in CC04_KANALY_WSZYSTKIE if c not in CC04_KANALY_GLOWNE]

# Pelny uklad kolumn CC-04 — DOKLADNIE jak w pliku multimetru (wszystkie kanaly po kolei:
# najpierw odczyty ChNNN, potem tempChNNN, na koncu roztempChNNN). Kolumny dobierane
# PO NAZWIE z naglowka pliku, wiec odporne na dodatkowe/inne kanaly.
CC04_KOLUMNY = (
    ['Data Czas', 'Tzadana', 'RHzadana', 'Todczytana', 'RHodczytana']
    + [f'Ch{ch}' for ch in CC04_KANALY_WSZYSTKIE]
    + ['tdp', 'thigro', '%rh']
    + [f'tempCh{ch}' for ch in CC04_KANALY_WSZYSTKIE]
    + ['roztdp(15min)']
    + [f'roztempCh{ch}' for ch in CC04_KANALY_WSZYSTKIE]
)


def _znajdz_naglowek(lines):
    """
    Znajduje wiersz naglowka kolumn (zaczyna sie od 'Data Czas') niezaleznie od
    dlugosci bloku metadanych. Zwraca (index_naglowka, mapa nazwa->pozycja) albo
    (None, None) gdy nie znaleziono.
    """
    for idx, line in enumerate(lines):
        if line.strip().lower().startswith('data czas'):
            names = [c.strip() for c in re.split(r'[;\t,]', line.strip())]
            return idx, {n: i for i, n in enumerate(names) if n}
    return None, None


def _mapa_kanal_pt(lines):
    """
    Mapa numer_kanalu -> 'Pt100-XX' z linii naglowka pliku CC-04, np.:
      "Czujnik wzorcowy: Pt100-09; Wejscie pomiarowe Ch: 101"  ->  {101: 'Pt100-09'}
    Uzywana do przypisania POPRAWNYCH numerow czujnikow do kanalow (glownych i zapasowych).
    """
    mapa = {}
    for line in lines:
        m = re.search(r'(Pt100-\d+).*?Ch:\s*0*(\d+)', line)
        if m:
            mapa[int(m.group(2))] = m.group(1)
    return mapa


def _naglowki_cc04(kanal_pt):
    """
    Naglowki (rzad 1) obserwacji CC-04 dla PELNEGO ukladu CC04_KOLUMNY — z nazwami
    czujnikow (Pt100-XX) zamiast surowych 'ChNNN'. Kolejnosc dokladnie jak CC04_KOLUMNY.
    """
    def et(ch):
        return kanal_pt.get(ch) or f'Ch{ch}'
    return (
        ['Data Czas', 'Tzadana', 'RHzadana', 'Todczytana', 'RHodczytana']
        + [f'Wskazania multimetru {et(ch)}' for ch in CC04_KANALY_WSZYSTKIE]
        + ['TPunktuRosy', 'Temperatura', '%RH']
        + [f'Temperatura {et(ch)}' for ch in CC04_KANALY_WSZYSTKIE]
        + ['RozrzutTPunktuRosy(15min)']
        + [f'RozrzutTemperatura_{et(ch)}(15min)' for ch in CC04_KANALY_WSZYSTKIE]
    )


def parse_txt_cc04(path: str):
    """
    Format CC-04:
      Linie 8-11 (ind. 7-10) : 4 czujniki, np. 'Czujnik wzorcowy: Pt100-09; Wejscie...'
      Naglowek kolumn        : wiersz zaczynajacy sie od 'Data Czas' (pozycja zmienna!)
      Dane pomiarowe         : wiersze ponizej naglowka

    Kolumny dobierane sa PO NAZWIE do ukladu kanonicznego CC04_KOLUMNY (21 kol.),
    co czyni parser odpornym na pliki z dodatkowymi kanalami (np. Przetwornik U) —
    kluczowe przy sklejaniu wielu plikow o roznej liczbie kolumn.
    Zwraca (sensor_names: list[str], rows: list[list[str]]).
    """
    lines = open_txt(path)

    # Numer Pt100 kazdego czujnika bierzemy PO NUMERZE KANALU (nie po kolejnosci linii,
    # bo dane sa z „co drugiego" kanalu). Nazwy do protokolu/obserwacji = kanaly GLOWNE
    # (101,103,105,107 -> Pt100-09,-13,-01,-18).
    kanal_do_pt = _mapa_kanal_pt(lines)
    sensor_names = [kanal_do_pt.get(ch, '') for ch in CC04_KANALY_GLOWNE]
    while len(sensor_names) < len(CC04_KANALY_GLOWNE):
        sensor_names.append('')

    n_kol = len(CC04_KOLUMNY)
    hdr_idx, colmap = _znajdz_naglowek(lines)
    rows = []

    if colmap is not None:
        # Mapowanie po nazwach kolumn — pelny uklad pliku (wszystkie kanaly po kolei).
        for line in lines[hdr_idx + 1:]:
            line = line.strip()
            if not line:
                continue
            parts = re.split(r'[;\t,]', line)
            row = []
            for name in CC04_KOLUMNY:
                ci = colmap.get(name)
                row.append(parts[ci] if (ci is not None and ci < len(parts)) else '')
            rows.append(row)
    else:
        # Fallback (stary tryb staloindeksowy) — gdy nie rozpoznano naglowka.
        for line in lines[17:]:
            line = line.strip()
            if not line:
                continue
            parts = re.split(r'[;\t,]', line)
            row = parts[:n_kol]
            while len(row) < n_kol:
                row.append('')
            rows.append(row)

    return sensor_names, rows


def _dostepne_txt():
    """Lista plikow .txt w FOLDER (bez plikow tymczasowych)."""
    try:
        return sorted(f for f in os.listdir(FOLDER)
                      if f.lower().endswith('.txt') and not f.startswith('~$'))
    except OSError:
        return []


def _znajdz_txt(nazwa):
    """
    Znajduje plik TXT tolerancyjnie (typowe pomylki w konfiguracji):
      - dokladna nazwa, z dodanym '.txt', ze spacja przed '.txt',
      - dopasowanie po nazwie bez rozszerzenia i nadmiaru spacji (rowne lub 'zaczyna sie od').
    Zwraca sciezke albo None.
    """
    if os.path.isabs(nazwa):
        return nazwa if os.path.exists(nazwa) else None
    baza = nazwa.strip()
    for k in (baza, baza + '.txt', baza.rstrip() + ' .txt'):
        p = os.path.join(FOLDER, k)
        if os.path.exists(p):
            return p

    def _norm(s):
        return re.sub(r'\s+', ' ', os.path.splitext(s)[0]).strip().lower()

    cel = _norm(baza)
    if cel:
        for f in _dostepne_txt():
            if _norm(f) == cel or _norm(f).startswith(cel):
                return os.path.join(FOLDER, f)
    return None


def resolve_txt_files():
    """
    Zwraca liste sciezek plikow TXT wejsciowych (posortowana chronologicznie).

    Zrodlo (wg priorytetu):
      1) zmienna OBS_TXT_FILES  — nazwy rozdzielone ';' (z panelu GUI),
      2) lista TXT_FILENAMES    — jesli niepusta,
      3) pojedynczy TXT_FILENAME.

    Nazwy wzgledne sa rozwiazywane wzgledem FOLDER. Sortowanie po nazwie pliku
    daje kolejnosc chronologiczna (nazwa zaczyna sie od 'YYYY-MM-DD HH.MM').
    """
    env = os.environ.get("OBS_TXT_FILES")
    if env:
        names = [n.strip() for n in env.split(';') if n.strip()]
    else:
        # TXT_FILENAMES (jesli zdefiniowane i niepuste) ma pierwszenstwo, w
        # przeciwnym razie TXT_FILENAME. globals().get() sprawia, ze zakomentowanie
        # TXT_FILENAMES nie wywala skryptu. Oba moga byc stringiem (jeden plik)
        # albo lista (wiele plikow) — normalizujemy do listy.
        src = globals().get("TXT_FILENAMES") or TXT_FILENAME
        names = [src] if isinstance(src, str) else list(src)
    names = [str(n).strip() for n in names if str(n).strip()]
    # Gdy lista pusta (np. TXT_FILENAMES = ["", ""]) — wroc do pojedynczego TXT_FILENAME.
    if not names and isinstance(TXT_FILENAME, str) and TXT_FILENAME.strip():
        names = [TXT_FILENAME.strip()]

    dostepne = _dostepne_txt()
    lista_txt = "\n".join(f"    • {f}" for f in dostepne) or "    (brak plikow .txt w folderze)"

    if not names:
        raise FileNotFoundError(
            "Nie podano pliku TXT — ustaw TXT_FILENAME = \"nazwa.txt\" (albo TXT_FILENAMES / "
            "OBS_TXT_FILES).\n  Dostepne pliki .txt:\n" + lista_txt)

    paths = []
    for n in names:
        p = _znajdz_txt(n)
        if p is None:
            raise FileNotFoundError(
                f"Nie znaleziono pliku TXT: '{n}'.\n  Dostepne pliki .txt:\n" + lista_txt)
        paths.append(p)

    paths.sort(key=lambda p: os.path.basename(p).lower())
    return paths


def combine_txt(paths, parse_one):
    """
    Parsuje wiele plikow TXT (przerwany pomiar) i skleja je w jeden ciag.

    - naglowek (czujnik/czujniki) bierzemy z PIERWSZEGO pliku,
    - wiersze laczymy w kolejnosci plikow (juz posortowanych chronologicznie),
    - duplikaty po znaczniku czasu (kolumna A) sa pomijane — usuwa to
      nakladajace sie probki na styku dwoch plikow.

    parse_one to parse_txt (CC) albo parse_txt_cc04 (CC-04); obie zwracaja
    (header, rows), wiec ta funkcja dziala dla obu formatow.
    """
    header = None
    combined = []
    seen_ts = set()
    for i, p in enumerate(paths):
        h, rows = parse_one(p)
        if i == 0:
            header = h
        added = 0
        for row in rows:
            ts = row[0].strip() if row and row[0] else ''
            if ts and ts in seen_ts:
                continue          # duplikat czasu (styk plikow) — pomijamy
            if ts:
                seen_ts.add(ts)
            combined.append(row)
            added += 1
        if len(paths) > 1:
            print(f"  + {os.path.basename(p)}: {len(rows)} wierszy, dodano {added}")
    return header, combined


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
    CC-04: krotka (dt, Tzad, RHzad, roztdp, l1, l2, l3, l4, Todcz, RHodcz)
    gdzie l1..l4 = rozrzuty 15-min czterech kanalow GLOWNYCH (do kryterium L).
    Indeksy w `rows` dobierane PO NAZWIE z CC04_KOLUMNY (pelny uklad pliku), by dzialalo
    niezaleznie od liczby/kolejnosci kanalow (obserwacja pokazuje wszystkie, analiza — glowne).
    """
    idx = {name: i for i, name in enumerate(CC04_KOLUMNY)}
    i_dt   = idx['Data Czas']
    i_tz   = idx['Tzadana']
    i_rhz  = idx['RHzadana']
    i_tod  = idx['Todczytana']
    i_rho  = idx['RHodczytana']
    i_rtdp = idx['roztdp(15min)']
    i_rl   = [idx[f'roztempCh{ch}'] for ch in CC04_KANALY_GLOWNE]   # 4 kanaly glowne

    def g(row, i):
        return row[i] if 0 <= i < len(row) else None

    result = []
    for row in rows:
        result.append((
            _s_to_dt(g(row, i_dt)),
            _s_to_float(g(row, i_tz)),
            _s_to_float(g(row, i_rhz)),
            _s_to_float(g(row, i_rtdp)),
            _s_to_float(g(row, i_rl[0])),
            _s_to_float(g(row, i_rl[1])),
            _s_to_float(g(row, i_rl[2])),
            _s_to_float(g(row, i_rl[3])),
            _s_to_float(g(row, i_tod)),   # Todczytana (odczyt komory)
            _s_to_float(g(row, i_rho)),   # RHodczytana (odczyt komory)
        ))
    return result


def _oblicz_start_okna(data, i, j, idx_t_odcz, idx_rh_odcz, min_stable, min_po,
                       prog_temp, prog_rh):
    """
    Zwraca start_idx okna analizy w segmencie stabilnym data[i:j] albo None.

    Zasada:
      1. Szukamy chwili wejscia ODCZYTOW komory w widelki wokol nastaw:
           • temperatura: |Todczytana - Tzadana| <= prog_temp (°C, bezwzglednie),
           • wilgotnosc (punkt z RH): |RHodczytana - RHzadana| w granicach prog_rh %.
         Punkt tylko-temperatura: liczy sie samo wejscie temperatury.
      2. Okno startuje min_po (2h) PO wejsciu. Gdy odczyty nie weszly w widelki —
         start liczony od min_stable (rozgrzewka) od poczatku punktu.
      3. Punkt trzymany jest zwykle ~2h, wiec samo "2h po wejsciu" czesto wypadaloby
         dokladnie na koncu punktu (okno = 0). Dlatego start jest cofany tak, by na
         koncu punktu zostal ZAWSZE ogon pomiarowy MIN_OKNO_ANALIZY (min. 5 minut do
         wyboru reprezentantow). To odtwarza dawne zachowanie (pomiar na koncu punktu).
      4. Punkty krotsze niz min_stable (przejsciowe/przerwane) sa pomijane (None).
    """
    seg_dt0    = data[i][0]
    seg_end_dt = data[j - 1][0]
    seg_t      = data[i][1]                   # Tzadana (nastawa temperatury)
    seg_c      = data[i][2]                   # RHzadana (nastawa wilgotnosci)
    if seg_dt0 is None or seg_end_dt is None:
        return None
    if seg_end_dt - seg_dt0 < min_stable:
        return None                           # punkt zbyt krotki — pomijamy

    temp_only = (seg_c is None or seg_c == 0.0)

    entry_dt = None
    for ki in range(i, j):
        t_odcz = data[ki][idx_t_odcz]         # Todczytana (odczyt komory)
        if seg_t is None or t_odcz is None or abs(seg_t - t_odcz) > prog_temp:
            continue                          # temperatura poza widelkami
        if temp_only:
            entry_dt = data[ki][0]            # sama temperatura w widelkach
            break
        rh = data[ki][idx_rh_odcz]            # RHodczytana (odczyt komory)
        if rh is not None and not _rozne_procent(seg_c, rh, prog_rh):
            entry_dt = data[ki][0]            # OBA (T i RH) w widelkach
            break

    if entry_dt is not None:
        start_time = entry_dt + min_po        # 2h po wejsciu w widelki
    else:
        start_time = seg_dt0 + min_stable     # brak wejscia — rozgrzewka od poczatku

    # Nie "zjadaj" calego punktu: zostaw ogon pomiarowy na koncu.
    latest = seg_end_dt - MIN_OKNO_ANALIZY
    if start_time > latest:
        start_time = latest
    if start_time < seg_dt0:
        start_time = seg_dt0

    for ki in range(i, j):
        if data[ki][0] and data[ki][0] >= start_time:
            return ki
    return None


def _find_all_analysis_windows(data, min_stable, idx_t_odcz=5, idx_rh_odcz=6,
                               min_po=None, prog_temp=None, prog_rh=None):
    """
    Dla każdego segmentu stabilnego B+C wyznacza okno analizy (start_idx, end_idx),
    gdzie end_idx = pierwszy wiersz ze zmianą B lub C (wyłącznie).

    Start okna liczy _oblicz_start_okna: okno rusza min_po (2h) od wejścia odczytów w
    widełki (T: |Todczytana-Tzadana| <= prog_temp °C; RH: prog_rh % od RHzadana). Dla
    tylko-temp liczy się sama temperatura. Fallback: po min_stable od początku.
    idx_t_odcz / idx_rh_odcz = indeksy Todczytana / RHodczytana w krotce (5/6 dla CC, 8/9 dla CC-04).
    """
    if min_po is None:
        min_po = STABILIZACJA_PO_RH
    if prog_temp is None:
        prog_temp = PROG_WEJSCIA_TEMP
    if prog_rh is None:
        prog_rh = PROG_WEJSCIA_RH_PROC

    windows = []
    n = len(data)
    i = 0
    while i < n:
        seg_b, seg_c, seg_dt0 = data[i][1], data[i][2], data[i][0]
        if seg_b is None or seg_c is None or seg_dt0 is None:
            i += 1
            continue

        # Koniec bieżącego segmentu stabilnego (stałe B i C)
        j = i + 1
        while j < n:
            b, c = data[j][1], data[j][2]
            if b != seg_b or c != seg_c:
                break
            j += 1

        start_idx = _oblicz_start_okna(data, i, j, idx_t_odcz, idx_rh_odcz,
                                       min_stable, min_po, prog_temp, prog_rh)
        if start_idx is not None:
            windows.append((start_idx, j))

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


def _reps_ostatnie_minuty(data, start_idx, end_idx, k=5):
    """
    Awaryjny wybor reprezentantow: po jednym wierszu na minute, k OSTATNICH minut
    okna (najblizej konca punktu). Uzywane gdy nie ma 5 kolejnych minut spelniajacych
    kryteria — bierzemy po prostu najblizsze po czasie (koniec punktu = najstabilniej).
    """
    by_min = {}
    for i in range(start_idx, end_idx):
        dt = data[i][0]
        if dt is None:
            continue
        m = dt.replace(second=0, microsecond=0)
        cur = by_min.get(m)
        if cur is None or dt.second < data[cur][0].second:
            by_min[m] = i
    minuty = sorted(by_min)
    if not minuty:
        return None
    return [by_min[m] for m in minuty[-k:]]


def _fallback_reprezentanci(data, start_idx, end_idx, score_fn):
    """
    Wybor reprezentantow BEZ twardych kryteriow (K/L, filtr nastawy) — dla punktow,
    ktore kryterium NIE przeszly, ale i tak chcemy je pokazac/wpisac (na pomaranczowo).
    Najpierw najlepsze 5 kolejnych minut wg score_fn; gdy sie nie da — 5 ostatnich minut.
    """
    valid = [i for i in range(start_idx, end_idx)
             if data[i][0] is not None and score_fn(i) is not None]
    reps = _find_best_minute_reps(data, valid, start_idx, end_idx, score_fn)
    if reps:
        return reps
    return _reps_ostatnie_minuty(data, start_idx, end_idx, 5)


def _oznacz_blok(ws, data, start_idx, end_idx, rep_indices, n_cols,
                 fill_light, fill_dark, powod=None):
    """
    Koloruje 5-minutowy blok (fill_light) + reprezentantow (fill_dark, pogrubienie).
    Gdy podano `powod` (punkt nie przeszedl kryterium) — dopisuje komentarz w komorce
    NA PRAWO od bloku (kolumna n_cols+2, wiersz pierwszego reprezentanta).
    Zwraca liste wierszy jasnego bloku.
    """
    rep_times  = [data[i][0] for i in rep_indices]
    blok_start = rep_times[0].replace(second=0, microsecond=0)
    blok_end   = blok_start + datetime.timedelta(minutes=5)
    blok = [i for i in range(start_idx, end_idx)
            if data[i][0] is not None and blok_start <= data[i][0] <= blok_end]
    for i in blok:
        _apply_fill(ws, 2 + i, fill_light, n_cols)
    for i in rep_indices:
        _apply_fill(ws, 2 + i, fill_dark, n_cols)
        _apply_bold(ws, 2 + i, n_cols)
    if powod:
        kom = ws.cell(row=2 + rep_indices[0], column=n_cols + 1)
        kom.value = f"UWAGA (nie na zielono): {powod}. Wybrano 5 wierszy najblizszych czasowo."
        kom.fill  = fill_dark
        fo = kom.font
        kom.font = Font(name=fo.name, size=fo.size, bold=True, italic=True, color='9C5700')
    return blok


def _process_segment(ws, data, start_idx, end_idx, seg_num, file_type='CC'):
    """
    Analizuje jeden segment i koloruje go w arkuszu obserwacji:
      • kryteria spelnione    -> ZIELONY (jasny blok + ciemni reprezentanci),
      • kryteria niespelnione -> POMARANCZOWY + komentarz z POWODEM (punktu NIE pomijamy;
        wybieramy najblizsze po czasie).
    Zwraca krotke (rep_indices, powod):
      powod = None  -> punkt zielony (OK),
      powod = tekst -> punkt pomaranczowy (przyczyna),
      (None, None)  -> brak jakichkolwiek danych do oznaczenia.
    """
    t0 = data[start_idx][0]
    t1 = data[end_idx - 1][0]
    b  = data[start_idx][1]
    c  = data[start_idx][2]
    print(f"  Segment {seg_num}: Excel row {2+start_idx}–{2+end_idx-1}  "
          f"T={b} RH={c}  ({t0} – {t1})")

    is_cc04    = (file_type == 'CC04')
    temp_only  = (c is not None and c == 0.0)
    n_cols     = len(CC04_KOLUMNY) if is_cc04 else 12
    n_l        = '4L' if is_cc04 else 'L'
    crit_label = n_l if temp_only else f"K+{n_l}"   # skrot techniczny (do konsoli)
    # Opis warunku stabilnosci po polsku (do komentarza w arkuszu — zrozumialy dla operatora)
    if temp_only:
        warunek_txt = "rozrzut temperatury czujnikow w 15 min <= 0,1 st.C"
    else:
        warunek_txt = "rozrzut temperatury czujnikow w 15 min <= 0,1 st.C oraz rozrzut punktu rosy w normie"

    # Score function (wspolna dla sciezki normalnej i awaryjnej)
    if temp_only:
        if is_cc04:
            def score_fn(i):
                vals = [v for v in data[i][4:8] if v is not None]
                return max(vals) if vals else None
        else:
            score_fn = lambda i: data[i][4]
    else:
        score_fn = lambda i: data[i][3]   # K (roztdp) dla CC i CC-04

    powod       = None    # None => zielony (OK); tekst => pomaranczowy (przyczyna)
    rep_indices = None

    # 1) Filtr zgodnosci nastawy z odczytem komory (mediana z okna vs nastawa)
    if FILTR_NASTAWA_ODCZYT:
        d_idx, e_idx = (8, 9) if is_cc04 else (5, 6)
        d_vals = [data[i][d_idx] for i in range(start_idx, end_idx) if data[i][d_idx] is not None]
        e_vals = [data[i][e_idx] for i in range(start_idx, end_idx) if data[i][e_idx] is not None]
        d_med  = statistics.median(d_vals) if d_vals else None
        e_med  = statistics.median(e_vals) if e_vals else None
        odrzuc_t  = _rozne_procent(b, d_med, MAX_ROZNICA_PROCENT)
        odrzuc_rh = (not temp_only) and _rozne_procent(c, e_med, MAX_ROZNICA_PROCENT)
        if odrzuc_t or odrzuc_rh:
            czesci = []
            if odrzuc_t:
                czesci.append(f"temperatura: nastawa {b} st.C, a komora ~{round(d_med, 2) if d_med is not None else '-'} st.C")
            if odrzuc_rh:
                czesci.append(f"wilgotnosc: nastawa {c}%, a komora ~{round(e_med, 2) if e_med is not None else '-'}%")
            powod = ("komora nie osiagnela nastawy — " + "; ".join(czesci)
                     + f" (roznica > {MAX_ROZNICA_PROCENT:g}%)")

    # 2) Wiersze spelniajace kryteria K/L (tylko jesli filtr nie odrzucil)
    if powod is None:
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

        print(f"    Wiersze spelniajace kryteria {crit_label}: {len(valid_indices)}")
        if not valid_indices:
            powod = f"punkt niestabilny — w zadnym momencie nie spelniono warunku: {warunek_txt}"
        else:
            rep_indices = _find_best_minute_reps(data, valid_indices, start_idx, end_idx, score_fn)
            if rep_indices is None:
                powod = f"komora za malo stabilna — brak 5 kolejnych minut, gdzie {warunek_txt}"

    # 3a) ZIELONY — kryteria spelnione
    if powod is None:
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

        blok = _oznacz_blok(ws, data, start_idx, end_idx, rep_indices, n_cols,
                            FILL_LIGHT, FILL_DARK, powod=None)
        if blok:
            print(f"    Wiersze jasno-zielone: {len(blok)} (Excel {2+blok[0]}–{2+blok[-1]})")
        return rep_indices, None

    # 3b) POMARANCZOWY — kryterium niespelnione: nie pomijamy, wybieramy najblizsze po czasie
    fb = _fallback_reprezentanci(data, start_idx, end_idx, score_fn)
    if not fb:
        print(f"    [POMINIETO] {powod} — brak nawet danych zastepczych.")
        return None, None
    _oznacz_blok(ws, data, start_idx, end_idx, fb, n_cols,
                 FILL_WARN_LIGHT, FILL_WARN_DARK, powod=powod)
    print(f"    [POMARANCZOWY] {powod} — oznaczono {len(fb)} reprezentantow (komentarz w kol. {n_cols+2}).")
    return fb, powod


def analyze_and_highlight(ws, rows, file_type='CC'):
    """
    Dla każdego segmentu stabilnego B+C (>= STABILIZACJA_MIN):
      - koloruje najlepsze 5-minutowe okno na #E2EFDA
      - oznacza 5 wierszy reprezentacyjnych (#A9D08E + bold)
    """
    data = _prepare_data_cc04(rows) if file_type == 'CC04' else _prepare_data(rows)

    # pozycje odczytow komory w krotce: Todczytana / RHodczytana
    idx_t_odcz  = 8 if file_type == 'CC04' else 5
    idx_rh_odcz = 9 if file_type == 'CC04' else 6
    windows = _find_all_analysis_windows(data, STABILIZACJA_MIN, idx_t_odcz, idx_rh_odcz)
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
    ostrzezenia = 0
    for num, (start_idx, end_idx) in enumerate(filtered, 1):
        rep_idx, powod = _process_segment(ws, data, start_idx, end_idx, num, file_type)
        if rep_idx is not None:
            found += 1
            if powod:
                ostrzezenia += 1
            segments.append((rep_idx, powod))
    print(f"  Oznaczono segmentow: {found}/{len(filtered)} "
          f"(zielonych: {found - ostrzezenia}, pomaranczowych: {ostrzezenia})")
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
    for rep_indices, _powod in rep_groups:
        tts = [_s_to_dt(rows_obs[i][0]) for i in rep_indices]
        punkty.append(None if any(t is None for t in tts) else tts)

    oznaczenia_per_plik = {}   # fname -> set(row_idx)
    dev = 0                    # numer przyrzadu (0-based) -> para kolumn w prawo
    uzyte = []                 # per przyrzad (w kolejnosci dev): (serial, temps, rhs)

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
        temps = [k[1] for m in dopasowania.values() for k in m if k[1] is not None]
        rhs   = [k[2] for m in dopasowania.values() for k in m if k[2] is not None]
        uzyte.append((_serial_z_wyniku(fname), temps, rhs))
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
                temps2 = [k[4] for m in dopasowania.values() for k in m if k[4] is not None]
                uzyte.append((_serial_z_wyniku(fname), temps2, []))
                dev += 1

    if dev == 0:
        print(f"  [WYNIKI] Zaden plik nie pasowal czasowo do punktow — brak danych srodowiskowych.")

    # Oznacz wiersze w plikach wynikow
    for fname, row_set in oznaczenia_per_plik.items():
        _oznacz_wyniki_xlsx(os.path.join(WYNIKI_FOLDER, fname), sorted(row_set))

    return uzyte


def _serial_z_wyniku(fname):
    """'TMM230200349_wynik.xlsx' -> 'TMM230200349' (nr fabryczny do dopasowania z PZ)."""
    base = os.path.splitext(fname)[0]
    return re.sub(r'_wynik$', '', base, flags=re.I)


# Kolumny Strony 2 (tabela przyrzadow, od wiersza 11):
#  B=2 wytworca(obiekt) C=3 wytworca(czujnik) D=4 typ E=5 nr fabr F=6 nr ewid
#  G=7 adres H=8 typ(czujnik) I=9 nr fabr(czujnik) J=10 nr ewid(czujnik)
#  K=11 rozdz. t  L=12 rozdz. RH  O=15 nr zlecenia
STRONA2_PIERWSZY_WIERSZ = 11


def wypelnij_strone2_z_pz(ws2, uzyte, pz_mapa, zest):
    """
    Wypelnia tabele przyrzadow na Stronie 2 protokolu na podstawie PZ.

    `uzyte` = lista (serial, temps, rhs) w kolejnosci kolumn pomiarowych Strony 3
    (i-ty przyrzad -> wiersz 11+i). Dopasowanie przyrzadu z PZ po nr fabrycznym
    (serial z nazwy pliku wyniku). Rozdzielczosc: z Zestawienia (po producencie+typie),
    a gdy brak — z wahania cyfr po przecinku w danych pomiarowych.
    Przyrzady bez dopasowania w PZ (np. mierniki reczne bez pliku logera) zostaja
    do recznego uzupelnienia (log ostrzegawczy).
    """
    if not uzyte:
        return
    print("\n  Wypelnianie tabeli przyrzadow (Strona 2) z PZ...")
    for i, (serial, temps, rhs) in enumerate(uzyte):
        w = STRONA2_PIERWSZY_WIERSZ + i
        dev = pz_mapa.get(pz_dane.normalizuj_serial(serial)) if pz_mapa else None
        if dev is None:
            print(f"    wiersz {w}: brak dopasowania w PZ (serial '{serial}') — uzupelnij recznie.")
            continue

        # Rozdzielczosc: Zestawienie -> fallback z danych
        t_res, rh_res = pz_dane.rozdzielczosc_zestawienie(zest, dev.wytworca, dev.typ)
        zrodlo = "Zestawienie"
        if t_res is None:
            t_res = pz_dane.rozdzielczosc_z_kolumny(temps); zrodlo = "dane"
        if rh_res is None:
            rh_res = pz_dane.rozdzielczosc_z_kolumny(rhs)
            if rh_res is None:       # przyrzad tylko-temperatura — brak danych RH
                rh_res = t_res

        def _set(col, val):
            ws2.cell(row=w, column=col).value = val if (val not in (None, "")) else "-"

        _set(2, dev.wytworca)            # B
        _set(3, dev.czuj_wytworca)       # C
        _set(4, dev.typ)                 # D
        _set(5, dev.nr_fabr)             # E
        _set(6, dev.nr_ewid)             # F
        _set(7, "")                      # G (adres) — brak w PZ
        _set(8, dev.czuj_typ)            # H
        _set(9, dev.czuj_nr_fabr)        # I
        _set(10, "")                     # J (nr ewid czujnika) — brak w PZ
        ws2.cell(row=w, column=11).value = t_res    # K
        ws2.cell(row=w, column=12).value = rh_res   # L
        ws2.cell(row=w, column=15).value = dev.nr_zlecenia   # O
        print(f"    wiersz {w}: {dev.wytworca} / {dev.typ} / {dev.nr_fabr} "
              f"(zlec {dev.nr_zlecenia}, K={t_res} L={rh_res} [{zrodlo}])")


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
        print('  [Excel COM] Brak pywin32 (win32com) — nie moge wstawic blokow.')
        return False

    xl = wb = None
    try:
        # DEDYKOWANA, nowa instancja Excela (DispatchEx) — NIE doczepiamy sie do
        # ewentualnego otwartego Excela uzytkownika (zajety/modalny -> wyjatek i tylko
        # 6 punktow w protokole). Fallback do Dispatch, gdyby DispatchEx zawiodl.
        try:
            xl = win32com.client.DispatchEx('Excel.Application')
        except Exception:
            xl = win32com.client.Dispatch('Excel.Application')
        xl.Visible         = False
        xl.DisplayAlerts   = False
        xl.ScreenUpdating  = False
        try:
            xl.AskToUpdateLinks = False
            xl.EnableEvents     = False
        except Exception:
            pass

        # UpdateLinks=0: nie aktualizuj linkow zewn. przy otwarciu (unika siegania do sieci).
        wb  = xl.Workbooks.Open(os.path.abspath(filepath), UpdateLinks=0, ReadOnly=False)
        ws  = None
        for idx in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(idx).Name == sheet_name:
                ws = wb.Sheets(idx)
                break
        if ws is None:
            print(f'  [Excel COM] Brak arkusza "{sheet_name}" — nie wstawiam blokow.')
            return False

        addr = f'{block_start}:{block_start + block_size - 1}'
        for _ in range(n_extra):
            ws.Rows(addr).Copy()
            ws.Rows(addr).Insert(Shift=-4121)   # xlShiftDown
            xl.CutCopyMode = False

        wb.Save()
        return True

    except Exception as exc:
        print(f'  [Excel COM] Blad wstawiania blokow: {type(exc).__name__}: {exc}')
        return False
    finally:
        if wb is not None:
            try: wb.Close(False)
            except: pass
        if xl is not None:
            try: xl.Quit()
            except: pass


def _zapisz_bezpiecznie(wb, path, opis="plik"):
    """
    Zapisuje skoroszyt; gdy plik jest OTWARTY w Excelu (zablokowany) — daje czytelny
    komunikat zamiast surowego PermissionError. To najczestsza przyczyna „braku danych":
    plik otwarty w Excelu w trakcie pracy skryptu przerywa zapis.
    """
    try:
        wb.save(path)
    except PermissionError:
        raise PermissionError(
            f"\n  Nie moge zapisac ({opis}):\n    {path}\n"
            f"  >>> Plik jest OTWARTY w Excelu. ZAMKNIJ go i uruchom skrypt ponownie. <<<\n"
            f"  (Otwarcie pliku w trakcie pracy skryptu przerywa zapis danych — stad „pusty\" protokol.)"
        )


def _round_hm(dt):
    """Zaokrągla datetime do pelnej minuty i zwraca string HH:MM."""
    if dt is None:
        return None
    if dt.second >= 30:
        dt = dt + datetime.timedelta(minutes=1)
    return dt.strftime('%H:%M')


def generuj_protokol(rep_groups, rows, measurement_id, obs_type, sensor_names=None,
                     pz_mapa=None, zest=None):
    """
    Tworzy plik protokolu na podstawie reprezentacyjnych wierszy z obserwacji.

    rep_groups : lista krotek (list[5 int], powod) – indeksy reprezentantow (0-based
                 do `rows`) dla kazdego punktu; powod=None => zielony (OK),
                 powod=tekst => pomaranczowy (kryterium niespelnione, wpisany mimo to)
    rows       : surowe wiersze danych z pliku TXT (list of list[str])
    obs_type   : 'CC' lub 'CC04'
    """
    TEMPLATE_BLOCKS = 6     # liczba blokow (punktow) w szablonie
    BLOCK_START_ROW = 20    # pierwszy wiersz Excel pierwszego bloku
    BLOCK_SIZE      = 5     # wierszy na jeden punkt

    if obs_type == 'CC04':
        tmpl_name   = PROTOKOL_CC04_TEMPLATE
        # Do protokolu ida TYLKO kanaly GLOWNE — indeksy odczytow ChNNN + tdp w rows[]
        # (dobierane po nazwie z pelnego ukladu CC04_KOLUMNY). K,L,M,N,O = 4 kanaly + tdp.
        _idx        = {name: i for i, name in enumerate(CC04_KOLUMNY)}
        src_indices = [_idx[f'Ch{ch}'] for ch in CC04_KANALY_GLOWNE] + [_idx['tdp']]
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

    for punkt_idx, (rep_indices, powod) in enumerate(rep_groups):
        r0 = BLOCK_START_ROW + punkt_idx * BLOCK_SIZE

        # Numer punktu w kolumnie A (A20=1, A25=2, ...). Wstawianie blokow kopiuje
        # szablonowy „1" do kazdego bloku — nadpisujemy poprawnym numerem kolejnym.
        ws3.cell(row=r0, column=1).value = punkt_idx + 1

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

        # Punkt NIE przeszedl kryterium (w obserwacji pomaranczowy) — oznacz TAK SAMO
        # w protokole: kolumna B (nastawa T) calego bloku na pomaranczowo + nota z powodem
        # w kolumnie poza obszarem druku (widoczna operatorowi w Excelu).
        if powod:
            for row_off in range(BLOCK_SIZE):
                ws3.cell(row=r0 + row_off, column=2).fill = FILL_WARN_DARK
            nota = ws3.cell(row=r0, column=20)
            nota.value = f"UWAGA (nie na zielono): {powod}"
            fo = nota.font
            nota.font = Font(name=fo.name, size=fo.size, bold=True, italic=True, color='9C5700')
            print(f"    [PROTOKOL] Punkt {punkt_idx+1} oznaczony na pomaranczowo: {powod}")

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
    uzyte = _wypelnij_wyniki_srodowiskowe(ws3, rep_groups, rows, obs_type)

    # Tabela przyrzadow (Strona 2) z PZ — dopasowanie po nr fabrycznym do kolumn Strony 3.
    if 'Strona 2' in proto_wb.sheetnames:
        wypelnij_strone2_z_pz(proto_wb['Strona 2'], uzyte, pz_mapa, zest)
    else:
        print("  [PZ] Brak arkusza 'Strona 2' — pomijam tabele przyrzadow.")

    _zapisz_bezpiecznie(proto_wb, out_path, "protokol")
    print(f"Zapisano protokol: {out_name}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    txt_files = resolve_txt_files()
    first_name = os.path.basename(txt_files[0])
    measurement_id = parse_measurement_id(first_name)

    # Typ pliku i naglowek wyznaczamy z pierwszego pliku (przerwany pomiar to
    # ten sam typ i ten sam czujnik we wszystkich czesciach).
    raw_lines = open_txt(txt_files[0])
    file_type = detect_file_type(raw_lines)

    # Dane przyrzadow z PZ wczytujemy NA STARCIE (przed analiza), zeby byly gotowe
    # do wypelnienia Strony 2 przy budowie protokolu.
    print("Wczytywanie danych przyrzadow (PZ + Zestawienie)...")
    pz_mapa, _pz_lista = pz_dane.wczytaj_pz(PZ_FOLDER)
    zest = pz_dane.wczytaj_zestawienie(ZESTAWIENIE_PLIK)

    print(f"Numer pomiaru : {measurement_id}")
    if len(txt_files) == 1:
        print(f"Plik wejsciowy: {first_name}")
    else:
        print(f"Pliki wejsciowe ({len(txt_files)}):")
        for p in txt_files:
            print(f"    • {os.path.basename(p)}")
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

        sensor_names, rows = combine_txt(txt_files, parse_txt_cc04)
        print(f"Czujniki wzorcowe: {sensor_names}")
        print(f"Wierszy danych   : {len(rows)}")

        wb = openpyxl.load_workbook(output_path)
        ws = wb.active

        # PELNY uklad jak w pliku multimetru: WSZYSTKIE kanaly po kolei (A..AG = 33 kol).
        # Analiza i protokol biora tylko kanaly GLOWNE (po nazwie); reszta jest do wgladu.
        N_KOL = len(CC04_KOLUMNY)
        kanal_pt = _mapa_kanal_pt(raw_lines)

        # Podpisy z szablonu (X92/X93) + ich STYL (ramka, format daty) czytamy PRZED
        # nadpisaniem danymi. W pelnym ukladzie kol. 24-25 to juz dane, wiec podpisy,
        # ramke i date przenosimy na PRAWO od danych, a stare komorki czyscimy (inaczej
        # zostaje pusta ramka i data pokazuje np. „19.02.1900" na miejscu danych).
        sig1 = ws.cell(row=92, column=24).value
        sig2 = ws.cell(row=93, column=24).value
        sig_border = _copy_obj(ws.cell(row=92, column=24).border)
        date_nf    = ws.cell(row=92, column=25).number_format or 'yyyy-mm-dd'

        # Dane pomiarowe A2:AG(ostatni) — wszystkie kolumny pliku
        for r_i, row in enumerate(rows):
            for c_i in range(min(len(row), N_KOL)):
                ws.cell(row=2 + r_i, column=1 + c_i).value = to_value(row[c_i])

        # Naglowki (rzad 1) — z nazwami czujnikow (Pt100-XX) zamiast surowych 'ChNNN'
        for c_i, lbl in enumerate(_naglowki_cc04(kanal_pt)):
            ws.cell(row=1, column=1 + c_i).value = lbl
        print(f"  Czujniki glowne : {sensor_names}")
        print(f"  Czujniki zapasowe: {[kanal_pt.get(c) or c for c in CC04_KANALY_ZAPASOWE]}")

        # Podpisy + data PRZENIESIONE na prawo od danych (z ramka jak w oryginale)
        SIG_COL = N_KOL + 3
        for i, nazwisko in enumerate((sig1, sig2)):
            cn = ws.cell(row=92 + i, column=SIG_COL)
            cn.value  = nazwisko
            cn.border = _copy_obj(sig_border)
            cd = ws.cell(row=92 + i, column=SIG_COL + 1)
            cd.value  = today
            cd.border = _copy_obj(sig_border)
            cd.number_format = date_nf

        # Wyczysc STARE komorki podpisow (kol 24-25, w. 92-93) — teraz to zwykle dane:
        # bez ramki i bez formatu daty.
        _pusta_ramka = Border()
        for _r in (92, 93):
            for _col in (24, 25):
                _oc = ws.cell(row=_r, column=_col)
                _oc.border = _pusta_ramka
                _oc.number_format = 'General'

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

        sensor_name, rows = combine_txt(txt_files, parse_txt)
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

    _zapisz_bezpiecznie(wb, output_path, "obserwacje")
    print(f"\nZapisano: {output_name}")

    if rep_groups:
        generuj_protokol(rep_groups, rows, measurement_id, file_type,
                         sensor_names=sensor_names, pz_mapa=pz_mapa, zest=zest)

    print("Gotowe!")


if __name__ == '__main__':
    main()
