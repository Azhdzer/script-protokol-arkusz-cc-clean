# -*- coding: utf-8 -*-
"""
generuj_arkusze.py

Automatyczne tworzenie kopii arkuszy obliczeniowych na podstawie pliku protokołu.

Etap 1 – tworzenie nazwanych kopii szablonu
Etap 2 – zarządzanie zakładkami (zmiana nazw, usuwanie/kopiowanie)
Etap 3 – wypełnianie C15:C19 i D15:D19 (dane stałe z Strona 3, kol. L i M)
Etap 4 – wypełnianie E15:E19 i F15:F19 (dane zmienne per kopia, kol. Q/R i dalej)
Etap 5 – wypełnianie komórek nagłówkowych i stopkowych
         (E4:F4, G6, E5, E6, H57, B228, H228, B230:C230, H230:I230)
Etap 6 – wypełnianie arkusza Wyniki: F24 (per kopia z Strona 3 Q17/S17/…),
         C28, C32 (daty), E28:G28, E32:G32 (podpisy)
Etap 7 – tworzenie kopii dokumentów Word (świadectwa wzorcowania):
         podmiana placeholderów, tabela kalibracyjna D246–G246 per zakładka

UWAGA: openpyxl odczytuje wartości zapisane przez Excel (data_only=True).
Jeśli plik protokołu nie był przeliczony i zapisany w Excelu, niektóre komórki
mogą zwrócić None zamiast wartości formuły.
"""

import os
import shutil
import datetime
import math
from copy import deepcopy
from itertools import groupby
import openpyxl
import xlwings as xw

try:
    from docx import Document as DocxDocument
    _DOCX_OK = True
except ImportError:
    DocxDocument = None  # type: ignore
    _DOCX_OK = False

# =============================================================================
# KONFIGURACJA  ← edytuj tutaj przed uruchomieniem
# =============================================================================

FOLDER           = r"."   # folder z plikami xlsx; "." = ten sam co skrypt
                           # możesz podać pełną ścieżkę, np. r"C:\Moje\Pliki"

PROTOKOL_PLIK    = "116_LA_TH_2026 - protokół CC.xlsx"
SZABLON_PLIK     = "xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.11 z 29.08.2025 - 1 - RH (CC).xlsx"

ARKUSZ_STRONA2   = "Strona 2"   # arkusz z listą kopii do wygenerowania
ARKUSZ_STRONA3   = "Strona 3"   # arkusz z definicją zakładek

START_ROW_S2     = 11    # wiersz startowy w Strona 2 (kolumna A)
START_ROW_S3     = 20    # wiersz startowy w Strona 3 (kolumna A)
BLOK_S3          = 5     # liczba wierszy zajmowanych przez jedną zakładkę w Strona 3

ARKUSZ_CHRONIONY = "Wyniki"   # ta zakładka nigdy nie jest usuwana ani zmieniana

START_COL_E_S3   = 17    # kolumna Q (1-indexed) — źródło dla E15:E19, kopia 1
START_COL_F_S3   = 18    # kolumna R (1-indexed) — źródło dla F15:F19, kopia 1
KROK_COL_EF      = 2     # przesunięcie kolumny dla każdej kolejnej kopii (co 2 w prawo)

ARKUSZ_WNIOSKI   = "Wyniki"          # ostatni arkusz — nie jest modyfikowany
PODPISUJACY_1    = "Artsiom Azhdzer"  # B230:C230 (scalona) — podpisujacy z lewej
PODPISUJACY_2    = "Marek Szpakowski" # H230:I230 (scalona) — podpisujacy z prawej

SZABLON_WORD        = "xxx_yyy_LA_TH_2026 - tylko temp.docx"  # szablon Word; "" = pomiń Etap 7
NR_SW_POCZATKOWY    = 770   # numer świadectwa pierwszej kopii (771, 772, ... dla kolejnych)

# Pliki linkowane wymagane do przeliczenia formul kalibracyjnych (D246/F246/G246).
# Muszą być otwarte w tej samej sesji Excel — podaj dokładne nazwy z rozszerzeniem.
PLIKI_LINKOWANE     = [
    "Obliczenia tdp, RH, C.xls",
    "Wzory.xls",
]

# =============================================================================
# FUNKCJE POMOCNICZE
# =============================================================================

def _cell_to_str(value):
    """
    Konwertuje wartość komórki do stringa.
    - None          → ""
    - float bez ułamka (np. 20.0) → "20"
    - float z ułamkiem (np. 20.3) → "20.3"
    - pozostałe     → str(value).strip()
    """
    if value is None:
        return ""
    if isinstance(value, float):
        return str(int(value)) if value == int(value) else str(value)
    return str(value).strip()


def _round_half_away_from_zero(value):
    """Zaokragla do calkowitych jak w Excel ROUND(..., 0)."""
    if value >= 0:
        return int(math.floor(value + 0.5))
    return int(math.ceil(value - 0.5))


def _cell_to_sheet_name_part(value):
    """
    Konwertuje wartosc do fragmentu nazwy zakladki i zaokragla liczby do calosci.
    Przyklad: 24.9 -> "25", 48.5 -> "49", -19.9 -> "-20".
    """
    if value is None:
        return ""

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, (int, float)):
        return str(_round_half_away_from_zero(float(value)))

    text = str(value).strip()
    if not text:
        return ""

    # Dla tekstow numerycznych (np. "24,9" lub "24.9")
    probe = text.replace(",", ".")
    try:
        num = float(probe)
    except ValueError:
        return text
    return str(_round_half_away_from_zero(num))


_EXCEL_NIEDOZWOLONE_ZNAKI = set(r'[]:*?/\\')


def _unikalne_nazwy_zakladek(dane_zakladek):
    """
    Buduje unikalne nazwy arkuszy zgodne z ograniczeniami Excela:
    - max 31 znakow
    - bez: []:*?/\\
    - bez duplikatow (case-insensitive)
    """
    wynik = []
    zajete = set()

    for i, zd in enumerate(dane_zakladek, start=1):
        raw = _cell_to_str((zd or {}).get("nazwa"))
        base = "".join("_" if ch in _EXCEL_NIEDOZWOLONE_ZNAKI else ch for ch in raw)
        base = base.strip().strip("'")
        if not base:
            base = f"Zakladka {i}"
        base = base[:31]

        candidate = base
        licznik = 2
        while candidate.lower() in zajete:
            suffix = f" ({licznik})"
            max_len = max(1, 31 - len(suffix))
            candidate = base[:max_len].rstrip() + suffix
            licznik += 1

        zajete.add(candidate.lower())
        wynik.append(candidate)

    return wynik


MIESIACE_GEN = {
    1: "stycznia",  2: "lutego",       3: "marca",     4: "kwietnia",
    5: "maja",      6: "czerwca",      7: "lipca",     8: "sierpnia",
    9: "września", 10: "października", 11: "listopada", 12: "grudnia",
}


# =============================================================================
# WCZYTYWANIE DANYCH Z PROTOKOŁU
# =============================================================================

def wczytaj_dane_z_protokolu_s2(sciezka, arkusz, start_row):
    """
    Czyta listę kopii do wygenerowania z Strona 2.

    Kolumna A (1)  – warunek zatrzymania (pusta = koniec)
    Kolumna E (5)  – wartość zastępująca 'RH (CC)' w nazwie kopii
    Kolumna O (15) – wartość zastępująca 'xxx' w nazwie kopii

    Zwraca: [{"O": "133", "E": "1010223"}, ...]
    """
    wb = openpyxl.load_workbook(sciezka, read_only=True, data_only=True)
    if arkusz not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Brak arkusza '{arkusz}' w pliku '{sciezka}'")

    ws = wb[arkusz]
    dane = []
    row = start_row

    while True:
        val_a = ws.cell(row=row, column=1).value
        if val_a is None or str(val_a).strip() == "":
            break

        val_O = ws.cell(row=row, column=15).value   # kolumna O
        val_E = ws.cell(row=row, column=5).value    # kolumna E

        if val_O is None or val_E is None:
            print(f"  [OSTRZEŻENIE] Wiersz {row} w '{arkusz}': brak wartości O lub E — pomijam wiersz.")
            row += 1
            continue

        dane.append({"O": _cell_to_str(val_O), "E": _cell_to_str(val_E)})
        row += 1

    wb.close()
    return dane


def wczytaj_zakladki_z_protokolu_s3(sciezka, arkusz, start_row, blok):
    """
    Czyta definicje zakładek z Strona 3 w blokach po 5 wierszy.

    Dla każdego bloku (startowy wiersz r0):
      - Kolumna A (1) w r0 – warunek zatrzymania (pusta = koniec)
      - Kolumna B (2) w r0 – pierwsza część nazwy zakładki
      - Kolumna C (3) w r0 – druga część nazwy zakładki  → nazwa = "{B}, {C}"
      - Kolumna L (12), wiersze r0..r0+4 – dane dla C15:C19
      - Kolumna M (13), wiersze r0..r0+4 – dane dla D15:D19

    Zwraca: [{"nazwa": "20.3, -", "L_dane": [5 val], "M_dane": [5 val]}, ...]
    """
    wb = openpyxl.load_workbook(sciezka, read_only=True, data_only=True)
    if arkusz not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Brak arkusza '{arkusz}' w pliku '{sciezka}'")

    ws = wb[arkusz]
    wynik = []
    blok_idx = 0

    while True:
        r0 = start_row + blok_idx * blok
        val_a = ws.cell(row=r0, column=1).value
        if val_a is None or str(val_a).strip() == "":
            break

        val_b = _cell_to_sheet_name_part(ws.cell(row=r0, column=2).value)
        val_c = _cell_to_sheet_name_part(ws.cell(row=r0, column=3).value)
        nazwa = f"{val_b}, {val_c}"

        L_dane = [ws.cell(row=r0 + k, column=12).value for k in range(blok)]
        M_dane = [ws.cell(row=r0 + k, column=13).value for k in range(blok)]

        wynik.append({"nazwa": nazwa, "L_dane": L_dane, "M_dane": M_dane})
        blok_idx += 1

    wb.close()
    return wynik


def wczytaj_dane_ef_z_protokolu_s3(sciezka, arkusz, n_kopii, n_zakladek,
                                    start_row, blok,
                                    start_col_e, start_col_f, krok):
    """
    Czyta dane E/F z Strona 3 dla każdej kopii i każdej zakładki.

    Kopia j (0-bazowany):
      kolumna E → start_col_e + j * krok   (Q dla j=0, S dla j=1, U dla j=2, …)
      kolumna F → start_col_f + j * krok   (R dla j=0, T dla j=1, V dla j=2, …)
    Zakładka i (0-bazowany):
      wiersze → start_row + i * blok  ..  start_row + i * blok + (blok-1)

    Zwraca: dane_ef[j][i] = {"E_dane": [5 val], "F_dane": [5 val]}
    """
    wb = openpyxl.load_workbook(sciezka, read_only=True, data_only=True)
    if arkusz not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Brak arkusza '{arkusz}' w pliku '{sciezka}'")

    ws = wb[arkusz]
    dane_ef = []

    for j in range(n_kopii):
        col_e = start_col_e + j * krok
        col_f = start_col_f + j * krok
        zakl = []
        for i in range(n_zakladek):
            r0 = start_row + i * blok
            zakl.append({
                "E_dane": [ws.cell(row=r0 + k, column=col_e).value for k in range(blok)],
                "F_dane": [ws.cell(row=r0 + k, column=col_f).value for k in range(blok)],
            })
        dane_ef.append(zakl)

    wb.close()
    return dane_ef


def wczytaj_wszystko_xlwings(sciezka, ark_s2, ark_s3,
                              start_s2, start_s3, blok,
                              start_col_e, start_col_f, krok):
    """
    Otwiera plik protokołu RAZ przez xlwings (COM Excel) i odczytuje:
      - listę kopii z Strona 2
      - definicje zakładek + dane L/M z Strona 3
      - dane E/F dla każdej kopii z Strona 3
    Używaj tej funkcji zamiast trzech osobnych wywołań opartych na openpyxl,
    gdy protokół zawiera formuły z niezakeszowanymi wartościami.
    Zwraca: (dane_s2, dane_zakladek, dane_ef)
    """
    app = xw.App(visible=False, add_book=False)
    app.api.AutomationSecurity = 1   # msoAutomationSecurityLow — wlacza makra bez pytania
    app.api.DisplayAlerts = False
    app.api.AskToUpdateLinks = False  # aktualizuj linki zewnetrzne bez pytania
    try:
        wb = app.books.open(sciezka)

        # --- Strona 2: lista kopii ---
        ws2 = wb.sheets[ark_s2]
        dane_s2 = []
        row = start_s2
        while True:
            val_a = ws2.cells(row, 1).value
            if val_a is None or str(val_a).strip() == "":
                break
            val_O = ws2.cells(row, 15).value   # kolumna O
            val_E = ws2.cells(row, 5).value    # kolumna E
            if val_O is None or val_E is None:
                print(f"  [OSTRZEŻENIE] Wiersz {row} w '{ark_s2}': brak O lub E — pomijam.")
                row += 1
                continue
            val_B = ws2.cells(row, 2).value   # kolumna B -> E5 kopii
            val_D = ws2.cells(row, 4).value   # kolumna D -> E6 kopii
            val_K = ws2.cells(row, 11).value  # kolumna K -> H57 kopii
            dane_s2.append({
                "O": _cell_to_str(val_O),
                "E": _cell_to_str(val_E),
                "B": val_B,
                "D": val_D,
                "K": val_K,
            })
            row += 1

        # --- Strona 3: zakładki + dane L/M ---
        ws3 = wb.sheets[ark_s3]
        dane_zakladek = []
        blok_idx = 0
        while True:
            r0 = start_s3 + blok_idx * blok
            val_a = ws3.cells(r0, 1).value
            if val_a is None or str(val_a).strip() == "":
                break
            val_b = _cell_to_sheet_name_part(ws3.cells(r0, 2).value)
            val_c = _cell_to_sheet_name_part(ws3.cells(r0, 3).value)
            nazwa = f"{val_b}, {val_c}"
            # Odczyt całego bloku kolumny L i M naraz (szybszy COM)
            L_range = ws3.range(ws3.cells(r0, 12), ws3.cells(r0 + blok - 1, 12)).value
            M_range = ws3.range(ws3.cells(r0, 13), ws3.cells(r0 + blok - 1, 13)).value
            L_dane = L_range if isinstance(L_range, list) else [L_range]
            M_dane = M_range if isinstance(M_range, list) else [M_range]
            k4_val = ws3.cells(r0 + 1, 5).value  # E(r0+1): 2. wiersz bloku, kol. E -> K4 kopii
            dane_zakladek.append({"nazwa": nazwa, "L_dane": L_dane, "M_dane": M_dane, "K4_val": k4_val})
            blok_idx += 1

        # --- Strona 3: dane E/F per kopia ---
        n_kopii = len(dane_s2)
        n_zakladek = len(dane_zakladek)
        dane_ef = []

        if n_kopii > 0 and n_zakladek > 0:
            total_rows = n_zakladek * blok
            for j in range(n_kopii):
                col_e = start_col_e + j * krok
                col_f = start_col_f + j * krok
                # Odczyt całej kolumny E i F dla tej kopii naraz
                e_range = ws3.range(
                    ws3.cells(start_s3, col_e),
                    ws3.cells(start_s3 + total_rows - 1, col_e)
                ).value
                f_range = ws3.range(
                    ws3.cells(start_s3, col_f),
                    ws3.cells(start_s3 + total_rows - 1, col_f)
                ).value
                e_vals = e_range if isinstance(e_range, list) else [e_range]
                f_vals = f_range if isinstance(f_range, list) else [f_range]
                zakl = []
                for i in range(n_zakladek):
                    start = i * blok
                    zakl.append({
                        "E_dane": e_vals[start:start + blok],
                        "F_dane": f_vals[start:start + blok],
                    })
                dane_ef.append(zakl)

        # --- Strona 3: F24 dla arkusza Wyniki, per kopia (wiersz 17) ---
        f24_per_kopia = []
        for j in range(n_kopii):
            col = start_col_e + j * krok
            f24_per_kopia.append(ws3.cells(17, col).value)

        wb.close()
        return dane_s2, dane_zakladek, dane_ef, f24_per_kopia
    finally:
        app.quit()


# =============================================================================
# GENEROWANIE NAZWY KOPII
# =============================================================================

def generuj_nazwe_pliku(szablon_nazwa, wartosc_O, wartosc_E):
    """
    Zastępuje:
      'xxx'     → wartosc_O
      'RH (CC)' → wartosc_E
    w nazwie pliku szablonu.
    """
    nazwa = szablon_nazwa.replace("xxx", wartosc_O, 1)
    nazwa = nazwa.replace("RH (CC)", wartosc_E, 1)
    return nazwa


def _parsuj_nazwe_pliku(nowa_nazwa):
    """
    Parsuje nazwe pliku kopii i zwraca (prefiks, rok, numer_fabryczny, typ).
    Przyklad: '133_LA_TH_2026 - ILAJ 5.4_11#21 - ... - 1010223.xlsx'
              -> ('133', '2026', '1010223', 'ILAJ 5.4_11#21')
    Prefiks         : czesc przed '_LA_TH_'
    Rok             : 4 cyfry po '_LA_TH_'
    Numer fabryczny : ostatni segment po ' - ' (bez rozszerzenia)
    Typ             : pierwszy segment po roku (model/typ przyrzadu pomiarowego)
    """
    nazwa = os.path.splitext(nowa_nazwa)[0]
    if "_LA_TH_" in nazwa:
        prefiks, reszta = nazwa.split("_LA_TH_", 1)
        rok = reszta[:4]
        czesci_po_roku = [p.strip() for p in reszta[4:].split(" - ") if p.strip()]
        typ = czesci_po_roku[0] if czesci_po_roku else ""
    else:
        prefiks = nazwa.split("_")[0]
        rok = str(datetime.datetime.now().year)
        typ = ""
    numer_fabryczny = nazwa.rsplit(" - ", 1)[-1]
    return prefiks, rok, numer_fabryczny, typ


# =============================================================================
# FORMATOWANIE DAT (POLSKIE) I OBSŁUGA DOKUMENTÓW WORD
# =============================================================================

def _formatuj_date(d):
    """Formatuje datetime/date jako 'DD miesiaca YYYY r.' (np. '04 maja 2026 r.')"""
    if isinstance(d, datetime.datetime):
        d = d.date()
    return f"{d.day:02d} {MIESIACE_GEN[d.month]} {d.year} r."


def _formatuj_daty_wzorcowania(daty_raw):
    """
    Formatuje liste dat jako polska date wzorcowania.
    Regula: ciagle 3+ dni → zakres (kreseczka); 1-2 dni → lista.
    Przy przejsciu miesiaca: nazwy obu miesiecy, rok tylko na koncu.

    Przyklady:
      [19.05, 20.05]         → '19, 20 maja 2026 r.'
      [10.05, 11.05, 12.05]  → '10 - 12 maja 2026 r.'
      [29.04, 04.05]         → '29 kwietnia, 04 maja 2026 r.'
      [18.05, 19.05, 21.05]  → '18, 19, 21 maja 2026 r.'
    """
    daty = []
    for d in daty_raw:
        if d is None:
            continue
        if isinstance(d, datetime.datetime):
            d = d.date()
        elif not isinstance(d, datetime.date):
            continue
        daty.append(d)
    daty = sorted(set(daty))
    if not daty:
        return ""

    rok = daty[-1].year

    # Wyznacz segmenty: ciagle biegi >= 3 dni → zakres; reszta → pojedyncze daty
    segmenty = []
    i = 0
    while i < len(daty):
        j = i
        while j + 1 < len(daty) and (daty[j + 1] - daty[j]).days == 1:
            j += 1
        bieg = daty[i:j + 1]
        if len(bieg) >= 3:
            segmenty.append(("zakres", bieg))
        else:
            for d in bieg:
                segmenty.append(("pojedyncza", [d]))
        i = j + 1

    # Buduj pary (rok, miesiac, fragment_tekstowy)
    pary = []
    for typ_seg, daty_seg in segmenty:
        d0 = daty_seg[0]
        dN = daty_seg[-1]
        if typ_seg == "zakres" and d0.month == dN.month:
            pary.append((d0.year, d0.month, f"{d0.day:02d} - {dN.day:02d}"))
        else:
            for d in daty_seg:
                pary.append((d.year, d.month, f"{d.day:02d}"))

    # Grupuj po (rok, miesiac) i scal fragmenty
    grupy = []
    for (yr, mo), grp in groupby(pary, key=lambda x: (x[0], x[1])):
        fragmenty = [item[2] for item in grp]
        grupy.append((yr, mo, ", ".join(fragmenty)))

    # Formatuj wynik: miesiac i rok tylko na ostatniej grupie
    czesci = []
    for idx, (yr, mo, dni_str) in enumerate(grupy):
        if idx < len(grupy) - 1:
            czesci.append(f"{dni_str} {MIESIACE_GEN[mo]}")
        else:
            czesci.append(f"{dni_str} {MIESIACE_GEN[mo]} {yr} r.")
    return ", ".join(czesci)


def _zastap_tekst_w_dok(doc, placeholder, value):
    """
    Zamienia placeholder we wszystkich paragrafach, tabelach i naglowkach/stopkach
    dokumentu Word. Obsluguje placeholder w pojedynczym runie jak i rozbity miedzy runami
    (skleja caly paragraf w pierwszym runie i podmienia tekst).
    """
    value = str(value) if value is not None else ""

    def _zastap_w_para(para):
        if placeholder not in para.text:
            return
        # Proba zamiany w pojedynczym runie (zachowuje formatowanie)
        for run in para.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, value)
                return
        # Placeholder rozbity miedzy runami: skleja caly paragraf w pierwszym runie
        pelny = "".join(r.text for r in para.runs)
        if placeholder in pelny:
            if para.runs:
                para.runs[0].text = pelny.replace(placeholder, value)
                for r in para.runs[1:]:
                    r.text = ""

    def _zastap_w_kontenerze(kontener):
        for para in kontener.paragraphs:
            _zastap_w_para(para)
        for table in kontener.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        _zastap_w_para(para)

    _zastap_w_kontenerze(doc)

    # Naglowki i stopki wszystkich sekcji dokumentu
    for section in doc.sections:
        for hdr_ftr in (
            section.header, section.footer,
            section.even_page_header, section.even_page_footer,
            section.first_page_header, section.first_page_footer,
        ):
            _zastap_w_kontenerze(hdr_ftr)


def _uzupelnij_tabele_kalibracji(doc, punkty):
    """
    Wyszukuje tabele kalibracyjna w dokumencie (po '[wartość_odn_1]'),
    dodaje wiersze jesli jest wiecej niz 3 punktow pomiarowych,
    a nastepnie wypelnia wszystkie komorki wartosciami z 'punkty'.
    punkty = [{"wartosc_odn": ..., "zmierzona": ..., "poprawka": ..., "niepewnosc": ...}, ...]
    """
    cal_table = None
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "[wartość_odn_1]" in cell.text or "[wartosc_odn_1]" in cell.text:
                    cal_table = table
                    break
            if cal_table:
                break
        if cal_table:
            break
    if cal_table is None:
        return  # brak tabeli kalibracyjnej w dokumencie

    n = len(punkty)
    template_row_idx = 2  # 0-indexed: trzeci wiersz szablonu (z placeholderami _3)

    # Dodaj brakujace wiersze jesli jest wiecej niz 3 punktow
    for k in range(3, n):
        src_tr = cal_table.rows[template_row_idx]._tr
        new_tr = deepcopy(src_tr)
        cal_table._tbl.append(new_tr)
        # Podmieniaj numery w placeholderach: _3] → _(k+1)]
        new_row = cal_table.rows[-1]
        for cell in new_row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.text = run.text.replace("_3]", f"_{k + 1}]")

    def _fmt(v):
        if v is None:
            return ""
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v).replace(".", ",")  # separator dziesiętny: przecinek (PL)

    for k, punkt in enumerate(punkty):
        idx = k + 1
        val_odn = _fmt(punkt.get("wartosc_odn"))
        val_zmi = _fmt(punkt.get("zmierzona"))
        val_pop = _fmt(punkt.get("poprawka"))
        val_nie = _fmt(punkt.get("niepewnosc"))
        # Probujemy obie wersje: z polskimi znakami i bez (ASCII), dla pewnosci
        for placeholder, wartosc in [
            (f"[wartość_odn_{idx}]", val_odn),
            (f"[wartosc_odn_{idx}]",  val_odn),
            (f"[zmierzona_{idx}]",    val_zmi),
            (f"[poprawka_{idx}]",     val_pop),
            (f"[niepewność_{idx}]",  val_nie),
            (f"[niepewnosc_{idx}]",   val_nie),
        ]:
            _zastap_tekst_w_dok(doc, placeholder, wartosc)


def generuj_nazwe_word(nr_sw, prefiks, rok):
    """Generuje nazwe pliku Word: '{nr_sw}_{prefiks}_LA_TH_{rok}.docx'"""
    return f"{nr_sw}_{prefiks}_LA_TH_{rok}.docx"


def utworz_kopie_word(folder, szablon_word, dane_s2, kopie_excel,
                      dane_zakladek, dane_kalibracji, nr_sw_poczatkowy):
    """
    Dla kazdej kopii Excel tworzy kopie dokumentu Word i wypelnia placeholdery:
      [data]             aktualna data ('DD miesiaca YYYY r.')
      [nr_sw]            numer swiadectwa (nr_sw_poczatkowy + j)
      [nr_zl]            nr zamowienia (z rekordu Strona 2 kol. D → E6 kopii)
      [nr_fabr]          nr fabryczny  (z nazwy pliku → G6 kopii)
      [wytworca]         wytworca      (z rekordu Strona 2 kol. B → E5 kopii)
      [typ]              typ/model     (segment po roku w nazwie pliku)
      [data_wzorcowania] daty pomiaru  (K4_val zakładek, sformatowane)
      [Podpis]           PODPISUJACY_2
      tabela kalibracyjna: D246/E246/F246/G246 per zakladka robocza
    Zwraca liste nazw utworzonych plikow Word.
    """
    if not _DOCX_OK:
        print("  [BLAD] Brak biblioteki 'python-docx'. Zainstaluj: pip install python-docx")
        return []

    szablon_path = os.path.join(folder, szablon_word)
    n = len(kopie_excel)
    nazwy_word = []

    dzis = datetime.datetime.now()
    data_str = _formatuj_date(dzis)

    # Daty wzorcowania wspolne dla calego protokolu (z K4_val zakładek)
    daty_k4 = [zd.get("K4_val") for zd in dane_zakladek]
    data_wzorcowania = _formatuj_daty_wzorcowania(daty_k4)

    for j, (rekord, nowa_nazwa_xlsx) in enumerate(zip(dane_s2, kopie_excel)):
        nr_sw = nr_sw_poczatkowy + j
        prefiks, rok, nr_fab, typ = _parsuj_nazwe_pliku(nowa_nazwa_xlsx)
        nowa_nazwa_docx = generuj_nazwe_word(nr_sw, prefiks, rok)
        sciezka_docx = os.path.join(folder, nowa_nazwa_docx)

        shutil.copy2(szablon_path, sciezka_docx)
        doc = DocxDocument(sciezka_docx)

        for placeholder, wartosc in {
            "[data]":             data_str,
            "[nr_sw]":            str(nr_sw),
            "[nr_zl]":            prefiks,          # prefiks z nazwy pliku (przed _LA_TH_)
            "[nr_fabr]":          nr_fab,
            "[wytworca]":         str(rekord.get("B") or ""),
            "[typ]":              str(rekord.get("D") or ""),  # E6 pierwszej zakladki
            "[data_wzorcowania]": data_wzorcowania,
            "[Podpis]":           PODPISUJACY_2,
        }.items():
            _zastap_tekst_w_dok(doc, placeholder, wartosc)

        punkty = dane_kalibracji[j] if j < len(dane_kalibracji) else []
        _uzupelnij_tabele_kalibracji(doc, punkty)

        doc.save(sciezka_docx)
        nazwy_word.append(nowa_nazwa_docx)
        print(f"  [Word {j + 1:>{len(str(n))}}/{n}] {nowa_nazwa_docx}")

    return nazwy_word


# =============================================================================
# ZARZĄDZANIE ZAKŁADKAMI I WYPEŁNIANIE KOMÓREK
# =============================================================================

def _dopasuj_liczbe_zakładek(wb, working, n, chroniony, first_ws_name):
    """
    Usuwa nadmiarowe zakładki robocze (od końca) lub dodaje kopie
    pierwszej zakładki roboczej gdy brakuje.
    Zwraca zaktualizowaną listę nazw zakładek roboczych.
    """
    if len(working) > n:
        # Usuń nadmiarowe od końca listy
        for ws_name in working[n:]:
            del wb[ws_name]
        working = working[:n]

    elif len(working) < n:
        # Dodaj kopie pierwszej zakładki roboczej
        first_ws = wb[first_ws_name]
        for _ in range(n - len(working)):
            new_ws = wb.copy_worksheet(first_ws)
            # Przesuń nową zakładkę o jedną pozycję w lewo (tuż przed Wyniki)
            if chroniony in wb.sheetnames:
                wb.move_sheet(new_ws.title, offset=-1)
            working.append(new_ws.title)

    return working


def dostosuj_zakladki_i_wypelnij(sciezka_pliku, dane_zakladek, dane_ef_kopia, chroniony):
    """
    Otwiera skopiowany plik xlsx i:
      1. Dopasowuje liczbę zakładek roboczych do len(dane_zakladek).
      2. Zmienia nazwy zakładek roboczych zgodnie z dane_zakladek[i]["nazwa"].
      3. Wypełnia w każdej zakładce roboczej:
           C15:C19 ← dane_zakladek[i]["L_dane"]
           D15:D19 ← dane_zakladek[i]["M_dane"]
           E15:E19 ← dane_ef_kopia[i]["E_dane"]
           F15:F19 ← dane_ef_kopia[i]["F_dane"]
      4. Zapisuje plik (jedno wb.save() na końcu).
    Zakładka ARKUSZ_CHRONIONY ("Wyniki") nie jest nigdy modyfikowana.
    """
    wb = openpyxl.load_workbook(sciezka_pliku, keep_links=False)

    wyniki_ok = chroniony in wb.sheetnames
    if not wyniki_ok:
        print(f"  [OSTRZEŻENIE] Brak zakładki '{chroniony}' w pliku.")

    # Pobierz zakładki robocze (wszystkie poza chronioną)
    working = [s for s in wb.sheetnames if s != chroniony]

    if not working:
        print(f"  [BŁĄD] Brak zakładek roboczych — pomijam plik.")
        wb.close()
        return

    if not dane_zakladek:
        print(f"  [OSTRZEŻENIE] Brak definicji zakładek — pomijam modyfikację zakładek.")
        wb.close()
        return

    N = len(dane_zakladek)
    first_ws_name = working[0]

    # --- Etap 2: dopasuj liczbę zakładek roboczych ---
    working = _dopasuj_liczbe_zakładek(wb, working, N, chroniony, first_ws_name)

    # Odśwież listę po zmianach struktury
    working = [s for s in wb.sheetnames if s != chroniony]

    # --- Etap 2: zmień nazwy przez tymczasowe (unika konfliktów między starymi a nowymi) ---
    for i, ws_name in enumerate(working):
        wb[ws_name].title = f"__tmp_{i}__"

    docelowe_nazwy = _unikalne_nazwy_zakladek(dane_zakladek)
    working_tmp = [s for s in wb.sheetnames if s != chroniony]
    for i, ws_name in enumerate(working_tmp):
        wb[ws_name].title = docelowe_nazwy[i]

    # --- Etap 3 & 4: wypełnij komórki w każdej zakładce roboczej ---
    working_final = [s for s in wb.sheetnames if s != chroniony]

    ADRESY_C = ["C15", "C16", "C17", "C18", "C19"]
    ADRESY_D = ["D15", "D16", "D17", "D18", "D19"]
    ADRESY_E = ["E15", "E16", "E17", "E18", "E19"]
    ADRESY_F = ["F15", "F16", "F17", "F18", "F19"]

    for i, ws_name in enumerate(working_final):
        ws = wb[ws_name]
        zd = dane_zakladek[i]

        # C15:C19 ← L_dane (stałe dla wszystkich kopii)
        for offset, addr in enumerate(ADRESY_C):
            val = zd["L_dane"][offset]
            if val is not None:
                ws[addr] = val

        # D15:D19 ← M_dane (stałe dla wszystkich kopii)
        for offset, addr in enumerate(ADRESY_D):
            val = zd["M_dane"][offset]
            if val is not None:
                ws[addr] = val

        # E15:E19 i F15:F19 ← dane zależne od numeru kopii
        if i < len(dane_ef_kopia):
            ef = dane_ef_kopia[i]
            for offset, addr in enumerate(ADRESY_E):
                val = ef["E_dane"][offset]
                if val is not None:
                    ws[addr] = val
            for offset, addr in enumerate(ADRESY_F):
                val = ef["F_dane"][offset]
                if val is not None:
                    ws[addr] = val

    wb.save(sciezka_pliku)


# =============================================================================
# ZAPIS DO KOPII PRZEZ XLWINGS (COM Excel) — zachowuje formuly
# =============================================================================

def _dostosuj_xlwings(app, sciezka_pliku, dane_zakladek, dane_ef_kopia, rekord, nowa_nazwa, chroniony, f24_val):
    """
    Otwiera kopie xlsx przez xlwings (COM Excel):
      1. Dopasowuje liczbe zakladek roboczych do len(dane_zakladek).
      2. Zmienia nazwy zakladek (przez tymczasowe, zeby uniknac konfliktow).
      3. Wypelnia C15:C19 / D15:D19 / E15:E19 / F15:F19 (dane z Strona 3).
      4. Wypelnia komorki naglowkowe i stopkowe w zakladkach roboczych (Etap 5).
      5. Wypelnia F24, C28, C32, E28:G28, E32:G32 w arkuszu Wyniki (Etap 6).
            6. Zapisuje plik.
    Uzycie COM Excel gwarantuje zachowanie wszystkich formul w arkuszu.
        Zwraca: None
    """
    wb = app.books.open(sciezka_pliku)
    try:
        wykluczone = {chroniony, ARKUSZ_WNIOSKI}
        working = [s.name for s in wb.sheets if s.name not in wykluczone]

        if not working:
            print("  [BLAD] Brak zakladek roboczych — pomijam.")
            return
        if not dane_zakladek:
            return

        N = len(dane_zakladek)
        first_ws_name = working[0]

        # --- Dopasuj liczbe zakladek ---
        if len(working) > N:
            for ws_name in working[N:]:
                wb.sheets[ws_name].delete()
            working = working[:N]

        elif len(working) < N:
            for _ in range(N - len(working)):
                before = {s.name for s in wb.sheets}
                src = wb.sheets[first_ws_name]
                if chroniony in {s.name for s in wb.sheets}:
                    src.api.Copy(Before=wb.sheets[chroniony].api)
                else:
                    src.api.Copy(After=wb.sheets[-1].api)
                after = {s.name for s in wb.sheets}
                new_name = (after - before).pop()
                working.append(new_name)

        working = [s.name for s in wb.sheets if s.name not in wykluczone]

        # --- Zmien nazwy przez tymczasowe (unika konfliktow) ---
        for i, ws_name in enumerate(working):
            wb.sheets[ws_name].name = f"__tmp_{i}__"

        docelowe_nazwy = _unikalne_nazwy_zakladek(dane_zakladek)
        working_tmp = [s.name for s in wb.sheets if s.name not in wykluczone]
        for i, ws_name in enumerate(working_tmp):
            wb.sheets[ws_name].name = docelowe_nazwy[i]

        # --- Wypelnij komorki ---
        working_final = [s.name for s in wb.sheets if s.name not in wykluczone]

        prefiks, rok, nr_fab, _ = _parsuj_nazwe_pliku(nowa_nazwa)
        dzis = datetime.datetime.now()

        for i, ws_name in enumerate(working_final):
            ws = wb.sheets[ws_name]
            zd = dane_zakladek[i]

            for offset in range(5):
                val = zd["L_dane"][offset]
                if val is not None:
                    ws.cells(15 + offset, 3).value = val   # C15:C19

            for offset in range(5):
                val = zd["M_dane"][offset]
                if val is not None:
                    ws.cells(15 + offset, 4).value = val   # D15:D19

            if i < len(dane_ef_kopia):
                ef = dane_ef_kopia[i]
                for offset in range(5):
                    val = ef["E_dane"][offset]
                    if val is not None:
                        ws.cells(15 + offset, 5).value = val  # E15:E19
                for offset in range(5):
                    val = ef["F_dane"][offset]
                    if val is not None:
                        ws.cells(15 + offset, 6).value = val  # F15:F19

            # --- Etap 5: naglowki i stopki ---
            ws.range("E4").value = f"{prefiks}/LA/TH/{rok}"
            ws.range("G6").value = nr_fab
            k4 = zd.get("K4_val")
            if k4 is not None:
                ws.range("K4").value = k4
            if rekord.get("B") is not None:
                ws.range("E5").value = rekord["B"]
            if rekord.get("D") is not None:
                ws.range("E6").value = rekord["D"]
            if rekord.get("K") is not None:
                ws.range("H57").value = rekord["K"]
            ws.range("B228").value = dzis
            ws.range("H228").value = dzis
            ws.range("B230").value = PODPISUJACY_1
            ws.range("H230").value = PODPISUJACY_2

        # --- Etap 6: arkusz Wyniki ---
        if chroniony in {s.name for s in wb.sheets}:
            ws_w = wb.sheets[chroniony]
            if f24_val is not None:
                ws_w.range("F24").value = f24_val
            ws_w.range("C28").value = dzis
            ws_w.range("C32").value = dzis
            ws_w.range("E28").value = PODPISUJACY_1   # scalone E28:G28
            ws_w.range("E32").value = PODPISUJACY_2   # scalone E32:G32

        app.api.CalculateFullRebuild()  # przelicz formuly (w tym odwolania do plikow linkowanych) przed zapisem
        wb.save()  # zapisuje plik
    finally:
        wb.close()


def _odczytaj_kalibracje_xlwings(app, sciezka_pliku, chroniony, enable_diag=False):
    """Czyta dane kalibracyjne z zakladek roboczych po pelnym przeliczeniu formul."""
    wb = app.books.open(sciezka_pliku)
    kalibracja = []
    try:
        # Wymus aktualizacje linkow zewnetrznych (bez tego Excel moze zostawiac
        # stare cache'owane wartosci, gdy DisplayAlerts=False suprimuje dialog).
        try:
            sources = wb.api.LinkSources(1)  # 1 = xlLinkTypeExcelLinks
            if sources:
                for src in sources:
                    wb.api.UpdateLink(Name=src, Type=1)
        except Exception as exc:
            print(f"    [UWAGA] UpdateLink: {exc}")

        wykluczone = {chroniony, ARKUSZ_WNIOSKI}

        app.api.CalculateFullRebuild()

        def _brak_wartosci(v):
            return v is None or (isinstance(v, str) and v.strip() == "")

        working_final_list = [s.name for s in wb.sheets if s.name not in wykluczone]
        _diag_done = False
        for ws_name in working_final_list:
            ws_cal = wb.sheets[ws_name]

            d = ws_cal.range("D246").value
            e = ws_cal.range("E246").value
            f = ws_cal.range("F246").value
            g = ws_cal.range("G246").value

            # Diagnostyka posrednich komorek (tylko pierwsza kopia, gdy enable_diag=True).
            if enable_diag and not _diag_done and (d is None or f is None or g is None):
                print(f"      [DIAG D246] formula={ws_cal.range('D246').formula!r}  value={ws_cal.range('D246').value!r}")
                print(f"      [DIAG F246] formula={ws_cal.range('F246').formula!r}  value={ws_cal.range('F246').value!r}")
                print(f"      [DIAG G246] formula={ws_cal.range('G246').formula!r}  value={ws_cal.range('G246').value!r}")
                for cell_addr in ('H57', 'G137', 'E220', 'G220', 'J220'):
                    rng = ws_cal.range(cell_addr)
                    print(f"      [DIAG {cell_addr}] formula={rng.formula!r}  value={rng.value!r}")
                _diag_done = True

            # Fallback: szukaj 4 niepustych wartosci w szerszym pasie,
            # bo przy scaleniach dane moga wpasc nie w D:E:F:G.
            if _brak_wartosci(d) or _brak_wartosci(f) or _brak_wartosci(g):
                pas = ws_cal.range("C246:J246").value
                if isinstance(pas, list):
                    kandydaci = [v for v in pas if not _brak_wartosci(v)]
                    if len(kandydaci) >= 4:
                        d, e, f, g = kandydaci[:4]

            print(f"    [Kalibracja {ws_name}] REF={d!r}  ZM={e!r}  POP={f!r}  NIEP={g!r}")
            kalibracja.append({
                "wartosc_odn": d,
                "zmierzona":   e,
                "poprawka":    f,
                "niepewnosc":  g,
            })
    finally:
        wb.close()
    return kalibracja


# =============================================================================
# TWORZENIE KOPII
# =============================================================================

def utworz_kopie(folder, szablon_plik, dane, dane_zakladek, dane_ef, chroniony, f24_per_kopia):
    """
    Dla każdego rekordu z Strona 2:
      1. Generuje nazwę kopii i kopiuje szablon (shutil.copy2).
      2. Modyfikuje kopie przez xlwings (jeden proces Excel dla wszystkich).
    Zwraca listę nazw utworzonych kopii.
    """
    sciezka_szablonu = os.path.join(folder, szablon_plik)
    n = len(dane)
    kopie = []

    # Krok 1: utwórz wszystkie kopie (szybkie kopiowanie pliku)
    for j, rekord in enumerate(dane):
        nowa_nazwa = generuj_nazwe_pliku(szablon_plik, rekord["O"], rekord["E"])
        sciezka_kopii = os.path.join(folder, nowa_nazwa)
        if os.path.exists(sciezka_kopii):
            print(f"  [UWAGA] Plik już istnieje, nadpisuję: {nowa_nazwa}")
        shutil.copy2(sciezka_szablonu, sciezka_kopii)
        kopie.append((j, nowa_nazwa, sciezka_kopii, rekord))

    print(f"  Skopiowano {n} plików. Modyfikuję zakładki i komórki przez Excel COM...")

    # Krok 2: modyfikuj wszystkie kopie w jednej sesji Excel (COM)
    app = xw.App(visible=False, add_book=False)
    app.api.AutomationSecurity = 1   # msoAutomationSecurityLow — wlacza makra bez pytania
    app.api.DisplayAlerts = False
    app.api.AskToUpdateLinks = False  # aktualizuj linki zewnetrzne bez pytania
    try:
        # Otwieramy pliki linkowane PRZED modyfikacja kopii.
        # Dzieki temu formuly odwolujace sie do tych plikow oblicza sie
        # przy kazdym CalculateFullRebuild+wb.save() i sa zakeszowane w kopii.
        linked_wbs = []
        for plik_link in PLIKI_LINKOWANE:
            sciezka_link = os.path.join(folder, plik_link)
            if os.path.exists(sciezka_link):
                try:
                    lwb = app.books.open(sciezka_link)
                    linked_wbs.append(lwb)
                    print(f"  Otwarto plik linkowany: {plik_link}")
                except Exception as exc:
                    print(f"  [UWAGA] Nie mozna otworzyc pliku linkowanego: {plik_link} — {exc}")
            else:
                print(f"  [UWAGA] Brak pliku linkowanego w folderze: {plik_link}")

        for j, nowa_nazwa, sciezka_kopii, rekord in kopie:
            print(f"[{j+1:>{len(str(n))}}/{n}] {nowa_nazwa}")
            dane_ef_kopia = dane_ef[j] if j < len(dane_ef) else []
            f24_val = f24_per_kopia[j] if j < len(f24_per_kopia) else None
            _dostosuj_xlwings(app, sciezka_kopii, dane_zakladek, dane_ef_kopia, rekord, nowa_nazwa, chroniony, f24_val)

        print("  Odczyt kalibracji po zakonczeniu wypelniania wszystkich kopii...")
        dane_kalibracji = []
        for j, nowa_nazwa, sciezka_kopii, _ in kopie:
            print(f"    [Kalibracja {j+1:>{len(str(n))}}/{n}] {nowa_nazwa}")
            kal = _odczytaj_kalibracje_xlwings(app, sciezka_kopii, chroniony, enable_diag=(j == 0))
            dane_kalibracji.append(kal or [])

        for lwb in linked_wbs:
            try:
                lwb.close()
            except Exception:
                pass
    finally:
        app.quit()

    nazwy = [nazwa for (_, nazwa, _, _) in kopie]
    return nazwy, dane_kalibracji


# =============================================================================
# GŁÓWNY PUNKT WEJŚCIA
# =============================================================================

def main():
    protokol = os.path.join(FOLDER, PROTOKOL_PLIK)
    szablon  = os.path.join(FOLDER, SZABLON_PLIK)

    SEP = "=" * 65

    # --- Walidacja plików wejściowych ---
    if not os.path.exists(protokol):
        print(f"[BŁĄD] Plik protokołu nie istnieje:\n  {protokol}")
        return
    if not os.path.exists(szablon):
        print(f"[BŁĄD] Plik szablonu nie istnieje:\n  {szablon}")
        return

    # --- Etapy 1-4: odczyt wszystkich danych przez xlwings (Excel COM) ---
    # xlwings uruchamia Excel i odczytuje prawdziwe wyliczone wartości formuł,
    # co rozwiązuje problem None przy data_only=True w openpyxl.
    print(SEP)
    print("ETAP 1-4  Odczyt danych z protokołu (przez Excel COM / xlwings)")
    print(SEP)
    print("  Uruchamiam Excel w tle — proszę czekać…")
    try:
        dane_s2, dane_zakladek, dane_ef, f24_per_kopia = wczytaj_wszystko_xlwings(
            protokol,
            ark_s2=ARKUSZ_STRONA2,
            ark_s3=ARKUSZ_STRONA3,
            start_s2=START_ROW_S2,
            start_s3=START_ROW_S3,
            blok=BLOK_S3,
            start_col_e=START_COL_E_S3,
            start_col_f=START_COL_F_S3,
            krok=KROK_COL_EF,
        )
    except Exception as e:
        print(f"[BŁĄD] Odczyt przez xlwings nie powiódł się: {e}")
        return

    print(f"  Znaleziono kopii: {len(dane_s2)}")
    if not dane_s2:
        print("[BŁĄD] Brak danych w Strona 2 — skrypt kończy pracę.")
        return

    if not dane_zakladek:
        print("[OSTRZEŻENIE] Brak definicji zakładek w Strona 3.")
    else:
        print(f"  Zakładki ({len(dane_zakladek)}):")
        for zd in dane_zakladek:
            print(f"    • {zd['nazwa']}")

    print(f"  Wczytano dane E/F: {len(dane_ef)} kopii × {len(dane_zakladek)} zakładek.")
    print(f"  Wczytano dane F24 (Wyniki): {len(f24_per_kopia)} wartości.")

    # --- Tworzenie i przetwarzanie kopii ---
    print(SEP)
    print("ETAP 1+2+3+4  Tworzenie kopii i dostosowanie zakładek")
    print(SEP)

    nazwy, dane_kalibracji = utworz_kopie(
        FOLDER, SZABLON_PLIK,
        dane_s2, dane_zakladek, dane_ef,
        ARKUSZ_CHRONIONY,
        f24_per_kopia,
    )

    print(SEP)
    print(
        f"Zakończono etapy Excel. Utworzono {len(nazwy)} kopii, "
        f"każda z {len(dane_zakladek)} zakładkami + {ARKUSZ_CHRONIONY}."
    )
    print(SEP)

    # --- Etap 7: dokumenty Word (świadectwa wzorcowania) ---
    if SZABLON_WORD:
        szablon_word_path = os.path.join(FOLDER, SZABLON_WORD)
        if not os.path.exists(szablon_word_path):
            print(f"[UWAGA] Plik szablonu Word nie istnieje: {SZABLON_WORD} — pomijam Etap 7")
        else:
            print(SEP)
            print("ETAP 7  Tworzenie świadectw wzorcowania (dokumenty Word)")
            print(SEP)
            nazwy_word = utworz_kopie_word(
                FOLDER, SZABLON_WORD, dane_s2, nazwy,
                dane_zakladek, dane_kalibracji, NR_SW_POCZATKOWY,
            )
            print(SEP)
            print(f"Zakończono Etap 7. Utworzono {len(nazwy_word)} dokumentów Word.")
            print(SEP)


if __name__ == "__main__":
    main()
