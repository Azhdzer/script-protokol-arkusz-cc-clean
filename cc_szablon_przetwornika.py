# -*- coding: utf-8 -*-
"""
cc_szablon_przetwornika.py — szablon arkusza obliczeniowego dla PRZETWORNIKA.

Kazda zakladka robocza konczy sie blokiem 'Wyniki do umieszczenia na swiadectwie
wzorcowania' — osobna tabela dla kazdego rodzaju przyrzadu:

    239  TERMOHIGROMETR                                   (D..K)
    246  TERMOMETR ELEKTRONICZNY                          (D..G)
    253  PRZETWORNIK TEMPERATURY I WILGOTNOSCI WZGLEDNEJ  (D..O)
    260  PRZETWORNIK TEMPERATURY                          (D..I)
    267  PRZETWORNIK WILGOTNOSCI WZGLEDNEJ                (D..J)

'Wyniki' standardowego szablonu zbieraja dwie z nich: termohigrometr (239, B:I)
i 'tylko temperatura' (246, O:R). Szablon przetwornika robi to samo dla
przetwornikow — wyglad (ramki, scalenia, szerokosci, formaty liczb) skopiowany
z blokow zakladki:

    B:M   ← 253  przetwornik t + RH   (tylko RH: kolumny temperatury '-')
    O:T   ← 260  przetwornik t        (dla t + RH: punkty tylko temperaturowe)

Krok 3 zostawia tabele pasujace do przetwornika (t + RH: 253 i 260; RH: 253
z '-' w kolumnach t — jak w recznie przygotowanych wynikach; t: 260). Naglowki przepisuje z zakladki, w ktorej L10 jest
juz ustawione — wiec 'pradu/napiecia', I/U i mA/V sa wlasciwe.

Modul robi szablon przetwornika z szablonu standardowego — przez Excela (wlasna
instancja), wiec linki zewnetrzne, formaty i walidacje zostaja. Uruchom ponownie
po nowej wersji szablonu standardowego:

    python cc_szablon_przetwornika.py [szablon_zrodlowy.xlsx] [szablon_docelowy.xlsx]
"""

import os
import re
import shutil
import sys

import cc_excel

DOMYSLNE_ZRODLO = ("xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 "
                   "z 17.06.2026 - 1 - RH (CC).xlsx")
DOMYSLNY_CEL = ("xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 "
                "z 17.06.2026 - 1 - Przetwornik (CC).xlsx")


def _litera(nr):
    s = ""
    while nr:
        nr, r = divmod(nr - 1, 26)
        s = chr(65 + r) + s
    return s


def _numer(litera):
    n = 0
    for ch in litera:
        n = n * 26 + ord(ch) - 64
    return n


def _kolumny(od, do):
    return [_litera(i) for i in range(_numer(od), _numer(do) + 1)]


# nazwa, kolumny 'Wynikow', wiersz wynikow zakladki, kolumny zakladki, naglowek zakladki,
# rodzaj ('rh' — tylko punkty z wilgotnoscia; 't' — reprezentant na temperature)
TABELE = (
    {"nazwa": "th", "wyniki": _kolumny("B", "M"), "wiersz": 253, "zrodlo": _kolumny("D", "O"),
     "naglowek": (250, 252), "rodzaj": "rh"},
    {"nazwa": "t", "wyniki": _kolumny("O", "T"), "wiersz": 260, "zrodlo": _kolumny("D", "I"),
     "naglowek": (257, 259), "rodzaj": "t"},
)
# Histereza: RHm (I253); przetwornik t — tm (F260).
HISTEREZA = {"th": "I253", "t": "F260"}
# Przetwornik tylko RH: tabela 253 z '-' w kolumnach temperatury (Itm, tm, Δt, Ut) —
# jak w recznie przygotowanych 'Wynikach' (230, 263051).
KOLUMNY_T_TABELI_TH = ("D", "F", "H")
KOLUMNY_UT_TABELI_TH = ("J", "K")

_RE_ARKUSZ = re.compile(r"^=(?:'((?:[^']|'')+)'|([^!'][^!]*))!")


def _arkusz(formula):
    m = _RE_ARKUSZ.match(str(formula or ""))
    if not m:
        return None
    return (m.group(1) or m.group(2)).replace("''", "'")


def _ref(arkusz, komorka):
    return f"='{arkusz.replace(chr(39), chr(39) * 2)}'!{komorka}"


def formuly_wiersza(arkusz, tabela):
    """{kolumna Wynikow: formula} jednej zakladki dla tabeli z TABELE."""
    return {kw: _ref(arkusz, f"{kz}{tabela['wiersz']}")
            for kw, kz in zip(tabela["wyniki"], tabela["zrodlo"])}


def tabela(nazwa):
    return next(t for t in TABELE if t["nazwa"] == nazwa)


def zakres(tab, od_w, do_w):
    return f"{tab['wyniki'][0]}{od_w}:{tab['wyniki'][-1]}{do_w}"


def przepisz_naglowki(ws_wyniki, ws_zakladki, tabele):
    """Naglowki (wiersze 1-3) tabel 'Wynikow' = wartosci naglowka bloku zakladki."""
    for tab in tabele:
        h0, h1 = tab["naglowek"]
        zr = ws_zakladki.Range(f"{tab['zrodlo'][0]}{h0}:{tab['zrodlo'][-1]}{h1}")
        ws_wyniki.Range(zakres(tab, 1, 3)).Value = zr.Value


def zbuduj(zrodlo, cel, app=None):
    """Tworzy szablon przetwornika `cel` z szablonu standardowego `zrodlo`."""
    shutil.copy2(zrodlo, cel)
    try:
        os.chmod(cel, 0o666)
    except OSError:
        pass
    wlasny = app is None
    if wlasny:
        import xlwings as xw
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
    try:
        wb = app.books.api.Open(os.path.abspath(cel), UpdateLinks=0)
        ws = wb.Worksheets("Wyniki")
        arkusze = []
        for r in range(4, 10):
            a = _arkusz(ws.Range(f"B{r}").Formula)
            if a:
                arkusze.append((r, a))
        if not arkusze:
            raise ValueError("W 'Wynikach' szablonu nie ma formul B4:B9 z nazwami zakladek.")
        z0 = wb.Worksheets(arkusze[0][1])

        # Stare tabele (termohigrometr B:I, tylko temperatura O:R) — precz, z formatami.
        ws.Range("B1:AB9").UnMerge()
        ws.Range("B1:AB9").Clear()

        # Naglowki z zakladki z L10 = 'Prąd' (w szablonie L10 = '-', formuly daja '-').
        l10 = z0.Range("L10").Value
        z0.Range("L10").Value = "Prąd"
        z0.Calculate()
        for tab in TABELE:
            h0, h1 = tab["naglowek"]
            s0, s1 = tab["zrodlo"][0], tab["zrodlo"][-1]
            z0.Range(f"{s0}{h0}:{s1}{h1}").Copy()
            cel_n = ws.Range(zakres(tab, 1, 3))
            cel_n.PasteSpecial(-4122)            # formaty (ramki, scalenia)
            cel_n.PasteSpecial(8)                # szerokosci kolumn
            z0.Range(f"{s0}{tab['wiersz']}:{s1}{tab['wiersz']}").Copy()
            ws.Range(zakres(tab, 4, 9)).PasteSpecial(-4122)
            app.api.CutCopyMode = False
        przepisz_naglowki(ws, z0, TABELE)
        z0.Range("L10").Value = l10

        for r, a in arkusze:
            for tab in TABELE:
                for kol, f in formuly_wiersza(a, tab).items():
                    ws.Range(f"{kol}{r}").Formula = f

        # Histereza: RHm punktu i jego powtorki (t + RH).
        nazwy = [a for _r, a in arkusze]
        baza = next((a for a in nazwy if a + " (2)" in nazwy), None)
        if baza:
            k = HISTEREZA["th"]
            ws.Range("J23").Formula = (f"=ABS(VALUE({_ref(baza, k)[1:]})"
                                       f"-VALUE({_ref(baza + ' (2)', k)[1:]}))")
        wb.Save()
        wb.Close(False)
    finally:
        if wlasny:
            cc_excel.zamknij_excel_xlwings(app)
    return cel


# ─── Szablon przetwornika TYLKO RH — 'Wyniki' 1:1 z wzorcowego pliku ─────────
# Operator przygotowal 'Wyniki' dla przetwornika wilgotnosci recznie (zlecenie 230,
# 263051): tabela B:L z wiersza 253 zakladek, kolumny temperatury '-', URH w mA i %.
# Szablon RH dostaje DOKLADNIE ten wyglad (formaty, scalenia, szerokosci, wysokosci,
# formuly) — krok 3 zmienia tylko formuly wg punktow.
DOMYSLNY_CEL_RH = ("xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 "
                   "z 17.06.2026 - 1 - Przetwornik RH (CC).xlsx")
ZAKRES_WZORU_RH = "A1:S12"
# kolumna 'Wynikow' -> komorka wiersza 253 ('-' = stala kreska)
KOLUMNY_RH = (("B", "D253"), ("C", "E253"), ("D", "-"), ("E", "G253"), ("F", "-"),
              ("G", "I253"), ("H", "-"), ("I", "K253"), ("J", "-"), ("K", "N253"),
              ("L", "O253"))


def zbuduj_rh(zrodlo, wzor_wynikow, cel, app=None):
    """
    Szablon przetwornika RH: szablon standardowy `zrodlo` + blok 'Wynikow'
    A1:S12 przeniesiony 1:1 z pliku `wzor_wynikow` (wyglad i formuly).
    """
    shutil.copy2(zrodlo, cel)
    try:
        os.chmod(cel, 0o666)
    except OSError:
        pass
    wlasny = app is None
    if wlasny:
        import xlwings as xw
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
    try:
        wb = app.books.api.Open(os.path.abspath(cel), UpdateLinks=0)
        wz = app.books.api.Open(os.path.abspath(wzor_wynikow), UpdateLinks=0, ReadOnly=True)
        ws, ww = wb.Worksheets("Wyniki"), wz.Worksheets("Wyniki")
        ws.Range(ZAKRES_WZORU_RH).UnMerge()
        ws.Range(ZAKRES_WZORU_RH).Clear()
        ww.Range(ZAKRES_WZORU_RH).Copy()
        ws.Range("A1").PasteSpecial(-4122)          # formaty, ramki, scalenia
        ws.Range("A1").PasteSpecial(8)              # szerokosci kolumn
        app.api.CutCopyMode = False
        for r in range(1, 13):
            ws.Rows(r).RowHeight = ww.Rows(r).RowHeight
            for c in range(1, 20):
                f = ww.Cells(r, c).Formula
                if f not in (None, ""):
                    # wzor ma w E3 odwolanie do L1 zamiast $L$10 — poprawiamy
                    ws.Cells(r, c).Formula = str(f).replace("'!L1=", "'!$L$10=")
        nazwy = [_arkusz(ws.Range(f"B{r}").Formula) for r in range(4, 10)]
        nazwy = [n for n in nazwy if n]
        baza = next((a for a in nazwy if a + " (2)" in nazwy), None)
        if baza:
            ws.Range("J23").Formula = (f"=ABS(VALUE({_ref(baza, 'I253')[1:]})"
                                       f"-VALUE({_ref(baza + ' (2)', 'I253')[1:]}))")
        wz.Close(False)
        wb.Save()
        wb.Close(False)
    finally:
        if wlasny:
            cc_excel.zamknij_excel_xlwings(app)
    return cel


# ─── Szablon WYSWIETLACZ + WYJSCIE ('TH + Przetwornik') — z recznie przygotowanego ────
# Zlecenie 239, VAISALA PTU303 D4850103: przyrzad wzorcowany z wyswietlacza i z wyjscia
# napieciowego naraz. Operator przygotowal arkusz recznie: zakladki jak w szablonie
# standardowym (K10 'TH/Przetwornik', E/F z wyswietlacza, G/H z wyjscia), 'Wyniki' z
# trzema sekcjami (239/246, 253/260, 274/279) i kryteriami zgodnosci. Szablon to TEN
# plik: 'Wyniki' zostaja 1:1, zakladki wracaja do wartosci szablonu standardowego
# (bez pomiarow, dat, podpisow), zostaje tyle zakladek, ile ma szablon standardowy.
DOMYSLNY_CEL_TH = ("xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 "
                   "z 17.06.2026 - 1 - TH + Przetwornik (CC).xlsx")
ZAKRES_ZAKLADKI = "A1:BT300"
ZAKRES_WYNIKOW_TH = "A1:AJ400"
_RE_NAZWA_ARKUSZA = re.compile(r"'((?:[^']|'')+)'!")


def przemapuj_arkusze(formula, mapa):
    """W formule odwolania 'Arkusz'!... do arkuszy z `mapa` -> nowa nazwa."""
    def _zm(m):
        nazwa = m.group(1).replace("''", "'")
        nowa = mapa.get(nazwa, nazwa)
        return "'" + nowa.replace("'", "''") + "'!"
    return _RE_NAZWA_ARKUSZA.sub(_zm, formula)


def mapa_usuwanych_arkuszy(formuly, zostaja):
    """
    Arkusze spoza `zostaja` (w kolejnosci pojawienia sie w formulach) -> kolejno
    zostajace arkusze, w kolko. Formuly 'Wynikow' nie dostaja #ADR! po usunieciu zakladek.
    """
    kolejnosc = []
    for f in formuly:
        for m in _RE_NAZWA_ARKUSZA.finditer(str(f or "")):
            n = m.group(1).replace("''", "'")
            if n not in kolejnosc:
                kolejnosc.append(n)
    obce = [n for n in kolejnosc if n not in zostaja]
    return {n: zostaja[i % len(zostaja)] for i, n in enumerate(obce)}


def roznice_do_wyczyszczenia(wzor, szablon):
    """
    Komorki zakladki wzoru recznego do przywrocenia z szablonu standardowego:
    [(wiersz, kolumna, wartosc_szablonu)] (od 1). Stale (pomiary, daty, numery,
    podpisy) wracaja do szablonu; formuly zmienione w recznym wzorze (np. odwolanie
    do przesunietej stopki 'Wynikow': M50, H56) zostaja.
    """
    wynik = []
    for r, (w_w, w_s) in enumerate(zip(wzor, szablon), start=1):
        for c, (a, b) in enumerate(zip(w_w, w_s), start=1):
            a, b = a or "", b or ""
            if a == b:
                continue
            if str(a).startswith("=") and str(b).startswith("="):
                continue
            wynik.append((r, c, b))
    return wynik


def zbuduj_th(wzor_reczny, zrodlo, cel, app=None):
    """Szablon 'TH + Przetwornik' z recznie przygotowanego arkusza `wzor_reczny`."""
    shutil.copy2(wzor_reczny, cel)
    try:
        os.chmod(cel, 0o666)
    except OSError:
        pass
    wlasny = app is None
    if wlasny:
        import xlwings as xw
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
    try:
        app.api.EnableEvents = False
        st = app.books.api.Open(os.path.abspath(zrodlo), UpdateLinks=0, ReadOnly=True)
        app.api.Calculation = -4135                  # reczne — bez przeliczania 15 zakladek
        wb = app.books.api.Open(os.path.abspath(cel), UpdateLinks=0)
        zakl_s = [s.Name for s in st.Worksheets if s.Name != "Wyniki"]
        zakl_w = [s.Name for s in wb.Worksheets if s.Name != "Wyniki"]
        zostaja = zakl_w[:len(zakl_s)]
        ws = wb.Worksheets("Wyniki")

        # 1. 'Wyniki': odwolania do usuwanych zakladek -> zakladki, ktore zostaja.
        formuly = ws.Range(ZAKRES_WYNIKOW_TH).Formula
        plaskie = [f for w in formuly for f in w]
        mapa = mapa_usuwanych_arkuszy(plaskie, zostaja)
        for r, wiersz in enumerate(formuly, start=1):
            for c, f in enumerate(wiersz, start=1):
                if isinstance(f, str) and f.startswith("="):
                    nowa = przemapuj_arkusze(f, mapa)
                    if nowa != f:
                        ws.Cells(r, c).Formula = nowa
        for n in zakl_w[len(zostaja):]:
            wb.Worksheets(n).Delete()

        # 2. Zakladki: stale wracaja do szablonu standardowego.
        wzorzec = st.Worksheets(zakl_s[0]).Range(ZAKRES_ZAKLADKI).Formula
        for n in zostaja:
            z = wb.Worksheets(n)
            for r, c, v in roznice_do_wyczyszczenia(z.Range(ZAKRES_ZAKLADKI).Formula, wzorzec):
                kom = z.Cells(r, c)
                if kom.HasArray:
                    continue
                if v == "":
                    (kom.MergeArea if kom.MergeCells else kom).ClearContents()
                else:
                    kom.Formula = v

        # 3. Nazwy zakladek jak w szablonie standardowym.
        for i, n in enumerate(zostaja):
            wb.Worksheets(n).Name = f"__tmp_{i}__"
        for i, n in enumerate(zostaja):
            wb.Worksheets(f"__tmp_{i}__").Name = zakl_s[i]

        # 4. Stopka 'Wynikow': odleglosc jak w szablonie, bez dat i podpisow.
        kol_a = ws.Range("A1:A400").Formula
        stopka = next((i for i, (v,) in enumerate(kol_a, start=1)
                       if str(v or "").strip().upper().startswith("ODLEG")), None)
        if stopka:
            ws.Range(f"F{stopka}").Value = st.Worksheets("Wyniki").Range("F24").Value
            for dr in (4, 8):
                for kol in ("C", "E"):
                    ws.Range(f"{kol}{stopka + dr}").MergeArea.ClearContents()
        ws.Activate()
        # Zapis w trybie automatycznym (jak szablon standardowy), ale bez przeliczania
        # pustych zakladek — przeliczenie ich z linkami potrafi trwac kilkanascie minut
        # (krok 3, zlecenie 239). EnableCalculation nie jest zapisywane w pliku.
        for n in zakl_s[:len(zostaja)]:
            wb.Worksheets(n).EnableCalculation = False
        app.api.Calculation = -4105
        wb.Save()
        wb.Close(False)
        st.Close(False)
    finally:
        try:
            app.api.EnableEvents = True
        except Exception:
            pass
        if wlasny:
            cc_excel.zamknij_excel_xlwings(app)
    return cel


# ─── Swiadectwo WYSWIETLACZ + WYJSCIE — szablon z wydanego swiadectwa ───────────────
# Wzor: 1276/239/LA/TH/2026 (ang.). Dane klienta i przyrzadu -> placeholdery, tabele
# wynikow puste (krok 3 wypelnia je z sekcji 'Wynikow'), 'voltage'/'V' -> placeholdery,
# zeby ten sam szablon pasowal do wyjscia pradowego.
DOMYSLNY_CEL_WORD_TH = "xxx_yyy_LA_TH_2026 - TH + przetwornik (ang).docx"
_RE_OBIEKT = re.compile(r"type:\s*(?P<typ>.+?),\s*serial number:\s*(?P<nr_fabr>.+?),"
                        r"(?:\s*identification number:\s*(?P<nr_ewid>.+?),)?"
                        r"\s*manufacturer:\s*(?P<wytworca>.+?),"
                        r"\s*year of production:\s*(?P<rok_produkcji>\d{4})")
_RE_ZAKRES_DWOCH = re.compile(r"\((-?[\d,.]+)\s+(?:÷\s*)?(-?[\d,.]+)\)")


def zamien_w_akapicie(para, stary, nowy):
    """Pierwsze wystapienie `stary` w akapicie (takze rozbite miedzy runy) -> `nowy`."""
    runy = para.runs
    teksty = [r.text for r in runy]
    pelny = "".join(teksty)
    pos = pelny.find(stary)
    if pos < 0 or not stary:
        return False
    run_znaku = [i for i, t in enumerate(teksty) for _ in t]
    i, j = run_znaku[pos], run_znaku[pos + len(stary) - 1]
    poczatek = sum(len(t) for t in teksty[:i])
    laczny = "".join(teksty[i:j + 1])
    runy[i].text = laczny[:pos - poczatek] + nowy + laczny[pos - poczatek + len(stary):]
    for r in runy[i + 1:j + 1]:
        r.text = ""
    return True


def _akapity(doc):
    """Wszystkie akapity: tresc, tabele, naglowki i stopki."""
    def z(kontener):
        yield from kontener.paragraphs
        for t in kontener.tables:
            for row in t.rows:
                for c in row.cells:
                    yield from z(c)
    yield from z(doc)
    for s in doc.sections:
        yield from z(s.header)
        yield from z(s.footer)


def _tylko_tekst(cell, tekst):
    """Komorka z jednym akapitem i jednym runem `tekst` (format pierwszego runu)."""
    pierwszy = cell.paragraphs[0]
    for p in cell.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    for r in pierwszy.runs[1:]:
        r._r.getparent().remove(r._r)
    if pierwszy.runs:
        pierwszy.runs[0].text = tekst
    else:
        pierwszy.add_run(tekst)


def zbuduj_word_th(wzor_docx, cel, podpis=None):
    """Szablon swiadectwa wyswietlacz + wyjscie z wydanego swiadectwa `wzor_docx`."""
    from docx import Document
    import generuj_arkusze as A
    doc = Document(wzor_docx)
    podpis = podpis or A.PODPISUJACY_2

    # 1. Tabele wynikow — puste wiersze (format zostaje).
    for t in A._tabele_wynikow_swiadectwa(doc):
        for row in list(t.rows)[A.WIERSZE_NAGLOWKA_SWIADECTWA:]:
            for c in A._komorki_wiersza(row):
                A._ustaw_tekst_komorki(c, "")

    # 2. Tabela opisowa (pierwsza): przedmiot, zleceniodawca, warunki, data.
    for row in doc.tables[0].rows:
        etykieta = row.cells[0].text.strip().upper()
        tresc = row.cells[-1]
        if etykieta.startswith("OBJECT"):
            m = _RE_OBIEKT.search(tresc.text)
            if not m:
                raise ValueError(f"Nie rozpoznano opisu przyrzadu: {tresc.text!r}")
            for pole, wartosc in m.groupdict().items():
                if wartosc:
                    for p in tresc.paragraphs:
                        zamien_w_akapicie(p, wartosc.strip(), f"[{pole}]")
        elif etykieta.startswith("APPLICANT"):
            _tylko_tekst(tresc, "[zleceniodawca]")
        elif etykieta.startswith("ENVIRONMENTAL"):
            for p, (a, b) in zip(tresc.paragraphs, (("[temp_min]", "[temp_max]"),
                                                    ("[wilg_min]", "[wilg_max]"))):
                m = _RE_ZAKRES_DWOCH.search(p.text)
                if m:
                    zamien_w_akapicie(p, m.group(1), a)
                    zamien_w_akapicie(p, m.group(2), b)
        elif etykieta.startswith("DATE OF CALIBRATION"):
            _tylko_tekst(tresc, "[data_wzorcowania]")

    # 3. Naglowek: data wydania, numer swiadectwa.
    for s in doc.sections:
        for t in s.header.tables:
            for row in t.rows:
                for c in row.cells:
                    txt = c.text.strip()
                    m = re.fullmatch(r"\d+/\d+/LA/TH/(\d{4})", txt)
                    if m:
                        _tylko_tekst(c, f"[nr_sw]/[nr_zl]/LA/TH/{m.group(1)}")
                    elif re.fullmatch(r"\d{1,2} [A-Z][a-z]+ \d{4}", txt):
                        _tylko_tekst(c, "[data]")

    # 4. Tresc: podpis, przypis histerezy, zakresy, rodzaj wyjscia.
    for p in list(_akapity(doc)):
        if podpis and podpis in p.text:
            zamien_w_akapicie(p, podpis, "[Podpis]")
        m = re.search(r"Point \((-?[\d.,]+) °C, ([\d.,]+) %rh\)", p.text)
        if m:
            zamien_w_akapicie(p, m.group(0), "Point ([hist_t] °C, [hist_rh] %rh)")
        m = re.search(r"tm\s*=\s*\((-?[\d,.]+) ÷ (-?[\d,.]+)\) °C", p.text)
        if m:
            zamien_w_akapicie(p, f"({m.group(1)} ÷ {m.group(2)}) °C", "([t_min] ÷ [t_max]) °C")
        m = re.search(r"RHm\s*=\s*\((-?[\d,.]+) ÷ (-?[\d,.]+)\) %", p.text)
        if m:
            zamien_w_akapicie(p, f"({m.group(1)} ÷ {m.group(2)}) %", "([rh_min] ÷ [rh_max]) %")
        m = re.search(r"\((-?[\d,.]+) ÷ (-?[\d,.]+)\) (V|mA) DC", p.text)
        if m:
            zamien_w_akapicie(p, m.group(0), "([wy_min] ÷ [wy_max]) [jedn_wy] DC")
        for r in p.runs:
            if "voltage" in r.text:
                r.text = r.text.replace("voltage", "[wyjscie_en]")
        runy = p.runs
        for r, nast in zip(runy, runy[1:]):
            if r.text == "V" and nast.font.subscript:
                r.text = "[symbol_en]"
        if p.text.strip() == "V" and len([r for r in runy if r.text.strip()]) == 1:
            for r in runy:
                if r.text.strip() == "V":
                    r.text = r.text.replace("V", "[jedn_wy]")
    doc.save(cel)
    return cel


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--word":
        # python cc_szablon_przetwornika.py --word <wydane swiadectwo.docx> [cel]
        ce = sys.argv[3] if len(sys.argv) > 3 else DOMYSLNY_CEL_WORD_TH
        print(f"Szablon swiadectwa wyswietlacz + wyjscie: {zbuduj_word_th(sys.argv[2], ce)}")
    elif len(sys.argv) > 1 and sys.argv[1] == "--th":
        # python cc_szablon_przetwornika.py --th <recznie przygotowany arkusz> [cel]
        ce = sys.argv[3] if len(sys.argv) > 3 else DOMYSLNY_CEL_TH
        print(f"Szablon wyswietlacz + wyjscie: {zbuduj_th(sys.argv[2], DOMYSLNE_ZRODLO, ce)}")
    elif len(sys.argv) > 1 and sys.argv[1] == "--rh":
        # python cc_szablon_przetwornika.py --rh <plik z wzorcowymi 'Wynikami'> [cel]
        ce = sys.argv[3] if len(sys.argv) > 3 else DOMYSLNY_CEL_RH
        print(f"Szablon przetwornika RH: {zbuduj_rh(DOMYSLNE_ZRODLO, sys.argv[2], ce)}")
    else:
        zr = sys.argv[1] if len(sys.argv) > 1 else DOMYSLNE_ZRODLO
        ce = sys.argv[2] if len(sys.argv) > 2 else DOMYSLNY_CEL
        print(f"Szablon przetwornika: {zbuduj(zr, ce)}")
