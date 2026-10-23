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
przetwornikow — trzy tabele obok siebie, wyglad (ramki, scalenia, szerokosci,
formaty liczb) skopiowany z blokow zakladki:

    B:M   ← 253  przetwornik t + RH
    O:T   ← 260  przetwornik t           (dla t + RH: punkty tylko temperaturowe)
    V:AB  ← 267  przetwornik RH

Krok 3 zostawia tabele pasujace do przetwornika (t + RH: 253 i 260; t: 260;
RH: 267), a pozostale czysci. Naglowki przepisuje z zakladki, w ktorej L10 jest
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
    {"nazwa": "rh", "wyniki": _kolumny("V", "AB"), "wiersz": 267, "zrodlo": _kolumny("D", "J"),
     "naglowek": (264, 266), "rodzaj": "rh"},
)
# Histereza: RHm (t+RH: I253, RH: G267); przetwornik t — tm (F260).
HISTEREZA = {"th": "I253", "rh": "G267", "t": "F260"}

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


if __name__ == "__main__":
    zr = sys.argv[1] if len(sys.argv) > 1 else DOMYSLNE_ZRODLO
    ce = sys.argv[2] if len(sys.argv) > 2 else DOMYSLNY_CEL
    print(f"Szablon przetwornika: {zbuduj(zr, ce)}")
