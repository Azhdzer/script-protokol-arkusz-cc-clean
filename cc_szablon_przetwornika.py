# -*- coding: utf-8 -*-
"""
cc_szablon_przetwornika.py — szablon arkusza obliczeniowego dla PRZETWORNIKA.

Standardowy szablon (… - 1 - RH (CC).xlsx) liczy przetwornik juz w zakladkach
roboczych (blok 'PRZETWORNIK TEMPERATURY I WILGOTNOSCI WZGLEDNEJ', wiersz 253),
ale arkusz 'Wyniki' zbiera wyniki TERMOHIGROMETRU (wiersz 239). Dla przetwornika
'Wyniki' musialy byc przebudowywane recznie (zlecenie 230):

    B  t      ='<zakladka>'!D253        G  RHm   ='<zakladka>'!I253
    C  RH     ='<zakladka>'!E253        H  Δt    ='<zakladka>'!J253
    D  Itm    ='<zakladka>'!F253        I  ΔRH   ='<zakladka>'!K253
    E  IRHm   ='<zakladka>'!G253        J  Ut    ='<zakladka>'!M253
    F  tm     ='<zakladka>'!H253        K  URH   ='<zakladka>'!O253

Naglowki D/E i jednostki zaleza od L10 zakladki (Prąd / Napięcie), histereza
liczy sie z RHm (I253) punktu i jego powtorki '(2)'.

Ten modul robi z szablonu standardowego szablon przetwornika — przez Excela
(wlasna instancja), wiec linki zewnetrzne, formaty i walidacje zostaja.
Uruchom ponownie, gdy przyjdzie nowa wersja szablonu standardowego:

    python cc_szablon_przetwornika.py [szablon_zrodlowy.xlsx] [szablon_docelowy.xlsx]
"""

import os
import re
import shutil
import sys

import cc_excel

WIERSZ_WYNIKOW = 253
# kolumna Wynikow -> kolumna wiersza 253 zakladki
KOLUMNY_WYNIKOW = (("B", "D"), ("C", "E"), ("D", "F"), ("E", "G"), ("F", "H"),
                   ("G", "I"), ("H", "J"), ("I", "K"), ("J", "M"), ("K", "O"))
# kolumny Wynikow dotyczace temperatury / wilgotnosci (bez wartosci odniesienia B/C)
KOLUMNY_T = ("D", "F", "H", "J")
KOLUMNY_RH = ("E", "G", "I", "K")
KOMORKA_HISTEREZY_ZAKLADKI = f"I{WIERSZ_WYNIKOW}"
# Formaty jak w recznie przygotowanych 'Wynikach' (zlecenie 230): prad/napiecie
# z 4 / 3 miejscami, wilgotnosc z 1, temperatura z 2.
FORMATY_LICZB = {"B": "0.00", "C": "0.0", "D": "0.0000", "E": "0.000", "F": "0.00",
                 "G": "0.0", "H": "0.00", "I": "0.0", "J": "0.00", "K": "0.0"}

DOMYSLNE_ZRODLO = ("xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 "
                   "z 17.06.2026 - 1 - RH (CC).xlsx")
DOMYSLNY_CEL = ("xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 "
                "z 17.06.2026 - 1 - Przetwornik (CC).xlsx")

_RE_ARKUSZ = re.compile(r"^=(?:'((?:[^']|'')+)'|([^!'][^!]*))!")


def _arkusz(formula):
    m = _RE_ARKUSZ.match(str(formula or ""))
    if not m:
        return None
    return (m.group(1) or m.group(2)).replace("''", "'")


def _ref(arkusz, komorka):
    return f"='{arkusz.replace(chr(39), chr(39) * 2)}'!{komorka}"


def naglowki(pierwsza):
    """Naglowki wierszy 1-3 'Wynikow' przetwornika (zalezne od L10 zakladki)."""
    a = f"'{pierwsza}'!$L$10"
    jedn = f'=IF({a}="-","-",IF({a}="Napięcie","V","mA"))'
    return {
        "A1": "Wyniki", "B1": "Wartość wielkości odniesienia",
        "D1": f'=IF({a}="-","-",IF({a}="Napięcie","Zmierzona wartość napięcia",'
              f'"Zmierzona wartość prądu"))',
        "F1": f'=IF({a}="-","-",IF({a}="Napięcie","Wartość temperatury i wilgotności '
              f'względnej odpowiadająca Utm oraz URHm","Wartość temperatury i wilgotności '
              f'względnej odpowiadająca Itm oraz IRHm"))',
        "H1": "Poprawka", "J1": "Niepewność pomiaru",
        "A2": "dla t i rh", "B2": "t", "C2": "RH",
        "D2": f'=IF({a}="-","-",IF({a}="Napięcie","Utm","Itm"))',
        "E2": f'=IF({a}="-","-",IF({a}="Napięcie","URHm","IRHm"))',
        "F2": "tm", "G2": "RHm", "H2": "Δt", "I2": "ΔRH", "J2": "Ut", "K2": "URH",
        "B3": "°C", "C3": "%rh", "D3": jedn, "E3": jedn, "F3": "°C", "G3": "%rh",
        "H3": "°C", "I3": "%rh", "J3": "°C", "K3": "%rh",
    }


def formuly_wiersza(arkusz):
    """{kolumna Wynikow: formula} dla jednej zakladki."""
    return {kw: _ref(arkusz, f"{kz}{WIERSZ_WYNIKOW}") for kw, kz in KOLUMNY_WYNIKOW}


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
        pierwsza = next((a for _r, a in arkusze if a == "23,30"), arkusze[0][1])

        # Tabela wynikow: B:K (o dwie kolumny szersza niz dla termohigrometru).
        ws.Range("H1:I1").UnMerge()
        ws.Range("H1:I9").Copy()
        ws.Range("J1:K9").PasteSpecial(-4122)          # xlPasteFormats
        app.api.CutCopyMode = False
        for adres in ("B1:C1", "D1:E1", "F1:G1", "H1:I1", "J1:K1"):
            ws.Range(adres).UnMerge()
            ws.Range(adres).Merge()
        for adres, wartosc in naglowki(pierwsza).items():
            ws.Range(adres).Formula = wartosc
        for r, a in arkusze:
            for kol, f in formuly_wiersza(a).items():
                ws.Range(f"{kol}{r}").Formula = f
                ws.Range(f"{kol}{r}").NumberFormat = FORMATY_LICZB[kol]
            ws.Range(f"O{r}:R{r}").ClearContents()     # tabela 'tylko temperatura' — nie dotyczy

        # Histereza: RHm punktu i jego powtorki.
        nazwy = [a for _r, a in arkusze]
        baza = next((a for a in nazwy if a + " (2)" in nazwy), None)
        if baza:
            ws.Range("J23").Formula = (
                f"=ABS(VALUE({_ref(baza, KOMORKA_HISTEREZY_ZAKLADKI)[1:]})"
                f"-VALUE({_ref(baza + ' (2)', KOMORKA_HISTEREZY_ZAKLADKI)[1:]}))")
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
