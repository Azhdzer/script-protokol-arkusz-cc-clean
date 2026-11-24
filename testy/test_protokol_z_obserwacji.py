# -*- coding: utf-8 -*-
"""
Krok 2b — protokol z punktow poprawionych recznie w obserwacji (zlecenie 243).

Krok 2 wybral czesc punktow inaczej, niz trzeba (61,2/25 i 61/59 poza tolerancja
nastaw, 70,2 zamiast 70,3). Operator oznaczyl wlasne punkty jak zwykle: 5 wierszy
reprezentantow z wypelnieniem i pogrubieniem — kolorem MOTYWU ('zielony, akcent 6,
jasniejszy 40%'), a nie #A9D08E jak skrypt — i zdjal pogrubienie z odrzuconego wyboru.
Liczby w tescie sa zmyslone; uklad jak w obserwacji CC.
"""

import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color

import generuj_obserwacje as G

START = datetime.datetime(2026, 11, 19, 16, 0, 0)
ZIELONY_SKRYPT = PatternFill("solid", fgColor="A9D08E")
ZIELONY_MOTYW = PatternFill("solid", fgColor=Color(theme=9, tint=0.3999755851924192))
JASNY = PatternFill("solid", fgColor="E2EFDA")
POMARANCZ = PatternFill("solid", fgColor="F4B183")
BIALY = PatternFill("solid", fgColor="FFFFFF")


def dane(odcinki):
    """odcinki = [(T, RH, ile_minut)] -> data jak _prepare_data (co 1 min)."""
    wynik, t = [], START
    for temp, rh, n in odcinki:
        for _ in range(n):
            wynik.append((t, temp, rh, 0.02, 0.02, temp, rh))
            t += datetime.timedelta(minutes=1)
    return wynik


def arkusz(data, oznaczenia):
    """oznaczenia: {indeks_danych: (wypelnienie, pogrubienie)}."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(1, 1, "Data Czas")
    for i, d in enumerate(data):
        c = ws.cell(2 + i, 1, d[0])
        if i in oznaczenia:
            fill, bold = oznaczenia[i]
            c.fill = fill
            c.font = Font(bold=bold)
    return ws


def punkt(start, fill=ZIELONY_MOTYW, ile=5, co=1):
    return {start + k * co: (fill, True) for k in range(ile)}


DANE = dane([(60.6, 25.5, 60), (61.2, 25.0, 60), (61.0, 59.0, 60), (70.2, 0.0, 30), (70.3, 0.0, 60)])


class TestPunktyZObserwacji(unittest.TestCase):

    def test_punkty_operatora_kolorem_motywu_i_skryptu(self):
        ozn = {**punkt(40, ZIELONY_SKRYPT), **punkt(100), **punkt(170, co=2)}
        grupy, problemy = G.punkty_z_obserwacji(arkusz(DANE, ozn), DANE)
        self.assertEqual(problemy, [])
        self.assertEqual([g for g, _p in grupy],
                         [[40, 41, 42, 43, 44], [100, 101, 102, 103, 104],
                          [170, 172, 174, 176, 178]])

    def test_zdjete_pogrubienie_i_jasny_blok_nie_sa_punktem(self):
        ozn = {**punkt(40), **{i: (JASNY, False) for i in range(195, 205)},
               **{i: (ZIELONY_SKRYPT, False) for i in range(200, 205)}}
        grupy, _ = G.punkty_z_obserwacji(arkusz(DANE, ozn), DANE)
        self.assertEqual(len(grupy), 1)

    def test_biale_wypelnienie_z_pogrubieniem_to_nie_punkt(self):
        grupy, _ = G.punkty_z_obserwacji(arkusz(DANE, punkt(40, BIALY)), DANE)
        self.assertEqual(grupy, [])

    def test_sasiednie_punkty_rozdzielone_po_nastawie(self):
        """Ostatnie minuty 60,6/25,5 i pierwsze 61,2/25 — dwie grupy, nie jedna z 10."""
        ozn = {**punkt(55), **punkt(60)}
        grupy, problemy = G.punkty_z_obserwacji(arkusz(DANE, ozn), DANE)
        self.assertEqual(problemy, [])
        self.assertEqual(len(grupy), 2)

    def test_zla_liczba_wierszy_to_problem(self):
        _g, problemy = G.punkty_z_obserwacji(arkusz(DANE, punkt(40, ile=4)), DANE)
        self.assertEqual(len(problemy), 1)
        self.assertIn("4 wierszy", problemy[0])

    def test_wiersz_spoza_txt_to_problem(self):
        ws = arkusz(DANE, punkt(40))
        c = ws.cell(2 + 300, 1, datetime.datetime(2030, 1, 1))
        c.fill, c.font = ZIELONY_MOTYW, Font(bold=True)
        _g, problemy = G.punkty_z_obserwacji(ws, DANE)
        self.assertTrue(any("brak takiego czasu" in p for p in problemy))

    def test_punkty_po_czasie_nie_po_numerze_wiersza(self):
        """Wiersz usuniety recznie w obserwacji nie przesuwa punktow."""
        ws = arkusz(DANE, punkt(100))
        ws.delete_rows(5, 1)
        grupy, problemy = G.punkty_z_obserwacji(ws, DANE)
        self.assertEqual(problemy, [])
        self.assertEqual(grupy[0][0], [100, 101, 102, 103, 104])

    def test_pomaranczowy_punkt_zachowuje_ostrzezenie(self):
        grupy, _ = G.punkty_z_obserwacji(arkusz(DANE, punkt(40, POMARANCZ)), DANE)
        self.assertIsNotNone(grupy[0][1])


class TestPunktyPZ(unittest.TestCase):

    def test_przypisanie_i_powtorzenie_bez_punktu(self):
        grupy = [([40, 41, 42, 43, 44], None), ([100, 101, 102, 103, 104], None),
                 ([230, 231, 232, 233, 234], None)]
        wynik = G.przypisz_punkty_pz(DANE, grupy, [(60, 25), (70, None)], 5, 6)
        self.assertEqual(wynik, [(60, 25), None, (70, None)])


class TestWyborObserwacji(unittest.TestCase):

    def test_numer_pomiaru_z_nazwy_obserwacji(self):
        self.assertEqual(G.numer_pomiaru_z_obserwacji("243_LA_TH_2026 - obserwacje.xlsx"), "243")
        self.assertEqual(G.numer_pomiaru_z_obserwacji(
            "242, 245, 256_LA_TH_2026 - obserwacje CC.xlsx"), "242, 245, 256")
        self.assertIsNone(G.numer_pomiaru_z_obserwacji("obserwacje.xlsx"))

    def test_nazwa_protokolu_jak_w_kroku_2(self):
        nr = G.numer_pomiaru_z_obserwacji(G.nazwa_obserwacji("243", "CC04"))
        self.assertEqual(nr, "243")


class TestCzyszczenieOznaczen(unittest.TestCase):

    def test_stare_oznaczenia_zdjete_naglowek_i_podpis_zostaja(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        for c, n in enumerate(("Czas", "Temperatura [°C]", "Wilgotność [%RH]"), 1):
            ws.cell(1, c, n).font = Font(bold=True)
        for r in range(2, 12):
            for c in range(1, 4):
                ws.cell(r, c, r)
        for c in range(1, 4):
            ws.cell(5, c).fill = ZIELONY_SKRYPT
            ws.cell(5, c).font = Font(bold=True)
        ws.cell(5, 5, 3).fill = ZIELONY_SKRYPT       # numer punktu
        ws.cell(1, 7, "Opracował:")                  # podpis za znacznikiem
        G._wyczysc_oznaczenia_punktow(ws, 3, 5)
        self.assertFalse(ws.cell(5, 1).fill.fill_type)
        self.assertFalse(ws.cell(5, 1).font.b)
        self.assertIsNone(ws.cell(5, 5).value)
        self.assertTrue(ws.cell(1, 1).font.b)
        self.assertEqual(ws.cell(1, 7).value, "Opracował:")


if __name__ == "__main__":
    unittest.main(verbosity=2)
