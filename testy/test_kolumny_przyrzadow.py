# -*- coding: utf-8 -*-
"""
Dokladanie par kolumn przyrzadow na Stronie 3 (ponad 10 z szablonu) — przez Excel.

Zgloszenie (20 przyrzadow, CC-04): nowe pary wstawiane ZA ostatnia para zostawialy
scalone naglowki 'Wskazania' / 'Przyrzady wzorcowane' (wiersze 8-9), obszar wydruku
i podzial strony (niebieska linia) na 10. parze, a gruba ramka konczyla sie w srodku.
Teraz kopia przedostatniej pary idzie PRZED ostatnia — jak przy recznym kopiowaniu.
"""

import io
import os
import shutil
import sys
import unittest
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.utils import get_column_letter as L

import generuj_obserwacje as G
from wspolne import KORZEN, nowa_piaskownica

SZABLONY = {
    "CC-04": ("xxx_LA_TH_2026 - protokół CC-04.xlsx", G.WYNIKI_START_COL_CC04),
    "CC": ("xxx_LA_TH_2026 - protokół CC.xlsx", G.WYNIKI_START_COL_CC),
}


def rozszerz(nazwa, ile):
    plik, start = SZABLONY[nazwa]
    zrodlo = os.path.join(KORZEN, plik)
    if not os.path.exists(zrodlo):
        raise unittest.SkipTest(f"brak szablonu {plik}")
    folder = nowa_piaskownica(f"kolumny_przyrzadow_{nazwa}_{ile}")
    cel = os.path.join(folder, "protokol.xlsx")
    shutil.copy2(zrodlo, cel)
    with redirect_stdout(io.StringIO()):
        dodane = G._rozszerz_kolumny_przyrzadow(cel, "Strona 3", start, ile)
    return openpyxl.load_workbook(cel)["Strona 3"], start, dodane


class _Wspolne:
    NAZWA = None
    ILE = None

    @classmethod
    def setUpClass(cls):
        cls.ws, cls.start, cls.dodane = rozszerz(cls.NAZWA, cls.ILE)
        plik, _ = SZABLONY[cls.NAZWA]
        cls.wzor = openpyxl.load_workbook(os.path.join(KORZEN, plik))["Strona 3"]
        cls.ostatnia = cls.start + (cls.ILE - 1) * 2          # pierwsza kolumna ostatniej pary
        cls.koniec = L(cls.ostatnia + 1)

    def test_dodano_pary(self):
        self.assertEqual(self.dodane, self.ILE - 10)

    def test_numery_przyrzadow(self):
        self.assertEqual([self.ws.cell(10, self.start + 2 * i).value for i in range(self.ILE)],
                         list(range(1, self.ILE + 1)))

    def test_scalone_naglowki_do_konca(self):
        zakresy = {str(r) for r in self.ws.merged_cells.ranges}
        self.assertIn(f"I8:{self.koniec}8", zakresy)
        self.assertIn(f"{L(self.start)}9:{self.koniec}9", zakresy)

    def test_obszar_wydruku_i_podzial_strony(self):
        self.assertTrue(str(self.ws.print_area).endswith(f"${self.koniec}$55"), self.ws.print_area)
        self.assertEqual([L(b.id) for b in self.ws.col_breaks.brk], [L(self.ostatnia + 2)])

    def test_gruba_ramka_na_koncu_jak_w_szablonie(self):
        wzor_ost = self.start + 18                           # ostatnia (10.) para szablonu
        for r in (8, 9, 15, 25):
            with self.subTest(wiersz=r):
                self.assertEqual(self.ws.cell(r, self.ostatnia + 1).border.right.style,
                                 self.wzor.cell(r, wzor_ost + 1).border.right.style)
                # srodkowa para jak srodkowa w szablonie (bez ramki konca tabeli)
                self.assertEqual(self.ws.cell(r, self.ostatnia - 1).border.right.style,
                                 self.wzor.cell(r, wzor_ost - 1).border.right.style)

    def test_pary_scalone_w_naglowku(self):
        zakresy = {str(r) for r in self.ws.merged_cells.ranges}
        for i in range(self.ILE):
            c = self.start + 2 * i
            self.assertIn(f"{L(c)}10:{L(c + 1)}11", zakresy)

    def test_naglowek_kolumny_t_relatywny(self):
        f = str(self.ws.cell(18, self.ostatnia).value)
        self.assertIn(f"{L(self.ostatnia)}10", f)

    def test_kolumny_za_tabela_przesuniete(self):
        # Blok na prawo od przyrzadow (AM1:AP7 w CC-04) przesuniety o dodane pary.
        zakresy = {str(r) for r in self.wzor.merged_cells.ranges if r.min_row == 1 and r.max_row == 7}
        przes = 2 * (self.ILE - 10)
        for z in zakresy:
            r = openpyxl.worksheet.cell_range.CellRange(z)
            self.assertIn(f"{L(r.min_col + przes)}1:{L(r.max_col + przes)}7",
                          {str(x) for x in self.ws.merged_cells.ranges})


class TestCC04Dwadziescia(_Wspolne, unittest.TestCase):
    NAZWA, ILE = "CC-04", 20


class TestCCDwanascie(_Wspolne, unittest.TestCase):
    NAZWA, ILE = "CC", 12


if __name__ == "__main__":
    unittest.main(verbosity=2)
