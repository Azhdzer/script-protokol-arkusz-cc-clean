# -*- coding: utf-8 -*-
"""
Nazwa kopii arkusza obliczeniowego (krok 3 i 6): po nr fabrycznym stoi jeszcze
nr ewidencyjny — '243_LA_TH_2026 - ... - 1 - 84186135 - 2122-DW-T.xlsx'.

Nr fabryczny i numer przyrzadu czytane z nazwy (G6, PZ, swiadectwa w kroku 5)
musza wychodzic tak samo z nowych nazw i ze starych (bez nr ewidencyjnego).
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_arkusze as G
import generuj_swiadectwa as S

SZABLON = "xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 z 17.06.2026 - 1 - RH (CC).xlsx"
RDZEN = "243_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 z 17.06.2026"


class _BezPZ(unittest.TestCase):
    def setUp(self):
        self._pz = G._PZ_CACHE
        G._PZ_CACHE = {}

    def tearDown(self):
        G._PZ_CACHE = self._pz


class TestNazwaKopii(_BezPZ):

    def nazwa(self, e, f, nr=1):
        return G._bezpieczna_nazwa_pliku(
            G.nazwa_kopii(SZABLON, {"O": "243", "E": e, "F": f}, nr_przyrzadu=nr))

    def test_fabryczny_i_ewidencyjny(self):
        self.assertEqual(self.nazwa("84186135", "2122-DW-T"),
                         f"{RDZEN} - 1 - 84186135 - 2122-DW-T.xlsx")

    def test_bez_ewidencyjnego_jak_dotad(self):
        for f in ("", "-", None):
            self.assertEqual(self.nazwa("84186135", f, 4), f"{RDZEN} - 4 - 84186135.xlsx")

    def test_bez_fabrycznego_ewidencyjny_raz(self):
        self.assertEqual(self.nazwa("-", "2122-DW-T"), f"{RDZEN} - 1 - 2122-DW-T.xlsx")

    def test_ten_sam_numer_nie_dublowany(self):
        self.assertEqual(self.nazwa("84186135", "84186135"), f"{RDZEN} - 1 - 84186135.xlsx")

    def test_znaki_niedozwolone(self):
        self.assertEqual(self.nazwa("Q/4", "12/DW"), f"{RDZEN} - 1 - Q_4 - 12_DW.xlsx")

    def test_ewidencyjny_z_pz_gdy_strona2_pusta(self):
        class _Dev:
            nr_ewid = "2122-DW-T"
        G._PZ_CACHE = {G.pz_dane.normalizuj_serial("84186135"): _Dev()}
        self.assertEqual(self.nazwa("84186135", "-"),
                         f"{RDZEN} - 1 - 84186135 - 2122-DW-T.xlsx")


class TestOdczytZNazwy(unittest.TestCase):

    def test_nowa_nazwa(self):
        n = f"{RDZEN} - 12 - 84404577 - 2122-DW-T.xlsx"
        self.assertEqual(G._parsuj_nazwe_pliku(n), ("243", "2026", "84404577", "ILAJ 5.4_11#21"))
        self.assertEqual(S._klucz_kolejnosci(n)[:2], ("243", 12))

    def test_stara_nazwa(self):
        n = f"{RDZEN} - 13 - 84404769.xlsx"
        self.assertEqual(G._parsuj_nazwe_pliku(n), ("243", "2026", "84404769", "ILAJ 5.4_11#21"))
        self.assertEqual(S._klucz_kolejnosci(n)[:2], ("243", 13))

    def test_nazwa_bez_numeru_przyrzadu(self):
        self.assertEqual(G._parsuj_nazwe_pliku("133_LA_TH_2026 - ILAJ 5.4_11#21 - 1010223.xlsx")[2],
                         "1010223")


if __name__ == "__main__":
    unittest.main(verbosity=2)
