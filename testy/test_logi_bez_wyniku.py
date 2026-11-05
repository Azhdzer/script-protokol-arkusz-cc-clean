# -*- coding: utf-8 -*-
"""
Krok 2 ostrzega o logach przyrzadow z PZ, ktore nie maja pliku w 'wyniki'.

Zlecenie 232: w folderze kroku 1 bylo 14 logow, w protokole 11 przyrzadow. Lista plikow
kroku 1 pamieta poprzedni wybor, a nowe pliki pojawiaja sie niezaznaczone — trzy
przyrzady nie przeszly przez krok 1 i wypadly z protokolu bez slowa.
"""

import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_obserwacje as G
from wspolne import nowa_piaskownica


class TestLogiBezWyniku(unittest.TestCase):

    def setUp(self):
        f = nowa_piaskownica("logi_bez_wyniku")
        self.logi, self.wyniki = os.path.join(f, "logi"), os.path.join(f, "wyniki")
        os.makedirs(self.logi)
        os.makedirs(self.wyniki)
        for n in ("37297967.xls", "36679283.xls", "99999999.xls", "~$36679283.xls"):
            open(os.path.join(self.logi, n), "wb").close()
        for n in ("232_LA_TH_2026_37297967_opracowanie danych.xlsx", "37172104_wynik.xlsx"):
            open(os.path.join(self.wyniki, n), "wb").close()
        self.pz = {"37297967": object(), "36679283": object()}

    def bez_wyniku(self):
        with mock.patch.object(G, "FOLDER_LOGOW", self.logi), \
                mock.patch.object(G, "WYNIKI_FOLDER", self.wyniki):
            return G.logi_bez_wyniku(self.pz)

    def test_log_z_pz_bez_wyniku(self):
        self.assertEqual(self.bez_wyniku(), ["36679283.xls"])

    def test_swiezy_wynik_tez_liczy_sie(self):
        open(os.path.join(self.wyniki, "36679283_wynik.xlsx"), "wb").close()
        self.assertEqual(self.bez_wyniku(), [])

    def test_bez_pz_cisza(self):
        self.pz = {}
        self.assertEqual(self.bez_wyniku(), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
