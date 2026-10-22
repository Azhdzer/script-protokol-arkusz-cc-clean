# -*- coding: utf-8 -*-
"""
Log PuTTY z wartosciami opisanymi JEDNOSTKAMI zamiast etykiet (D0540007.txt):

    =~=~=~=~=~= PuTTY log 2026.10.19 11:22:06 =~=~=~=~=~=
    2026-10-19 11:22:07    1012.9 mbar   22.4 'C   54 %RH

Parser szukal tylko 'T=' i 'RH=' i konczyl: 'Nie znaleziono linii pomiarowych'.
Uklad z etykietami musi dzialac jak dotad — nowy uklad jest DODANY obok.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import analizuj_excele as A
from wspolne import nowa_piaskownica

NAGLOWEK = "=~=~=~=~=~=~=~=~=~=~=~= PuTTY log 2026.10.19 11:22:06 =~=~=~=~=~=~=~=~=~=~=~=\r\n"


def przetworz(tresc, nazwa="log.txt"):
    folder = nowa_piaskownica("putty_jednostki")
    sciezka = os.path.join(folder, nazwa)
    with open(sciezka, "w", encoding="utf-8", newline="") as f:
        f.write(tresc)
    wyj = os.path.join(folder, "wyniki")
    os.makedirs(wyj, exist_ok=True)
    import io
    from contextlib import redirect_stdout
    from pathlib import Path
    with redirect_stdout(io.StringIO()):
        A.parse_putty(Path(sciezka), wyj)
    pliki = [p for p in os.listdir(wyj) if p.endswith(".xlsx")]
    if not pliki:
        return []
    wb = openpyxl.load_workbook(os.path.join(wyj, pliki[0]))
    wiersze = list(wb.active.iter_rows(min_row=2, values_only=True))
    wb.close()
    return wiersze


class TestWartosciZJednostkami(unittest.TestCase):

    def test_temperatura_i_wilgotnosc_po_jednostkach(self):
        w = przetworz(NAGLOWEK
                      + "2026-10-19 11:22:07    1012.9 mbar   22.4 'C   54 %RH\r\n"
                      + "2026-10-22 06:46:41     999.3 mbar   59.9 'C   96 %RH\r\n")
        self.assertEqual([(r[1], r[2]) for r in w], [(22.4, 54), (59.9, 96)])

    def test_cisnienie_nie_jest_temperatura(self):
        w = przetworz(NAGLOWEK + "2026-10-19 11:22:07    1012.9 mbar   22.4 'C   54 %RH\r\n")
        self.assertNotEqual(w[0][1], 1012.9)

    def test_znak_stopnia(self):
        w = przetworz(NAGLOWEK + "2026-10-19 11:22:07  1013 hPa  -5.3 °C  41.5 %RH\r\n")
        self.assertEqual((w[0][1], w[0][2]), (-5.3, 41.5))

    def test_urwana_linia_pomijana(self):
        w = przetworz(NAGLOWEK
                      + "2026-10-19 11:34:37    1012.9 mbar   22.4 'C  \r\n"
                      + "2026-10-19 11:35:07    1012.9 mbar   22.5 'C   54 %RH\r\n")
        self.assertEqual(len(w), 1)
        self.assertEqual(w[0][1], 22.5)


class TestUkladZEtykietamiBezZmian(unittest.TestCase):

    def test_etykiety_w_dowolnej_kolejnosci(self):
        w = przetworz(NAGLOWEK
                      + "2026-07-02 12:10:13 P=  998.2 hPa T= 24.0 'C RH= 55.5 %RH\r\n"
                      + "2026-07-02 11:20:26 T=   23.4 RH=    60.3 P=  998.1\r\n")
        self.assertEqual([(r[1], r[2]) for r in w], [(23.4, 60.3), (24.0, 55.5)])   # posortowane po czasie


if __name__ == "__main__":
    unittest.main(verbosity=2)
