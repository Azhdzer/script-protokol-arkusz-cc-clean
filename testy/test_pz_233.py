# -*- coding: utf-8 -*-
"""
Uklad PZ 233 — w protokole (Strona 2) nie bylo przyrzadow z PZ:

  * naglowek 'Obiekty wzorcowania' BEZ dwukropka (sekcja nie byla znajdowana),
  * numery rozerwane przez wyjustowany tekst PDF: 'nr fabr.: 371 72104',
    'nr ewid.: U90337 -18',
  * ponad 9 pozycji — '10) ' i dalsze sklejaly sie z pozycja 9 i przepadaly
    (to samo w PZ 07, 29, 48, 68, 99, 117 ...).

Numery w tescie sa zmyslone; uklad tekstu jak w PZ.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pz_dane as P


def _pz(n):
    wpisy = "".join(
        f"{i}) Termometr elektroniczny (rejestrator) typ: testo 174T, "
        f"nr fabr.: 371 {90000 + i}, nr ewid.: X123 -{i + 10}, \nwytwórca: Testo. \n"
        for i in range(1, n + 1))
    return ("Numery zleceń laboratorium: 233/LA/TH/2026 \n"
            "Obiekty wzorcowania \n" + wpisy +
            "Metoda wzorcowania: \nMetoda porównawcza w komorze klimatycznej. \n")


class TestPZ233(unittest.TestCase):

    def test_naglowek_bez_dwukropka_i_sklejone_numery(self):
        lista = P.parsuj_tekst(_pz(3))
        self.assertEqual([(p.nr_fabr, p.nr_ewid, p.typ, p.pozycja) for p in lista], [
            ("37190001", "X123-11", "testo 174T", 1),
            ("37190002", "X123-12", "testo 174T", 2),
            ("37190003", "X123-13", "testo 174T", 3)])

    def test_ponad_dziewiec_pozycji(self):
        lista = P.parsuj_tekst(_pz(28))
        self.assertEqual([p.pozycja for p in lista], list(range(1, 29)))
        self.assertEqual(lista[9].nr_fabr, "37190010")
        self.assertEqual(lista[-1].nr_ewid, "X123-38")

    def test_liczba_w_opisie_nie_jest_pozycja(self):
        # '12) ' poza kolejnoscia (po 1) nie otwiera nowej pozycji
        self.assertEqual(P._podziel_na_wpisy("1) Rejestrator (kanał 12) nr fabr.: 5 2) Drugi"),
                         ["", "1", "Rejestrator (kanał 12) nr fabr.: 5 ", "2", "Drugi"])

    def test_sklej_numer(self):
        for wej, wyj in (("371 72104", "37172104"), ("U90337 -18", "U90337-18"),
                         ("2 537/10", "2537/10"), ("PP/W/ZLM 12 021/T-04", "PP/W/ZLM 12021/T-04"),
                         ("524000646 04/2025", "524000646 04/2025"),   # numer + data produkcji
                         ("A 21120971", "A 21120971"), ("", ""), (None, None)):
            with self.subTest(wej=wej):
                self.assertEqual(P.sklej_numer(wej), wyj)


if __name__ == "__main__":
    unittest.main(verbosity=2)
