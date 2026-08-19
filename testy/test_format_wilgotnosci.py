# -*- coding: utf-8 -*-
"""
Testy formatowania warunkow srodowiskowych w swiadectwach Word.

Zgloszenie: w KAZDYM swiadectwie wilgotnosc wzgledna wychodzila z doklejonym
',0' — 'Wilgotnosc wzgledna: (30,0 ÷ 54,0) %'. Wilgotnosc podaje sie w pelnych
procentach, wiec ten przecinek sugerowal rozdzielczosc, ktorej pomiar nie ma.

Temperatura otoczenia zostaje bez zmian — jedno miejsce po przecinku.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_arkusze as A

F = A._formatuj_zakres_srodowiskowy


class TestWilgotnoscBezPrzecinka(unittest.TestCase):

    def test_wartosc_calkowita(self):
        self.assertEqual(F(30.0, 0), "30")

    def test_ulamek_w_dol(self):
        self.assertEqual(F(53.4, 0), "53")

    def test_ulamek_w_gore(self):
        self.assertEqual(F(53.6, 0), "54")

    def test_polowka_zaokragla_w_gore(self):
        """Wbudowane round() dalo by 30 (zaokraglenie bankowe) — tu ma byc 31."""
        self.assertEqual(F(30.5, 0), "31")

    def test_druga_polowka_tez_w_gore(self):
        self.assertEqual(F(31.5, 0), "32")

    def test_brak_przecinka_w_wyniku(self):
        for wartosc in (30.0, 44.44, 53.99, 12.5):
            with self.subTest(wartosc=wartosc):
                self.assertNotIn(",", F(wartosc, 0))

    def test_wartosc_z_tekstu(self):
        self.assertEqual(F("47.4", 0), "47")


class TestTemperaturaBezZmian(unittest.TestCase):

    def test_jedno_miejsce_po_przecinku(self):
        self.assertEqual(F(21.8, 1), "21,8")

    def test_domyslnie_jedno_miejsce(self):
        self.assertEqual(F(23.25), "23,2")

    def test_przecinek_dziesietny_po_polsku(self):
        self.assertIn(",", F(23.3, 1))

    def test_wartosc_calkowita_dostaje_zero(self):
        self.assertEqual(F(23.0, 1), "23,0")


class TestPrzypadkiBrzegowe(unittest.TestCase):

    def test_brak_wartosci(self):
        self.assertEqual(F(None, 0), "—")
        self.assertEqual(F(None, 1), "—")

    def test_wartosc_nieliczbowa_zostaje_tekstem(self):
        self.assertEqual(F("brak danych", 0), "brak danych")

    def test_zero(self):
        self.assertEqual(F(0.0, 0), "0")

    def test_wartosc_ujemna_temperatury(self):
        """Temperatura otoczenia moze byc ujemna — format nie moze sie wywrocic."""
        self.assertEqual(F(-19.1, 1), "-19,1")


if __name__ == "__main__":
    unittest.main(verbosity=2)
