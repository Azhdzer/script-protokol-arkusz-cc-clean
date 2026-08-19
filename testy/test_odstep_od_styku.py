# -*- coding: utf-8 -*-
"""
Testy odstepu 5 wierszy reprezentacyjnych od KONCA punktu pomiarowego.

Zgloszenie z realnego pomiaru: reprezentanci wypadli tuz przed zmiana nastawy
(punkt konczyl sie o 02:52, wybrane minuty siegaly 02:51). Na takim styku komora
zaczyna juz przechodzic do kolejnego punktu, a 15-minutowe rozrzuty lapia probki
zza granicy — odczyty wychodza "rozmazane".

Przyczyna: przy plaskich odczytach wszystkie okna maja identyczna srednia, a remis
rozstrzygany jest na korzysc POZNIEJSZEGO okna (porownanie <=). Bez odstepu wybor
zawsze przyklejal sie do konca punktu.
"""

import datetime
import io
import os
import sys
import unittest
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_obserwacje as G

POCZATEK = datetime.datetime(2026, 8, 14, 2, 30, 0)


def punkt(ile_minut, score=0.02):
    """Plaski punkt pomiarowy: odczyt co minute, wszedzie ten sam rozrzut."""
    dane = [(POCZATEK + datetime.timedelta(minutes=i), 25.0, 50.0, score, score)
            for i in range(ile_minut)]
    return dane, list(range(ile_minut))


def minuty(dane, indeksy):
    return [dane[i][0].strftime("%H:%M") for i in indeksy]


class TestOdstepOdZmianyNastawy(unittest.TestCase):

    def wybierz(self, dane, valid):
        return G._find_best_minute_reps(dane, valid, 0, len(dane),
                                        lambda i: dane[i][3])

    def test_reprezentanci_nie_dotykaja_konca_punktu(self):
        dane, valid = punkt(23)                      # 02:30..02:52
        wybor = self.wybierz(dane, valid)
        odstep = dane[-1][0] - dane[wybor[-1]][0]
        self.assertGreaterEqual(odstep, G.ODSTEP_OD_KONCA_PUNKTU)

    def test_bez_odstepu_wybor_przykleja_sie_do_konca(self):
        """Dowod, ze problem byl realny — z odstepem 0 wraca stare zachowanie."""
        dane, valid = punkt(23)
        stary = G.ODSTEP_OD_KONCA_PUNKTU
        G.ODSTEP_OD_KONCA_PUNKTU = datetime.timedelta(0)
        try:
            wybor = self.wybierz(dane, valid)
        finally:
            G.ODSTEP_OD_KONCA_PUNKTU = stary
        self.assertEqual(dane[wybor[-1]][0], dane[-1][0])

    def test_wybrane_minuty_sa_kolejne(self):
        dane, valid = punkt(23)
        wybor = self.wybierz(dane, valid)
        czasy = [dane[i][0] for i in wybor]
        odstepy = {(czasy[i + 1] - czasy[i]) for i in range(4)}
        self.assertEqual(odstepy, {datetime.timedelta(minutes=1)})

    def test_wybiera_piec_wierszy(self):
        dane, valid = punkt(23)
        self.assertEqual(len(self.wybierz(dane, valid)), 5)

    def test_lepszy_rozrzut_wygrywa_mimo_odstepu(self):
        """Odstep ogranicza tylko koniec okna — nadal szukamy najmniejszego rozrzutu."""
        dane, valid = punkt(30, score=0.5)
        dane = [(t, tz, rh, 0.01 if 5 <= i <= 9 else k, l)
                for i, (t, tz, rh, k, l) in enumerate(dane)]
        wybor = self.wybierz(dane, valid)
        self.assertEqual(minuty(dane, wybor),
                         ["02:35", "02:36", "02:37", "02:38", "02:39"])

    def test_krotki_punkt_dostaje_ostrzezenie(self):
        """Gdy punkt jest za krotki na odstep, bierzemy okno mimo wszystko."""
        dane, valid = punkt(6)                       # tylko 6 minut
        bufor = io.StringIO()
        with redirect_stdout(bufor):
            wybor = self.wybierz(dane, valid)
        self.assertIsNotNone(wybor)
        self.assertIn("za krotki", bufor.getvalue())

    def test_dlugi_punkt_nie_ostrzega(self):
        dane, valid = punkt(23)
        bufor = io.StringIO()
        with redirect_stdout(bufor):
            self.wybierz(dane, valid)
        self.assertEqual(bufor.getvalue(), "")

    def test_zbyt_malo_danych_zwraca_none(self):
        dane, valid = punkt(3)
        self.assertIsNone(self.wybierz(dane, valid))


class TestUstawienieOdstepu(unittest.TestCase):

    def setUp(self):
        import cc_config as C
        self.ust = C.WG_ENV["OBS_ODSTEP_KONIEC"]

    def test_domyslnie_dwie_minuty(self):
        self.assertEqual(self.ust.domyslna, 2)

    def test_da_sie_wylaczyc(self):
        self.assertEqual(self.ust.minimum, 0)

    def test_jest_w_grupie_okna_analizy(self):
        self.assertEqual(self.ust.grupa, "Okno analizy")


if __name__ == "__main__":
    unittest.main(verbosity=2)
