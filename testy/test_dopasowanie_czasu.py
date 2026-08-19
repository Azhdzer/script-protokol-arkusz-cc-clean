# -*- coding: utf-8 -*-
"""
Testy raportowania odchylki czasu przy dopasowaniu wynikow loggerow.

Kontekst: tolerancja dopasowania (OBS_TOL) to PROG ODRZUCENIA, a nie okno
wyszukiwania — algorytm zawsze bierze najblizszy rekord loggera. Przy gestym
zapisie (co 1 min) odchylka wynosi sekundy, ale przy rzadkim zapisie albo
dziurze w danych moze siegnac kilkunastu minut, a punkt pomiarowy trwa godziny.
Bez liczby w logu roznica miedzy jednym a drugim byla niewidoczna.
"""

import datetime
import io
import os
import sys
import unittest
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_obserwacje as G

T0 = datetime.datetime(2026, 8, 13, 14, 0, 0)


def cele(*przesuniecia_s):
    """Piec docelowych czasow wiersza protokolu."""
    return [T0 + datetime.timedelta(seconds=s) for s in przesuniecia_s]


def trafienia(*przesuniecia_s):
    """Piec dopasowanych rekordow loggera (krotka: czas, temp, rh, idx, temp2)."""
    return [(T0 + datetime.timedelta(seconds=s), 23.0, 45.0, i, None)
            for i, s in enumerate(przesuniecia_s)]


class TestCzasTxt(unittest.TestCase):

    def test_sekundy(self):
        self.assertEqual(G._czas_txt(12), "12 s")

    def test_granica_przechodzi_na_minuty(self):
        self.assertEqual(G._czas_txt(90), "1,5 min")

    def test_minuty_z_przecinkiem(self):
        self.assertEqual(G._czas_txt(930), "15,5 min")

    def test_brak_danych(self):
        self.assertEqual(G._czas_txt(None), "?")


class TestOdchylkiDopasowania(unittest.TestCase):

    def test_idealne_trafienie(self):
        dop = {0: trafienia(0, 60, 120, 180, 240)}
        punkty = [cele(0, 60, 120, 180, 240)]
        sr, mx = G._odchylki_dopasowania(dop, punkty)
        self.assertEqual((sr, mx), (0.0, 0.0))

    def test_liczy_srednia_i_maksimum(self):
        dop = {0: trafienia(10, 60, 120, 180, 240)}      # pierwszy o 10 s obok
        punkty = [cele(0, 60, 120, 180, 240)]
        sr, mx = G._odchylki_dopasowania(dop, punkty)
        self.assertEqual(mx, 10.0)
        self.assertEqual(sr, 2.0)

    def test_odchylka_jest_bezwzgledna(self):
        """Rekord sprzed czasu pomiaru liczy sie tak samo jak po nim."""
        dop = {0: trafienia(-30, 60, 120, 180, 240)}
        punkty = [cele(0, 60, 120, 180, 240)]
        _sr, mx = G._odchylki_dopasowania(dop, punkty)
        self.assertEqual(mx, 30.0)

    def test_wiele_punktow(self):
        """Maksimum liczone po WSZYSTKICH punktach — drugi jest caly o 10 min obok."""
        dop = {0: trafienia(0, 60, 120, 180, 240),
               1: trafienia(600, 660, 720, 780, 840)}
        punkty = [cele(0, 60, 120, 180, 240),
                  cele(0, 60, 120, 180, 240)]
        sr, mx = G._odchylki_dopasowania(dop, punkty)
        self.assertEqual(mx, 600.0)
        self.assertEqual(sr, 300.0)      # pierwszy punkt idealny, drugi o 600 s

    def test_korekta_zegara_jest_uwzgledniona(self):
        """Po korekcie zegara odchylke liczymy od czasow PRZESUNIETYCH."""
        shift = datetime.timedelta(minutes=5)
        dop = {0: trafienia(300, 360, 420, 480, 540)}
        punkty = [cele(0, 60, 120, 180, 240)]
        sr, mx = G._odchylki_dopasowania(dop, punkty, shift)
        self.assertEqual((sr, mx), (0.0, 0.0))

    def test_brak_dopasowan(self):
        self.assertEqual(G._odchylki_dopasowania({}, []), (None, None))


class TestOstrzezenieODopasowaniu(unittest.TestCase):
    """Ostrzezenie ma sie pojawiac tylko wtedy, gdy naprawde jest o czym mowic."""

    def komunikat(self, odch_max, tol_s):
        bufor = io.StringIO()
        with redirect_stdout(bufor):
            G._ostrzez_o_odchylce("37025098_wynik.xlsx", odch_max, tol_s)
        return bufor.getvalue()

    def test_gesty_zapis_nie_ostrzega(self):
        """Logger co 1 min przy tolerancji 30 min — odchylka rzedu sekund."""
        self.assertEqual(self.komunikat(25.0, 1800.0), "")

    def test_dopasowanie_sprzed_kwadransa_ostrzega(self):
        tresc = self.komunikat(900.0, 1800.0)
        self.assertIn("UWAGA", tresc)
        self.assertIn("15,0 min", tresc)

    def test_prog_nie_przekracza_dwoch_minut(self):
        """Nawet przy ogromnej tolerancji ostrzegamy powyzej 2 minut."""
        self.assertIn("UWAGA", self.komunikat(180.0, 14400.0))

    def test_ciasna_tolerancja_obniza_prog(self):
        """Przy tolerancji 30 s ostrzegamy juz powyzej ~7 s."""
        self.assertIn("UWAGA", self.komunikat(10.0, 30.0))

    def test_brak_danych_nie_ostrzega(self):
        self.assertEqual(self.komunikat(None, 1800.0), "")


class TestTolerancjaDopasowania(unittest.TestCase):
    """
    Loggery zapisuja co 1 min, a czesc z nich (Aranet / Efento) potrafi pominac
    2-3 probki. Tolerancja ma pokrywac taka przerwe, ale nie caly punkt pomiarowy
    — te trwaja po 2 h.
    """

    def setUp(self):
        import cc_config as C
        self.ust = C.WG_ENV["OBS_TOL"]

    def test_domyslna_pokrywa_przerwe_w_zapisie(self):
        """Przerwa 3 min daje odchylke do ~2 min — musi sie zmiescic."""
        self.assertGreaterEqual(self.ust.domyslna, 2.0)

    def test_domyslna_jest_ulamkiem_dlugosci_punktu(self):
        """Punkt trwa ok. 2 h; tolerancja rzedu pol godziny byla za luzna."""
        self.assertLessEqual(self.ust.domyslna, 5.0)

    def test_minimum_pozwala_na_pol_minuty(self):
        self.assertLessEqual(self.ust.minimum, 0.5)

    def test_typ_pozwala_na_ulamki(self):
        self.assertEqual(self.ust.typ, "liczba")

    def test_prog_ostrzezenia_przy_domyslnej_lapie_przerwe(self):
        """Przy 3 min ostrzezenie ma sie odezwac dla odchylki ~2 min."""
        tol_s = self.ust.domyslna * 60
        bufor = io.StringIO()
        with redirect_stdout(bufor):
            G._ostrzez_o_odchylce("plik.xlsx", 120.0, tol_s)
        self.assertIn("UWAGA", bufor.getvalue())

    def test_prog_ostrzezenia_milczy_przy_zapisie_co_minute(self):
        """Normalne dopasowanie (do 30 s) nie moze zasmiecac logu."""
        tol_s = self.ust.domyslna * 60
        bufor = io.StringIO()
        with redirect_stdout(bufor):
            G._ostrzez_o_odchylce("plik.xlsx", 28.0, tol_s)
        self.assertEqual(bufor.getvalue(), "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
