# -*- coding: utf-8 -*-
"""
Testy przypisywania punktow z PZ do segmentow obserwacji.

Zgloszenie z pomiaru 191: punkt (23 °C, 30 %rh) trafil na segment o nastawie
dokladnie 23,0/30,0 — a to byl etap SUSZENIA, na ktorym komora punktu nie
utrzymala (odczyt ~45 %rh). Prawdziwy pomiar mial nastawe 23,0/28,0 (operator
celowo zaniza nastawe, bo komora „przestrzeliwuje" do 30 %) i zostal odrzucony
jako „spoza zamowienia".

Przyczyna: ranking kandydatow liczyl odleglosc po NASTAWIE komory. Teraz
decyduje ODCZYT — to, co komora naprawde utrzymywala.
"""

import datetime
import io
import os
import sys
import unittest
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_obserwacje as G

T0 = datetime.datetime(2026, 8, 17, 14, 0, 0)


def segment(nastawa_t, nastawa_rh, odczyt_t, odczyt_rh, ile=10, przesun_h=0):
    """
    Wiersze jednego segmentu w ukladzie CC:
    (dt, Tzadana, RHzadana, K, L, Todczytana, RHodczytana).
    """
    poczatek = T0 + datetime.timedelta(hours=przesun_h)
    return [(poczatek + datetime.timedelta(minutes=i), nastawa_t, nastawa_rh,
             0.02, 0.02, odczyt_t, odczyt_rh) for i in range(ile)]


def zbuduj(segmenty):
    """Skleja segmenty w jedna liste `data` i zwraca (data, windows)."""
    dane, okna = [], []
    for seg in segmenty:
        start = len(dane)
        dane.extend(seg)
        okna.append((start, len(dane)))
    return dane, okna


def wybierz(dane, okna, punkty):
    """Zwraca (indeksy wybranych okien, log)."""
    with redirect_stdout(io.StringIO()) as buf:
        wynik = G._wybierz_okna_wg_pz(dane, okna, punkty, 5, 6)
    indeksy = [okna.index(okno) for okno, _punkt in wynik]
    return indeksy, buf.getvalue()


class TestWyborPoOdczycie(unittest.TestCase):

    def test_wygrywa_segment_ktory_utrzymal_punkt(self):
        """
        Segment 0: nastawa 23/28, komora trzyma 30 %  <- prawdziwy pomiar
        Segment 1: nastawa 23/30, komora poszla na 45 %  <- suszenie
        """
        dane, okna = zbuduj([
            segment(23.0, 28.0, 23.0, 30.0),
            segment(23.0, 30.0, 23.0, 45.0, przesun_h=5),
        ])
        indeksy, _log = wybierz(dane, okna, [(23.0, 30.0)])
        self.assertEqual(indeksy, [0])

    def test_segment_suszenia_zostaje_pominiety(self):
        dane, okna = zbuduj([
            segment(23.0, 28.0, 23.0, 30.0),
            segment(23.0, 30.0, 23.0, 45.0, przesun_h=5),
        ])
        _indeksy, log = wybierz(dane, okna, [(23.0, 30.0)])
        self.assertIn("Segment 2", log)
        self.assertIn("spoza zamowienia", log)

    def test_nastawa_rowna_punktowi_nie_daje_przewagi(self):
        """Sama zgodnosc nastawy nie moze przewazyc nad tym, co komora trzymala."""
        dane, okna = zbuduj([
            segment(23.0, 30.0, 23.0, 44.0),      # nastawa idealna, odczyt zly
            segment(23.0, 27.0, 23.0, 30.0, przesun_h=5),   # nastawa gorsza, odczyt idealny
        ])
        indeksy, _log = wybierz(dane, okna, [(23.0, 30.0)])
        self.assertEqual(indeksy, [1])

    def test_bez_odczytow_dziala_jak_dawniej(self):
        """Gdy kolumn odczytu brak, wracamy do porownania po nastawie."""
        dane, okna = zbuduj([
            segment(23.0, 28.0, None, None),
            segment(23.0, 30.0, None, None, przesun_h=5),
        ])
        indeksy, _log = wybierz(dane, okna, [(23.0, 30.0)])
        self.assertEqual(indeksy, [1])

    def test_punkt_tylko_temperaturowy(self):
        dane, okna = zbuduj([
            segment(37.0, 0.0, 37.4, 0.0),
            segment(4.0, 0.0, 4.2, 0.0, przesun_h=5),
        ])
        indeksy, _log = wybierz(dane, okna, [(4.0, None)])
        self.assertEqual(indeksy, [1])

    def test_powtorzony_punkt_bierze_dwa_segmenty(self):
        """Histereza: 60 %rh zamowione dwa razy — musza wejsc dwa rozne okna."""
        dane, okna = zbuduj([
            segment(23.0, 58.0, 23.0, 60.0),
            segment(23.0, 58.0, 23.0, 60.0, przesun_h=5),
        ])
        indeksy, _log = wybierz(dane, okna, [(23.0, 60.0), (23.0, 60.0)])
        self.assertEqual(sorted(indeksy), [0, 1])

    def test_kolejnosc_wynikow_jest_chronologiczna(self):
        dane, okna = zbuduj([
            segment(35.0, 57.0, 35.0, 60.0),
            segment(10.0, 59.0, 10.0, 60.0, przesun_h=5),
        ])
        indeksy, _log = wybierz(dane, okna, [(10.0, 60.0), (35.0, 60.0)])
        self.assertEqual(indeksy, [0, 1])

    def test_brak_kandydata_jest_zglaszany(self):
        dane, okna = zbuduj([segment(23.0, 28.0, 23.0, 30.0)])
        _indeksy, log = wybierz(dane, okna, [(50.0, 20.0)])
        self.assertIn("BRAK pasujacego segmentu", log)


class TestTolerancjaNadalObowiazuje(unittest.TestCase):
    """Odczyt decyduje o wyborze, ale kandydatow nadal zawezamy po nastawie."""

    def test_nastawa_poza_tolerancja_odpada(self):
        # nastawa RH 20 % przy punkcie 30 % to roznica 10 > TOL_PUNKT_RH (4)
        dane, okna = zbuduj([segment(23.0, 20.0, 23.0, 30.0)])
        _indeksy, log = wybierz(dane, okna, [(23.0, 30.0)])
        self.assertIn("BRAK pasujacego segmentu", log)


if __name__ == "__main__":
    unittest.main(verbosity=2)
