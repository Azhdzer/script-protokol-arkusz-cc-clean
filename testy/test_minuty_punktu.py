# -*- coding: utf-8 -*-
"""
Wiersze przyrzadu dobierane do PELNYCH minut punktu (zlecenie 232, 234).

Obserwacja: 18:19:47, 18:20:46, 18:22:14, 18:23:14, 18:24:14 = minuty 20-24.
Logger z zapisem na :25 dostawal 19:25, 20:25, 22:25, 23:25, 24:25 — minuta 21
wypadala, a przyrzad byl czytany z innych minut niz obserwacja.
"""

import datetime as D
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_obserwacje as G


def t(tekst):
    return D.datetime.strptime("2026-10-30 " + tekst, "%Y-%m-%d %H:%M:%S")


OBSERWACJA = [t(x) for x in ("18:19:47", "18:20:46", "18:22:14", "18:23:14", "18:24:14")]


def logger(sekunda, co_ile_s=60):
    """Rekordy [(dt, temp, rh, row_idx, temp2)] od 18:00 z zapisem na podanej sekundzie."""
    start = t(f"18:00:{sekunda:02d}")
    return [(start + D.timedelta(seconds=co_ile_s * i), 30.3, None, i + 2, None)
            for i in range(60 * 60 // co_ile_s)]


def wybrane(rekordy):
    return [k[0].strftime("%H:%M:%S")
            for k in G._znajdz_5_wierszy(rekordy, G._minuty_punktu(OBSERWACJA))]


class TestMinutyPunktu(unittest.TestCase):

    def test_pelne_minuty(self):
        self.assertEqual([x.strftime("%H:%M:%S") for x in G._minuty_punktu(OBSERWACJA)],
                         ["18:20:00", "18:21:00", "18:22:00", "18:23:00", "18:24:00"])

    def test_logger_na_25_sekundzie_jak_w_zgloszeniu(self):
        self.assertEqual(wybrane(logger(25)),
                         ["18:20:25", "18:21:25", "18:22:25", "18:23:25", "18:24:25"])

    def test_zadna_minuta_nie_wypada_dla_kazdej_sekundy_zapisu(self):
        for sekunda in range(60):
            with self.subTest(sekunda=sekunda):
                czasy = [k[0] for k in G._znajdz_5_wierszy(logger(sekunda), G._minuty_punktu(OBSERWACJA))]
                self.assertEqual({(b - a).total_seconds() for a, b in zip(czasy, czasy[1:])}, {60.0})

    def test_gesty_logger_trafia_w_pelna_minute(self):
        self.assertEqual(wybrane(logger(0, co_ile_s=10)),
                         ["18:20:00", "18:21:00", "18:22:00", "18:23:00", "18:24:00"])

    def test_pol_minuty_w_gore_jak_w_obserwacji(self):
        self.assertEqual(G._minuty_punktu([t("18:19:30")]), [t("18:20:00")])
        self.assertEqual(G._minuty_punktu([t("18:19:29")]), [t("18:19:00")])

    def test_powtorzona_minuta_zostawia_czasy(self):
        czasy = [t("18:20:47"), t("18:21:03"), t("18:22:03")]
        self.assertEqual(G._minuty_punktu(czasy), czasy)

    def test_brak_czasu_zostawia_czasy(self):
        czasy = [t("18:20:47"), None]
        self.assertEqual(G._minuty_punktu(czasy), czasy)


if __name__ == "__main__":
    unittest.main(verbosity=2)
