# -*- coding: utf-8 -*-
"""
Korekta nastawy T w trakcie punktu (zlecenie 233, 29.10.2026):

    07:35  37,0 / 35     <- skrypt konczyl tu punkt i bral okno sprzed korekty
    07:36  37,3 / 35
    ...    37,4 / 35     <- operator podciagnal nastawe, zeby lepiej trafic w 37 °C
    09:56  37,4 / 65     <- dopiero tu zaczyna sie nastepny punkt

Zmiana nastawy T do OBS_KOREKTA_T (0,5 °C) od poczatku punktu to ten sam punkt;
zmiana RH zawsze konczy punkt.
"""

import datetime
import io
import os
import sys
import unittest
from contextlib import redirect_stdout
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_obserwacje as G

START = datetime.datetime(2026, 10, 29, 5, 0)


def wiersze(odcinki):
    """Wiersze pliku CC: odcinki = [(T_zad, RH_zad, T_odcz, minut), ...]."""
    rows, t = [], START
    for t_zad, rh_zad, t_odcz, minut in odcinki:
        for _ in range(minut):
            rows.append([t.strftime("%Y-%m-%d %H:%M:%S"), str(t_zad), str(rh_zad),
                         str(t_odcz), str(rh_zad), "114.3", "19.0", str(t_odcz),
                         str(rh_zad), str(t_odcz), "0.05", "0.02"])
            t += datetime.timedelta(minutes=1)
    return rows


def okna(odcinki, tol=0.5):
    data = G._prepare_data(wiersze(odcinki))
    with mock.patch.object(G, "TOL_KOREKTY_NASTAWY_T", tol), redirect_stdout(io.StringIO()):
        w = G._find_all_analysis_windows(data, G.STABILIZACJA_MIN)
    return [(data[si][0].strftime("%H:%M"), data[ei - 1][0].strftime("%H:%M"), data[ei - 1][1])
            for si, ei in w]


PRZEBIEG_233 = [
    (37.0, 35.0, 36.5, 155),     # 05:00-07:35 — dojscie i trzymanie 37 / 35
    (37.3, 35.0, 36.8, 10),      # 07:35 korekta
    (37.4, 35.0, 37.0, 130),     # do 09:55
    (37.4, 65.0, 37.0, 180),     # 09:55 — nastepny punkt (RH)
]


class TestKorektaNastawy(unittest.TestCase):

    def test_korekta_to_ten_sam_punkt(self):
        w = okna(PRZEBIEG_233)
        self.assertEqual(len(w), 2, w)
        # Punkt 37 / 35 konczy sie na zmianie RH, z nastawa po korekcie.
        self.assertEqual(w[0][1:], ("09:54", 37.4))

    def test_tolerancja_zero_jak_dawniej(self):
        w = okna(PRZEBIEG_233, tol=0.0)
        self.assertEqual([k for _p, k, _n in w][:1], ["07:34"])

    def test_wieksza_zmiana_to_nowy_punkt(self):
        w = okna([(37.0, 35.0, 37.0, 180), (37.8, 35.0, 37.8, 180)])
        self.assertEqual(len(w), 2, w)

    def test_tolerancja_liczona_od_poczatku_punktu(self):
        # 37,0 -> 37,4 -> 37,8: pelzanie po 0,4 nie skleja calego zakresu
        w = okna([(37.0, 35.0, 37.0, 180), (37.4, 35.0, 37.4, 60), (37.8, 35.0, 37.8, 180)])
        self.assertEqual(len(w), 2, w)

    def test_rampa_wejscia_nie_jest_nastawa_odniesienia(self):
        """Zlecenie 233/240: rampa ... -26,6 -> -27,0 -> -27,3 (kroki po 1 min), trzymane
        -27,3, potem korekta -27,6. Od -27,0 to 0,6 °C i punkt sie rozcinal."""
        w = okna([(-25.0, 0.0, -25.0, 180), (-26.3, 0.0, -26.3, 1), (-26.6, 0.0, -26.6, 1),
                  (-27.0, 0.0, -27.0, 1),
                  (-27.3, 0.0, -27.3, 160), (-27.6, 0.0, -27.6, 180), (20.0, 0.0, 20.0, 180)])
        punkt = [x for x in w if x[2] in (-27.3, -27.6)]
        self.assertEqual(len(punkt), 1, w)
        self.assertEqual(punkt[0][2], -27.6)

    def test_rampa_wyjscia_konczy_punkt(self):
        """70,3 trzymane, potem rampa 70,2 -> 70,1 -> ... co minute: punkt konczy sie
        na poczatku rampy, a nie 0,5 °C dalej."""
        rampa = [(round(70.2 - k * 0.1, 1), 0.0, 70.0, 1) for k in range(15)]
        w = okna([(70.3, 0.0, 70.3, 180)] + rampa + [(60.0, 0.0, 60.0, 180)])
        self.assertEqual(w[0][1:], ("07:59", 70.3))

    def test_zmiana_rh_zawsze_nowy_punkt(self):
        w = okna([(25.0, 30.0, 25.0, 180), (25.0, 32.0, 25.0, 180)])
        self.assertEqual(len(w), 2, w)

    def test_widelki_wg_nastawy_po_korekcie(self):
        # Komora trzyma 37,4 dopiero po korekcie — wejscie w widelki liczy sie od niej,
        # okno startuje 2 h po wejsciu, a nie od poczatku punktu.
        w = okna([(37.0, 35.0, 37.6, 60), (37.4, 35.0, 37.4, 240), (20.0, 35.0, 20.0, 180)])
        self.assertEqual(w[0][0], "08:00")


if __name__ == "__main__":
    unittest.main(verbosity=2)
