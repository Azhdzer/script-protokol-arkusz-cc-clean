# -*- coding: utf-8 -*-
"""
Zapas sredniej ponad najlepsze okno (zlecenie 232, punkt 37 °C / 65 %):

    12:15-12:19  K = 0,06   <- najlepsza srednia, wybierane dotad
    ...          K = 0,07 ... 0,12
    14:02        zmiana nastawy

Przyrzady dochodza do nastaw wolniej niz komora. Z okien spelniajacych kryterium,
ktorych srednia miesci sie w najlepsza + OBS_ZAPAS_KRYT (0,05), bierzemy NAJPOZNIEJSZE;
najlepsze zostaje jasno zaznaczone jako alternatywa (bez pogrubienia, nie do protokolu).
"""

import datetime
import io
import os
import sys
import unittest
from contextlib import redirect_stdout
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import generuj_obserwacje as G

POCZATEK = datetime.datetime(2026, 10, 29, 12, 0)


def punkt(k_wg_minuty):
    """Punkt z odczytem co minute; K (score) zalezy od minuty."""
    dane = [(POCZATEK + datetime.timedelta(minutes=i), 37.4, 65.0, k, k)
            for i, k in enumerate(k_wg_minuty)]
    return dane, list(range(len(dane)))


# 12:00-12:14 K 0,09 | 12:15-12:19 K 0,06 | 12:20-13:39 K 0,08 | 13:40-14:01 K 0,12
K_232 = [0.09] * 15 + [0.06] * 5 + [0.08] * 80 + [0.12] * 22


def wybor(dane, valid, zapas, alternatywa=None):
    with mock.patch.object(G, "ZAPAS_KRYTERIUM", zapas), redirect_stdout(io.StringIO()):
        w = G._find_best_minute_reps(dane, valid, 0, len(dane), lambda i: dane[i][3],
                                     alternatywa=alternatywa)
    return [dane[i][0].strftime("%H:%M") for i in w]


class TestWyborOkna(unittest.TestCase):

    def test_najpozniejsze_w_zapasie(self):
        dane, valid = punkt(K_232)
        alt = []
        # Okna z K 0,08 (<= 0,06 + 0,05) az do styku z 0,12; okno z dwiema minutami
        # 0,12 ma srednia 0,096 — tez w zapasie.
        w = wybor(dane, valid, 0.05, alt)
        self.assertEqual(w, ["13:38", "13:39", "13:40", "13:41", "13:42"])
        self.assertEqual(len(alt), 1)
        najlepsza, okno, wybrana = alt[0]
        self.assertAlmostEqual(najlepsza, 0.06)
        self.assertEqual(dane[okno[0]][0].strftime("%H:%M"), "12:15")
        self.assertLessEqual(wybrana, 0.06 + 0.05 + 1e-9)

    def test_zapas_zero_jak_dawniej(self):
        dane, valid = punkt(K_232)
        alt = []
        self.assertEqual(wybor(dane, valid, 0.0, alt)[0], "12:15")
        self.assertEqual(alt, [])

    def test_poza_zapasem_nie_brane(self):
        dane, valid = punkt([0.06] * 5 + [0.40] * 30)
        self.assertEqual(wybor(dane, valid, 0.05)[0], "12:00")

    def test_odstep_od_konca_nadal_dziala(self):
        dane, valid = punkt([0.06] * 30)
        w = wybor(dane, valid, 0.05)
        self.assertLessEqual(datetime.datetime.strptime(w[-1], "%H:%M").minute,
                             29 - int(G.ODSTEP_OD_KONCA_PUNKTU.total_seconds() // 60))

    def test_wybrane_minuty_kolejne(self):
        dane, valid = punkt(K_232)
        w = [datetime.datetime.strptime(x, "%H:%M") for x in wybor(dane, valid, 0.05)]
        self.assertEqual({(b - a).seconds for a, b in zip(w, w[1:])}, {60})


class TestWybierzOkno(unittest.TestCase):

    def test_remis_pozniejsze(self):
        self.assertEqual(G._wybierz_okno([(0.1, [1]), (0.1, [2])], 0.0), [2])

    def test_brak_kandydatow(self):
        self.assertIsNone(G._wybierz_okno([], 0.05))

    def test_alternatywa_tylko_gdy_rozne(self):
        alt = []
        G._wybierz_okno([(0.06, [1]), (0.20, [2])], 0.05, alt)
        self.assertEqual(alt, [])


class TestOznaczenieAlternatywy(unittest.TestCase):

    def test_jasno_bez_pogrubienia_z_dopiskiem(self):
        ws = openpyxl.Workbook().active
        for r in range(1, 12):
            for c in range(1, 4):
                ws.cell(r, c).value = r
        with mock.patch.object(G, "ZAPAS_KRYTERIUM", 0.05), redirect_stdout(io.StringIO()):
            G._oznacz_alternatywe(ws, (0.06, [3, 4, 5, 6, 7], 0.095), 3)
        for r in range(5, 10):
            self.assertEqual(ws.cell(r, 1).fill.fgColor.rgb[-6:], "E2EFDA")
            self.assertFalse(ws.cell(r, 1).font.bold)
        self.assertIn("Alternatywa", ws.cell(5, 5).value)
        self.assertIn("0.060", ws.cell(5, 5).value)


if __name__ == "__main__":
    unittest.main(verbosity=2)
