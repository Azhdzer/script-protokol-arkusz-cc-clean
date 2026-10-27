# -*- coding: utf-8 -*-
"""
Punkt samej temperatury zapisany BEZ nawiasu (zlecenie 232):

    Wzorcowanie w punktach pomiarowych:
    5 °C, (37 °C, 35 %rh); (37 °C, 65 %rh),
    z niepewnoscia rozszerzona (k = 2): Ut  (0,2 ÷ 0,4) °C, Urh  1 %rh

Parser znal tylko listy w nawiasach '(0; 10) °C' i pary '(25 °C, 30 %rh)', wiec
'5 °C' przepadalo — obserwacja nie brala punktu 5 °C, choc byl w PZ i w pomiarze.
Ten sam zapis w PZ 08/09/75/99/103/116 ('2 °C, (10 °C, 70 %rh) ...') i 71
('1) 37 °C, z niepewnoscia ...'). Do tego '(-15; -30 °C)' — jednostka w nawiasie (49/2025).
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pz_dane as P


class TestPojedynczaTemperatura(unittest.TestCase):

    def test_232(self):
        self.assertEqual(P._punkty_z_fragmentu(
            "5 °C, (37 °C, 35 %rh); (37 °C, 65 %rh),\n"
            "z niepewnością rozszerzoną (k = 2): Ut  (0,2 ÷ 0,4) °C, Urh  1 %rh"),
            [(5.0, None), (37.0, 35.0), (37.0, 65.0)])

    def test_pozycja_tylko_jedna_temperatura(self):
        self.assertEqual(P._punkty_z_fragmentu(
            "1) 37 °C, z niepewnością rozszerzoną (k = 2): Ut  (0,2 ÷ 0,4) °C."), [(37.0, None)])

    def test_niepewnosc_to_nie_punkt(self):
        for frag in ("(25 °C, 30 %rh), z niepewnością rozszerzoną (k = 2): Ut ≥ 0,2 °C",
                     "(10 °C, 60 %rh) U ≥ 0,015°C",
                     "(-40; -10) °C with expanded uncertainty U = 0,1 °C"):
            with self.subTest(frag=frag):
                self.assertNotIn(0.2, [t for t, _rh in P._punkty_z_fragmentu(frag)])
                self.assertNotIn(0.015, [t for t, _rh in P._punkty_z_fragmentu(frag)])
                self.assertNotIn(0.1, [t for t, _rh in P._punkty_z_fragmentu(frag)])

    def test_pary_i_listy_bez_zmian(self):
        self.assertEqual(P._punkty_z_fragmentu("(-20; 0; 40) °C, (25 °C, 30 %rh)"),
                         [(-20.0, None), (0.0, None), (40.0, None), (25.0, 30.0)])
        self.assertEqual(P._punkty_z_fragmentu(
            "- temperature: 10 °C; relative humidity: 22 %rh, 45 %rh"), [(10.0, 22.0), (10.0, 45.0)])

    def test_jednostka_w_nawiasie(self):
        self.assertEqual(P._punkty_z_fragmentu(
            "1) (-15; -30 °C), z niepewnością rozszerzoną (k = 2): Ut  (0,2 ÷ 0,4) °C."),
            [(-15.0, None), (-30.0, None)])


if __name__ == "__main__":
    unittest.main(verbosity=2)
