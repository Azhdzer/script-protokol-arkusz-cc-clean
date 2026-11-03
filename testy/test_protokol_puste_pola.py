# -*- coding: utf-8 -*-
"""
Puste pola protokolu (zlecenie 232):
  * Strona 2, 'Uwagi' bez tresci -> '-' na szarym tle,
  * PZ: nr ewidencyjne w nawiasach za numerami fabrycznymi
    'nr fabr.: 36668832 (H1), 36679283 (H2)' -> nr ewid. H1, H2 (dotad puste).
Strona 3 (czujniki bez danych, dl. rurki i forma kondensacji) — sprawdzane na pelnym
przebiegu kroku 2 (test_obieg) i recznie na zleceniu 232.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import PatternFill

import generuj_obserwacje as G
import pz_dane as P

SZARY = PatternFill(fill_type="solid", fgColor="BFBFBF")


class TestUwagi(unittest.TestCase):

    def test_puste_dostaja_kreske_na_szarym(self):
        ws = openpyxl.Workbook().active
        w0 = G.WIERSZ_1_PRZYRZADU_S2
        ws.cell(w0, 16).value = "Uz = (24,0 ± 0,1) V DC."
        ws.cell(w0 + 1, 16).value = "  "
        G._uwagi_bez_tresci(ws, 3, SZARY)
        self.assertEqual(ws.cell(w0, 16).value, "Uz = (24,0 ± 0,1) V DC.")
        self.assertIsNone(ws.cell(w0, 16).fill.fill_type)
        for w in (w0 + 1, w0 + 2):
            self.assertEqual(ws.cell(w, 16).value, "-")
            self.assertEqual(ws.cell(w, 16).fill.fgColor.rgb[-6:], "BFBFBF")
        self.assertIsNone(ws.cell(w0 + 3, 16).value)        # poza przyrzadami — bez zmian


def pz(lista):
    return ("Obiekty wzorcowania:\n"
            f"Termohigrometr (rejestrator, 3 szt.) typ: testo 174H, nr fabr.: {lista}, "
            "wytwórca: Testo.\nMetoda wzorcowania: \nMetoda porównawcza w komorze klimatycznej.\n")


class TestEwidencyjneWNawiasach(unittest.TestCase):

    def test_lista_z_nawiasami(self):
        lista = P.parsuj_tekst(pz("36668832 (H1), 36679283 (H2), 36668513 \n(H10)"))
        self.assertEqual([(p.nr_fabr, p.nr_ewid) for p in lista],
                         [("36668832", "H1"), ("36679283", "H2"), ("36668513", "H10")])

    def test_dopisek_ze_spacjami_to_nie_ewid(self):
        lista = P.parsuj_tekst(pz("21201805 (kanał pomiarowy nr: 1)"))
        self.assertEqual([(p.nr_fabr, p.nr_ewid) for p in lista], [("21201805", "")])

    def test_slowo_bez_cyfry_to_nie_ewid(self):
        lista = P.parsuj_tekst(pz("21201805 (rejestrator)"))
        self.assertEqual(lista[0].nr_ewid, "")

    def test_etykieta_nr_ewid_ma_pierwszenstwo(self):
        tekst = ("Obiekty wzorcowania:\nTermohigrometr typ: testo 174H, nr fabr.: 36668832 (H1), "
                 "nr ewid.: Q/LOG/5, wytwórca: Testo.\nMetoda wzorcowania: \nx\n")
        self.assertEqual(P.parsuj_tekst(tekst)[0].nr_ewid, "Q/LOG/5")

    def test_bez_nawiasow_bez_zmian(self):
        lista = P.parsuj_tekst(pz("36668832, 36679283"))
        self.assertEqual([p.nr_ewid for p in lista], ["", ""])


if __name__ == "__main__":
    unittest.main(verbosity=2)
