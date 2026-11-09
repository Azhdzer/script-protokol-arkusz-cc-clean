# -*- coding: utf-8 -*-
"""
Zakladki bez pomiarow (punkty z PZ do domierzenia) nie sa przeliczane.

Zlecenie 239 (przetwornik, 19 zakladek, w tym puste '20, 15', '40, 15', '60, 15'):
krok 3 wisial >13 min na wlaczeniu trybu automatycznego — formuly pustych zakladek
(UDF-y z Wzory/Obliczenia) na pustych odczytach nie konczyly sie. Bez pustych
zakladek ten sam krok przeliczal skoroszyt w ~2 s.

Prawdziwy Excel (wlasna instancja), pliki tylko z piaskownicy.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import cc_excel
import generuj_arkusze as A
from wspolne import nowa_piaskownica

try:
    import xlwings as xw
except ImportError:          # pragma: no cover
    xw = None


def arkusz(sciezka):
    wb = openpyxl.Workbook()
    wb.active.title = "20, 45"
    wb["20, 45"]["E15"] = 20.1
    wb.create_sheet("20, 15")                       # punkt z PZ bez pomiarow
    wb.create_sheet("40, 15")["H17"] = 2.5          # przetwornik: odczyt w H
    wb.create_sheet(A.ARKUSZ_WYNIKI)
    wb.save(sciezka)


@unittest.skipIf(xw is None, "brak xlwings")
class TestPusteZakladkiBezPrzeliczania(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.f = nowa_piaskownica("puste_zakladki_przeliczanie")
        cls.sciezka = os.path.join(cls.f, "239_LA_TH_2026 - Wzór - 1 - D4850103.xlsx")
        arkusz(cls.sciezka)
        cls.app = xw.App(visible=False, add_book=False)
        cls.app.api.DisplayAlerts = False

    @classmethod
    def tearDownClass(cls):
        cc_excel.zamknij_excel_xlwings(cls.app)

    def setUp(self):
        self.wb = A._open_book_hidden(self.app, self.sciezka, update_links=False)

    def tearDown(self):
        self.wb.close()

    def wlaczone(self):
        return {s.name: bool(s.api.EnableCalculation) for s in self.wb.sheets}

    def test_otwarcie_wylacza_tylko_puste(self):
        self.assertEqual(self.wlaczone(), {"20, 45": True, "20, 15": False,
                                           "40, 15": True, A.ARKUSZ_WYNIKI: True})

    def test_po_uzupelnieniu_wraca_przeliczanie(self):
        self.wb.sheets["20, 15"].range("F16").value = 15.2
        self.assertEqual(A._przeliczanie_zakladek_wg_danych(self.wb), [])
        self.assertTrue(self.wlaczone()["20, 15"])

    def test_tryb_przeliczania_przywrocony(self):
        self.assertEqual(self.app.api.Calculation, -4105)


if __name__ == "__main__":
    unittest.main(verbosity=2)
