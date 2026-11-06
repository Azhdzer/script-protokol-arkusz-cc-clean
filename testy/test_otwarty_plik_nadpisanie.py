# -*- coding: utf-8 -*-
"""
Krok 2 nadpisuje protokol / obserwacje otwarte w Excelu uzytkownika.

Zgloszenie: protokol z poprzedniego przebiegu byl otwarty, krok 2 przeliczal caly
pomiar i dopiero na koncu padal PermissionError (WinError 32). Teraz ten JEDEN
skoroszyt jest zapisywany (zmiany uzytkownika trafiaja do kopii zapasowej)
i zamykany; Excel i jego pozostale skoroszyty zostaja.

Testy z prawdziwym Excelem uruchamiaja WLASNA instancje (udaje Excela uzytkownika)
i dzialaja tylko na plikach z piaskownicy.
"""

import io
import os
import sys
import time
import unittest
from contextlib import redirect_stdout
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import cc_excel
import generuj_obserwacje as G
from wspolne import nowa_piaskownica

try:
    import win32com.client  # noqa: F401
    _COM = True
except ImportError:          # pragma: no cover
    _COM = False


def plik(sciezka, a1="stare"):
    wb = openpyxl.Workbook()
    wb.active["A1"] = a1
    wb.save(sciezka)


class TestPlikBlokady(unittest.TestCase):

    def test_excel_i_word(self):
        f = nowa_piaskownica("blokada")
        p = os.path.join(f, "233_LA_TH_2026 - protokół CC-04.xlsx")
        self.assertIsNone(cc_excel.plik_blokady(p))
        open(os.path.join(f, "~$233_LA_TH_2026 - protokół CC-04.xlsx"), "w").close()
        self.assertTrue(cc_excel.plik_blokady(p))
        d = os.path.join(f, "1047_233_LA_TH_2026.docx")
        open(os.path.join(f, "~$47_233_LA_TH_2026.docx"), "w").close()
        self.assertTrue(cc_excel.plik_blokady(d))


class TestNadpiszSzablonem(unittest.TestCase):

    def setUp(self):
        self.f = nowa_piaskownica("nadpisz_szablonem")
        self.szablon, self.cel = os.path.join(self.f, "s.xlsx"), os.path.join(self.f, "c.xlsx")
        plik(self.szablon, "szablon")
        plik(self.cel, "stary protokol")

    def test_zablokowany_jest_zwalniany_przed_kopia(self):
        kolejnosc = []
        with mock.patch.object(cc_excel, "plik_blokady", return_value="~$c.xlsx"), \
                mock.patch.object(cc_excel, "zapisz_i_zamknij_otwarty",
                                  side_effect=lambda p: kolejnosc.append("zamknij") or True), \
                mock.patch.object(G.cc_kopia, "kopia_przed_nadpisaniem",
                                  side_effect=lambda p: kolejnosc.append("kopia")), \
                redirect_stdout(io.StringIO()):
            G._nadpisz_szablonem(self.szablon, self.cel)
        self.assertEqual(kolejnosc, ["zamknij", "kopia"])
        self.assertEqual(openpyxl.load_workbook(self.cel).active["A1"].value, "szablon")

    def test_nie_da_sie_nadpisac_czytelny_blad(self):
        bufor = io.StringIO()
        with mock.patch.object(G.shutil, "copy2", side_effect=PermissionError(32, "uzywany")), \
                redirect_stdout(bufor), self.assertRaises(SystemExit):
            G._nadpisz_szablonem(self.szablon, self.cel)
        self.assertIn("Zamknij go", bufor.getvalue())


@unittest.skipUnless(_COM, "brak pywin32")
class TestZapiszIZamknijOtwarty(unittest.TestCase):
    """Prawdziwy Excel: niezapisana zmiana uzytkownika trafia do pliku, reszta zostaje."""

    def test_zapisuje_zamyka_tylko_ten_skoroszyt(self):
        f = nowa_piaskownica("zapisz_zamknij_otwarty")
        cel, inny = os.path.join(f, "protokol.xlsx"), os.path.join(f, "inny.xlsx")
        plik(cel)
        plik(inny)
        xl = cc_excel.nowy_excel_com()
        pid = cc_excel.pid_excela_com(xl)
        try:
            wb = xl.Workbooks.Open(cel)
            xl.Workbooks.Open(inny)
            wb.Worksheets(1).Range("A1").Value = "poprawka uzytkownika"
            time.sleep(0.5)
            self.assertTrue(cc_excel.plik_blokady(cel))
            with redirect_stdout(io.StringIO()):
                self.assertTrue(cc_excel.zapisz_i_zamknij_otwarty(cel))
            self.assertIsNone(cc_excel.plik_blokady(cel))
            self.assertEqual(openpyxl.load_workbook(cel).active["A1"].value, "poprawka uzytkownika")
            self.assertEqual(xl.Workbooks.Count, 1)
            self.assertTrue(cc_excel.proces_zyje(pid))
        finally:
            with redirect_stdout(io.StringIO()):
                cc_excel.zamknij_excel_com(xl)

    def test_nieotwarty_plik_bez_excela(self):
        f = nowa_piaskownica("zapisz_zamknij_nieotwarty")
        cel = os.path.join(f, "x.xlsx")
        plik(cel)
        self.assertTrue(cc_excel.zapisz_i_zamknij_otwarty(cel))


if __name__ == "__main__":
    unittest.main(verbosity=2)
