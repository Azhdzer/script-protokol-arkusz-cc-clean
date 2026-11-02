# -*- coding: utf-8 -*-
"""
Blok podpisow obserwacji CC-04 (zlecenie 232, 234): przy 21 kolumnach danych trafia
dokladnie w X:Y, czyli w stare komorki szablonu. Czyszczenie po zapisie zdzieralo
z nich format i ramke — w obserwacji stalo '46328' zamiast daty.
"""

import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

import generuj_obserwacje as G

DATA = datetime.date(2026, 11, 2)
RAMKA = Border(left=Side(style="thin"), right=Side(style="thin"),
               top=Side(style="thin"), bottom=Side(style="thin"))


def szablon(format_daty="mm-dd-yy"):
    ws = openpyxl.Workbook().active
    for wiersz in (92, 93):
        for kol, nf in ((24, "General"), (25, format_daty)):
            c = ws.cell(row=wiersz, column=kol)
            c.font, c.border, c.number_format = Font(bold=True), RAMKA, nf
            c.alignment = Alignment(horizontal="center")
    ws.cell(row=92, column=24).value = "Artsiom Azhdzer"
    ws.cell(row=93, column=24).value = "Marek Szpakowski"
    ws.column_dimensions["X"].width = 20.57
    ws.column_dimensions["Y"].width = 9.57
    return ws


def przenies(ws, kol_podpisu):
    styl = G._styl_podpisow_cc04(ws)
    G._przenies_podpisy_cc04(ws, ("Artsiom Azhdzer", "Marek Szpakowski"), DATA, kol_podpisu, styl)


class TestPodpisyCC04(unittest.TestCase):

    def sprawdz_blok(self, ws, kol):
        for wiersz, nazwisko in ((92, "Artsiom Azhdzer"), (93, "Marek Szpakowski")):
            cn, cd = ws.cell(row=wiersz, column=kol), ws.cell(row=wiersz, column=kol + 1)
            self.assertEqual(cn.value, nazwisko)
            self.assertEqual(cd.value, DATA)
            self.assertTrue(cd.is_date, cd.number_format)
            self.assertEqual(cd.number_format, "mm-dd-yy")
            for c in (cn, cd):
                self.assertEqual(c.border.left.style, "thin")
                self.assertTrue(c.font.bold)
                self.assertEqual(c.alignment.horizontal, "center")
        self.assertEqual(ws.column_dimensions[get_column_letter(kol)].width, 20.57)
        self.assertEqual(ws.column_dimensions[get_column_letter(kol + 1)].width, 9.57)

    def test_kazde_polozenie_bloku(self):
        # 23-25 zachodza na stare X:Y — tam psul sie format; 19 i 28 to polozenia obok.
        for kol in (19, 23, 24, 25, 28):
            with self.subTest(kol=kol):
                ws = szablon()
                przenies(ws, kol)
                self.sprawdz_blok(ws, kol)

    def test_stare_komorki_poza_blokiem_bez_ramki(self):
        ws = szablon()
        przenies(ws, 28)
        for adres in ("X92", "Y92", "X93", "Y93"):
            self.assertIsNone(ws[adres].border.left.style, adres)
            self.assertFalse(ws[adres].font.bold, adres)

    def test_szablon_bez_formatu_daty(self):
        ws = szablon(format_daty="General")
        przenies(ws, 24)
        self.assertTrue(ws["Y92"].is_date)
        self.assertEqual(ws["Y92"].number_format, "dd.mm.yyyy")


if __name__ == "__main__":
    unittest.main(verbosity=2)
