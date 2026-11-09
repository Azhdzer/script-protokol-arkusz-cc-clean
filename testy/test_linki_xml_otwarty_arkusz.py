# -*- coding: utf-8 -*-
"""
Przywracanie linkow serwerowych w arkuszu, ktory jest otwarty w Excelu.

Zlecenie 239: uzytkownik otworzyl arkusz, gdy krok 3 jeszcze konczyl. os.replace nie
przeszedl, linki zostaly lokalne (C:\\...\\Wzory.xls), a obok zostal plik
'... .xlsx._xltmp_'. Teraz skoroszyt jest najpierw zapisywany i zamykany, a gdy
podmiana i tak sie nie uda — plik tymczasowy jest sprzatany.
"""

import io
import os
import sys
import unittest
import zipfile
from contextlib import redirect_stdout
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_arkusze as A
from wspolne import nowa_piaskownica

RELS = ('<?xml version="1.0"?><Relationships><Relationship Id="rId1" '
        'Target="file:///C:\\Users\\x\\Wzory.xls" TargetMode="External"/></Relationships>')
SERWER = {"Wzory.xls": r"\\plum4\LabPomiarowe\Wzory.xls"}


def arkusz(sciezka):
    with zipfile.ZipFile(sciezka, "w") as z:
        z.writestr("xl/workbook.xml", "<workbook/>")
        z.writestr("xl/externalLinks/_rels/externalLink1.xml.rels", RELS)


def cel_linku(sciezka):
    with zipfile.ZipFile(sciezka) as z:
        return z.read("xl/externalLinks/_rels/externalLink1.xml.rels").decode()


class TestLinkiXmlOtwartyArkusz(unittest.TestCase):

    def setUp(self):
        self.f = nowa_piaskownica("linki_xml_otwarty")
        self.p = os.path.join(self.f, "239_LA_TH_2026 - Wzór - 1 - D4850103.xlsx")
        arkusz(self.p)

    def test_otwarty_jest_zapisywany_i_zamykany_przed_podmiana(self):
        kolejnosc = []
        with mock.patch.object(A.cc_excel, "plik_blokady", return_value="~$x"), \
                mock.patch.object(A.cc_excel, "zapisz_i_zamknij_otwarty",
                                  side_effect=lambda p: kolejnosc.append("zamknij") or True), \
                redirect_stdout(io.StringIO()):
            A._przywroc_linki_w_xml(self.p, SERWER, cicho=True)
        self.assertEqual(kolejnosc, ["zamknij"])
        self.assertIn(r"\\plum4\LabPomiarowe\Wzory.xls", cel_linku(self.p))

    def test_nieudana_podmiana_nie_zostawia_pliku_tymczasowego(self):
        bufor = io.StringIO()
        with mock.patch.object(A.os, "replace", side_effect=PermissionError(13, "zajety")), \
                redirect_stdout(bufor):
            A._przywroc_linki_w_xml(self.p, SERWER, cicho=True)
        self.assertFalse(os.path.exists(self.p + "._xltmp_"))
        self.assertIn("linki zostaja jak byly", bufor.getvalue())
        self.assertIn("C:\\Users\\x\\Wzory.xls", cel_linku(self.p))

    def test_zamkniety_bez_excela(self):
        with mock.patch.object(A.cc_excel, "zapisz_i_zamknij_otwarty") as z, \
                redirect_stdout(io.StringIO()):
            A._przywroc_linki_w_xml(self.p, SERWER, cicho=True)
        z.assert_not_called()
        self.assertIn(r"\\plum4\LabPomiarowe\Wzory.xls", cel_linku(self.p))


if __name__ == "__main__":
    unittest.main(verbosity=2)
