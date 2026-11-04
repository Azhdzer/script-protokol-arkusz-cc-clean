# -*- coding: utf-8 -*-
"""
GEN_NR_SW_WG_EWID — numery swiadectw w kolejnosci nr EWIDENCYJNEGO.

Klienci chca, zeby swiadectwo nr 1047 dostal przyrzad U90337-18, 1048 — U90337-36
itd. (zlecenie 233), a nie przyrzad o najmniejszym nr fabrycznym (kolejnosc kopii).
Domyslnie (wylaczone) numeracja bez zmian.
"""

import io
import os
import re
import sys
import unittest
from contextlib import redirect_stdout
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_arkusze as A
from wspolne import nowa_piaskownica

try:
    from docx import Document
except ImportError:          # pragma: no cover
    Document = None


class _PZ:
    def __init__(self, nr_ewid):
        self.nr_ewid, self.uzytkownik, self.zleceniodawca = nr_ewid, "", ""
        self.wytworca, self.typ = "", ""


class TestNumerySwiadectw(unittest.TestCase):

    def test_domyslnie_kolejnosc_kopii(self):
        self.assertEqual(A.numery_swiadectw(["U90337-36", "U90337-18"], 1047), [1047, 1048])

    def test_wg_ewid_naturalnie(self):
        ewid = ["U90337-36", "U90337-9", "U90337-100", "U90337-18"]
        self.assertEqual(A.numery_swiadectw(ewid, 1047, wg_ewid=True), [1049, 1047, 1050, 1048])

    def test_bez_ewid_na_koncu(self):
        self.assertEqual(A.numery_swiadectw(["", "B-2", None, "B-1"], 10, wg_ewid=True),
                         [12, 11, 13, 10])

    def test_wylaczone_wg_nr_fabrycznego(self):
        fabr = ["37192", "9", "37191", "10"]
        self.assertEqual(A.numery_swiadectw(["U-1", "U-2", "U-3", "U-4"], 1047, False, fabr),
                         [1050, 1047, 1049, 1048])

    def test_wg_ewid_bez_ewid_wg_fabrycznego(self):
        self.assertEqual(A.numery_swiadectw(["", "B-1", ""], 10, True, ["Z2", "Z9", "Z1"]),
                         [12, 10, 11])

    def test_ewid_z_pz_gdy_strona2_pusta(self):
        with mock.patch.object(A, "_pz_mapa_arkusze", return_value={"371": _PZ("U-5")}):
            self.assertEqual(A._nr_ewid_kopii({"F": None}, "233_LA_TH_2026 - testo 174T - 371.xlsx"),
                             "U-5")
            self.assertEqual(A._nr_ewid_kopii({"F": "U-7"}, "233_LA_TH_2026 - testo 174T - 371.xlsx"),
                             "U-7")


@unittest.skipIf(Document is None, "brak python-docx")
class TestSwiadectwaWgEwid(unittest.TestCase):

    def test_word_dostaje_numer_wg_ewid(self):
        folder = nowa_piaskownica("nr_sw_wg_ewid")
        doc = Document()
        doc.add_paragraph("Swiadectwo [nr_sw] dla [nr_ewid]")
        doc.save(os.path.join(folder, "W.docx"))
        # kopie w kolejnosci nr fabrycznego, nr ewidencyjne "pomieszane"
        kopie = [f"233_LA_TH_2026 - testo 174T - 3719{i}.xlsx" for i in (1, 2, 3)]
        s2 = [{"F": "U90337-41"}, {"F": "U90337-18"}, {"F": "U90337-36"}]
        with mock.patch.object(A, "_pz_mapa_arkusze", return_value={}), \
                redirect_stdout(io.StringIO()):
            nazwy = A.utworz_kopie_word(folder, "W.docx", "W.docx", "W.docx", s2, kopie,
                                        [], [[], [], []], 1047, wg_ewid=True)
        wynik = {}
        for nazwa in nazwy:
            tekst = " ".join(p.text for p in Document(os.path.join(folder, nazwa)).paragraphs)
            wynik[re.match(r"(\d+)_", nazwa).group(1)] = tekst
        self.assertEqual(wynik, {"1047": "Swiadectwo 1047 dla U90337-18",
                                 "1048": "Swiadectwo 1048 dla U90337-36",
                                 "1049": "Swiadectwo 1049 dla U90337-41"})


@unittest.skipIf(Document is None, "brak python-docx")
class TestBezNrEwid(unittest.TestCase):
    """Brak nr ewid. w PZ -> w opisie obiektu nie ma ', nr ewid.: -'."""

    def swiadectwo(self, ewid):
        folder = nowa_piaskownica("bez_nr_ewid")
        doc = Document()
        p = doc.add_paragraph()
        # podzial na runy jak w prawdziwym szablonie 'zakres.docx'
        for t in ("Termohigrometr (rejestrator) typ: ", "[typ]", ", nr fabr.: ", "[nr_fabr]",
                  ",", " nr ewid.:", " [nr_ewid],", " ", "wytwórca: ", "[wytworca]", "."):
            p.add_run(t)
        doc.save(os.path.join(folder, "W.docx"))
        with mock.patch.object(A, "_pz_mapa_arkusze", return_value={}), \
                redirect_stdout(io.StringIO()):
            nazwy = A.utworz_kopie_word(folder, "W.docx", "W.docx", "W.docx",
                                        [{"F": ewid, "B": "Lascar", "D": "EL-USB-2-LCD"}],
                                        ["240_LA_TH_2026 - Wzór - 1 - 000007263.xlsx"], [], [[]], 1)
        return Document(os.path.join(folder, nazwy[0])).paragraphs[0].text

    def test_kreska_usuwa_fragment(self):
        self.assertEqual(self.swiadectwo("-"),
                         "Termohigrometr (rejestrator) typ: EL-USB-2-LCD, nr fabr.: 000007263, "
                         "wytwórca: Lascar.")

    def test_pusty_usuwa_fragment(self):
        self.assertNotIn("ewid", self.swiadectwo(""))

    def test_z_numerem_bez_zmian(self):
        self.assertIn(", nr ewid.: U90337-18, wytwórca", self.swiadectwo("U90337-18"))


@unittest.skipIf(Document is None, "brak python-docx")
class TestKopiaPoprzedniegoSwiadectwa(unittest.TestCase):

    def test_kopia_ma_poprzednia_tresc_a_nie_szablon(self):
        """Kopia robiona byla PO nadpisaniu szablonem — zachowywala pusty szablon."""
        folder = nowa_piaskownica("kopia_swiadectwa")
        szablon = Document()
        szablon.add_paragraph("Swiadectwo [nr_sw]")
        szablon.save(os.path.join(folder, "W.docx"))
        stare = Document()
        stare.add_paragraph("POPRZEDNIE SWIADECTWO")
        stare.save(os.path.join(folder, "1047_233_LA_TH_2026.docx"))
        with mock.patch.object(A, "_pz_mapa_arkusze", return_value={}), \
                redirect_stdout(io.StringIO()):
            A.utworz_kopie_word(folder, "W.docx", "W.docx", "W.docx", [{"F": "X"}],
                                ["233_LA_TH_2026 - testo - 1"], [], [[]], 1047)
        kopie = os.listdir(os.path.join(folder, "kopie_zapasowe"))
        self.assertEqual(len(kopie), 1)
        tekst = " ".join(p.text for p in Document(
            os.path.join(folder, "kopie_zapasowe", kopie[0])).paragraphs)
        self.assertIn("POPRZEDNIE SWIADECTWO", tekst)


if __name__ == "__main__":
    unittest.main(verbosity=2)
