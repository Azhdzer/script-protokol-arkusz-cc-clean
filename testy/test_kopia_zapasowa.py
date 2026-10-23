# -*- coding: utf-8 -*-
"""
Kopia zapasowa przed nadpisaniem (cc_kopia).

Zgloszenie (zlecenie 230): ponowne uruchomienie kroku 2 utworzylo protokol od
nowa i przepadlo reczne wypelnienie przetwornikow na Stronach 2-3. Kazdy plik,
ktory krok 2/3/4 nadpisuje, trafia najpierw do 'kopie_zapasowe'.
"""

import os
import sys
import time
import unittest
from contextlib import redirect_stdout
import io

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import cc_kopia
from wspolne import KORZEN, nowa_piaskownica


def plik(folder, nazwa, tresc):
    p = os.path.join(folder, nazwa)
    with open(p, "w", encoding="utf-8") as f:
        f.write(tresc)
    return p


class TestKopia(unittest.TestCase):

    def setUp(self):
        self.folder = nowa_piaskownica("kopia_zapasowa")

    def kopia(self, sciezka, **kw):
        with redirect_stdout(io.StringIO()):
            return cc_kopia.kopia_przed_nadpisaniem(sciezka, **kw)

    def test_brak_pliku_nic_nie_robi(self):
        self.assertIsNone(self.kopia(os.path.join(self.folder, "brak.xlsx")))

    def test_kopia_z_trescia_i_data_w_nazwie(self):
        p = plik(self.folder, "230_LA_TH_2026 - protokół CC.xlsx", "recznie wypelniony")
        k = self.kopia(p)
        self.assertEqual(os.path.dirname(k), os.path.join(self.folder, cc_kopia.FOLDER_KOPII))
        self.assertRegex(os.path.basename(k),
                         r"^230_LA_TH_2026 - protokół CC \(\d{4}-\d\d-\d\d \d\d\.\d\d\.\d\d\)\.xlsx$")
        with open(k, encoding="utf-8") as f:
            self.assertEqual(f.read(), "recznie wypelniony")

    def test_zachowuje_tylko_najnowsze(self):
        p = plik(self.folder, "a.xlsx", "x")
        for i in range(4):
            t = time.time() - 1000 + i * 10
            os.utime(p, (t, t))
            self.kopia(p, zachowaj=2)
        kopie = os.listdir(os.path.join(self.folder, cc_kopia.FOLDER_KOPII))
        self.assertEqual(len(kopie), 2)


class TestKrokiRobiaKopie(unittest.TestCase):
    """Straznik w zrodlach: kazde nadpisanie poprzedzone kopia."""

    def zrodlo(self, nazwa):
        with open(os.path.join(KORZEN, nazwa), encoding="utf-8") as f:
            return f.read()

    def test_krok2_protokol_i_obserwacja(self):
        s = self.zrodlo("generuj_obserwacje.py")
        self.assertRegex(s, r"cc_kopia\.kopia_przed_nadpisaniem\(out_path\)[^\n]*\n"
                            r"\s*shutil\.copy2\(tmpl_path, out_path\)")
        self.assertEqual(s.count("cc_kopia.kopia_przed_nadpisaniem(output_path)"), 2)

    def test_krok3_kopie_i_swiadectwa(self):
        s = self.zrodlo("generuj_arkusze.py")
        self.assertIn("cc_kopia.kopia_przed_nadpisaniem(sciezka_kopii)", s)
        self.assertIn("cc_kopia.kopia_przed_nadpisaniem(sciezka_docx)", s)

    def test_krok4(self):
        self.assertIn("cc_kopia.kopia_przed_nadpisaniem(sciezka)", self.zrodlo("wypelnij_warunki.py"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
