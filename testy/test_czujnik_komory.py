# -*- coding: utf-8 -*-
"""
Testy czujnika wzorcowego komory CC.

Zgloszenie: po wymianie czujnika w komorze CC (Pt100-11 -> Pt100-31) kopie
arkusza obliczeniowego nadal pokazywaly stary numer. Przyczyna: numer siedzial
WYLACZNIE w szablonie arkusza (komorka K11 kazdej zakladki), a krok 3 tej
komorki dla komory CC w ogole nie dotykal — ustawial ja tylko dla CC-04, gdzie
czujnik wybierany jest wg naroznika. Nie bylo wiec miejsca, w ktorym mozna
byloby ten numer zmienic.

Teraz steruje nim panel (GEN_CC_CZUJNIK) i to on jest zrodlem prawdy — komorki
kopii sa nadpisywane, tak samo jak od dawna dzieje sie to dla CC-04. Wartosci
domyslne rejestru sa rowne tym z szablonu, wiec dopoki nikt nic nie zmieni,
kopie wygladaja dokladnie jak dotad. Pilnuje tego TestZgodnoscZSzablonem.

Pelne sprawdzenie „numer z panelu jest w gotowej kopii" robi test_obieg.py
(test_3_czujnik_wzorcowy_z_panelu_trafia_do_kopii) — wymaga Excela.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import cc_config as C
from wspolne import KORZEN, SZABLON_ARKUSZA, wartosci_z_modulu

KOMORKI = ("K11", "K12", "K13", "K17")


class TestRejestr(unittest.TestCase):
    """Ustawienia komory CC musza byc widoczne i opisane w panelu."""

    def ust(self, env):
        return C.WG_ENV[env]

    def test_wszystkie_cztery_istnieja(self):
        for env in ("GEN_CC_CZUJNIK", "GEN_CC_PRZYRZAD",
                    "GEN_CC_WEJSCIE", "GEN_CC_KOMORA"):
            with self.subTest(env=env):
                self.assertIn(env, C.WG_ENV)

    def test_naleza_do_kroku_arkusze(self):
        for env in ("GEN_CC_CZUJNIK", "GEN_MAP_CC04"):
            with self.subTest(env=env):
                self.assertEqual(self.ust(env).krok, "ark")

    def test_czujniki_obu_komor_w_jednej_grupie(self):
        """CC i CC-04 maja stac obok siebie — szuka sie ich razem."""
        self.assertEqual(self.ust("GEN_CC_CZUJNIK").grupa,
                         self.ust("GEN_MAP_CC04").grupa)

    def test_numery_czujnikow_sa_na_poziomie_podstawowym(self):
        """Wymiana czujnika to zwykla praca w laboratorium, nie ustawienie eksperckie."""
        for env in ("GEN_CC_CZUJNIK", "GEN_MAP_CC04"):
            with self.subTest(env=env):
                self.assertEqual(self.ust(env).poziom, C.PODSTAWOWY)

    def test_opis_uprzedza_ze_szablon_jest_nadpisywany(self):
        self.assertIn("szablon", self.ust("GEN_CC_CZUJNIK").opis.lower())


class TestZgodnoscZSzablonem(unittest.TestCase):
    """
    Domyslne wartosci panelu musza odpowiadac temu, co stoi w szablonie arkusza
    obliczeniowego. Gdyby sie rozjechaly, samo wlaczenie zapisu zmienialoby
    dokumenty bez wiedzy uzytkownika.
    """

    @classmethod
    def setUpClass(cls):
        sciezka = os.path.join(KORZEN, SZABLON_ARKUSZA)
        if not os.path.exists(sciezka):
            raise unittest.SkipTest(f"brak szablonu: {SZABLON_ARKUSZA}")
        wb = openpyxl.load_workbook(sciezka, read_only=True)
        try:
            zakladki = [n for n in wb.sheetnames if n != "Wyniki"]
            cls.z_szablonu = {k: wb[zakladki[0]][k].value for k in KOMORKI}
            cls.wszystkie = [{k: wb[n][k].value for k in KOMORKI} for n in zakladki]
        finally:
            wb.close()

    def test_szablon_ma_ten_sam_komplet_w_kazdej_zakladce(self):
        for i, komplet in enumerate(self.wszystkie):
            with self.subTest(zakladka=i):
                self.assertEqual(komplet, self.z_szablonu)

    def test_domyslny_czujnik_zgodny_z_szablonem(self):
        self.assertEqual(C.WG_ENV["GEN_CC_CZUJNIK"].domyslna, self.z_szablonu["K11"])

    def test_domyslny_multimetr_zgodny_z_szablonem(self):
        self.assertEqual(C.WG_ENV["GEN_CC_PRZYRZAD"].domyslna, self.z_szablonu["K12"])

    def test_domyslne_wejscie_i_oznaczenie_zgodne_z_szablonem(self):
        self.assertEqual(C.WG_ENV["GEN_CC_WEJSCIE"].domyslna, self.z_szablonu["K13"])
        self.assertEqual(C.WG_ENV["GEN_CC_KOMORA"].domyslna, self.z_szablonu["K17"])


class TestOdczytWSkrypcie(unittest.TestCase):
    """Krok 3 czyta ustawienia do PARAMETRY_CC — stad ida do komorek kopii."""

    def parametry(self, env=None):
        wyrazenia = {k: f"m.PARAMETRY_CC[{k!r}]" for k in KOMORKI}
        odczyt = wartosci_z_modulu("generuj_arkusze", list(wyrazenia.values()), env)
        return {k: odczyt[w] for k, w in wyrazenia.items()}

    def test_domyslnie_wartosci_z_rejestru(self):
        self.assertEqual(self.parametry()["K11"],
                         C.WG_ENV["GEN_CC_CZUJNIK"].domyslna)

    def test_zmiana_w_panelu_dociera_do_skryptu(self):
        p = self.parametry({"GEN_CC_CZUJNIK": "Pt100-31", "GEN_CC_PRZYRZAD": "K2002",
                            "GEN_CC_WEJSCIE": "104", "GEN_CC_KOMORA": "CC-02"})
        self.assertEqual(p, {"K11": "Pt100-31", "K12": "K2002",
                             "K13": "104", "K17": "CC-02"})

    def test_puste_pole_wraca_do_wartosci_domyslnej(self):
        """
        Pusto = wartosc domyslna, jak w calym panelu (cc_config.tekst). Czyli
        komorka NIGDY nie zostaje pusta — wyczyszczone pole nie skasuje numeru
        czujnika w gotowej kopii.
        """
        self.assertEqual(self.parametry({"GEN_CC_CZUJNIK": ""})["K11"],
                         C.WG_ENV["GEN_CC_CZUJNIK"].domyslna)

    def test_zmiana_jednego_pola_nie_rusza_pozostalych(self):
        p = self.parametry({"GEN_CC_CZUJNIK": "Pt100-31"})
        self.assertEqual(p["K12"], C.WG_ENV["GEN_CC_PRZYRZAD"].domyslna)
        self.assertEqual(p["K17"], C.WG_ENV["GEN_CC_KOMORA"].domyslna)


if __name__ == "__main__":
    unittest.main(verbosity=2)
