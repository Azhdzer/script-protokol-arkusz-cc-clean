# -*- coding: utf-8 -*-
"""
Testy parsowania pozycji 'Obiekty wzorcowania' z PZ.

PZ przychodzi w kilku ukladach. Trzeci z nich (PZ 197) wychodzil na jaw dopiero
w praktyce: gdy na te same punkty idzie DUZO przyrzadow, etykieta 'nr fabr.:'
stoi RAZ — na koncu naglowka pozycji — a numery sa dopiero w podpunktach:

    Termohigrometr (rejestrator, 9 szt.) typ: testo 174H, nr fabr.:
      • 83623973, nr wew.: UR00045;
      • 83617608, nr wew.: UR00052;
      ...
    wytworca: Testo.

Parser szukal 'nr fabr.:' wewnatrz podpunktu, wiec wszystkie 9 przyrzadow
dostawalo PUSTY numer fabryczny, a ostatni jeszcze nr ewidencyjny sklejony
z ogonem 'UR00044; wytworca: Testo'.

Testy pracuja na fragmentach tekstu, a nie na plikach PDF — PZ zawiera dane
zleceniodawcy i nie trafia do repozytorium.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pz_dane


def parsuj(tekst):
    return pz_dane._parsuj_wpis(tekst, "197")


class TestEtykietaWNaglowku(unittest.TestCase):
    """Uklad z PZ 197 — 'nr fabr.:' raz, numery w podpunktach."""

    WPIS = ("Termohigrometr (rejestrator, 9 szt.) typ: testo 174H, nr fabr.:\n"
            "• 83623973, nr wew.: UR00045;\n"
            "• 83617608, nr wew.: UR00052;\n"
            "• 83677891, nr wew.: UR00044;\n"
            "wytwórca: Testo.")

    def setUp(self):
        self.przyrzady = parsuj(self.WPIS)

    def test_wszystkie_przyrzady_rozpoznane(self):
        self.assertEqual(len(self.przyrzady), 3)

    def test_numery_fabryczne_wypelnione(self):
        self.assertEqual([p.nr_fabr for p in self.przyrzady],
                         ["83623973", "83617608", "83677891"])

    def test_numery_ewidencyjne_wypelnione(self):
        self.assertEqual([p.nr_ewid for p in self.przyrzady],
                         ["UR00045", "UR00052", "UR00044"])

    def test_ostatni_nie_zbiera_ogona_z_wytworca(self):
        """Regresja: 'UR00044; wytworca: Testo' zamiast samego 'UR00044'."""
        ostatni = self.przyrzady[-1]
        self.assertNotIn(";", ostatni.nr_ewid)
        self.assertNotIn("wytw", ostatni.nr_ewid.lower())

    def test_typ_i_wytworca_dziedziczone_z_naglowka(self):
        for p in self.przyrzady:
            with self.subTest(fabr=p.nr_fabr):
                self.assertEqual(p.typ, "testo 174H")
                self.assertEqual(p.wytworca, "Testo")


class TestEtykietaWKazdymPodpunkcie(unittest.TestCase):
    """Uklad A — kazdy podpunkt ma wlasna etykiete 'nr fabr.:'."""

    WPIS = ("Termometr (rejestrator, 2 szt.) typ: testo 175T2, "
            "• nr fabr.: 40118669, nr ewid.: Q/LOG/36, "
            "• nr fabr.: 40118614, nr ewid.: Q/LOG/37, "
            "wytwórca: Testo.")

    def test_nadal_dziala(self):
        p = parsuj(self.WPIS)
        self.assertEqual([x.nr_fabr for x in p], ["40118669", "40118614"])
        self.assertEqual([x.nr_ewid for x in p], ["Q/LOG/36", "Q/LOG/37"])


class TestJednaLinia(unittest.TestCase):
    """Uklad B — jeden przyrzad w jednej linii."""

    def test_nadal_dziala(self):
        p = parsuj("typ: M1, nr fabr.: TMM160500502, nr ewid.: Q/LOG/19, "
                   "wytwórca: Tempmate.")
        self.assertEqual(len(p), 1)
        self.assertEqual(p[0].nr_fabr, "TMM160500502")
        self.assertEqual(p[0].nr_ewid, "Q/LOG/19")
        self.assertEqual(p[0].wytworca, "Tempmate")

    def test_wariant_z_nr_wewn(self):
        p = parsuj("typ: 174H, nr fabr.: 123456, nr wewn.: CL-1318A, wytwórca: Testo.")
        self.assertEqual(p[0].nr_fabr, "123456")
        self.assertEqual(p[0].nr_ewid, "CL-1318A")


class TestOdpornoscNaFalszywePozytywy(unittest.TestCase):
    """
    Rozpoznanie numeru „z poczatku podpunktu" wolno wlaczac WYLACZNIE wtedy,
    gdy naglowek konczy sie sama etykieta — inaczej z opisow robilyby sie
    fikcyjne numery fabryczne.
    """

    def test_bez_etykiety_w_naglowku_nie_zgaduje(self):
        wpis = ("Termohigrometr (2 szt.) typ: testo 174H, "
                "• nr ewid.: UR00045, "
                "• nr ewid.: UR00052, "
                "wytwórca: Testo.")
        p = parsuj(wpis)
        self.assertEqual([x.nr_fabr for x in p], ["", ""])
        self.assertEqual([x.nr_ewid for x in p], ["UR00045", "UR00052"])

    def test_srednik_konczy_numer_ewidencyjny(self):
        p = parsuj("typ: X, nr fabr.: 111, nr wew.: AB-1; wytwórca: Testo.")
        self.assertEqual(p[0].nr_ewid, "AB-1")


class TestPunktyMieszane(unittest.TestCase):
    """
    Sekcja 'Zakres wzorcowania' laczy oba zapisy punktow: liste samych temperatur
    i punkty z wilgotnoscia. Zgloszenie z PZ 197: w protokole zabraklo punktow
    -20, 0 i 40 °C.

    Przyczyna: gdy trafil sie choc jeden punkt z wilgotnoscia, parser konczyl
    prace i listy '(-20; 0; 40) °C' juz nie szukal. Punkty do protokolu wybiera
    sie wg PZ, wiec te trzy po prostu z niego znikaly.
    """

    FRAGMENT = ("(-20; 0; 40) °C, (25 °C, 30 %rh); (25 °C, 60 %rh); "
                "(25 °C, 85 %rh); (25 °C, 60 %rh)")

    def punkty(self, frag=None):
        return pz_dane._punkty_z_fragmentu(frag if frag is not None else self.FRAGMENT)

    def test_wszystkie_punkty_rozpoznane(self):
        self.assertEqual(len(self.punkty()), 7)

    def test_punkty_samej_temperatury_nie_gina(self):
        tylko_temp = [t for t, rh in self.punkty() if rh is None]
        self.assertEqual(tylko_temp, [-20.0, 0.0, 40.0])

    def test_punkty_z_wilgotnoscia_zachowane(self):
        z_rh = [(t, rh) for t, rh in self.punkty() if rh is not None]
        self.assertEqual(z_rh, [(25.0, 30.0), (25.0, 60.0), (25.0, 85.0), (25.0, 60.0)])

    def test_kolejnosc_jak_w_zamowieniu(self):
        """Lista temperatur stoi w PZ pierwsza — ma byc pierwsza takze u nas."""
        self.assertEqual(self.punkty()[0], (-20.0, None))
        self.assertEqual(self.punkty()[3], (25.0, 30.0))

    def test_powtorzony_punkt_histerezy_zostaje(self):
        """Drugi raz 60 %rh to osobny punkt — nie wolno go scalic."""
        self.assertEqual(sum(1 for t, rh in self.punkty() if (t, rh) == (25.0, 60.0)), 2)

    def test_sama_lista_temperatur(self):
        self.assertEqual(self.punkty("(0; 10; 20) °C"),
                         [(0.0, None), (10.0, None), (20.0, None)])

    def test_same_punkty_z_wilgotnoscia(self):
        self.assertEqual(self.punkty("(25 °C, 30 %rh); (25 °C, 60 %rh)"),
                         [(25.0, 30.0), (25.0, 60.0)])

    def test_kilka_list_temperatur(self):
        self.assertEqual(self.punkty("(0; 10) °C oraz (30; 40) °C"),
                         [(0.0, None), (10.0, None), (30.0, None), (40.0, None)])

    def test_brak_punktow(self):
        self.assertEqual(self.punkty("brak danych o punktach"), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
