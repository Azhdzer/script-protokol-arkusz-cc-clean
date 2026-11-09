# -*- coding: utf-8 -*-
"""
Kazdy punkt z PZ przyrzadu dostaje zakladke i wiersz w Wynikach, takze bez pomiarow
w protokole (zlecenie 234: logger Q/LOG/22 nie zapisal punktu 0 °C). Operator domierza
go pozniej i nie kopiuje zakladek recznie. Zakladka jest pusta — bez odczytow komory
i daty z tego wsadu — a do swiadectwa punkt wchodzi dopiero po uzupelnieniu.
"""

import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_arkusze as A


def zakladka(nazwa, blok=0):
    return {"nazwa": nazwa, "block_idx": blok, "L_dane": [1] * 5, "M_dane": [2] * 5, "K4_val": "data"}


def ef(wartosc=30.0):
    return {"E_dane": [wartosc] * 5, "F_dane": [None] * 5}


def nazwy(lista):
    return [z["nazwa"] for z in lista]


class TestPunktyBezZakladki(unittest.TestCase):

    def test_brakujacy_punkt(self):
        self.assertEqual(A._punkty_bez_zakladki([(0.0, None), (10.0, None), (30.0, None)],
                                                ["30, -", "10, -"]), [(0.0, None)])

    def test_komplet(self):
        self.assertEqual(A._punkty_bez_zakladki([(30.0, None)], ["30, -"]), [])

    def test_zakladka_od_nastawy_komory_w_tolerancji(self):
        self.assertEqual(A._punkty_bez_zakladki([(20.0, None)], ["20.3, -"]), [])

    def test_blisko_ale_poza_tolerancja(self):
        self.assertEqual(A._punkty_bez_zakladki([(20.0, None)], ["25, -"]), [(20.0, None)])

    def test_punkt_temperatury_pokryty_zakladka_z_rh(self):
        self.assertEqual(A._punkty_bez_zakladki([(30.0, None)], ["30, 50"]), [])

    def test_punkt_z_rh_nie_jest_pokryty_zakladka_bez_rh(self):
        self.assertEqual(A._punkty_bez_zakladki([(30.0, 50.0)], ["30, -"]), [(30.0, 50.0)])

    def test_rozne_wilgotnosci_tej_samej_temperatury(self):
        self.assertEqual(A._punkty_bez_zakladki([(25.0, 30.0), (25.0, 50.0), (25.0, 80.0)],
                                                ["25, 30", "25, 80"]), [(25.0, 50.0)])

    def test_jedna_zakladka_pokrywa_jeden_punkt(self):
        # (20) dokladnie i (21) w tolerancji — zakladka '20, -' nalezy do punktu 20.
        self.assertEqual(A._punkty_bez_zakladki([(21.0, None), (20.0, None)], ["20, -"]), [(21.0, None)])

    def test_punkt_temperatury_nie_zabiera_zakladki_punktowi_rh(self):
        """Zlecenie 239: '20' (tylko t) bral '20, 15', a 20/15 dostawal pusta '20, 15 (2)'."""
        nazwy = ["20, 15", "20, 30", "20, 45", "40, 15", "40, 45", "5, -"]
        pz = [(5.0, None), (20.0, None), (40.0, None),
              (20.0, 15.0), (20.0, 45.0), (40.0, 15.0), (40.0, 45.0)]
        self.assertEqual(A._punkty_bez_zakladki(pz, nazwy), [])

    def test_punkt_temperatury_bez_zadnej_zakladki_tej_temperatury(self):
        self.assertEqual(A._punkty_bez_zakladki([(60.0, None), (40.0, 15.0)], ["40, 15"]),
                         [(60.0, None)])

    def test_powtorka_histerezy_w_pz(self):
        self.assertEqual(A._punkty_bez_zakladki([(23.0, 60.0), (23.0, 60.0)], ["23, 60"]), [(23.0, 60.0)])


class TestDolaczanie(unittest.TestCase):

    def test_jak_na_zrzucie_0_stopni_na_koncu(self):
        z, e, dodane = A._dolacz_punkty_z_pz(
            [zakladka("30, -"), zakladka("20, -"), zakladka("10, -")], [ef(), ef(), ef()],
            [(0.0, None), (10.0, None), (20.0, None), (30.0, None)])
        self.assertEqual(nazwy(z), ["30, -", "20, -", "10, -", "0, -"])
        self.assertEqual(dodane, ["0, -"])
        self.assertEqual(len(e), 4)
        self.assertEqual(e[3], {"E_dane": [None] * 5, "F_dane": [None] * 5})
        self.assertEqual(e[0]["E_dane"][0], 30.0)

    def test_srodkowy_punkt_na_swoim_miejscu(self):
        z, e, _ = A._dolacz_punkty_z_pz([zakladka("30, -"), zakladka("10, -")], [ef(30.0), ef(10.0)],
                                        [(10.0, None), (20.0, None), (30.0, None)])
        self.assertEqual(nazwy(z), ["30, -", "20, -", "10, -"])
        self.assertEqual([x["E_dane"][0] for x in e], [30.0, None, 10.0])

    def test_kolejnosc_rosnaca_zachowana(self):
        z, _e, _ = A._dolacz_punkty_z_pz([zakladka("10, -"), zakladka("30, -")], [ef(), ef()],
                                         [(10.0, None), (20.0, None), (30.0, None), (40.0, None)])
        self.assertEqual(nazwy(z), ["10, -", "20, -", "30, -", "40, -"])

    def test_zaburzona_kolejnosc_na_koniec(self):
        z, _e, _ = A._dolacz_punkty_z_pz(
            [zakladka("10, -"), zakladka("30, -"), zakladka("20, -")], [ef()] * 3,
            [(10.0, None), (20.0, None), (30.0, None), (0.0, None)])
        self.assertEqual(nazwy(z)[-1], "0, -")

    def test_pusta_zakladka_bez_danych_wsadu(self):
        z, _e, _ = A._dolacz_punkty_z_pz([zakladka("30, -")], [ef()], [(0.0, None)])
        pusta = z[1]
        self.assertTrue(pusta["bez_pomiarow"])
        self.assertIsNone(pusta["K4_val"])
        self.assertEqual(pusta["L_dane"], [None] * 5)
        self.assertEqual(pusta["M_dane"], [None] * 5)
        self.assertNotIn("block_idx", pusta)
        self.assertIn("0, -", A.ZAKLADKI_Z_PZ)

    def test_nie_zmienia_wejscia_ani_innych_przyrzadow(self):
        wspolne = [zakladka("30, -")]
        A._dolacz_punkty_z_pz(wspolne, [ef()], [(0.0, None)])
        self.assertEqual(nazwy(wspolne), ["30, -"])

    def test_bez_pz_bez_zmian(self):
        z, e, dodane = A._dolacz_punkty_z_pz([zakladka("30, -")], [ef()], [])
        self.assertEqual((nazwy(z), len(e), dodane), (["30, -"], 1, []))


class TestPerKopia(unittest.TestCase):

    def buduj(self, punkty_wg_serialu, flaga=True):
        bloki = [zakladka("30, -", 0), zakladka("0, -", 1)]
        dane_ef = [[ef(), ef()], [ef(), {"E_dane": [None] * 5, "F_dane": [None] * 5}]]
        s2 = [{"E": "AAA", "F": "Q/1"}, {"E": "BBB", "F": "Q/2"}]
        with mock.patch.object(A, "_PZ_PUNKTY_PRZYRZADOW_CACHE", punkty_wg_serialu), \
                mock.patch.object(A, "ZAKLADKI_WSZYSTKICH_PUNKTOW_PZ", flaga), \
                mock.patch.object(A, "USUWAJ_PUSTE_BLOKI_KOPII_S3", True):
            return A._zbuduj_dane_zakladek_per_kopia(bloki, dane_ef, s2)

    def test_tylko_przyrzad_bez_punktu_dostaje_pusta(self):
        wynik = self.buduj({"AAA": [(0.0, None), (30.0, None)], "BBB": [(0.0, None), (30.0, None)]})
        self.assertEqual([nazwy(w) for w in wynik], [["30, -", "0, -"], ["30, -", "0, -"]])
        self.assertFalse(wynik[0][1].get("bez_pomiarow"))
        self.assertTrue(wynik[1][1].get("bez_pomiarow"))

    def test_punkt_spoza_pz_nie_dostaje_zakladki(self):
        wynik = self.buduj({"BBB": [(30.0, None)]})
        self.assertEqual(nazwy(wynik[1]), ["30, -"])

    def test_po_numerze_ewidencyjnym(self):
        wynik = self.buduj({A.pz_dane.normalizuj_serial("Q/2"): [(0.0, None), (30.0, None)]})
        self.assertEqual(nazwy(wynik[1]), ["30, -", "0, -"])

    def test_flaga_wylaczona(self):
        wynik = self.buduj({"BBB": [(0.0, None), (30.0, None)]}, flaga=False)
        self.assertEqual(nazwy(wynik[1]), ["30, -"])


class TestZakladkaBezPomiarow(unittest.TestCase):

    def arkusz(self, wartosci):
        ws = mock.Mock()
        ws.range.return_value.value = wartosci
        return ws

    def test_pusta(self):
        self.assertTrue(A._zakladka_bez_pomiarow(self.arkusz([[None, "", " ", None]] * 5)))

    def test_z_odczytem(self):
        wiersze = [[None] * 4 for _ in range(5)]
        wiersze[2][0] = 0.3
        self.assertFalse(A._zakladka_bez_pomiarow(self.arkusz(wiersze)))

    def test_przetwornik_kolumny_g_h(self):
        wiersze = [[None] * 4 for _ in range(5)]
        wiersze[0][3] = 8.37
        self.assertFalse(A._zakladka_bez_pomiarow(self.arkusz(wiersze)))


if __name__ == "__main__":
    unittest.main(verbosity=2)
