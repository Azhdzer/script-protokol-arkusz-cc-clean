# -*- coding: utf-8 -*-
"""
Testy odsiewania przyrzadow bez pomiarow (generuj_arkusze).

Kontekst: lista kopii do zrobienia pochodzi ze Strony 2 (tabela przyrzadow),
a nie z kolorow na Stronie 3. Zeby wygenerowac dokumenty dla JEDNEGO przyrzadu,
uzytkownik wyszarza pomiary pozostalych — i oczekuje jednej kopii i jednego
swiadectwa. Wczesniej powstawaly komplet pustych kopii (sam arkusz Wyniki)
i bezuzyteczne swiadectwa z zerowa tabela kalibracji.

Testy sa czysto obliczeniowe — nie uruchamiaja Excela.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_arkusze as A


def blok(*wartosci):
    """Blok pomiarowy Strony 3 z podanymi wartosciami w kolumnie E."""
    return {"E_dane": list(wartosci), "F_dane": []}


PUSTY = {"E_dane": [None, None], "F_dane": [None, None]}


def przyrzad(nr_fabryczny):
    return {"O": "termohigrometr", "E": nr_fabryczny}


class TestOdfiltrowaniePrzyrzadow(unittest.TestCase):

    def setUp(self):
        self.s2 = [przyrzad("37025098"), przyrzad("37025105"),
                   przyrzad("37025156"), przyrzad("37025720")]
        self.f24 = [1.1, 2.2, 3.3, 4.4]

    def test_wyszarzenie_trzech_zostawia_jeden(self):
        """Sytuacja z realnego uzycia: aktywny tylko czwarty przyrzad."""
        ef = [[PUSTY, PUSTY], [PUSTY, PUSTY], [PUSTY, PUSTY], [blok(38.0), PUSTY]]
        s2, _ef, _f24, pominiete = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual(len(s2), 1)
        self.assertEqual(len(pominiete), 3)
        self.assertEqual(s2[0]["E"], "37025720")

    def test_zachowany_numer_z_protokolu(self):
        """
        Nazwa kopii ma dalej zgadzac sie z pozycja na Stronie 2 — czwarty
        przyrzad zostaje czwartym, a nie staje sie pierwszym.
        """
        ef = [[PUSTY], [PUSTY], [PUSTY], [blok(38.0)]]
        s2, _ef, _f24, _pom = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual(s2[0]["_nr_przyrzadu"], 4)

    def test_pominiete_niosa_swoje_numery(self):
        ef = [[PUSTY], [blok(25.0)], [PUSTY], [blok(38.0)]]
        _s2, _ef, _f24, pominiete = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual([nr for nr, _r in pominiete], [1, 3])

    def test_dane_ef_i_f24_ida_razem_z_przyrzadem(self):
        """Po filtracji indeksy musza sie zgadzac, inaczej dane sie rozjada."""
        ef = [[PUSTY], [blok(25.0)], [PUSTY], [blok(38.0)]]
        _s2, ef_po, f24_po, _pom = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual(ef_po, [[blok(25.0)], [blok(38.0)]])
        self.assertEqual(f24_po, [2.2, 4.4])

    def test_wszystkie_z_danymi_nic_nie_gubi(self):
        ef = [[blok(1.0)], [blok(2.0)], [blok(3.0)], [blok(4.0)]]
        s2, _ef, f24_po, pominiete = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual(len(s2), 4)
        self.assertEqual(pominiete, [])
        self.assertEqual(f24_po, self.f24)

    def test_wszystkie_wyszarzone_daja_pusta_liste(self):
        ef = [[PUSTY], [PUSTY], [PUSTY], [PUSTY]]
        s2, _ef, _f24, pominiete = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual(s2, [])
        self.assertEqual(len(pominiete), 4)

    def test_brak_wpisu_ef_traktowany_jak_brak_danych(self):
        """Krotsza lista dane_ef niz Strona 2 nie moze wywalic skryptu."""
        ef = [[blok(1.0)]]
        s2, _ef, _f24, pominiete = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual(len(s2), 1)
        self.assertEqual(len(pominiete), 3)

    def test_wartosc_zero_to_wciaz_pomiar(self):
        """0.0 jest poprawnym odczytem — nie wolno go uznac za pusty blok."""
        ef = [[blok(0.0)], [PUSTY], [PUSTY], [PUSTY]]
        s2, _ef, _f24, _pom = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual(len(s2), 1)
        self.assertEqual(s2[0]["_nr_przyrzadu"], 1)

    def test_dane_tylko_w_kolumnie_f_licza_sie(self):
        ef = [[{"E_dane": [], "F_dane": [45.0]}], [PUSTY], [PUSTY], [PUSTY]]
        s2, _ef, _f24, _pom = A._odfiltruj_przyrzady_bez_danych(
            self.s2, ef, self.f24)
        self.assertEqual(len(s2), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
