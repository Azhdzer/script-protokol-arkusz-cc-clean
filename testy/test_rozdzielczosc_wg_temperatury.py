# -*- coding: utf-8 -*-
"""
Testy rozdzielczosci temperatury zaleznej od punktu.

Rzadki przypadek (zlecenie 195, Onset UX100-001): przyrzad ma INNA rozdzielczosc
w roznych temperaturach. Na Stronie 2 protokolu w kolumnie 'Rozdzielczosc
odczytu t' stoi wtedy tekst:

    dla 70°C t: 0,056 °C
    dla 20°C t: 0,024 °C
    dla -10°C t: 0,034 °C

Krok 3 wpisywal ten tekst wprost do H57 kazdej zakladki kopii — a H57 musi byc
liczba, bo od niej liczy sie budzet niepewnosci. Teraz kazda zakladka dostaje
wartosc dla SWOJEJ temperatury (z nazwy zakladki '70, -' / '-10, -').

Zapisy z Zestawienia sa rozne — temperatura przed wartoscia, za nia, zakresy
'(10 ÷ 25) °C' i '(4-5) °C' — dlatego sprawdzamy kazdy z nich.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_arkusze as A
import pz_dane as P

ZGLOSZENIE = "dla 70°C t: 0,056 °C\ndla 20°C t: 0,024 °C\ndla -10°C t: 0,034 °C"


def wartosc(tekst, temp):
    return P.rozdzielczosc_dla_temperatury(tekst, temp)[0]


class TestPrzypadekZeZgloszenia(unittest.TestCase):

    def test_kazda_temperatura_ma_swoja_wartosc(self):
        self.assertEqual(wartosc(ZGLOSZENIE, 70), 0.056)
        self.assertEqual(wartosc(ZGLOSZENIE, 20), 0.024)
        self.assertEqual(wartosc(ZGLOSZENIE, -10), 0.034)

    def test_nastawa_komory_obok_okraglej_temperatury(self):
        """Zakladki maja nastawy (70,3 / -10,3), wpisy — okragle wartosci."""
        self.assertEqual(wartosc(ZGLOSZENIE, 70.3), 0.056)
        self.assertEqual(wartosc(ZGLOSZENIE, -10.3), 0.034)

    def test_temperatura_spoza_listy_nie_zgaduje(self):
        """Brak wpisu dla 45 °C — lepiej nic nie wpisac niz wpisac cudza wartosc."""
        v, opis = P.rozdzielczosc_dla_temperatury(ZGLOSZENIE, 45)
        self.assertIsNone(v)
        self.assertIn("45", opis)

    def test_opis_mowi_skad_wartosc(self):
        _v, opis = P.rozdzielczosc_dla_temperatury(ZGLOSZENIE, 70.3)
        self.assertEqual(opis, "dla 70 °C")

    def test_wynik_jest_liczba_a_nie_tekstem(self):
        self.assertIsInstance(wartosc(ZGLOSZENIE, 20), float)


class TestZapisyZZestawienia(unittest.TestCase):
    """Rozne formy, w jakich rozdzielczosc wg temperatury stoi w Zestawieniu."""

    def test_temperatura_z_dwukropkiem(self):
        self.assertEqual(wartosc("-20 °C: 0,047 °C\n5 °C: 0,026 °C", -20), 0.047)
        self.assertEqual(wartosc("-20 °C: 0,047 °C\n5 °C: 0,026 °C", 5), 0.026)

    def test_zakres_temperatur(self):
        tekst = "(10 ÷ 25) °C\nt: 0,024 °C\ndla 35 °C\nt: 0,027 °C"
        self.assertEqual(wartosc(tekst, 20), 0.024)
        self.assertEqual(wartosc(tekst, 35), 0.027)

    def test_zakres_z_myslnikiem(self):
        self.assertEqual(wartosc("dla (4-5) °C t: 0,026 °C\ndla 0°C t: 0,027 °C", 4.6),
                         0.026)

    def test_temperatura_za_wartoscia(self):
        """'t: 0,024 °C / dla (15 ÷ 35) °C' — zakres dotyczy wartosci PRZED nim."""
        tekst = "t: 0,024 °C\ndla (15 ÷ 35) °C\n\nRH: 0,1 %\ndla (30 ÷ 70) % RH"
        self.assertEqual(P.rozdzielczosci_wg_temperatury(tekst), [((15.0, 35.0), 0.024)])

    def test_temperatura_za_wartoscia_w_jednej_linii(self):
        tekst = "t: 0,5 °C\nt: 0,33 °C - dla 20 °C"
        self.assertEqual(wartosc(tekst, 20), 0.33)

    def test_linia_wilgotnosci_nie_jest_rozdzielczoscia_temperatury(self):
        tekst = "dla 35 °C\nt: 0,027 °C\nRH: 0,01 %"
        self.assertEqual(P.rozdzielczosci_wg_temperatury(tekst), [((35.0, 35.0), 0.027)])

    def test_temperatury_ujemne(self):
        tekst = "dla -25°C t: 0,056 °C\ndla -30°C t: 0,069 °C"
        self.assertEqual(wartosc(tekst, -30), 0.069)
        self.assertEqual(wartosc(tekst, -25), 0.056)

    def test_najblizszy_wpis_wygrywa(self):
        """80 i 70 °C sa obok siebie — 72 °C to 70, nie 80."""
        tekst = "dla 80°C t: 0,074 °C\ndla 70°C t: 0,056 °C"
        self.assertEqual(wartosc(tekst, 72), 0.056)


class TestDotychczasoweWartosci(unittest.TestCase):
    """Regresja: zwykle wpisy w kolumnie K maja dzialac dokladnie jak dotad."""

    def test_liczba_bez_zmian(self):
        self.assertEqual(P.rozdzielczosc_dla_temperatury(0.1, 70), (0.1, None))

    def test_liczba_calkowita_bez_zmian(self):
        self.assertEqual(P.rozdzielczosc_dla_temperatury(1, 70), (1, None))

    def test_myslnik_bez_zmian(self):
        self.assertEqual(P.rozdzielczosc_dla_temperatury("-", 70), ("-", None))

    def test_pusta_komorka(self):
        self.assertEqual(P.rozdzielczosc_dla_temperatury(None, 70), (None, None))

    def test_jedna_ogolna_wartosc_w_tekscie(self):
        """'t: 0,024 °C' bez temperatury — ta sama wartosc w kazdej zakladce."""
        for t in (-10, 20, 70):
            with self.subTest(temp=t):
                self.assertEqual(wartosc("t: 0,024 °C\nRH: 0,07 %", t), 0.024)

    def test_wartosc_ogolna_jako_zapas(self):
        """Gdy dla tej temperatury nie ma wpisu, a jest wartosc ogolna — bierzemy ja."""
        tekst = "t: 0,5 °C\nt: 0,33 °C - dla 20 °C"
        v, opis = P.rozdzielczosc_dla_temperatury(tekst, 70)
        self.assertEqual(v, 0.5)
        self.assertIn("ogolna", opis)

    def test_bez_temperatury_zakladki_nie_zgaduje(self):
        v, _opis = P.rozdzielczosc_dla_temperatury(ZGLOSZENIE, None)
        self.assertIsNone(v)


class TestTemperaturaZNazwyZakladki(unittest.TestCase):
    """Temperatura punktu bierze sie z nazwy zakladki kopii."""

    def test_punkt_samej_temperatury(self):
        self.assertEqual(A._t_z_nazwy_zakladki("70, -"), 70.0)

    def test_temperatura_ujemna(self):
        self.assertEqual(A._t_z_nazwy_zakladki("-10, 30"), -10.0)

    def test_powtorka_histerezy(self):
        self.assertEqual(A._t_z_nazwy_zakladki("25, 51 (2)"), 25.0)

    def test_nazwa_nieliczbowa(self):
        self.assertIsNone(A._t_z_nazwy_zakladki("Wyniki"))

    def test_pusta_nazwa(self):
        self.assertIsNone(A._t_z_nazwy_zakladki(""))
        self.assertIsNone(A._t_z_nazwy_zakladki(None))


if __name__ == "__main__":
    unittest.main(verbosity=2)
