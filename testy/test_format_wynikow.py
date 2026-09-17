# -*- coding: utf-8 -*-
"""
Testy wygladu plikow wynikowych kroku 1.

Pliki z 'wyniki/' oglada sie recznie — sprawdza sie w nich, z jakich probek
loggera powstal punkt protokolu. Wychodzily jako surowy zrzut: kolumna czasu
byla za waska, wiec Excel pokazywal '########', naglowek nie roznil sie od
danych, nie bylo ramek ani wyrownania.

Teraz kazdy plik wynikowy (takze zestawienie zbiorcze) dostaje wyglad tabeli:
pogrubiony, wysrodkowany naglowek na szarym tle, ramki, wysrodkowane dane,
pelny format daty i przyklejony naglowek przy przewijaniu.

Osobno pilnujemy, ze oznaczenia punktow nanoszone pozniej przez krok 2
(zielony/pomaranczowy) NIE kasuja tego formatowania.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
import pandas as pd

import analizuj_excele as A
import generuj_obserwacje as G
from wspolne import nowa_piaskownica


def przykladowa_ramka(wierszy=10, naglowek_wilg="Wilgotność [%RH]"):
    return pd.DataFrame({
        "Czas": pd.date_range("2026-09-15 14:21:00", periods=wierszy, freq="min"),
        "Temperatura [°C]": [26.179 - i * 0.01 for i in range(wierszy)],
        naglowek_wilg: [47.949 - i * 0.1 for i in range(wierszy)],
    })


class TestWygladArkusza(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.folder = nowa_piaskownica("format_wynikow")
        cls.sciezka = os.path.join(cls.folder, "wynik.xlsx")
        A._zapisz_z_formatem(przykladowa_ramka(), cls.sciezka)
        cls.wb = openpyxl.load_workbook(cls.sciezka)
        cls.ws = cls.wb.active

    @classmethod
    def tearDownClass(cls):
        cls.wb.close()

    # --- naglowek ---
    def test_naglowek_pogrubiony(self):
        for kol in range(1, 4):
            with self.subTest(kolumna=kol):
                self.assertTrue(self.ws.cell(row=1, column=kol).font.bold)

    def test_naglowek_wysrodkowany(self):
        self.assertEqual(self.ws.cell(row=1, column=1).alignment.horizontal, "center")

    def test_naglowek_ma_tlo(self):
        self.assertEqual(self.ws.cell(row=1, column=1).fill.fgColor.rgb, "00D9D9D9")

    def test_naglowek_zawija_dlugi_tekst(self):
        self.assertTrue(self.ws.cell(row=1, column=2).alignment.wrap_text)

    def test_naglowek_zostaje_przy_przewijaniu(self):
        self.assertEqual(self.ws.freeze_panes, "A2")

    # --- czas ---
    def test_czas_ma_pelny_format(self):
        self.assertEqual(self.ws["A2"].number_format, A.FORMAT_CZASU)

    def test_kolumna_czasu_dosc_szeroka(self):
        """Za waska kolumna to wlasnie '########' zamiast godziny."""
        szer = self.ws.column_dimensions["A"].width
        self.assertGreaterEqual(szer, len("2026-09-15 14:21:00"))

    def test_czas_zostal_data_a_nie_tekstem(self):
        self.assertIsInstance(self.ws["A2"].value, __import__("datetime").datetime)

    # --- dane ---
    def test_dane_wysrodkowane(self):
        self.assertEqual(self.ws["B2"].alignment.horizontal, "center")

    def test_dane_maja_ramki(self):
        for adres in ("A2", "B2", "C2"):
            with self.subTest(komorka=adres):
                self.assertTrue(self.ws[adres].border.left.style)

    def test_ostatni_wiersz_tez_sformatowany(self):
        ostatni = self.ws.max_row
        self.assertTrue(self.ws.cell(row=ostatni, column=1).border.bottom.style)
        self.assertEqual(self.ws.cell(row=ostatni, column=1).number_format,
                         A.FORMAT_CZASU)

    def test_wartosci_nie_zostaly_zmienione(self):
        self.assertAlmostEqual(self.ws["B2"].value, 26.179, places=3)


class TestKolorowNaglowkow(unittest.TestCase):
    """
    Tlo naglowka mowi, jaka wielkosc stoi w kolumnie. W zestawieniu zbiorczym
    Temp i Wilg kolejnych przyrzadow stoja na przemian — sam szary kolor niczego
    tam nie rozdziela.
    """

    def kolor(self, naglowek):
        return A.kolor_naglowka(naglowek)

    def test_czas_szary(self):
        for naglowek in ("Czas", "Timestamp (Local)", "Data"):
            with self.subTest(naglowek=naglowek):
                self.assertEqual(self.kolor(naglowek), A.KOLOR_NAGL_CZAS)

    def test_temperatura_brzoskwiniowa(self):
        for naglowek in ("Temperatura [°C]", "Temp KH30-03AA", "Temperature (°C)"):
            with self.subTest(naglowek=naglowek):
                self.assertEqual(self.kolor(naglowek), A.KOLOR_NAGL_TEMP)

    def test_wilgotnosc_blekitna(self):
        for naglowek in ("Wilgotność [%RH]", "Wilg KH30-03AA", "Humidity (%RH)"):
            with self.subTest(naglowek=naglowek):
                self.assertEqual(self.kolor(naglowek), A.KOLOR_NAGL_WILG)

    def test_wilgotnosc_ma_pierwszenstwo_przed_temperatura(self):
        """'Wilgotność [%RH]' nie moze wyjsc na temperature przez slowo '[°C]'."""
        self.assertEqual(self.kolor("Wilgotność wzgledna przy 25 °C"),
                         A.KOLOR_NAGL_WILG)

    def test_nieznana_wielkosc_dostaje_kolor_zapasowy(self):
        self.assertEqual(self.kolor("Ciśnienie [hPa]"), A.KOLOR_NAGL_INNE)

    def test_wszystkie_kolory_sa_rozne(self):
        kolory = {A.KOLOR_NAGL_CZAS, A.KOLOR_NAGL_TEMP,
                  A.KOLOR_NAGL_WILG, A.KOLOR_NAGL_INNE}
        self.assertEqual(len(kolory), 4)

    def test_kolory_trafiaja_do_pliku(self):
        folder = nowa_piaskownica("format_kolory")
        sciezka = os.path.join(folder, "wynik.xlsx")
        A._zapisz_z_formatem(przykladowa_ramka(3), sciezka)
        wb = openpyxl.load_workbook(sciezka)
        self.addCleanup(wb.close)
        ws = wb.active
        self.assertEqual(ws["A1"].fill.fgColor.rgb[-6:], A.KOLOR_NAGL_CZAS)
        self.assertEqual(ws["B1"].fill.fgColor.rgb[-6:], A.KOLOR_NAGL_TEMP)
        self.assertEqual(ws["C1"].fill.fgColor.rgb[-6:], A.KOLOR_NAGL_WILG)


class TestSzerokosciKolumn(unittest.TestCase):
    """Naglowki zestawienia zbiorczego zawieraja nazwy plikow — bywaja dlugie."""

    @classmethod
    def setUpClass(cls):
        cls.folder = nowa_piaskownica("format_szerokosci")

    def szerokosc(self, naglowek):
        sciezka = os.path.join(self.folder, "x.xlsx")
        A._zapisz_z_formatem(przykladowa_ramka(3, naglowek), sciezka)
        wb = openpyxl.load_workbook(sciezka)
        self.addCleanup(wb.close)
        return wb.active.column_dimensions["C"].width

    def test_krotki_naglowek_ma_minimalna_szerokosc(self):
        self.assertEqual(self.szerokosc("RH"), A.SZEROKOSC_MIN)

    def test_dlugi_naglowek_nie_rozpycha_bez_konca(self):
        dlugi = "Wilg KH30-03AA_2026915122158_bardzo_dluga_nazwa"
        self.assertEqual(self.szerokosc(dlugi), A.SZEROKOSC_MAX)


class TestFormatPrzezywaOznaczenia(unittest.TestCase):
    """
    Krok 2 koloruje w tych plikach wybrane okna. Nie wolno mu przy okazji
    zgubic ramek, wyrownania ani formatu daty — inaczej po kroku 2 plik znow
    wygladalby jak surowy zrzut.
    """

    @classmethod
    def setUpClass(cls):
        folder = nowa_piaskownica("format_oznaczenia")
        cls.sciezka = os.path.join(folder, "wynik.xlsx")
        A._zapisz_z_formatem(przykladowa_ramka(20), cls.sciezka)
        G._oznacz_wyniki_xlsx(cls.sciezka, {0: {5, 6, 7}, 1: {12, 13}},
                              {1: "punkt na styku"})
        cls.wb = openpyxl.load_workbook(cls.sciezka)
        cls.ws = cls.wb.active

    @classmethod
    def tearDownClass(cls):
        cls.wb.close()

    def test_naglowek_dalej_wyrozniony(self):
        self.assertTrue(self.ws["A1"].font.bold)
        self.assertEqual(self.ws["A1"].fill.fgColor.rgb, "00D9D9D9")

    def test_oznaczony_wiersz_zachowuje_ramke_i_format(self):
        c = self.ws.cell(row=5, column=1)
        self.assertTrue(c.border.left.style)
        self.assertEqual(c.number_format, A.FORMAT_CZASU)
        self.assertEqual(c.alignment.horizontal, "center")

    def test_kolory_punktow_naniesione(self):
        self.assertEqual(self.ws.cell(row=5, column=1).fill.fgColor.rgb,
                         G.FILL_DARK.fgColor.rgb)
        self.assertEqual(self.ws.cell(row=12, column=1).fill.fgColor.rgb,
                         G.FILL_WARN_DARK.fgColor.rgb)

    def test_szerokosc_i_przypiecie_naglowka_zostaja(self):
        self.assertGreaterEqual(self.ws.column_dimensions["A"].width, 20)
        self.assertEqual(self.ws.freeze_panes, "A2")

    def test_niezaznaczony_wiersz_nadal_bialy_ale_sformatowany(self):
        c = self.ws.cell(row=4, column=1)
        self.assertIn(c.fill.fgColor.rgb, (None, "00000000"))
        self.assertTrue(c.border.left.style)


if __name__ == "__main__":
    unittest.main(verbosity=2)
