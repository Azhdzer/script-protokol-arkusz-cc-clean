# -*- coding: utf-8 -*-
"""
Testy CSV z blokiem opisowym nad naglowkiem danych (loggery typu KH30).

Zgloszenie: krok 1 nie potrafil odczytac plikow loggera KH30 —
'No time column detected. Available columns: [Local Timezone is UTC +02:00:00,
Unnamed: 1, ...]'. Plik zaczyna sie od kilku wierszy opisu urzadzenia, a
prawdziwy naglowek danych stoi dopiero w 10. wierszu:

    Local Timezone is UTC +02:00:00;;;;;;;
    Device Name;First Timestamp (UTC+0);...
    KH30-03AA;15.09.2026 12:21:58;...
    ...
    ID;Imprecise time;...;Timestamp (UTC+0);Timestamp (Local);Temperature (°C);Humidity (%RH)
    0;Inactive;...;15.09.2026 12:21;15.09.2026 14:21;26.179;47.949

Trzy pulapki tego ukladu:

  1. Naglowek nizej niz pierwszy wiersz — pandas brala za nazwy kolumn tekst opisu.
  2. Czas podany DWA razy: w UTC i lokalnie. Do protokolu idzie czas LOKALNY —
     pomiar opisujemy zegarem laboratorium.
  3. Slowa kluczowe lapia kolumny STATUSU: 'Imprecise time' (wartosci
     'Inactive') pasuje do 'time', a 'Temperature Alert' do 'Temperature' —
     i obie stoja PRZED prawdziwymi kolumnami.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd

import analizuj_excele as A
from wspolne import nowa_piaskownica

OPIS = [
    "Local Timezone is UTC +02:00:00;;;;;;;",
    "Device Name;First Timestamp (UTC+0);Last Timestamp (UTC+0);;;;;",
    "KH30-03AA;15.09.2026 12:21:58;17.09.2026 04:42:58;;;;;",
    "The 'IsTimeImprecise' field signals whether or not the timestamp is reliable.;;;;;;;",
    ";;;;;;;",
    ";;;;;;Temperature (°C);Humidity (%RH)",
    "Higher Alarm Limit;;;;;;Disabled;Disabled",
    "Lower Alarm Limit;;;;;;Disabled;Disabled",
    ";;;;;;;",
]
NAGLOWEK = ("ID;Imprecise time;Temperature Alert;Humidity Alert;"
            "Timestamp (UTC+0);Timestamp (Local);Temperature (°C);Humidity (%RH)")


def plik_kh30(sciezka, wierszy=20):
    """Plik w ukladzie loggera KH30 — z blokiem opisowym i CRLF."""
    linie = list(OPIS) + [NAGLOWEK]
    for i in range(wierszy):
        linie.append(
            f"{i};Inactive;Inactive;Inactive;"
            f"15.09.2026 12:{21 + i:02d};15.09.2026 14:{21 + i:02d};"
            f"{26.179 - i * 0.01:.3f};{47.949 - i * 0.1:.3f}")
    with open(sciezka, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write("\n".join(linie) + "\n")
    return sciezka


class TestSzukanieNaglowka(unittest.TestCase):
    """Naglowek danych = wiersz, ktory ma NARAZ czas i wielkosc mierzona."""

    @classmethod
    def setUpClass(cls):
        cls.folder = nowa_piaskownica("csv_blok_opisowy")
        cls.plik = plik_kh30(os.path.join(cls.folder, "KH30-03AA.csv"))

    def test_znajduje_wlasciwy_wiersz(self):
        self.assertEqual(A.znajdz_wiersz_naglowka(self.plik, ";", "utf-8-sig"), 9)

    def test_blok_opisowy_nie_udaje_naglowka(self):
        """Wiersz 'Device Name;First Timestamp...' ma czas, ale nie ma temperatury."""
        self.assertNotEqual(A.znajdz_wiersz_naglowka(self.plik, ";", "utf-8-sig"), 1)

    def test_zly_separator_nic_nie_znajduje(self):
        self.assertIsNone(A.znajdz_wiersz_naglowka(self.plik, "\t", "utf-8-sig"))

    def test_plik_bez_takiego_ukladu_zwraca_none(self):
        zwykly = os.path.join(self.folder, "zwykly.csv")
        with open(zwykly, "w", encoding="utf-8") as f:
            f.write("a;b;c\n1;2;3\n")
        self.assertIsNone(A.znajdz_wiersz_naglowka(zwykly, ";", "utf-8"))

    def test_wczytanie_daje_wlasciwe_kolumny(self):
        df = A.read_csv_z_blokiem_opisowym(self.plik)
        self.assertIsNotNone(df)
        self.assertIn("Timestamp (Local)", df.columns)
        self.assertIn("Temperature (°C)", df.columns)


class TestWyborCzasu(unittest.TestCase):
    """Czas lokalny ma pierwszenstwo; kolumny statusu odpadaja."""

    KOLUMNY = ["Imprecise time", "Timestamp (UTC+0)", "Timestamp (Local)"]

    def ramka(self):
        return pd.DataFrame({
            "Imprecise time": ["Inactive"] * 5,
            "Timestamp (UTC+0)": [f"15.09.2026 12:{21+i:02d}" for i in range(5)],
            "Timestamp (Local)": [f"15.09.2026 14:{21+i:02d}" for i in range(5)],
        })

    def test_lokalny_idzie_pierwszy(self):
        self.assertEqual(A.preferuj_czas_lokalny(self.KOLUMNY)[0], "Timestamp (Local)")

    def test_bez_lokalnego_kolejnosc_bez_zmian(self):
        kol = ["Data", "Godzina"]
        self.assertEqual(A.preferuj_czas_lokalny(kol), kol)

    def test_jedna_kolumna_czasu_nietknieta(self):
        self.assertEqual(A.preferuj_czas_lokalny(["Czas"]), ["Czas"])

    def test_kolumna_statusu_odpada(self):
        zostaly = A.kolumny_czasu_z_datami(self.ramka(), self.KOLUMNY)
        self.assertNotIn("Imprecise time", zostaly)

    def test_gdy_zadna_nie_ma_dat_lista_zostaje(self):
        """Wtedy blad ma sie zglosic jak dotad, a nie zniknac po cichu."""
        df = pd.DataFrame({"Imprecise time": ["Inactive"] * 5})
        self.assertEqual(A.kolumny_czasu_z_datami(df, ["Imprecise time"]),
                         ["Imprecise time"])

    def test_razem_wychodzi_czas_lokalny(self):
        df = self.ramka()
        kol = A.preferuj_czas_lokalny(A.kolumny_czasu_z_datami(df, self.KOLUMNY))
        czasy = A.build_times(df, kol)
        self.assertEqual(czasy.iloc[0].hour, 14)


class TestWyborKolumnyWartosci(unittest.TestCase):
    """'Temperature Alert' pasuje do slowa, ale liczb nie ma."""

    def ramka(self):
        return pd.DataFrame({
            "Temperature Alert": ["Inactive"] * 5,
            "Temperature (°C)": [26.179, 26.35, 26.265, 25.923, 25.667],
        })

    def test_pomija_kolumne_statusu(self):
        df = self.ramka()
        kol = A.find_cols(df.columns, A.TEMP_KW)
        self.assertEqual(A.pierwsza_kolumna_z_liczbami(df, kol), "Temperature (°C)")

    def test_pierwsza_z_danymi_ma_pierwszenstwo(self):
        """Gdy pierwsza dopasowana kolumna ma liczby — zostaje wybrana."""
        df = pd.DataFrame({"Temperatura": [1.0, 2.0], "Temp zapasowa": [3.0, 4.0]})
        kol = A.find_cols(df.columns, A.TEMP_KW)
        self.assertEqual(A.pierwsza_kolumna_z_liczbami(df, kol), kol[0])

    def test_pusta_lista_daje_none(self):
        self.assertIsNone(A.pierwsza_kolumna_z_liczbami(pd.DataFrame(), []))


class TestPelnyPrzebieg(unittest.TestCase):
    """Plik KH30 -> arkusz .xlsx, tak jak robi to krok 1 panelu."""

    @classmethod
    def setUpClass(cls):
        import io
        from contextlib import redirect_stdout
        from pathlib import Path
        folder = nowa_piaskownica("csv_kh30")
        wyjscie = os.path.join(folder, "wyniki")
        os.makedirs(wyjscie, exist_ok=True)
        plik = plik_kh30(os.path.join(folder, "KH30-03AA_2026.csv"))
        with redirect_stdout(io.StringIO()):
            A.parse_csv_generic(Path(plik), wyjscie)
        cls.df = pd.read_excel(os.path.join(wyjscie, "KH30-03AA_2026_wynik.xlsx"))

    def test_wszystkie_wiersze_przepisane(self):
        self.assertEqual(len(self.df), 20)

    def test_czas_jest_lokalny_a_nie_utc(self):
        """W pliku UTC startuje 12:21, lokalny 14:21 (strefa UTC+02:00)."""
        self.assertEqual(pd.Timestamp(self.df["Czas"].iloc[0]).hour, 14)

    def test_temperatura_z_wlasciwej_kolumny(self):
        self.assertAlmostEqual(self.df["Temperatura [°C]"].iloc[0], 26.179, places=3)

    def test_wilgotnosc_z_wlasciwej_kolumny(self):
        self.assertAlmostEqual(self.df["Wilgotność [%RH]"].iloc[0], 47.949, places=3)

    def test_kolumny_statusu_nie_trafiaja_do_wyniku(self):
        self.assertEqual(len(self.df.columns), 3)


if __name__ == "__main__":
    unittest.main(verbosity=2)
