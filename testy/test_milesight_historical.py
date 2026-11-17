# -*- coding: utf-8 -*-
"""
Eksport Milesight EM320-TH 'Historical Data' (.xlsx, zlecenie 249).

Oryginal: arkusz 'in', KAZDY wiersz CSV w jednej komorce kolumny A, najnowszy
zapis na gorze, '℃' zepsute kodowaniem:

    Time,Temp/â„ƒ,RH/%
    2026-11-11 09:35:11,26.7,31

Czesc plikow operator rozbil 'Tekstem jako kolumny' (przecinek + spacja), wiec
data i godzina trafily do osobnych kolumn:
    * naglowek rozbity:     Time | Temp/℃ | RH/% | (pusto)  — o kolumne za waski,
    * naglowek nierozbity:  'Time,Temp/℃,RH/%' w A1, reszta pusta.

Krok 1 we wszystkich trzech wariantach konczyl sie 'Brak prawidlowych danych'.
Numery i wartosci w tescie sa zmyslone; uklad jak w eksporcie.
"""

import datetime
import io
import os
import sys
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
import pandas as pd

import analizuj_excele as A
import pz_dane as P
from wspolne import nowa_piaskownica

EUI = "24E1240000000A01"
NAGLOWEK = "Time,Temp/â„\x83,RH/%"
START = datetime.datetime(2026, 11, 16, 6, 52, 41)

# (minuta, T, RH) — RH bywa calkowita, bywa z polowka
POMIARY = [(i, round(22.0 + 0.1 * i, 1), 30 if i % 2 else 30.5) for i in range(10)]


def _czas(i):
    return START + datetime.timedelta(minutes=i)


def _wiersze(wariant):
    """Wiersze arkusza 'in' (najnowsze na gorze, jak w eksporcie)."""
    dane = list(reversed(POMIARY))
    if wariant == "spakowany":
        return [[NAGLOWEK]] + [[f"{_czas(i):%Y-%m-%d %H:%M:%S},{t},{rh}"] for i, t, rh in dane]
    rozbite = [[datetime.datetime.combine(_czas(i).date(), datetime.time()),
                _czas(i).time(), t, rh] for i, t, rh in dane]
    if wariant == "rozbity_naglowek":
        return [["Time", "Temp/â„\x83", "RH/%", None]] + rozbite
    if wariant == "spakowany_naglowek":
        return [[NAGLOWEK, None, None, None]] + rozbite
    raise ValueError(wariant)


def przetworz(wariant, nazwa=None):
    folder = nowa_piaskownica("milesight_" + wariant)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "in"
    for w in _wiersze(wariant):
        ws.append(w)
    sciezka = Path(folder) / (nazwa or f"21 - {EUI}-Historical Data-2026-11-17_12-33.xlsx")
    wb.save(sciezka)
    wyj = Path(folder) / "wyniki"
    wyj.mkdir()
    bufor = io.StringIO()
    with mock.patch.object(A, "_mapa_pz", return_value={}), redirect_stdout(bufor):
        A.process_file(sciezka, wyj)
    pliki = list(wyj.glob("*_wynik.xlsx"))
    return pliki, bufor.getvalue()


class TestWariantyEksportu(unittest.TestCase):

    def sprawdz(self, wariant):
        pliki, log = przetworz(wariant)
        self.assertEqual(len(pliki), 1, log)
        df = pd.read_excel(pliki[0])
        self.assertEqual(len(df), len(POMIARY), log)
        # posortowane rosnaco i wartosci na swoich miejscach
        self.assertEqual(df["Czas"].iloc[0], pd.Timestamp(_czas(0)))
        self.assertEqual(df["Czas"].iloc[-1], pd.Timestamp(_czas(len(POMIARY) - 1)))
        for i, t, rh in POMIARY:
            wiersz = df[df["Czas"] == pd.Timestamp(_czas(i))].iloc[0]
            self.assertAlmostEqual(wiersz["Temperatura [°C]"], t)
            self.assertAlmostEqual(wiersz["Wilgotność [%RH]"], rh)
        return log

    def test_csv_w_jednej_kolumnie(self):
        self.assertIn("CSV w jednej kolumnie", self.sprawdz("spakowany"))

    def test_rozbity_z_przesunietym_naglowkiem(self):
        self.assertIn("naglowek przesuniety", self.sprawdz("rozbity_naglowek"))

    def test_rozbite_dane_naglowek_w_a1(self):
        self.sprawdz("spakowany_naglowek")


class TestNumerZNazwyPliku(unittest.TestCase):

    def test_eui_z_nazwy(self):
        for nazwa in (
            f"21 - {EUI}-Historical Data-2026-11-17_12-33.xlsx",
            f"10 - {EUI}-Historical Data-2026-11-17_12-38 (1).xlsx",
            f"9 - {EUI}-Historical Data-2026-11-17_12-30 1_wynik.xlsx",
            f"{EUI}-Historical Data-2026-11-17_12-30.xlsx",
        ):
            self.assertEqual(P.serial_z_nazwy_pliku(nazwa), EUI, nazwa)

    def test_inne_nazwy_bez_zmian(self):
        self.assertEqual(P.serial_z_nazwy_pliku("1970325 2026-07-31 12.19.00_wynik.xlsx"), "1970325")
        self.assertEqual(P.serial_z_nazwy_pliku("37025105_2_wynik.xlsx"), "37025105")


class TestZwykleTabeleNietkniete(unittest.TestCase):
    """Wyrownanie naglowka nie moze ruszac poprawnych tabel."""

    def test_zwykla_tabela(self):
        df = pd.DataFrame({"Czas": ["2026-11-16 06:52:41"] * 3, "Temp": [1.0, 2, 3], "RH": [4.0, 5, 6]})
        self.assertIs(A.scal_rozdzielona_date_i_godzine(df), df)

    def test_osobne_kolumny_daty_i_godziny_z_naglowkami(self):
        """'Data | Godzina | T | RH' z pelnym naglowkiem — to robi build_times, nie wyrownanie."""
        df = pd.DataFrame({"Data": ["2026-11-16"] * 3, "Godzina": ["06:52:41"] * 3,
                           "Temp": [1.0, 2, 3], "RH": [4.0, 5, 6]})
        self.assertIs(A.scal_rozdzielona_date_i_godzine(df), df)

    def test_przecinek_w_naglowku_czasu_przy_pelnych_kolumnach(self):
        df = pd.DataFrame({"Data, czas": ["2026-11-16 06:52:41"] * 3, "Temp": [1.0, 2, 3]})
        self.assertIs(A.scal_rozdzielona_date_i_godzine(df), df)

    def test_rozpakowanie_odrzuca_dane_w_innych_kolumnach(self):
        df_raw = pd.DataFrame([[NAGLOWEK, None], ["2026-11-16 06:52:41,22.0,30", 5]])
        self.assertIsNone(A.rozpakuj_csv_z_kolumny(df_raw))

    def test_rozpakowanie_bez_naglowka(self):
        df_raw = pd.DataFrame([["a,b,c"], ["1,2,3"], ["4,5,6"], ["7,8,9"]])
        self.assertIsNone(A.rozpakuj_csv_z_kolumny(df_raw))


if __name__ == "__main__":
    unittest.main(verbosity=2)
