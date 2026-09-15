# -*- coding: utf-8 -*-
"""
Testy rozpoznawania kanalow po JEDNOSTCE stojacej w kolumnie obok.

Zgloszenie: krok 1 nie potrafil zamienic pliku tekstowego na arkusz — konczyl
sie bledem 'No temperature column with data'. Logger zapisuje kanaly ogolnie:

    Position  Date        Time      Ch1_Value  Ch1_Unit  Ch2_Value  Ch2_unit
    1         12.08.2025  10:23:08   000046,9      %RH    000024,5  DEGREE C

Nazwa kolumny ('Ch1_Value') nie mowi nic o wielkosci, wiec dopasowanie po
slowach kluczowych ('temp', 'humid', ...) nie mialo za co chwycic. Fallback
pozycyjny tez nie pomagal, bo liczyl liczby bez uwzglednienia przecinka
dziesietnego — '000024,5' wygladalo dla niego na tekst.

Co gorsza, gdyby fallback zadzialal, wzialby PIERWSZA kolumne liczbowa —
czyli wilgotnosc — i wpisal ja do protokolu jako temperature.

Rozwiazanie: rola kolumny wynika z TRESCI sasiedniej kolumny jednostki. Ta
proba wchodzi DOPIERO wtedy, gdy zawiodly obie dotychczasowe (slowa kluczowe
i fallback pozycyjny) — pliki rozpoznawane do tej pory ida ta sama droga
co wczesniej.
"""

import io
import os
import sys
import unittest
from contextlib import redirect_stdout
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd

import analizuj_excele as A
from wspolne import nowa_piaskownica


NAGLOWEK = ("Position\tDate\tTime\tCh1_Value\tCh1_Unit\tCh2_Value\tCh2_unit"
            "\tCh3_Value\tCh3_unit")


def _pole(wartosc):
    """Zapis loggera: dopelnienie zerami do 6 cyfr i przecinek dziesietny."""
    return "  {:09.1f}".format(wartosc).replace(".", ",")


def plik_loggera(sciezka, wierszy=30, jedn_temp="  DEGREE C"):
    """Kopia ukladu z prawdziwego pliku — razem z CRLF na koncach linii."""
    linie = [NAGLOWEK]
    for i in range(wierszy):
        linie.append("\t".join([
            str(i + 1), "12.08.2025", "10:23:{:02d}".format(i % 60),
            _pole(46.9 + i * 0.1), "       %RH",
            _pole(24.5 + i * 0.01), jedn_temp,
            _pole(1005.0), "       hpa",
        ]))
    with open(sciezka, "w", encoding="utf-8", newline="\r\n") as f:
        f.write("\n".join(linie) + "\n")
    return sciezka


class TestDominujacaJednostka(unittest.TestCase):
    """Kolumna jednostki = ten sam krotki, nieliczbowy napis w kolko."""

    def jedn(self, wartosci):
        return A._dominujaca_jednostka(pd.Series(wartosci))

    def test_staly_napis_jest_jednostka(self):
        self.assertEqual(self.jedn(["  %RH", "%RH", " %RH ", "%RH"]), "%RH")

    def test_kolumna_liczb_nie_jest_jednostka(self):
        self.assertIsNone(self.jedn(["24,5", "24,6", "24,7", "24,8"]))

    def test_kolumna_zmiennych_napisow_nie_jest_jednostka(self):
        self.assertIsNone(self.jedn(["ok", "blad", "start", "stop", "ok"]))

    def test_za_malo_wierszy_to_nie_dowod(self):
        self.assertIsNone(self.jedn(["%RH", "%RH"]))

    def test_dlugi_opis_to_nie_jednostka(self):
        self.assertIsNone(self.jedn(["pomiar wilgotnosci wzgl."] * 10))

    def test_pojedyncze_odstepstwo_nie_psuje_rozpoznania(self):
        """Jeden uszkodzony wiersz na 20 nie moze przekreslic calej kolumny."""
        self.assertEqual(self.jedn(["DEGREE C"] * 19 + ["???"]), "DEGREE C")


class TestPrzypisanieRol(unittest.TestCase):
    """Kolumna wartosci dostaje role z jednostki stojacej PO PRAWEJ."""

    def ramka(self, jedn1="%RH", jedn2="DEGREE C", jedn3="hpa"):
        n = 10
        return pd.DataFrame({
            "Position":  range(1, n + 1),
            "Ch1_Value": ["  000046,9"] * n,
            "Ch1_Unit":  [jedn1] * n,
            "Ch2_Value": ["  000024,5"] * n,
            "Ch2_unit":  [jedn2] * n,
            "Ch3_Value": ["  001005,0"] * n,
            "Ch3_unit":  [jedn3] * n,
        })

    def test_temperatura_z_kolumny_ze_stopniami(self):
        self.assertEqual(A._kolumny_wg_jednostek(self.ramka())["temp"], "Ch2_Value")

    def test_wilgotnosc_z_kolumny_z_procentami(self):
        self.assertEqual(A._kolumny_wg_jednostek(self.ramka())["wilg"], "Ch1_Value")

    def test_cisnienie_pomijane(self):
        role = A._kolumny_wg_jednostek(self.ramka())
        self.assertNotIn("Ch3_Value", role.values())

    def test_warianty_zapisu_stopni(self):
        for jedn in ("DEGREE C", "deg C", "°C", " C ", "℃"):
            with self.subTest(jednostka=jedn):
                self.assertEqual(
                    A._kolumny_wg_jednostek(self.ramka(jedn2=jedn)).get("temp"),
                    "Ch2_Value")

    def test_warianty_zapisu_wilgotnosci(self):
        for jedn in ("%RH", "% rh", "%R.H.", "%"):
            with self.subTest(jednostka=jedn):
                self.assertEqual(
                    A._kolumny_wg_jednostek(self.ramka(jedn1=jedn)).get("wilg"),
                    "Ch1_Value")

    def test_fahrenheit_nie_udaje_celsjusza(self):
        """Wartosc w F wpisana jako °C to blad pomiarowy — kanal ma wypasc."""
        with redirect_stdout(io.StringIO()):
            role = A._kolumny_wg_jednostek(self.ramka(jedn2="DEGREE F"))
        self.assertIsNone(role.get("temp"))

    def test_bez_kolumn_jednostek_nic_nie_zwraca(self):
        df = pd.DataFrame({"Czas": ["10:00"] * 5, "Temperatura": ["24,5"] * 5})
        self.assertEqual(A._kolumny_wg_jednostek(df), {})


class TestLiczenieLiczb(unittest.TestCase):
    """Zapis europejski ('24,5') to liczba — na tym wykladal sie fallback."""

    def test_przecinek_dziesietny_liczy_sie_jako_liczba(self):
        df = pd.DataFrame({"a": ["  000024,5", "  000024,6", "  000024,7"]})
        self.assertEqual(A._licz_liczby(df, "a"), 3)

    def test_tekst_sie_nie_liczy(self):
        df = pd.DataFrame({"a": ["%RH", "%RH", "%RH"]})
        self.assertEqual(A._licz_liczby(df, "a"), 0)


class TestPelnyPrzebieg(unittest.TestCase):
    """Plik TXT -> arkusz .xlsx, tak jak robi to krok 1 panelu."""

    @classmethod
    def setUpClass(cls):
        folder = nowa_piaskownica("kanaly_jedn")
        wejscie = os.path.join(folder, "wejscie")
        cls.wyjscie = os.path.join(folder, "wyniki")
        os.makedirs(wejscie, exist_ok=True)
        os.makedirs(cls.wyjscie, exist_ok=True)
        txt = plik_loggera(os.path.join(wejscie, "HBA01001_test.txt"))

        with redirect_stdout(io.StringIO()) as buf:
            A.parse_txt_generic(Path(txt), cls.wyjscie)
        cls.log = buf.getvalue()
        cls.wynik = os.path.join(cls.wyjscie, "HBA01001_test_wynik.xlsx")
        cls.df = pd.read_excel(cls.wynik)

    def test_arkusz_powstal(self):
        self.assertTrue(os.path.isfile(self.wynik))

    def test_sa_obie_wielkosci(self):
        self.assertIn("Temperatura [°C]", self.df.columns)
        self.assertIn("Wilgotność [%RH]", self.df.columns)

    def test_temperatura_to_nie_wilgotnosc(self):
        """Ch1 to %RH — nie wolno jej wpisac jako temperatury."""
        self.assertAlmostEqual(self.df["Temperatura [°C]"].iloc[0], 24.5, places=1)
        self.assertAlmostEqual(
            self.df["Wilgotność [%RH]"].iloc[0], 46.9, places=1)

    def test_cisnienie_nie_trafia_do_arkusza(self):
        self.assertEqual(len(self.df.columns), 3)

    def test_data_czytana_po_europejsku(self):
        """'12.08.2025' to 12 sierpnia, nie 8 grudnia."""
        pierwszy = pd.Timestamp(self.df["Czas"].iloc[0])
        self.assertEqual((pierwszy.year, pierwszy.month, pierwszy.day), (2025, 8, 12))

    def test_wszystkie_wiersze_przepisane(self):
        self.assertEqual(len(self.df), 30)

    def test_log_mowi_skad_wzieto_kanaly(self):
        self.assertIn("po jednostkach", self.log)


if __name__ == "__main__":
    unittest.main(verbosity=2)
