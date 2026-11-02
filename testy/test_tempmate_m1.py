# -*- coding: utf-8 -*-
"""
Raport Tempmate M1 (.xls, zlecenie 234): arkusz 'List' z metadanymi i pusty 'Graph'.

    Summary
    Start Time | 30-Oct-26 13:50:25 | | Mean Kinetic Temperature (MKT) | 20.6°C
    ...
    No. | Time               | Temperature°C
    1   | 30-Oct-26 13:50:25 | 25.1

Wiersz 'Start Time ... Temperature' uchodzil za naglowek tabeli ('Brak prawidlowych
danych'), a pusty 'Graph' konczyl sie bledem 'single positional indexer is out-of-bounds'.
Numery i wartosci w tescie sa zmyslone; uklad jak w raporcie.
"""

import io
import os
import sys
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import analizuj_excele as A
import pz_dane as P
from wspolne import nowa_piaskownica

SERIAL = "TMM000000001"


def raport(nazwa_testu, dane, naglowek_temp="Temperature°C", serial=SERIAL):
    """Arkusze jak w .xls Tempmate; .xlsx, bo tym samym kodem czyta oba."""
    folder = nowa_piaskownica("tempmate_" + nazwa_testu)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "List"
    wiersze = [
        ["Data Report"], ["File created on:02-Nov-26 08:57:55"], ["Device Information"],
        ["Device Model", "M1", None, "Sensor Type", "Temperature(Internal)"],
        ["Serial Number", serial], [], ["Summary"],
        ["Start Time", "30-Oct-26 13:50:25", None, "Mean Kinetic Temperature (MKT)", "20.6°C"],
        ["Stop Time", "30-Oct-26 13:55:25", None, "Average", "25.0°C"],
        [], ["No.", "Time", naglowek_temp],
    ] + dane
    for w in wiersze:
        ws.append(w)
    wb.create_sheet("Graph")
    sciezka = Path(folder) / f"{SERIAL}_20261102085741_LOG 21.xlsx"
    wb.save(sciezka)
    wyj = Path(folder) / "wyniki"
    wyj.mkdir()
    bufor = io.StringIO()
    with mock.patch.object(A, "_mapa_pz", return_value={}), redirect_stdout(bufor):
        A.parse_excel_inspect(sciezka, wyj)
    return wyj, bufor.getvalue()


def odczyt(plik):
    wb = openpyxl.load_workbook(plik, data_only=True)
    try:
        return [tuple(w) for w in wb.worksheets[0].iter_rows(min_row=2, values_only=True)]
    finally:
        wb.close()


class TestTempmateM1(unittest.TestCase):

    def test_pomiary_i_nazwa_z_numeru_seryjnego(self):
        wyj, log = raport("podstawowy", [["1", "30-Oct-26 13:50:25", 25.1],
                                         ["2", "30-Oct-26 13:51:25", 25.2]])
        self.assertEqual([p.name for p in wyj.iterdir()], [f"{SERIAL}_wynik.xlsx"])
        wiersze = odczyt(wyj / f"{SERIAL}_wynik.xlsx")
        self.assertEqual([w[1] for w in wiersze], [25.1, 25.2])
        self.assertEqual(str(wiersze[0][0]), "2026-10-30 13:50:25")
        self.assertIn("Tempmate M1", log)
        self.assertNotIn("Błąd", log)
        self.assertIn("pusty arkusz", log)

    def test_low_voltage_i_dziura_w_logu(self):
        dane = [[str(i + 1), f"30-Oct-26 13:{50 + i}:25", 25.0] for i in range(6)]
        dane += [["7", "Low Voltage", None], ["8", "30-Oct-26 16:00:25", 24.0]]
        wyj, log = raport("low_voltage", dane)
        self.assertEqual(len(odczyt(wyj / f"{SERIAL}_wynik.xlsx")), 7)
        self.assertIn("1 wierszy bez pomiaru: 'Low Voltage'", log)
        self.assertIn("Dziura w zapisie: 2026-10-30 13:55:25", log)

    def test_rowny_zapis_bez_ostrzezen(self):
        dane = [[str(i + 1), f"30-Oct-26 13:{50 + i}:25", 25.0] for i in range(6)]
        _wyj, log = raport("rowny", dane)
        self.assertNotIn("⚠", log)

    def test_fahrenheit_nie_jest_brany_za_celsjusze(self):
        wyj, log = raport("fahrenheit", [["1", "30-Oct-26 13:50:25", 77.2]],
                          naglowek_temp="Temperature°F")
        self.assertEqual(list(wyj.iterdir()), [])
        self.assertIn("Fahrenheita", log)


class TestPusteArkusze(unittest.TestCase):

    def test_pusty_arkusz_nie_przerywa_kolejnych(self):
        folder = nowa_piaskownica("tempmate_pusty_pierwszy")
        wb = openpyxl.Workbook()
        wb.active.title = "Graph"
        ws = wb.create_sheet("Dane")
        ws.append(["Czas", "Temperatura"])
        ws.append(["2026-10-30 13:50:25", 25.1])
        sciezka = Path(folder) / "plik.xlsx"
        wb.save(sciezka)
        wyj = Path(folder) / "wyniki"
        wyj.mkdir()
        with mock.patch.object(A, "_mapa_pz", return_value={}), redirect_stdout(io.StringIO()):
            A.parse_excel_inspect(sciezka, wyj)
        self.assertEqual(len(list(wyj.iterdir())), 1)


class TestDopasowanieDoPZ(unittest.TestCase):

    def test_dopisek_log_z_nazwy_pliku(self):
        for nazwa, serial in (("40125624_LOG 66.xls", "40125624"),
                              ("40125624_LOG 66_wynik.xlsx", "40125624"),
                              ("40125624_2_LOG 66_wynik.xlsx", "40125624"),
                              ("40125624 log66.xls", "40125624"),
                              ("LOG 66.xls", "LOG 66"),
                              ("37025105_2_wynik.xlsx", "37025105"),
                              ("KATALOG.xls", "KATALOG")):
            with self.subTest(nazwa=nazwa):
                self.assertEqual(P.serial_z_nazwy_pliku(nazwa), serial)

    def test_typ_bez_dwukropka(self):
        tekst = ("Obiekty wzorcowania:\n"
                 "1) Termometr elektroniczny (rejestrator, 2 szt.) typ: Tempmate M1:\n"
                 "• nr fabr.: TMM000000001, nr ewid.: Q/LOG/18,\n"
                 "• nr fabr.: TMM000000002, nr ewid.: Q/LOG/21,\n"
                 "wytwórca: Tempmate.\n"
                 "Metoda wzorcowania: \nMetoda porównawcza w komorze klimatycznej.\n")
        lista = P.parsuj_tekst(tekst)
        self.assertEqual([(p.typ, p.nr_fabr, p.nr_ewid) for p in lista], [
            ("Tempmate M1", "TMM000000001", "Q/LOG/18"),
            ("Tempmate M1", "TMM000000002", "Q/LOG/21")])
        self.assertEqual(lista[0].wytworca, "Tempmate")


if __name__ == "__main__":
    unittest.main(verbosity=2)
