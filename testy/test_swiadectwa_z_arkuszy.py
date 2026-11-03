# -*- coding: utf-8 -*-
"""
Krok 5 — swiadectwa Word z gotowych arkuszy obliczeniowych (folder 'Etap_5_ark_obl_to_word').

Arkusze zrobione wczesniej (bez Worda) i uzupelnione po kilku pomiarach. Swiadectwa
maja wyjsc tak samo jak z kroku 3: wyniki z arkusza, dane przyrzadu z PZ, data z K4
zakladek, a warunki srodowiskowe z protokolow lezacych obok (zakladka -> blok Strony 3
o tej samej dacie i nastawie). Pelne porownanie z krokiem 3 na zleceniu 234 — recznie.
"""

import datetime
import os
import sys
import unittest
import unittest.mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import generuj_swiadectwa as S
from wspolne import nowa_piaskownica

D = datetime.date


def arkusz(sciezka, zakladki=("30, -", "0, -"), przetwornik=False):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for i, n in enumerate(zakladki):
        ws = wb.create_sheet(n)
        ws["K4"] = "30.10.2026"
        ws["E5"], ws["E6"], ws["G6"] = "Tempmate", "M1", "TMM1"
        ws["K10"] = "Przetwornik" if przetwornik else "Termohigrometr"
        ws["K15"], ws["M22"] = "001", 4
    wb.create_sheet("Wyniki")
    wb.save(sciezka)


def protokol(sciezka, bloki, seriale=()):
    """bloki: [(T, RH, 'dd.mm.rrrr', start, koniec, (t1, rh1), (t2, rh2))]"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Strona 3"
    ws2 = wb.create_sheet("Strona 2")
    for i, s in enumerate(seriale):
        ws2.cell(11 + i, 1).value, ws2.cell(11 + i, 5).value = i + 1, s
    for i, (t, rh, data, st, kon, fg1, fg2) in enumerate(bloki):
        r0 = 20 + i * 5
        ws.cell(r0, 1).value = i + 1
        ws.cell(r0, 2).value, ws.cell(r0, 3).value = t, rh
        ws.cell(r0 + 1, 5).value, ws.cell(r0 + 2, 5).value, ws.cell(r0 + 3, 5).value = data, st, kon
        (ws.cell(r0 + 2, 6).value, ws.cell(r0 + 2, 7).value) = fg1
        (ws.cell(r0 + 3, 6).value, ws.cell(r0 + 3, 7).value) = fg2
    wb.save(sciezka)


class TestPlikiWFolderze(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.f = nowa_piaskownica("swiadectwa_pliki")
        for n in ("234_LA_TH_2026 - ILAJ - Wzór - 10 - TMM9.xlsx",
                  "234_LA_TH_2026 - ILAJ - Wzór - 2 - TMM2.xlsx",
                  "xxx_LA_TH_2026 - ILAJ - szablon.xlsx",
                  "234_LA_TH_2026_TMM2_opracowanie danych.xlsx"):
            arkusz(os.path.join(cls.f, n))
        protokol(os.path.join(cls.f, "232, 234_LA_TH_2026 - protokół CC-04.xlsx"), [])
        wb = openpyxl.Workbook()
        wb.save(os.path.join(cls.f, "inny.xlsx"))             # bez 'Wyniki'

    def test_arkusze_w_kolejnosci_przyrzadow(self):
        self.assertEqual(S.arkusze_w_folderze(self.f),
                         ["234_LA_TH_2026 - ILAJ - Wzór - 2 - TMM2.xlsx",
                          "234_LA_TH_2026 - ILAJ - Wzór - 10 - TMM9.xlsx"])

    def test_protokoly(self):
        self.assertEqual(S.protokoly_w_folderze(self.f), ["232, 234_LA_TH_2026 - protokół CC-04.xlsx"])

    def test_brak_folderu(self):
        self.assertEqual(S.arkusze_w_folderze(os.path.join(self.f, "nie_ma")), [])


class TestDopasowanieBloku(unittest.TestCase):

    BLOKI = [
        {"protokol": "p1", "nr": 1, "t": 30.3, "rh": 0.0, "data": D(2026, 10, 30)},
        {"protokol": "p1", "nr": 2, "t": 0.2, "rh": None, "data": D(2026, 11, 1)},
        {"protokol": "p2", "nr": 1, "t": 0.3, "rh": None, "data": D(2026, 11, 5)},   # domiar
        {"protokol": "p1", "nr": 3, "t": 25.0, "rh": 50.0, "data": D(2026, 10, 31)},
    ]

    def nr(self, nazwa, k4):
        b = S.dopasuj_blok({"nazwa": nazwa, "K4_val": k4}, self.BLOKI)
        return (b["protokol"], b["nr"]) if b else None

    def test_po_dacie_i_nastawie(self):
        self.assertEqual(self.nr("30, -", "30.10.2026"), ("p1", 1))

    def test_domiar_z_innego_protokolu(self):
        self.assertEqual(self.nr("0, -", "05.11.2026"), ("p2", 1))
        self.assertEqual(self.nr("0, -", datetime.datetime(2026, 11, 1)), ("p1", 2))

    def test_punkt_z_wilgotnoscia(self):
        self.assertEqual(self.nr("25, 50", "31.10.2026"), ("p1", 3))
        self.assertIsNone(self.nr("25, -", "31.10.2026"))       # tylko t != blok z RH

    def test_inna_data_albo_bez_daty(self):
        self.assertIsNone(self.nr("30, -", "31.10.2026"))
        self.assertIsNone(self.nr("30, -", None))


class TestZlecenieRozbiteNaProtokoly(unittest.TestCase):
    """Jedno zlecenie mierzone w 3 czesciach: te same punkty tego samego dnia w 3 protokolach."""

    @classmethod
    def setUpClass(cls):
        f = nowa_piaskownica("swiadectwa_czesci")
        for i, (st, seriale) in enumerate((("08:00", ["TMM1", "TMM2"]),
                                           ("11:00", ["TMM3"]),
                                           ("14:00", ["Q/4"])), start=1):
            protokol(os.path.join(f, f"233_LA_TH_2026 - protokol cz{i}.xlsx"),
                     [(30.2, 0, "30.10.2026", st, st[:3] + "05", (22.0 + i, 40), (22.0 + i, 40))],
                     seriale)
        cls.bloki = []
        for p in S.protokoly_w_folderze(f):
            cls.bloki += S.bloki_protokolu(os.path.join(f, p))

    def protokol_przyrzadu(self, serial):
        return S.dopasuj_blok({"nazwa": "30, -", "K4_val": "30.10.2026"}, self.bloki, serial)["protokol"]

    def test_seriale_ze_strony2(self):
        self.assertEqual(self.bloki[0]["seriale"], {"TMM1", "TMM2"})
        self.assertEqual(self.bloki[2]["seriale"], {"Q_4"})

    def test_blok_z_protokolu_przyrzadu(self):
        self.assertEqual(self.protokol_przyrzadu("TMM2"), "233_LA_TH_2026 - protokol cz1.xlsx")
        self.assertEqual(self.protokol_przyrzadu("TMM3"), "233_LA_TH_2026 - protokol cz2.xlsx")
        self.assertEqual(self.protokol_przyrzadu("Q_4"), "233_LA_TH_2026 - protokol cz3.xlsx")

    def test_przyrzad_spoza_protokolow_bierze_punkt_po_dacie(self):
        self.assertEqual(self.protokol_przyrzadu("INNY"), "233_LA_TH_2026 - protokol cz1.xlsx")

    def test_punktu_brak_w_jego_protokole_domiar_z_innego(self):
        bloki = [dict(b) for b in self.bloki]
        bloki[1]["t"] = 10.0                       # cz2 nie ma punktu 30
        b = S.dopasuj_blok({"nazwa": "30, -", "K4_val": "30.10.2026"}, bloki, "TMM3")
        self.assertIsNotNone(b)


class TestProtokolyZFolderuGlownego(unittest.TestCase):
    """Protokoly czesci zlecenia leza w folderze glownym, arkusze w folderze kroku 5."""

    def test_tylko_protokoly_tego_zlecenia(self):
        glowny = nowa_piaskownica("swiadectwa_glowny")
        ark = os.path.join(glowny, "Etap_5")
        os.makedirs(ark)
        arkusz(os.path.join(ark, "232_LA_TH_2026 - ILAJ - Wzór - 1 - TMM1.xlsx"))
        for n in ("232_LA_TH_2026 - protokół CC.xlsx", "232_LA_TH_2026 - protokół_2 CC-04.xlsx",
                  "231, 232_LA_TH_2026 - protokół CC.xlsx", "230_LA_TH_2026 - protokół CC.xlsx",
                  "2232_LA_TH_2026 - protokół CC.xlsx"):
            protokol(os.path.join(glowny, n), [])
        protokol(os.path.join(ark, "232_LA_TH_2026 - protokół_3 CC-04.xlsx"), [])
        with unittest.mock.patch.object(S, "FOLDER", glowny), \
                unittest.mock.patch.object(S, "FOLDER_ARKUSZY", ark):
            wynik = S.sciezki_protokolow(S.arkusze_w_folderze(ark))
        self.assertEqual([os.path.relpath(p, glowny) for p in wynik],
                         [os.path.join("Etap_5", "232_LA_TH_2026 - protokół_3 CC-04.xlsx"),
                          "231, 232_LA_TH_2026 - protokół CC.xlsx",
                          "232_LA_TH_2026 - protokół CC.xlsx",
                          "232_LA_TH_2026 - protokół_2 CC-04.xlsx"])


class TestWarunkiBloku(unittest.TestCase):

    def test_z_fg_protokolu_bez_czujnika(self):
        blok = {"data": D(2026, 10, 30), "start": "08:00", "koniec": "08:05",
                "fg": [(22.1, 41.0), (22.3, 40.0)]}
        wywolania = []
        self.assertEqual(S.th_bloku(blok, lambda r, m: wywolania.append((r, m)) or []),
                         [(22.1, 41.0), (22.3, 40.0)])
        self.assertEqual(wywolania, [])


class TestOdczytArkusza(unittest.TestCase):

    def test_dane_i_przetwornik(self):
        f = nowa_piaskownica("swiadectwa_odczyt")
        p = os.path.join(f, "a.xlsx")
        arkusz(p, przetwornik=True)
        d = S.odczytaj_arkusz(p)
        self.assertEqual([z["nazwa"] for z in d["zakladki"]], ["30, -", "0, -"])
        self.assertEqual(d["zakladki"][0]["K4_val"], "30.10.2026")
        self.assertEqual((d["E5"], d["E6"], d["G6"]), ("Tempmate", "M1", "TMM1"))
        self.assertEqual(d["przetwornik"]["K15"], "001")
        self.assertEqual(d["przetwornik"]["M22"], 4)

    def test_rekord_z_pz(self):
        class Dev:
            nr_ewid, uwagi = "Q/LOG/21", "Uz = (24 ± 1) V DC."
        dane = {"E5": "Tempmate", "E6": "M1", "G6": "X", "przetwornik": None}
        r = S.rekord_arkusza("234_LA_TH_2026 - ILAJ - 4 - TMM160500533.xlsx", dane,
                             {S.pz_dane.normalizuj_serial("TMM160500533"): Dev()})
        self.assertEqual((r["E"], r["F"], r["B"], r["D"], r["P"]),
                         ("TMM160500533", "Q/LOG/21", "Tempmate", "M1", "Uz = (24 ± 1) V DC."))


if __name__ == "__main__":
    unittest.main(verbosity=2)
