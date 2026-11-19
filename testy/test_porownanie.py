# -*- coding: utf-8 -*-
"""
Krok 6 — arkusz porownawczy z protokolu z przekladkami (zlecenie 243, protokol CC-04_1).

Uklad testowego protokolu (zmyslone liczby, uklad jak w oryginale):

    wiersz 14   czujnik wzorcowy:   P1 CC-04-LG   P2 CC-04-PD   P3 -
    wiersz 17   odleglosc:          ≤ 4 cm
    blok 1  (20)  70,5 / -      P1 t zielona, RH szara  | P2 t szara        | P3 odczyt bez czujnika
    blok 2  (25)  25 / 60       P1 zielona             | P2 t i RH szare   | P3 -
    przekladka (30/31)          P1 CC-04-PG / ≤ 2 cm   | P2 -              | P3 CC-04-LD
    blok 3  (32)  70,2 / -      P1 zielona             | P2 -              | P3 zielona
    blok 4  (37)  25,1 / 59,5   P1 zielona             | P2 'Przestal rejestrowac'
    (42)  'minimum:'

Krok 3 czyta bloki co 5 wierszy od 20 — na przekladce (wiersz 30) konczyl.
"""

import io
import os
import sys
import unittest
from contextlib import redirect_stdout
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color

import generuj_arkusze as G
import generuj_porownanie as P
from wspolne import nowa_piaskownica

ZIELONY = PatternFill("solid", fgColor="CCFFCC")
SZARY_RGB = PatternFill("solid", fgColor="BFBFBF")
# 'Tlo 1, ciemniejsze 25%' — tak zapisane sa szare komorki w czesci protokolu.
SZARY_MOTYW = PatternFill("solid", fgColor=Color(theme=0, tint=-0.249977111117893))

KOL = {1: 19, 2: 21, 3: 23}          # S, U, W — kolumna t przyrzadu (RH = +1)
RS = {"K": 127.1, "L": 127.2, "M": 127.3, "N": 127.4}


def _blok(ws, r0, nr, tk, rhk, przyrzady):
    """przyrzady: {nr: (t, rh, wypelnienie_t, wypelnienie_rh)}."""
    ws.cell(r0, 1, nr)
    ws.cell(r0, 2, tk)
    ws.cell(r0, 3, rhk)
    ws.cell(r0 + 1, 5, "09.11.2026")
    ws.cell(r0 + 2, 5, "20:23")
    ws.cell(r0 + 3, 5, "20:27")
    for i in range(5):
        ws.cell(r0 + i, 8, i + 1)
        for k, (lit, v) in enumerate(RS.items()):
            ws.cell(r0 + i, 11 + k, round(v + r0 / 100 + i / 1000, 4))
        ws.cell(r0 + i, 15, 3.6)
        for p, (t, rh, f_t, f_rh) in przyrzady.items():
            c = ws.cell(r0 + i, KOL[p], t)
            c.fill = f_t
            c = ws.cell(r0 + i, KOL[p] + 1, rh)
            c.fill = f_rh


def protokol(nazwa="999_LA_TH_2026 - protokół CC-04_1.xlsx", seriale=("11111111", "22222222", "33333333")):
    folder = nowa_piaskownica("porownanie_" + Path(nazwa).stem.replace(" ", "_")[:20])
    wb = openpyxl.Workbook()
    wb.active.title = "Strona 1"
    wb["Strona 1"]["I53"] = "-"
    s2 = wb.create_sheet("Strona 2")
    for j, ser in enumerate(seriale):
        r = 11 + j
        for c, v in ((2, "Testo"), (4, "174H"), (5, ser), (6, f"EW-{j + 1}"), (11, 0.1),
                     (12, 0.1), (13, "-"), (14, "-"), (15, "999"), (16, "-")):
            s2.cell(r, c, v)
    s3 = wb.create_sheet("Strona 3")
    for p, tag in ((1, "CC-04-LG"), (2, "CC-04-PD"), (3, "-")):
        s3.cell(14, KOL[p], tag)
        s3.cell(17, KOL[p], "≤ 4 cm")
        s3.cell(19, KOL[p], "°C")
        s3.cell(19, KOL[p] + 1, "%")
    _blok(s3, 20, 1, 70.5, "-", {1: (70.4, 9.4, ZIELONY, SZARY_RGB),
                                  2: (70.6, 9.7, SZARY_MOTYW, SZARY_MOTYW),
                                  3: (70.3, 8.0, ZIELONY, SZARY_RGB)})
    _blok(s3, 25, 2, 25, 60, {1: (24.8, 60.5, ZIELONY, ZIELONY),
                              2: (24.9, 61.0, SZARY_RGB, SZARY_MOTYW)})
    for p, tag, odl in ((1, "CC-04-PG", "≤ 2 cm"), (2, "-", "-"), (3, "CC-04-LD", "≤ 4 cm")):
        s3.cell(30, KOL[p], tag)
        s3.cell(31, KOL[p], odl)
    _blok(s3, 32, 3, 70.2, "-", {1: (70.1, 28.0, ZIELONY, SZARY_RGB),
                                  3: (70.0, 25.0, ZIELONY, SZARY_RGB)})
    _blok(s3, 37, 4, 25.1, 59.5, {1: (24.9, 59.9, ZIELONY, ZIELONY),
                                   2: ("Przestał rejestrowac", None, ZIELONY, ZIELONY)})
    s3.cell(42, 1, "minimum:")
    sciezka = os.path.join(folder, nazwa)
    wb.save(sciezka)
    return sciezka


def czytaj(sciezka, **kw):
    with redirect_stdout(io.StringIO()) as buf:
        wynik = P.wczytaj_protokol(sciezka, **kw)
    return wynik, buf.getvalue()


class TestUklad(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.sciezka = protokol()
        (cls.s2, cls.zakl, cls.ef, cls.f24, _mm), cls.log = czytaj(cls.sciezka)

    def test_bloki_za_przekladka_sa_czytane(self):
        self.assertEqual([z["wiersz_s3"] for z in self.zakl], [20, 25, 32, 37])

    def test_uklad_bloki_i_przekladki(self):
        wb = openpyxl.load_workbook(self.sciezka)
        uklad = P.bloki_i_przekladki(wb["Strona 3"])
        self.assertEqual([r for rodzaj, r in uklad if rodzaj == "blok"], [20, 25, 32, 37])
        self.assertEqual([r for rodzaj, r in uklad if rodzaj == "przekladka"], [30, 31])

    def test_czujnik_zmienia_sie_na_przekladce(self):
        p1 = self.ef[0]
        self.assertEqual([b.get("CC04_TAG") for b in p1], ["LG", "LG", "PG", "PG"])
        self.assertEqual(p1[0]["C_dane"][0], 127.1 + 0.20)            # LG -> kol. K
        self.assertAlmostEqual(p1[2]["C_dane"][0], 127.2 + 0.32)      # PG -> kol. L
        self.assertEqual(p1[2]["ODLEGLOSC"], "≤ 2 cm")

    def test_minus_w_przekladce_nie_zmienia_czujnika(self):
        self.assertEqual(self.ef[1][3].get("CC04_TAG"), "PD")

    def test_odczyt_bez_czujnika_pomijany_z_ostrzezeniem(self):
        self.assertEqual(self.ef[2][0], {})
        self.assertIn("brak czujnika", self.log)
        self.assertEqual(self.ef[2][2]["CC04_TAG"], "LD")
        self.assertAlmostEqual(self.ef[2][2]["C_dane"][0], 127.3 + 0.32)  # LD -> kol. M

    def test_rekord_czujnik_i_odleglosc_z_pierwszego_bloku(self):
        self.assertEqual(self.s2[2]["CC04_TAG"], "LD")
        self.assertEqual(self.f24, ["≤ 4 cm", "≤ 4 cm", "≤ 4 cm"])
        self.assertEqual([r["_nr_przyrzadu"] for r in self.s2], [1, 2, 3])

    def test_tekst_zamiast_odczytu_pominiety(self):
        self.assertFalse(G._czy_blok_ef_aktywny(self.ef[1][3]))
        self.assertIn("Przestał rejestrowac", self.log)


class TestSzareKomorki(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.sciezka = protokol("998_LA_TH_2026 - protokół CC-04_1.xlsx")
        (_s2, _z, cls.ef, _f, _m), _ = czytaj(cls.sciezka)
        (_s2, _z, cls.ef_bez, _f, _m), _ = czytaj(cls.sciezka, uwzglednij_szare=False)

    def test_kolor_motywu_z_odcieniem(self):
        self.assertEqual(P.z_odcieniem("FFFFFF", -0.249977111117893), "#BFBFBF")

    def test_szara_rh_w_punkcie_samej_temperatury_pomijana(self):
        b = self.ef[0][0]
        self.assertEqual(b["E_dane"], [70.4] * 5)
        self.assertEqual(b["F_dane"], [None] * 5)
        self.assertFalse(b["SZARY"])

    def test_szara_temperatura_brana_do_porownania(self):
        b = self.ef[1][0]
        self.assertEqual(b["E_dane"], [70.6] * 5)
        self.assertEqual(b["F_dane"], [None] * 5)     # punkt '70, -' zostaje temperaturowy
        self.assertTrue(b["SZARY"])

    def test_caly_szary_punkt_rh_brany_razem_z_wilgotnoscia(self):
        b = self.ef[1][1]
        self.assertEqual(b["F_dane"], [61.0] * 5)
        self.assertTrue(b["SZARY"])

    def test_bez_szarych(self):
        self.assertFalse(G._czy_blok_ef_aktywny(self.ef_bez[1][0]))
        self.assertFalse(G._czy_blok_ef_aktywny(self.ef_bez[1][1]))
        self.assertEqual(self.ef_bez[0][0]["E_dane"], [70.4] * 5)


class TestNazwyIPolaczenie(unittest.TestCase):

    def test_powtorzenia_maja_nazwe_z_pz_i_sufiks(self):
        (_s2, zakl, ef, _f, _m), _ = czytaj(protokol("997_LA_TH_2026 - protokół CC-04_1.xlsx"))
        with redirect_stdout(io.StringIO()):
            P.nazwy_zakladek_porownania(zakl, punkty=[(70, None), (25, 60)])
        self.assertEqual([z["nazwa"] for z in zakl], ["70, -", "25, 60", "70, -", "25, 60"])
        aktywne = [z for z, b in zip(zakl, ef[0]) if G._czy_blok_ef_aktywny(b)]
        self.assertEqual(G._unikalne_nazwy_zakladek(aktywne),
                         ["70, -", "25, 60", "70, - (2)", "25, 60 (2)"])

    def test_bez_pasujacego_punktu_nazwa_z_nastawy(self):
        (_s2, zakl, _e, _f, _m), _ = czytaj(protokol("996_LA_TH_2026 - protokół CC-04_1.xlsx"))
        with redirect_stdout(io.StringIO()):
            P.nazwy_zakladek_porownania(zakl, punkty=[(40, 25)])
        self.assertEqual(zakl[0]["nazwa"], "71, -")

    def test_dwa_protokoly_laczone_po_numerze_fabrycznym(self):
        a, _ = czytaj(protokol("995_LA_TH_2026 - protokół CC-04_1.xlsx"))
        b, _ = czytaj(protokol("994_LA_TH_2026 - protokół CC-04_1.xlsx",
                               seriale=("22222222", "44444444", "55555555")))
        rek, zakl, ef, f24 = P.scal_protokoly([a[:4], b[:4]])
        self.assertEqual([r["E"] for r in rek],
                         ["11111111", "22222222", "33333333", "44444444", "55555555"])
        self.assertEqual(len(zakl), 8)
        self.assertTrue(all(len(e) == 8 for e in ef))
        # 22222222: bloki z pierwszego protokolu (jako P2) i z drugiego (jako P1)
        self.assertTrue(G._czy_blok_ef_aktywny(ef[1][0]))
        self.assertTrue(G._czy_blok_ef_aktywny(ef[1][4]))
        # 11111111 nie ma go w drugim protokole — puste bloki
        self.assertFalse(any(G._czy_blok_ef_aktywny(e) for e in ef[0][4:]))
        self.assertEqual(len(f24), 5)


class TestWyborPrzyrzadow(unittest.TestCase):

    REK = [{"_nr_przyrzadu": i, "E": f"8418{i:04d}", "F": f"21{i:02d}-DW-T"} for i in range(1, 8)]

    def test_puste_wszystkie(self):
        self.assertEqual(P.wybrane_przyrzady(self.REK, ""), list(range(7)))

    def test_numery_i_zakres(self):
        self.assertEqual(P.wybrane_przyrzady(self.REK, "2, 5-6"), [1, 4, 5])

    def test_numer_fabryczny_i_ewidencyjny(self):
        self.assertEqual(P.wybrane_przyrzady(self.REK, "84180003; 2107-DW-T"), [2, 6])


class TestFolderKroku6(unittest.TestCase):

    def test_protokol_najpierw_z_folderu_kroku_6(self):
        folder = nowa_piaskownica("porownanie_folder")
        krok6 = os.path.join(folder, "Etap_6_porownanie")
        os.makedirs(krok6)
        for f in (os.path.join(krok6, "a.xlsx"), os.path.join(folder, "a.xlsx"),
                  os.path.join(folder, "b.xlsx")):
            open(f, "w").close()
        from unittest import mock
        with mock.patch.object(P, "FOLDER", folder), mock.patch.object(P, "FOLDER_WYJ", krok6):
            self.assertEqual(P.sciezka_protokolu("a.xlsx"), os.path.join(krok6, "a.xlsx"))
            self.assertEqual(P.sciezka_protokolu("b.xlsx"), os.path.join(folder, "b.xlsx"))
            self.assertEqual(P.sciezka_protokolu("c.xlsx"), os.path.join(krok6, "c.xlsx"))

    def test_stary_folder_wynikow_przechodzi_na_folder_kroku(self):
        import json
        from unittest import mock
        import cc_config as C
        folder = nowa_piaskownica("porownanie_ustawienia")
        plik = os.path.join(folder, "cc_ustawienia.json")
        with open(plik, "w", encoding="utf-8") as f:
            json.dump({"POR_FOLDER": "porownanie"}, f)
        with mock.patch.object(C, "PLIK_USTAWIEN", plik):
            self.assertEqual(C.wczytaj()["POR_FOLDER"], "Etap_6_porownanie")
        with open(plik, "w", encoding="utf-8") as f:
            json.dump({"POR_FOLDER": "moj_folder"}, f)
        with mock.patch.object(C, "PLIK_USTAWIEN", plik):
            self.assertEqual(C.wczytaj()["POR_FOLDER"], "moj_folder")


if __name__ == "__main__":
    unittest.main(verbosity=2)
