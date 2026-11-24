# -*- coding: utf-8 -*-
"""
Krok 5 — data wzorcowania z PROTOKOLU, nie z K4 zakladki (zlecenie 243).

Arkusze kroku 5 bywaja skladane recznie: do starego arkusza wkleja sie odczyty z nowego
pomiaru, a K4 zostaje stare (11.11 przy danych z 17.11) albo wpisane jedno dla wszystkich
(18.11) — 17.11 znikal ze swiadectwa. Zakladka trafia teraz na blok protokolu po swoich
odczytach (E/F przyrzadu i Rs wzorca w C), a data i warunki sa z tego bloku. Protokol
'mega' ma wiersze-przekladki miedzy blokami, a Strona 2 bywa bez numerow w kolumnie A.
Liczby zmyslone.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import generuj_arkusze as A
import generuj_swiadectwa as S
from wspolne import nowa_piaskownica

KOL = A.START_COL_E_S3 + A.PRZESUNIECIE_STARTU_KOL_CC04 + A.KROK_COL_EF   # 2. przyrzad


def protokol(sciezka, bloki):
    """bloki: [(T, RH, data, t_przyrzadu, rh_przyrzadu, rs)] albo 'przekladka'."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = A.ARKUSZ_STRONA3
    ws2 = wb.create_sheet(A.ARKUSZ_STRONA2)
    for i, s in enumerate(("111", "84329000")):                 # bez 'Nr' w kolumnie A
        ws2.cell(A.START_ROW_S2 + i, 2).value, ws2.cell(A.START_ROW_S2 + i, 5).value = "Testo", s
    r, nr = A.START_ROW_S3, 0
    for b in bloki:
        if b == "przekladka":
            ws.cell(r, KOL).value = "CC-04-PG"
            r += 1
            continue
        nr += 1
        t, rh, data, tp, rhp, rs = b
        ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value = nr, t, rh
        ws.cell(r + 1, 5).value = data
        for i in range(5):
            ws.cell(r + i, KOL).value, ws.cell(r + i, KOL + 1).value = tp, rhp
            ws.cell(r + i, 12).value = rs + i / 10000
        r += A.BLOK_S3
    wb.save(sciezka)


def zakladka(nazwa, k4, tp, rhp, rs):
    return {"nazwa": nazwa, "K4_val": k4, "E": [tp] * 5, "F": [rhp] * 5,
            "C": [rs + i / 10000 for i in range(5)]}


class TestDataZProtokolu(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        f = nowa_piaskownica("swiadectwa_daty_protokol")
        p = os.path.join(f, "243_LA_TH_2026 - protokół CC-04_mega.xlsx")
        protokol(p, [(70.5, None, "09.11.2026", 70.5, 12.4, 127.17),
                     (25.1, 25.5, "11.11.2026", 25.0, 27.0, 109.60),
                     "przekladka", "przekladka",
                     (25.1, 25.5, "17.11.2026", 25.0, 32.3, 109.70),
                     (25.1, 60.0, "17.11.2026", 25.0, 67.0, 109.71),
                     (25.1, 60.0, "18.11.2026", 25.0, 68.8, 109.71)])
        cls.bloki = S.bloki_protokolu(p)

    def test_bloki_za_przekladkami_i_strona2_bez_numerow(self):
        self.assertEqual(len(self.bloki), 5)
        self.assertEqual(self.bloki[0]["seriale"], {"111", "84329000"})

    def test_data_bloku_o_tych_samych_odczytach_nie_z_k4(self):
        zd = zakladka("25, 25", "18.11.2026", 25.0, 32.3, 109.70)
        b = S.dopasuj_blok(zd, self.bloki, "84329000")
        self.assertEqual((b["nr"], b["data"].day), (3, 17))
        S._data_z_protokolu(zd, b)
        self.assertEqual(zd["K4_val"], "17.11.2026")

    def test_stare_k4_przy_nowych_odczytach(self):
        zd = zakladka("25, 60 (2)", "12.11.2026", 25.0, 68.8, 109.71)
        self.assertEqual(S.dopasuj_blok(zd, self.bloki, "84329000")["data"].day, 18)

    def test_inny_wzorzec_to_inny_blok(self):
        zd = zakladka("25, 25", "11.11.2026", 25.0, 32.3, 109.99)
        self.assertEqual(S.bloki_po_odczytach(zd, self.bloki, "84329000"), [])
        # bez pasujacych odczytow — jak dotad po K4 i nastawie
        self.assertEqual(S.dopasuj_blok(zd, self.bloki, "84329000")["data"].day, 11)

    def test_odczyty_innego_przyrzadu_nie_pasuja(self):
        zd = zakladka("25, 25", "18.11.2026", 25.0, 32.3, 109.70)
        self.assertEqual(S.bloki_po_odczytach(zd, self.bloki, "111"), [])

    def test_ta_sama_data_bez_komunikatu(self):
        zd = zakladka("70, -", "09.11.2026", 70.5, 12.4, 127.17)
        S._data_z_protokolu(zd, S.dopasuj_blok(zd, self.bloki, "84329000"))
        self.assertEqual(zd["K4_val"], "09.11.2026")


class TestOdczytyZakladki(unittest.TestCase):

    def test_odczytaj_arkusz_czyta_c_e_f(self):
        p = os.path.join(nowa_piaskownica("swiadectwa_odczyty_zakladki"), "a.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "25, 25"
        for i in range(5):
            ws.cell(15 + i, 3).value, ws.cell(15 + i, 5).value, ws.cell(15 + i, 6).value = 109.7, 25.0, 32.3
        wb.create_sheet("Wyniki")
        wb.save(p)
        z = S.odczytaj_arkusz(p)["zakladki"][0]
        self.assertEqual((z["C"][0], z["E"][4], z["F"][2]), (109.7, 25.0, 32.3))


if __name__ == "__main__":
    unittest.main(verbosity=2)
