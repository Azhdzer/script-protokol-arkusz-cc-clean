# -*- coding: utf-8 -*-
"""
Przyrzad wzorcowany z wyswietlacza i z wyjscia analogowego naraz (zlecenie 239,
VAISALA PTU303 D4850103). Na Stronie 2 dwa wiersze z tym samym nr fabrycznym:
jeden z rozdzielczoscia odczytu (wyswietlacz), drugi ze skalowaniem 0-5 V (wyjscie).
Krok 3 robil dwa arkusze, jakby to byly dwa przyrzady. Teraz — jeden arkusz wg
recznie przygotowanego wzoru: E/F zakladek z wyswietlacza, G/H z wyjscia, K10
'TH/Przetwornik', 'Wyniki' z trzema sekcjami, ktore zmieniaja tylko liczbe wierszy.
"""

import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_arkusze as A
import cc_szablon_przetwornika as SZ


def blok(e=None, f=None):
    return {"E_dane": [e] * 5, "F_dane": [f] * 5, "E_fmt": ["0.0"] * 5, "F_fmt": ["0.0"] * 5}


def wyswietlacz(serial="D4850103"):
    return {"E": serial, "F": "EV363465", "K": 0.1, "L": 0.1, "PRZETWORNIK": None}


def wyjscie(serial="D4850103"):
    return {"E": serial, "F": "EV363465", "M": "0-5V (-40 / +60)°C", "P": "ch temp 119",
            "PRZETWORNIK": {"K10": "Przetwornik", "L10": "Napięcie", "K14": "1586A",
                            "K15": 119, "K16": 109, "M22": 0, "O22": 5, "_szablon": "P.xlsx"}}


class TestPolaczenie(unittest.TestCase):

    def test_para_to_jeden_przyrzad(self):
        s2 = [wyswietlacz(), wyjscie(), {"E": "Q/1", "PRZETWORNIK": None}]
        ef = [[blok(20.0, 45.7)], [blok(2.9997, 2.2836)], [blok(21.0, None)]]
        s2n, efn, f24, pary = A._polacz_wyswietlacz_z_wyjsciem(s2, ef, ["BS", "x", "> 4 cm"])
        self.assertEqual([r["E"] for r in s2n], ["D4850103", "Q/1"])
        self.assertEqual(f24, ["BS", "> 4 cm"])
        self.assertEqual(len(pary), 1)
        r = s2n[0]
        self.assertTrue(r["POLACZONY"])
        self.assertEqual(r["_nr_przyrzadu"], 1)
        self.assertEqual(s2n[1]["_nr_przyrzadu"], 3)
        self.assertEqual(r["PRZETWORNIK"]["K10"], "TH/Przetwornik")
        self.assertEqual(r["PRZETWORNIK"]["L10"], "Napięcie")
        self.assertEqual(r["PRZETWORNIK"]["_szablon"], A.SZABLON_TH_PRZETWORNIKA)
        b = efn[0][0]
        self.assertEqual((b["E_dane"][0], b["F_dane"][0]), (20.0, 45.7))
        self.assertEqual((b["G_dane"][0], b["H_dane"][0]), (2.9997, 2.2836))

    def test_wyjscie_przed_wyswietlaczem(self):
        s2n, efn, _f, pary = A._polacz_wyswietlacz_z_wyjsciem(
            [wyjscie(), wyswietlacz()], [[blok(3.0, 2.3)], [blok(20.0, 45.0)]], [None, None])
        self.assertEqual(len(s2n), 1)
        self.assertEqual(s2n[0]["_nr_przyrzadu"], 2)          # numer wiersza wyswietlacza
        self.assertEqual(efn[0][0]["G_dane"][0], 3.0)

    def test_serial_po_normalizacji(self):
        s2n, *_ = A._polacz_wyswietlacz_z_wyjsciem(
            [wyswietlacz("D4850103"), wyjscie(" d4850103 ")], [[blok(1)], [blok(2)]], [None, None])
        self.assertEqual(len(s2n), 1)

    def test_dwa_termohigrometry_bez_wyjscia_zostaja(self):
        s2n, *_ , pary = A._polacz_wyswietlacz_z_wyjsciem(
            [wyswietlacz(), wyswietlacz()], [[blok(1)], [blok(2)]], [None, None])
        self.assertEqual((len(s2n), pary), (2, []))

    def test_sam_przetwornik_zostaje(self):
        s2n, *_ , pary = A._polacz_wyswietlacz_z_wyjsciem([wyjscie()], [[blok(1)]], [None])
        self.assertEqual((len(s2n), pary), (1, []))
        self.assertFalse(s2n[0].get("POLACZONY"))

    def test_blok_tylko_z_danymi_wyjscia_jest_aktywny(self):
        self.assertTrue(A._czy_blok_ef_aktywny({"E_dane": [None] * 5, "F_dane": [None] * 5,
                                                "G_dane": [2.99] + [None] * 4}))


# Uklad recznego wzoru: naglowek, 2 wiersze sekcji wyswietlacza, naglowek, 2 wiersze
# wyjscia (°C/%), naglowek, 2 wiersze wyjscia (V), stopka.
def siatka():
    pusty = [""] * 30

    def w(**kom):
        r = list(pusty)
        for kol, f in kom.items():
            r[SZ._numer(kol) - 1] = f
        return r
    return [
        w(B="Wartość wielkości odniesienia"),
        w(B="='10,22'!D239", I="='10,22'!K239", M="=F2+H2", W="='-25'!D246", Z="='-25'!G246",
          AA="=Y2+Z2"),
        w(B="='10,45'!D239", I="='10,45'!K239", M="=F3+H3", W="='5'!D246", Z="='5'!G246",
          AA="=Y3+Z3"),
        w(D='=IF(\'10,22\'!$L$10="-","-","Zmierzona wartość napięcia")'),
        w(B="='10,22'!D253", K="='10,22'!O253", M="=H5+J5", W="='-25'!D260"),
        w(B="='10,45'!D253", K="='10,45'!O253", M="=H6+J6", W="='5'!D260"),
        w(),
        w(B="='10,22'!D274", W="='-25'!D279", AD="='-25'!K279"),
        w(B="='10,45'!D274", W="='5'!D279", AD="='5'!K279"),
        w(A="ODLEGŁOŚĆ OD CZUJNIKA WZORCOWEGO", J="=ABS(VALUE('20,45'!E219)-VALUE('20,45 (2)'!E219))"),
    ]


class TestSekcjeWynikow(unittest.TestCase):

    def test_trzy_sekcje(self):
        sekcje = A._sekcje_wynikow(siatka())
        self.assertEqual([(s["start"], s["koniec"]) for s in sekcje], [(2, 3), (5, 6), (8, 9)])

    def test_kolumny_wzoru(self):
        kol = A._sekcje_wynikow(siatka())[0]["kolumny"]
        self.assertEqual(kol, [(1, "rh", "D239"), (8, "rh", "K239"), (12, "rh", None),
                               (22, "t", "D246"), (25, "t", "G246"), (26, "t", None)])

    def test_wiersze_zakladek_tylko_t(self):
        kol = A._sekcje_wynikow(siatka())[2]["kolumny"]
        self.assertEqual([(c, r) for c, r, _k in kol], [(1, "rh"), (22, "t"), (29, "t")])

    def test_wiersz_r1c1(self):
        kol = A._sekcje_wynikow(siatka())[0]["kolumny"]
        wzor = [""] * 27
        wzor[12], wzor[26] = "=RC[-7]+RC[-5]", "=RC[-2]+RC[-1]"
        w = A._wiersz_r1c1(wzor, kol, "20, 45 (2)", None, 10)
        self.assertEqual(w[1], "='20, 45 (2)'!R[229]C[2]")      # B10 -> D239
        self.assertEqual(w[8], "='20, 45 (2)'!R[229]C[2]")      # I10 -> K239
        self.assertEqual(w[12], "=RC[-7]+RC[-5]")
        self.assertEqual((w[22], w[25], w[26]), ("", "", ""))     # tabela t bez punktu

    def test_apostrof_w_nazwie(self):
        kol = [(1, "rh", "D239")]
        self.assertEqual(A._wiersz_r1c1(["", ""], kol, "x'y", None, 4)[1], "='x''y'!R[235]C[2]")

    def test_uklad_rozpoznany(self):
        ws = mock.Mock()
        ws.range.return_value.formula = tuple(tuple(w) for w in siatka())
        uklad = A._uklad_wynikow(ws)
        self.assertTrue(uklad.get("polaczony"))
        self.assertEqual(uklad["histereza"], "E219")

    def test_zwykly_szablon_nie_jest_polaczony(self):
        ws = mock.Mock()
        tylko_pierwsza = siatka()[:4]
        ws.range.return_value.formula = tuple(tuple(w) for w in tylko_pierwsza)
        self.assertFalse(A._uklad_wynikow(ws).get("polaczony"))


class TestKolejnoscJakWeWzorze(unittest.TestCase):
    """Zakladki zlecenia 239 -> wiersze jak w recznie przygotowanych 'Wynikach'."""

    ZAKLADKI = ["10, 22", "20, 15", "10, 45", "10, 95", "20, 45", "20, 95", "20, 45 (2)",
                "40, 15", "40, 45", "40, 95", "60, 15", "60, 45", "60, 95", "5, -", "-25, -"]

    def test_kolejnosc(self):
        wpisy = [{"arkusz": n, "val_b": A._t_z_nazwy_zakladki(n),
                  "ma_rh": A._rh_z_nazwy_zakladki(n) is not None,
                  "rh_ref": A._rh_z_nazwy_zakladki(n),
                  "powtorka": A._to_jest_powtorka_kolizji(n)} for n in self.ZAKLADKI]
        rh, t = A._kolejnosc_wierszy_wynikow(wpisy)
        self.assertEqual(rh, ["10, 22", "10, 45", "10, 95", "20, 15", "20, 45", "20, 95",
                              "20, 45 (2)", "40, 15", "40, 45", "40, 95", "60, 15", "60, 45",
                              "60, 95"])
        self.assertEqual(t, ["-25, -", "5, -", "10, 45", "20, 45", "40, 45", "60, 45"])


class TestBudowaSzablonu(unittest.TestCase):

    def test_przemapowanie_arkuszy(self):
        self.assertEqual(SZ.przemapuj_arkusze("='20,45 (2)'!D239+'10,22'!$L$10",
                                              {"20,45 (2)": "23,60"}),
                         "='23,60'!D239+'10,22'!$L$10")

    def test_mapa_w_kolko_po_zostajacych(self):
        f = ["='a'!D1", "='b'!D1", "='x'!D1", "='y'!D1", "='z'!D1", "=IF('a'!L1)"]
        self.assertEqual(SZ.mapa_usuwanych_arkuszy(f, ["a", "b"]), {"x": "a", "y": "b", "z": "a"})

    def test_roznice_stale_wracaja_formuly_zostaja(self):
        wzor = [["239/LA/TH/2026", "=Wyniki!F56", "107.81"], ["", "=SUM(A1)", "TH/Przetwornik"]]
        szab = [["/LA/TH/2026", "=Wyniki!F24", ""], ["", "=SUM(A1)", "Termohigrometr"]]
        self.assertEqual(SZ.roznice_do_wyczyszczenia(wzor, szab),
                         [(1, 1, "/LA/TH/2026"), (1, 3, ""), (2, 3, "Termohigrometr")])


if __name__ == "__main__":
    unittest.main(verbosity=2)
