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


def dokument_z_tabelami(liczby_wierszy=(3, 2, 3, 3, 2)):
    from docx import Document
    doc = Document()
    doc.add_paragraph("1) Point ([hist_t] °C, [hist_rh] %rh) was repeated.")
    for n in liczby_wierszy:
        t = doc.add_table(rows=3 + n, cols=3)
        t.rows[2].cells[2].text = "pass/fail"
        for r in t.rows[3:]:
            for c in r.cells:
                c.paragraphs[0].add_run("")
    return doc


def wiersz(arkusz, *dane, zg=("tak",)):
    return {"arkusz": arkusz, "dane": list(dane), "zgodnosc": [A.ZGODNOSC_EN[z] for z in zg]}


class TestSwiadectwo(unittest.TestCase):

    def sekcje(self):
        rh = [wiersz("10, 22", "10.1", "22.2"), wiersz("20, 45", "20.1", "44.6"),
              wiersz("20, 45 (2)", "20.1", "44.4", zg=("nie",))]
        t = [wiersz("-25, -", "-25.1", "-25.3")]
        return [{"rh": rh, "t": t}, {"rh": rh[:2], "t": t}, {"rh": rh, "t": t + t}]

    def test_tabele_i_przypis(self):
        doc = dokument_z_tabelami()
        A._wypelnij_swiadectwo_th_przetw(doc, self.sekcje())
        tabele = A._tabele_wynikow_swiadectwa(doc)
        self.assertEqual([len(t.rows) - 3 for t in tabele], [3, 1, 2, 3, 2])
        w = tabele[0].rows[5]
        self.assertEqual(w.cells[0].text, "20.1")
        self.assertEqual(w.cells[1].text, "   44.4 1)")
        self.assertTrue(w.cells[1].paragraphs[0].runs[-1].font.superscript)
        self.assertEqual(w.cells[2].text, "fail")
        self.assertEqual(tabele[1].rows[3].cells[2].text, "pass")
        self.assertIn("Point (20 °C, 45 %rh)", doc.paragraphs[0].text)

    def test_bez_powtorki_przypis_znika(self):
        doc = dokument_z_tabelami()
        sek = [{"rh": [wiersz("10, 22", "10.1", "22.2")], "t": []}] * 3
        A._wypelnij_swiadectwo_th_przetw(doc, sek)
        self.assertFalse(any("[hist_t]" in p.text for p in doc.paragraphs))

    def test_wiecej_wierszy_niz_w_szablonie(self):
        doc = dokument_z_tabelami((2, 2, 2, 2, 2))
        rh = [wiersz(f"{t}, 50", str(t), "50.0") for t in range(5)]
        A._wypelnij_swiadectwo_th_przetw(doc, [{"rh": rh, "t": []}] * 3)
        t0 = A._tabele_wynikow_swiadectwa(doc)[0]
        self.assertEqual([r.cells[0].text for r in t0.rows[3:]], ["0", "1", "2", "3", "4"])

    def test_daty_po_angielsku(self):
        self.assertEqual(A.data_po_angielsku("05 ÷ 09 listopada 2026 r."), "05 ÷ 09 November 2026")
        self.assertEqual(A.data_po_angielsku("29 kwietnia, 04 maja 2026 r."), "29 April, 04 May 2026")

    def test_placeholdery_wyjscia(self):
        rek = {"P": "", "F": "EV1", "PRZETWORNIK": {"L10": "Prąd", "M22": 4, "O22": 20}}
        p = A.placeholdery_th_przetw(rek, "10 listopada 2026 r.", "")
        self.assertEqual((p["[wyjscie_en]"], p["[symbol_en]"], p["[jedn_wy]"], p["[wy_min]"]),
                         ("current", "I", "mA", "4"))
        rek["PRZETWORNIK"]["L10"] = "Napięcie"
        p = A.placeholdery_th_przetw(rek, "10 listopada 2026 r.", "")
        self.assertEqual((p["[wyjscie_en]"], p["[symbol_en]"], p["[jedn_wy]"], p["[data]"]),
                         ("voltage", "V", "V", "10 November 2026"))


class TestSzablonSwiadectwa(unittest.TestCase):

    def test_zamiana_rozbita_miedzy_runy(self):
        from docx import Document
        p = Document().add_paragraph()
        for t in ("(", "2", "0", ",", "7", "  2", "4", ") °C"):
            p.add_run(t)
        self.assertTrue(SZ.zamien_w_akapicie(p, "20,7", "[temp_min]"))
        self.assertEqual(p.text, "([temp_min]  24) °C")

    def test_opis_przyrzadu(self):
        m = SZ._RE_OBIEKT.search("Temperature and humidity transmitter type: PTU303, serial number: "
                                 "D4850103, identification number: EV363465, manufacturer: VAISALA, "
                                 "year of production: 2008.")
        self.assertEqual(m.group("typ", "nr_fabr", "nr_ewid", "wytworca", "rok_produkcji"),
                         ("PTU303", "D4850103", "EV363465", "VAISALA", "2008"))

    def test_rok_produkcji_z_pz(self):
        import pz_dane
        tekst = ("Calibration objects: 1) Temperature and humidity transmitter type: PTU303, "
                 "serial number: D4850103, identification number: EV363465, manufacturer: "
                 "VAISALA, year of production: 2008. 2) Pressure transmitter type: X, serial "
                 "number: P1, year of production: 2015.")
        self.assertEqual(pz_dane.rok_produkcji(tekst, "D4850103"), "2008")
        self.assertEqual(pz_dane.rok_produkcji(tekst, "P1"), "2015")
        self.assertEqual(pz_dane.rok_produkcji("typ: X, nr fabr.: A1. 2) typ: Y, nr fabr.: B2, "
                                               "rok produkcji: 2010.", "A1"), "")

    def test_rok_produkcji_w_swiadectwie(self):
        rek = {"P": "", "F": "", "PRZETWORNIK": {"L10": "Napięcie"}}
        self.assertEqual(A.placeholdery_th_przetw(rek, "", "", "2008")["[rok_produkcji]"], "2008")
        self.assertEqual(A.placeholdery_th_przetw(rek, "", "")["[rok_produkcji]"], "(UZUPEŁNIJ)")

    def test_zakres_warunkow(self):
        m = SZ._RE_ZAKRES_DWOCH.search("Relative humidity: (29  50) %")
        self.assertEqual(m.groups(), ("29", "50"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
