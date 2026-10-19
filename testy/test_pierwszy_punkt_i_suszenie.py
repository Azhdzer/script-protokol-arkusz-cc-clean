# -*- coding: utf-8 -*-
"""
Pierwszy punkt pliku i odroznienie suszenia od punktu (zlecenie 201).

Zgloszenie: skrypt pomijal pierwszy punkt zakresu — 23 °C przy nastawie 28 %
(komora dochodzi wtedy do 30 %, wzorzec pokazywal rowne 30,0 %). Podejrzenie
padlo na suszenie, ale przyczyna byla inna:

  1. Okno punktu liczone jest od 2 h stabilizacji po zmianie nastawy. Plik
     zaczyna sie jednak od komory JUZ stojacej na pierwszej nastawie — zapis
     wlacza sie, gdy punkt jest gotowy. Ten pierwszy odcinek trwal 97 min,
     krocej niz 2 h, i w ogole nie stawal sie oknem.

  2. Gdy zamowienie zawiera 23 °C / 30 %, pasuja do niego TRZY odcinki: pierwszy
     punkt, suszenie 23 °C / 28 % po serii 60/85/60 % i postoj 23 °C / 30 % na
     koncu. Ranking po odczycie komory wybieral... postoj — jego nastawa to rowno 30.

Regula rozrozniajaca (wg operatora): niska wilgotnosc ROZPOCZYNAJACA serie przy
danej temperaturze to punkt; niska wilgotnosc PO serii wyzszych wilgotnosci przy
tej samej temperaturze to suszenie przed przejsciem do 10 / 35 °C.
"""

import datetime
import io
import os
import sys
import unittest
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generuj_obserwacje as G

START = datetime.datetime(2026, 10, 15, 13, 44)


def wiersze(odcinki, start=START):
    """
    Wiersze pliku CC: odcinki = [(T, RH_nastawa, RH_odczyt, minut), ...].
    Komora trzyma nastawe temperatury; wzorzec i rozrzuty stabilne.
    """
    rows, t = [], start
    for temp, rh_zad, rh_odcz, minut in odcinki:
        for _ in range(minut):
            rows.append([t.strftime("%Y-%m-%d %H:%M:%S"), str(temp), str(rh_zad),
                         str(temp), str(rh_odcz), "108.6", "5.1", str(temp),
                         str(rh_odcz), str(temp), "0.05", "0.02"])
            t += datetime.timedelta(minutes=1)
    return rows


def okna(odcinki):
    data = G._prepare_data(wiersze(odcinki))
    with redirect_stdout(io.StringIO()):
        w = G._find_all_analysis_windows(data, G.STABILIZACJA_MIN)
    return data, w


def nastawy(data, w):
    return [(data[si][1], data[si][2]) for si, _ei in w]


# Przebieg z pliku 201 (skrocony do ksztaltu): punkt 30 % na starcie,
# seria 60/85/60, suszenie 28 %, 10 °C, 35 °C, postoj 30 %.
PRZEBIEG_201 = [
    (23.0, 28.0, 28.0, 97),     # pierwszy punkt — tylko 97 min od startu pliku
    (22.9, 58.5, 58.5, 180),
    (22.9, 82.5, 82.5, 210),
    (22.9, 58.5, 58.5, 180),
    (23.0, 28.0, 28.0, 360),    # suszenie
    (9.8, 59.0, 59.0, 210),
    (35.0, 57.5, 57.5, 210),
    (23.0, 30.0, 45.0, 600),    # postoj na koncu
]


class TestPierwszyOdcinekPliku(unittest.TestCase):

    def test_krotki_pierwszy_odcinek_jest_oknem(self):
        data, w = okna(PRZEBIEG_201)
        self.assertEqual(nastawy(data, w)[0], (23.0, 28.0))

    def test_okno_na_koncu_odcinka(self):
        """Reprezentanci z ogona punktu, jak w kazdym innym punkcie."""
        data, w = okna(PRZEBIEG_201)
        si, ei = w[0]
        self.assertEqual(data[ei - 1][0] - data[si][0], G.MIN_OKNO_ANALIZY)

    def test_ten_sam_krotki_odcinek_w_srodku_pliku_odpada(self):
        """W srodku pomiaru 97 min po zmianie nastawy to za malo — jak dotad."""
        odcinki = [(22.9, 58.5, 58.5, 180), (23.0, 28.0, 28.0, 97), (22.9, 82.5, 82.5, 210)]
        data, w = okna(odcinki)
        self.assertNotIn((23.0, 28.0), nastawy(data, w))

    def test_pierwszy_odcinek_w_trakcie_dochodzenia_odpada(self):
        """Komora dopiero dochodzi do nastawy na starcie — to nie gotowy punkt."""
        odcinki = [(23.0, 28.0, 45.0, 97), (22.9, 58.5, 58.5, 180)]
        data, w = okna(odcinki)
        self.assertNotIn((23.0, 28.0), nastawy(data, w))

    def test_dlugi_pierwszy_odcinek_bez_zmian(self):
        odcinki = [(23.0, 28.0, 28.0, 180), (22.9, 58.5, 58.5, 180)]
        data, w = okna(odcinki)
        self.assertEqual(nastawy(data, w)[0], (23.0, 28.0))


class TestSuszenieCzyPunkt(unittest.TestCase):

    def test_niska_wilgotnosc_na_poczatku_serii_to_punkt(self):
        data, w = okna(PRZEBIEG_201)
        self.assertNotIn(0, G.okna_suszenia(data, w))

    def test_niska_wilgotnosc_po_serii_to_suszenie(self):
        data, w = okna(PRZEBIEG_201)
        idx = nastawy(data, w).index((23.0, 28.0), 1)     # drugie wystapienie
        self.assertIn(idx, G.okna_suszenia(data, w))

    def test_postoj_na_koncu_to_suszenie(self):
        data, w = okna(PRZEBIEG_201)
        self.assertIn(nastawy(data, w).index((23.0, 30.0)), G.okna_suszenia(data, w))

    def test_seria_przy_innej_temperaturze_nie_przesadza(self):
        """23 °C / 30 % po serii przy 35 °C — to start serii 23 °C, nie suszenie."""
        odcinki = [(35.0, 57.5, 57.5, 210), (23.0, 30.0, 30.0, 210), (22.9, 58.5, 58.5, 180)]
        data, w = okna(odcinki)
        self.assertEqual(G.okna_suszenia(data, w), set())

    def test_punkt_poza_strefa_suszenia_nigdy(self):
        """10 °C / 30 % po 10 °C / 60 % — temperatura poza strefa suszenia."""
        odcinki = [(9.8, 59.0, 59.0, 210), (9.8, 30.0, 30.0, 210)]
        data, w = okna(odcinki)
        self.assertEqual(G.okna_suszenia(data, w), set())


class TestWyborWgPZ(unittest.TestCase):

    def wybierz(self, punkty):
        data, w = okna(PRZEBIEG_201)
        with redirect_stdout(io.StringIO()) as buf:
            wybrane = G._wybierz_okna_wg_pz(data, w, punkty)
        self.log = buf.getvalue()
        return data, w, wybrane

    def test_punkt_23_30_trafia_w_pierwszy_odcinek(self):
        """Ani suszenie 28 %, ani postoj 30 % na koncu — pierwszy punkt pliku."""
        data, w, wybrane = self.wybierz([(23, 30), (23, 60), (23, 85), (23, 60)])
        okno_30 = next(o for o, p in wybrane if p == (23, 30))
        self.assertEqual(okno_30, w[0])

    def test_suszenie_i_postoj_spoza_zamowienia(self):
        _d, w, wybrane = self.wybierz([(23, 30), (23, 60), (23, 85), (23, 60)])
        uzyte = {o for o, _p in wybrane}
        self.assertEqual(len(uzyte), 4)
        self.assertNotIn(w[4], uzyte)          # suszenie
        self.assertNotIn(w[-1], uzyte)         # postoj

    def test_gdy_pasuje_tylko_suszenie_bierze_je_z_ostrzezeniem(self):
        """Zamowiony punkt nie moze zniknac po cichu."""
        odcinki = [(22.9, 58.5, 58.5, 180), (23.0, 28.0, 28.0, 360)]
        data, w = okna(odcinki)
        with redirect_stdout(io.StringIO()) as buf:
            wybrane = G._wybierz_okna_wg_pz(data, w, [(23, 30)])
        self.assertEqual(len(wybrane), 1)
        self.assertIn("suszenie", buf.getvalue())


if __name__ == "__main__":
    unittest.main(verbosity=2)
