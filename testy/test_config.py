# -*- coding: utf-8 -*-
"""
Testy rejestru ustawien (cc_config.py) — spojnosc definicji, konwersja do env,
trwalosc pliku i odpornosc helperow czytajacych zmienne srodowiskowe.
"""

import datetime
import json
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import cc_config as C
from wspolne import nowa_piaskownica

ZNANE_TYPY = {"tekst", "calk", "liczba", "flaga", "minuty", "plik", "folder",
              "pliki", "kolor", "tabela"}
ZNANE_KROKI = {k for k, _e, _o in C.KROKI}


class TestRejestr(unittest.TestCase):
    """Definicje ustawien musza byc spojne — panel buduje z nich formularze."""

    def test_env_sa_unikalne(self):
        nazwy = [u.env for u in C.USTAWIENIA]
        duplikaty = {n for n in nazwy if nazwy.count(n) > 1}
        self.assertFalse(duplikaty, f"Zdublowane nazwy zmiennych: {duplikaty}")

    def test_pola_opisowe_wypelnione(self):
        for u in C.USTAWIENIA:
            with self.subTest(env=u.env):
                self.assertTrue(u.etykieta.strip(), "pusta etykieta")
                self.assertIn(u.typ, ZNANE_TYPY)
                self.assertIn(u.krok, ZNANE_KROKI)
                self.assertIn(u.poziom, (C.PODSTAWOWY, C.ZAAWANSOWANY))
                self.assertTrue(u.grupa.strip(), "pusta grupa")

    def test_typ_domyslnej_zgodny_z_deklaracja(self):
        oczekiwane = {
            "flaga": bool,
            "calk": int, "minuty": int,
            "liczba": (int, float),
            "tekst": str, "plik": str, "folder": str, "kolor": str,
            "pliki": list, "tabela": list,
        }
        for u in C.USTAWIENIA:
            with self.subTest(env=u.env):
                self.assertIsInstance(u.domyslna, oczekiwane[u.typ])

    def test_typy_wymagajace_dodatkow(self):
        for u in C.USTAWIENIA:
            with self.subTest(env=u.env):
                if u.typ == "plik":
                    self.assertTrue(u.wzorzec, "typ 'plik' wymaga wzorca rozszerzen")
                if u.typ == "tabela":
                    self.assertTrue(u.kolumny, "typ 'tabela' wymaga naglowkow kolumn")
                    for wiersz in u.domyslna:
                        self.assertEqual(len(wiersz), len(u.kolumny))

    def test_zakresy_liczbowe_maja_sens(self):
        for u in C.USTAWIENIA:
            if u.typ in ("calk", "liczba", "minuty") and None not in (u.minimum, u.maksimum):
                with self.subTest(env=u.env):
                    self.assertLess(u.minimum, u.maksimum)
                    self.assertGreaterEqual(u.domyslna, u.minimum)
                    self.assertLessEqual(u.domyslna, u.maksimum)

    def test_kolory_sa_szesnastkowe(self):
        for u in C.USTAWIENIA:
            if u.typ == "kolor":
                with self.subTest(env=u.env):
                    self.assertRegex(u.domyslna, r"^#[0-9A-Fa-f]{6}$")

    def test_kazdy_krok_ma_ustawienia_podstawowe(self):
        for klucz in ("analiza", "obs", "ark"):
            with self.subTest(krok=klucz):
                self.assertTrue(C.dla_kroku(klucz, C.PODSTAWOWY))

    def test_dla_kroku_dzieli_rozlacznie(self):
        for klucz in ZNANE_KROKI:
            podst = C.dla_kroku(klucz, C.PODSTAWOWY)
            zaaw = C.dla_kroku(klucz, C.ZAAWANSOWANY)
            self.assertEqual(len(podst) + len(zaaw), len(C.dla_kroku(klucz)))


class TestEksportDoEnv(unittest.TestCase):
    """Wartosci musza trafiac do podprocesu w formacie, ktory skrypt zrozumie."""

    def test_flaga(self):
        self.assertEqual(C.WG_ENV["GEN_WORD"].do_env(True), "1")
        self.assertEqual(C.WG_ENV["GEN_WORD"].do_env(False), "0")

    def test_lista_plikow_laczona_srednikiem(self):
        u = C.WG_ENV["OBS_TXT_FILES"]
        self.assertEqual(u.do_env(["a.txt", "b.txt"]), "a.txt;b.txt")
        self.assertEqual(u.do_env([]), "")

    def test_tabela_jako_json(self):
        u = C.WG_ENV["GEN_MAP_CC04"]
        odczyt = json.loads(u.do_env(u.domyslna))
        self.assertEqual(odczyt, u.domyslna)

    def test_do_env_obejmuje_caly_rejestr(self):
        env = C.do_env(C.domyslne())
        self.assertEqual(set(env), set(C.WG_ENV))
        for wartosc in env.values():
            self.assertIsInstance(wartosc, str)


class TestTrwalosc(unittest.TestCase):
    """Zapis/odczyt cc_ustawienia.json — w piaskownicy, nie w projekcie."""

    def setUp(self):
        self.folder = nowa_piaskownica("config")
        self.stara = C.PLIK_USTAWIEN
        C.PLIK_USTAWIEN = os.path.join(self.folder, "cc_ustawienia.json")

    def tearDown(self):
        C.PLIK_USTAWIEN = self.stara

    def test_zapis_i_odczyt(self):
        wart = C.domyslne()
        wart["GEN_NR_SW"] = 2222
        wart["OBS_TXT_FILES"] = ["x.txt"]
        self.assertIsNone(C.zapisz(wart))
        odczyt = C.wczytaj()
        self.assertEqual(odczyt["GEN_NR_SW"], 2222)
        self.assertEqual(odczyt["OBS_TXT_FILES"], ["x.txt"])

    def test_brak_pliku_daje_domyslne(self):
        self.assertEqual(C.wczytaj(), C.domyslne())

    def test_uszkodzony_plik_nie_blokuje_startu(self):
        with open(C.PLIK_USTAWIEN, "w", encoding="utf-8") as f:
            f.write("{to nie jest JSON")
        self.assertEqual(C.wczytaj(), C.domyslne())

    def test_nieznane_klucze_sa_pomijane(self):
        with open(C.PLIK_USTAWIEN, "w", encoding="utf-8") as f:
            json.dump({"GEN_NR_SW": 7, "USUNIETE_USTAWIENIE": 1}, f)
        odczyt = C.wczytaj()
        self.assertEqual(odczyt["GEN_NR_SW"], 7)
        self.assertNotIn("USUNIETE_USTAWIENIE", odczyt)

    def test_brakujace_klucze_uzupelniane_domyslnymi(self):
        with open(C.PLIK_USTAWIEN, "w", encoding="utf-8") as f:
            json.dump({"GEN_NR_SW": 7}, f)
        odczyt = C.wczytaj()
        self.assertEqual(set(odczyt), set(C.domyslne()))


class TestHelperyOdczytu(unittest.TestCase):
    """
    Helpery czytajace env. Kluczowa wlasciwosc: smieciowa wartosc NIE moze wywalic
    skryptu w polowie wzorcowania — ma sie cofnac do wartosci domyslnej.
    """

    def setUp(self):
        self.zmienione = []

    def tearDown(self):
        for k in self.zmienione:
            os.environ.pop(k, None)

    def ustaw(self, nazwa, wartosc):
        os.environ[nazwa] = wartosc
        self.zmienione.append(nazwa)

    def test_flaga_rozpoznaje_warianty(self):
        for tekst in ("1", "true", "TAK", "yes", "on", "True"):
            self.ustaw("T_FLAGA", tekst)
            self.assertTrue(C.flaga("T_FLAGA", False), tekst)
        for tekst in ("0", "false", "nie", "off", "cokolwiek"):
            self.ustaw("T_FLAGA", tekst)
            self.assertFalse(C.flaga("T_FLAGA", True), tekst)

    def test_flaga_bez_zmiennej_daje_domyslna(self):
        self.assertTrue(C.flaga("T_BRAK", True))
        self.assertFalse(C.flaga("T_BRAK", False))

    def test_liczba_przyjmuje_przecinek(self):
        self.ustaw("T_L", "3,5")
        self.assertEqual(C.liczba("T_L", 1.0), 3.5)

    def test_liczba_ze_smieci_wraca_do_domyslnej(self):
        self.ustaw("T_L", "abc")
        self.assertEqual(C.liczba("T_L", 1.25), 1.25)

    def test_calk_ucina_czesc_ulamkowa(self):
        self.ustaw("T_C", "7.9")
        self.assertEqual(C.calk("T_C", 1), 7)

    def test_calk_ze_smieci_wraca_do_domyslnej(self):
        self.ustaw("T_C", "-")
        self.assertEqual(C.calk("T_C", 42), 42)

    def test_tekst_przycina_biale_znaki(self):
        self.ustaw("T_T", "  S8000-02  ")
        self.assertEqual(C.tekst("T_T", "x"), "S8000-02")

    def test_pusty_tekst_daje_domyslna(self):
        self.ustaw("T_T", "   ")
        self.assertEqual(C.tekst("T_T", "domyslna"), "domyslna")

    def test_minuty_daja_timedelta(self):
        self.ustaw("T_M", "90")
        self.assertEqual(C.minuty("T_M", datetime.timedelta(hours=2)),
                         datetime.timedelta(minutes=90))

    def test_minuty_ze_smieci_wracaja_do_domyslnej(self):
        self.ustaw("T_M", "duzo")
        dom = datetime.timedelta(hours=2)
        self.assertEqual(C.minuty("T_M", dom), dom)

    def test_lista_dzieli_i_przycina(self):
        self.ustaw("T_LS", " a.xls ; b.xls ;; ")
        self.assertEqual(C.lista("T_LS", []), ["a.xls", "b.xls"])

    def test_tabela_z_json(self):
        self.ustaw("T_TB", '[["LG","P1","P2","3","T"]]')
        self.assertEqual(C.tabela("T_TB", []), [["LG", "P1", "P2", "3", "T"]])

    def test_tabela_ze_zlego_json_wraca_do_domyslnej(self):
        self.ustaw("T_TB", "{nie lista}")
        self.assertEqual(C.tabela("T_TB", [["x"]]), [["x"]])

    def test_sciezka_wzgledna_rozwijana_od_bazy(self):
        self.ustaw("T_S", "wyniki")
        self.assertEqual(C.sciezka("T_S", "inne", r"C:\baza"),
                         os.path.join(r"C:\baza", "wyniki"))

    def test_sciezka_bezwzgledna_zostaje(self):
        self.ustaw("T_S", r"D:\gdzie\indziej")
        self.assertEqual(C.sciezka("T_S", "inne", r"C:\baza"), r"D:\gdzie\indziej")


if __name__ == "__main__":
    unittest.main(verbosity=2)
