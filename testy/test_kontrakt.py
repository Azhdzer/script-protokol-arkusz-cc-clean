# -*- coding: utf-8 -*-
"""
Testy KONTRAKTU panel ↔ skrypty.

To jest najwazniejszy plik w tym katalogu. Sprawdza, ze kazde ustawienie widoczne
w panelu naprawde dociera do skryptu i zmienia wlasciwa stala. Poprzedni panel
psul sie dokladnie tutaj: pokazywal opcje, ktorych skrypt nigdy nie czytal, wiec
uzytkownik "ustawial" cos, co nie mialo zadnego wplywu na wynik.

Dla kazdego ustawienia sprawdzamy dwie rzeczy:
  1) bez zmiennej srodowiskowej stala ma wartosc DOMYSLNA (reczne uruchomienie
     skryptu dziala jak dawniej),
  2) z ustawiona zmienna stala przyjmuje NOWA wartosc (panel realnie steruje).

Kazdy modul importowany jest w osobnym procesie — stale konfiguracyjne wyliczaja
sie raz, przy imporcie.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import cc_config as C
from wspolne import KORZEN, wartosci_z_modulu

# Staly, sztuczny folder roboczy — dzieki niemu sciezki wyliczane ze zmiennej
# CC_FOLDER sa przewidywalne i niezalezne od maszyny.
BAZA = r"C:\baza_testowa"

# (env, wyrazenie na module, wartosc testowa, oczekiwany str po podstawieniu,
#  oczekiwany str przy braku zmiennej = wartosc domyslna skryptu)
KONTRAKT = {
    "analizuj_excele": [
        ("ANL_INPUT",  "m.DEFAULT_INPUT",  "logi_dut",
         os.path.join(BAZA, "logi_dut"), os.path.join(BAZA, "excel_do_analizy")),
        ("ANL_OUTPUT", "m.DEFAULT_OUTPUT", "wyjscie",
         os.path.join(BAZA, "wyjscie"), os.path.join(BAZA, "wyniki")),
        ("ANL_DEBUG",  "m.DEBUG", "1", "True", "False"),
        ("ANL_PLIKI",  "m.WYBRANE_PLIKI", "a.xls;b.xls",
         "['a.xls', 'b.xls']", "[]"),
    ],
    "generuj_obserwacje": [
        # --- pliki i podpisy ---
        ("OBS_TEMPLATE",      "m.TEMPLATE", "T_obs.xlsx", "T_obs.xlsx",
         "xxx_LA_TH_2026 - obserwacje CC.xlsx"),
        ("OBS_CC04_TEMPLATE", "m.CC04_TEMPLATE", "T_cc04.xlsx", "T_cc04.xlsx",
         "szablon_LA_TH_2026 - obserwacje.xlsx"),
        ("OBS_PROT_CC",       "m.PROTOKOL_CC_TEMPLATE", "P_cc.xlsx", "P_cc.xlsx",
         "xxx_LA_TH_2026 - protokół CC.xlsx"),
        ("OBS_PROT_CC04",     "m.PROTOKOL_CC04_TEMPLATE", "P_cc04.xlsx", "P_cc04.xlsx",
         "xxx_LA_TH_2026 - protokół CC-04.xlsx"),
        ("OBS_PODPIS",        "m.PODPIS", "Jan Kowalski", "Jan Kowalski",
         "Artsiom Azhdzer"),
        ("OBS_PODPIS_SPR",    "m.PODPIS_SPRAWDZIL", "Anna Nowak", "Anna Nowak",
         "Marek Szpakowski"),
        # --- okno analizy ---
        ("OBS_STAB_MIN",   "m.STABILIZACJA_MIN",   "45", "0:45:00", "2:00:00"),
        ("OBS_STAB_PO_RH", "m.STABILIZACJA_PO_RH", "75", "1:15:00", "2:00:00"),
        ("OBS_PROG_T",     "m.PROG_WEJSCIA_TEMP",  "0.9", "0.9", "0.4"),
        ("OBS_PROG_RH",    "m.PROG_WEJSCIA_RH_PROC", "6.5", "6.5", "3.0"),
        ("OBS_MIN_OKNO",   "m.MIN_OKNO_ANALIZY",   "25", "0:25:00", "0:15:00"),
        ("OBS_ODSTEP_KONIEC", "m.ODSTEP_OD_KONCA_PUNKTU", "5", "0:05:00", "0:02:00"),
        # --- suszenie ---
        ("OBS_SUSZ_T_MIN",  "m.SUSZENIE_T_ZAKRES[0]", "18.5", "18.5", "21.0"),
        ("OBS_SUSZ_T_MAX",  "m.SUSZENIE_T_ZAKRES[1]", "31.5", "31.5", "27.0"),
        ("OBS_SUSZ_RH_MAX", "m.SUSZENIE_RH_MAX",      "44.0", "44.0", "50.0"),
        # --- punkty z PZ ---
        ("OBS_PZ_PUNKTY",    "m.WYBIERAJ_PUNKTY_WG_PZ", "0", "False", "True"),
        ("OBS_TOL_PUNKT_T",  "m.TOL_PUNKT_T",  "2.5", "2.5", "1.5"),
        ("OBS_TOL_PUNKT_RH", "m.TOL_PUNKT_RH", "6.0", "6.0", "4.0"),
        # --- dopasowanie wynikow ---
        ("OBS_TOL",                "m.WYNIKI_TOLERANCJA_MIN",   "0.5", "0.5", "3.0"),
        ("OBS_POMIJAJ_PUSTE_KOL",  "m.POMIJAJ_PUSTE_KOLUMNY",   "0", "False", "True"),
        ("OBS_MAX_ROZN_PRZYRZAD",  "m.MAX_ROZNICA_PRZYRZAD_C",  "8.0", "8.0", "5.0"),
        ("OBS_KOREKTA_ZEGARA",     "m.KOREKTA_ZEGARA",          "0", "False", "True"),
        ("OBS_KZ_MAX",             "m.KOREKTA_ZEGARA_MAX_MIN",  "240", "240", "360"),
        ("OBS_KZ_KROK",            "m.KOREKTA_ZEGARA_KROK_MIN", "10", "10", "5"),
        # --- filtr nastawa/odczyt ---
        ("OBS_FILTR",      "m.FILTR_NASTAWA_ODCZYT", "0", "False", "True"),
        ("OBS_PROG",       "m.MAX_ROZNICA_PROCENT",  "17.5", "17.5", "10.0"),
        ("OBS_TOL_ABS_T",  "m.TOL_ABS_TEMP",         "2.0", "2.0", "1.0"),
        ("OBS_TOL_ABS_RH", "m.TOL_ABS_RH",           "3.5", "3.5", "2.0"),
        # --- zdjecia ---
        ("OBS_FOTO",        "m.KOPIUJ_FOTO",     "1", "True", "False"),
        ("OBS_FOTO_ZRODLO", "m.FOTO_ZRODLO",     r"D:\zdjecia", r"D:\zdjecia",
         r"\\83b\Zdjęcia"),
        ("OBS_FOTO_TOL",    "m.FOTO_TOLERANCJA", "4", "0:04:00", "0:01:00"),
        ("OBS_FOTO_CEL",    "m.FOTO_FOLDER", "zdjecia_188",
         os.path.join(BAZA, "zdjecia_188"), os.path.join(BAZA, "foto")),
        # --- sciezki wspolne ---
        ("CC_PZ_FOLDER",   "m.PZ_FOLDER", "PZ_inne",
         os.path.join(BAZA, "PZ_inne"), os.path.join(BAZA, "PZ")),
        ("CC_ZESTAWIENIE", "m.ZESTAWIENIE_PLIK", "Zest2.xlsx",
         os.path.join(BAZA, "Zest2.xlsx"),
         os.path.join(BAZA, "Zestawienie wzorcowanych przyrządów.xlsx")),
        ("ANL_OUTPUT",     "m.WYNIKI_FOLDER", "wyjscie",
         os.path.join(BAZA, "wyjscie"), os.path.join(BAZA, "wyniki")),
    ],
    "generuj_arkusze": [
        # --- dane wejsciowe ---
        ("CC_PROTOKOL", "m.PROTOKOL_PLIK", "prot_X.xlsx", "prot_X.xlsx", None),
        ("CC_SZABLON",  "m.SZABLON_PLIK",  "wzor_X.xlsx", "wzor_X.xlsx", None),
        # --- naglowki i podpisy ---
        ("GEN_PODPIS_1",   "m.PODPISUJACY_1", "Jan Kowalski", "Jan Kowalski",
         "Artsiom Azhdzer"),
        ("GEN_PODPIS_2",   "m.PODPISUJACY_2", "Anna Nowak", "Anna Nowak",
         "Marek Szpakowski"),
        ("GEN_K18_CC",     "m.HIGROMETR_K18_WG_KOMORY['CC']", "OPTIDEW",
         "OPTIDEW", "S8000-02"),
        ("GEN_K18_CC04",   "m.HIGROMETR_K18_WG_KOMORY['CC-04']", "S8000-07",
         "S8000-07", "S8000"),
        ("GEN_K18_DOM",    "m.HIGROMETR_K18_DOMYSLNY", "BRAK", "BRAK", "S8000"),
        ("GEN_NR_SW",      "m.NR_SW_POCZATKOWY", "2500", "2500", "1047"),
        ("GEN_NR_POM",     "m.NR_POMIESZCZENIA", "12", "12", "9"),
        ("GEN_MODEL_CZUJ", "m.MODEL_CZUJNIKA", "MX1102", "MX1102", "MX1101-02"),
        # --- szablony Word ---
        ("GEN_WORD_TEMP", "m.SZABLON_WORD_TYLKO_TEMP", "W_temp.docx", "W_temp.docx",
         "xxx_yyy_LA_TH_2026 - tylko temp.docx"),
        ("GEN_WORD_RH",   "m.SZABLON_WORD_Z_RH", "W_rh.docx", "W_rh.docx",
         "xxx_yyy_LA_TH_2026 - zakres.docx"),
        ("GEN_WORD_MIX",  "m.SZABLON_WORD_MIESZANY", "W_mix.docx", "W_mix.docx",
         "xxx_yyy_LA_TH_2026 - zakres + temp.docx"),
        # --- etapy ---
        ("GEN_EXCEL", "m.GENERUJ_EXCEL", "0", "False", "True"),
        ("GEN_WORD",  "m.GENERUJ_WORD",  "0", "False", "True"),
        ("GEN_PUSTE", "m.USUWAJ_PUSTE_BLOKI_KOPII_S3", "0", "False", "True"),
        ("GEN_POMIJAJ_PUSTE", "m.POMIJAJ_PRZYRZADY_BEZ_DANYCH", "0", "False", "True"),
        # --- pliki linkowane ---
        ("GEN_LINKOWANE", "m.PLIKI_LINKOWANE", "A.xls;B.xls",
         "['A.xls', 'B.xls']", "['Obliczenia tdp, RH, C.xls', 'Wzory.xls']"),
        ("GEN_LINK_OBLICZENIA", "m.LINKI_SERWEROWE['Obliczenia tdp, RH, C.xls']",
         r"\\serwer\Obl.xls", r"\\serwer\Obl.xls",
         r"\\plum4\LabPomiarowe\Obliczenia tdp, RH, C.xls"),
        ("GEN_LINK_WZORY", "m.LINKI_SERWEROWE['Wzory.xls']",
         r"\\serwer\Wzory.xls", r"\\serwer\Wzory.xls",
         r"\\plum4\LabPomiarowe\Wzory.xls"),
        # --- stabilnosc Excela / wyglad ---
        ("GEN_AUTOREC",     "m.CZYSC_AUTORECOVER", "0", "False", "True"),
        ("GEN_AUTOREC_DNI", "m.CZYSC_AUTORECOVER_STARSZE_NIZ_DNI", "5", "5", "1"),
        ("GEN_PROG_OSTRZ",  "m.PROG_OSTRZEZENIA_KOPII", "25", "25", "10"),
        ("GEN_TAB_RATIO",   "m.TAB_RATIO", "0.5", "0.5", "0.85"),
        ("GEN_TOL_CZUJ",    "m.TOLERANCJA_CZUJNIKA_MIN", "45.5", "45.5", "2.0"),
        # --- filtr kolorow Strony 3 ---
        ("GEN_FILTR_KOLOR", "m.FILTRUJ_KOLOR_S3", "0", "False", "True"),
        ("GEN_KOLOR_AKT",   "m.KOLOR_AKTYWNY_S3", "#00FF00", "#00FF00", "#CCFFCC"),
        ("GEN_KOLOR_POM",   "m.KOLOR_POMIJANY_S3", "#808080", "#808080", "#BFBFBF"),
        ("GEN_INNE_KOLORY", "m.BIERZ_INNE_KOLORY_S3", "1", "True", "False"),
        # --- mapowanie CC-04 ---
        ("GEN_MAP_CC04", "m.MAPOWANIE_TYPU_CC04['LG']['K11']",
         '[["LG","Pt100-99","1586A-02","101","CC-04-LG"]]', "Pt100-99", "Pt100-09"),
        # --- PZ ---
        ("CC_PZ_FOLDER", "m.PZ_FOLDER_ARK", "PZ_inne",
         os.path.join(BAZA, "PZ_inne"), os.path.join(BAZA, "PZ")),
    ],
}

# Ustawienia, ktorych nie da sie sprawdzic przez stala modulu — dzialaja dopiero
# w trakcie przebiegu. Maja wlasne testy nizej / w test_obieg.py.
POKRYTE_W_CZASIE_PRZEBIEGU = {"CC_FOLDER", "OBS_TXT_FILES"}


class TestPokrycie(unittest.TestCase):
    """Zadne ustawienie panelu nie moze zostac bez podlaczenia do skryptu."""

    def test_kazde_ustawienie_jest_sprawdzane(self):
        sprawdzane = {e for wpisy in KONTRAKT.values() for e, *_ in wpisy}
        brakujace = set(C.WG_ENV) - sprawdzane - POKRYTE_W_CZASIE_PRZEBIEGU
        self.assertFalse(
            brakujace,
            f"Ustawienia w panelu bez testu kontraktu: {sorted(brakujace)}")

    def test_nie_testujemy_nieistniejacych_ustawien(self):
        sprawdzane = {e for wpisy in KONTRAKT.values() for e, *_ in wpisy}
        nadmiarowe = sprawdzane - set(C.WG_ENV)
        self.assertFalse(nadmiarowe,
                         f"Test odwoluje sie do nieznanych ustawien: {sorted(nadmiarowe)}")

    def test_nazwy_wystepuja_w_zrodlach(self):
        """Kazda nazwa zmiennej musi pojawic sie w kodzie ktoregos ze skryptow."""
        zrodla = ""
        for plik in ("analizuj_excele.py", "generuj_obserwacje.py",
                     "generuj_arkusze.py", "app_gui.py"):
            with open(os.path.join(KORZEN, plik), encoding="utf-8") as f:
                zrodla += f.read()
        for env in C.WG_ENV:
            with self.subTest(env=env):
                self.assertIn(env, zrodla,
                              f"'{env}' nie wystepuje w zadnym skrypcie")


class TestWartosciDomyslne(unittest.TestCase):
    """
    Bez panelu skrypty musza dzialac jak przed zmianami — inaczej reczne
    uruchomienie 'python generuj_obserwacje.py' dawaloby inne wyniki.
    """

    @classmethod
    def setUpClass(cls):
        cls.odczyt = {}
        for modul, wpisy in KONTRAKT.items():
            wyrazenia = sorted({w for _e, w, *_r in wpisy})
            cls.odczyt[modul] = wartosci_z_modulu(
                modul, wyrazenia, env={"CC_FOLDER": BAZA})

    def test_domyslne_zgodne_z_kodem(self):
        for modul, wpisy in KONTRAKT.items():
            for env, wyrazenie, _wart, _oczek, domyslna in wpisy:
                if domyslna is None:
                    continue          # brak sensownej domyslnej (np. nazwa protokolu)
                with self.subTest(modul=modul, env=env):
                    self.assertEqual(self.odczyt[modul][wyrazenie], domyslna)

    def test_domyslne_rejestru_zgodne_ze_skryptami(self):
        """Rejestr panelu i stale w skryptach nie moga sie rozjechac."""
        proste = {
            "GEN_NR_SW": "1047", "GEN_NR_POM": "9", "GEN_MODEL_CZUJ": "MX1101-02",
            "GEN_K18_CC": "S8000-02", "GEN_K18_CC04": "S8000",
            "OBS_PROG": "10.0", "OBS_TOL": "3.0", "OBS_PROG_T": "0.4",
            "GEN_TAB_RATIO": "0.85", "GEN_KOLOR_AKT": "#CCFFCC",
        }
        for env, oczekiwana in proste.items():
            with self.subTest(env=env):
                self.assertEqual(str(C.WG_ENV[env].domyslna), oczekiwana)


class TestSterowanieZPanelu(unittest.TestCase):
    """
    Ustawienie zmienione w panelu MUSI zmienic stala w skrypcie. To jest test,
    ktory wykrylby stary problem "panel pokazuje opcje, skrypt jej nie czyta".
    """

    @classmethod
    def setUpClass(cls):
        cls.odczyt = {}
        for modul, wpisy in KONTRAKT.items():
            env = {"CC_FOLDER": BAZA}
            for nazwa, _wyr, wartosc, *_r in wpisy:
                env[nazwa] = wartosc
            wyrazenia = sorted({w for _e, w, *_r in wpisy})
            cls.odczyt[modul] = wartosci_z_modulu(modul, wyrazenia, env=env)

    def test_kazde_ustawienie_dociera_do_skryptu(self):
        for modul, wpisy in KONTRAKT.items():
            for env, wyrazenie, wartosc, oczekiwana, _dom in wpisy:
                with self.subTest(modul=modul, env=env):
                    self.assertEqual(
                        self.odczyt[modul][wyrazenie], oczekiwana,
                        f"{env}='{wartosc}' nie zmienilo {wyrazenie}")

    def test_wartosci_testowe_roznia_sie_od_domyslnych(self):
        """Inaczej test przechodzilby, nawet gdyby ustawienie bylo ignorowane."""
        for modul, wpisy in KONTRAKT.items():
            for env, _wyr, _wart, oczekiwana, domyslna in wpisy:
                if domyslna is None:
                    continue
                with self.subTest(modul=modul, env=env):
                    self.assertNotEqual(oczekiwana, domyslna)


if __name__ == "__main__":
    unittest.main(verbosity=2)
