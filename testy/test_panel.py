# -*- coding: utf-8 -*-
"""
Testy okna panelu (app_gui.py) — budowa krokow, zbieranie ustawien, walidacja
danych wejsciowych przed startem i zachowanie kolejki przy bledzie.

Zaden test nie dotyka plikow projektu: folder roboczy i plik ustawien wskazuja
na piaskownice.
"""

import os
import shutil
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from PySide6.QtWidgets import QApplication, QMessageBox

import cc_config as C
import app_gui
from wspolne import (KORZEN, nowa_piaskownica, PLIKI_OBSERWACJA, FOLDERY_OBSERWACJA,
                     PLIKI_ARKUSZE, SZABLON_ARKUSZA)

_app = QApplication.instance() or QApplication([])


class BazaOkna(unittest.TestCase):
    """Wspolne rusztowanie: okno wskazujace wylacznie na piaskownice."""

    PLIKI = ()
    FOLDERY = ()

    def setUp(self):
        self.folder = nowa_piaskownica(self.nazwa_piaskownicy(),
                                       self.PLIKI, self.FOLDERY)
        self._stary_plik = C.PLIK_USTAWIEN
        C.PLIK_USTAWIEN = os.path.join(self.folder, "cc_ustawienia.json")
        self.okno = app_gui.Okno()
        self.okno.wartosci["CC_FOLDER"] = self.folder
        self.okno._odswiez_folder()

    def tearDown(self):
        self.okno.proc = None
        # Zapis ustawien jest odroczony (QTimer 600 ms). Gdyby wystrzelil juz po
        # przywroceniu prawdziwej sciezki, test zapisalby swoje ustawienia do
        # folderu projektu. Zatrzymujemy go, zanim oddamy sciezke.
        self.okno._zapis.stop()
        self.okno.deleteLater()
        C.PLIK_USTAWIEN = self._stary_plik

    def nazwa_piaskownicy(self):
        return "panel_" + self._testMethodName


class TestBudowaOkna(BazaOkna):

    def test_ma_wszystkie_strony(self):
        self.assertEqual(self.okno.stos.count(), 5)

    def test_kroki_uruchamialne_maja_status(self):
        self.assertEqual(set(self.okno.statusy), {"analiza", "obs", "ark"})

    def test_pola_pokrywaja_caly_rejestr(self):
        # CC_FOLDER ma wlasne miejsce w naglowku; OBS_TXT_FILES i ANL_PLIKI to
        # listy plikow z zaznaczaniem (ListaPlikow), nie zwykle pola formularza.
        oczekiwane = set(C.WG_ENV) - {"CC_FOLDER", "OBS_TXT_FILES", "ANL_PLIKI"}
        self.assertEqual(set(self.okno.pola), oczekiwane)

    def test_listy_plikow_sa_podpiete_do_ustawien(self):
        """Kazda lista z zaznaczaniem musi odpowiadac wpisowi w rejestrze."""
        for env in ("OBS_TXT_FILES", "ANL_PLIKI"):
            with self.subTest(env=env):
                self.assertIn(env, C.WG_ENV)
                self.assertEqual(C.WG_ENV[env].typ, "pliki")

    def test_lista_kontrolna_niepusta(self):
        self.assertGreaterEqual(len(self.okno.kontrolne), 5)

    def test_przelaczanie_stron(self):
        for nr in range(5):
            self.okno._przelacz(nr)
            self.assertEqual(self.okno.stos.currentIndex(), nr)

    def test_kazdy_krok_ma_przycisk_uruchomienia(self):
        for klucz in ("analiza", "obs", "ark"):
            with self.subTest(krok=klucz):
                self.assertTrue(hasattr(self.okno, f"btn_{klucz}"))

    def test_przycisk_uruchomienia_nie_przewija_sie_z_formularzem(self):
        """
        Przycisk musi siedziec w szapce strony, poza obszarem przewijania —
        inaczej przy dlugim formularzu trzeba go szukac scrollem.
        """
        from PySide6.QtWidgets import QScrollArea
        for nr, klucz in ((1, "analiza"), (2, "obs"), (3, "ark")):
            with self.subTest(krok=klucz):
                przycisk = getattr(self.okno, f"btn_{klucz}")
                przewijane = self.okno.stos.widget(nr).findChild(QScrollArea)
                self.assertIsNotNone(przewijane)
                self.assertFalse(przewijane.isAncestorOf(przycisk),
                                 "przycisk wpadl do obszaru przewijania")

    def test_kazde_pole_zaawansowane_jest_na_stronie_zaawansowanej(self):
        zaaw = [u.env for u in C.USTAWIENIA if u.poziom == C.ZAAWANSOWANY]
        strona = self.okno.stos.widget(4)
        obecne = {p.ust.env for p in strona.findChildren(app_gui.Pole)}
        self.assertEqual(set(zaaw), obecne)

    def test_kazde_pole_podstawowe_jest_na_stronie_swojego_kroku(self):
        """
        Ustawienie oznaczone jako podstawowe nie moze zginac w Zaawansowanych —
        na tym potknely sie ustawienia zdjec, ktore trafily nie tam, gdzie
        podejmuje sie decyzje o ich zbieraniu.
        """
        strony = {"przygotowanie": 0, "analiza": 1, "obs": 2, "ark": 3}
        # CC_FOLDER ma wlasne miejsce w naglowku, listy plikow to osobne widgety.
        bez_pola = {"CC_FOLDER", "OBS_TXT_FILES", "ANL_PLIKI"}
        for u in C.USTAWIENIA:
            if u.poziom != C.PODSTAWOWY or u.env in bez_pola:
                continue
            with self.subTest(env=u.env, krok=u.krok):
                strona = self.okno.stos.widget(strony[u.krok])
                obecne = {p.ust.env for p in strona.findChildren(app_gui.Pole)}
                self.assertIn(u.env, obecne)

    def test_ustawienia_zdjec_sa_na_stronie_obserwacji(self):
        """Wlacznik zdjec i sciezka zrodlowa maja byc pod reka przy kroku 2."""
        strona = self.okno.stos.widget(2)
        obecne = {p.ust.env for p in strona.findChildren(app_gui.Pole)}
        for env in ("OBS_FOTO", "OBS_FOTO_ZRODLO"):
            with self.subTest(env=env):
                self.assertIn(env, obecne)

    def test_zdjecia_sa_w_tej_samej_karcie_co_lista_plikow(self):
        """
        Regresja z realnego uzycia: jako osobna karta ponizej wypadaly poza
        dolna krawedz okna na laptopie i uzytkownik ich nie znajdowal.
        Musza siedziec w karcie z lista plikow TXT.
        """
        karta = self.okno.lista_txt.parentWidget()
        w_karcie = {p.ust.env for p in karta.findChildren(app_gui.Pole)}
        self.assertIn("OBS_FOTO", w_karcie)
        self.assertIn("OBS_FOTO_ZRODLO", w_karcie)


class TestZbieranieUstawien(BazaOkna):

    def test_zmiana_pola_trafia_do_wartosci(self):
        self.okno.pola["GEN_NR_SW"].ustaw(3333)
        self.okno._zbierz()
        self.assertEqual(self.okno.wartosci["GEN_NR_SW"], 3333)

    def test_zaznaczenie_logow_trafia_do_wartosci(self):
        """Lista z kroku 1 musi byc ustawieniem, a nie ozdoba."""
        wej = os.path.join(self.folder, "excel_do_analizy")
        os.makedirs(wej, exist_ok=True)
        open(os.path.join(wej, "logger_a.csv"), "w").close()
        open(os.path.join(wej, "logger_b.csv"), "w").close()
        self.okno.lista_logow.odswiez()
        self.okno.lista_logow.ustaw_zaznaczone(["logger_b.csv"])
        self.okno._zbierz()
        self.assertEqual(self.okno.wartosci["ANL_PLIKI"], ["logger_b.csv"])
        self.assertEqual(C.do_env(self.okno.wartosci)["ANL_PLIKI"], "logger_b.csv")

    def test_zaznaczenie_txt_trafia_do_wartosci(self):
        open(os.path.join(self.folder, "pomiar.txt"), "w").close()
        self.okno.lista_txt.odswiez()
        self.okno.lista_txt.ustaw_zaznaczone(["pomiar.txt"])
        self.okno._zbierz()
        self.assertEqual(self.okno.wartosci["OBS_TXT_FILES"], ["pomiar.txt"])

    def test_eksport_env_jest_kompletny(self):
        self.okno._zbierz()
        env = C.do_env(self.okno.wartosci)
        self.assertEqual(set(env), set(C.WG_ENV))

    def test_zapis_tworzy_plik_ustawien(self):
        self.okno.pola["GEN_K18_CC"].ustaw("OPTIDEW")
        self.okno._zapisz()
        self.assertTrue(os.path.exists(C.PLIK_USTAWIEN))
        self.assertEqual(C.wczytaj()["GEN_K18_CC"], "OPTIDEW")

    def test_folder_roboczy_nie_jest_nadpisywany_przez_zbierz(self):
        """Folder ustawia sie w naglowku — _zbierz() nie moze go cofnac."""
        self.okno._zbierz()
        self.assertEqual(self.okno.wartosci["CC_FOLDER"], self.folder)

    def test_przywroc_domyslne_zachowuje_folder(self):
        self.okno.pola["GEN_NR_SW"].ustaw(9999)
        pierwotny = QMessageBox.question
        QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.Yes)
        try:
            self.okno._przywroc_domyslne()
        finally:
            QMessageBox.question = pierwotny
        self.assertEqual(self.okno.wartosci["GEN_NR_SW"], 1047)
        self.assertEqual(self.okno.wartosci["CC_FOLDER"], self.folder)


class TestWalidacjaBrakow(BazaOkna):
    """Pusta piaskownica — panel ma wypisac konkretne braki, nie ruszac na slepo."""

    def test_analiza_zglasza_brak_folderu(self):
        problemy = self.okno._problemy(["analiza"])
        self.assertTrue(any("Brak folderu z logami" in p for p in problemy))

    def test_analiza_zglasza_pusty_folder(self):
        os.makedirs(os.path.join(self.folder, "excel_do_analizy"), exist_ok=True)
        problemy = self.okno._problemy(["analiza"])
        self.assertTrue(any("jest pusty" in p for p in problemy))

    def test_obserwacja_zglasza_brak_txt(self):
        problemy = self.okno._problemy(["obs"])
        self.assertTrue(any("TXT" in p for p in problemy))

    def test_obserwacja_zglasza_brak_szablonow(self):
        problemy = self.okno._problemy(["obs"])
        self.assertTrue(any("Nie znaleziono pliku" in p for p in problemy))

    def test_arkusze_zglaszaja_brak_protokolu(self):
        problemy = self.okno._problemy(["ark"])
        self.assertTrue(any("Plik protokolu" in p for p in problemy))

    def test_arkusze_zglaszaja_brak_plikow_linkowanych(self):
        problemy = self.okno._problemy(["ark"])
        self.assertTrue(any("Wzory.xls" in p for p in problemy))

    def test_word_wylaczony_nie_wymaga_szablonow_word(self):
        self.okno.pola["GEN_WORD"].ustaw(False)
        self.okno._zbierz()
        problemy = self.okno._problemy(["ark"])
        self.assertFalse(any("szablonu Word" in p for p in problemy))

    def test_word_wlaczony_wymaga_szablonow_word(self):
        self.okno.pola["GEN_WORD"].ustaw(True)
        self.okno._zbierz()
        problemy = self.okno._problemy(["ark"])
        self.assertTrue(any("szablonu Word" in p for p in problemy))

    def test_nieistniejacy_folder_roboczy_przerywa_od_razu(self):
        self.okno.wartosci["CC_FOLDER"] = os.path.join(self.folder, "nie_ma")
        problemy = self.okno._problemy(["analiza", "obs", "ark"])
        self.assertEqual(len(problemy), 1)
        self.assertIn("Folder roboczy nie istnieje", problemy[0])


def _xlsx_z_zakladkami(sciezka, nazwy):
    """
    Tworzy minimalny plik xlsx o zadanych nazwach zakladek.

    Testy walidacji sprawdzaja tylko obecnosc plikow i nazwy zakladek, wiec nie
    potrzebuja prawdziwych protokolow. Wczesniej kopiowaly konkretne pliki
    pomiarowe z folderu projektu — i przestawaly dzialac, gdy uzytkownik
    sprzatnal dane zakonczonego zlecenia.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nazwa in nazwy:
        wb.create_sheet(nazwa)
    wb.save(sciezka)
    return sciezka


class TestWalidacjaKompletnegoZestawu(BazaOkna):
    """Komplet plikow na miejscu — walidacja ma milczec."""

    # Szablony obserwacji/protokolow to stale pliki projektu; reszte budujemy
    # syntetycznie, zeby test nie zalezal od danych biezacego zlecenia.
    PLIKI = tuple(PLIKI_OBSERWACJA[2:])

    def setUp(self):
        super().setUp()
        # Krok 1: wystarczy jakikolwiek plik o obslugiwanym rozszerzeniu.
        wejscie = os.path.join(self.folder, "excel_do_analizy")
        os.makedirs(wejscie, exist_ok=True)
        open(os.path.join(wejscie, "logger.csv"), "w").close()

        # Krok 2: plik TXT pomiaru + folder PZ.
        open(os.path.join(self.folder, "2026-01-01 08.00_999.txt"), "w").close()
        os.makedirs(os.path.join(self.folder, "PZ"), exist_ok=True)
        open(os.path.join(self.folder, "PZ", "PZ 999.pdf"), "w").close()

        # Krok 3: protokol (Strona 2/3), szablon arkusza (Wyniki), szablony Word
        # i pliki linkowane — same nazwy i zakladki, bez zawartosci.
        _xlsx_z_zakladkami(os.path.join(self.folder, "999 - protokół CC.xlsx"),
                           ["Strona 1", "Strona 2", "Strona 3"])
        _xlsx_z_zakladkami(os.path.join(self.folder, "999 - Wzór ark. obl.xlsx"),
                           ["23,30", "Wyniki"])
        for nazwa in ("xxx_yyy_LA_TH_2026 - tylko temp.docx",
                      "xxx_yyy_LA_TH_2026 - zakres.docx",
                      "xxx_yyy_LA_TH_2026 - zakres + temp.docx",
                      "Wzory.xls", "Obliczenia tdp, RH, C.xls"):
            open(os.path.join(self.folder, nazwa), "w").close()

        self.okno._odswiez_folder()
        self.okno.lista_txt.ustaw_zaznaczone(["2026-01-01 08.00_999.txt"])
        self.okno.pola["CC_PROTOKOL"].ustaw("999 - protokół CC.xlsx")
        self.okno.pola["CC_SZABLON"].ustaw("999 - Wzór ark. obl.xlsx")
        self.okno._zbierz()

    def test_analiza_bez_zastrzezen(self):
        self.assertEqual(self.okno._problemy(["analiza"]), [])

    def test_obserwacja_bez_zastrzezen(self):
        self.assertEqual(self.okno._problemy(["obs"]), [])

    def test_arkusze_bez_zastrzezen(self):
        self.assertEqual(self.okno._problemy(["ark"]), [])

    def test_caly_obieg_bez_zastrzezen(self):
        self.assertEqual(self.okno._problemy(["analiza", "obs", "ark"]), [])

    def test_arkusz_obserwacji_wskazany_jako_protokol_jest_wykryty(self):
        """
        Regresja: kolko myszy nad lista podmienialo protokol na arkusz
        obserwacji. Plik istnieje, wiec sama kontrola obecnosci go przepuszcza —
        musi zadzialac sprawdzenie zawartosci.
        """
        _xlsx_z_zakladkami(os.path.join(self.folder, "podmiana.xlsx"),
                           ["obserwacje"])
        self.okno.pola["CC_PROTOKOL"].ustaw("podmiana.xlsx")
        self.okno._zbierz()
        problemy = self.okno._problemy(["ark"])
        self.assertTrue(any("nie wyglada na protokol" in x for x in problemy),
                        problemy)

    def test_zly_plik_jako_szablon_arkusza_jest_wykryty(self):
        _xlsx_z_zakladkami(os.path.join(self.folder, "nie_szablon.xlsx"),
                           ["Strona 1", "Strona 2", "Strona 3"])
        self.okno.pola["CC_SZABLON"].ustaw("nie_szablon.xlsx")
        self.okno._zbierz()
        problemy = self.okno._problemy(["ark"])
        self.assertTrue(any("nie wyglada na arkusz obliczeniowy" in x
                            for x in problemy), problemy)

    def test_komunikat_wymienia_znalezione_zakladki(self):
        _xlsx_z_zakladkami(os.path.join(self.folder, "podmiana.xlsx"),
                           ["obserwacje"])
        self.okno.pola["CC_PROTOKOL"].ustaw("podmiana.xlsx")
        self.okno._zbierz()
        tresc = " ".join(self.okno._problemy(["ark"]))
        self.assertIn("obserwacje", tresc)


class TestKolejkaKrokow(BazaOkna):
    """Obieg musi przerwac sie na pierwszym bledzie, a nie brnac dalej."""

    def test_blad_czysci_kolejke(self):
        self.okno.kolejka = ["obs", "ark"]
        self.okno._koniec("analiza", "1 · Analiza logow", 1)
        self.assertEqual(self.okno.kolejka, [])

    def test_blad_ustawia_status_bledu(self):
        self.okno.kolejka = []
        self.okno._koniec("analiza", "1 · Analiza logow", 2)
        self.assertEqual(self.okno.statusy["analiza"].text(), "✕")

    def test_sukces_ustawia_status_ok(self):
        self.okno.kolejka = []
        self.okno._koniec("obs", "2 · Obserwacja", 0)
        self.assertEqual(self.okno.statusy["obs"].text(), "✓")

    def test_blad_odblokowuje_przyciski(self):
        self.okno._zajety(True)
        self.okno.kolejka = ["ark"]
        self.okno._koniec("obs", "2 · Obserwacja", 1)
        self.assertTrue(self.okno.btn_obieg.isEnabled())

    def test_zajety_blokuje_uruchamianie(self):
        self.okno._zajety(True)
        self.assertFalse(self.okno.btn_obieg.isEnabled())
        self.assertFalse(self.okno.btn_analiza.isEnabled())
        self.assertTrue(self.okno.btn_stop.isEnabled())
        self.okno._zajety(False)
        self.assertTrue(self.okno.btn_obieg.isEnabled())
        self.assertFalse(self.okno.btn_stop.isEnabled())

    def test_log_odnotowuje_pominiete_kroki(self):
        self.okno.kolejka = ["obs", "ark"]
        self.okno._koniec("analiza", "1 · Analiza logow", 1)
        self.assertIn("pominieto 2", self.okno.log.toPlainText())


class TestZmianaFolderu(BazaOkna):

    def test_listy_przeladowuja_sie_po_zmianie_folderu(self):
        inny = nowa_piaskownica("panel_inny_folder")
        open(os.path.join(inny, "nowy_pomiar.txt"), "w").close()
        self.okno.wartosci["CC_FOLDER"] = inny
        self.okno._odswiez_folder()
        nazwy = [self.okno.lista_txt.lista.item(i).data(
                    app_gui.Qt.UserRole)
                 for i in range(self.okno.lista_txt.lista.count())]
        self.assertEqual(nazwy, ["nowy_pomiar.txt"])

    def test_naglowek_pokazuje_sciezke(self):
        self.assertEqual(self.okno.lbl_folder.text(), self.folder)


if __name__ == "__main__":
    unittest.main(verbosity=2)
