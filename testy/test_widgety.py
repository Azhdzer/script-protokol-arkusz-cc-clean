# -*- coding: utf-8 -*-
"""
Testy widgetow panelu (cc_widgets.py) — pola ustawien, lista plikow, kolorowanie
logu, wykrywanie utworzonych plikow i lista kontrolna swiezosci.
"""

import datetime
import os
import sys
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QFont
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

import cc_config as C
import cc_widgets as W
from wspolne import KORZEN, nowa_piaskownica

_app = QApplication.instance() or QApplication([])


def _dotknij(sciezka, wiek_dni=0, tresc=b"x"):
    """Tworzy plik o zadanym wieku (do testow listy kontrolnej)."""
    with open(sciezka, "wb") as f:
        f.write(tresc)
    if wiek_dni:
        kiedy = time.time() - wiek_dni * 86400
        os.utime(sciezka, (kiedy, kiedy))
    return sciezka


class TestFormatowanie(unittest.TestCase):

    def test_ludzki_rozmiar(self):
        self.assertEqual(W.ludzki_rozmiar(512), "512 B")
        self.assertEqual(W.ludzki_rozmiar(2048), "2,0 kB")
        self.assertEqual(W.ludzki_rozmiar(5 * 1024 * 1024), "5,0 MB")

    def test_ile_temu_dzis(self):
        self.assertTrue(W.ile_temu(time.time()).startswith("dzis"))

    def test_ile_temu_wczoraj(self):
        wczoraj = time.time() - 86400
        self.assertTrue(W.ile_temu(wczoraj).startswith("wczoraj"))

    def test_ile_temu_dawno(self):
        self.assertIn("dni temu", W.ile_temu(time.time() - 12 * 86400))


class TestNazwyArkuszy(unittest.TestCase):
    """Podglad zawartosci xlsx czytany wprost z ZIP-a — na prawdziwym pliku."""

    def test_prawdziwy_plik_wynikow(self):
        plik = os.path.join(KORZEN, "wyniki", "37025098_wynik.xlsx")
        if not os.path.exists(plik):
            self.skipTest("brak wyniki/37025098_wynik.xlsx")
        self.assertEqual(W.nazwy_arkuszy(plik), ["Sheet1"])

    def test_protokol_ma_strony(self):
        plik = os.path.join(KORZEN, "xxx_LA_TH_2026 - protokół CC.xlsx")
        if not os.path.exists(plik):
            self.skipTest("brak szablonu protokolu")
        arkusze = W.nazwy_arkuszy(plik)
        self.assertIn("Strona 2", arkusze)
        self.assertIn("Strona 3", arkusze)

    def test_plik_nie_bedacy_xlsx(self):
        folder = nowa_piaskownica("arkusze")
        plik = _dotknij(os.path.join(folder, "nie_xlsx.xlsx"), tresc=b"to nie ZIP")
        self.assertEqual(W.nazwy_arkuszy(plik), [])

    def test_brakujacy_plik(self):
        self.assertEqual(W.nazwy_arkuszy(r"C:\nie\ma\takiego.xlsx"), [])


class TestPole(unittest.TestCase):
    """Kazdy typ pola musi poprawnie oddawac i przyjmowac wartosc."""

    def setUp(self):
        self.folder = nowa_piaskownica("pole")

    def zbuduj(self, env, wartosc=None):
        u = C.WG_ENV[env]
        return W.Pole(u, u.domyslna if wartosc is None else wartosc,
                      folder_cb=lambda: self.folder)

    def test_tekst(self):
        p = self.zbuduj("OBS_PODPIS")
        self.assertEqual(p.wartosc(), "Artsiom Azhdzer")
        p.ustaw("Jan Kowalski")
        self.assertEqual(p.wartosc(), "Jan Kowalski")

    def test_flaga(self):
        p = self.zbuduj("GEN_WORD")
        self.assertTrue(p.wartosc())
        p.ustaw(False)
        self.assertFalse(p.wartosc())

    def test_calk(self):
        p = self.zbuduj("GEN_NR_SW")
        self.assertEqual(p.wartosc(), 1047)
        p.ustaw(2500)
        self.assertEqual(p.wartosc(), 2500)

    def test_liczba(self):
        p = self.zbuduj("OBS_PROG_T")
        self.assertAlmostEqual(p.wartosc(), 0.4)
        p.ustaw(1.25)
        self.assertAlmostEqual(p.wartosc(), 1.25)

    def test_minuty(self):
        p = self.zbuduj("OBS_STAB_PO_RH")
        self.assertEqual(p.wartosc(), 120)
        p.ustaw(45)
        self.assertEqual(p.wartosc(), 45)

    def test_kolor(self):
        p = self.zbuduj("GEN_KOLOR_AKT")
        self.assertEqual(p.wartosc(), "#CCFFCC")
        p.ustaw("#00FF00")
        self.assertEqual(p.wartosc(), "#00FF00")

    def test_folder(self):
        p = self.zbuduj("CC_PZ_FOLDER")
        self.assertEqual(p.wartosc(), "PZ")
        p.ustaw(r"D:\inne\PZ")
        self.assertEqual(p.wartosc(), r"D:\inne\PZ")

    def test_tabela_pomija_puste_wiersze(self):
        p = self.zbuduj("GEN_MAP_CC04")
        # ustaw() dokleja pusty wiersz na dopisanie — nie moze trafic do wyniku
        self.assertEqual(p.wartosc(), C.WG_ENV["GEN_MAP_CC04"].domyslna)

    def test_tabela_przyjmuje_zmiany(self):
        p = self.zbuduj("GEN_MAP_CC04")
        nowa = [["LG", "Pt100-99", "1586A-02", "101", "CC-04-LG"]]
        p.ustaw(nowa)
        self.assertEqual(p.wartosc(), nowa)

    def test_plik_wybiera_wg_podpowiedzi(self):
        """Przy pierwszym uruchomieniu ma trafic w protokol, nie w pierwszy xlsx."""
        for nazwa in ("aaa_obserwacje.xlsx", "zzz_protokół CC.xlsx"):
            _dotknij(os.path.join(self.folder, nazwa))
        p = self.zbuduj("CC_PROTOKOL", wartosc="")
        self.assertEqual(p.wartosc(), "zzz_protokół CC.xlsx")

    def test_plik_pomija_pliki_tymczasowe_excela(self):
        _dotknij(os.path.join(self.folder, "~$protokół CC.xlsx"))
        _dotknij(os.path.join(self.folder, "protokół CC.xlsx"))
        p = self.zbuduj("CC_PROTOKOL", wartosc="")
        self.assertEqual(p.wartosc(), "protokół CC.xlsx")

    def test_brak_pliku_wykrywany(self):
        p = self.zbuduj("CC_PROTOKOL", wartosc="nie_ma_mnie.xlsx")
        self.assertTrue(p.brak_pliku())

    def test_istniejacy_plik_nie_jest_zglaszany(self):
        _dotknij(os.path.join(self.folder, "jest.xlsx"))
        p = self.zbuduj("CC_PROTOKOL", wartosc="jest.xlsx")
        self.assertFalse(p.brak_pliku())

    def test_zapisany_plik_spoza_folderu_zostaje_widoczny(self):
        """Nie wolno po cichu podmienic wyboru uzytkownika na inny plik."""
        _dotknij(os.path.join(self.folder, "inny protokół.xlsx"))
        p = self.zbuduj("CC_PROTOKOL", wartosc="usuniety protokół.xlsx")
        self.assertEqual(p.wartosc(), "usuniety protokół.xlsx")

    def test_sygnal_zmiany(self):
        p = self.zbuduj("GEN_NR_SW")
        licznik = []
        p.zmienione.connect(lambda: licznik.append(1))
        p._kontrolka.setValue(1234)
        self.assertTrue(licznik)


class TestOdpornoscNaKolko(unittest.TestCase):
    """
    Kolko myszy nad polem NIE moze zmieniac jego wartosci.

    Regresja z prawdziwego uzycia: przewijajac dlugi formularz uzytkownik
    najechal kolkiem na liste "Szablon arkusza obliczeniowego" i po cichu
    podmienil plik na arkusz obserwacji. Zmiana zapisala sie do ustawien
    i wyszla dopiero przy uruchomieniu.
    """

    def setUp(self):
        self.folder = nowa_piaskownica("kolko")
        for nazwa in ("aaa protokół CC.xlsx", "bbb protokół CC.xlsx",
                      "ccc protokół CC.xlsx"):
            _dotknij(os.path.join(self.folder, nazwa))

    def krec(self, widget, obroty=-5):
        """Symuluje ruch kolka nad widgetem."""
        from PySide6.QtCore import QPointF
        from PySide6.QtGui import QWheelEvent
        srodek = QPointF(widget.rect().center())
        zdarzenie = QWheelEvent(
            srodek, QPointF(widget.mapToGlobal(widget.rect().center())),
            QPoint(0, obroty * 24), QPoint(0, obroty * 120),
            Qt.NoButton, Qt.NoModifier, Qt.NoScrollPhase, False)
        QApplication.sendEvent(widget, zdarzenie)
        return zdarzenie

    def test_lista_rozwijana_nie_reaguje_na_kolko(self):
        u = C.WG_ENV["CC_PROTOKOL"]
        pole = W.Pole(u, "bbb protokół CC.xlsx", folder_cb=lambda: self.folder)
        przed = pole.wartosc()
        self.krec(pole._kontrolka)
        self.assertEqual(pole.wartosc(), przed)

    def test_pole_calkowite_nie_reaguje_na_kolko(self):
        u = C.WG_ENV["GEN_NR_SW"]
        pole = W.Pole(u, 1047, folder_cb=lambda: self.folder)
        self.krec(pole._kontrolka)
        self.assertEqual(pole.wartosc(), 1047)

    def test_pole_dziesietne_nie_reaguje_na_kolko(self):
        u = C.WG_ENV["OBS_PROG_T"]
        pole = W.Pole(u, 0.4, folder_cb=lambda: self.folder)
        self.krec(pole._kontrolka)
        self.assertAlmostEqual(pole.wartosc(), 0.4)

    def test_kolko_jest_oddawane_w_gore(self):
        """Odrzucenie zdarzenia pozwala przewinac strone pod spodem."""
        u = C.WG_ENV["GEN_NR_SW"]
        pole = W.Pole(u, 1047, folder_cb=lambda: self.folder)
        zdarzenie = self.krec(pole._kontrolka)
        self.assertFalse(zdarzenie.isAccepted())

    def test_klawiatura_nadal_zmienia_wartosc(self):
        """Blokujemy tylko kolko — normalna edycja ma dzialac."""
        u = C.WG_ENV["GEN_NR_SW"]
        pole = W.Pole(u, 1047, folder_cb=lambda: self.folder)
        QTest.keyClick(pole._kontrolka, Qt.Key_Up)
        self.assertEqual(pole.wartosc(), 1048)


class TestPrzelacznik(unittest.TestCase):
    """
    Przelacznik z suwakiem. Kluczowe: polozenie galki musi ZAWSZE zgadzac sie
    ze stanem — takze gdy stan ustawiono przy wyciszonych sygnalach.
    """

    def test_domyslnie_wylaczony_ma_galke_z_lewej(self):
        p = W.Przelacznik("Etap")
        p.grab()
        self.assertEqual(p.pozycja, 0.0)

    def test_ustawiony_programowo_ma_galke_z_prawej(self):
        p = W.Przelacznik("Etap")
        p.setChecked(True)
        p.grab()
        self.assertEqual(p.pozycja, 1.0)

    def test_ustawienie_przy_wyciszonych_sygnalach_dociaga_galke(self):
        """
        Regresja: Pole.ustaw() wola blockSignals, wiec 'toggled' nie leci
        i animacja nie startuje. Bez dociagniecia przelacznik pokazywalby
        stan odwrotny do prawdziwego.
        """
        p = W.Przelacznik("Etap")
        p.blockSignals(True)
        p.setChecked(True)
        p.blockSignals(False)
        p.grab()
        self.assertEqual(p.pozycja, 1.0)

    def test_wylaczenie_przy_wyciszonych_sygnalach_dociaga_galke(self):
        p = W.Przelacznik("Etap")
        p.setChecked(True)
        p.grab()
        p.blockSignals(True)
        p.setChecked(False)
        p.blockSignals(False)
        p.grab()
        self.assertEqual(p.pozycja, 0.0)

    def test_pole_typu_flaga_pokazuje_wlaczony_stan(self):
        """Test na styku: wartosc z ustawien musi byc widoczna na przelaczniku."""
        u = C.WG_ENV["GEN_WORD"]
        pole = W.Pole(u, True, folder_cb=os.getcwd)
        pole._kontrolka.grab()
        self.assertTrue(pole.wartosc())
        self.assertEqual(pole._kontrolka.pozycja, 1.0)

    def test_klikniecie_w_podpis_przelacza(self):
        """Celem jest caly widget, nie sam tor przelacznika."""
        p = W.Przelacznik("Generuj swiadectwa Word")
        p.resize(p.sizeHint())
        punkt = p.rect().center()
        punkt.setX(p.width() - 12)          # na tekscie, daleko od toru
        QTest.mouseClick(p, Qt.LeftButton, Qt.NoModifier, punkt)
        self.assertTrue(p.isChecked())

    def test_klikniecie_zglasza_toggled(self):
        p = W.Przelacznik("Etap")
        p.resize(p.sizeHint())
        stany = []
        p.toggled.connect(stany.append)
        QTest.mouseClick(p, Qt.LeftButton, Qt.NoModifier, p.rect().center())
        self.assertEqual(stany, [True])

    def test_szerokosc_uwzglednia_podpis(self):
        krotki = W.Przelacznik("A")
        dlugi = W.Przelacznik("Usuwaj puste bloki Strony 3")
        self.assertGreater(dlugi.sizeHint().width(), krotki.sizeHint().width())

    def test_wylaczony_rysuje_sie_bez_bledu(self):
        p = W.Przelacznik("Etap")
        p.setEnabled(False)
        self.assertFalse(p.grab().isNull())


class TestListaPlikow(unittest.TestCase):

    def setUp(self):
        self.folder = nowa_piaskownica("lista")
        for i, nazwa in enumerate(("stary.txt", "sredni.txt", "nowy.txt")):
            _dotknij(os.path.join(self.folder, nazwa), wiek_dni=(2 - i) * 3)
        _dotknij(os.path.join(self.folder, "nieistotny.xlsx"))
        _dotknij(os.path.join(self.folder, "~$ukryty.txt"))
        self.lista = W.ListaPlikow((".txt",), lambda: self.folder)
        self.lista.odswiez()

    def nazwy(self):
        return [self.lista.lista.item(i).data(Qt.UserRole)
                for i in range(self.lista.lista.count())]

    def test_filtruje_po_rozszerzeniu(self):
        self.assertNotIn("nieistotny.xlsx", self.nazwy())

    def test_pomija_pliki_tymczasowe_excela(self):
        self.assertNotIn("~$ukryty.txt", self.nazwy())

    def test_sortuje_od_najnowszego(self):
        self.assertEqual(self.nazwy(), ["nowy.txt", "sredni.txt", "stary.txt"])

    def test_domyslnie_nic_nie_zaznaczone(self):
        self.assertEqual(self.lista.zaznaczone(), [])

    def test_zaznaczanie_i_odczyt(self):
        self.lista.ustaw_zaznaczone(["nowy.txt", "stary.txt"])
        self.assertEqual(sorted(self.lista.zaznaczone()), ["nowy.txt", "stary.txt"])

    def test_odswiezenie_zachowuje_zaznaczenie(self):
        self.lista.ustaw_zaznaczone(["sredni.txt"])
        self.lista.odswiez()
        self.assertEqual(self.lista.zaznaczone(), ["sredni.txt"])

    def test_nowy_plik_pojawia_sie_po_odswiezeniu(self):
        _dotknij(os.path.join(self.folder, "dopisany.txt"))
        self.lista.odswiez()
        self.assertIn("dopisany.txt", self.nazwy())

    def test_podsumowanie_informuje_o_sklejaniu(self):
        self.lista.ustaw_zaznaczone(["nowy.txt", "stary.txt"])
        self.assertIn("sklejone", self.lista.podsumowanie.text())

    # ── przyciski (klikamy naprawde, nie wolamy metod) ───────────────────
    def przycisk(self, napis):
        from PySide6.QtWidgets import QPushButton
        return next(b for b in self.lista.findChildren(QPushButton)
                    if b.text() == napis)

    def test_przycisk_odswiez_pokazuje_nowe_pliki(self):
        """
        Regresja: `clicked` niesie argument bool, ktory trafial w parametr
        `zachowaj` metody odswiez() i wywracal ja (TypeError). Qt polykalo
        wyjatek — przycisk milczaco nie robil nic i trzeba bylo przelaczac
        zakladki, zeby lista sie odswiezyla.
        """
        _dotknij(os.path.join(self.folder, "dopisany.txt"))
        self.przycisk("Odswiez").click()
        self.assertIn("dopisany.txt", self.nazwy())

    def test_przycisk_odswiez_zachowuje_zaznaczenie(self):
        self.lista.ustaw_zaznaczone(["nowy.txt"])
        self.przycisk("Odswiez").click()
        self.assertEqual(self.lista.zaznaczone(), ["nowy.txt"])

    def test_przycisk_zaznacz_wszystko_dziala(self):
        self.przycisk("Zaznacz wszystko").click()
        self.assertEqual(len(self.lista.zaznaczone()), self.lista.lista.count())

    def test_przycisk_odznacz_wszystko_dziala(self):
        self.lista.zaznacz_wszystko()
        self.przycisk("Odznacz wszystko").click()
        self.assertEqual(self.lista.zaznaczone(), [])

    def test_odswiez_odporne_na_argument_sygnalu(self):
        """Nawet podpiete wprost pod sygnal nie moze sie wywrocic."""
        _dotknij(os.path.join(self.folder, "kolejny.txt"))
        self.lista.odswiez(False)          # tak wygladal blad
        self.assertIn("kolejny.txt", self.nazwy())

    def test_pusty_folder(self):
        pusty = nowa_piaskownica("lista_pusta")
        lista = W.ListaPlikow((".txt",), lambda: pusty)
        lista.odswiez()
        self.assertEqual(lista.lista.count(), 0)
        self.assertIn("Brak", lista.podsumowanie.text())

    # ── zaznaczanie hurtem ────────────────────────────────────────────────
    def test_zaznacz_wszystko(self):
        self.lista.zaznacz_wszystko()
        self.assertEqual(sorted(self.lista.zaznaczone()), sorted(self.nazwy()))

    def test_odznacz_wszystko(self):
        self.lista.zaznacz_wszystko()
        self.lista.odznacz_wszystko()
        self.assertEqual(self.lista.zaznaczone(), [])

    def test_zaznacz_wszystko_zglasza_zmiane_raz(self):
        """Sygnal ma poleciec raz na operacje, nie raz na plik."""
        licznik = []
        self.lista.zmienione.connect(lambda: licznik.append(1))
        self.lista.zaznacz_wszystko()
        self.assertEqual(len(licznik), 1)

    def test_zaznacz_wszystko_aktualizuje_podsumowanie(self):
        self.lista.zaznacz_wszystko()
        self.assertIn("3 z 3", self.lista.podsumowanie.text())

    # ── klikniecie w dowolne miejsce wiersza ──────────────────────────────
    def klik_w_wiersz(self, nr, przesuniecie_x):
        """Klika w wiersz `nr` w podanej odleglosci od jego lewej krawedzi."""
        element = self.lista.lista.item(nr)
        prostokat = self.lista.lista.visualItemRect(element)
        punkt = prostokat.center()
        punkt.setX(prostokat.left() + przesuniecie_x)
        QTest.mousePress(self.lista.lista.viewport(), Qt.LeftButton,
                         Qt.NoModifier, punkt)

    def test_klikniecie_w_nazwe_pliku_zaznacza(self):
        """Sedno poprawki: nie trzeba celowac w maly kwadracik."""
        self.klik_w_wiersz(0, 160)          # daleko od checkboxa, na tekscie
        self.assertEqual(self.lista.zaznaczone(), ["nowy.txt"])

    def test_ponowne_klikniecie_odznacza(self):
        self.klik_w_wiersz(0, 160)
        self.klik_w_wiersz(0, 160)
        self.assertEqual(self.lista.zaznaczone(), [])

    def test_klikniecie_w_kwadracik_przelacza_tylko_raz(self):
        """Trafienie w sam checkbox nie moze przelaczyc stanu dwukrotnie."""
        self.klik_w_wiersz(0, 12)
        self.assertEqual(self.lista.zaznaczone(), ["nowy.txt"])

    def test_klikniecie_zglasza_zmiane(self):
        licznik = []
        self.lista.zmienione.connect(lambda: licznik.append(1))
        self.klik_w_wiersz(1, 160)
        self.assertTrue(licznik)

    def test_klikniecie_w_pustke_nic_nie_zaznacza(self):
        prostokat = self.lista.lista.visualItemRect(self.lista.lista.item(2))
        punkt = prostokat.center()
        punkt.setY(prostokat.bottom() + 60)   # ponizej ostatniego wiersza
        QTest.mousePress(self.lista.lista.viewport(), Qt.LeftButton,
                         Qt.NoModifier, punkt)
        self.assertEqual(self.lista.zaznaczone(), [])

    # ── wyroznienie zaznaczonego wiersza ──────────────────────────────────
    def test_zaznaczony_wiersz_jest_pogrubiony(self):
        self.lista.ustaw_zaznaczone(["nowy.txt"])
        self.assertTrue(self.lista.lista.item(0).font().bold())

    def test_niezaznaczony_wiersz_nie_jest_pogrubiony(self):
        self.lista.ustaw_zaznaczone(["nowy.txt"])
        self.assertFalse(self.lista.lista.item(1).font().bold())

    def test_odznaczenie_zdejmuje_pogrubienie(self):
        self.lista.ustaw_zaznaczone(["nowy.txt"])
        self.lista.odznacz_wszystko()
        self.assertFalse(self.lista.lista.item(0).font().bold())


class TestWidokLogu(unittest.TestCase):
    """Blad musi byc widoczny na pierwszy rzut oka, nie zagubiony w scianie tekstu."""

    def barwa(self, linia):
        widok = W.WidokLogu()
        return widok._format(linia).foreground().color().name().upper()

    def test_blad_na_czerwono(self):
        self.assertEqual(self.barwa("  !!! BLAD: brak kolumny"), "#FF8A80")

    def test_traceback_na_czerwono(self):
        self.assertEqual(self.barwa("Traceback (most recent call last):"), "#FF8A80")

    def test_uwaga_na_pomaranczowo(self):
        self.assertEqual(self.barwa("  [UWAGA] Brak pliku czujnika"), "#FFD180")

    def test_ok_na_zielono(self):
        self.assertEqual(self.barwa("  [OK] Zapisano protokol"), "#B9F6CA")

    def test_zwykly_tekst_neutralny(self):
        self.assertEqual(self.barwa("Numer pomiaru : 188"), "#D6E2F0")

    def test_dopisywanie_zachowuje_tresc(self):
        widok = W.WidokLogu()
        widok.dopisz("pierwsza\n")
        widok.dopisz("druga\n")
        self.assertIn("pierwsza", widok.toPlainText())
        self.assertIn("druga", widok.toPlainText())


class TestPanelWynikow(unittest.TestCase):
    """Roznica migawek folderu — to ona pokazuje uzytkownikowi, co powstalo."""

    def setUp(self):
        self.folder = nowa_piaskownica("wyniki_panel")
        os.makedirs(os.path.join(self.folder, "wyniki"), exist_ok=True)
        self.panel = W.PanelWynikow(lambda: self.folder)

    def wiersze(self):
        d = self.panel.drzewo
        return {d.topLevelItem(i).text(0): d.topLevelItem(i).text(1)
                for i in range(d.topLevelItemCount())}

    def test_nowy_plik_oznaczony_jako_nowy(self):
        self.panel.zapamietaj_stan()
        _dotknij(os.path.join(self.folder, "wyniki", "swiezy.xlsx"))
        self.panel.pokaz_zmiany("test")
        self.assertEqual(self.wiersze().get(os.path.join("wyniki", "swiezy.xlsx")),
                         "nowy")

    def test_zmieniony_plik_oznaczony_jako_zmieniony(self):
        plik = _dotknij(os.path.join(self.folder, "protokol.xlsx"), tresc=b"stara")
        self.panel.zapamietaj_stan()
        time.sleep(0.05)
        _dotknij(plik, tresc=b"nowa tresc, inny rozmiar")
        self.panel.pokaz_zmiany("test")
        self.assertEqual(self.wiersze().get("protokol.xlsx"), "zmieniony")

    def test_brak_zmian_jest_komunikowany(self):
        _dotknij(os.path.join(self.folder, "nietkniety.xlsx"))
        self.panel.zapamietaj_stan()
        self.panel.pokaz_zmiany("test")
        self.assertEqual(self.panel.drzewo.topLevelItemCount(), 0)
        self.assertIn("zaden plik", self.panel.naglowek.text())

    def test_naglowek_liczy_pliki(self):
        self.panel.zapamietaj_stan()
        _dotknij(os.path.join(self.folder, "a.xlsx"))
        _dotknij(os.path.join(self.folder, "b.docx"))
        self.panel.pokaz_zmiany("3 · Arkusze")
        self.assertIn("utworzono 2", self.panel.naglowek.text())

    def test_pokazuje_zakladki_prawdziwego_xlsx(self):
        import shutil
        zrodlo = os.path.join(KORZEN, "xxx_LA_TH_2026 - protokół CC.xlsx")
        if not os.path.exists(zrodlo):
            self.skipTest("brak szablonu protokolu")
        self.panel.zapamietaj_stan()
        shutil.copy2(zrodlo, os.path.join(self.folder, "protokol.xlsx"))
        self.panel.pokaz_zmiany("test")
        szczegol = self.panel.drzewo.topLevelItem(0).text(3)
        self.assertIn("Strona 2", szczegol)

    def test_pliki_tymczasowe_excela_pomijane(self):
        self.panel.zapamietaj_stan()
        _dotknij(os.path.join(self.folder, "~$protokol.xlsx"))
        self.panel.pokaz_zmiany("test")
        self.assertEqual(self.panel.drzewo.topLevelItemCount(), 0)


class TestWierszKontrolny(unittest.TestCase):
    """Przypomnienie o nieaktualnych plikach — sedno kroku 'Przygotowanie'."""

    def setUp(self):
        self.folder = nowa_piaskownica("kontrolny")

    def wiersz(self, nazwa, **kw):
        w = W.WierszKontrolny("Tytul", "Opis",
                              lambda: os.path.join(self.folder, nazwa), **kw)
        w.odswiez()
        return w

    def test_brak_pliku_to_blad(self):
        w = self.wiersz("nie_ma.xls", prog_dni=30)
        self.assertEqual(w.ikona.text(), "✕")
        self.assertIn("BRAK", w.stan.text())

    def test_swiezy_plik_jest_ok(self):
        _dotknij(os.path.join(self.folder, "Wzory.xls"), wiek_dni=1)
        w = self.wiersz("Wzory.xls", prog_dni=30)
        self.assertEqual(w.ikona.text(), "✓")
        self.assertNotIn("ZAKTUALIZUJ", w.stan.text())

    def test_stary_plik_wola_o_aktualizacje(self):
        _dotknij(os.path.join(self.folder, "Wzory.xls"), wiek_dni=45)
        w = self.wiersz("Wzory.xls", prog_dni=30)
        self.assertEqual(w.ikona.text(), "!")
        self.assertIn("ZAKTUALIZUJ", w.stan.text())

    def test_brak_progu_nie_ostrzega(self):
        _dotknij(os.path.join(self.folder, "cokolwiek.xls"), wiek_dni=500)
        w = self.wiersz("cokolwiek.xls")
        self.assertEqual(w.ikona.text(), "✓")

    def test_pusty_folder_ostrzega(self):
        os.makedirs(os.path.join(self.folder, "PZ"), exist_ok=True)
        w = self.wiersz("PZ", katalog=True, wzorzec=(".pdf",))
        self.assertEqual(w.ikona.text(), "!")
        self.assertIn("pusty", w.stan.text())

    def test_folder_z_plikami_liczy_je(self):
        pz = os.path.join(self.folder, "PZ")
        os.makedirs(pz, exist_ok=True)
        _dotknij(os.path.join(pz, "a.pdf"))
        _dotknij(os.path.join(pz, "b.pdf"))
        w = self.wiersz("PZ", katalog=True, wzorzec=(".pdf",))
        self.assertEqual(w.ikona.text(), "✓")
        self.assertIn("2 plikow", w.stan.text())

    def test_folder_liczy_tylko_pasujace_rozszerzenia(self):
        pz = os.path.join(self.folder, "PZ")
        os.makedirs(pz, exist_ok=True)
        _dotknij(os.path.join(pz, "a.pdf"))
        _dotknij(os.path.join(pz, "notatka.txt"))
        w = self.wiersz("PZ", katalog=True, wzorzec=(".pdf",))
        self.assertIn("1 plik ", w.stan.text() + " ")


class TestSiatkaKart(unittest.TestCase):
    """
    Lista kontrolna jako siatka kart. Wczesniej kazda pozycja byla paskiem na
    cala szerokosc — tresci malo, pustego miejsca po prawej duzo, a komplet
    pozycji nie miescil sie bez przewijania.
    """

    def siatka(self, ile_kart, szerokosc, min_szerokosc=250):
        from PySide6.QtWidgets import QLabel
        s = W.SiatkaKart(min_szerokosc=min_szerokosc)
        for i in range(ile_kart):
            s.dodaj(QLabel(f"karta {i}"))
        s.resize(szerokosc, 400)
        # Widget nie jest pokazany, wiec Qt nie dostarcza resizeEvent —
        # przeliczamy uklad wprost. Sciezke przez zdarzenie sprawdza osobny test.
        s._ustaw_uklad()
        return s

    def kolumny_uzyte(self, siatka):
        pozycje = [siatka._siatka.getItemPosition(i)
                   for i in range(siatka._siatka.count())]
        return len({kol for _w, kol, _rs, _cs in pozycje})

    def test_szerokie_okno_daje_wiele_kolumn(self):
        s = self.siatka(7, 1250)
        self.assertGreaterEqual(self.kolumny_uzyte(s), 4)

    def test_waskie_okno_zwija_do_jednej_kolumny(self):
        s = self.siatka(7, 260)
        self.assertEqual(self.kolumny_uzyte(s), 1)

    def test_nie_wiecej_kolumn_niz_kart(self):
        s = self.siatka(2, 2000)
        self.assertEqual(self.kolumny_uzyte(s), 2)

    def test_wszystkie_karty_trafiaja_do_siatki(self):
        s = self.siatka(7, 1250)
        self.assertEqual(s._siatka.count(), 7)
        self.assertEqual(len(s.karty()), 7)

    def test_zmiana_szerokosci_przeklada_karty(self):
        s = self.siatka(7, 1250)
        szerokie = self.kolumny_uzyte(s)
        s.resize(260, 400)
        s._ustaw_uklad()
        self.assertLess(self.kolumny_uzyte(s), szerokie)

    def test_zdarzenie_zmiany_rozmiaru_przeklada_karty(self):
        """W dzialajacym oknie uklad odswieza sie sam, przez resizeEvent."""
        from PySide6.QtCore import QSize
        from PySide6.QtGui import QResizeEvent
        s = self.siatka(7, 260)
        self.assertEqual(self.kolumny_uzyte(s), 1)
        s.resize(1250, 400)
        s.resizeEvent(QResizeEvent(QSize(1250, 400), QSize(260, 400)))
        self.assertGreaterEqual(self.kolumny_uzyte(s), 4)

    def test_pusta_siatka_nie_wywala(self):
        s = W.SiatkaKart()
        s.resize(800, 200)
        s._ustaw_uklad()
        self.assertEqual(s._siatka.count(), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
