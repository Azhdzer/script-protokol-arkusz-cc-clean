# -*- coding: utf-8 -*-
"""
Skrypty zamykaja WYLACZNIE Excela, ktorego same uruchomily (cc_excel.py).

Zgloszenie: uruchomienie kroku zamykalo Excela uzytkownika. Przyczyna lezala
w kroku 2 — przy usuwaniu blokow protokolu wolal on
win32com.client.Dispatch('Excel.Application'), ktore NIE tworzy nowego Excela,
tylko PODPINA SIE pod otwarty. Potem ukrywal go (Visible=False), wylaczal
pytania (DisplayAlerts=False) i konczyl Quit() — zamykajac okna uzytkownika
bez pytania o niezapisane zmiany. Dwie inne funkcje mialy to samo jako zapas.

Do tego zawieszony przy duzym pliku Excel zostawal w tle i blokowal pliki;
jedynym wyjsciem bylo recznie zamykanie WSZYSTKICH Exceli. Teraz ubijany jest
tylko proces o PID naszej instancji — i tylko gdy po Quit() dalej zyje.

Testy z Excelem pomijane przy CC_TESTY_SZYBKIE=1.
"""

import os
import re
import subprocess
import sys
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import cc_excel
from wspolne import KORZEN

SZYBKIE = os.environ.get("CC_TESTY_SZYBKIE", "").strip() in ("1", "true", "tak")


def procesy_excela():
    import win32process
    return {p for p in win32process.EnumProcesses()
            if (cc_excel.nazwa_procesu(p) or "").lower() == "excel.exe"}


class TestBezpieczenstwoUbijania(unittest.TestCase):
    """Ubijamy po PID — wiec PID musi naprawde nalezec do Excela."""

    def test_nazwa_wlasnego_procesu(self):
        self.assertTrue(cc_excel.nazwa_procesu(os.getpid()).lower().startswith("python"))

    def test_wlasny_proces_zyje(self):
        self.assertTrue(cc_excel.proces_zyje(os.getpid()))

    def test_nieistniejacy_proces_nie_zyje(self):
        self.assertFalse(cc_excel.proces_zyje(0))
        self.assertIsNone(cc_excel.nazwa_procesu(0))

    def test_proces_inny_niz_excel_nie_jest_ubijany(self):
        """Gdyby Windows nadal numer PID innemu programowi — nie wolno go ruszyc."""
        p = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"],
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            self.assertFalse(cc_excel.ubij_excela(p.pid))
            time.sleep(0.3)
            self.assertIsNone(p.poll(), "proces nie-Excel zostal zakonczony")
        finally:
            p.kill()
            p.wait()

    def test_brak_pid_nic_nie_robi(self):
        self.assertTrue(cc_excel.dopilnuj_zamkniecia(None))


class TestZadenSkryptNiePodpinaSiePodExcelaUzytkownika(unittest.TestCase):
    """
    Straznik w zrodlach: 'Dispatch(' podpina sie pod otwartego Excela, wiec jego
    Quit() zamyka okna uzytkownika. W skryptach wolno tylko przez cc_excel.
    """

    SKRYPTY = ("generuj_obserwacje.py", "generuj_arkusze.py", "wypelnij_warunki.py",
               "analizuj_excele.py")

    def zrodlo(self, plik):
        with open(os.path.join(KORZEN, plik), encoding="utf-8") as f:
            return f.read()

    def test_brak_dispatch_do_excela(self):
        wzorzec = re.compile(r"""\bDispatch\(\s*['"]Excel\.Application""")
        for plik in self.SKRYPTY:
            with self.subTest(plik=plik):
                self.assertIsNone(wzorzec.search(self.zrodlo(plik)))

    def test_brak_gołego_quit_w_kroku_3(self):
        """app.quit() bez pilnowania zostawial zawieszone procesy w tle."""
        self.assertNotRegex(self.zrodlo("generuj_arkusze.py"), r"\bapp\.quit\(\)")

    def test_nowa_instancja_to_dispatchex(self):
        self.assertIn('DispatchEx("Excel.Application")', self.zrodlo("cc_excel.py"))


@unittest.skipIf(SZYBKIE, "CC_TESTY_SZYBKIE=1 — testy z Excelem pominiete")
class TestNaPrawdziwymExcelu(unittest.TestCase):

    def setUp(self):
        self.przed = procesy_excela()

    def tearDown(self):
        # Nasze ewentualne resztki — nigdy procesy sprzed testu.
        for pid in procesy_excela() - self.przed:
            cc_excel.ubij_excela(pid)

    def test_nowa_instancja_to_osobny_proces(self):
        xl = cc_excel.nowy_excel_com()
        try:
            pid = cc_excel.pid_excela_com(xl)
            self.assertIsNotNone(pid)
            self.assertNotIn(pid, self.przed, "podpieto sie pod istniejacego Excela")
        finally:
            cc_excel.zamknij_excel_com(xl)

    def test_zawieszona_instancja_ubita_a_inne_zostaja(self):
        xl = cc_excel.nowy_excel_com()
        pid = cc_excel.pid_excela_com(xl)
        import io
        from contextlib import redirect_stdout
        with redirect_stdout(io.StringIO()):
            # bez Quit() — tak wyglada zawieszony Excel
            self.assertTrue(cc_excel.dopilnuj_zamkniecia(pid, czekaj_s=1))
        del xl
        self.assertFalse(cc_excel.proces_zyje(pid))
        for p in self.przed:
            with self.subTest(pid=p):
                self.assertTrue(cc_excel.proces_zyje(p), "ubito Excela sprzed testu")

    def test_xlwings_zamyka_swoj_proces(self):
        import xlwings as xw
        app = xw.App(visible=False, add_book=False)
        pid = cc_excel.pid_excela_xlwings(app)
        cc_excel.zamknij_excel_xlwings(app)
        self.assertFalse(cc_excel.proces_zyje(pid))

    def test_usuwanie_blokow_protokolu_nie_rusza_innych_exceli(self):
        """Dokladnie ta sciezka zamykala Excela uzytkownika."""
        import shutil
        from wspolne import nowa_piaskownica
        import generuj_obserwacje as G
        folder = nowa_piaskownica("excel_wlasne_instancje")
        prot = os.path.join(folder, "protokol.xlsx")
        shutil.copy2(os.path.join(KORZEN, "xxx_LA_TH_2026 - protokół CC.xlsx"), prot)
        os.chmod(prot, 0o666)
        self.assertTrue(G._delete_rows_via_excel(prot, "Strona 3", 60, 5))
        koniec = time.time() + cc_excel.CZEKAJ_NA_ZAMKNIECIE_S + 5
        while time.time() < koniec and procesy_excela() - self.przed:
            time.sleep(0.5)
        self.assertEqual(procesy_excela() - self.przed, set(),
                         "nasz Excel zostal w tle")
        for p in self.przed:
            with self.subTest(pid=p):
                self.assertTrue(cc_excel.proces_zyje(p), "zamknieto Excela sprzed testu")


if __name__ == "__main__":
    unittest.main(verbosity=2)
