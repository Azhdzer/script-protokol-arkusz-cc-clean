# -*- coding: utf-8 -*-
"""
app_gui.py — panel sterujacy calym obiegiem wzorcowania (PySide6).

Zalozenie: wszystko, co dotad wymagalo edycji skryptu w Notatniku, ustawia sie
tutaj. Rejestr ustawien (typy, wartosci domyslne, opisy) mieszka w cc_config.py;
panel buduje z niego formularze automatycznie i przekazuje wartosci do skryptow
przez zmienne srodowiskowe.

Uklad okna:
  • gora        — folder roboczy,
  • lewa listwa — kroki obiegu (Przygotowanie / 1 Analiza / 2 Obserwacja /
                  3 Arkusze) ze statusem + "Uruchom caly obieg",
  • srodek      — formularz aktywnego kroku,
  • dol         — zywy log i lista plikow, ktore powstaly.

Uruchomienie (zrodlo):  .venv\\Scripts\\python.exe app_gui.py
Zamrozony .exe uruchamia workery jako podprocesy (patrz app_entry.py).
"""

import datetime
import os
import sys
import traceback

from PySide6.QtCore import Qt, QProcess, QProcessEnvironment, QTimer
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QFrame, QScrollArea, QStackedWidget, QFileDialog, QSplitter,
    QTabWidget, QMessageBox, QSizePolicy, QButtonGroup,
)

import cc_config as C
from cc_widgets import (
    Pole, ListaPlikow, WidokLogu, PanelWynikow, WierszKontrolny, SiatkaKart,
    nazwy_arkuszy,
    BG, CARD, STROKE, ACCENT, ACCENT2, ACCENT3, TEXT, MUTED, FIELD,
    OK_C, WARN_C, ERR_C,
)

if getattr(sys, "frozen", False):
    HERE = os.path.dirname(sys.executable)
    _BUNDLE = getattr(sys, "_MEIPASS", HERE)
else:
    HERE = os.path.dirname(os.path.abspath(__file__))
    _BUNDLE = HERE


def zasob(rel):
    return os.path.join(_BUNDLE, rel)


# Kroki uruchamialne: klucz -> (etykieta, worker CC_WORKER, plik skryptu)
URUCHAMIALNE = {
    "analiza": ("1 · Analiza logow",  "analiza",    "analizuj_excele.py"),
    "obs":     ("2 · Obserwacja",     "obserwacje", "generuj_obserwacje.py"),
    "ark":     ("3 · Arkusze i Word", "arkusze",    "generuj_arkusze.py"),
}

# Ustawienia kroku 2 pokazywane WEWNATRZ karty "Dane wejsciowe", tuz pod lista
# plikow TXT — razem odpowiadaja na pytanie "co bierzemy do tego pomiaru".
W_KARCIE_WEJSCIOWEJ = ("OBS_FOTO", "OBS_FOTO_ZRODLO")

# Symbole statusu kroku w lewej listwie.
STATUS = {
    "gotowy":   ("○", MUTED),
    "trwa":     ("◐", ACCENT),
    "ok":       ("✓", OK_C),
    "blad":     ("✕", ERR_C),
}


class Okno(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generator protokolow i arkuszy wzorcowania")
        self.resize(1320, 900)
        self.setMinimumSize(1060, 700)

        self.wartosci = C.wczytaj()
        if not os.path.isdir(str(self.wartosci.get("CC_FOLDER") or "")):
            self.wartosci["CC_FOLDER"] = HERE

        self.pola = {}          # env -> Pole
        self.statusy = {}       # klucz kroku -> QLabel z symbolem
        self.kontrolne = []     # wiersze listy kontrolnej
        self.proc = None
        self.kolejka = []       # kroki czekajace na uruchomienie
        self.biezacy = None

        # Zapis ustawien jest odroczony — bez tego kazde nacisniecie strzalki
        # w polu liczbowym zapisywaloby plik na dysk.
        self._zapis = QTimer(self)
        self._zapis.setSingleShot(True)
        self._zapis.setInterval(600)
        self._zapis.timeout.connect(self._zapisz)

        korzen = QWidget()
        self.setCentralWidget(korzen)
        zewn = QVBoxLayout(korzen)
        zewn.setContentsMargins(16, 14, 16, 14)
        zewn.setSpacing(12)

        zewn.addWidget(self._naglowek())

        podzial = QSplitter(Qt.Vertical)
        podzial.setChildrenCollapsible(False)

        srodek = QWidget()
        cialo = QHBoxLayout(srodek)
        cialo.setContentsMargins(0, 0, 0, 0)
        cialo.setSpacing(12)
        cialo.addWidget(self._listwa(), 0)
        cialo.addWidget(self._strony(), 1)
        podzial.addWidget(srodek)
        podzial.addWidget(self._dol())
        podzial.setStretchFactor(0, 3)
        podzial.setStretchFactor(1, 2)
        # Formularz kroku dostaje wiekszosc miejsca; log da sie rozciagnac
        # myszka, gdy bedzie potrzebny.
        podzial.setSizes([720, 190])
        zewn.addWidget(podzial, 1)

        self.setStyleSheet(self._qss())
        self._przelacz(0)
        self._odswiez_folder()

    # ── folder roboczy ───────────────────────────────────────────────────
    @property
    def folder(self):
        return str(self.wartosci.get("CC_FOLDER") or HERE)

    def _naglowek(self):
        w = QFrame()
        w.setObjectName("Karta")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(18, 13, 18, 13)
        lay.setSpacing(6)

        t = QLabel("Generator protokolow i arkuszy wzorcowania")
        t.setObjectName("H1")
        lay.addWidget(t)

        row = QHBoxLayout()
        row.setSpacing(8)
        et = QLabel("Folder roboczy:")
        et.setObjectName("OpisPola")
        self.lbl_folder = QLabel(self.folder)
        self.lbl_folder.setObjectName("Sciezka")
        self.lbl_folder.setTextInteractionFlags(Qt.TextSelectableByMouse)
        b1 = QPushButton("Zmien…")
        b1.setObjectName("Ghost")
        b1.setCursor(Qt.PointingHandCursor)
        b1.clicked.connect(self._wybierz_folder)
        b2 = QPushButton("Otworz w Eksploratorze")
        b2.setObjectName("Ghost")
        b2.setCursor(Qt.PointingHandCursor)
        b2.clicked.connect(lambda: os.startfile(self.folder)
                           if os.path.isdir(self.folder) else None)
        row.addWidget(et, 0)
        row.addWidget(self.lbl_folder, 1)
        row.addWidget(b1, 0)
        row.addWidget(b2, 0)
        lay.addLayout(row)
        return w

    def _wybierz_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Wybierz folder roboczy", self.folder)
        if d:
            self.wartosci["CC_FOLDER"] = d
            self._odswiez_folder()
            self._zapisz()

    def _folder_logow(self):
        """Folder wejsciowy kroku 1 — wzgledny liczony od folderu roboczego."""
        wej = str(self.wartosci.get("ANL_INPUT", "excel_do_analizy"))
        return wej if os.path.isabs(wej) else os.path.join(self.folder, wej)

    def _odswiez_folder(self):
        """Po zmianie folderu przeladowuje wszystkie listy plikow i kontrolki."""
        self.lbl_folder.setText(self.folder)
        for pole in self.pola.values():
            pole.odswiez_pliki()
        self.lista_txt.odswiez(zachowaj=self.wartosci.get("OBS_TXT_FILES") or [])
        self.lista_logow.odswiez(zachowaj=self.wartosci.get("ANL_PLIKI") or [])
        for wiersz in self.kontrolne:
            wiersz.odswiez()

    # ── lewa listwa krokow ───────────────────────────────────────────────
    def _listwa(self):
        w = QFrame()
        w.setObjectName("Karta")
        w.setFixedWidth(272)
        lay = QVBoxLayout(w)
        lay.setContentsMargins(12, 14, 12, 14)
        lay.setSpacing(6)

        lay.addWidget(self._podpis("KROKI OBIEGU"))
        self.grupa = QButtonGroup(self)
        self.grupa.setExclusive(True)

        pozycje = [(k, e) for k, e, _o in C.KROKI] + [("zaawansowane", "Zaawansowane")]
        for i, (klucz, etykieta) in enumerate(pozycje):
            row = QWidget()
            rl = QHBoxLayout(row)
            rl.setContentsMargins(0, 0, 0, 0)
            rl.setSpacing(6)

            znak = QLabel("")
            znak.setFixedWidth(16)
            znak.setAlignment(Qt.AlignCenter)
            if klucz in URUCHAMIALNE:
                self.statusy[klucz] = znak
                self._status(klucz, "gotowy")
            rl.addWidget(znak, 0)

            b = QPushButton(etykieta)
            b.setObjectName("Krok")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.setChecked(i == 0)
            b.clicked.connect(lambda _c, n=i: self._przelacz(n))
            self.grupa.addButton(b)
            rl.addWidget(b, 1)
            lay.addWidget(row)

        lay.addSpacing(10)
        lay.addWidget(self._linia())
        lay.addSpacing(6)

        self.btn_obieg = QPushButton("▶▶  Uruchom caly obieg")
        self.btn_obieg.setObjectName("Glowny")
        self.btn_obieg.setMinimumHeight(46)
        self.btn_obieg.setCursor(Qt.PointingHandCursor)
        self.btn_obieg.setToolTip("Analiza logow  →  Obserwacja i protokol  →  "
                                  "Arkusze i swiadectwa.\nPrzerywa sie na pierwszym bledzie.")
        self.btn_obieg.clicked.connect(lambda: self._uruchom(["analiza", "obs", "ark"]))
        lay.addWidget(self.btn_obieg)

        self.btn_stop = QPushButton("Przerwij")
        self.btn_stop.setObjectName("Ghost")
        self.btn_stop.setCursor(Qt.PointingHandCursor)
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._przerwij)
        lay.addWidget(self.btn_stop)

        lay.addStretch(1)

        self.lbl_stan = QLabel("Gotowy.")
        self.lbl_stan.setObjectName("OpisPola")
        self.lbl_stan.setWordWrap(True)
        lay.addWidget(self.lbl_stan)
        return w

    def _status(self, klucz, stan):
        znak, kolor = STATUS[stan]
        et = self.statusy.get(klucz)
        if et:
            et.setText(znak)
            et.setStyleSheet(f"color:{kolor}; font-size:14px; font-weight:700;")

    # ── strony krokow ────────────────────────────────────────────────────
    def _strony(self):
        self.stos = QStackedWidget()
        self.stos.addWidget(self._strona_przygotowanie())
        self.stos.addWidget(self._strona_analiza())
        self.stos.addWidget(self._strona_obserwacja())
        self.stos.addWidget(self._strona_arkusze())
        self.stos.addWidget(self._strona_zaawansowane())
        return self.stos

    def _przelacz(self, nr):
        self.stos.setCurrentIndex(nr)
        if nr == 0:
            for wiersz in self.kontrolne:
                wiersz.odswiez()
        elif nr == 1:
            # Folder wejsciowy mogl sie zmienic w ustawieniach — przeladuj liste,
            # zachowujac dotychczasowy wybor plikow.
            self.lista_logow.odswiez(zachowaj=self.wartosci.get("ANL_PLIKI") or [])
        elif nr == 2:
            self.lista_txt.odswiez()

    def _przewijalna(self):
        """Zwraca (scroll, layout) — kazda strona kroku jest przewijalna."""
        scroll = QScrollArea()
        scroll.setObjectName("Przewijanie")
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        wnetrze = QWidget()
        scroll.setWidget(wnetrze)
        lay = QVBoxLayout(wnetrze)
        lay.setContentsMargins(2, 2, 12, 2)
        lay.setSpacing(12)
        return scroll, lay

    def _karta(self, tytul=None):
        k = QFrame()
        k.setObjectName("Karta")
        lay = QVBoxLayout(k)
        lay.setContentsMargins(18, 15, 18, 16)
        lay.setSpacing(11)
        if tytul:
            lay.addWidget(self._podpis(tytul))
        return k, lay

    def _podpis(self, tekst):
        w = QLabel(tekst.upper())
        w.setObjectName("Podpis")
        return w

    def _linia(self):
        w = QFrame()
        w.setObjectName("Linia")
        w.setFixedHeight(1)
        return w

    def _naglowek_kroku(self, klucz):
        """
        Szapka strony kroku: nazwa, opis i PRZYCISK URUCHOMIENIA.

        Przycisk siedzi tutaj, a nie na dole formularza, bo strona kroku bywa
        dluga (kilkanascie kart ustawien) — schowany pod nimi wymagalby
        przewijania za kazdym razem.
        """
        etykieta, opis = next(((e, o) for k, e, o in C.KROKI if k == klucz),
                              ("", ""))
        karta, lay = self._karta()
        row = QHBoxLayout()
        row.setSpacing(14)

        kol = QVBoxLayout()
        kol.setSpacing(3)
        t = QLabel(etykieta)
        t.setObjectName("H2")
        o = QLabel(opis)
        o.setObjectName("OpisPola")
        o.setWordWrap(True)
        kol.addWidget(t)
        kol.addWidget(o)
        row.addLayout(kol, 1)

        if klucz in URUCHAMIALNE:
            b = QPushButton(f"▶  Uruchom krok {etykieta.split(' ·')[0]}")
            b.setObjectName("Glowny")
            b.setMinimumHeight(46)
            b.setMinimumWidth(210)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _c, k=klucz: self._uruchom([k]))
            setattr(self, f"btn_{klucz}", b)
            row.addWidget(b, 0, Qt.AlignVCenter)

        lay.addLayout(row)
        return karta

    def _strona_kroku(self, klucz):
        """Zwraca (strona, layout_tresci). Szapka nie przewija sie z trescia."""
        strona = QWidget()
        zewn = QVBoxLayout(strona)
        zewn.setContentsMargins(0, 0, 0, 0)
        zewn.setSpacing(12)
        zewn.addWidget(self._naglowek_kroku(klucz))
        scroll, lay = self._przewijalna()
        zewn.addWidget(scroll, 1)
        return strona, lay

    # ── budowa pol z rejestru ────────────────────────────────────────────
    def _zbuduj_pole(self, ust):
        """Tworzy Pole, rejestruje je i podpina zapis ustawien."""
        pole = Pole(ust, self.wartosci.get(ust.env, ust.domyslna),
                    folder_cb=lambda: self.folder)
        pole.zmienione.connect(self._zmiana)
        self.pola[ust.env] = pole
        return pole

    def _dodaj_pola(self, lay, ustawienia):
        """Wstawia karty pogrupowane wg `grupa`, zachowujac kolejnosc rejestru."""
        grupy = []
        for u in ustawienia:
            if u.typ == "pliki":
                continue          # obsluzone osobno (ListaPlikow)
            if not grupy or grupy[-1][0] != u.grupa:
                grupy.append((u.grupa, []))
            grupy[-1][1].append(u)

        for nazwa, lista in grupy:
            karta, kl = self._karta(nazwa)
            for u in lista:
                kl.addWidget(self._zbuduj_pole(u))
            lay.addWidget(karta)

    # ── strona: przygotowanie ────────────────────────────────────────────
    def _strona_przygotowanie(self):
        scroll, lay = self._przewijalna()

        karta, kl = self._karta()
        t = QLabel("Przygotowanie")
        t.setObjectName("H2")
        o = QLabel("Zanim uruchomisz obieg, upewnij sie, ze ponizsze pliki sa "
                   "aktualne. Skrypty ich nie sprawdzaja — wygeneruja swiadectwa "
                   "ze starych danych bez slowa protestu.")
        o.setObjectName("OpisPola")
        o.setWordWrap(True)
        kl.addWidget(t)
        kl.addWidget(o)
        lay.addWidget(karta)

        karta2, kl2 = self._karta("Lista kontrolna")
        rok = datetime.date.today().year

        def sciezka(wzgledna):
            return (wzgledna if os.path.isabs(wzgledna)
                    else os.path.join(self.folder, wzgledna))

        def plik_czujnika():
            nr = self.wartosci.get("GEN_NR_POM", 9)
            model = self.wartosci.get("GEN_MODEL_CZUJ", "MX1101-02")
            return sciezka(f"Pom. nr {nr} ({model}) - {rok}.xlsx")

        pozycje = [
            ("Logi z przyrzadow (DUT)",
             "Wgraj pliki z loggerow do tego folderu — krok 1 zamieni je na wyniki/*.xlsx.",
             lambda: sciezka(str(self.wartosci.get("ANL_INPUT", "excel_do_analizy"))),
             None, True, None),
            ("PZ — Potwierdzenia zamowienia (PDF)",
             "Z nich powstaje tabela przyrzadow na Stronie 2 protokolu.",
             lambda: sciezka(str(self.wartosci.get("CC_PZ_FOLDER", "PZ"))),
             None, True, (".pdf",)),
            ("Pliki TXT multimetru",
             "Pomiar z komory. Zaznaczysz je w kroku 2.",
             lambda: self.folder, None, True, (".txt",)),
            ("Czujnik srodowiskowy — Pom. nr N",
             "Warunki srodowiskowe do Etapu 7. Pobierz swiezy odczyt z rejestratora.",
             plik_czujnika, 7, False, None),
            ("Zestawienie wzorcowanych przyrzadow.xlsx",
             "Zrodlo rozdzielczosci t/RH dla kolumn K/L Strony 2.",
             lambda: sciezka(str(self.wartosci.get(
                 "CC_ZESTAWIENIE", "Zestawienie wzorcowanych przyrządów.xlsx"))),
             30, False, None),
            ("Wzory.xls",
             "Plik linkowany — bez niego formuly kalibracyjne sie nie policza. "
             "Skopiuj swiezy z \\\\plum4\\LabPomiarowe.",
             lambda: sciezka("Wzory.xls"), 30, False, None),
            ("Obliczenia tdp, RH, C.xls",
             "Drugi plik linkowany, tak samo z \\\\plum4\\LabPomiarowe.",
             lambda: sciezka("Obliczenia tdp, RH, C.xls"), 30, False, None),
        ]
        # Karty w siatce, a nie paski na cala szerokosc — pozycje sa krotkie,
        # wiec w rzedzie miesci sie ich kilka i widac caly komplet bez przewijania.
        siatka = SiatkaKart(min_szerokosc=250)
        for tytul, opis, cb, prog, katalog, wzor in pozycje:
            wiersz = WierszKontrolny(tytul, opis, cb, prog_dni=prog,
                                     katalog=katalog, wzorzec=wzor)
            self.kontrolne.append(wiersz)
            siatka.dodaj(wiersz)
        kl2.addWidget(siatka)

        odsw = QPushButton("Sprawdz ponownie")
        odsw.setObjectName("Ghost")
        odsw.setCursor(Qt.PointingHandCursor)
        odsw.clicked.connect(lambda: [w.odswiez() for w in self.kontrolne])
        kl2.addWidget(odsw)
        lay.addWidget(karta2)

        # Folder roboczy (jedyne ustawienie tego kroku) ma wlasne miejsce w
        # naglowku okna — nie dublujemy go tutaj jako pola formularza.
        lay.addStretch(1)
        return scroll

    # ── strona: analiza ──────────────────────────────────────────────────
    def _strona_analiza(self):
        strona, lay = self._strona_kroku("analiza")

        u_logi = C.WG_ENV["ANL_PLIKI"]
        karta, kl = self._karta(u_logi.grupa)
        et = QLabel(u_logi.etykieta)
        et.setObjectName("EtykietaPola")
        opis = QLabel(u_logi.opis)
        opis.setObjectName("OpisPola")
        opis.setWordWrap(True)
        kl.addWidget(et)
        kl.addWidget(opis)
        self.lista_logow = ListaPlikow(tuple(u_logi.wzorzec), self._folder_logow)
        self.lista_logow.zmienione.connect(self._zmiana)
        kl.addWidget(self.lista_logow)
        lay.addWidget(karta)

        self._dodaj_pola(lay, C.dla_kroku("analiza", C.PODSTAWOWY))
        lay.addStretch(1)
        return strona

    # ── strona: obserwacja ───────────────────────────────────────────────
    def _strona_obserwacja(self):
        strona, lay = self._strona_kroku("obs")

        u_txt = C.WG_ENV["OBS_TXT_FILES"]
        karta, kl = self._karta(u_txt.grupa)
        et = QLabel(u_txt.etykieta)
        et.setObjectName("EtykietaPola")
        op = QLabel(u_txt.opis)
        op.setObjectName("OpisPola")
        op.setWordWrap(True)
        kl.addWidget(et)
        kl.addWidget(op)
        self.lista_txt = ListaPlikow((".txt",), lambda: self.folder)
        self.lista_txt.zmienione.connect(self._zmiana)
        kl.addWidget(self.lista_txt)

        # Zdjecia siedza w TEJ SAMEJ karcie co wybor plikow, a nie w osobnej
        # nizej: to czesc odpowiedzi na pytanie "co bierzemy do tego pomiaru",
        # a przy nizszym oknie osobna karta wypadala poza widoczny obszar.
        kl.addWidget(self._linia())
        for env in W_KARCIE_WEJSCIOWEJ:
            kl.addWidget(self._zbuduj_pole(C.WG_ENV[env]))
        lay.addWidget(karta)

        self._dodaj_pola(lay, [u for u in C.dla_kroku("obs", C.PODSTAWOWY)
                               if u.env not in W_KARCIE_WEJSCIOWEJ])
        lay.addStretch(1)
        return strona

    # ── strona: arkusze ──────────────────────────────────────────────────
    def _strona_arkusze(self):
        strona, lay = self._strona_kroku("ark")
        self._dodaj_pola(lay, C.dla_kroku("ark", C.PODSTAWOWY))
        lay.addStretch(1)
        return strona

    # ── strona: zaawansowane ─────────────────────────────────────────────
    def _strona_zaawansowane(self):
        scroll, lay = self._przewijalna()

        karta, kl = self._karta()
        t = QLabel("Zaawansowane")
        t.setObjectName("H2")
        o = QLabel("Ustawienia, ktore zmienia sie rzadko: progi analizy okna, "
                   "korekta zegara loggerow, filtr kolorow Strony 3, sciezki "
                   "serwerowe, mapowanie CC-04. Zmieniaj swiadomie — wplywaja na "
                   "wynik wzorcowania.")
        o.setObjectName("OpisPola")
        o.setWordWrap(True)
        kl.addWidget(t)
        kl.addWidget(o)
        lay.addWidget(karta)

        for klucz, etykieta, _opis in C.KROKI:
            zaaw = C.dla_kroku(klucz, C.ZAAWANSOWANY)
            if not zaaw:
                continue
            naglowek = QLabel(etykieta)
            naglowek.setObjectName("Sekcja")
            lay.addWidget(naglowek)
            self._dodaj_pola(lay, zaaw)

        przywroc = QPushButton("Przywroc wartosci domyslne")
        przywroc.setObjectName("Ghost")
        przywroc.setCursor(Qt.PointingHandCursor)
        przywroc.clicked.connect(self._przywroc_domyslne)
        lay.addWidget(przywroc)
        lay.addStretch(1)
        return scroll

    def _przywroc_domyslne(self):
        odp = QMessageBox.question(
            self, "Przywrocic domyslne?",
            "Wszystkie ustawienia wroca do wartosci wpisanych w kodzie skryptow.\n"
            "Folder roboczy i wybor plikow TXT zostana zachowane.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if odp != QMessageBox.Yes:
            return
        dom = C.domyslne()
        zachowaj = ("CC_FOLDER", "OBS_TXT_FILES")
        for env, pole in self.pola.items():
            if env in zachowaj:
                continue
            self.wartosci[env] = dom[env]
            pole.ustaw(dom[env])
        self._zapisz()
        self.lbl_stan.setText("Przywrocono ustawienia domyslne.")

    # ── dol: log + wyniki ────────────────────────────────────────────────
    def _dol(self):
        karta = QFrame()
        karta.setObjectName("Karta")
        lay = QVBoxLayout(karta)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(8)

        self.zakladki = QTabWidget()
        self.zakladki.setObjectName("Zakladki")

        self.log = WidokLogu()
        self.zakladki.addTab(self.log, "Log")

        self.wyniki = PanelWynikow(lambda: self.folder)
        self.zakladki.addTab(self.wyniki, "Utworzone pliki")
        lay.addWidget(self.zakladki, 1)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.pasek = QLabel("Gotowy.")
        self.pasek.setObjectName("OpisPola")
        b_zapisz = QPushButton("Zapisz log do pliku")
        b_zapisz.setObjectName("Ghost")
        b_zapisz.setCursor(Qt.PointingHandCursor)
        b_zapisz.clicked.connect(self._zapisz_log)
        b_czysc = QPushButton("Wyczysc log")
        b_czysc.setObjectName("Ghost")
        b_czysc.setCursor(Qt.PointingHandCursor)
        b_czysc.clicked.connect(self.log.clear)
        row.addWidget(self.pasek, 1)
        row.addWidget(b_czysc, 0)
        row.addWidget(b_zapisz, 0)
        lay.addLayout(row)

        karta.setMinimumHeight(180)
        karta.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        return karta

    def _zapisz_log(self):
        domyslna = os.path.join(
            self.folder, f"log_{datetime.datetime.now():%Y-%m-%d_%H%M}.txt")
        sciezka, _ = QFileDialog.getSaveFileName(
            self, "Zapisz log", domyslna, "Plik tekstowy (*.txt)")
        if not sciezka:
            return
        try:
            with open(sciezka, "w", encoding="utf-8") as f:
                f.write(self.log.toPlainText())
            self.pasek.setText(f"Log zapisany: {os.path.basename(sciezka)}")
        except OSError as e:
            QMessageBox.warning(self, "Nie udalo sie zapisac", str(e))

    # ── ustawienia ───────────────────────────────────────────────────────
    def _zmiana(self):
        # Wartosci zbieramy od razu (inne widoki czytaja self.wartosci), a sam
        # zapis na dysk odkladamy — inaczej kazde klikniecie strzalki w polu
        # liczbowym oznaczaloby zapis pliku.
        self._zbierz()
        self._zapis.start()

    def _zbierz(self):
        for env, pole in self.pola.items():
            self.wartosci[env] = pole.wartosc()
        self.wartosci["OBS_TXT_FILES"] = self.lista_txt.zaznaczone()
        self.wartosci["ANL_PLIKI"] = self.lista_logow.zaznaczone()

    def _zapisz(self):
        self._zbierz()
        blad = C.zapisz(self.wartosci)
        if blad:
            self.pasek.setText(f"Nie udalo sie zapisac ustawien: {blad}")

    # ── walidacja przed uruchomieniem ────────────────────────────────────
    def _problemy(self, kroki):
        """Lista czytelnych zastrzezen; pusta = wszystko na miejscu."""
        p = []
        if not os.path.isdir(self.folder):
            return [f"Folder roboczy nie istnieje: {self.folder}"]

        if "analiza" in kroki:
            wej = str(self.wartosci.get("ANL_INPUT", "excel_do_analizy"))
            pelna = wej if os.path.isabs(wej) else os.path.join(self.folder, wej)
            rozsz = (".csv", ".xls", ".xlsx", ".txt", ".pdf", ".log")
            if not os.path.isdir(pelna):
                p.append(f"Brak folderu z logami: {pelna}")
            else:
                obecne = [f for f in os.listdir(pelna) if f.lower().endswith(rozsz)]
                if not obecne:
                    p.append(f"Folder z logami jest pusty: {pelna}")
                else:
                    wybrane = self.wartosci.get("ANL_PLIKI") or []
                    brak = [n for n in wybrane if n not in obecne]
                    if brak:
                        p.append("Zaznaczone pliki do analizy nie istnieja juz w "
                                 f"folderze: {', '.join(brak)}")

        if "obs" in kroki:
            if not self.lista_txt.zaznaczone():
                p.append("Nie zaznaczono zadnego pliku TXT multimetru (krok 2).")
            for env in ("OBS_TEMPLATE", "OBS_CC04_TEMPLATE",
                        "OBS_PROT_CC", "OBS_PROT_CC04", "CC_ZESTAWIENIE"):
                pole = self.pola.get(env)
                if pole and pole.brak_pliku():
                    p.append(f"Nie znaleziono pliku: {pole.ust.etykieta} "
                             f"→ '{pole.wartosc() or '(nie wybrano)'}'")

        if "ark" in kroki:
            for env in ("CC_PROTOKOL", "CC_SZABLON"):
                pole = self.pola.get(env)
                if pole and pole.brak_pliku():
                    p.append(f"Nie znaleziono pliku: {pole.ust.etykieta} "
                             f"→ '{pole.wartosc() or '(nie wybrano)'}'")
            p += self._problemy_zawartosci()
            if self.wartosci.get("GEN_WORD"):
                for env in ("GEN_WORD_TEMP", "GEN_WORD_RH", "GEN_WORD_MIX"):
                    pole = self.pola.get(env)
                    if pole and pole.brak_pliku():
                        p.append(f"Brak szablonu Word: '{pole.wartosc()}'")
            for nazwa in str(self.wartosci.get("GEN_LINKOWANE", "")).split(";"):
                nazwa = nazwa.strip()
                if nazwa and not os.path.exists(os.path.join(self.folder, nazwa)):
                    p.append(f"Brak pliku linkowanego w folderze roboczym: {nazwa} "
                             f"(formuly kalibracyjne moga sie nie policzyc)")
        return p

    def _problemy_zawartosci(self):
        """
        Sprawdza, czy wybrane pliki SA tym, za co sie podaja.

        Sama obecnosc pliku nie wystarcza: latwo wskazac arkusz obserwacji tam,
        gdzie ma byc protokol. Skrypt ruszylby, otworzyl Excela na dwie minuty
        i zrobil bezsensowne kopie. Rozpoznajemy po zakladkach — czytane wprost
        z archiwum, wiec sprawdzenie jest natychmiastowe.
        """
        wymagane = {
            "CC_PROTOKOL": (("Strona 2", "Strona 3"),
                            "to nie wyglada na protokol (brak Strony 2 i 3)"),
            "CC_SZABLON":  (("Wyniki",),
                            "to nie wyglada na arkusz obliczeniowy (brak zakladki 'Wyniki')"),
        }
        problemy = []
        for env, (zakladki, komunikat) in wymagane.items():
            pole = self.pola.get(env)
            if not pole or pole.brak_pliku():
                continue          # brak pliku zglosza inne testy
            nazwa = pole.wartosc()
            sciezka = (nazwa if os.path.isabs(nazwa)
                       else os.path.join(self.folder, nazwa))
            obecne = nazwy_arkuszy(sciezka)
            if not obecne:
                continue          # nieczytelny xlsx — niech zglosi to sam skrypt
            brakujace = [z for z in zakladki if z not in obecne]
            if brakujace:
                problemy.append(
                    f"{pole.ust.etykieta}: '{nazwa}' — {komunikat}. "
                    f"Znalezione zakladki: {', '.join(obecne[:5])}"
                    + (" …" if len(obecne) > 5 else ""))
        return problemy

    # ── uruchamianie ─────────────────────────────────────────────────────
    def _uruchom(self, kroki):
        if self.proc and self.proc.state() != QProcess.NotRunning:
            self.lbl_stan.setText("Trwa juz inne zadanie — poczekaj albo przerwij.")
            return

        self._zapisz()

        problemy = self._problemy(kroki)
        if problemy:
            tresc = "\n".join(f"  •  {x}" for x in problemy)
            odp = QMessageBox.warning(
                self, "Sprawdz dane wejsciowe",
                f"Znaleziono {len(problemy)} zastrzezen:\n\n{tresc}\n\n"
                "Uruchomic mimo to?",
                QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel)
            if odp != QMessageBox.Yes:
                return

        for k in kroki:
            self._status(k, "gotowy")
        self.kolejka = list(kroki)
        self.log.clear()
        self.zakladki.setCurrentIndex(0)
        self._zajety(True)
        self._nastepny()

    def _nastepny(self):
        if not self.kolejka:
            self._zajety(False)
            self.lbl_stan.setText("Gotowy.")
            return

        klucz = self.kolejka.pop(0)
        etykieta, worker, skrypt = URUCHAMIALNE[klucz]
        self.biezacy = klucz
        self._status(klucz, "trwa")
        self.lbl_stan.setText(f"Trwa: {etykieta}")
        self.pasek.setText(f"Trwa: {etykieta}")

        srodowisko = QProcessEnvironment.systemEnvironment()
        for nazwa, wartosc in C.do_env(self.wartosci).items():
            srodowisko.insert(nazwa, wartosc)
        srodowisko.insert("PYTHONIOENCODING", "utf-8")
        srodowisko.insert("PYTHONUNBUFFERED", "1")
        srodowisko.insert("CC_WORKER", worker)

        self.wyniki.zapamietaj_stan()

        self.proc = QProcess(self)
        self.proc.setProcessEnvironment(srodowisko)
        self.proc.setWorkingDirectory(self.folder)
        self.proc.setProcessChannelMode(QProcess.MergedChannels)
        self.proc.readyReadStandardOutput.connect(self._wyjscie)
        self.proc.finished.connect(lambda kod, _s: self._koniec(klucz, etykieta, kod))
        self.proc.errorOccurred.connect(lambda _e: self._blad_startu(klucz, etykieta))

        self.log.dopisz(f"\n{'═' * 72}\n  {etykieta}\n{'═' * 72}\n")

        if getattr(sys, "frozen", False):
            self.proc.start(sys.executable, [])
        else:
            sciezka = os.path.join(HERE, skrypt)
            if not os.path.exists(sciezka):
                self.log.dopisz(f"!!! BLAD: brak pliku skryptu {sciezka}\n")
                self._koniec(klucz, etykieta, 1)
                return
            self.proc.start(sys.executable, ["-u", sciezka])

    def _wyjscie(self):
        dane = bytes(self.proc.readAllStandardOutput())
        self.log.dopisz(dane.decode("utf-8", errors="replace"))

    def _blad_startu(self, klucz, etykieta):
        if self.proc and self.proc.error() == QProcess.FailedToStart:
            self.log.dopisz(f"!!! BLAD: nie udalo sie uruchomic '{etykieta}'.\n")

    def _koniec(self, klucz, etykieta, kod):
        ok = (kod == 0)
        self._status(klucz, "ok" if ok else "blad")
        self.log.dopisz(f"\n[{etykieta}: {'zakonczono OK' if ok else f'BLAD, kod {kod}'}]\n")

        self.wyniki.pokaz_zmiany(etykieta)
        self.zakladki.setCurrentIndex(1 if ok else 0)
        self.pasek.setText(f"{etykieta}: {'gotowe' if ok else f'blad (kod {kod})'}")

        # Nowe pliki wynikowe moga zmienic listy wyboru (np. swiezy protokol).
        for pole in self.pola.values():
            pole.odswiez_pliki()

        if not ok:
            pominiete = len(self.kolejka)
            self.kolejka = []
            if pominiete:
                self.log.dopisz(f"[obieg przerwany — pominieto {pominiete} krok(i)]\n")
            self.lbl_stan.setText(f"{etykieta}: blad (kod {kod}).")
            self._zajety(False)
            return

        self._nastepny()

    def _przerwij(self):
        self.kolejka = []
        if self.proc and self.proc.state() != QProcess.NotRunning:
            self.proc.kill()
            self.log.dopisz("\n[przerwano przez uzytkownika]\n")
        self._zajety(False)
        self.lbl_stan.setText("Przerwano.")

    def _zajety(self, tak):
        self.btn_obieg.setEnabled(not tak)
        self.btn_stop.setEnabled(tak)
        for klucz in URUCHAMIALNE:
            b = getattr(self, f"btn_{klucz}", None)
            if b:
                b.setEnabled(not tak)

    # ── zamkniecie ───────────────────────────────────────────────────────
    def closeEvent(self, e):
        if self.proc and self.proc.state() != QProcess.NotRunning:
            odp = QMessageBox.question(
                self, "Zadanie w toku",
                "Trwa uruchomiony krok. Zamknac panel i przerwac go?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if odp != QMessageBox.Yes:
                e.ignore()
                return
            self.proc.kill()
        self._zapisz()
        super().closeEvent(e)

    # ── styl ─────────────────────────────────────────────────────────────
    def _qss(self):
        return f"""
        QWidget {{ background:{BG}; color:{TEXT};
            font-family:'Segoe UI Variable','Segoe UI',sans-serif; font-size:13px; }}
        /* Etykiety musza byc przezroczyste — inaczej dziedzicza tlo okna i
           rysuja szare paski na bialych kartach. */
        QLabel {{ background:transparent; }}
        QCheckBox {{ background:transparent; }}
        #Karta {{ background:{CARD}; border:1px solid {STROKE}; border-radius:10px; }}
        #H1 {{ font-size:19px; font-weight:700; }}
        #H2 {{ font-size:16px; font-weight:700; }}
        #Podpis {{ font-size:11px; font-weight:700; color:{ACCENT3}; letter-spacing:1px; }}
        #Sekcja {{ font-size:13px; font-weight:700; color:{ACCENT3};
            padding:6px 2px 0 2px; }}
        #Sciezka {{ color:{MUTED}; font-size:11px; }}
        #OpisPola {{ color:{MUTED}; font-size:11px; }}
        #EtykietaPola {{ font-weight:600; }}
        #Linia {{ background:{STROKE}; }}
        #Przewijanie {{ border:none; background:transparent; }}
        #Przewijanie > QWidget > QWidget {{ background:transparent; }}

        /* Przelacznik rysuje sie sam (cc_widgets.Przelacznik) — QSS ustawia mu
           tylko czcionke. Reguly ::indicator celowo nie ma: arkusz stylow nie
           narysuje ruchomej galki, a wlaczylby wlasne malowanie wskaznika. */
        QCheckBox#Toggle {{ padding:4px 2px; font-weight:600; }}

        QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox {{ background:{FIELD};
            border:1px solid #D6DCE4; border-radius:7px; padding:6px 8px; min-height:20px; }}
        QLineEdit:hover, QComboBox:hover, QDoubleSpinBox:hover, QSpinBox:hover {{
            border:1px solid {ACCENT2}; }}
        QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QSpinBox:focus {{
            border:1px solid {ACCENT}; }}
        QComboBox::drop-down {{ border:none; width:22px; }}
        QComboBox QAbstractItemView {{ background:{CARD}; border:1px solid {STROKE};
            selection-background-color:{ACCENT}; selection-color:white; outline:0; }}

        QPushButton#Glowny {{ background:{ACCENT}; color:white; border:none;
            border-radius:8px; font-size:14px; font-weight:700; padding:9px 26px; }}
        QPushButton#Glowny:hover {{ background:{ACCENT2}; }}
        QPushButton#Glowny:pressed {{ background:{ACCENT3}; }}
        QPushButton#Glowny:disabled {{ background:#9DBBDC; }}
        QPushButton#Ghost {{ background:transparent; border:1px solid {STROKE};
            border-radius:7px; padding:6px 12px; color:{MUTED}; font-weight:600; }}
        QPushButton#Ghost:hover {{ background:#F3F6FC; border-color:{ACCENT2}; color:{ACCENT}; }}
        QPushButton#Ghost:disabled {{ color:#B7BDC6; border-color:#EDEFF2; }}
        QPushButton#Probka {{ border:1px solid #B6C2D1; border-radius:5px; }}

        QPushButton#Krok {{ background:transparent; border:1px solid transparent;
            border-radius:8px; padding:9px 12px; text-align:left; font-weight:600;
            color:{TEXT}; }}
        QPushButton#Krok:hover {{ background:#EAF3FB; }}
        QPushButton#Krok:checked {{ background:{ACCENT}; color:white; }}

        QFrame#Kontrolny {{ background:{FIELD}; border:1px solid {STROKE};
            border-radius:8px; }}
        #KontrolnyTytul {{ font-weight:700; }}

        QListWidget#ListaPlikow {{ background:{FIELD}; border:1px solid #D6DCE4;
            border-radius:8px; padding:4px; }}
        /* Celowo BRAK reguly ::item — kazda z nich przelacza Qt na rysowanie
           elementow przez arkusz stylow, a wtedy tlo ustawiane per wiersz
           (ListaPlikow._odswiez_wyglad) przestaje dzialac. Wysokosc wiersza
           daje setSizeHint. */
        /* Powiekszony kwadracik — domyslne 13 px trudno trafic przy wierszu 50 px.
           Klikalny jest i tak caly wiersz (patrz _ListaKlikalna). */
        QListWidget#ListaPlikow::indicator {{ width:20px; height:20px;
            border-radius:5px; border:2px solid #9AA7B5; background:{CARD};
            margin-left:6px; margin-right:4px; }}
        QListWidget#ListaPlikow::indicator:hover {{ border-color:{ACCENT}; }}
        QListWidget#ListaPlikow::indicator:checked {{ background:{ACCENT};
            border-color:{ACCENT}; }}

        QTreeWidget#Wyniki {{ background:{FIELD}; border:1px solid #D6DCE4;
            border-radius:8px; alternate-background-color:#F5F7FA; }}
        QTreeWidget#Wyniki::item {{ padding:4px 2px; }}
        QHeaderView::section {{ background:{CARD}; border:none;
            border-bottom:1px solid {STROKE}; padding:6px; font-weight:700;
            color:{ACCENT3}; }}
        QTableWidget {{ background:{FIELD}; border:1px solid #D6DCE4;
            border-radius:8px; gridline-color:{STROKE}; }}

        QTabWidget#Zakladki::pane {{ border:1px solid {STROKE}; border-radius:8px;
            top:-1px; }}
        QTabBar::tab {{ background:transparent; padding:7px 18px; margin-right:2px;
            border-top-left-radius:8px; border-top-right-radius:8px;
            color:{MUTED}; font-weight:600; }}
        QTabBar::tab:selected {{ background:{CARD}; color:{ACCENT};
            border:1px solid {STROKE}; border-bottom:1px solid {CARD}; }}
        QTabBar::tab:hover:!selected {{ color:{ACCENT}; }}

        #Log {{ background:#0F172A; color:#D6E2F0; border:1px solid {STROKE};
            border-radius:8px; padding:6px; }}

        QSplitter::handle {{ background:transparent; height:8px; }}
        QScrollBar:vertical {{ background:transparent; width:10px; margin:2px; }}
        QScrollBar::handle:vertical {{ background:#C7CDD6; border-radius:5px; min-height:30px; }}
        QScrollBar::handle:vertical:hover {{ background:{ACCENT2}; }}
        QScrollBar:horizontal {{ background:transparent; height:10px; margin:2px; }}
        QScrollBar::handle:horizontal {{ background:#C7CDD6; border-radius:5px; min-width:30px; }}
        QScrollBar::add-line, QScrollBar::sub-line {{ width:0; height:0; }}
        """


def _podepnij_raportowanie_bledow(okno):
    """
    Kieruje nieobsluzone wyjatki do logu panelu.

    Qt polyka wyjatki rzucone w slotach: przycisk po prostu przestaje dzialac,
    bez zadnego sladu na ekranie (w zamrozonym .exe nie ma nawet konsoli).
    Tak wlasnie ukryl sie blad przycisku "Odswiez". Od teraz kazdy taki wyjatek
    leci do zakladki Log i do crash.log obok aplikacji.
    """
    poprzedni = sys.excepthook

    def hak(typ, wartosc, slad):
        tresc = "".join(traceback.format_exception(typ, wartosc, slad))
        try:
            okno.log.dopisz("\n!!! BLAD w interfejsie:\n" + tresc)
            okno.zakladki.setCurrentIndex(0)
            okno.pasek.setText("Blad w interfejsie — szczegoly w logu.")
        except Exception:
            pass
        try:
            with open(os.path.join(HERE, "crash.log"), "a", encoding="utf-8") as f:
                f.write(f"\n=== GUI {datetime.datetime.now():%Y-%m-%d %H:%M:%S} ===\n")
                f.write(tresc)
        except OSError:
            pass
        poprzedni(typ, wartosc, slad)

    sys.excepthook = hak


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    ico = zasob("app.ico")
    if os.path.exists(ico):
        app.setWindowIcon(QIcon(ico))
    okno = Okno()
    _podepnij_raportowanie_bledow(okno)
    # Panel otwiera sie zmaksymalizowany: formularze krokow sa dlugie, a log
    # i lista utworzonych plikow potrzebuja miejsca. Rozmiar z resize() zostaje
    # jako ten, do ktorego wraca okno po klikniecu "Przywroc w dol".
    okno.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
