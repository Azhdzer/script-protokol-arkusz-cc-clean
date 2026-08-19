# -*- coding: utf-8 -*-
"""
cc_widgets.py — elementy interfejsu panelu (app_gui.py).

Zawiera to, co da sie opisac w oderwaniu od samego okna:

  Pole            — jedno ustawienie z rejestru cc_config zamienione na widget
                    (pole tekstowe / liczba / przelacznik / lista plikow / kolor / tabela)
  ListaPlikow     — przewijana lista plikow z folderu z zaznaczaniem (TXT multimetru)
  WidokLogu       — zywy log z kolorowaniem bledow, ostrzezen i potwierdzen
  PanelWynikow    — co powstalo po uruchomieniu: nowe/zmienione pliki + zawartosc xlsx
  WierszKontrolny — jeden wiersz listy kontrolnej "co zaktualizowac przed startem"
"""

from __future__ import annotations

import datetime
import os
import re
import zipfile

from PySide6.QtCore import (Qt, Signal, QSize, QRectF, QEasingCurve, Property,
                            QPropertyAnimation)
from PySide6.QtGui import QColor, QTextCursor, QTextCharFormat, QFont, QPainter
from PySide6.QtWidgets import (
    QWidget, QLabel, QLineEdit, QCheckBox, QComboBox, QSpinBox, QDoubleSpinBox,
    QPushButton, QVBoxLayout, QHBoxLayout, QFrame, QPlainTextEdit, QFileDialog,
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QTreeWidget,
    QTreeWidgetItem, QColorDialog, QHeaderView, QAbstractItemView, QSizePolicy,
    QMessageBox, QGridLayout,
)

# ─── Paleta (Windows 11 / Fluent, jasny motyw) ───────────────────────────────
BG      = "#F3F3F3"
CARD    = "#FFFFFF"
STROKE  = "#E5E7EB"
ACCENT  = "#0067C0"
ACCENT2 = "#1975C5"
ACCENT3 = "#005BA1"
TEXT    = "#1A1A1A"
MUTED   = "#5E5E5E"
FIELD   = "#FBFBFB"
OK_C    = "#137333"
WARN_C  = "#B45309"
ERR_C   = "#B3261E"


def ludzki_rozmiar(bajty):
    """'18,3 MB' — rozmiar pliku w czytelnej postaci."""
    x = float(bajty)
    for jedn in ("B", "kB", "MB", "GB"):
        if x < 1024 or jedn == "GB":
            return f"{x:.1f} {jedn}".replace(".", ",") if jedn != "B" else f"{int(x)} B"
        x /= 1024
    return f"{x:.1f} GB"


def ile_temu(znacznik):
    """'dzis 09:27' / 'wczoraj' / '12 dni temu' — wiek pliku po ludzku."""
    dt = datetime.datetime.fromtimestamp(znacznik)
    dni = (datetime.date.today() - dt.date()).days
    if dni == 0:
        return f"dzis {dt:%H:%M}"
    if dni == 1:
        return f"wczoraj {dt:%H:%M}"
    return f"{dni} dni temu ({dt:%d.%m})"


def nazwy_arkuszy(sciezka, limit=40):
    """
    Nazwy zakladek pliku xlsx — czytane wprost z archiwum ZIP (xl/workbook.xml),
    bez openpyxl. Dzieki temu podglad zawartosci jest natychmiastowy nawet dla
    plikow po kilkanascie MB.
    """
    try:
        with zipfile.ZipFile(sciezka) as z:
            xml = z.read("xl/workbook.xml").decode("utf-8", errors="replace")
    except (OSError, KeyError, zipfile.BadZipFile):
        return []
    nazwy = re.findall(r'<sheet[^>]*\bname="([^"]*)"', xml)
    return [n.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
            for n in nazwy[:limit]]


# ═════════════════════════════════════════════════════════════════════════════
# Kontrolki odporne na kolko myszy
# ═════════════════════════════════════════════════════════════════════════════
# Domyslnie Qt pozwala zmieniac wartosc QComboBox / QSpinBox kolkiem myszy.
# Na dlugim, przewijanym formularzu to pulapka: przewijajac strone nad polem
# "Plik protokolu" mozna po cichu podmienic wybrany plik, a nad "Numer
# swiadectwa" — przestawic numeracje. Zmiana nie rzuca sie w oczy, bo wzrok
# sledzi przewijanie, a nie pole. Dlatego kolko jest tu ZAWSZE oddawane w gore
# (do obszaru przewijania), a wartosc zmienia sie wylacznie klikiem lub
# klawiatura.

class _BezKolka:
    """Domieszka: kolko myszy przewija strone, a nie zmienia wartosci pola."""

    def wheelEvent(self, zdarzenie):
        zdarzenie.ignore()


class ListaRozwijana(_BezKolka, QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.StrongFocus)


class PoleCalkowite(_BezKolka, QSpinBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.StrongFocus)


class PoleDziesietne(_BezKolka, QDoubleSpinBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.StrongFocus)


# ═════════════════════════════════════════════════════════════════════════════
# Przelacznik (switch) z suwakiem
# ═════════════════════════════════════════════════════════════════════════════
class Przelacznik(QCheckBox):
    """
    Przelacznik w stylu Windows 11: tor + okragly suwak przesuwajacy sie na bok.

    Rysowany recznie, bo arkusz stylow Qt nie potrafi narysowac RUCHOMEJ galki —
    da sie nim pomalowac tylko caly prostokat wskaznika. Efekt byl taki, ze
    wlaczony i wylaczony przelacznik roznily sie wylacznie kolorem wypelnienia,
    bez widocznego suwaka.

    Pozostaje zwyklym QCheckBox — isChecked()/setChecked()/toggled dzialaja jak
    dotad, wiec reszta panelu nie musi o nim nic wiedziec.
    """

    SZEROKOSC = 46      # dlugosc toru
    WYSOKOSC = 24       # grubosc toru
    LUZ = 3             # odstep galki od krawedzi toru
    ODSTEP_TEKSTU = 12

    def __init__(self, tekst="", parent=None):
        super().__init__(tekst, parent)
        self.setCursor(Qt.PointingHandCursor)
        self._pozycja = 1.0 if self.isChecked() else 0.0
        self._animacja = QPropertyAnimation(self, b"pozycja", self)
        self._animacja.setDuration(130)
        self._animacja.setEasingCurve(QEasingCurve.InOutCubic)
        self.toggled.connect(self._przesun)

    # ── animowana pozycja galki: 0.0 = wylaczony, 1.0 = wlaczony ─────────
    def _daj_pozycje(self):
        return self._pozycja

    def _ustaw_pozycje(self, wartosc):
        self._pozycja = wartosc
        self.update()

    pozycja = Property(float, _daj_pozycje, _ustaw_pozycje)

    def _przesun(self, wlaczony):
        docelowa = 1.0 if wlaczony else 0.0
        self._animacja.stop()
        # Przelacznika, ktorego nie widac, nie ma po co animowac — przy budowie
        # formularza powstaje ich kilkanascie naraz, jeszcze przed pokazaniem
        # okna. Ustawiamy polozenie od razu.
        if not self.isVisible():
            self._ustaw_pozycje(docelowa)
            return
        self._animacja.setStartValue(self._pozycja)
        self._animacja.setEndValue(docelowa)
        self._animacja.start()

    # ── geometria i reakcja na klikniecie ────────────────────────────────
    def sizeHint(self):
        szer_tekstu = self.fontMetrics().horizontalAdvance(self.text())
        wysokosc = max(self.WYSOKOSC + 6, self.fontMetrics().height() + 6)
        return QSize(self.SZEROKOSC + self.ODSTEP_TEKSTU + szer_tekstu + 4,
                     wysokosc)

    def minimumSizeHint(self):
        return self.sizeHint()

    def hitButton(self, punkt):
        """Klikalny jest caly widget — takze podpis obok przelacznika."""
        return self.rect().contains(punkt)

    def _dociagnij_galke(self):
        """
        Dosuwa galke do aktualnego stanu, gdy animacja nie biegnie.

        Konieczne, bo stan bywa ustawiany przy WYCISZONYCH sygnalach
        (Pole.ustaw() wola blockSignals, zeby wpisanie wartosci z pliku
        ustawien nie wygladalo jak edycja uzytkownika). Wtedy `toggled` nie
        leci, animacja nie startuje i galka zostalaby w poprzednim polozeniu —
        przelacznik pokazywalby stan odwrotny do prawdziwego.
        """
        if self._animacja.state() == QPropertyAnimation.Running:
            return
        docelowa = 1.0 if self.isChecked() else 0.0
        if self._pozycja != docelowa:
            self._pozycja = docelowa

    # ── rysowanie ────────────────────────────────────────────────────────
    def paintEvent(self, _zdarzenie):
        self._dociagnij_galke()
        malarz = QPainter(self)
        malarz.setRenderHint(QPainter.Antialiasing)

        gora = (self.height() - self.WYSOKOSC) / 2
        tor = QRectF(0, gora, self.SZEROKOSC, self.WYSOKOSC)

        malarz.setPen(Qt.NoPen)
        malarz.setBrush(self._kolor_toru())
        malarz.drawRoundedRect(tor, self.WYSOKOSC / 2, self.WYSOKOSC / 2)

        srednica = self.WYSOKOSC - 2 * self.LUZ
        przesuw = (self.SZEROKOSC - 2 * self.LUZ - srednica) * self._pozycja
        malarz.setBrush(QColor("#FFFFFF" if self.isEnabled() else "#F2F4F7"))
        malarz.drawEllipse(QRectF(self.LUZ + przesuw, gora + self.LUZ,
                                  srednica, srednica))

        if self.text():
            malarz.setPen(QColor(TEXT if self.isEnabled() else MUTED))
            malarz.setFont(self.font())
            lewa = self.SZEROKOSC + self.ODSTEP_TEKSTU
            malarz.drawText(QRectF(lewa, 0, self.width() - lewa, self.height()),
                            Qt.AlignVCenter | Qt.AlignLeft, self.text())

    def _kolor_toru(self):
        """Tor plynnie przechodzi z szarego w akcent razem z ruchem galki."""
        if not self.isEnabled():
            return QColor("#E1E5EA")
        wylaczony, wlaczony = QColor("#C3CBD6"), QColor(ACCENT)
        if self.underMouse():
            wylaczony, wlaczony = QColor("#AEB8C6"), QColor(ACCENT2)
        skladowe = [
            round(a + (b - a) * self._pozycja)
            for a, b in ((wylaczony.red(), wlaczony.red()),
                         (wylaczony.green(), wlaczony.green()),
                         (wylaczony.blue(), wlaczony.blue()))
        ]
        return QColor(*skladowe)

    def enterEvent(self, zdarzenie):
        super().enterEvent(zdarzenie)
        self.update()

    def leaveEvent(self, zdarzenie):
        super().leaveEvent(zdarzenie)
        self.update()


# ═════════════════════════════════════════════════════════════════════════════
# Pole ustawienia
# ═════════════════════════════════════════════════════════════════════════════
class Pole(QWidget):
    """
    Jedno ustawienie z rejestru cc_config jako gotowy wiersz formularza:
    etykieta + kontrolka + (opcjonalnie) opis pod spodem.

    API:  wartosc()  -> aktualna wartosc,  ustaw(v) -> wpisz wartosc.
    Sygnal `zmienione` leci przy kazdej edycji (panel zapisuje ustawienia).
    """

    zmienione = Signal()

    def __init__(self, ust, wartosc, folder_cb=None, parent=None):
        super().__init__(parent)
        self.ust = ust
        self._folder_cb = folder_cb or (lambda: os.getcwd())
        self._kontrolka = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(3)

        typ = ust.typ

        if typ == "flaga":
            # Przelacznik ma etykiete w sobie — nie dublujemy jej nad kontrolka.
            self._kontrolka = Przelacznik(ust.etykieta)
            self._kontrolka.setObjectName("Toggle")
            self._kontrolka.toggled.connect(lambda _: self.zmienione.emit())
            lay.addWidget(self._kontrolka)
        else:
            et = QLabel(ust.etykieta)
            et.setObjectName("EtykietaPola")
            et.setWordWrap(True)
            lay.addWidget(et)
            lay.addWidget(self._zbuduj_kontrolke(typ))

        if ust.opis:
            op = QLabel(ust.opis)
            op.setObjectName("OpisPola")
            op.setWordWrap(True)
            lay.addWidget(op)

        self.ustaw(wartosc)

    # ── budowa kontrolki wg typu ─────────────────────────────────────────
    def _zbuduj_kontrolke(self, typ):
        u = self.ust

        if typ in ("calk", "minuty"):
            w = PoleCalkowite()
            w.setRange(int(u.minimum if u.minimum is not None else 0),
                       int(u.maksimum if u.maksimum is not None else 1_000_000))
            w.setSuffix(u.przyrostek or (" min" if typ == "minuty" else ""))
            w.valueChanged.connect(lambda _: self.zmienione.emit())
            self._kontrolka = w
            return w

        if typ == "liczba":
            w = PoleDziesietne()
            w.setDecimals(2)
            w.setRange(float(u.minimum if u.minimum is not None else -1e6),
                       float(u.maksimum if u.maksimum is not None else 1e6))
            w.setSingleStep(float(u.krok_wart or 0.5))
            w.setSuffix(u.przyrostek)
            w.valueChanged.connect(lambda _: self.zmienione.emit())
            self._kontrolka = w
            return w

        if typ == "plik":
            ramka = QWidget()
            row = QHBoxLayout(ramka)
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(6)

            w = ListaRozwijana()
            w.setToolTip(u.opis or "Wybierz plik z folderu roboczego")
            w.currentTextChanged.connect(lambda _: self.zmienione.emit())

            btn_wybierz = QPushButton("Wybierz…")
            btn_wybierz.setObjectName("Ghost")
            btn_wybierz.setCursor(Qt.PointingHandCursor)
            btn_wybierz.setToolTip("Wskaz plik w oknie wyboru")
            btn_wybierz.clicked.connect(self._wybierz_plik)

            # Podpis slowny, a nie symbol strzalki: znaki w rodzaju U+27F3 nie
            # sa gwarantowane w czcionce interfejsu i potrafia wyjsc kwadracikiem.
            btn_odswiez = QPushButton("Odswiez")
            btn_odswiez.setObjectName("Ghost")
            btn_odswiez.setCursor(Qt.PointingHandCursor)
            btn_odswiez.setToolTip("Przeladuj liste plikow z folderu roboczego")
            btn_odswiez.clicked.connect(self.odswiez_pliki)

            row.addWidget(w, 1)
            row.addWidget(btn_wybierz, 0)
            row.addWidget(btn_odswiez, 0)

            self._kontrolka = w
            self.odswiez_pliki()
            return ramka

        if typ == "folder":
            ramka = QWidget()
            row = QHBoxLayout(ramka)
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(6)
            w = QLineEdit()
            w.textChanged.connect(lambda _: self.zmienione.emit())
            btn = QPushButton("Wybierz…")
            btn.setObjectName("Ghost")
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(self._wybierz_folder)
            row.addWidget(w, 1)
            row.addWidget(btn, 0)
            self._kontrolka = w
            return ramka

        if typ == "kolor":
            ramka = QWidget()
            row = QHBoxLayout(ramka)
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(6)
            w = QLineEdit()
            w.setMaximumWidth(110)
            w.textChanged.connect(self._odswiez_probke)
            w.textChanged.connect(lambda _: self.zmienione.emit())
            self._probka = QPushButton("")
            self._probka.setObjectName("Probka")
            self._probka.setFixedSize(34, 26)
            self._probka.setCursor(Qt.PointingHandCursor)
            self._probka.clicked.connect(self._wybierz_kolor)
            row.addWidget(w, 0)
            row.addWidget(self._probka, 0)
            row.addStretch(1)
            self._kontrolka = w
            return ramka

        if typ == "tabela":
            w = QTableWidget(0, len(u.kolumny))
            w.setHorizontalHeaderLabels(u.kolumny)
            w.verticalHeader().setVisible(False)
            w.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            w.setMinimumHeight(150)
            w.itemChanged.connect(lambda _: self.zmienione.emit())
            self._kontrolka = w
            return w

        w = QLineEdit()          # "tekst" i wszystko nieznane
        w.textChanged.connect(lambda _: self.zmienione.emit())
        self._kontrolka = w
        return w

    # ── obsluga ──────────────────────────────────────────────────────────
    def _wybierz_plik(self):
        """
        Wybor pliku przez okno systemowe.

        Plik musi lezec w folderze roboczym: skrypty skladaja sciezke jako
        os.path.join(FOLDER, nazwa), a z nazwy szablonu buduja tez nazwy kopii
        wyjsciowych. Plik z innego katalogu dalby albo bledna sciezke, albo
        dziwaczne nazwy swiadectw — dlatego taki wybor odrzucamy z wyjasnieniem.
        """
        folder = self._folder_cb()
        rozszerzenia = " ".join(f"*{r}" for r in (self.ust.wzorzec or (".xlsx",)))
        sciezka, _ = QFileDialog.getOpenFileName(
            self, self.ust.etykieta, folder,
            f"Obslugiwane pliki ({rozszerzenia});;Wszystkie pliki (*)")
        if not sciezka:
            return

        sciezka = os.path.normpath(sciezka)
        if os.path.normcase(os.path.dirname(sciezka)) != os.path.normcase(
                os.path.normpath(folder)):
            QMessageBox.information(
                self, "Plik poza folderem roboczym",
                "Wybrany plik lezy poza folderem roboczym:\n\n"
                f"{sciezka}\n\n"
                "Skrypty szukaja plikow wejsciowych w folderze roboczym i z ich "
                "nazw buduja nazwy kopii. Skopiuj plik do:\n\n"
                f"{folder}\n\ni wybierz go ponownie.")
            return

        nazwa = os.path.basename(sciezka)
        if self._kontrolka.findText(nazwa) < 0:
            self._kontrolka.insertItem(0, nazwa)
        self._kontrolka.setCurrentText(nazwa)

    def _wybierz_folder(self):
        start = self._kontrolka.text() or self._folder_cb()
        if not os.path.isabs(start):
            start = os.path.join(self._folder_cb(), start)
        d = QFileDialog.getExistingDirectory(self, self.ust.etykieta, start)
        if d:
            self._kontrolka.setText(d)

    def _wybierz_kolor(self):
        biezacy = QColor(self._kontrolka.text() or "#FFFFFF")
        k = QColorDialog.getColor(biezacy, self, self.ust.etykieta)
        if k.isValid():
            self._kontrolka.setText(k.name().upper())

    def _odswiez_probke(self, tekst):
        k = QColor(tekst)
        tlo = k.name() if k.isValid() else "#FFFFFF"
        self._probka.setStyleSheet(
            f"background:{tlo}; border:1px solid #B6C2D1; border-radius:5px;")

    def odswiez_pliki(self):
        """Dla typu 'plik' — przeladowuje liste z aktualnego folderu roboczego."""
        if self.ust.typ != "plik":
            return
        biezaca = self._kontrolka.currentText()
        folder = self._folder_cb()
        try:
            pliki = sorted(
                f for f in os.listdir(folder)
                if f.lower().endswith(tuple(self.ust.wzorzec or (".xlsx",)))
                and not f.startswith("~$"))
        except OSError:
            pliki = []
        self._kontrolka.blockSignals(True)
        self._kontrolka.clear()
        self._kontrolka.addItems(pliki)
        # Zapisany wybor moze wskazywac plik, ktorego juz nie ma w folderze —
        # zostawiamy go widocznego z ostrzezeniem, zamiast po cichu podmieniac.
        if biezaca:
            if biezaca in pliki:
                self._kontrolka.setCurrentText(biezaca)
            else:
                self._kontrolka.insertItem(0, biezaca)
                self._kontrolka.setCurrentIndex(0)
                self._kontrolka.setItemData(
                    0, "Tego pliku nie ma w folderze roboczym", Qt.ToolTipRole)
        elif self.ust.podpowiedz:
            # Pierwsze uruchomienie: zamiast pierwszego pliku alfabetycznie
            # wybierz ten, ktorego nazwa pasuje do podpowiedzi.
            frag = self.ust.podpowiedz.lower()
            trafienie = next((f for f in pliki if frag in f.lower()), None)
            if trafienie:
                self._kontrolka.setCurrentText(trafienie)
        self._kontrolka.blockSignals(False)

    def brak_pliku(self):
        """True, gdy wybrany plik nie istnieje (typ 'plik'/'folder')."""
        if self.ust.typ not in ("plik", "folder"):
            return False
        v = self.wartosc()
        if not v:
            return True
        p = v if os.path.isabs(v) else os.path.join(self._folder_cb(), v)
        return not os.path.exists(p)

    # ── wartosc ──────────────────────────────────────────────────────────
    def wartosc(self):
        typ, w = self.ust.typ, self._kontrolka
        if typ == "flaga":
            return w.isChecked()
        if typ in ("calk", "minuty"):
            return w.value()
        if typ == "liczba":
            return w.value()
        if typ == "plik":
            return w.currentText()
        if typ == "tabela":
            dane = []
            for r in range(w.rowCount()):
                wiersz = []
                for c in range(w.columnCount()):
                    it = w.item(r, c)
                    wiersz.append(it.text().strip() if it else "")
                if any(wiersz):
                    dane.append(wiersz)
            return dane
        return w.text()

    def ustaw(self, v):
        typ, w = self.ust.typ, self._kontrolka
        w.blockSignals(True)
        try:
            if typ == "flaga":
                w.setChecked(bool(v))
            elif typ in ("calk", "minuty"):
                w.setValue(int(v or 0))
            elif typ == "liczba":
                w.setValue(float(v or 0))
            elif typ == "plik":
                # Pusta wartosc = nic jeszcze nie wybrano; zostawiamy wybor
                # zrobiony przez odswiez_pliki() na podstawie podpowiedzi.
                if v:
                    if w.findText(v) < 0:
                        w.insertItem(0, v)
                    w.setCurrentText(v)
            elif typ == "tabela":
                wiersze = v or []
                w.setRowCount(len(wiersze) + 1)   # pusty wiersz na dopisanie
                for r in range(w.rowCount()):
                    for c in range(w.columnCount()):
                        zrodlo = wiersze[r] if r < len(wiersze) else []
                        tekst = str(zrodlo[c]) if c < len(zrodlo) else ""
                        w.setItem(r, c, QTableWidgetItem(tekst))
            else:
                w.setText("" if v is None else str(v))
                if typ == "kolor":
                    self._odswiez_probke(w.text())
        finally:
            w.blockSignals(False)


# ═════════════════════════════════════════════════════════════════════════════
# Lista plikow z zaznaczaniem (TXT multimetru, logi do analizy)
# ═════════════════════════════════════════════════════════════════════════════
class _ListaKlikalna(QListWidget):
    """
    Lista, w ktorej klikniecie w DOWOLNE miejsce wiersza przelacza pole wyboru.

    Domyslnie Qt reaguje tylko na trafienie w kwadracik ~13 px — przy wierszu
    wysokim na 50 px latwo w niego nie trafic. Tutaj celem jest caly wiersz.
    Zdarzenie jest konsumowane (bez wywolania klasy bazowej), wiec klikniecie
    w sam kwadracik nie przelacza stanu dwa razy.
    """

    def mousePressEvent(self, zdarzenie):
        element = self.itemAt(zdarzenie.position().toPoint())
        if element is None:
            super().mousePressEvent(zdarzenie)
            return
        self.setCurrentItem(element)
        element.setCheckState(
            Qt.Unchecked if element.checkState() == Qt.Checked else Qt.Checked)
        zdarzenie.accept()


class ListaPlikow(QWidget):
    """
    Przewijana lista plikow z folderu, kazdy z polem wyboru, data i rozmiarem.
    Sortowana od najnowszego — plik z dzisiejszego pomiaru jest zawsze na gorze.

    Zaznaczony wiersz jest wyroznany tlem i pogrubieniem, zeby po zaznaczeniu
    kilku plikow od razu bylo widac, ktore wejda do pomiaru.
    """

    zmienione = Signal()

    def __init__(self, wzorzec, folder_cb, parent=None):
        super().__init__(parent)
        self.wzorzec = wzorzec
        self._folder_cb = folder_cb

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        gora = QHBoxLayout()
        gora.setSpacing(8)
        self.podsumowanie = QLabel("—")
        self.podsumowanie.setObjectName("OpisPola")
        btn_odsw = QPushButton("Odswiez")
        btn_odsw.setObjectName("Ghost")
        btn_odsw.setCursor(Qt.PointingHandCursor)
        # UWAGA: `clicked` niesie argument bool (stan wcisniecia). Podpiete
        # wprost do odswiez() trafialby on w parametr `zachowaj` i wywracal
        # metode (TypeError: 'bool' object is not iterable) — przycisk milczaco
        # nie robil nic. Lambda odcina argument.
        btn_odsw.clicked.connect(lambda: self.odswiez())
        btn_wsz = QPushButton("Zaznacz wszystko")
        btn_wsz.setObjectName("Ghost")
        btn_wsz.setCursor(Qt.PointingHandCursor)
        btn_wsz.clicked.connect(self.zaznacz_wszystko)
        btn_nic = QPushButton("Odznacz wszystko")
        btn_nic.setObjectName("Ghost")
        btn_nic.setCursor(Qt.PointingHandCursor)
        btn_nic.clicked.connect(self.odznacz_wszystko)
        gora.addWidget(self.podsumowanie, 1)
        gora.addWidget(btn_wsz, 0)
        gora.addWidget(btn_nic, 0)
        gora.addWidget(btn_odsw, 0)
        lay.addLayout(gora)

        self.lista = _ListaKlikalna()
        self.lista.setObjectName("ListaPlikow")
        # 3 pelne wiersze po 50 px (reszta przez przewijanie). Wyzsza lista
        # spychala ustawienia zdjec pod dolna krawedz okna na laptopie.
        self.lista.setMinimumHeight(160)
        # Stan niesie pole wyboru, nie zaznaczenie wiersza. Szara belka
        # "biezacego elementu" tylko mylila — wygladala jak wybor pliku.
        self.lista.setSelectionMode(QAbstractItemView.NoSelection)
        self.lista.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.lista.itemChanged.connect(self._po_zmianie)
        lay.addWidget(self.lista, 1)

    def _ustaw_wszystkim(self, stan):
        self.lista.blockSignals(True)
        for i in range(self.lista.count()):
            self.lista.item(i).setCheckState(stan)
        self.lista.blockSignals(False)
        self._po_zmianie()

    def zaznacz_wszystko(self):
        self._ustaw_wszystkim(Qt.Checked)

    def odznacz_wszystko(self):
        self._ustaw_wszystkim(Qt.Unchecked)

    def _po_zmianie(self, *_):
        self._odswiez_wyglad()
        self._odswiez_podsumowanie()
        self.zmienione.emit()

    def _odswiez_wyglad(self):
        """
        Zaznaczony wiersz ma byc widoczny na pierwszy rzut oka: pogrubienie,
        tlo i granatowy tekst.

        Tlo ustawiamy przez role elementu, a nie przez QSS — arkusz stylow nie
        zna stanu "zaznaczony checkbox" (selektor ::item:checked nie istnieje).
        Dlatego regula ::item w QSS celowo NIE istnieje, bo nadpisalaby to.

        Sygnaly listy sa na czas zmiany wygladu wyciszone: setFont/setBackground
        tez emituja itemChanged, a to wpadaloby z powrotem w _po_zmianie i
        mnozylo sygnal `zmienione` przez liczbe wierszy.
        """
        self.lista.blockSignals(True)
        try:
            for i in range(self.lista.count()):
                element = self.lista.item(i)
                zaznaczony = element.checkState() == Qt.Checked
                czcionka = element.font()
                czcionka.setBold(zaznaczony)
                element.setFont(czcionka)
                element.setBackground(QColor("#D3E6FA") if zaznaczony
                                      else QColor(0, 0, 0, 0))
                element.setForeground(QColor(ACCENT3) if zaznaczony
                                      else QColor(TEXT))
        finally:
            self.lista.blockSignals(False)

    def _odswiez_podsumowanie(self):
        n = len(self.zaznaczone())
        wsz = self.lista.count()
        if wsz == 0:
            self.podsumowanie.setText("Brak pasujacych plikow w folderze roboczym.")
        elif n == 0:
            self.podsumowanie.setText(f"{wsz} plikow — nic nie zaznaczono.")
        elif n == 1:
            self.podsumowanie.setText(f"Zaznaczono 1 z {wsz} plikow.")
        else:
            self.podsumowanie.setText(
                f"Zaznaczono {n} z {wsz} plikow — zostana sklejone chronologicznie.")

    def odswiez(self, zachowaj=None):
        """
        Przeladowuje liste; `zachowaj` to nazwy do ponownego zaznaczenia
        (None = zachowaj biezacy wybor).
        """
        if zachowaj is not None and not isinstance(zachowaj, (list, tuple, set)):
            # Ktos podpial metode wprost pod sygnal niosacy argument (np. clicked).
            # Zamiast wywrocic sie po cichu, zachowujemy sie jak przy braku listy.
            zachowaj = None
        wybrane = set(zachowaj if zachowaj is not None else self.zaznaczone())
        folder = self._folder_cb()
        try:
            wpisy = [(e.name, e.stat()) for e in os.scandir(folder)
                     if e.is_file() and e.name.lower().endswith(self.wzorzec)
                     and not e.name.startswith("~$")]
        except OSError:
            wpisy = []
        wpisy.sort(key=lambda x: x[1].st_mtime, reverse=True)

        self.lista.blockSignals(True)
        self.lista.clear()
        for nazwa, st in wpisy:
            it = QListWidgetItem(f"{nazwa}\n      {ile_temu(st.st_mtime)}  ·  "
                                 f"{ludzki_rozmiar(st.st_size)}")
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if nazwa in wybrane else Qt.Unchecked)
            it.setData(Qt.UserRole, nazwa)
            it.setToolTip("Kliknij w dowolnym miejscu wiersza, aby zaznaczyc")
            it.setSizeHint(QSize(0, 50))
            self.lista.addItem(it)
        self.lista.blockSignals(False)
        self._odswiez_wyglad()
        self._odswiez_podsumowanie()

    def zaznaczone(self):
        return [self.lista.item(i).data(Qt.UserRole)
                for i in range(self.lista.count())
                if self.lista.item(i).checkState() == Qt.Checked]

    def ustaw_zaznaczone(self, nazwy):
        self.odswiez(zachowaj=nazwy)


# ═════════════════════════════════════════════════════════════════════════════
# Zywy log
# ═════════════════════════════════════════════════════════════════════════════
class WidokLogu(QPlainTextEdit):
    """
    Log podprocesu z kolorowaniem. Bledy i ostrzezenia sa widoczne od razu —
    nie trzeba ich szukac w scianie tekstu.
    """

    _WZORCE = (
        (re.compile(r"!!!\s*BLAD|Traceback|Error|Blad:|BLAD", re.I), "#FF8A80", True),
        (re.compile(r"UWAGA|WARN|Ostrzez", re.I),                   "#FFD180", False),
        (re.compile(r"\[OK\]|Gotowe|zakonczono|Zapisano",  re.I),   "#B9F6CA", False),
        (re.compile(r"^={5,}|^-{5,}|^\s*Etap\s|^\s*\d\.\s", re.M),  "#82B1FF", False),
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Log")
        self.setReadOnly(True)
        self.setMaximumBlockCount(20000)   # log potrafi urosnac; trzymamy pamiec w ryzach
        self.setFont(QFont("Cascadia Mono", 9))

    def dopisz(self, tekst):
        kursor = self.textCursor()
        kursor.movePosition(QTextCursor.End)
        for linia in tekst.splitlines(keepends=True):
            kursor.insertText(linia, self._format(linia))
        self.setTextCursor(kursor)
        self.ensureCursorVisible()

    def _format(self, linia):
        fmt = QTextCharFormat()
        for wzor, kolor, pogrub in self._WZORCE:
            if wzor.search(linia):
                fmt.setForeground(QColor(kolor))
                if pogrub:
                    fmt.setFontWeight(QFont.Bold)
                return fmt
        fmt.setForeground(QColor("#D6E2F0"))
        return fmt


# ═════════════════════════════════════════════════════════════════════════════
# Panel wynikow — co powstalo
# ═════════════════════════════════════════════════════════════════════════════
class PanelWynikow(QWidget):
    """
    Pokazuje pliki utworzone i zmienione przez ostatni przebieg.

    Dziala na roznicy migawek folderu (przed / po). Dla plikow xlsx dopisuje
    liste zakladek — od razu widac, czy kopia dostala wlasciwe punkty pomiarowe.
    Dwuklik otwiera plik w Excelu/Wordzie.
    """

    OBSERWOWANE = ("", "wyniki", "foto", "excel_do_analizy")

    def __init__(self, folder_cb, parent=None):
        super().__init__(parent)
        self._folder_cb = folder_cb
        self._migawka = {}

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        self.naglowek = QLabel("Uruchom krok — tutaj pojawi sie lista utworzonych plikow.")
        self.naglowek.setObjectName("OpisPola")
        self.naglowek.setWordWrap(True)
        lay.addWidget(self.naglowek)

        self.drzewo = QTreeWidget()
        self.drzewo.setObjectName("Wyniki")
        self.drzewo.setColumnCount(4)
        self.drzewo.setHeaderLabels(["Plik", "Stan", "Rozmiar", "Zawartosc"])
        self.drzewo.setRootIsDecorated(False)
        self.drzewo.setAlternatingRowColors(True)
        self.drzewo.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.drzewo.setMinimumHeight(120)
        hdr = self.drzewo.header()
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.Stretch)
        self.drzewo.itemDoubleClicked.connect(self._otworz)
        lay.addWidget(self.drzewo, 1)

    # ── migawki ──────────────────────────────────────────────────────────
    def _skanuj(self):
        baza = self._folder_cb()
        stan = {}
        for pod in self.OBSERWOWANE:
            katalog = os.path.join(baza, pod) if pod else baza
            try:
                for e in os.scandir(katalog):
                    if e.is_file() and not e.name.startswith("~$"):
                        st = e.stat()
                        stan[e.path] = (st.st_mtime, st.st_size)
            except OSError:
                continue
        return stan

    def zapamietaj_stan(self):
        """Migawka PRZED uruchomieniem — do niej porownamy stan po zakonczeniu."""
        self._migawka = self._skanuj()

    def pokaz_zmiany(self, etykieta_kroku):
        po = self._skanuj()
        nowe, zmienione = [], []
        for sciezka, (mtime, rozmiar) in po.items():
            przed = self._migawka.get(sciezka)
            if przed is None:
                nowe.append((sciezka, rozmiar))
            elif przed[0] != mtime or przed[1] != rozmiar:
                zmienione.append((sciezka, rozmiar))

        self.drzewo.clear()
        if not nowe and not zmienione:
            self.naglowek.setText(
                f"{etykieta_kroku}: zaden plik nie zostal utworzony ani zmieniony.")
            return

        self.naglowek.setText(
            f"{etykieta_kroku}: utworzono {len(nowe)}, zaktualizowano {len(zmienione)}.")
        for sciezka, rozmiar in sorted(nowe) + sorted(zmienione):
            stan = "nowy" if (sciezka, rozmiar) in nowe else "zmieniony"
            self._dodaj(sciezka, stan, rozmiar)
        self.drzewo.resizeColumnToContents(0)

    def _dodaj(self, sciezka, stan, rozmiar):
        baza = self._folder_cb()
        try:
            nazwa = os.path.relpath(sciezka, baza)
        except ValueError:
            nazwa = sciezka

        szczegol = ""
        if sciezka.lower().endswith((".xlsx", ".xlsm")):
            ark = nazwy_arkuszy(sciezka)
            if ark:
                szczegol = f"{len(ark)} zakladek: " + ", ".join(ark[:8])
                if len(ark) > 8:
                    szczegol += f", … (+{len(ark) - 8})"
        elif sciezka.lower().endswith(".docx"):
            szczegol = "swiadectwo Word"

        it = QTreeWidgetItem([nazwa, stan, ludzki_rozmiar(rozmiar), szczegol])
        it.setData(0, Qt.UserRole, sciezka)
        it.setForeground(1, QColor(OK_C if stan == "nowy" else ACCENT3))
        it.setToolTip(3, szczegol)
        it.setToolTip(0, f"{sciezka}\n(dwuklik otwiera plik)")
        self.drzewo.addTopLevelItem(it)

    def _otworz(self, item, _kol):
        sciezka = item.data(0, Qt.UserRole)
        if sciezka and os.path.exists(sciezka):
            try:
                os.startfile(sciezka)   # noqa: S606 — Windows, celowo
            except OSError:
                pass
class SiatkaKart(QWidget):
    """
    Karty ulozone w siatke, z liczba kolumn dobierana do szerokosci okna.

    Pozycje listy kontrolnej sa krotkie. Rozciagniete na cala szerokosc
    zostawialy pas pustego miejsca po prawej i zmuszaly do przewijania, zeby
    zobaczyc komplet. W siatce mieszcza sie po kilka w rzedzie.

    Liczba kolumn wynika z `min_szerokosc`: przy waskim oknie karty ustawiaja
    sie w jednej kolumnie, przy szerokim — w trzech czy czterech.
    """

    def __init__(self, min_szerokosc=250, odstep=10, parent=None):
        super().__init__(parent)
        self._min_szerokosc = min_szerokosc
        self._karty = []
        self._kolumny = 0
        self._siatka = QGridLayout(self)
        self._siatka.setContentsMargins(0, 0, 0, 0)
        self._siatka.setSpacing(odstep)

    def dodaj(self, karta):
        self._karty.append(karta)
        self._ustaw_uklad(wymus=True)

    def karty(self):
        return list(self._karty)

    def _ile_kolumn(self):
        odstep = self._siatka.spacing()
        dostepna = max(self.width(), self._min_szerokosc)
        mozliwe = (dostepna + odstep) // (self._min_szerokosc + odstep)
        return max(1, min(len(self._karty) or 1, int(mozliwe)))

    def _ustaw_uklad(self, wymus=False):
        kolumny = self._ile_kolumn()
        if kolumny == self._kolumny and not wymus:
            return
        self._kolumny = kolumny
        for karta in self._karty:
            self._siatka.removeWidget(karta)
        for i, karta in enumerate(self._karty):
            self._siatka.addWidget(karta, i // kolumny, i % kolumny)
        for kol in range(max(self._siatka.columnCount(), kolumny)):
            self._siatka.setColumnStretch(kol, 1 if kol < kolumny else 0)

    def resizeEvent(self, zdarzenie):
        super().resizeEvent(zdarzenie)
        self._ustaw_uklad()




# ═════════════════════════════════════════════════════════════════════════════
# Lista kontrolna "co zaktualizowac przed startem"
# ═════════════════════════════════════════════════════════════════════════════
class WierszKontrolny(QFrame):
    """
    Jeden wiersz listy kontrolnej: nazwa pliku/folderu, jego wiek i przyciski.

    Sedno: pokazac WPROST, ze np. 'Wzory.xls' ma 12 dni, zanim uruchomisz obieg
    i wygenerujesz swiadectwa ze starych wspolczynnikow.
    """

    def __init__(self, tytul, opis, sciezka_cb, prog_dni=None, katalog=False,
                 wzorzec=None, parent=None):
        super().__init__(parent)
        self.setObjectName("Kontrolny")
        self._sciezka_cb = sciezka_cb
        self._prog_dni = prog_dni
        self._katalog = katalog
        self._wzorzec = wzorzec

        # Uklad KARTY (a nie paska na cala szerokosc): pozycje listy kontrolnej
        # sa krotkie, wiec rozciagniete w wiersz zostawialy pas pustego miejsca.
        self.setMinimumWidth(250)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(4)

        gora = QHBoxLayout()
        gora.setSpacing(8)
        self.ikona = QLabel("•")
        self.ikona.setFixedWidth(16)
        self.ikona.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        tytul_l = QLabel(tytul)
        tytul_l.setObjectName("KontrolnyTytul")
        tytul_l.setWordWrap(True)
        gora.addWidget(self.ikona, 0, Qt.AlignTop)
        gora.addWidget(tytul_l, 1)
        lay.addLayout(gora)

        self.stan = QLabel("—")
        self.stan.setObjectName("OpisPola")
        self.stan.setWordWrap(True)
        lay.addWidget(self.stan)

        opis_l = QLabel(opis)
        opis_l.setObjectName("OpisPola")
        opis_l.setWordWrap(True)
        lay.addWidget(opis_l, 1)

        self.btn = QPushButton("Otworz")
        self.btn.setObjectName("Ghost")
        self.btn.setCursor(Qt.PointingHandCursor)
        self.btn.clicked.connect(self._otworz)
        dol = QHBoxLayout()
        dol.addStretch(1)
        dol.addWidget(self.btn, 0)
        lay.addLayout(dol)

    def _otworz(self):
        p = self._sciezka_cb()
        if p and os.path.exists(p):
            try:
                os.startfile(p)   # noqa: S606 — Windows, celowo
            except OSError:
                pass

    def odswiez(self):
        p = self._sciezka_cb()
        if not p or not os.path.exists(p):
            self._ustaw("✕", ERR_C, "BRAK — bez tego pliku krok nie ruszy.")
            return

        if self._katalog:
            try:
                pliki = [e for e in os.scandir(p) if e.is_file()
                         and (not self._wzorzec
                              or e.name.lower().endswith(self._wzorzec))]
            except OSError:
                pliki = []
            if not pliki:
                self._ustaw("!", WARN_C, "Folder jest pusty.")
                return
            najnowszy = max(e.stat().st_mtime for e in pliki)
            ile = "1 plik" if len(pliki) == 1 else f"{len(pliki)} plikow"
            self._ustaw("✓", OK_C, f"{ile} · najnowszy {ile_temu(najnowszy)}")
            return

        st = os.stat(p)
        wiek_dni = (datetime.date.today()
                    - datetime.date.fromtimestamp(st.st_mtime)).days
        tekst = f"{ile_temu(st.st_mtime)} · {ludzki_rozmiar(st.st_size)}"
        if self._prog_dni is not None and wiek_dni > self._prog_dni:
            self._ustaw("!", WARN_C, f"{tekst} — ZAKTUALIZUJ przed uruchomieniem.")
        else:
            self._ustaw("✓", OK_C, tekst)

    def _ustaw(self, znak, kolor, tekst):
        self.ikona.setText(znak)
        self.ikona.setStyleSheet(f"color:{kolor}; font-size:15px; font-weight:700;")
        self.stan.setText(tekst)
        self.stan.setStyleSheet(f"color:{kolor}; font-weight:600;")
