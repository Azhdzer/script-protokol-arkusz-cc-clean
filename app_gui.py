# -*- coding: utf-8 -*-
"""
app_gui.py — panel sterujacy (desktop, PySide6) dla generatora arkuszy/protokolow.

Uklad (styl Windows 11 / Fluent):
  • lewa kolumna — PLIKI WEJSCIOWE (folder + wybor plikow) → FUNKCJE → PARAMETRY,
  • srodek       — PODGLAD: makiety "Przed / Po wygenerowaniu" dla Protokolu,
                   arkusza Wyniki i swiadectwa Word (statyczne obrazki),
  • dol          — przyciski akcji (Obserwacja / Protokol / Word) + zywy log.

Uruchomienie (zrodlo):  .venv\\Scripts\\python.exe app_gui.py
Zamrozony .exe uruchamia workery jako podprocesy (patrz app_entry.py).
"""

import os
import sys

from PySide6.QtCore import Qt, QProcess, QProcessEnvironment, QSettings
from PySide6.QtGui import QFont, QTextCursor, QPixmap, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QCheckBox, QPushButton,
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QFileDialog,
    QDoubleSpinBox, QSpinBox, QFrame, QPlainTextEdit, QSizePolicy, QScrollArea,
    QButtonGroup,
)

# ─── Sciezki (dziala z zrodla i po zamrozeniu PyInstaller) ─────────────────────
if getattr(sys, "frozen", False):
    HERE = os.path.dirname(sys.executable)
    _BUNDLE = getattr(sys, "_MEIPASS", HERE)
else:
    HERE = os.path.dirname(os.path.abspath(__file__))
    _BUNDLE = HERE


def resource_path(rel):
    """Zwraca sciezke do zasobu (assets/...) — z bundla przy zamrozeniu."""
    return os.path.join(_BUNDLE, rel)


# ─── Paleta (Windows 11 / Fluent, jasny motyw) ─────────────────────────────────
BG      = "#F3F3F3"   # tlo okna (Win11 mica-like)
CARD    = "#FFFFFF"
STROKE  = "#E5E7EB"   # subtelna ramka karty
ACCENT  = "#0067C0"   # systemowy niebieski Win11
ACCENT2 = "#1975C5"
ACCENT3 = "#005BA1"
TEXT    = "#1A1A1A"
MUTED   = "#5E5E5E"
FIELD   = "#FBFBFB"

# ─── Funkcje (przelaczniki) ───────────────────────────────────────────────────
# key, label, env (zmienna srodowiskowa dla skryptu), domyslnie, opis
FUNKCJE = [
    dict(key="excel",   label="Generuj arkusze Excel",        env="GEN_EXCEL",   default=True,
         tip="Kopie arkuszy obliczeniowych + wypelnienie zakladek."),
    dict(key="word",    label="Generuj swiadectwa Word",      env="GEN_WORD",    default=True,
         tip="Dokumenty Word (swiadectwa wzorcowania)."),
    dict(key="puste",   label="Usuwaj puste bloki Strona 3",  env="GEN_PUSTE",   default=True,
         tip="Kopia dostaje tylko aktywne bloki Strona 3."),
    dict(key="autorec", label="Sprzataj AutoRecover",         env="GEN_AUTOREC", default=True,
         tip="Czysci stare pliki autoodzyskiwania (zapobiega crashom Excela)."),
    dict(key="filtr",   label="Filtr nastawa/odczyt >prog",   env="OBS_FILTR",   default=True,
         tip="Odrzuca punkt gdy odczyt komory rozni sie od nastawy > progu."),
]

# Nazwy plikow wejsciowych domyslnie wybierane (dopasowanie po fragmencie nazwy).
_DOMYSLNY_PROTOKOL = "protokół CC-04"
_DOMYSLNY_SZABLON  = "Wzór ark. obl"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generator arkuszy i protokolow")
        self.resize(1180, 760)
        self.setMinimumSize(980, 640)
        self.checks = {}
        self.proc = None
        self.settings = QSettings("LabTH", "GeneratorArkuszy")
        self.folder = self.settings.value("folder", HERE)
        if not os.path.isdir(self.folder):
            self.folder = HERE

        self.doc_key = "protokol"     # protokol | wyniki | word
        self.state_pelny = False      # False = "Przed", True = "Po"

        root = QWidget(); self.setCentralWidget(root)
        outer = QVBoxLayout(root); outer.setContentsMargins(18, 16, 18, 16); outer.setSpacing(12)

        outer.addWidget(self._header())

        body = QHBoxLayout(); body.setSpacing(14)
        body.addWidget(self._sidebar(), 0)
        body.addWidget(self._center(), 1)
        outer.addLayout(body, 1)

        outer.addWidget(self._footer())

        self.setStyleSheet(self._qss())
        self._skanuj_folder()
        self._refresh_preview()

    # ── Naglowek ─────────────────────────────────────────────────────────
    def _header(self):
        w = QFrame(); w.setObjectName("Header")
        lay = QVBoxLayout(w); lay.setContentsMargins(22, 16, 22, 16); lay.setSpacing(3)
        t = QLabel("Generator arkuszy i protokolow"); t.setObjectName("H1")
        s = QLabel("Wybierz pliki i funkcje po lewej. Podglad po srodku pokazuje, "
                   "jak dokument wyglada przed i po wygenerowaniu.")
        s.setObjectName("H2"); s.setWordWrap(True)
        lay.addWidget(t); lay.addWidget(s)
        return w

    # ── Lewy panel (przewijalny) ─────────────────────────────────────────
    def _sidebar(self):
        scroll = QScrollArea(); scroll.setObjectName("SideScroll")
        scroll.setWidgetResizable(True); scroll.setFixedWidth(346)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        holder = QWidget(); scroll.setWidget(holder)
        col = QVBoxLayout(holder); col.setContentsMargins(0, 0, 8, 0); col.setSpacing(14)

        # — PLIKI WEJSCIOWE —
        c1 = self._card(); l1 = self._card_layout(c1)
        l1.addWidget(self._cap("PLIKI WEJSCIOWE"))

        frow = QHBoxLayout(); frow.setSpacing(8)
        self.lbl_folder = QLabel(); self.lbl_folder.setObjectName("Path")
        self.lbl_folder.setToolTip("Folder z plikami wejsciowymi")
        self.lbl_folder.setTextInteractionFlags(Qt.TextSelectableByMouse)
        btn_folder = QPushButton("Zmien…"); btn_folder.setObjectName("Ghost")
        btn_folder.setCursor(Qt.PointingHandCursor); btn_folder.clicked.connect(self._wybierz_folder)
        frow.addWidget(self.lbl_folder, 1); frow.addWidget(btn_folder, 0)
        l1.addWidget(QLabel("Folder roboczy")); l1.addLayout(frow)

        l1.addWidget(self._field_label("Plik protokolu (CC-04)"))
        self.cb_protokol = QComboBox(); self.cb_protokol.setToolTip("Arkusz protokolu do przetworzenia")
        l1.addWidget(self.cb_protokol)

        l1.addWidget(self._field_label("Plik szablonu (Wzor ark. obl.)"))
        self.cb_szablon = QComboBox(); self.cb_szablon.setToolTip("Szablon arkusza obliczeniowego")
        l1.addWidget(self.cb_szablon)
        col.addWidget(c1)

        # — FUNKCJE —
        c2 = self._card(); l2 = self._card_layout(c2)
        l2.addWidget(self._cap("FUNKCJE"))
        for spec in FUNKCJE:
            cb = QCheckBox(spec["label"]); cb.setObjectName("Toggle")
            cb.setChecked(self.settings.value(f"f_{spec['key']}", spec["default"], type=bool))
            cb.setToolTip(spec["tip"]); cb.setCursor(Qt.PointingHandCursor)
            cb.stateChanged.connect(self._refresh_preview)
            self.checks[spec["key"]] = cb
            l2.addWidget(cb)
        col.addWidget(c2)

        # — PARAMETRY —
        c3 = self._card(); l3 = self._card_layout(c3)
        l3.addWidget(self._cap("PARAMETRY"))
        grid = QGridLayout(); grid.setVerticalSpacing(10); grid.setHorizontalSpacing(10)
        grid.addWidget(QLabel("Prog nastawa/odczyt [%]"), 0, 0)
        self.sp_prog = QDoubleSpinBox(); self.sp_prog.setRange(0, 100)
        self.sp_prog.setValue(self.settings.value("prog", 10.0, type=float)); self.sp_prog.setSingleStep(1)
        self.sp_prog.valueChanged.connect(self._refresh_preview)
        grid.addWidget(self.sp_prog, 0, 1)
        grid.addWidget(QLabel("Tolerancja czasu [min]"), 1, 0)
        self.sp_tol = QSpinBox(); self.sp_tol.setRange(1, 240)
        self.sp_tol.setValue(self.settings.value("tol", 30, type=int))
        grid.addWidget(self.sp_tol, 1, 1)
        l3.addLayout(grid)
        col.addWidget(c3)

        col.addStretch(1)
        return scroll

    # ── Srodek: podglad (makiety przed/po) ───────────────────────────────
    def _center(self):
        card = self._card(); lay = self._card_layout(card, margin=(18, 16, 18, 18))
        lay.addWidget(self._cap("PODGLAD"))

        # segmented: wybor dokumentu
        seg = QHBoxLayout(); seg.setSpacing(0)
        self.seg_group = QButtonGroup(self); self.seg_group.setExclusive(True)
        for i, (key, label) in enumerate([("protokol", "Protokol"), ("wyniki", "Arkusz Wyniki"),
                                          ("word", "Swiadectwo Word")]):
            b = QPushButton(label); b.setObjectName("Seg"); b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor); b.setChecked(i == 0)
            b.setProperty("seg_pos", "left" if i == 0 else ("right" if i == 2 else "mid"))
            b.clicked.connect(lambda _c, k=key: self._set_doc(k))
            self.seg_group.addButton(b); seg.addWidget(b)
        seg.addStretch(1)

        # przelacznik Przed / Po
        self.sw_stan = QPushButton("Pokaz: PO wygenerowaniu"); self.sw_stan.setObjectName("StateSw")
        self.sw_stan.setCheckable(True); self.sw_stan.setCursor(Qt.PointingHandCursor)
        self.sw_stan.toggled.connect(self._set_stan)
        seg.addWidget(self.sw_stan)
        lay.addLayout(seg)

        self.opis = QLabel(); self.opis.setObjectName("Opis"); self.opis.setWordWrap(True)
        lay.addWidget(self.opis)

        # obszar obrazka
        self.preview = QLabel(); self.preview.setObjectName("Preview")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview.setMinimumHeight(320)
        lay.addWidget(self.preview, 1)
        return card

    # ── Dol: przyciski + log ─────────────────────────────────────────────
    def _footer(self):
        card = self._card(); lay = self._card_layout(card, margin=(16, 12, 16, 12))
        row = QHBoxLayout(); row.setSpacing(12)
        self.btn_obs  = QPushButton("Obserwacja")
        self.btn_prot = QPushButton("Protokol")
        self.btn_word = QPushButton("Swiadectwo Word")
        for b, obj in ((self.btn_obs, "Secondary"), (self.btn_prot, "Primary"), (self.btn_word, "Primary")):
            b.setObjectName(obj); b.setMinimumHeight(44); b.setCursor(Qt.PointingHandCursor)
        self.btn_obs.clicked.connect(lambda: self._run("obserwacja"))
        self.btn_prot.clicked.connect(lambda: self._run("protokol"))
        self.btn_word.clicked.connect(lambda: self._run("word"))
        row.addWidget(self.btn_obs); row.addStretch(1)
        row.addWidget(self.btn_prot); row.addWidget(self.btn_word)
        lay.addLayout(row)

        head = QHBoxLayout()
        self.status = QLabel("Gotowy."); self.status.setObjectName("Status")
        self.btn_log = QPushButton("Pokaz log ▾"); self.btn_log.setObjectName("Ghost")
        self.btn_log.setCheckable(True); self.btn_log.setCursor(Qt.PointingHandCursor)
        self.btn_log.toggled.connect(self._toggle_log)
        head.addWidget(self.status); head.addStretch(1); head.addWidget(self.btn_log)
        lay.addLayout(head)

        self.log = QPlainTextEdit(); self.log.setObjectName("Log"); self.log.setReadOnly(True)
        self.log.setFixedHeight(150); self.log.setVisible(False)
        lay.addWidget(self.log)
        return card

    # ── Helpery UI ────────────────────────────────────────────────────────
    def _card(self):
        c = QFrame(); c.setObjectName("Card"); return c

    def _card_layout(self, card, margin=(18, 16, 18, 16)):
        lay = QVBoxLayout(card); lay.setContentsMargins(*margin); lay.setSpacing(9); return lay

    def _cap(self, text):
        w = QLabel(text); w.setObjectName("Cap"); return w

    def _field_label(self, text):
        w = QLabel(text); w.setContentsMargins(0, 6, 0, 0); return w

    # ── Pliki wejsciowe ─────────────────────────────────────────────────────
    def _wybierz_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Wybierz folder roboczy", self.folder)
        if d:
            self.folder = d; self.settings.setValue("folder", d); self._skanuj_folder()

    def _skanuj_folder(self):
        self.lbl_folder.setText(self.folder)
        try:
            xlsx = sorted(f for f in os.listdir(self.folder)
                          if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"))
        except OSError:
            xlsx = []
        for cb, frag, key in ((self.cb_protokol, _DOMYSLNY_PROTOKOL, "protokol_plik"),
                              (self.cb_szablon, _DOMYSLNY_SZABLON, "szablon_plik")):
            cb.blockSignals(True); cb.clear(); cb.addItems(xlsx)
            zapis = self.settings.value(key, "")
            wybor = zapis if zapis in xlsx else next((f for f in xlsx if frag.lower() in f.lower()), None)
            if wybor:
                cb.setCurrentText(wybor)
            cb.blockSignals(False)
        if not xlsx:
            self.status.setText("Brak plikow .xlsx w wybranym folderze.")

    # ── Podglad ─────────────────────────────────────────────────────────────
    def _set_doc(self, key):
        self.doc_key = key; self._refresh_preview()

    def _set_stan(self, on):
        self.state_pelny = on
        self.sw_stan.setText("Pokaz: PRZED (pusty)" if on else "Pokaz: PO wygenerowaniu")
        self._refresh_preview()

    def _stan(self, key):
        return self.checks[key].isChecked()

    def _refresh_preview(self):
        stan = "pelny" if self.state_pelny else "pusto"
        path = resource_path(os.path.join("assets", f"{self.doc_key}_{stan}.png"))
        pix = QPixmap(path)
        if pix.isNull():
            self.preview.setText("(brak makiety podgladu)")
        else:
            w = max(360, self.preview.width() - 24)
            self.preview.setPixmap(pix.scaledToWidth(min(w, pix.width()),
                                                     Qt.SmoothTransformation))
        # opis kontekstowy
        prog = self.sp_prog.value()
        if self.state_pelny:
            if self._stan("filtr"):
                txt = f"Po wygenerowaniu: filtr aktywny (prog {prog:g}%) — punkty poza nastawa sa odrzucane (czerwone)."
            else:
                txt = "Po wygenerowaniu: filtr wylaczony — wszystkie punkty sa brane."
        else:
            txt = "Przed uruchomieniem: pusty szablon. Wlacz przelacznik po prawej, aby zobaczyc wynik."
        self.opis.setText(txt)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if hasattr(self, "preview"):
            self._refresh_preview()

    # ── Uruchamianie (QProcess — bez zawieszania) ────────────────────────
    def _run(self, tryb):
        if self.proc and self.proc.state() != QProcess.NotRunning:
            self.status.setText("Trwa juz inne zadanie — poczekaj.")
            return
        worker = "obserwacje" if tryb == "obserwacja" else "arkusze"
        skrypt = f"generuj_{worker}.py"
        frozen = getattr(sys, "frozen", False)
        if not frozen:
            sciezka = os.path.join(HERE, skrypt)
            if not os.path.exists(sciezka):
                self.status.setText(f"Brak pliku: {skrypt}")
                return

        env = QProcessEnvironment.systemEnvironment()
        for spec in FUNKCJE:
            env.insert(spec["env"], "1" if self._stan(spec["key"]) else "0")
        if tryb == "word":
            env.insert("GEN_WORD", "1"); env.insert("GEN_EXCEL", "0")
        env.insert("OBS_PROG", str(self.sp_prog.value()))
        env.insert("OBS_TOL", str(self.sp_tol.value()))
        env.insert("PYTHONIOENCODING", "utf-8")
        # Wybor plikow wejsciowych → nadpisuje domyslne w generuj_arkusze.py
        env.insert("CC_FOLDER", self.folder)
        if self.cb_protokol.currentText():
            env.insert("CC_PROTOKOL", self.cb_protokol.currentText())
        if self.cb_szablon.currentText():
            env.insert("CC_SZABLON", self.cb_szablon.currentText())

        self._zapisz_ustawienia()

        self.proc = QProcess(self)
        self.proc.setProcessEnvironment(env)
        self.proc.setWorkingDirectory(self.folder)
        self.proc.setProcessChannelMode(QProcess.MergedChannels)
        self.proc.readyReadStandardOutput.connect(self._on_output)
        self.proc.finished.connect(lambda code, _st: self._on_finished(tryb, code))

        self.log.clear(); self.btn_log.setChecked(True)
        self._busy(True)
        self.status.setText(f"'{tryb}' uruchomiono…")

        if frozen:
            env.insert("CC_WORKER", worker)
            self.proc.setProcessEnvironment(env)
            self.log.appendPlainText(f"$ {os.path.basename(sys.executable)} [{worker}]\n")
            self.proc.start(sys.executable, [])
        else:
            self.log.appendPlainText(f"$ python {skrypt}\n")
            self.proc.start(sys.executable, [sciezka])

    def _on_output(self):
        data = bytes(self.proc.readAllStandardOutput()).decode("utf-8", errors="replace")
        self.log.moveCursor(QTextCursor.End)
        self.log.insertPlainText(data)
        self.log.moveCursor(QTextCursor.End)

    def _on_finished(self, tryb, code):
        self._busy(False)
        ok = (code == 0)
        self.status.setText(f"'{tryb}': {'zakonczono ✓' if ok else f'blad (kod {code})'}")
        self.log.appendPlainText(f"\n[proces zakonczony, kod {code}]")

    def _busy(self, b):
        for x in (self.btn_obs, self.btn_prot, self.btn_word):
            x.setEnabled(not b)

    def _toggle_log(self, on):
        self.log.setVisible(on)
        self.btn_log.setText("Ukryj log ▴" if on else "Pokaz log ▾")

    def _zapisz_ustawienia(self):
        s = self.settings
        s.setValue("folder", self.folder)
        s.setValue("protokol_plik", self.cb_protokol.currentText())
        s.setValue("szablon_plik", self.cb_szablon.currentText())
        s.setValue("prog", self.sp_prog.value()); s.setValue("tol", self.sp_tol.value())
        for spec in FUNKCJE:
            s.setValue(f"f_{spec['key']}", self._stan(spec["key"]))

    def closeEvent(self, e):
        self._zapisz_ustawienia(); super().closeEvent(e)

    # ── Styl (Windows 11 / Fluent) ───────────────────────────────────────
    def _qss(self):
        return f"""
        QWidget {{ background:{BG}; color:{TEXT}; font-family:'Segoe UI Variable','Segoe UI',sans-serif; font-size:13px; }}
        #Header {{ background:{CARD}; border:1px solid {STROKE}; border-radius:10px; }}
        #H1 {{ font-size:19px; font-weight:700; color:{TEXT}; }}
        #H2 {{ font-size:12px; color:{MUTED}; }}
        #Card {{ background:{CARD}; border:1px solid {STROKE}; border-radius:10px; }}
        #SideScroll {{ border:none; background:transparent; }}
        #Cap {{ font-size:11px; font-weight:700; color:{ACCENT3}; letter-spacing:1px; }}
        #Path {{ color:{MUTED}; font-size:11px; }}
        QLabel {{ color:{TEXT}; }}

        QCheckBox#Toggle {{ spacing:10px; padding:5px 2px; }}
        QCheckBox#Toggle::indicator {{ width:38px; height:20px; border-radius:10px;
            background:#CBD5E1; border:1px solid #B6C2D1; }}
        QCheckBox#Toggle::indicator:checked {{ background:{ACCENT}; border:1px solid {ACCENT}; }}
        QCheckBox#Toggle::indicator:hover {{ border:1px solid {ACCENT2}; }}

        QComboBox, QDoubleSpinBox, QSpinBox {{ background:{FIELD}; border:1px solid #D6DCE4;
            border-radius:7px; padding:6px 8px; min-height:20px; }}
        QComboBox:hover, QDoubleSpinBox:hover, QSpinBox:hover {{ border:1px solid {ACCENT2}; }}
        QComboBox::drop-down {{ border:none; width:22px; }}
        QComboBox QAbstractItemView {{ background:{CARD}; border:1px solid {STROKE};
            selection-background-color:{ACCENT}; selection-color:white; outline:0; }}

        QPushButton#Primary {{ background:{ACCENT}; color:white; border:none; border-radius:8px;
            font-size:14px; font-weight:700; padding:8px 26px; }}
        QPushButton#Primary:hover {{ background:{ACCENT2}; }}
        QPushButton#Primary:pressed {{ background:{ACCENT3}; }}
        QPushButton#Primary:disabled {{ background:#9DBBDC; }}
        QPushButton#Secondary {{ background:{CARD}; color:{ACCENT}; border:1px solid {ACCENT};
            border-radius:8px; font-size:14px; font-weight:700; padding:8px 22px; }}
        QPushButton#Secondary:hover {{ background:#EAF3FB; }}
        QPushButton#Secondary:disabled {{ color:#9DBBDC; border-color:#9DBBDC; }}
        QPushButton#Ghost {{ background:transparent; border:1px solid {STROKE}; border-radius:7px;
            padding:5px 12px; color:{MUTED}; font-weight:600; }}
        QPushButton#Ghost:hover {{ background:#F3F6FC; border-color:{ACCENT2}; color:{ACCENT}; }}

        QPushButton#Seg {{ background:{FIELD}; border:1px solid #D6DCE4; padding:7px 16px;
            color:{MUTED}; font-weight:600; }}
        QPushButton#Seg[seg_pos="left"] {{ border-top-left-radius:8px; border-bottom-left-radius:8px; }}
        QPushButton#Seg[seg_pos="right"] {{ border-top-right-radius:8px; border-bottom-right-radius:8px; }}
        QPushButton#Seg:checked {{ background:{ACCENT}; color:white; border:1px solid {ACCENT}; }}
        QPushButton#Seg:hover:!checked {{ background:#EAF3FB; color:{ACCENT}; }}

        QPushButton#StateSw {{ background:#EAF3FB; border:1px solid {ACCENT}; border-radius:8px;
            padding:7px 14px; color:{ACCENT}; font-weight:700; }}
        QPushButton#StateSw:checked {{ background:{ACCENT}; color:white; }}

        #Opis {{ color:{MUTED}; font-size:12px; padding:2px 0; }}
        #Preview {{ background:#FAFBFC; border:1px dashed #CBD5E1; border-radius:10px; padding:12px; }}
        #Status {{ color:{MUTED}; }}
        #Log {{ background:#0F172A; color:#D1FAE5; border-radius:8px;
            font-family:'Cascadia Mono','Consolas',monospace; font-size:12px; }}
        QScrollBar:vertical {{ background:transparent; width:10px; margin:2px; }}
        QScrollBar::handle:vertical {{ background:#C7CDD6; border-radius:5px; min-height:30px; }}
        QScrollBar::handle:vertical:hover {{ background:{ACCENT2}; }}
        QScrollBar::add-line, QScrollBar::sub-line {{ height:0; }}
        """


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    ico = resource_path("app.ico")
    if os.path.exists(ico):
        app.setWindowIcon(QIcon(ico))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
