# -*- coding: utf-8 -*-
"""
testy/_podglad.py — narzedzie pomocnicze (nie test).

Renderuje okno panelu do pliku PNG, zeby dalo sie obejrzec uklad bez klikania
po aplikacji. Przydatne przy zmianach w wygladzie.

Uruchomienie:
    .venv\\Scripts\\python.exe testy\\_podglad.py            # krok 2 (Obserwacja)
    .venv\\Scripts\\python.exe testy\\_podglad.py 4          # strona Zaawansowane

Wynik: testy/_piaskownica/podglad_<nr>.png
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QApplication

import app_gui
import cc_config as C

NR_STRONY = int(sys.argv[1]) if len(sys.argv) > 1 else 2
WYJSCIE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_piaskownica")


def main():
    os.makedirs(WYJSCIE, exist_ok=True)
    # Podglad nie moze zostawic po sobie cc_ustawienia.json w folderze projektu
    # — panel zapisuje ustawienia z opoznieniem, wiec przekierowujemy zapis
    # do piaskownicy.
    C.PLIK_USTAWIEN = os.path.join(WYJSCIE, "cc_ustawienia_podgladu.json")
    app = QApplication(sys.argv)
    okno = app_gui.Okno()
    okno.resize(1400, 940)
    okno.show()

    def zrzut():
        okno.grupa.buttons()[NR_STRONY].setChecked(True)
        okno._przelacz(NR_STRONY)
        app.processEvents()
        # Na stronie Obserwacji zaznacz dwa najnowsze pliki — widac wtedy,
        # jak wyglada wiersz zaznaczony.
        if NR_STRONY == 2:
            for i in range(min(2, okno.lista_txt.lista.count())):
                okno.lista_txt.lista.item(i).setCheckState(Qt.Checked)
            okno.lista_txt._po_zmianie()
            app.processEvents()
        sciezka = os.path.join(WYJSCIE, f"podglad_{NR_STRONY}.png")
        okno.grab().save(sciezka)
        print(f"zapisano: {sciezka}")
        app.quit()

    QTimer.singleShot(800, zrzut)
    app.exec()


if __name__ == "__main__":
    main()
