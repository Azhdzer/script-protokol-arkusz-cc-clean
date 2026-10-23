# -*- coding: utf-8 -*-
"""
cc_kopia.py — kopia zapasowa pliku PRZED jego nadpisaniem.

Krok 2 tworzy protokol od nowa z szablonu, a krok 3 — kopie arkuszy. Gdy plik
o tej nazwie juz byl, przepadaly reczne poprawki (zlecenie 230: recznie
wypelnione przetworniki na Stronach 2-3 zniknely po ponownym uruchomieniu
obiegu). Teraz poprzednia wersja laduje w podfolderze 'kopie_zapasowe'
folderu roboczego, z data i godzina w nazwie:

    kopie_zapasowe/230_LA_TH_2026 - protokół CC (2026-10-23 08.31.27).xlsx

Trzymamy ZACHOWAJ najnowszych kopii kazdego pliku — starsze sa usuwane.
"""

import datetime
import glob
import os
import shutil

FOLDER_KOPII = "kopie_zapasowe"
ZACHOWAJ = 20


def kopia_przed_nadpisaniem(sciezka, zachowaj=ZACHOWAJ):
    """
    Gdy `sciezka` istnieje — kopiuje ja do <folder pliku>/kopie_zapasowe.
    Zwraca sciezke kopii albo None. Bledy sa zglaszane, ale nie przerywaja kroku.
    """
    if not sciezka or not os.path.isfile(sciezka):
        return None
    folder = os.path.join(os.path.dirname(os.path.abspath(sciezka)), FOLDER_KOPII)
    rdzen, rozsz = os.path.splitext(os.path.basename(sciezka))
    znacznik = datetime.datetime.fromtimestamp(os.path.getmtime(sciezka)).strftime(
        "%Y-%m-%d %H.%M.%S")
    cel = os.path.join(folder, f"{rdzen} ({znacznik}){rozsz}")
    try:
        os.makedirs(folder, exist_ok=True)
        if not os.path.exists(cel):
            shutil.copy2(sciezka, cel)
        print(f"  [Kopia zapasowa] {os.path.basename(sciezka)} -> {FOLDER_KOPII}\\{os.path.basename(cel)}")
    except OSError as e:
        print(f"  [Kopia zapasowa] NIE udalo sie zachowac '{os.path.basename(sciezka)}': {e}")
        return None
    _przytnij(folder, rdzen, rozsz, zachowaj)
    return cel


def _przytnij(folder, rdzen, rozsz, zachowaj):
    wzorzec = os.path.join(folder, glob.escape(rdzen) + " (*)" + glob.escape(rozsz))
    kopie = sorted(glob.glob(wzorzec), key=os.path.getmtime, reverse=True)
    for stara in kopie[zachowaj:]:
        try:
            os.remove(stara)
        except OSError:
            pass
