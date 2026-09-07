# -*- coding: utf-8 -*-
r"""
zrob_skrot.py — tworzy skrot uruchamiajacy panel BEZ okna konsoli.

Skrot celuje w `.venv\Scripts\pythonw.exe` (wersja Pythona bez konsoli)
z argumentem `app_gui.py`. Dwuklik otwiera samo okno panelu.

Uruchomienie:
    .venv\Scripts\python.exe zrob_skrot.py            # skrot w folderze projektu
    .venv\Scripts\python.exe zrob_skrot.py --pulpit   # dodatkowo na pulpicie

Skrot zawiera BEZWZGLEDNE sciezki, wiec nie trafia do repozytorium — po
przeniesieniu lub sklonowaniu projektu uruchom ten skrypt ponownie.

Alternatywa bez Pythona na maszynie docelowej: `powershell -File build.ps1`
buduje samodzielny `dist\ProtokolCC.exe`.
"""

import os
import sys

NAZWA = "Generator protokolow.lnk"
OPIS = "Generator protokolow i arkuszy wzorcowania (bez konsoli)"


def utworz(katalog_docelowy, korzen):
    """Tworzy skrot w podanym katalogu. Zwraca sciezke do niego."""
    import win32com.client

    pythonw = os.path.join(korzen, ".venv", "Scripts", "pythonw.exe")
    if not os.path.exists(pythonw):
        raise FileNotFoundError(
            f"Brak {pythonw}\n"
            "Utworz virtualenv (.venv) i zainstaluj requirements.txt.")

    sciezka = os.path.join(katalog_docelowy, NAZWA)
    powloka = win32com.client.Dispatch("WScript.Shell")
    skrot = powloka.CreateShortCut(sciezka)
    skrot.TargetPath = pythonw
    skrot.Arguments = "app_gui.py"
    skrot.WorkingDirectory = korzen
    skrot.Description = OPIS
    ikona = os.path.join(korzen, "app.ico")
    if os.path.exists(ikona):
        skrot.IconLocation = ikona
    skrot.save()
    return sciezka


def main():
    korzen = os.path.dirname(os.path.abspath(__file__))
    cele = [korzen]
    if "--pulpit" in sys.argv:
        cele.append(os.path.join(os.path.expanduser("~"), "Desktop"))

    for katalog in cele:
        try:
            print("utworzono:", utworz(katalog, korzen))
        except Exception as e:                    # noqa: BLE001 — komunikat dla uzytkownika
            print(f"[BLAD] {katalog}: {type(e).__name__}: {e}")
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
