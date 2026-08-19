# -*- coding: utf-8 -*-
"""
testy/wspolne.py — narzedzia wspolne dla testow.

Zasada: KAZDY test, ktory cokolwiek zapisuje, robi to w `testy/_piaskownica/<nazwa>/`
— nigdy w folderze projektu. Piaskownica dostaje KOPIE prawdziwych plikow
wejsciowych, wiec testy dzialaja na realnych danych, nie na atrapach, a mimo to
nie moga uszkodzic protokolow uzytkownika.
"""

import json
import os
import shutil
import stat
import subprocess
import sys

KORZEN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PIASKOWNICA = os.path.join(KORZEN, "testy", "_piaskownica")
PYTHON = sys.executable

# Skrypty-workery uruchamiane jako podprocesy (jak robi to panel).
WORKERY = {
    "analiza":    "analizuj_excele.py",
    "obserwacje": "generuj_obserwacje.py",
    "arkusze":    "generuj_arkusze.py",
}

# Prawdziwe pliki wejsciowe projektu potrzebne poszczegolnym krokom.
PLIKI_OBSERWACJA = [
    "2026-08-06 12.10_188.txt",
    "2026-08-10 13.57_188.txt",
    "xxx_LA_TH_2026 - obserwacje CC.xlsx",
    "szablon_LA_TH_2026 - obserwacje.xlsx",
    "xxx_LA_TH_2026 - protokół CC.xlsx",
    "xxx_LA_TH_2026 - protokół CC-04.xlsx",
    "Zestawienie wzorcowanych przyrządów.xlsx",
]
FOLDERY_OBSERWACJA = ["PZ", "wyniki"]

PLIKI_ARKUSZE = [
    "xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.12 z 17.06.2026 - 1 - RH (CC).xlsx",
    "Wzory.xls",
    "Obliczenia tdp, RH, C.xls",
    "Pom. nr 9 (MX1101-02) - 2026.xlsx",
    "xxx_yyy_LA_TH_2026 - tylko temp.docx",
    "xxx_yyy_LA_TH_2026 - zakres.docx",
    "xxx_yyy_LA_TH_2026 - zakres + temp.docx",
    "xxx_yyy_LA_TH_2026 - tylko temp (uzytkownik).docx",
    "xxx_yyy_LA_TH_2026 - zakres (uzytkownik).docx",
    "xxx_yyy_LA_TH_2026 - zakres + temp (uzytkownik).docx",
]

SZABLON_ARKUSZA = PLIKI_ARKUSZE[0]
PROTOKOL_CC = "188_LA_TH_2026 - protokół CC.xlsx"
OBSERWACJA_CC = "188_LA_TH_2026 - obserwacje CC.xlsx"

# Gotowy, wypelniony protokol lezacy w projekcie. Pozwala testowac krok 3 nawet
# wtedy, gdy pliki TXT poprzedniego zlecenia zostaly juz z folderu usuniete.
PROTOKOL_GOTOWY = "188_LA_TH_2026 - protokół CC_2.xlsx"


def dostepne(nazwy):
    """Ktore z podanych plikow/folderow naprawde sa w projekcie."""
    return [n for n in nazwy if os.path.exists(os.path.join(KORZEN, n))]


def utworz_logi_probne(folder, nazwy=("LOG_A.csv", "LOG_B.csv"), wierszy=120):
    """
    Tworzy proste logi CSV rozpoznawane przez analizuj_excele (format csv_generic).

    Dzieki nim testy kroku 1 dzialaja niezaleznie od tego, czy w projekcie leza
    akurat dane jakiegos zlecenia — uzytkownik sprzata je po zakonczeniu pracy.
    """
    import csv
    import datetime

    os.makedirs(folder, exist_ok=True)
    poczatek = datetime.datetime(2026, 1, 1, 8, 0)
    for nazwa in nazwy:
        with open(os.path.join(folder, nazwa), "w", newline="", encoding="utf-8") as f:
            zapis = csv.writer(f)
            zapis.writerow(["Czas", "Temperatura [°C]", "Wilgotność [%RH]"])
            for i in range(wierszy):
                zapis.writerow([
                    (poczatek + datetime.timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"),
                    round(23 + i * 0.01, 2),
                    round(45 + i * 0.02, 2),
                ])
    return list(nazwy)


def _zdejmij_tylko_do_odczytu(funkcja, sciezka, _wyjatek):
    """
    Handler dla shutil.rmtree. Czesc plikow wzorcowych projektu ('Wzory.xls',
    'Obliczenia tdp, RH, C.xls') ma atrybut tylko-do-odczytu. copy2 przenosi go
    na kopie, przez co kolejne sprzatanie piaskownicy nie moglo ich usunac.
    """
    try:
        os.chmod(sciezka, stat.S_IWRITE)
        funkcja(sciezka)
    except OSError:
        pass


def _kopiuj_zapisywalny(zrodlo, cel):
    """Kopiuje plik i zdejmuje z kopii atrybut tylko-do-odczytu."""
    shutil.copy2(zrodlo, cel)
    try:
        os.chmod(cel, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass


def nowa_piaskownica(nazwa, pliki=(), foldery=()):
    """
    Tworzy czysta piaskownice `testy/_piaskownica/<nazwa>` i kopiuje do niej
    wskazane pliki/foldery z korzenia projektu. Zwraca sciezke bezwzgledna.
    """
    cel = os.path.join(PIASKOWNICA, nazwa)
    if os.path.isdir(cel):
        shutil.rmtree(cel, onexc=_zdejmij_tylko_do_odczytu)
    os.makedirs(cel, exist_ok=True)
    for p in pliki:
        zrodlo = os.path.join(KORZEN, p)
        if os.path.exists(zrodlo):
            _kopiuj_zapisywalny(zrodlo, os.path.join(cel, p))
    for f in foldery:
        zrodlo = os.path.join(KORZEN, f)
        if os.path.isdir(zrodlo):
            shutil.copytree(zrodlo, os.path.join(cel, f),
                            copy_function=_kopiuj_zapisywalny)
    return cel


def uruchom_worker(worker, folder, dodatkowe_env=None, limit_s=900):
    """
    Uruchamia worker tak samo jak panel: osobny proces, ustawienia przez env,
    katalog roboczy = piaskownica. Zwraca (kod_wyjscia, polaczone_wyjscie).
    """
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    env["CC_FOLDER"] = folder
    env.update(dodatkowe_env or {})
    # Zmienne z poprzednich testow nie moga wyciec do tego przebiegu.
    for klucz in [k for k in env if k.startswith(("OBS_", "GEN_", "ANL_"))]:
        if not (dodatkowe_env or {}).get(klucz):
            env.pop(klucz, None)
    env.update(dodatkowe_env or {})

    wynik = subprocess.run(
        [PYTHON, "-u", os.path.join(KORZEN, WORKERY[worker])],
        cwd=folder, env=env, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=limit_s,
    )
    return wynik.returncode, (wynik.stdout or "") + (wynik.stderr or "")


def wartosci_z_modulu(modul, wyrazenia, env=None):
    """
    Importuje modul w OSOBNYM procesie z podanym srodowiskiem i zwraca slownik
    {wyrazenie: str(wartosc)}.

    Osobny proces jest konieczny, bo stale konfiguracyjne skryptow wyliczaja sie
    RAZ, przy imporcie — w jednym procesie nie da sie sprawdzic dwoch roznych
    zestawow ustawien.
    """
    kod = (
        "import json, sys\n"
        f"import {modul} as m\n"
        f"wyr = {list(wyrazenia)!r}\n"
        "print('@@@' + json.dumps({w: str(eval(w, {'m': m})) for w in wyr}))\n"
    )
    srodowisko = dict(os.environ)
    srodowisko["PYTHONIOENCODING"] = "utf-8"
    # Czysty start: zadna zmienna kroku nie moze pochodzic z zewnatrz.
    for klucz in [k for k in srodowisko if k.startswith(("OBS_", "GEN_", "ANL_", "CC_"))]:
        srodowisko.pop(klucz, None)
    srodowisko.update(env or {})

    wynik = subprocess.run([PYTHON, "-c", kod], cwd=KORZEN, env=srodowisko,
                           capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=300)
    if wynik.returncode != 0:
        raise AssertionError(
            f"Import '{modul}' nie powiodl sie (kod {wynik.returncode}):\n"
            f"{wynik.stdout}\n{wynik.stderr}")
    for linia in (wynik.stdout or "").splitlines():
        if linia.startswith("@@@"):
            return json.loads(linia[3:])
    raise AssertionError(f"Brak wyniku z '{modul}':\n{wynik.stdout}\n{wynik.stderr}")
