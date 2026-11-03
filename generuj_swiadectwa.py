# -*- coding: utf-8 -*-
"""
generuj_swiadectwa.py — krok 5: swiadectwa Word z GOTOWYCH arkuszy obliczeniowych.

Krok 3 robi arkusze i od razu swiadectwa. Bywa jednak, ze arkusze powstaja bez Worda,
a potem sa uzupelniane (domierzone punkty, poprawki) — czasem przez kilka pomiarow.
Ten krok konczy obieg: bierze arkusze z folderu 'Etap_5_ark_obl_to_word' i robi
swiadectwa tak samo jak krok 3 (te same funkcje z generuj_arkusze):

  * wyniki kalibracji — z arkusza (zakladki punktow po przeliczeniu w Excelu),
  * producent / typ / nr fabryczny / zlecenie — z arkusza (E5, E6, nazwa pliku),
  * nr ewidencyjny, uzytkownik, zleceniodawca, zasilanie — z PZ (folder PZ albo ten folder),
  * data wzorcowania — z K4 zakladek,
  * warunki srodowiskowe — z protokolow lezacych w tym folderze: zakladka trafia na blok
    Strony 3 o tej samej dacie i nastawie, a t/RH pomieszczenia sa brane z F/G protokolu
    albo z pliku czujnika (jak w kroku 3/4). Kilka protokolow = kilka sesji pomiarowych.

Swiadectwa zapisywane sa obok arkuszy. Arkusze nie sa zmieniane (linki zewnetrzne wracaja
na serwer, poprzednia wersja w 'kopie_zapasowe').
"""

import datetime
import os
import re
import sys

import openpyxl

import cc_config as C
import cc_kopia
import pz_dane
# Te same funkcje i stale co w kroku 3 — swiadectwo ma byc identyczne.
import generuj_arkusze as A

FOLDER = A.FOLDER
FOLDER_ARKUSZY = C.sciezka("SW_FOLDER", "Etap_5_ark_obl_to_word", FOLDER)
NR_SW_POCZATKOWY = C.calk("SW_NR_SW", 1047)
NR_SW_WG_EWID = C.flaga("SW_NR_SW_WG_EWID", False)

# Pliki w folderze, ktore nie sa arkuszami obliczeniowymi przyrzadow.
_RE_NIE_ARKUSZ = re.compile(r'protok|obserwac|opracowanie|zestawienie', re.I)
_RE_NR_PRZYRZADU = re.compile(r' - (\d+) - [^-]+$')


def _nazwy_arkuszy(sciezka):
    try:
        wb = openpyxl.load_workbook(sciezka, read_only=True)
    except Exception:
        return []
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _klucz_kolejnosci(nazwa):
    """Wg zlecenia i numeru przyrzadu z nazwy ('... - 4 - TMM160500533.xlsx'), jak w kroku 3."""
    prefiks, _rok, _nr, _typ = A._parsuj_nazwe_pliku(nazwa)
    m = _RE_NR_PRZYRZADU.search(os.path.splitext(nazwa)[0])
    return (prefiks, int(m.group(1)) if m else 10 ** 6, nazwa)


def arkusze_w_folderze(folder):
    """Arkusze obliczeniowe przyrzadow (z zakladka 'Wyniki') w kolejnosci przyrzadow."""
    if not os.path.isdir(folder):
        return []
    wynik = []
    for n in os.listdir(folder):
        if n.startswith(("~$", "xxx")) or not n.lower().endswith((".xlsx", ".xlsm")):
            continue
        if _RE_NIE_ARKUSZ.search(n):
            continue
        if A.ARKUSZ_WYNIKI in _nazwy_arkuszy(os.path.join(folder, n)):
            wynik.append(n)
    return sorted(wynik, key=_klucz_kolejnosci)


def protokoly_w_folderze(folder):
    """Protokoly (z 'Strona 3') — zrodlo godzin punktow do warunkow srodowiskowych."""
    if not os.path.isdir(folder):
        return []
    return sorted(n for n in os.listdir(folder)
                  if not n.startswith("~$") and n.lower().endswith((".xlsx", ".xlsm"))
                  and "protok" in n.lower()
                  and A.ARKUSZ_STRONA3 in _nazwy_arkuszy(os.path.join(folder, n)))


def odczytaj_arkusz(sciezka):
    """
    Dane przyrzadu z arkusza: {'E5', 'E6', 'G6', 'przetwornik', 'zakladki': [...]}.
    Zakladki punktow (bez 'Wyniki'): {'nazwa', 'K4_val'}. Przetwornik — komorki z jego
    zakladki (K10 'Przetwornik'), jak je wpisal krok 3.
    """
    wb = openpyxl.load_workbook(sciezka, read_only=True, data_only=True)
    try:
        zakladki = [n for n in wb.sheetnames if n != A.ARKUSZ_WYNIKI]
        dane = {"zakladki": [], "przetwornik": None, "E5": None, "E6": None, "G6": None}
        for i, n in enumerate(zakladki):
            ws = wb[n]
            dane["zakladki"].append({"nazwa": n, "K4_val": ws["K4"].value})
            if i == 0:
                for k in ("E5", "E6", "G6"):
                    dane[k] = ws[k].value
                if str(ws["K10"].value or "").strip() == "Przetwornik":
                    dane["przetwornik"] = {k: ws[k].value for k in A.KOMORKI_PRZETWORNIKA}
        return dane
    finally:
        wb.close()


def bloki_protokolu(sciezka):
    """
    Bloki Strony 3 protokolu: nastawy (B, C), data i godziny (E), t/RH pomieszczenia
    (F/G) — tak jak czyta je krok 4.
    """
    wb = openpyxl.load_workbook(sciezka, read_only=True, data_only=True)
    try:
        ws = wb[A.ARKUSZ_STRONA3]
        wynik, nr = [], 0
        while True:
            r0 = A.START_ROW_S3 + nr * A.BLOK_S3
            if ws.cell(row=r0, column=1).value in (None, ""):
                return wynik
            wynik.append({
                "protokol": os.path.basename(sciezka), "nr": nr + 1,
                "t": pz_dane._do_float(ws.cell(row=r0, column=2).value),
                "rh": pz_dane._do_float(ws.cell(row=r0, column=3).value),
                "data": A._parse_date_s3(ws.cell(row=r0 + 1, column=5).value),
                "start": ws.cell(row=r0 + 2, column=5).value,
                "koniec": ws.cell(row=r0 + 3, column=5).value,
                "fg": [(ws.cell(row=r0 + k, column=6).value, ws.cell(row=r0 + k, column=7).value)
                       for k in (2, 3)],
            })
            nr += 1
    finally:
        wb.close()


def dopasuj_blok(zakladka, bloki):
    """
    Blok protokolu tej zakladki: ta sama data (K4) i nastawa w tolerancji nazw zakladek
    kroku 3 (punkt tylko-temperatura <-> blok bez wilgotnosci). Najblizszy albo None.
    """
    data = A._parse_date_s3(zakladka.get("K4_val"))
    t = A._t_z_nazwy_zakladki(zakladka["nazwa"])
    rh = A._rh_z_nazwy_zakladki(zakladka["nazwa"])
    if data is None or t is None:
        return None
    pasujace = []
    for b in bloki:
        if b["data"] != data or b["t"] is None or abs(b["t"] - t) > A.TOL_NAZWA_T:
            continue
        rh_b = b["rh"] if b["rh"] not in (None, 0.0) else None
        if (rh is None) != (rh_b is None):
            continue
        if rh is not None and abs(rh - rh_b) > A.TOL_NAZWA_RH:
            continue
        pasujace.append((abs(b["t"] - t) + (abs(rh - rh_b) / 10.0 if rh is not None else 0), b))
    return min(pasujace, key=lambda x: x[0])[1] if pasujace else None


def th_bloku(blok, dane_miesiaca):
    """(t, RH) pomieszczenia na start i koniec bloku: z F/G protokolu, a gdy puste — z czujnika."""
    wynik = []
    for (t, rh), czas_val in zip(blok["fg"], (blok["start"], blok["koniec"])):
        if t not in (None, ""):
            wynik.append((t, rh if rh not in ("",) else None))
            continue
        czas = A._parse_time_s3(czas_val)
        if czas is None or blok["data"] is None:
            continue
        moment = datetime.datetime.combine(blok["data"], czas)
        t, rh, _odch = A._szukaj_th_w_danych_czujnika(
            dane_miesiaca(blok["data"].year, blok["data"].month), moment)
        if t is not None:
            wynik.append((t, rh))
    return wynik


def mapa_pz():
    """PZ z folderu PZ kroku 3 oraz z folderu arkuszy (gdy PZ przeniesiono razem z nimi)."""
    mapa = dict(A._pz_mapa_arkusze())
    if any(n.lower().endswith((".pdf", ".lnk")) and n.upper().startswith("PZ")
           for n in os.listdir(FOLDER_ARKUSZY)):
        mapa.update(pz_dane.wczytaj_pz(FOLDER_ARKUSZY)[0])
    A._PZ_CACHE = mapa
    return mapa


def rekord_arkusza(nazwa, dane, pz):
    """Rekord jak ze Strony 2 protokolu — tyle, ile potrzebuje utworz_kopie_word."""
    serial = A._parsuj_nazwe_pliku(nazwa)[2] or A._cell_to_str(dane.get("G6"))
    dev = pz.get(pz_dane.normalizuj_serial(serial))
    return {
        "E": serial,
        "F": (dev.nr_ewid if dev else "") or "",
        "B": dane.get("E5"),
        "D": dane.get("E6"),
        "P": (getattr(dev, "uwagi", "") if dev else "") or "",
        "PRZETWORNIK": dane.get("przetwornik"),
    }


def main():
    print("=" * 65)
    print("  Krok 5: swiadectwa Word z gotowych arkuszy obliczeniowych")
    print("=" * 65)
    print(f"  Folder arkuszy: {FOLDER_ARKUSZY}")
    if not os.path.isdir(FOLDER_ARKUSZY):
        print(f"[BŁĄD] Folder nie istnieje: {FOLDER_ARKUSZY}")
        return 1
    if not A._DOCX_OK:
        print("[BŁĄD] Brak biblioteki 'python-docx'.")
        return 1

    nazwy = arkusze_w_folderze(FOLDER_ARKUSZY)
    if not nazwy:
        print("[BŁĄD] Brak arkuszy obliczeniowych (pliki .xlsx z zakladka 'Wyniki').")
        return 1
    otwarte = [n for n in nazwy if os.path.exists(os.path.join(FOLDER_ARKUSZY, "~$" + n))]
    if otwarte:
        print("[BŁĄD] Arkusze OTWARTE w Excelu — zamknij je i uruchom krok ponownie:")
        for n in otwarte:
            print(f"    • {n}")
        return 1
    print(f"  Arkuszy: {len(nazwy)}")

    protokoly = protokoly_w_folderze(FOLDER_ARKUSZY)
    bloki = []
    for p in protokoly:
        bloki += bloki_protokolu(os.path.join(FOLDER_ARKUSZY, p))
    if protokoly:
        print(f"  Protokoly (godziny punktow): {', '.join(protokoly)} — blokow: {len(bloki)}")
    else:
        A._warn("Brak protokolu w folderze — swiadectwa BEZ warunkow srodowiskowych. "
                "Wloz do folderu protokol(y) z pomiarow, zeby je policzyc.")

    pz = mapa_pz()
    dane_arkuszy = [odczytaj_arkusz(os.path.join(FOLDER_ARKUSZY, n)) for n in nazwy]
    rekordy = [rekord_arkusza(n, d, pz) for n, d in zip(nazwy, dane_arkuszy)]

    cache_czujnika = {}

    def dane_miesiaca(rok, miesiac):
        if (rok, miesiac) not in cache_czujnika:
            plik = A._sciezka_pliku_czujnika(A.NR_POMIESZCZENIA, A.MODEL_CZUJNIKA, rok)
            cache_czujnika[(rok, miesiac)] = A._wczytaj_dane_czujnika_miesiac(plik, rok, miesiac)
        return cache_czujnika[(rok, miesiac)]

    kalibracje, klasy, zakresy = [], [], []
    app = linked = None
    sciezki = {}
    try:
        app, linked, sciezki = A._nowa_sesja_excel(FOLDER)
        for j, (nazwa, dane) in enumerate(zip(nazwy, dane_arkuszy)):
            sciezka = os.path.join(FOLDER_ARKUSZY, nazwa)
            print(f"\n[{j + 1}/{len(nazwy)}] {nazwa}")
            # Odczyt przepina linki pliku na lokalne Wzory/Obliczenia (jak krok 3) —
            # poprzednia wersja arkusza zostaje w kopiach zapasowych.
            cc_kopia.kopia_przed_nadpisaniem(sciezka)
            kal, klasa = A._odczytaj_kalibracje_xlwings(
                app, sciezka, A.ARKUSZ_WYNIKI, enable_diag=(j == 0), sciezki_linkowane=sciezki)
            kalibracje.append(kal or [])
            klasy.append(klasa)

            dane_th, bez_bloku = [], []
            for zd in dane["zakladki"]:
                if zd["nazwa"] not in {p["nazwa"] for p in (kal or [])}:
                    continue                      # zakladka bez pomiarow — poza swiadectwem
                blok = dopasuj_blok(zd, bloki)
                if blok is None:
                    bez_bloku.append(zd["nazwa"])
                    continue
                th = th_bloku(blok, dane_miesiaca)
                print(f"    [Srodowisko] '{zd['nazwa']}' -> {blok['protokol']}, punkt {blok['nr']}: "
                      + (", ".join(f"{t}/{rh}" for t, rh in th) or "BRAK danych czujnika"))
                dane_th += th
            if bez_bloku and protokoly:
                A._warn(f"Zakladki bez punktu w protokolach (data/nastawa): {', '.join(bez_bloku)} — "
                        f"ich warunki srodowiskowe NIE sa w zakresie swiadectwa.", indent="    ")
            zakresy.append(A._zakresy_z_wzorow(app, dane_th, A.NR_POMIESZCZENIA, A.MODEL_CZUJNIKA)
                           if dane_th else None)
    finally:
        A._zamknij_sesje_excel(app, linked)
        if A.LINKI_SERWEROWE:
            for nazwa in nazwy:
                try:
                    A._przywroc_linki_w_xml(os.path.join(FOLDER_ARKUSZY, nazwa), A.LINKI_SERWEROWE,
                                            cicho=True)
                except Exception as exc:
                    A._warn(f"Nie przywrocono linkow serwerowych w '{nazwa}': {exc}")

    print("\n" + "=" * 65)
    print("  Swiadectwa Word")
    print("=" * 65)
    utworzone = A.utworz_kopie_word(
        FOLDER, A.SZABLON_WORD_TYLKO_TEMP, A.SZABLON_WORD_Z_RH, A.SZABLON_WORD_MIESZANY,
        rekordy, nazwy, [], kalibracje, NR_SW_POCZATKOWY,
        dane_zakladek_per_kopia=[d["zakladki"] for d in dane_arkuszy],
        klasa_wilg_per_kopia=klasy, warunki_per_kopia=zakresy,
        wg_ewid=NR_SW_WG_EWID, folder_wyj=FOLDER_ARKUSZY)
    print("=" * 65)
    print(f"  Utworzono {len(utworzone)} swiadectw w: {FOLDER_ARKUSZY}")
    print("=" * 65)
    return 0 if len(utworzone) == len(nazwy) else 1


if __name__ == "__main__":
    sys.exit(main())
