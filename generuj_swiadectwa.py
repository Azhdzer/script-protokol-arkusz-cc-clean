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
  * data wzorcowania — z PROTOKOLU: zakladka trafia na blok Strony 3 po swoich odczytach
    (E/F przyrzadu i Rs wzorca w C); gdy zadnego bloku nie ma — z K4 zakladki.
    Arkusze skladane recznie z kilku pomiarow miewaja w K4 date innego dnia
    (zlecenie 243: punkty z 17.11 z K4 = 18.11 — 17.11 znikal ze swiadectwa),
  * warunki srodowiskowe — z protokolow z tego folderu i z folderu glownego (tego samego
    zlecenia): zakladka trafia na blok Strony 3 o tej samej dacie i nastawie, najpierw
    z protokolu, w ktorym jest przyrzad, a t/RH pomieszczenia sa brane z F/G protokolu
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
# Uklad Strony 3 z przekladkami (protokoly 'mega' z kroku 6).
from generuj_porownanie import bloki_i_przekladki, wartosc_pomiaru

FOLDER = A.FOLDER
FOLDER_ARKUSZY = C.sciezka("SW_FOLDER", "Etap_5_ark_obl_to_word", FOLDER)
NR_SW_POCZATKOWY = C.calk("SW_NR_SW", 1047)
NR_SW_WG_EWID = C.flaga("SW_NR_SW_WG_EWID", False)

# Pliki w folderze, ktore nie sa arkuszami obliczeniowymi przyrzadow.
_RE_NIE_ARKUSZ = re.compile(r'protok|obserwac|opracowanie|zestawienie', re.I)
# Pierwszy ' - N - ' w nazwie; po nr fabrycznym moze stac jeszcze nr ewidencyjny.
_RE_NR_PRZYRZADU = re.compile(r' - (\d+) - ')


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


_RE_ZLECENIA_PROTOKOLU = re.compile(r'^\s*(\d+(?:\s*,\s*\d+)*)\s*_LA_TH_', re.I)


def sciezki_protokolow(nazwy_arkuszy):
    """
    Protokoly do warunkow srodowiskowych: wszystkie z folderu arkuszy oraz z folderu
    glownego te, ktorych zlecenie ('232_LA_TH_...', '232, 234_LA_TH_...') ma ktorys
    arkusz — tam leza protokoly kolejnych czesci zlecenia (protokol_2, protokol_3...).
    """
    sciezki = [os.path.join(FOLDER_ARKUSZY, n) for n in protokoly_w_folderze(FOLDER_ARKUSZY)]
    if os.path.normcase(os.path.abspath(FOLDER)) == os.path.normcase(os.path.abspath(FOLDER_ARKUSZY)):
        return sciezki
    zlecenia = {A._parsuj_nazwe_pliku(n)[0].strip() for n in nazwy_arkuszy} - {""}
    juz = {os.path.basename(s).lower() for s in sciezki}
    for n in protokoly_w_folderze(FOLDER):
        m = _RE_ZLECENIA_PROTOKOLU.match(n)
        if m and n.lower() not in juz and zlecenia & set(re.split(r'\s*,\s*', m.group(1))):
            sciezki.append(os.path.join(FOLDER, n))
    return sciezki


def odczytaj_arkusz(sciezka):
    """
    Dane przyrzadu z arkusza: {'E5', 'E6', 'G6', 'przetwornik', 'zakladki': [...]}.
    Zakladki punktow (bez 'Wyniki'): {'nazwa', 'K4_val', 'E', 'F', 'C'} — E/F to odczyty
    przyrzadu (E15:F19), C — wzorca (C15:C19); po nich zakladka trafia na blok protokolu.
    Przetwornik — komorki z jego zakladki (K10 'Przetwornik'), jak je wpisal krok 3.
    """
    wb = openpyxl.load_workbook(sciezka, read_only=True, data_only=True)
    try:
        zakladki = [n for n in wb.sheetnames if n != A.ARKUSZ_WYNIKI]
        dane = {"zakladki": [], "przetwornik": None, "E5": None, "E6": None, "G6": None}
        for i, n in enumerate(zakladki):
            ws = wb[n]
            wiersze = list(ws.iter_rows(min_row=15, max_row=19,
                                        min_col=3, max_col=6, values_only=True))
            dane["zakladki"].append({"nazwa": n, "K4_val": ws["K4"].value,
                                     "C": [w[0] for w in wiersze], "E": [w[2] for w in wiersze],
                                     "F": [w[3] for w in wiersze]})
            if i == 0:
                for k in ("E5", "E6", "G6"):
                    dane[k] = ws[k].value
                k10 = str(ws["K10"].value or "").strip()
                if k10 in ("Przetwornik", A.K10_TH_PRZETWORNIK):
                    dane["przetwornik"] = {k: ws[k].value for k in A.KOMORKI_PRZETWORNIKA}
                    # Wyswietlacz + wyjscie (krok 3): swiadectwo z sekcji 'Wynikow'.
                    dane["polaczony"] = k10 == A.K10_TH_PRZETWORNIK
        return dane
    finally:
        wb.close()


def bloki_protokolu(sciezka):
    """
    Bloki Strony 3 protokolu: nastawy (B, C), data i godziny (E), t/RH pomieszczenia
    (F/G) — tak jak czyta je krok 4 — oraz odczyty: 'odczyty' {nr fabryczny: (E[5], F[5])}
    kazdego przyrzadu ze Strony 2 i 'wzorce' (kolumny K..O: Rs multimetrow, tdp).
    Wiersze-przekladki miedzy blokami (protokol 'mega') sa pomijane.
    """
    wb = openpyxl.load_workbook(sciezka, data_only=True)
    try:
        kolumny = _kolumny_przyrzadow(wb, A._czy_protokol_cc04(sciezka))
        seriale = frozenset(kolumny)
        ws = wb[A.ARKUSZ_STRONA3]
        wynik = []
        for nr, (rodzaj, r0) in enumerate(r for r in bloki_i_przekladki(ws) if r[0] == "blok"):
            kolumna = lambda k: [ws.cell(row=r0 + i, column=k).value for i in range(A.BLOK_S3)]
            wynik.append({
                "protokol": os.path.basename(sciezka), "nr": ws.cell(row=r0, column=1).value or nr + 1,
                "seriale": seriale,
                "t": pz_dane._do_float(ws.cell(row=r0, column=2).value),
                "rh": pz_dane._do_float(ws.cell(row=r0, column=3).value),
                "data": A._parse_date_s3(ws.cell(row=r0 + 1, column=5).value),
                "start": ws.cell(row=r0 + 2, column=5).value,
                "koniec": ws.cell(row=r0 + 3, column=5).value,
                "fg": [(ws.cell(row=r0 + k, column=6).value, ws.cell(row=r0 + k, column=7).value)
                       for k in (2, 3)],
                "odczyty": {s: (kolumna(k), kolumna(k + 1)) for s, k in kolumny.items()},
                "wzorce": [kolumna(k) for k in range(11, 16)],
            })
        return wynik
    finally:
        wb.close()


def _klucz_seriala(serial):
    """Nr fabryczny do porownan: 'Q/4' z protokolu = 'Q_4' z nazwy pliku arkusza."""
    return pz_dane.normalizuj_serial(A._bezpieczna_nazwa_pliku(str(serial)))


def _kolumny_przyrzadow(wb, cc04=False):
    """
    {znormalizowany nr fabryczny: kolumna odczytow t na Stronie 3} przyrzadow ze Strony 2
    (kolumna E) — kto byl w protokole i gdzie sa jego odczyty.
    """
    if A.ARKUSZ_STRONA2 not in wb.sheetnames:
        return {}
    ws, wynik, row = wb[A.ARKUSZ_STRONA2], {}, A.START_ROW_S2
    kol0 = A.START_COL_E_S3 + (A.PRZESUNIECIE_STARTU_KOL_CC04 if cc04 else 0)
    # Koniec listy = wiersz pusty w A/B/D/E/F/O (jak w kroku 6) — kolumna 'Nr' (A)
    # bywa niewypelniona (243 protokol CC-04_3), wiec sama nie konczy listy.
    while any(A._cell_to_str(ws.cell(row=row, column=c).value) for c in (1, 2, 4, 5, 6, 15)):
        serial = A._cell_to_str(ws.cell(row=row, column=5).value)
        if serial:
            wynik.setdefault(_klucz_seriala(serial), kol0 + (row - A.START_ROW_S2) * A.KROK_COL_EF)
        row += 1
    return wynik


def _seriale_strony2(wb):
    """Znormalizowane nr fabryczne przyrzadow ze Strony 2 (kolumna E) — kto byl w protokole."""
    return frozenset(_kolumny_przyrzadow(wb))


def _liczba(v):
    v = wartosc_pomiaru(v)
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _te_same_odczyty(zakladka, protokol, minimum=1):
    """Kazda liczba zakladki jest w protokolu na tej samej pozycji (puste w zakladce pomijane)."""
    pary = [(_liczba(a), _liczba(b)) for a, b in zip(zakladka or (), protokol or ())]
    pary = [(a, b) for a, b in pary if a is not None]
    return len(pary) >= minimum and all(b is not None and abs(a - b) < 1e-6 for a, b in pary)


def bloki_po_odczytach(zakladka, bloki, serial):
    """
    Bloki protokolow z tymi samymi odczytami co zakladka: E/F przyrzadu (o tym nr
    fabrycznym) i Rs wzorca (C zakladki = ktoras z kolumn K..O bloku). Niezalezne od K4.
    """
    if not serial or not zakladka.get("E"):
        return []
    klucz = _klucz_seriala(serial)
    c = zakladka.get("C")
    z_c = any(_liczba(v) is not None for v in c or ())
    wynik = []
    for b in bloki:
        od = (b.get("odczyty") or {}).get(klucz)
        if not od or not _te_same_odczyty(zakladka["E"], od[0], minimum=3):
            continue
        if not _te_same_odczyty(zakladka.get("F"), od[1], minimum=0):
            continue
        if z_c and not any(_te_same_odczyty(c, w) for w in b.get("wzorce") or ()):
            continue
        wynik.append(b)
    return wynik


def dopasuj_blok(zakladka, bloki, serial=None):
    """
    Blok protokolu tej zakladki: ta sama data (K4) i nastawa w tolerancji nazw zakladek
    kroku 3 (punkt tylko-temperatura <-> blok bez wilgotnosci). Najblizszy albo None.
    Z podanym nr fabrycznym najpierw protokoly, w ktorych przyrzad jest na Stronie 2
    (zlecenie rozbite na kilka protokolow z tymi samymi punktami tego samego dnia);
    gdy go nie ma w zadnym — dowolny protokol z tym punktem.

    Najpierw po ODCZYTACH (bloki_po_odczytach) — wtedy K4 zakladki nie musi sie zgadzac;
    z kilku takich blokow (ten sam pomiar w kilku protokolach) — ten z data z K4.
    """
    po_odczytach = bloki_po_odczytach(zakladka, bloki, serial)
    if po_odczytach:
        k4 = A._parse_date_s3(zakladka.get("K4_val"))
        return min(po_odczytach, key=lambda b: b["data"] != k4)
    if serial:
        klucz = _klucz_seriala(serial)
        swoje = [b for b in bloki if klucz in b.get("seriale", ())]
        if swoje:
            return dopasuj_blok(zakladka, swoje) or dopasuj_blok(zakladka, bloki)
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


def _data_z_protokolu(zakladka, blok):
    """Data wzorcowania zakladki = data bloku protokolu (K4 tylko gdy bloku nie ma)."""
    if blok.get("data") is None or A._parse_date_s3(zakladka.get("K4_val")) == blok["data"]:
        return
    nowa = blok["data"].strftime("%d.%m.%Y")
    stara = A._parse_date_s3(zakladka.get("K4_val"))
    print(f"    [Data] '{zakladka['nazwa']}': z protokolu {nowa} ({blok['protokol']}, punkt {blok['nr']}), "
          f"w K4 arkusza {stara.strftime('%d.%m.%Y') if stara else 'brak'}.")
    zakladka["K4_val"] = nowa


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
        "POLACZONY": bool(dane.get("polaczony")),
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

    protokoly = sciezki_protokolow(nazwy)
    otwarte = [p for p in protokoly if os.path.exists(os.path.join(
        os.path.dirname(p), "~$" + os.path.basename(p)))]
    if otwarte:
        A._warn("Protokol otwarty w Excelu — czytam ostatnio ZAPISANA wersje: "
                + ", ".join(os.path.basename(p) for p in otwarte))
    bloki = []
    for p in protokoly:
        b = bloki_protokolu(p)
        bloki += b
        print(f"  Protokol: {os.path.relpath(p, FOLDER)} — punktow: {len(b)}, "
              f"przyrzadow: {len(b[0]['seriale']) if b else 0}")
    if not protokoly:
        A._warn("Brak protokolu (w folderze arkuszy ani w glownym) — swiadectwa BEZ warunkow "
                "srodowiskowych. Wloz protokol(y) z pomiarow, zeby je policzyc.")

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
                blok = dopasuj_blok(zd, bloki, rekordy[j]["E"])
                if blok is None:
                    bez_bloku.append(zd["nazwa"])
                    continue
                _data_z_protokolu(zd, blok)
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
