# -*- coding: utf-8 -*-
"""
generuj_porownanie.py — KROK 6: arkusze obliczeniowe porownawcze.

Wejscie: jeden (albo kilka) protokolow CC / CC-04. Wyjscie: dla kazdego przyrzadu
arkusz obliczeniowy (ten sam szablon co w kroku 3) z zakladka dla KAZDEGO bloku,
w ktorym przyrzad ma odczyty — takze powtorzonych punktow ('70, -' i '70, - (2)')
i komorek szarych. W 'Wyniki' kazda zakladka ma swoj wiersz. Arkusz sluzy do
porownania, jak licza sie poszczegolne powtorzenia — swiadectw z niego nie ma.

Czym rozni sie "mega protokol" od zwyklego (zlecenie 243, protokol CC-04_1):
  * miedzy blokami punktow sa PRZEKLADKI — wiersze bez numeru w kol. A, w ktorych
    przyrzad dostaje nowy czujnik wzorcowy ('CC-04-LG' ...) i odleglosc ('≤ 4 cm').
    Przyrzad przeniesiony w inne miejsce komory od tego bloku liczy sie wzgledem
    innego Pt100. '-' w przekladce = bez zmian;
  * bloki nie leza co 5 wierszy od wiersza 20 — krok 3 zatrzymuje sie na pierwszej
    przekladce. Tu bloki szukane sa po numerze w kol. A;
  * te same punkty powtarzaja sie wiele razy, czesto bez wyszarzenia.

Protokol czytany jest openpyxl (dane w Stronie 3 to wartosci, nie formuly) — bez
Excela. Excel uruchamia dopiero wypelnianie kopii (generuj_arkusze.utworz_kopie),
tak samo jak w kroku 3, ale BEZ warunkow srodowiskowych i bez odczytu kalibracji.

Kilka protokolow: przyrzady laczone po nr fabrycznym (Strona 2, kol. E), bloki
kolejnych protokolow dopisywane po kolei — przyrzad z dwoch protokolow dostaje
zakladki z obu.
"""

import colorsys
import os
import re
import sys

import openpyxl

import cc_config as C
import cc_excel
import cc_widok
import generuj_arkusze as G
import pz_dane

FOLDER = G.FOLDER
PROTOKOLY = C.lista("POR_PROTOKOLY", [])
SZABLON = C.tekst("POR_SZABLON", "") or G.SZABLON_PLIK
# Folder kroku 6: tu leza protokoly do porownania i tu powstaja arkusze porownawcze
# (jak 'Etap_5_ark_obl_to_word' w kroku 5).
FOLDER_WYJ = C.sciezka("POR_FOLDER", "Etap_6_porownanie", FOLDER)
UWZGLEDNIJ_SZARE = C.flaga("POR_SZARE", True)
# '2, 5, 21-24' (pozycje ze Strony 2) albo numery fabryczne / ewidencyjne; puste = wszystkie.
WYBOR_PRZYRZADOW = C.tekst("POR_PRZYRZADY", "")

ARKUSZ_S1, ARKUSZ_S2, ARKUSZ_S3 = "Strona 1", G.ARKUSZ_STRONA2, G.ARKUSZ_STRONA3
WIERSZ_ODLEGLOSCI_S3 = 17
# Bez numeru bloku przez tyle wierszy z rzedu = koniec tabeli punktow.
MAKS_PUSTYCH_WIERSZY = 30


# ─── Kolory komorek (openpyxl) ───────────────────────────────────────────────
# Szare komorki protokolu bywaja zapisane jako kolor motywu z odcieniem
# ('Tlo 1, ciemniejsze 25%' = theme 0, tint -0.25), a nie jako #BFBFBF.
_KOLEJNOSC_MOTYWU = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                     "accent4", "accent5", "accent6", "hlink", "folHlink")
_MOTYW_DOMYSLNY = {"lt1": "FFFFFF", "dk1": "000000", "lt2": "E7E6E6", "dk2": "44546A",
                   "accent1": "4472C4", "accent2": "ED7D31", "accent3": "A5A5A5",
                   "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47",
                   "hlink": "0563C1", "folHlink": "954F72"}
_RE_KOLOR_MOTYWU = re.compile(
    rb'<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*'
    rb'<a:(?:srgbClr\s+val|sysClr\s[^>]*?lastClr)="([0-9A-Fa-f]{6})"')


def kolory_motywu(wb):
    """Kolory motywu skoroszytu w kolejnosci indeksow Excela (0 = Tlo 1, 1 = Tekst 1 ...)."""
    xml = getattr(wb, "loaded_theme", None) or b""
    mapa = {n.decode(): v.decode().upper() for n, v in _RE_KOLOR_MOTYWU.findall(xml)}
    return [mapa.get(k, _MOTYW_DOMYSLNY[k]) for k in _KOLEJNOSC_MOTYWU]


def z_odcieniem(rgb6, tint):
    """Kolor po odcieniu (tint) jak w Excelu: <0 przyciemnia, >0 rozjasnia."""
    r, g, b = (int(rgb6[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    return "#" + "".join(f"{round(x * 255):02X}" for x in colorsys.hls_to_rgb(h, l, s))


def kolor_komorki(cell, motyw):
    """'#RRGGBB' wypelnienia komorki albo None (brak wypelnienia)."""
    f = cell.fill
    if f is None or f.fill_type in (None, "none"):
        return None
    c = f.fgColor
    try:
        if c.type == "rgb" and isinstance(c.rgb, str):
            return "#" + c.rgb[-6:].upper()
        if c.type == "theme" and isinstance(c.theme, int) and c.theme < len(motyw):
            return z_odcieniem(motyw[c.theme], float(c.tint or 0))
        if c.type == "indexed" and isinstance(c.indexed, int):
            from openpyxl.styles.colors import COLOR_INDEX
            if c.indexed < len(COLOR_INDEX):
                return "#" + COLOR_INDEX[c.indexed][-6:].upper()
    except (TypeError, ValueError):
        return None
    return None


def czy_szara(cell, motyw):
    return (kolor_komorki(cell, motyw) or "").upper() == G.KOLOR_POMIJANY_S3.upper()


# ─── Wartosci ────────────────────────────────────────────────────────────────

def _format(cell):
    fmt = cell.number_format
    return None if not fmt or fmt.upper() in ("GENERAL", "@") else fmt


def wartosc_pomiaru(v):
    """Liczba albo liczba zapisana tekstem; tekst opisowy ('Przestal rejestrowac') -> None."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", ".")
    try:
        float(s)
    except ValueError:
        return None
    return v


def _liczba(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _numer_bloku(v):
    if _liczba(v):
        return True
    return isinstance(v, str) and v.strip().isdigit()


# ─── Uklad Strony 3: bloki i przekladki ──────────────────────────────────────

def bloki_i_przekladki(ws, start=G.START_ROW_S3, blok=G.BLOK_S3):
    """
    Wiersze Strony 3 w kolejnosci: [('blok', r0), ('przekladka', r), ...].

    Blok zaczyna sie od numeru w kol. A i ma `blok` wierszy. Wiersz bez numeru
    miedzy blokami to (potencjalna) przekladka. Tekst w kol. A ('minimum:') albo
    dluga przerwa konczy tabele punktow.
    """
    wynik, r, puste = [], start, 0
    while r <= ws.max_row and puste < MAKS_PUSTYCH_WIERSZY:
        a = ws.cell(r, 1).value
        if _numer_bloku(a):
            wynik.append(("blok", r))
            r += blok
            puste = 0
            continue
        if isinstance(a, str) and a.strip():
            break
        wynik.append(("przekladka", r))
        puste += 1
        r += 1
    # Przekladki po ostatnim bloku niczego nie zmieniaja.
    while wynik and wynik[-1][0] == "przekladka":
        wynik.pop()
    return wynik


def _odleglosc(v):
    s = G._cell_to_str(v)
    return s if "cm" in s.lower() else None


# ─── Odczyt jednego protokolu ────────────────────────────────────────────────

def wczytaj_protokol(sciezka, uwzglednij_szare=True):
    """
    Czyta protokol (openpyxl, bez Excela). Zwraca (dane_s2, dane_zakladek, dane_ef, f24)
    w ukladzie generuj_arkusze.wczytaj_wszystko_xlwings, z dodatkami:
      * zakladka: 'wiersz_s3' (pierwszy wiersz bloku), 'protokol',
      * blok przyrzadu: 'CC04_TAG' (czujnik wzorcowy WAZNY W TYM BLOKU),
        'SZARY' (odczyty wziete z szarych komorek), 'ODLEGLOSC',
      * rekord: '_nr_przyrzadu' (pozycja na Stronie 2), '_kol_s3'.
    """
    nazwa = os.path.basename(sciezka)
    wb = openpyxl.load_workbook(sciezka, data_only=True)
    motyw = kolory_motywu(wb)
    cc04 = G._czy_protokol_cc04(sciezka)
    przes = G.PRZESUNIECIE_STARTU_KOL_CC04 if cc04 else 0
    col_e0, krok = G.START_COL_E_S3 + przes, G.KROK_COL_EF
    blok = G.BLOK_S3

    # --- Strona 2: przyrzady ---
    ws2 = wb[ARKUSZ_S2]
    dane_s2 = []
    r = G.START_ROW_S2
    while r <= ws2.max_row:
        wartosci = {k: ws2.cell(r, c).value for k, c in
                    (("B", 2), ("D", 4), ("E", 5), ("F", 6), ("K", 11), ("L", 12),
                     ("M", 13), ("N", 14), ("O", 15), ("P", 16))}
        if all(G._cell_to_str(wartosci[k]) == "" for k in ("B", "D", "E", "F", "O")):
            break
        nr = r - G.START_ROW_S2 + 1
        if wartosci["O"] is None or wartosci["E"] is None:
            G._warn(f"{nazwa}: Strona 2, wiersz {r} — brak nr zlecenia (O) lub nr fabrycznego (E), "
                    f"przyrzad {nr} pominiety.")
            r += 1
            continue
        rek = dict(wartosci)
        rek["O"], rek["E"] = G._cell_to_str(wartosci["O"]), G._cell_to_str(wartosci["E"])
        rek.update({"IS_CC04_PROTO": cc04, "CC04_RAW": "", "CC04_TAG": None,
                    "_nr_przyrzadu": nr, "_kol_s3": col_e0 + (nr - 1) * krok,
                    "_protokol": nazwa})
        dane_s2.append(rek)
        r += 1

    # --- Strona 3: bloki, przekladki, odczyty ---
    ws3 = wb[ARKUSZ_S3]
    uklad = bloki_i_przekladki(ws3)
    tag = {}
    odl = {}
    for rek in dane_s2:
        k = rek["_kol_s3"]
        rek["JEDN_T"] = G._cell_to_str(ws3.cell(G.WIERSZ_JEDNOSTEK_S3, k).value)
        rek["JEDN_RH"] = G._cell_to_str(ws3.cell(G.WIERSZ_JEDNOSTEK_S3, k + 1).value)
        if cc04:
            raw = ws3.cell(G.WIERSZ_TYPU_CC04_S3, k).value
            rek["CC04_RAW"] = G._cell_to_str(raw)
            tag[k] = G._wykryj_tag_cc04(raw)
        odl[k] = _odleglosc(ws3.cell(WIERSZ_ODLEGLOSCI_S3, k).value)

    dane_zakladek = []
    dane_ef = [[] for _ in dane_s2]
    przekladki, bez_czujnika, teksty = 0, [], set()
    col_o25 = 18 if cc04 else 16
    for rodzaj, r0 in uklad:
        if rodzaj == "przekladka":
            zmiana = False
            for rek in dane_s2:
                k = rek["_kol_s3"]
                v = ws3.cell(r0, k).value
                t = G._wykryj_tag_cc04(v) if cc04 else None
                if t:
                    tag[k], zmiana = t, True
                d = _odleglosc(v)
                if d:
                    odl[k], zmiana = d, True
            przekladki += zmiana
            continue

        raw_b, raw_c = ws3.cell(r0, 2).value, ws3.cell(r0, 3).value
        o25 = G._cell_to_str(ws3.cell(r0, col_o25).value)
        zd = {
            "nazwa": f"{G._cell_to_sheet_name_part(raw_b)}, {G._cell_to_sheet_name_part(raw_c)}",
            "block_idx": len(dane_zakladek),
            "wiersz_s3": r0,
            "protokol": nazwa,
            "L_dane": [ws3.cell(r0 + i, 12).value for i in range(blok)],
            "L_fmt": [_format(ws3.cell(r0 + i, 12)) for i in range(blok)],
            "M_dane": [ws3.cell(r0 + i, 13).value for i in range(blok)],
            "M_fmt": [_format(ws3.cell(r0 + i, 13)) for i in range(blok)],
            "K4_val": ws3.cell(r0 + 1, 5).value,
            "czas_start": ws3.cell(r0 + 2, 5).value,
            "czas_koniec": ws3.cell(r0 + 3, 5).value,
            "o25_val": "N/A" if o25 in ("", "-") else o25,
            "nast_t": pz_dane._do_float(raw_b),
            "nast_rh": pz_dane._do_float(raw_c),
        }
        dane_zakladek.append(zd)

        for j, rek in enumerate(dane_s2):
            k = rek["_kol_s3"]
            komorki = {kol: [ws3.cell(r0 + i, kol) for i in range(blok)] for kol in (k, k + 1)}
            szare = {kol: [czy_szara(c, motyw) for c in komorki[kol]] for kol in komorki}
            # Szara temperatura = punkt wylaczony dla przyrzadu -> do porownania bierzemy.
            # Szara wilgotnosc tylko razem z szara temperatura w bloku Z nastawa RH;
            # sama szara RH (np. punkt samej temperatury '70, -') to wilgotnosc, ktorej
            # sie nie wzorcuje — wzieta zrobilaby z zakladki temperaturowej punkt RH.
            caly_punkt_szary = zd["nast_rh"] is not None and any(
                s for s, c in zip(szare[k], komorki[k]) if wartosc_pomiaru(c.value) is not None)
            e_dane, f_dane, szary = [], [], False
            for kol, cel in ((k, e_dane), (k + 1, f_dane)):
                bierz_szare = uwzglednij_szare and (kol == k or caly_punkt_szary)
                for c, s in zip(komorki[kol], szare[kol]):
                    v = wartosc_pomiaru(c.value)
                    if v is None and G._czy_wartosc_niepusta(c.value) and c.value != "-":
                        teksty.add(G._cell_to_str(c.value))
                    if v is not None and s:
                        if bierz_szare:
                            szary = True
                        else:
                            v = None
                    cel.append(v)
            ef = {"E_dane": e_dane, "E_fmt": [_format(ws3.cell(r0 + i, k)) for i in range(blok)],
                  "F_dane": f_dane, "F_fmt": [_format(ws3.cell(r0 + i, k + 1)) for i in range(blok)],
                  "SZARY": szary, "ODLEGLOSC": odl.get(k)}
            if cc04:
                t = tag.get(k)
                ef["CC04_TAG"] = t
                if t and G._czy_blok_ef_aktywny(ef):
                    kol_c = G.MAPOWANIE_KOLUMNY_C_CC04.get(t, 12)
                    ef["C_dane"] = [ws3.cell(r0 + i, kol_c).value for i in range(blok)]
                    ef["C_fmt"] = [_format(ws3.cell(r0 + i, kol_c)) for i in range(blok)]
                    ef["D_dane"] = [ws3.cell(r0 + i, G.KOLUMNA_D_CC04_S3).value for i in range(blok)]
                    ef["D_fmt"] = [_format(ws3.cell(r0 + i, G.KOLUMNA_D_CC04_S3)) for i in range(blok)]
                elif G._czy_blok_ef_aktywny(ef):
                    # Odczyty jest, czujnika wzorcowego nie — nie ma wzgledem czego liczyc.
                    bez_czujnika.append((rek["_nr_przyrzadu"], zd["nazwa"], r0))
                    ef = {}
            dane_ef[j].append(ef)

    # Rekord: czujnik i odleglosc z PIERWSZEGO bloku z odczytami (nazwa, log, F24 w Wyniki).
    f24 = []
    for j, rek in enumerate(dane_s2):
        pierwszy = next((ef for ef in dane_ef[j] if G._czy_blok_ef_aktywny(ef)), None)
        if pierwszy and pierwszy.get("CC04_TAG"):
            rek["CC04_TAG"] = pierwszy["CC04_TAG"]
            rek["CC04_RAW"] = f"CC-04-{pierwszy['CC04_TAG']}"
        f24.append((pierwszy or {}).get("ODLEGLOSC") or odl.get(rek["_kol_s3"]))

    try:
        multimetr_s1 = wb[ARKUSZ_S1][G.KOMORKA_MULTIMETRU_S1].value if ARKUSZ_S1 in wb.sheetnames else None
    except Exception:
        multimetr_s1 = None
    wb.close()

    print(f"  [{nazwa}] przyrzadow: {len(dane_s2)}, blokow: {len(dane_zakladek)}, "
          f"przekladek ze zmiana czujnika/odleglosci: {przekladki}")
    for nr, nz, r0 in bez_czujnika:
        G._warn(f"przyrzad {nr}, blok '{nz}' (wiersz {r0}): odczyty sa, ale brak czujnika "
                f"wzorcowego CC-04 (w wierszu 14 / przekladce '-') — blok pominiety.")
    if teksty:
        print(f"    Komorki z tekstem zamiast odczytu (pominiete): {', '.join(sorted(teksty))}")
    return dane_s2, dane_zakladek, dane_ef, f24, multimetr_s1


# ─── Nazwy zakladek ──────────────────────────────────────────────────────────

def nazwy_zakladek_porownania(dane_zakladek, punkty=None):
    """
    Nazwa bloku = NAJBLIZSZY punkt z PZ w tolerancji (G.TOL_NAZWA_T / _RH), a gdy
    zaden nie pasuje — nastawa komory, jak w kroku 3.

    W odroznieniu od kroku 3 punkt z PZ NIE jest zuzywany: powtorzenia (25,1/25,5
    i 25/25,5; 60,6/25,5 i 61,2/25,5) dostaja te sama nazwe, a w arkuszu przyrzadu
    sufiks ' (2)', ' (3)' — od razu widac, ktore zakladki porownywac.
    """
    if punkty is None:
        punkty = G._pz_punkty_arkusze()
    if not punkty:
        return 0
    zmienione = 0
    for zd in dane_zakladek:
        t_n, rh_n = zd.get("nast_t"), zd.get("nast_rh")
        if t_n is None:
            continue
        najlepszy = None
        for t_pz, rh_pz in punkty:
            if t_pz is None or abs(t_n - t_pz) > G.TOL_NAZWA_T:
                continue
            if rh_pz is None:
                if rh_n not in (None, 0.0):
                    continue
                dyst = abs(t_n - t_pz)
            else:
                if rh_n is None or abs(rh_n - rh_pz) > G.TOL_NAZWA_RH:
                    continue
                dyst = abs(t_n - t_pz) + abs(rh_n - rh_pz) / 10.0
            if najlepszy is None or dyst < najlepszy[0] - 1e-9:
                najlepszy = (dyst, t_pz, rh_pz)
        if najlepszy is None:
            continue
        _d, t_pz, rh_pz = najlepszy
        zd["nazwa"] = (f"{G._cell_to_sheet_name_part(t_pz)}, "
                       f"{G._cell_to_sheet_name_part(rh_pz) if rh_pz is not None else '-'}")
        G.ZAKLADKI_Z_PZ.add(zd["nazwa"])
        zmienione += 1
    print(f"  Nazwy zakladek wg punktow z PZ ({zmienione}/{len(dane_zakladek)}, powtorzenia "
          f"z ta sama nazwa): " + " | ".join(f"'{zd['nazwa']}'" for zd in dane_zakladek))
    return zmienione


# ─── Kilka protokolow ────────────────────────────────────────────────────────

def _klucz_przyrzadu(rek):
    return (pz_dane.normalizuj_serial(G._cell_to_str(rek.get("E")))
            or pz_dane.normalizuj_serial(G._cell_to_str(rek.get("F"))))


def scal_protokoly(odczyty):
    """
    odczyty: [(dane_s2, dane_zakladek, dane_ef, f24), ...] -> jedna lista zakladek
    i jeden przyrzad na nr fabryczny. Blok spoza protokolu przyrzadu = {} (nieaktywny).
    """
    rekordy, ef_all, f24_all, idx = [], [], [], {}
    zakl_all = []
    for dane_s2, zakl, dane_ef, f24 in odczyty:
        przed = len(zakl_all)
        zakl_all.extend(zakl)
        obecne = set()
        for j, rek in enumerate(dane_s2):
            k = _klucz_przyrzadu(rek) or f"#{len(rekordy)}"
            if k not in idx:
                idx[k] = len(rekordy)
                rekordy.append(rek)
                ef_all.append([{} for _ in range(przed)])
                f24_all.append(f24[j] if j < len(f24) else None)
            n = idx[k]
            ef_all[n].extend(dane_ef[j] if j < len(dane_ef) else [{} for _ in zakl])
            obecne.add(n)
        for n in range(len(rekordy)):
            if n not in obecne:
                ef_all[n].extend({} for _ in zakl)
    return rekordy, zakl_all, ef_all, f24_all


# ─── Wybor przyrzadow ────────────────────────────────────────────────────────

def wybrane_przyrzady(rekordy, wybor):
    """Filtr '2, 5, 21-24' / numery fabryczne lub ewidencyjne. Pusty wybor = wszystkie."""
    tokeny = [t.strip() for t in re.split(r"[;,]", wybor or "") if t.strip()]
    if not tokeny:
        return list(range(len(rekordy)))
    numery, klucze = set(), set()
    for t in tokeny:
        m = re.fullmatch(r"(\d{1,3})\s*-\s*(\d{1,3})", t)
        if m:
            numery.update(range(int(m.group(1)), int(m.group(2)) + 1))
        elif re.fullmatch(r"\d{1,3}", t):
            numery.add(int(t))
        else:
            klucze.add(pz_dane.normalizuj_serial(t))
        if re.fullmatch(r"\d{4,}", t):
            klucze.add(pz_dane.normalizuj_serial(t))
    return [j for j, rek in enumerate(rekordy)
            if rek.get("_nr_przyrzadu") in numery
            or pz_dane.normalizuj_serial(G._cell_to_str(rek.get("E"))) in klucze
            or pz_dane.normalizuj_serial(G._cell_to_str(rek.get("F"))) in klucze]


# ─── Uruchomienie ────────────────────────────────────────────────────────────

def _opis_przyrzadu(rek, ef_kopia, zakl):
    aktywne = [(zakl[i], ef) for i, ef in enumerate(ef_kopia) if G._czy_blok_ef_aktywny(ef)]
    szare = sum(1 for _z, ef in aktywne if ef.get("SZARY"))
    czujniki = []
    for _z, ef in aktywne:
        t = ef.get("CC04_TAG")
        if t and (not czujniki or czujniki[-1] != t):
            czujniki.append(t)
    nazwy = G._unikalne_nazwy_zakladek([z for z, _ef in aktywne])
    powt = sum(1 for n in nazwy if G._to_jest_powtorka_kolizji(n))
    return (f"{len(aktywne)} zakl. (powtorzen {powt}, z szarych komorek {szare})"
            + (f", czujnik: {' -> '.join(czujniki)}" if czujniki else ""))


def sciezka_protokolu(nazwa):
    """Protokol z folderu kroku 6; gdy go tam nie ma — z folderu roboczego (jak dawniej)."""
    if os.path.isabs(nazwa):
        return nazwa
    w_kroku = os.path.join(FOLDER_WYJ, nazwa)
    if os.path.exists(w_kroku):
        return w_kroku
    w_roboczym = os.path.join(FOLDER, nazwa)
    return w_roboczym if os.path.exists(w_roboczym) else w_kroku


def _main_impl():
    SEP = "=" * 65
    if not PROTOKOLY:
        print("[BLAD] Nie wybrano protokolu (krok 6, 'Protokoly do porownania').")
        return 1
    sciezki = [sciezka_protokolu(p) for p in PROTOKOLY]
    brak = [p for p in sciezki if not os.path.exists(p)]
    if brak:
        for p in brak:
            print(f"[BLAD] Brak protokolu: {p}")
        return 1
    if not os.path.exists(os.path.join(FOLDER, SZABLON)):
        print(f"[BLAD] Brak szablonu arkusza obliczeniowego: {os.path.join(FOLDER, SZABLON)}")
        return 1

    # Zachowanie kroku 3 dostosowane do arkusza porownawczego (patrz naglowek pliku).
    G.OBLICZAJ_WARUNKI_SRODOWISKOWE = False
    G.WYNIKI_WSZYSTKIE_ZAKLADKI = True
    G.ZAKLADKI_WSZYSTKICH_PUNKTOW_PZ = False
    G.USUWAJ_PUSTE_BLOKI_KOPII_S3 = True

    print(SEP)
    print(f"KROK 6  Arkusze porownawcze — odczyt {len(sciezki)} protokolu(ow) (bez Excela)")
    print(SEP)
    odczyty, multimetr = [], None
    for p in sciezki:
        s2, zakl, ef, f24, mm = wczytaj_protokol(p, uwzglednij_szare=UWZGLEDNIJ_SZARE)
        odczyty.append((s2, zakl, ef, f24))
        multimetr = multimetr or mm
    dane_s2, dane_zakladek, dane_ef, f24 = scal_protokoly(odczyty)
    if len(sciezki) > 1:
        print(f"  Po polaczeniu: {len(dane_s2)} przyrzadow, {len(dane_zakladek)} blokow.")

    nazwy_zakladek_porownania(dane_zakladek)
    G.uzupelnij_przetworniki(dane_s2, multimetr)

    wybrane = wybrane_przyrzady(dane_s2, WYBOR_PRZYRZADOW)
    if WYBOR_PRZYRZADOW.strip():
        print(f"  Wybor przyrzadow '{WYBOR_PRZYRZADOW}': {len(wybrane)} z {len(dane_s2)}.")
    wybrane = [j for j in wybrane if any(G._czy_blok_ef_aktywny(ef) for ef in dane_ef[j])]
    dane_s2 = [dane_s2[j] for j in wybrane]
    dane_ef = [dane_ef[j] for j in wybrane]
    f24 = [f24[j] for j in wybrane]
    dane_s2, dane_ef, f24, _pary = G._polacz_wyswietlacz_z_wyjsciem(dane_s2, dane_ef, f24)
    if not dane_s2:
        print("[BLAD] Zaden wybrany przyrzad nie ma odczytow — nie ma czego porownywac.")
        return 1

    print(f"  Przyrzady do arkuszy porownawczych ({len(dane_s2)}):")
    for rek, ef_k in zip(dane_s2, dane_ef):
        print(f"    • {rek['_nr_przyrzadu']:>2}. {G._numer_do_nazwy(rek) or '?'}: "
              f"{_opis_przyrzadu(rek, ef_k, dane_zakladek)}")

    os.makedirs(FOLDER_WYJ, exist_ok=True)
    print(SEP)
    print(f"  Arkusze -> {FOLDER_WYJ}")
    print(SEP)
    nazwy, *_ = G.utworz_kopie(FOLDER, SZABLON, dane_s2, dane_zakladek, dane_ef,
                               G.ARKUSZ_WYNIKI, f24, folder_wyj=FOLDER_WYJ,
                               odczyt_kalibracji=False)
    sciezki_kopii = [os.path.join(FOLDER_WYJ, G._bezpieczna_nazwa_pliku(n)) for n in nazwy]
    for p in sciezki_kopii:
        cc_widok.wymus_widok(p, G.TAB_RATIO, pierwszy=0)
    print(SEP)
    print(f"Zakonczono krok 6. Arkuszy porownawczych: {len(nazwy)} (folder: {FOLDER_WYJ}).")
    print(SEP)
    return 0


def main():
    _ar_user = G._czytaj_autorecover()
    if G.CZYSC_AUTORECOVER:
        try:
            G._wyczysc_autorecover_folder()
        except Exception as e:
            print(f"  [AutoRecover] Sprzatanie pominieto: {type(e).__name__}: {e}")
    if G.SPRZATAJ_ZAWIESZONE_EXCELE:
        cc_excel.sprzataj_zawieszone_excele()
    try:
        return _main_impl()
    finally:
        if _ar_user is not None:
            G._ustaw_autorecover(_ar_user)


if __name__ == "__main__":
    sys.exit(main())
