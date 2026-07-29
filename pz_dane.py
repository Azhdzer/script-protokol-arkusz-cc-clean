# -*- coding: utf-8 -*-
"""
pz_dane.py — wspolny modul do wczytywania danych przyrzadow z:

  1) plikow PDF "Potwierdzenie zamowienia" (folder PZ/),
  2) pliku "Zestawienie wzorcowanych przyrzadow.xlsx" (arkusz Termohigrometry).

Uzywany przez generuj_obserwacje.py (wypelnianie Strony 2 protokolu) oraz
generuj_arkusze.py (fallback, gdy Strona 2 pusta). Modul jest CZYSTY — tylko
czyta i parsuje, bez efektow ubocznych.

Kluczowe pojecia:
  • nr_fabryczny (serial) — laczy przyrzad z PZ z kolumna pomiarowa Strony 3
    (pliki wynikow nazywaja sie <SERIAL>_wynik.xlsx).
  • rozdzielczosc t/RH — z Zestawienia (po producencie+typie); gdy brak,
    liczona z wahania cyfr po przecinku w danych pomiarowych.
"""

import os
import re
import glob

try:
    import logging
    from pypdf import PdfReader
    logging.getLogger('pypdf').setLevel(logging.ERROR)
    _PDF_OK = True
except ImportError:
    _PDF_OK = False

try:
    import openpyxl
    _XLSX_OK = True
except ImportError:
    _XLSX_OK = False


# =============================================================================
# MODEL DANYCH
# =============================================================================

class PZPrzyrzad:
    """Jeden przyrzad z PZ (po rozwinieciu 'N szt.' — dokladnie jeden nr fabryczny)."""
    __slots__ = ("nr_zlecenia", "wytworca", "typ", "nr_fabr", "nr_ewid",
                 "czuj_wytworca", "czuj_typ", "czuj_nr_fabr", "uzytkownik")

    def __init__(self, nr_zlecenia="", wytworca="", typ="", nr_fabr="", nr_ewid="",
                 czuj_wytworca="", czuj_typ="", czuj_nr_fabr="", uzytkownik=""):
        self.nr_zlecenia = nr_zlecenia
        self.wytworca = wytworca
        self.typ = typ
        self.nr_fabr = nr_fabr
        self.nr_ewid = nr_ewid
        self.czuj_wytworca = czuj_wytworca
        self.czuj_typ = czuj_typ
        self.czuj_nr_fabr = czuj_nr_fabr
        self.uzytkownik = uzytkownik   # blok adresowy z pola 'UZYTKOWNIK:' (moze byc pusty)

    def __repr__(self):
        return (f"PZPrzyrzad(zlec={self.nr_zlecenia}, wytw={self.wytworca!r}, "
                f"typ={self.typ!r}, fabr={self.nr_fabr!r}, ewid={self.nr_ewid!r}, "
                f"czuj={self.czuj_typ!r}/{self.czuj_nr_fabr!r})")


# =============================================================================
# NORMALIZACJA
# =============================================================================

def normalizuj_serial(s):
    """Klucz do dopasowania: bez spacji, wielkie litery, bez kropek/przecinkow na koncu."""
    if s is None:
        return ""
    return re.sub(r'\s+', '', str(s)).strip().strip('.,;').upper()


def _norm_txt(s):
    """Normalizacja nazw (producent/typ) do dopasowania rozmytego."""
    if s is None:
        return ""
    return re.sub(r'[\s\-_.]+', '', str(s)).strip().lower()


def _oczysc(s):
    """Porzadkuje wartosc do wpisania: zbedne spacje, koncowe znaki interpunkcyjne."""
    if s is None:
        return ""
    return re.sub(r'\s+', ' ', str(s)).strip().strip('.,;').strip()


# =============================================================================
# PARSER PDF (Potwierdzenie zamowienia)
# =============================================================================

# Etykiety pol wewnatrz opisu przyrzadu — DWUJEZYCZNE (PL / EN; PZ bywa po angielsku).
# Granica konca wartosci serialu = poczatek nastepnej etykiety.
_END = (r'wytw[oó]rca|manufacturer|nr\s*wewn|identification|'
        r'typ\s*:|type\s*:|nr\s*kat|year\s+of\s+production')
_RE_TYP    = re.compile(r'(?:typ|type|nr\s*kat\.?)\s*:\s*([^,\n]+)', re.I)
_RE_WEWN   = re.compile(r'(?:nr\s*wewn\.?|identification\s+number)\s*:\s*([^,\n]+)', re.I)
_RE_WYTW   = re.compile(r'(?:wytw[oó]rca|manufacturer)\s*:\s*([^,.\n]+)', re.I)
# serial: "nr fabr.:" / "serial number(s):" albo bare "nr:" (nie "nr kat.", nie "nr wewn.")
_RE_FABR   = re.compile(
    r'(?:nr\s*fabr\.?|serial\s*numbers?)\s*:\s*(.+?)(?=,?\s*(?:' + _END + r')|$)', re.I)
_RE_NR     = re.compile(
    r'(?<!kat)(?<!wewn)\bnr\s*:\s*(.+?)(?=,?\s*(?:' + _END + r')|$)', re.I)
# podzial obiekt / czujnik pomiarowy (PL/EN, rozne sformulowania)
_RE_CZUJNIK_SPLIT = re.compile(
    r'\b(?:z\s+czujnikiem|oraz\s+czujnika|i\s+(?:zewn[eę]trznego\s+)?czujnik\w*'
    r'|with\s+\d+\b[^,]*?sensors?)\b', re.I)


def _wytnij_serial_liste(part):
    """Zwraca liste serial-i z czesci 'obiekt' (obsluguje kilka po przecinku)."""
    m = _RE_FABR.search(part) or _RE_NR.search(part)
    if not m:
        return []
    return [s.strip() for s in m.group(1).split(',') if s.strip()]


def _parsuj_pole(part):
    """Zwraca (typ, wytworca) z fragmentu tekstu."""
    mt = _RE_TYP.search(part)
    mw = _RE_WYTW.search(part)
    return (_oczysc(mt.group(1)) if mt else "",
            _oczysc(mw.group(1)) if mw else "")


def _parsuj_wpis(wpis, nr_zlecenia):
    """
    Parsuje jeden wpis 'Obiekty wzorcowania' -> lista PZPrzyrzad.
    Obsluguje: obiekt + opcjonalny czujnik ('z czujnikiem'/'oraz czujnika'),
    oraz kilka seriali (np. '2 szt.') -> po jednym rekordzie na serial.
    """
    czesci = _RE_CZUJNIK_SPLIT.split(wpis, maxsplit=1)
    obiekt = czesci[0]
    czujnik = czesci[1] if len(czesci) > 1 else ""

    typ, wytworca = _parsuj_pole(obiekt)
    mw = _RE_WEWN.search(obiekt)
    nr_ewid = _oczysc(mw.group(1)) if mw else ""
    serials = _wytnij_serial_liste(obiekt)

    czuj_typ, czuj_wytworca = _parsuj_pole(czujnik) if czujnik else ("", "")
    czuj_serials = _wytnij_serial_liste(czujnik) if czujnik else []
    czuj_serial = _oczysc(czuj_serials[0]) if czuj_serials else ""

    # Gdy wytworca podany raz na koncu (wspolny dla obiektu i czujnika) — trafia do
    # czesci 'czujnik'; obiekt zostaje bez wytworcy. Dziedziczymy go wtedy do obiektu.
    if not wytworca and czuj_wytworca:
        wytworca = czuj_wytworca

    out = []
    for s in serials or [""]:
        out.append(PZPrzyrzad(
            nr_zlecenia=nr_zlecenia,
            wytworca=wytworca, typ=typ,
            nr_fabr=_oczysc(s), nr_ewid=nr_ewid,
            czuj_wytworca=czuj_wytworca, czuj_typ=czuj_typ, czuj_nr_fabr=czuj_serial,
        ))
    return out


# Pole UZYTKOWNIK / USER (opcjonalne) — blok adresowy do '[użytkownik]' w Word.
# Naglowek zawsze wielkimi literami (jak APPLICANT/ZLECENIODAWCA) — bez re.I, by nie
# lapac slowa 'user' w zdaniach.
_RE_UZYT = re.compile(
    r'(?:UŻYTKOWNIK|UZYTKOWNIK|USER)\s*:(.*?)'
    r'(?:Uwaga\s*:|Note\s*:|Opis\s+us[łl]ugi|Service\s+description|$)', re.S)


def _parsuj_uzytkownik(text):
    """Zwraca blok adresowy uzytkownika (bez pustych linii i notatki 'Uwaga'), albo ''."""
    m = _RE_UZYT.search(text)
    if not m:
        return ""
    linie = [ln.strip() for ln in m.group(1).splitlines() if ln.strip()]
    return "\n".join(linie)


def _numer_zlecenia_th(text):
    """Zwraca numer zlecenia (samo '119'); przy wielu zleceniach preferuje /LA/TH/."""
    m = re.search(r'(?:Numer\s+zlecenia\s+laboratorium|Order\s+numbers?)\s*:\s*([^\n]+)',
                  text, re.I)
    if not m:
        return ""
    linia = m.group(1)
    th = re.search(r'(\d+)\s*/\s*LA\s*/\s*TH', linia, re.I)
    if th:
        return th.group(1)
    any_nr = re.search(r'(\d+)\s*/', linia)
    return any_nr.group(1) if any_nr else _oczysc(linia)


def parsuj_pdf(path):
    """Parsuje jeden PDF PZ -> lista PZPrzyrzad."""
    if not _PDF_OK:
        return []
    reader = PdfReader(path)
    text = "\n".join((p.extract_text() or "") for p in reader.pages)
    nr_zlec = _numer_zlecenia_th(text)
    uzyt = _parsuj_uzytkownik(text)

    m = re.search(
        r'(?:Obiekty\s+wzorcowania|Calibration\s+objects)\s*:(.*?)'
        r'(?:(?:Metoda\s+wzorcowania|Calibration\s+methods?)\s*:|$)', text, re.I | re.S)
    if not m:
        return []
    sekcja = re.sub(r'\s+', ' ', m.group(1)).strip()

    # podzial na wpisy numerowane '1) ... 2) ...'; brak numeracji => jeden wpis
    wpisy = re.split(r'(?<!\d)([1-9])\)\s', sekcja)
    przyrzady = []
    if len(wpisy) > 1:
        # re.split z grupa: [prefix, '1', tekst1, '2', tekst2, ...]
        for i in range(1, len(wpisy), 2):
            tekst = wpisy[i + 1] if i + 1 < len(wpisy) else ""
            przyrzady.extend(_parsuj_wpis(tekst, nr_zlec))
    else:
        przyrzady.extend(_parsuj_wpis(sekcja, nr_zlec))
    for p in przyrzady:
        p.uzytkownik = uzyt   # to samo dla wszystkich przyrzadow w danym PZ
    return przyrzady


def wczytaj_pz(folder_pz):
    """
    Wczytuje wszystkie PDFy z folder_pz -> dict[normalizuj_serial(nr_fabr) -> PZPrzyrzad].
    Zwraca (mapa, lista_wszystkich). Cichy gdy brak folderu/plikow.
    """
    mapa, lista = {}, []
    if not _PDF_OK:
        print("  [PZ] Brak biblioteki 'pypdf' — pomijam dane z PZ.")
        return mapa, lista
    if not os.path.isdir(folder_pz):
        print(f"  [PZ] Brak folderu '{folder_pz}' — pomijam dane z PZ.")
        return mapa, lista

    pliki = sorted(glob.glob(os.path.join(folder_pz, "*.pdf")))
    for path in pliki:
        try:
            przyrzady = parsuj_pdf(path)
        except Exception as e:
            print(f"  [PZ] Blad odczytu '{os.path.basename(path)}': {e}")
            continue
        for p in przyrzady:
            lista.append(p)
            # Klucz po nr fabrycznym I nr ewidencyjnym — pliki logerow bywaja nazwane
            # jednym albo drugim (np. Vaisala: plik 'EV363499' = nr ewidencyjny, a nr
            # fabryczny to 'D4940027'). Pierwszy przyrzad z danym kluczem wygrywa.
            for klucz in (normalizuj_serial(p.nr_fabr), normalizuj_serial(p.nr_ewid)):
                if klucz:
                    mapa.setdefault(klucz, p)
    print(f"  [PZ] Wczytano {len(pliki)} PDF, {len(lista)} przyrzadow "
          f"({len(mapa)} z nr fabrycznym).")
    return mapa, lista


# =============================================================================
# ZESTAWIENIE (rozdzielczosc t / RH)
# =============================================================================

_RE_RES_T  = re.compile(r't\s*:?\s*([\d.,]+)', re.I)
_RE_RES_RH = re.compile(r'(?:rh|wilg)\s*:?\s*([\d.,]+)', re.I)


def _do_float(s):
    try:
        return float(str(s).replace(',', '.'))
    except (ValueError, TypeError):
        return None


def wczytaj_zestawienie(sciezka, arkusz="Termohigrometry"):
    """
    Wczytuje liste wpisow: [(producent_norm, typ_norm, t_res, rh_res), ...].
    Producent w kol. B, Typ w kol. D (czesto wieloliniowy, np. 'wskaznik:\\ntesto 625\\n...'),
    Rozdzielczosc w kol. J (np. 't: 0,1 °C\\nRH: 0,1 %'). Cichy gdy brak pliku.
    Zwraca liste (nie dict) — dopasowanie jest rozmyte (podciag), patrz rozdzielczosc_zestawienie.
    """
    wpisy = []
    if not _XLSX_OK or not os.path.exists(sciezka):
        if not os.path.exists(sciezka):
            print(f"  [PZ] Brak pliku Zestawienia '{os.path.basename(sciezka)}' — "
                  f"rozdzielczosc tylko z danych.")
        return wpisy
    try:
        wb = openpyxl.load_workbook(sciezka, data_only=True, read_only=True)
    except Exception as e:
        print(f"  [PZ] Blad otwarcia Zestawienia: {e}")
        return wpisy
    if arkusz not in wb.sheetnames:
        print(f"  [PZ] Brak arkusza '{arkusz}' w Zestawieniu.")
        return wpisy
    ws = wb[arkusz]
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=11, values_only=True):
        producent, typ, rozdz = row[1], row[3], row[9]   # B, D, J
        if not (producent and typ and rozdz):
            continue
        rozdz = str(rozdz)
        mt = _RE_RES_T.search(rozdz)
        mrh = _RE_RES_RH.search(rozdz)
        t_res = _do_float(mt.group(1)) if mt else None
        rh_res = _do_float(mrh.group(1)) if mrh else None
        wpisy.append((_norm_txt(producent), _norm_txt(typ), t_res, rh_res))
    wb.close()
    print(f"  [PZ] Zestawienie: {len(wpisy)} przyrzadow z rozdzielczoscia.")
    return wpisy


def rozdzielczosc_zestawienie(zest, producent, typ):
    """
    Rozmyte dopasowanie (producent, typ) do Zestawienia. Kolumna Typ bywa wieloliniowa
    ('wskaznik:\\ntesto 625\\nczujnik:\\n...'), dlatego typ przyrzadu szukamy jako PODCIAG.
    Zwraca (t_res, rh_res) albo (None, None).
    """
    pn, tn = _norm_txt(producent), _norm_txt(typ)
    if not tn:
        return (None, None)
    for zp, zt, t_res, rh_res in zest:
        prod_ok = bool(pn) and (pn == zp or pn in zp or zp in pn)
        typ_ok = tn in zt or zt in tn
        if prod_ok and typ_ok:
            return (t_res, rh_res)
    return (None, None)


# =============================================================================
# ROZDZIELCZOSC Z DANYCH (fallback)
# =============================================================================

def rozdzielczosc_z_kolumny(values):
    """
    Rozdzielczosc = najmniejsze miejsce dziesietne, ktore sie zmienia w kolumnie:
      same calkowite (40.00...) -> 1;  zmiana na 1. miejscu (40.1) -> 0.1;
      na 2. (40.04) -> 0.01.  None gdy brak danych liczbowych.
    """
    vals = [v for v in values if isinstance(v, (int, float))]
    if not vals:
        return None
    max_dec = 0
    for v in vals:
        s = f"{float(v):.4f}".rstrip('0')
        dec = len(s.split('.')[1]) if '.' in s else 0
        if dec > max_dec:
            max_dec = dec
    return {0: 1.0, 1: 0.1, 2: 0.01, 3: 0.001}.get(max_dec, 10 ** (-max_dec))
