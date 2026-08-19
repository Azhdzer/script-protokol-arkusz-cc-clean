# -*- coding: utf-8 -*-
"""
cc_config.py — jedno zrodlo prawdy o ustawieniach calego obiegu.

Modul pelni dwie role:

  1) Dla PANELU (app_gui.py) — REJESTR ustawien: kazda opcja ma typ, wartosc
     domyslna, etykiete PL, opis, krok obiegu i poziom (podstawowy/zaawansowany).
     Panel buduje z tego formularze automatycznie, zapisuje wartosci do
     'cc_ustawienia.json' i eksportuje je do zmiennych srodowiskowych workera.

  2) Dla SKRYPTOW (generuj_*.py, analizuj_excele.py) — zestaw funkcji
     odczytujacych te zmienne: flaga(), liczba(), calk(), tekst(), minuty(), ...
     Skrypty zachowuja swoje stale jako wartosci domyslne — gdy panel nic nie
     poda, dzialaja dokladnie jak wczesniej (mozna je nadal uruchamiac recznie).

WAZNE: modul korzysta wylacznie z biblioteki standardowej. Importuja go workery,
wiec nie moze wciagac PySide6 ani innych ciezkich zaleznosci.
"""

from __future__ import annotations

import datetime
import json
import os
import sys

# ─────────────────────────────────────────────────────────────────────────────
# Sciezki (dziala z zrodla i po zamrozeniu PyInstaller)
# ─────────────────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    HERE = os.path.dirname(sys.executable)
else:
    HERE = os.path.dirname(os.path.abspath(__file__))

PLIK_USTAWIEN = os.path.join(HERE, "cc_ustawienia.json")

# Kroki obiegu — kolejnosc ma znaczenie (panel rysuje je w tej kolejnosci).
KROKI = [
    ("przygotowanie", "Przygotowanie",       "Sprawdz pliki wejsciowe przed uruchomieniem"),
    ("analiza",       "1 · Analiza logow",   "Logi z przyrzadow  ->  wyniki/<serial>_wynik.xlsx"),
    ("obs",           "2 · Obserwacja",      "TXT multimetru  ->  arkusz obserwacji + protokol"),
    ("ark",           "3 · Arkusze i Word",  "Protokol  ->  kopie arkuszy Excel + swiadectwa Word"),
]

PODSTAWOWY  = "podstawowy"
ZAAWANSOWANY = "zaawansowany"


# ─────────────────────────────────────────────────────────────────────────────
# Definicja pojedynczego ustawienia
# ─────────────────────────────────────────────────────────────────────────────
class Ust:
    """
    Jedno ustawienie widoczne w panelu i przekazywane do skryptu przez env.

    typ:
        "tekst"        — pole tekstowe
        "calk"         — liczba calkowita
        "liczba"       — liczba zmiennoprzecinkowa
        "flaga"        — przelacznik wl/wyl
        "minuty"       — liczba calkowita w minutach (skrypt robi z niej timedelta)
        "plik"         — plik z folderu roboczego (lista rozwijana; patrz `wzorzec`)
        "folder"       — sciezka do folderu (pole + przycisk "Wybierz")
        "pliki"        — wiele plikow z folderu (lista z zaznaczaniem)
        "kolor"        — kolor #RRGGBB
        "tabela"       — tabela (JSON w env); uzywane dla mapowania CC-04
    """

    __slots__ = ("env", "etykieta", "typ", "domyslna", "krok", "poziom",
                 "grupa", "opis", "wzorzec", "minimum", "maksimum", "krok_wart",
                 "przyrostek", "kolumny", "podpowiedz")

    def __init__(self, env, etykieta, typ, domyslna, krok, poziom=PODSTAWOWY,
                 grupa="Ogolne", opis="", wzorzec=None, minimum=None,
                 maksimum=None, krok_wart=None, przyrostek="", kolumny=None,
                 podpowiedz=None):
        self.env = env
        self.etykieta = etykieta
        self.typ = typ
        self.domyslna = domyslna
        self.krok = krok
        self.poziom = poziom
        self.grupa = grupa
        self.opis = opis
        self.wzorzec = wzorzec          # np. (".xlsx", ".xlsm") dla typu "plik"
        self.minimum = minimum
        self.maksimum = maksimum
        self.krok_wart = krok_wart
        self.przyrostek = przyrostek
        self.kolumny = kolumny or []    # naglowki dla typu "tabela"
        # Fragment nazwy pliku podpowiadany przy PIERWSZYM uruchomieniu, gdy nic
        # jeszcze nie wybrano — bez tego lista rozwijana wskazalaby pierwszy plik
        # alfabetycznie, czyli zwykle nie ten, o ktory chodzi.
        self.podpowiedz = podpowiedz

    def do_env(self, wartosc):
        """Zamienia wartosc na tekst przekazywany do podprocesu."""
        if self.typ == "flaga":
            return "1" if wartosc else "0"
        if self.typ == "pliki":
            return ";".join(wartosc or [])
        if self.typ == "tabela":
            return json.dumps(wartosc, ensure_ascii=False)
        return "" if wartosc is None else str(wartosc)


def _U(*a, **kw):
    return Ust(*a, **kw)


# ─────────────────────────────────────────────────────────────────────────────
# REJESTR USTAWIEN
# ─────────────────────────────────────────────────────────────────────────────
# Wspolne dla wszystkich krokow
_WSPOLNE = [
    _U("CC_FOLDER", "Folder roboczy", "folder", HERE, "przygotowanie",
       grupa="Sciezki",
       opis="Folder z protokolami, szablonami i plikami TXT. Wszystkie sciezki "
            "wzgledne liczone sa od niego."),
]

# ── KROK 1 — analizuj_excele.py ──────────────────────────────────────────────
_ANALIZA = [
    _U("ANL_PLIKI", "Pliki do przetworzenia", "pliki", [], "analiza",
       grupa="Dane wejsciowe",
       wzorzec=(".csv", ".xls", ".xlsx", ".txt", ".pdf", ".log"),
       opis="Nic nie zaznaczone = przetwarzane sa WSZYSTKIE pliki z folderu. "
            "Zaznacz wybrane, gdy w folderze leza tez logi z innego zlecenia."),
    _U("ANL_INPUT", "Folder z logami przyrzadow", "folder", "excel_do_analizy", "analiza",
       grupa="Sciezki",
       opis="Tu wgrywasz surowe pliki z loggerow DUT (csv / xlsx / txt / pdf / log)."),
    _U("ANL_OUTPUT", "Folder wynikow", "folder", "wyniki", "analiza",
       poziom=ZAAWANSOWANY, grupa="Sciezki",
       opis="Tu trafiaja znormalizowane pliki <serial>_wynik.xlsx czytane pozniej "
            "przez krok Obserwacja."),
    _U("ANL_DEBUG", "Tryb diagnostyczny (pelne bledy)", "flaga", False, "analiza",
       poziom=ZAAWANSOWANY, grupa="Diagnostyka",
       opis="Wypisuje pelny traceback przy bledzie parsowania pliku."),
]

# ── KROK 2 — generuj_obserwacje.py ───────────────────────────────────────────
_OBSERWACJA = [
    # --- podstawowe ---
    _U("OBS_TXT_FILES", "Pliki TXT multimetru", "pliki", [], "obs",
       grupa="Dane wejsciowe", wzorzec=(".txt",),
       opis="Zaznacz plik(i) pomiaru. Kilka plikow = przerwany pomiar — zostana "
            "sklejone chronologicznie z usunieciem duplikatow czasu."),
    # Zdjecia stoja tuz pod wyborem plikow TXT, bo to decyzja podejmowana raz na
    # zlecenie — razem z tym, co w ogole bierzemy do pomiaru. Schowane nizej
    # (za podpisami, szablonami i filtrami) trzeba bylo ich szukac scrollem.
    _U("OBS_FOTO", "Zbieraj zdjecia punktow pomiarowych", "flaga", False, "obs",
       grupa="Zdjecia",
       opis="Dla kazdego wybranego punktu kopiuje zdjecia z czasow jego 5 wierszy "
            "reprezentacyjnych (tych podswietlonych w obserwacji). Czas zdjecia "
            "czytany jest z nazwy pliku, a gdy jej nie ma — z daty modyfikacji."),
    _U("OBS_FOTO_ZRODLO", "Folder zrodlowy ze zdjeciami", "folder",
       r"\\83b\Zdjęcia", "obs", grupa="Zdjecia",
       opis="Stad kopiujemy. Zwykle folder z aparatu / dysku sieciowego "
            "z danego dnia wzorcowania."),
    _U("OBS_POMIJAJ_PUSTE_KOL", "Pomijaj puste kolumny w obserwacji", "flaga", True, "obs",
       grupa="Arkusz obserwacji",
       opis="Kolumna bez ani jednej wartosci (nieuzyty kanal multimetru) nie "
            "trafia do arkusza — tabela jest ciagla, bez dziur. Analizy i "
            "protokolu to nie dotyczy: pracuja na danych z pliku TXT, nie na "
            "kolumnach arkusza. Odwolania wykresow sa przeliczane automatycznie."),

    _U("OBS_PODPIS", "Pomiary wykonal(a)", "tekst", "Artsiom Azhdzer", "obs",
       grupa="Podpisy", opis="Trafia do arkusza obserwacji i na Strone 2 protokolu."),
    _U("OBS_PODPIS_SPR", "Protokol sprawdzil(a)", "tekst", "Marek Szpakowski", "obs",
       grupa="Podpisy", opis="Strona 2 protokolu."),
    _U("OBS_TEMPLATE", "Szablon obserwacji CC", "plik",
       "xxx_LA_TH_2026 - obserwacje CC.xlsx", "obs",
       grupa="Szablony", wzorzec=(".xlsx", ".xlsm"),
       opis="Szablon arkusza obserwacji dla komory CC."),
    _U("OBS_CC04_TEMPLATE", "Szablon obserwacji CC-04", "plik",
       "szablon_LA_TH_2026 - obserwacje.xlsx", "obs",
       grupa="Szablony", wzorzec=(".xlsx", ".xlsm"),
       opis="Szablon arkusza obserwacji dla komory CC-04."),
    _U("OBS_PROT_CC", "Szablon protokolu CC", "plik",
       "xxx_LA_TH_2026 - protokół CC.xlsx", "obs",
       grupa="Szablony", wzorzec=(".xlsx", ".xlsm")),
    _U("OBS_PROT_CC04", "Szablon protokolu CC-04", "plik",
       "xxx_LA_TH_2026 - protokół CC-04.xlsx", "obs",
       grupa="Szablony", wzorzec=(".xlsx", ".xlsm")),
    _U("CC_PZ_FOLDER", "Folder PZ (Potwierdzenia zamowienia)", "folder", "PZ", "obs",
       grupa="Dane przyrzadow",
       opis="PDF-y PZ (PL/EN) — z nich wypelniana jest tabela przyrzadow na Stronie 2."),
    _U("CC_ZESTAWIENIE", "Zestawienie wzorcowanych przyrzadow", "plik",
       "Zestawienie wzorcowanych przyrządów.xlsx", "obs",
       grupa="Dane przyrzadow", wzorzec=(".xlsx", ".xlsm"),
       opis="Zrodlo rozdzielczosci t/RH (kolumny K/L Strony 2)."),
    _U("OBS_FILTR", "Filtr: nastawa vs odczyt komory", "flaga", True, "obs",
       grupa="Filtr nastawa/odczyt",
       opis="Odrzuca segmenty, w ktorych komora nie osiagnela nastawy "
            "(przejscia, suszenie) — nie trafiaja do protokolu."),
    _U("OBS_PROG", "Prog wzgledny nastawa/odczyt", "liczba", 10.0, "obs",
       grupa="Filtr nastawa/odczyt", przyrostek=" %", minimum=0.0, maksimum=100.0,
       opis="Dozwolona roznica |nastawa-odczyt| / nastawa * 100."),
    _U("OBS_TOL", "Tolerancja dopasowania czasu", "liczba", 3.0, "obs",
       grupa="Dopasowanie wynikow", przyrostek=" min", minimum=0.1, maksimum=240.0,
       krok_wart=0.5,
       opis="PROG ODRZUCENIA, nie okno wyszukiwania: skrypt zawsze bierze "
            "NAJBLIZSZY rekord z wyniki/, a ta wartosc decyduje, czy plik w ogole "
            "pasuje do tego wzorcowania. Loggery ustawiane sa na zapis co 1 min, "
            "wiec normalne odchylki to sekundy; 3 min daja zapas na przerwy w "
            "zapisie (Aranet/Efento potrafia pominac 2-3 probki). Za ciasna "
            "wartosc ODRZUCA caly plik. Ulamki dozwolone — 0,5 = 30 s."),

    # --- zaawansowane: okno analizy ---
    _U("OBS_STAB_MIN", "Rozgrzewka (gdy odczyty nie wejda w widelki)", "minuty", 120, "obs",
       poziom=ZAAWANSOWANY, grupa="Okno analizy", minimum=0, maksimum=1440,
       opis="Gdy odczyty komory nigdy nie trafia w widelki nastawy, okno analizy "
            "liczone jest od tego czasu od poczatku punktu."),
    _U("OBS_STAB_PO_RH", "Stabilizacja po wejsciu w widelki", "minuty", 120, "obs",
       poziom=ZAAWANSOWANY, grupa="Okno analizy", minimum=0, maksimum=1440,
       opis="Ile czasu odliczyc od momentu, gdy odczyty komory weszly w widelki "
            "wokol nastaw."),
    _U("OBS_PROG_T", "Widelki wejscia — temperatura", "liczba", 0.4, "obs",
       poziom=ZAAWANSOWANY, grupa="Okno analizy", przyrostek=" °C",
       minimum=0.0, maksimum=20.0, krok_wart=0.1,
       opis="|T odczytana - T zadana| ponizej tej wartosci = komora weszla w nastawe."),
    _U("OBS_PROG_RH", "Widelki wejscia — wilgotnosc", "liczba", 3.0, "obs",
       poziom=ZAAWANSOWANY, grupa="Okno analizy", przyrostek=" %",
       minimum=0.0, maksimum=100.0,
       opis="Wzgledna roznica RH odczytanej od zadanej."),
    _U("OBS_ODSTEP_KONIEC", "Odstep reprezentantow od zmiany nastawy", "minuty", 2, "obs",
       grupa="Okno analizy", minimum=0, maksimum=60,
       opis="5 wierszy reprezentacyjnych nie moze konczyc sie tuz przed zmiana "
            "nastawy. Na takim styku komora zaczyna juz przechodzic do kolejnego "
            "punktu, a 15-minutowe rozrzuty lapia probki zza granicy — odczyty "
            "wychodza rozmazane. 0 = bez odstepu (zachowanie sprzed poprawki)."),
    _U("OBS_MIN_OKNO", "Gwarantowany ogon pomiarowy", "minuty", 15, "obs",
       poziom=ZAAWANSOWANY, grupa="Okno analizy", minimum=1, maksimum=600,
       opis="Gdy punkt trzymany jest krotko, start okna jest cofany tak, by na "
            "koncu punktu zostalo co najmniej tyle minut."),

    # --- zaawansowane: suszenie ---
    _U("OBS_SUSZ_T_MIN", "Suszenie — T od", "liczba", 21.0, "obs",
       poziom=ZAAWANSOWANY, grupa="Wykrywanie suszenia", przyrostek=" °C",
       minimum=-90.0, maksimum=200.0,
       opis="Zakres temperatury charakterystyczny dla suszenia komory."),
    _U("OBS_SUSZ_T_MAX", "Suszenie — T do", "liczba", 27.0, "obs",
       poziom=ZAAWANSOWANY, grupa="Wykrywanie suszenia", przyrostek=" °C",
       minimum=-90.0, maksimum=200.0),
    _U("OBS_SUSZ_RH_MAX", "Suszenie — RH ponizej", "liczba", 50.0, "obs",
       poziom=ZAAWANSOWANY, grupa="Wykrywanie suszenia", przyrostek=" %",
       minimum=0.0, maksimum=100.0),

    # --- zaawansowane: punkty z PZ ---
    _U("OBS_PZ_PUNKTY", "Wybieraj punkty wg 'Zakres wzorcowania' z PZ", "flaga", True, "obs",
       poziom=ZAAWANSOWANY, grupa="Wybor punktow z PZ",
       opis="Jeden wsad komory obsluguje czesto kilka zlecen. Gdy wlaczone, do "
            "protokolu trafiaja wylacznie punkty zamowione w PZ."),
    _U("OBS_TOL_PUNKT_T", "Tolerancja punktu — temperatura", "liczba", 1.5, "obs",
       poziom=ZAAWANSOWANY, grupa="Wybor punktow z PZ", przyrostek=" °C",
       minimum=0.0, maksimum=20.0, krok_wart=0.1,
       opis="Dopuszczalna roznica nastawy komory od punktu z PZ."),
    _U("OBS_TOL_PUNKT_RH", "Tolerancja punktu — wilgotnosc", "liczba", 4.0, "obs",
       poziom=ZAAWANSOWANY, grupa="Wybor punktow z PZ", przyrostek=" %",
       minimum=0.0, maksimum=100.0),

    # --- zaawansowane: dopasowanie wynikow ---
    _U("OBS_MAX_ROZN_PRZYRZAD", "Max roznica odczyt przyrzadu vs nastawa", "liczba", 5.0, "obs",
       poziom=ZAAWANSOWANY, grupa="Dopasowanie wynikow", przyrostek=" °C",
       minimum=0.0, maksimum=50.0, krok_wart=0.5,
       opis="Zabezpieczenie przed przypisaniem pliku wynikow z DRUGIEJ komory: "
            "przyrzad musi pokazywac mniej wiecej nastawe naszej komory."),
    _U("OBS_KOREKTA_ZEGARA", "Korekta zegara loggera", "flaga", True, "obs",
       poziom=ZAAWANSOWANY, grupa="Dopasowanie wynikow",
       opis="Tanie loggery bywaja rozjechane w czasie o godziny. Skrypt porownuje "
            "profil temperatury loggera z profilem komory i koryguje przesuniecie "
            "(same wartosci pomiarow zostaja nietkniete)."),
    _U("OBS_KZ_MAX", "Korekta zegara — max przesuniecie", "calk", 360, "obs",
       poziom=ZAAWANSOWANY, grupa="Dopasowanie wynikow", przyrostek=" min",
       minimum=0, maksimum=1440),
    _U("OBS_KZ_KROK", "Korekta zegara — rozdzielczosc", "calk", 5, "obs",
       poziom=ZAAWANSOWANY, grupa="Dopasowanie wynikow", przyrostek=" min",
       minimum=1, maksimum=60),

    # --- zaawansowane: filtr bezwzgledny ---
    _U("OBS_TOL_ABS_T", "Tolerancja bezwzgledna — temperatura", "liczba", 1.0, "obs",
       poziom=ZAAWANSOWANY, grupa="Filtr nastawa/odczyt", przyrostek=" °C",
       minimum=0.0, maksimum=20.0, krok_wart=0.1,
       opis="Odczyt w tych granicach od nastawy jest zgodny, nawet gdy prog "
            "wzgledny go odrzuca (wazne przy malych nastawach)."),
    _U("OBS_TOL_ABS_RH", "Tolerancja bezwzgledna — wilgotnosc", "liczba", 2.0, "obs",
       poziom=ZAAWANSOWANY, grupa="Filtr nastawa/odczyt", przyrostek=" %RH",
       minimum=0.0, maksimum=100.0, krok_wart=0.5),

    _U("OBS_FOTO_CEL", "Folder docelowy zdjec", "folder", "foto", "obs",
       poziom=ZAAWANSOWANY, grupa="Zdjecia",
       opis="Tu trafiaja skopiowane zdjecia, w podfolderach 'punkt_NN_<nastawa>'. "
            "Sciezka wzgledna liczona od folderu roboczego."),
    _U("OBS_FOTO_TOL", "Tolerancja czasu zdjecia", "minuty", 1, "obs",
       poziom=ZAAWANSOWANY, grupa="Zdjecia", minimum=1, maksimum=120,
       opis="Max odchylka czasu zdjecia od wiersza pomiarowego."),
]

# ── KROK 3 — generuj_arkusze.py ──────────────────────────────────────────────
_ARKUSZE = [
    # --- podstawowe ---
    _U("CC_PROTOKOL", "Plik protokolu", "plik", "", "ark",
       grupa="Dane wejsciowe", wzorzec=(".xlsx", ".xlsm"), podpowiedz="protok",
       opis="Wypelniony protokol CC / CC-04 (wynik kroku 2). Strona 2 = lista kopii "
            "do zrobienia, Strona 3 = punkty pomiarowe."),
    _U("CC_SZABLON", "Szablon arkusza obliczeniowego", "plik", "", "ark",
       grupa="Dane wejsciowe", wzorzec=(".xlsx", ".xlsm"), podpowiedz="ark. obl",
       opis="Wzor ark. obl. — z niego powstaje kopia dla kazdego przyrzadu."),
    _U("GEN_EXCEL", "Generuj arkusze Excel", "flaga", True, "ark",
       grupa="Etapy",
       opis="Wylaczone = korzysta z juz istniejacych kopii (przydatne, gdy chcesz "
            "wygenerowac sam Word)."),
    _U("GEN_WORD", "Generuj swiadectwa Word", "flaga", True, "ark",
       grupa="Etapy", opis="Etap 7 — dokumenty .docx."),
    _U("GEN_PUSTE", "Usuwaj puste bloki Strony 3", "flaga", True, "ark",
       grupa="Etapy",
       opis="Kopia dostaje tylko te zakladki, dla ktorych sa dane E/F."),
    _U("GEN_POMIJAJ_PUSTE", "Pomijaj przyrzady bez pomiarow", "flaga", True, "ark",
       grupa="Etapy",
       opis="Przyrzad wyszarzony w calosci na Stronie 3 nie dostaje ani kopii "
            "Excel, ani swiadectwa Word. Tak wybierasz pojedynczy przyrzad: "
            "wyszarz pomiary pozostalych."),
    _U("GEN_NR_SW", "Numer pierwszego swiadectwa", "calk", 1047, "ark",
       grupa="Naglowki", minimum=1, maksimum=99999,
       opis="Kolejne kopie dostaja numery rosnaco (1047, 1048, ...)."),
    _U("GEN_K18_CC", "Higrometr K18 — komora CC", "tekst", "S8000-02", "ark",
       grupa="Naglowki",
       opis="Higrometr punktu rosy wpisywany do K18 kazdej zakladki. Dla punktow "
            "tylko-temperatura zawsze wpisywane jest '-'."),
    _U("GEN_K18_CC04", "Higrometr K18 — komora CC-04", "tekst", "S8000", "ark",
       grupa="Naglowki"),
    _U("GEN_PODPIS_1", "Podpisujacy z lewej (B230)", "tekst", "Artsiom Azhdzer", "ark",
       grupa="Podpisy"),
    _U("GEN_PODPIS_2", "Podpisujacy z prawej (H230)", "tekst", "Marek Szpakowski", "ark",
       grupa="Podpisy"),
    _U("GEN_NR_POM", "Numer pomieszczenia srodowiskowego", "calk", 9, "ark",
       grupa="Warunki srodowiskowe", minimum=1, maksimum=999,
       opis="Szuka pliku 'Pom. nr <N> (<model>) - <rok>.xlsx' — najpierw lokalnie, "
            "potem na \\\\PLUM4."),
    _U("GEN_MODEL_CZUJ", "Model czujnika srodowiskowego", "tekst", "MX1101-02", "ark",
       grupa="Warunki srodowiskowe"),
    _U("GEN_WORD_TEMP", "Szablon Word — tylko temperatura", "plik",
       "xxx_yyy_LA_TH_2026 - tylko temp.docx", "ark",
       grupa="Szablony Word", wzorzec=(".docx",),
       opis="Uzywany, gdy zadna zakladka nie ma aktywnej wilgotnosci."),
    _U("GEN_WORD_RH", "Szablon Word — zakres (z RH)", "plik",
       "xxx_yyy_LA_TH_2026 - zakres.docx", "ark",
       grupa="Szablony Word", wzorzec=(".docx",),
       opis="Uzywany, gdy WSZYSTKIE zakladki maja aktywna wilgotnosc."),
    _U("GEN_WORD_MIX", "Szablon Word — zakres + temperatura", "plik",
       "xxx_yyy_LA_TH_2026 - zakres + temp.docx", "ark",
       grupa="Szablony Word", wzorzec=(".docx",),
       opis="Uzywany, gdy czesc zakladek ma wilgotnosc, a czesc nie (dwie tabele). "
            "Warianty '(uzytkownik)' dobierane sa automatycznie na podstawie PZ."),

    # --- zaawansowane ---
    _U("GEN_K18_DOM", "Higrometr K18 — domyslny", "tekst", "S8000", "ark",
       poziom=ZAAWANSOWANY, grupa="Naglowki",
       opis="Uzywany, gdy typ komory nie pasuje ani do CC, ani do CC-04."),
    _U("GEN_AUTOREC", "Sprzataj pliki autoodzyskiwania Excela", "flaga", True, "ark",
       poziom=ZAAWANSOWANY, grupa="Stabilnosc Excela",
       opis="Stare pliki .xar potrafia urosnac do setek MB i wywalac Excela przy starcie."),
    _U("GEN_AUTOREC_DNI", "Usuwaj autoodzyskiwanie starsze niz", "calk", 1, "ark",
       poziom=ZAAWANSOWANY, grupa="Stabilnosc Excela", przyrostek=" dni",
       minimum=0, maksimum=365,
       opis="0 = czysc wszystko. Wartosc 1 chroni dzisiejsze odzyskiwanie recznej pracy."),
    _U("GEN_PROG_OSTRZ", "Ostrzegaj od tylu kopiowanych zakladek", "calk", 10, "ark",
       poziom=ZAAWANSOWANY, grupa="Stabilnosc Excela", minimum=1, maksimum=200),
    _U("GEN_TAB_RATIO", "Szerokosc paska zakladek w kopii", "liczba", 0.85, "ark",
       poziom=ZAAWANSOWANY, grupa="Wyglad kopii", minimum=0.1, maksimum=1.0,
       krok_wart=0.05,
       opis="Domyslne 0,6 Excela chowa zakladki za strzalkami przy wielu punktach."),
    _U("GEN_TOL_CZUJ", "Tolerancja czasu czujnika srodowiskowego", "liczba", 2.0, "ark",
       poziom=ZAAWANSOWANY, grupa="Warunki srodowiskowe", przyrostek=" min",
       minimum=0.1, maksimum=600.0, krok_wart=0.5,
       opis="Czujnik Pom. nr 9 zapisuje co 60 s bez przerw, wiec najblizszy rekord "
            "jest w granicach ~30 s. Gdy nic nie miesci sie w tolerancji, komorki "
            "F/G zostaja PUSTE — lepiej uzupelnic recznie niz wpisac warunki "
            "z zupelnie innej chwili."),
    _U("GEN_LINKOWANE", "Pliki linkowane (oddzielone srednikiem)", "tekst",
       "Obliczenia tdp, RH, C.xls;Wzory.xls", "ark",
       poziom=ZAAWANSOWANY, grupa="Pliki linkowane",
       opis="Musza byc otwarte w tej samej sesji Excela, inaczej formuly "
            "kalibracyjne (D246/F246/G246) nie policza sie."),
    _U("GEN_LINK_OBLICZENIA", "Sciezka serwerowa — Obliczenia tdp, RH, C.xls", "tekst",
       r"\\plum4\LabPomiarowe\Obliczenia tdp, RH, C.xls", "ark",
       poziom=ZAAWANSOWANY, grupa="Pliki linkowane",
       opis="Przywracana w kopii na koncu, po odczytaniu kalibracji."),
    _U("GEN_LINK_WZORY", "Sciezka serwerowa — Wzory.xls", "tekst",
       r"\\plum4\LabPomiarowe\Wzory.xls", "ark",
       poziom=ZAAWANSOWANY, grupa="Pliki linkowane"),
    _U("GEN_FILTR_KOLOR", "Filtruj dane Strony 3 wg koloru", "flaga", True, "ark",
       poziom=ZAAWANSOWANY, grupa="Filtr kolorow Strony 3"),
    _U("GEN_KOLOR_AKT", "Kolor komorek aktywnych", "kolor", "#CCFFCC", "ark",
       poziom=ZAAWANSOWANY, grupa="Filtr kolorow Strony 3",
       opis="Komorki w tym kolorze sa brane do kopii."),
    _U("GEN_KOLOR_POM", "Kolor komorek pomijanych", "kolor", "#BFBFBF", "ark",
       poziom=ZAAWANSOWANY, grupa="Filtr kolorow Strony 3"),
    _U("GEN_INNE_KOLORY", "Bierz komorki w pozostalych kolorach", "flaga", False, "ark",
       poziom=ZAAWANSOWANY, grupa="Filtr kolorow Strony 3"),
    _U("GEN_MAP_CC04", "Mapowanie typu CC-04 -> stale K11/K12/K13/K17", "tabela",
       [
           ["LG", "Pt100-09", "1586A-02", "101", "CC-04-LG"],
           ["LD", "Pt100-01", "1586A-02", "105", "CC-04-LD"],
           ["PD", "Pt100-18", "1586A-02", "107", "CC-04-PD"],
           ["PG", "Pt100-13", "1586A-02", "103", "CC-04-PG"],
       ], "ark",
       poziom=ZAAWANSOWANY, grupa="Mapowanie CC-04",
       kolumny=["Tag (S14)", "K11", "K12", "K13", "K17"],
       opis="Tag czytany z wiersza 14 Strony 3 (S:T14, U:V14, ...) decyduje o "
            "stalych wpisywanych do zakladek roboczych kopii."),
]

USTAWIENIA = _WSPOLNE + _ANALIZA + _OBSERWACJA + _ARKUSZE
WG_ENV = {u.env: u for u in USTAWIENIA}


def dla_kroku(krok, poziom=None):
    """Ustawienia danego kroku; opcjonalnie tylko danego poziomu."""
    return [u for u in USTAWIENIA
            if u.krok == krok and (poziom is None or u.poziom == poziom)]


def domyslne():
    """Slownik env -> wartosc domyslna."""
    return {u.env: u.domyslna for u in USTAWIENIA}


# ─────────────────────────────────────────────────────────────────────────────
# Trwalosc (JSON obok aplikacji)
# ─────────────────────────────────────────────────────────────────────────────
def wczytaj():
    """
    Wczytuje zapisane ustawienia, uzupelniajac braki wartosciami domyslnymi.
    Uszkodzony plik nie blokuje startu — wracamy wtedy do domyslnych.
    """
    wart = domyslne()
    try:
        with open(PLIK_USTAWIEN, encoding="utf-8") as f:
            zapisane = json.load(f)
        if isinstance(zapisane, dict):
            for k, v in zapisane.items():
                if k in wart:
                    wart[k] = v
    except FileNotFoundError:
        pass
    except (OSError, ValueError):
        pass
    return wart


def zapisz(wartosci):
    """Zapisuje ustawienia. Zwraca None przy sukcesie albo tekst bledu."""
    try:
        with open(PLIK_USTAWIEN, "w", encoding="utf-8") as f:
            json.dump(wartosci, f, ensure_ascii=False, indent=2, sort_keys=True)
        return None
    except OSError as e:
        return f"{type(e).__name__}: {e}"


def do_env(wartosci):
    """Slownik zmiennych srodowiskowych dla podprocesu-workera."""
    out = {}
    for env, u in WG_ENV.items():
        if env in wartosci:
            out[env] = u.do_env(wartosci[env])
    return out


# ═════════════════════════════════════════════════════════════════════════════
# STRONA SKRYPTOW — odczyt zmiennych srodowiskowych
# Kazda funkcja przyjmuje wartosc domyslna i zwraca ja, gdy zmiennej brak albo
# jest niepoprawna. Dzieki temu skrypt uruchomiony recznie dziala jak dawniej.
# ═════════════════════════════════════════════════════════════════════════════
_PRAWDA = ("1", "true", "tak", "yes", "on")


def flaga(nazwa, domyslna):
    v = os.environ.get(nazwa)
    return domyslna if v is None or not v.strip() else v.strip().lower() in _PRAWDA


def liczba(nazwa, domyslna):
    v = os.environ.get(nazwa)
    if v is None or not v.strip():
        return domyslna
    try:
        return float(v.replace(",", "."))
    except ValueError:
        return domyslna


def calk(nazwa, domyslna):
    v = os.environ.get(nazwa)
    if v is None or not v.strip():
        return domyslna
    try:
        return int(float(v.replace(",", ".")))
    except ValueError:
        return domyslna


def tekst(nazwa, domyslna):
    v = os.environ.get(nazwa)
    return domyslna if v is None or not v.strip() else v.strip()


def minuty(nazwa, domyslna_td):
    """Zwraca timedelta. domyslna_td to timedelta uzywana, gdy zmiennej brak."""
    v = os.environ.get(nazwa)
    if v is None or not v.strip():
        return domyslna_td
    try:
        return datetime.timedelta(minutes=float(v.replace(",", ".")))
    except ValueError:
        return domyslna_td


def lista(nazwa, domyslna, sep=";"):
    """Lista tekstow rozdzielonych srednikiem."""
    v = os.environ.get(nazwa)
    if v is None or not v.strip():
        return domyslna
    return [c.strip() for c in v.split(sep) if c.strip()]


def tabela(nazwa, domyslna):
    """Lista wierszy (lista list) zapisana w env jako JSON."""
    v = os.environ.get(nazwa)
    if v is None or not v.strip():
        return domyslna
    try:
        dane = json.loads(v)
        return dane if isinstance(dane, list) else domyslna
    except ValueError:
        return domyslna


def sciezka(nazwa, domyslna, baza):
    """
    Sciezka do pliku/folderu. Wartosc wzgledna rozwijana jest wzgledem `baza`
    (folderu roboczego), bezwzgledna zostaje bez zmian.
    """
    v = tekst(nazwa, domyslna)
    if not v:
        return v
    return v if os.path.isabs(v) else os.path.join(baza, v)
