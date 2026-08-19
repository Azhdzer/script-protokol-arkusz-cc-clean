# Generator arkuszy i protokołów wzorcowania (CC / CC‑04)

Zestaw skryptów **Python** automatyzujący pełny obieg wzorcowania termohigrometrów
w akredytowanym laboratorium pomiarowym: od surowych logów czujników, przez analizę
stabilności i protokoły, po arkusze obliczeniowe **Excel** i świadectwa **Word** — spięty
desktopowym panelem **PySide6** i spakowany w jeden plik `ProtokolCC.exe`.

> 📊 **Pełna interaktywna infografika:** [`DOKUMENTACJA/infografika-projektu.html`](DOKUMENTACJA/infografika-projektu.html)
> (opisy każdej funkcji i ustawienia + diagramy). Otwórz w przeglądarce.

---

## 🗺️ Architektura — zależności modułów

```mermaid
flowchart LR
  subgraph IN["Wejscia (dane surowe)"]
    TXT["Plik TXT multimetru"]
    LOG["Logi czujnikow DUT<br/>csv / xlsx / txt / pdf / log"]
    PZ["PZ - Potwierdzenie zamowienia<br/>PDF (PL / EN)"]
    ZEST["Zestawienie<br/>wzorcowanych przyrzadow.xlsx"]
  end
  ANAL["analizuj_excele.py<br/>sniff_format - 11 parserow"]
  WYN["wyniki/&lt;serial&gt;_wynik.xlsx<br/>Czas | Temp | Wilg"]
  PZD["pz_dane.py<br/>wczytaj_pz + wczytaj_zestawienie"]
  OBS["generuj_obserwacje.py<br/>analiza stabilnosci + protokol"]
  PROTO["Protokol CC / CC-04<br/>Strona 2 przyrzady + Strona 3 pomiary"]
  ARK["generuj_arkusze.py<br/>kopie arkuszy + swiadectwa"]
  KOP["Kopie Excel"]
  DOC["Swiadectwa Word .docx"]
  GUI["app_gui.py + app_entry.py<br/>panel PySide6 / dyspozytor exe"]

  LOG --> ANAL --> WYN
  TXT --> OBS
  WYN --> OBS
  PZ --> PZD
  ZEST --> PZD
  PZD --> OBS
  OBS --> PROTO
  PROTO --> ARK
  PZD --> ARK
  ARK --> KOP
  ARK --> DOC
  GUI -.uruchamia.-> OBS
  GUI -.uruchamia.-> ARK

  classDef in fill:#dbeafe,stroke:#2563eb,color:#1e3a8a;
  classDef xf fill:#fef3c7,stroke:#d97706,color:#7c2d12;
  classDef out fill:#dcfce7,stroke:#16a34a,color:#14532d;
  classDef pz fill:#f3e8ff,stroke:#9333ea,color:#581c87;
  classDef gui fill:#fce7f3,stroke:#db2777,color:#831843;
  class TXT,LOG,PZ,ZEST in
  class ANAL,OBS,ARK xf
  class WYN,PROTO,KOP,DOC out
  class PZD pz
  class GUI gui
```

**Kluczowa zależność kolejności:** `analizuj_excele` musi wygenerować `wyniki/*` zanim
`generuj_obserwacje` dopasuje je czasowo do Strony 3. PZ wczytywany jest na starcie
obserwacji, a dopasowanie *przyrząd ↔ kolumna pomiarowa* idzie po **nr fabrycznym lub
nr ewidencyjnym** = serial z nazwy pliku wyniku.

---

## 🔄 Pipeline danych — przepływ działań (end‑to‑end)

```mermaid
flowchart TD
  A["Logi DUT + srodowiskowe"] --> B{{"analizuj_excele - sniff_format"}}
  B -->|"11 formatow"| C["wyniki/&lt;serial&gt;_wynik.xlsx"]
  D["TXT multimetru (1..N plikow)"] --> E["combine_txt - sklejenie + dedup po czasie"]
  E --> F["Arkusz obserwacji + analyze_and_highlight<br/>segmenty, okno 5-min, 5 reprezentantow"]
  F --> G["generuj_protokol"]
  C --> H["_wypelnij_wyniki_srodowiskowe<br/>dopasowanie 5 wierszy po czasie"]
  H --> G
  PZ["PZ/*.pdf"] --> I["pz_dane.wczytaj_pz<br/>klucz: nr_fabr + nr_ewid"]
  ZEST["Zestawienie.xlsx"] --> J["wczytaj_zestawienie - rozdzielczosc t/RH"]
  I --> K["wypelnij_strone2_z_pz"]
  J --> K
  G --> L["Protokol: Strona 3 (pomiary)"]
  K --> M["Protokol: Strona 2 (tabela przyrzadow)"]
  L --> N["generuj_arkusze._main_impl"]
  M --> N
  N --> O["7 etapow: kopie -> zakladki -> C/D/E/F -> naglowki (K11-K18) -> Wyniki -> srodowisko -> Word"]
  O --> P["Kopie Excel"]
  O --> Q["Swiadectwa Word .docx"]

  classDef in fill:#dbeafe,stroke:#2563eb,color:#1e3a8a;
  classDef xf fill:#fef3c7,stroke:#d97706,color:#7c2d12;
  classDef out fill:#dcfce7,stroke:#16a34a,color:#14532d;
  classDef pz fill:#f3e8ff,stroke:#9333ea,color:#581c87;
  class A,D,PZ,ZEST in
  class B,E,F,G,H,N,O xf
  class C,L,P,Q out
  class I,J,K,M pz
```

---

## 🌳 Pełny obieg pracy — od plików do świadectw (mega‑schemat)

Jedno drzewo spinające cały proces: co i gdzie wgrać ręcznie, które pliki wzorcowe zaktualizować,
a potem kolejno **analiza → obserwacja/protokół → arkusze/świadectwa** ze wszystkimi etapami.

```mermaid
flowchart TD
  subgraph F0["FAZA 0 - Przygotowanie reczne (zanim uruchomisz cokolwiek)"]
    direction TB
    P1["Wgraj pliki z przyrzadow (loggery DUT)<br/>-> folder excel_do_analizy/"]:::in
    P2["Wgraj PZ - Potwierdzenie zamowienia (PDF)<br/>-> folder PZ/"]:::in
    P6["Umiesc plik(i) TXT multimetru<br/>w folderze roboczym"]:::in
    P3["Zaktualizuj czujnik srodowiskowy<br/>Pom. nr 9 (MX1101-02) - 2026.xlsx"]:::in
    P4["Zaktualizuj pliki wzorcowe / linkowane<br/>Wzory.xls + Obliczenia tdp, RH, C.xls"]:::in
    P5["Zaktualizuj rozdzielczosci<br/>Zestawienie wzorcowanych przyrzadow.xlsx"]:::in
  end

  subgraph F1["1 - analizuj_excele.py - normalizacja logow"]
    direction TB
    A1{{"sniff_format() - wykrycie formatu z 1 KB"}}:::xf
    A2["11 parserow: tempmate/PDF, PuTTY/.log, Rotronic,<br/>ALMEMO, Comet/TFA, xTHERM, HOBO, Aranet, CSV/TXT/Excel"]:::xf
    A3["save_result() - Czas | Temperatura | Wilgotnosc"]:::xf
    A4[("wyniki/&lt;serial&gt;_wynik.xlsx")]:::out
    A1 --> A2 --> A3 --> A4
  end

  subgraph F2["2 - generuj_obserwacje.py - obserwacja + protokol"]
    direction TB
    B1["resolve_txt_files() - tolerancyjny wybor TXT"]:::xf
    B2["combine_txt() - sklejenie N plikow + dedup po czasie"]:::xf
    B3["Arkusz obserwacji (A2:U / A2:L)"]:::out
    B4["analyze_and_highlight()<br/>segmenty stabilne (B,C >= 1h45m)<br/>filtr suszenia / histerezy<br/>okno 5-min (min K/L) + 5 reprezentantow"]:::xf
    B5["wczytaj_pz() + wczytaj_zestawienie()"]:::pz
    B6["generuj_protokol()"]:::xf
    B8["_wypelnij_wyniki_srodowiskowe()<br/>dopasowanie 5 wierszy po czasie<br/>-> Q/R (CC) lub S/T (CC-04)"]:::xf
    B7["PROTOKOL - Strona 3 (pomiary + srodowisko)"]:::out
    B9["wypelnij_strone2_z_pz()<br/>match serial -> PZ, rozdzielczosc K/L"]:::pz
    B10["PROTOKOL - Strona 2 (tabela przyrzadow)"]:::out
    B1 --> B2 --> B3 --> B4 --> B6
    B5 --> B6
    B6 --> B8 --> B7
    B6 --> B7
    B5 --> B9 --> B10
  end

  subgraph F3["3 - generuj_arkusze.py - kopie arkuszy + swiadectwa (7 etapow)"]
    direction TB
    C0["_main_impl() -> wczytaj_wszystko_xlwings()<br/>Strona 2 (lista) + Strona 3 (zakladki, C/D, E/F, F24)"]:::xf
    C1["Etap 1 - kopie szablonu (per przyrzad)"]:::xf
    C2["Etap 2 - liczba i nazwy zakladek"]:::xf
    C34["Etap 3-4 - C15:C19, D15:D19, E15:E19, F15:F19"]:::xf
    C5["Etap 5 - naglowki: E4/G6/K4/E5/E6/H57,<br/>K11-K13/K17 (CC-04), K18 higrometr, podpisy"]:::xf
    C6["Etap 6 - Wyniki: F24, sort tabel,<br/>_aktualizuj_formule_histerezy (J23)"]:::xf
    C7["Etap 7 - warunki srodowiskowe<br/>czujnik + Wzory.xls -> F/G + min/max"]:::xf
    C8["_odczytaj_kalibracje_xlwings()<br/>D239:K239 (RH) / D246:G246 (temp)"]:::xf
    C9["utworz_kopie_word()<br/>szablon wg klasy, placeholdery,<br/>[uzytkownik], tabele kalibracji"]:::pz
    C10[("Kopie Excel - arkusze obliczeniowe")]:::out
    C11[("Swiadectwa Word - .docx")]:::out
    C0 --> C1 --> C2 --> C34 --> C5 --> C6 --> C7 --> C8 --> C9
    C7 --> C10
    C9 --> C11
  end

  P1 --> A1
  P6 --> B1
  A4 --> B8
  P2 --> B5
  P5 --> B5
  B7 --> C0
  B10 --> C0
  P3 --> C7
  P4 --> C7
  P4 --> C0
  P2 -.fallback.-> C9

  classDef in fill:#dbeafe,stroke:#2563eb,color:#1e3a8a;
  classDef xf fill:#fef3c7,stroke:#d97706,color:#7c2d12;
  classDef out fill:#dcfce7,stroke:#16a34a,color:#14532d;
  classDef pz fill:#f3e8ff,stroke:#9333ea,color:#581c87;
```

**Kolejność:** FAZA 0 (ręcznie) → `analizuj_excele` (→ `wyniki/*`) → `generuj_obserwacje`
(protokół: Strona 2 z PZ + Strona 3 z pomiarów) → `generuj_arkusze` (kopie Excel + świadectwa Word).
Pliki `Wzory.xls` / `Pom. nr 9` wchodzą dopiero w **Etap 7** arkuszy; `Zestawienie` i `PZ` — przy budowie Strony 2.

---

## 🖥️ Panel sterujący — jak się go używa

Okno prowadzi przez obieg krok po kroku (lewa listwa), a każdy krok ma własny formularz:

| Krok | Co robisz |
|---|---|
| **Przygotowanie** | Lista kontrolna: dla każdego pliku wejściowego widać **jego wiek** — `Wzory.xls` sprzed 40 dni dostaje pomarańczowe „ZAKTUALIZUJ przed uruchomieniem". Przycisk `Otworz` otwiera plik/folder. |
| **1 · Analiza logów** | Podgląd plików wgranych do `excel_do_analizy/` (data, rozmiar) + uruchomienie. |
| **2 · Obserwacja** | **Zaznaczanie plików TXT multimetru** z listy (posortowanej od najnowszego); kilka zaznaczeń = przerwany pomiar. Podpisy, szablony, PZ, progi filtra. |
| **3 · Arkusze i Word** | Protokół, szablon, numer świadectwa, higrometr K18, podpisy, szablony Word, etapy. |
| **Zaawansowane** | Reszta ustawień pogrupowana wg kroku: okno analizy, korekta zegara, filtr kolorów Strony 3, ścieżki serwerowe, mapowanie CC‑04. |

Na dole okna zawsze widoczne są dwie zakładki:

- **Log** — żywe wyjście workera, z kolorowaniem (`!!! BLAD` czerwony, `UWAGA` pomarańczowy,
  `[OK]` zielony) i zapisem do pliku,
- **Utworzone pliki** — różnica stanu folderu przed/po przebiegu: co powstało, co się
  zmieniło, ile waży i **jakie zakładki są w środku** (dwuklik otwiera plik).

`▶▶ Uruchom caly obieg` odpala kroki 1 → 2 → 3 po kolei i **przerywa się na pierwszym
błędzie**. Przed startem panel sprawdza dane wejściowe (brak zaznaczonego TXT, brak
protokołu, brak `Wzory.xls`…) i pokazuje listę zastrzeżeń do zatwierdzenia.

## ▶️ Uruchamianie — jeden exe, cztery tryby

Panel GUI uruchamia skrypty‑workery jako podprocesy. W zamrożonym `.exe` ten sam plik pełni
rolę i GUI, i workera — o trybie decyduje zmienna `CC_WORKER`.

```mermaid
sequenceDiagram
  participant U as Uzytkownik
  participant E as ProtokolCC.exe
  participant AE as app_entry.main()
  participant G as app_gui (GUI)
  participant W as worker (generuj_*)
  U->>E: dwuklik (brak CC_WORKER)
  E->>AE: start
  AE->>G: CC_WORKER puste -> uruchom GUI
  U->>G: klik "Uruchom caly obieg" albo pojedynczy krok
  G->>G: _zapisz() -> cc_ustawienia.json, walidacja wejscia
  G->>G: cc_config.do_env(wartosci) -> 65 zmiennych
  G->>E: QProcess.start + CC_WORKER=analiza/obserwacje/arkusze
  E->>AE: start (worker)
  AE->>W: CC_WORKER ustawione -> import generuj_*, main()
  W-->>G: stdout przez potok QProcess -> zywy log
  W-->>G: kod wyjscia (0 = OK)
```

---

## 🔗 Interakcje danych — Strona 2 ↔ Strona 3

```mermaid
flowchart LR
  subgraph P["Protokol (jeden plik xlsx)"]
    S2["Strona 2<br/>tabela przyrzadow B:J, K/L, O"]
    S3["Strona 3<br/>pomiary + srodowisko Q/R lub S/T"]
  end
  WYN["wyniki/&lt;serial&gt;_wynik.xlsx"]
  PZ["PZ + Zestawienie"]
  TXT["TXT multimetru"]
  TXT --> S3
  WYN -->|"dopasowanie po czasie"| S3
  PZ -->|"match serial = nazwa pliku wyniku"| S2
  S3 -->|"serial - kolumna (kolejnosc)"| S2
  S2 --> KOP["Kopie: E5=wytworca, E6=typ, G6=serial, K18=higrometr"]
  S3 --> KOP
  S2 --> DOC["Word: [wytworca][typ][nr_fabr][nr_ewid][uzytkownik]"]
  S3 --> DOC

  classDef in fill:#dbeafe,stroke:#2563eb,color:#1e3a8a;
  classDef xf fill:#fef3c7,stroke:#d97706,color:#7c2d12;
  classDef out fill:#dcfce7,stroke:#16a34a,color:#14532d;
  class WYN,PZ,TXT in
  class KOP,DOC xf
  class S2,S3 out
```

**Reguła sprzężenia:** i‑ty przyrząd na Stronie 2 (wiersz 11+i) odpowiada i‑tej parze kolumn
pomiarowych na Stronie 3 (S:T, U:V… dla CC‑04). Liczba przyrządów = liczbie kolumn pomiarowych.
Rozdzielczość K/L: z Zestawienia (producent+typ), a gdy brak — z wahania cyfr po przecinku w danych.

---

## 📦 Moduły

| Plik | Linie | Funkcje | Rola | Biblioteki |
|---|---:|---:|---|---|
| `generuj_arkusze.py` | 3402 | 97 | Kopie arkuszy obliczeniowych Excel + świadectwa Word | xlwings, openpyxl, python‑docx |
| `generuj_obserwacje.py` | 1549 | 41 | Arkusz obserwacji + protokół (CC / CC‑04) z pliku TXT | openpyxl, xlwings (COM) |
| `analizuj_excele.py` | 1093 | 28 | Uniwersalny parser loggerów → `wyniki/*.xlsx` | pandas, openpyxl, pypdf |
| `pz_dane.py` | 346 | 16 | Parser PZ (PDF, PL/EN) + Zestawienie rozdzielczości | pypdf, openpyxl |
| `app_gui.py` | ~740 | 35 | Panel sterujący PySide6 — kroki obiegu, log, wyniki | PySide6 |
| `cc_config.py` | ~500 | 12 | Rejestr **65 ustawień** + zapis `cc_ustawienia.json` + odczyt env | — (stdlib) |
| `cc_widgets.py` | ~560 | 30 | Widgety panelu: pola, lista plików, log, wyniki, checklista | PySide6 |
| `app_entry.py` | 85 | 3 | Dyspozytor zamrożonego exe (GUI ↔ worker) | — |
| `testy/` | ~1200 | 284 testów | Rejestr, kontrakt panel↔skrypty, widgety, panel, pełny obieg | unittest (stdlib) |

## 📥 Obsługiwane formaty loggerów (`analizuj_excele.py`)

`tempmate (PDF)` · `PuTTY/Vaisala (.log)` · `Rotronic HW4` · `ALMEMO` · `Comet/TFA` ·
`xTHERM (COM)` · `HOBO` · `Aranet` · `ElogVis` · generyczny `CSV` / `TXT` / `Excel`.
Wynik zawsze znormalizowany: **`Czas | Temperatura [°C] | Wilgotność [%RH]`**.

## 🔌 Kontrakt env‑var (GUI ↔ workery)

**Nie ma już potrzeby edytowania skryptów w Notatniku.** Wszystkie 65 ustawień
opisuje rejestr w [`cc_config.py`](cc_config.py) (typ, wartość domyślna, etykieta,
opis, krok obiegu, poziom podstawowy/zaawansowany). Panel buduje z niego formularze
automatycznie, zapisuje wybory do `cc_ustawienia.json` i przekazuje je workerowi
przez zmienne środowiskowe. Stałe w skryptach pełnią rolę **wartości domyślnych** —
uruchomienie `python generuj_obserwacje.py` bez panelu działa jak dawniej.

| Prefiks | Zakres | Przykłady |
|---|---|---|
| `CC_WORKER` | tryb exe | `analiza` / `obserwacje` / `arkusze` / (puste = GUI) |
| `CC_*` | wspólne | `CC_FOLDER`, `CC_PROTOKOL`, `CC_SZABLON`, `CC_PZ_FOLDER`, `CC_ZESTAWIENIE` |
| `ANL_*` | krok 1 — analiza | `ANL_PLIKI`, `ANL_INPUT`, `ANL_OUTPUT`, `ANL_DEBUG` |
| `OBS_*` | krok 2 — obserwacja | `OBS_TXT_FILES`, `OBS_PODPIS`, `OBS_STAB_PO_RH`, `OBS_PROG_T`, `OBS_KOREKTA_ZEGARA`, `OBS_FOTO`, `OBS_FOTO_ZRODLO`, `OBS_FOTO_CEL`, … |
| `GEN_*` | krok 3 — arkusze/Word | `GEN_EXCEL`, `GEN_WORD`, `GEN_POMIJAJ_PUSTE`, `GEN_NR_SW`, `GEN_K18_CC`, `GEN_PODPIS_1`, `GEN_MAP_CC04`, `GEN_KOLOR_AKT`, … |

Konwersję typów robią helpery `cc_config.flaga / liczba / calk / tekst / minuty /
lista / tabela / sciezka` — niepoprawna wartość cofa się do domyślnej zamiast wywalać skrypt.

---

## 🚀 Uruchomienie

```powershell
# Z źródła (panel GUI):
.venv\Scripts\python.exe app_gui.py

# Pojedyncze skrypty:
.venv\Scripts\python.exe analizuj_excele.py            # excel_do_analizy/ -> wyniki/
.venv\Scripts\python.exe generuj_obserwacje.py         # TXT -> obserwacja + protokol
.venv\Scripts\python.exe generuj_arkusze.py            # protokol -> kopie + Word

# Budowa jednego .exe:
powershell -ExecutionPolicy Bypass -File build.ps1     # -> dist\ProtokolCC.exe
```

### 🧪 Testy

```powershell
# Wszystko (z krokiem 3 — uruchamia Excel, ok. 3 min):
.venv\Scripts\python.exe -m unittest discover -s testy -t testy

# Szybko, bez Excela:
$env:CC_TESTY_SZYBKIE=1; .venv\Scripts\python.exe -m unittest discover -s testy -t testy
```

**284 testów** na czystym `unittest` (bez dodatkowych zależności), w tym pełny
obieg 1 → 2 → 3 na prawdziwym pomiarze 188. Wszystko dzieje się w
`testy/_piaskownica/` — testy nie dotykają plików projektu. Szczegóły:
[`testy/README.md`](testy/README.md).

Najważniejszy jest `testy/test_kontrakt.py`: pilnuje, że **każde** ustawienie
widoczne w panelu naprawdę zmienia stałą w skrypcie (i że bez panelu obowiązuje
wartość domyślna). Dodanie ustawienia do rejestru bez podłączenia go do skryptu
wywala suite.

**Wymagania środowiska:** Windows 10/11, zainstalowany **Excel** (xlwings/COM),
dostęp do sieciowego `\\plum4\LabPomiarowe\Wzory.xls`. Zależności: `requirements.txt`
(`openpyxl · xlwings · python-docx · PySide6 · pandas · pypdf`).

### Sterowanie etapami (`generuj_arkusze.py`)

Przełączniki `Generuj arkusze Excel` / `Generuj swiadectwa Word` w kroku 3 panelu
(env `GEN_EXCEL` / `GEN_WORD`): `wł/wł` = pełny obieg · `wł/wył` = tylko Excel ·
`wył/wł` = Word z już istniejących kopii.

### Warianty protokołu CC‑04

- Kolumny E/F na Stronie 3: CC = `Q/R`, CC‑04 = `S/T` (kolejne kopie +2 w prawo).
- Typ CC‑04 z wiersza 14 (`S:T14`…): tagi `LG` / `LD` / `PD` / `PG` → stałe `K11–K17`
  (Pt100‑09/01/18/13, 1586A‑02, …).

## 🆕 Ostatnie usprawnienia

- **Wybór pojedynczego przyrządu przez wyszarzenie:** przyrząd, którego wszystkie
  bloki na Stronie 3 są szare, nie dostaje już ani kopii Excel, ani świadectwa Word.
  Wcześniej powstawała pusta kopia (sam arkusz `Wyniki`) i świadectwo z zerową
  tabelą kalibracji — lista kopii idzie ze Strony 2, więc same kolory nie usuwały
  przyrządu z obiegu. Numer przyrządu z protokołu jest zachowany (`… - 4 - <serial>`).
- **Panel zamiast Notatnika:** wszystkie 65 ustawień z trzech skryptów ma teraz
  formularz w aplikacji (rejestr `cc_config.py`), zapisywany do `cc_ustawienia.json`.
  Doszły: lista kontrolna świeżości plików wejściowych, zaznaczanie plików TXT z listy,
  uruchamianie kroków po kolei z przerwaniem na błędzie, kolorowany log i panel
  „Utworzone pliki" pokazujący, co powstało i jakie ma zakładki.
- **PZ (Potwierdzenie zamówienia):** automatyczne wypełnianie tabeli przyrządów na Stronie 2
  (dwujęzycznie **PL/EN**), dopasowanie po nr fabrycznym **i** nr ewidencyjnym, blok `UŻYTKOWNIK/USER`
  → warianty szablonów Word `(uzytkownik)`.
- **analizuj_excele:** dodane formaty **PDF (tempmate)** i **PuTTY/Vaisala (.log)**.
- **generuj_obserwacje:** obsługa **wielu plików TXT** (przerwany pomiar, sklejanie + dedup),
  tolerancyjny wybór pliku, parser CC‑04 odporny na dodatkowe kanały.
- **generuj_arkusze:** konfigurowalny higrometr **`K18`** (temp‑only → „-"), automatyczna aktualizacja
  formuły histerezy **`Wyniki!J23`**, poprawka daty wzorcowania.

---

## 🤖 Agenty pomocnicze (`.claude/agents/`)

`python-pro` · `pyside6-desktop-expert` · `desktop-packager` · `debugger` ·
`code-reviewer` · `test-automator` · `project-visualizer` (analiza i wizualizacja architektury).

---

*Dokumentacja i diagramy generowane automatycznie z analizy kodu. Pełny opis każdej funkcji
i ustawienia: [`DOKUMENTACJA/infografika-projektu.html`](DOKUMENTACJA/infografika-projektu.html).*
