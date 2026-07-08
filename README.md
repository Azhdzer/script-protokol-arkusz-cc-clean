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

  classDef in fill:#12325a,stroke:#3b82f6,color:#dbeafe;
  classDef xf fill:#3a2c0c,stroke:#f59e0b,color:#fde68a;
  classDef out fill:#0f3d24,stroke:#22c55e,color:#bbf7d0;
  classDef pz fill:#2e1245,stroke:#a855f7,color:#e9d5ff;
  classDef gui fill:#3d1030,stroke:#ec4899,color:#fbcfe8;
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

  classDef in fill:#12325a,stroke:#3b82f6,color:#dbeafe;
  classDef xf fill:#3a2c0c,stroke:#f59e0b,color:#fde68a;
  classDef out fill:#0f3d24,stroke:#22c55e,color:#bbf7d0;
  classDef pz fill:#2e1245,stroke:#a855f7,color:#e9d5ff;
  class A,D,PZ,ZEST in
  class B,E,F,G,H,N,O xf
  class C,L,P,Q out
  class I,J,K,M pz
```

---

## ▶️ Uruchamianie — jeden exe, trzy tryby

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
  U->>G: klik "Obserwacja / Protokol / Word"
  G->>G: _run(tryb): zbierz env (GEN_*, OBS_*, CC_*)
  G->>E: QProcess.start + CC_WORKER=obserwacje/arkusze
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

  classDef in fill:#12325a,stroke:#3b82f6,color:#dbeafe;
  classDef xf fill:#3a2c0c,stroke:#f59e0b,color:#fde68a;
  classDef out fill:#0f3d24,stroke:#22c55e,color:#bbf7d0;
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
| `app_gui.py` | 472 | 26 | Panel sterujący PySide6 (Windows 11 Fluent) | PySide6 |
| `app_entry.py` | 83 | 3 | Dyspozytor zamrożonego exe (GUI ↔ worker) | — |

## 📥 Obsługiwane formaty loggerów (`analizuj_excele.py`)

`tempmate (PDF)` · `PuTTY/Vaisala (.log)` · `Rotronic HW4` · `ALMEMO` · `Comet/TFA` ·
`xTHERM (COM)` · `HOBO` · `Aranet` · `ElogVis` · generyczny `CSV` / `TXT` / `Excel`.
Wynik zawsze znormalizowany: **`Czas | Temperatura [°C] | Wilgotność [%RH]`**.

## 🔌 Kontrakt env‑var (GUI ↔ workery)

| Zmienna | Znaczenie |
|---|---|
| `CC_WORKER` | tryb: `arkusze` / `obserwacje` / (puste = GUI) |
| `CC_FOLDER` | folder roboczy |
| `CC_PROTOKOL` / `CC_SZABLON` | wybrany protokół / szablon |
| `CC_PZ_FOLDER` | folder PZ |
| `OBS_TXT_FILES` | lista TXT (średnik) — przerwany pomiar |
| `OBS_FILTR` / `OBS_PROG` / `OBS_TOL` | filtr / próg % / tolerancja min |
| `GEN_EXCEL` / `GEN_WORD` / `GEN_PUSTE` / `GEN_AUTOREC` | etapy i zachowania |

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

**Wymagania środowiska:** Windows 10/11, zainstalowany **Excel** (xlwings/COM),
dostęp do sieciowego `\\plum4\LabPomiarowe\Wzory.xls`. Zależności: `requirements.txt`
(`openpyxl · xlwings · python-docx · PySide6 · pandas · pypdf`).

### Sterowanie etapami (`generuj_arkusze.py`)

- `GENERUJ_EXCEL=True/False`, `GENERUJ_WORD=True/False`:
  `True/True` = pełny obieg · `True/False` = tylko Excel · `False/True` = Word z istniejących kopii.

### Warianty protokołu CC‑04

- Kolumny E/F na Stronie 3: CC = `Q/R`, CC‑04 = `S/T` (kolejne kopie +2 w prawo).
- Typ CC‑04 z wiersza 14 (`S:T14`…): tagi `LG` / `LD` / `PD` / `PG` → stałe `K11–K17`
  (Pt100‑09/01/18/13, 1586A‑02, …).

## 🆕 Ostatnie usprawnienia

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
