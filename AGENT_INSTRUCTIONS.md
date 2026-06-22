# Instrukcja dla agenta — Etap 1 + Etap 2: Tworzenie kopii arkuszy obliczeniowych z automatycznym zarządzaniem zakładkami

## Cel zadania

Napisz skrypt Python, który:
1. Odczytuje dane z pliku protokołu (Excel)
2. Na podstawie tych danych tworzy nazwane kopie pliku szablonu (Excel)
3. Zapisuje kopie w tym samym folderze co oryginały
4. W każdej kopii automatycznie dostosowuje zakładki (sheets): usuwa zbędne i zmienia nazwy pozostałych zgodnie z danymi z protokołu

---

## Środowisko

- System operacyjny: Windows
- Język skryptu: **Python 3**
- Wymagane biblioteki: `openpyxl`, `shutil`, `os`, `re`
- Skrypt ma działać w tym samym folderze, w którym znajdują się pliki Excel
- **Uwaga:** `openpyxl` nie obsługuje formuł Excel ani makr VBA — skrypt operuje wyłącznie na strukturze zakładek (ich nazwach i kolejności), nie zmienia zawartości komórek szablonu

---

## Pliki wejściowe

### Plik szablonu
Nazwa pliku szablonu:
```
xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.11 z 29.08.2025 - 1 - RH (CC).xlsx
```
Plik szablonu jest stały — jego nazwa nie zmienia się. Skrypt powinien zawierać zmienną konfiguracyjną z pełną nazwą tego pliku.

### Plik protokołu
Aktualna nazwa pliku protokołu:
```
133; 148_LA_TH_2026 - protokół CC.xlsx
```
> **Uwaga:** Nazwa pliku protokołu może się zmieniać między uruchomieniami. Dlatego skrypt musi zawierać **zmienną konfiguracyjną** (na górze pliku), w której użytkownik może wpisać aktualną nazwę pliku protokołu przed uruchomieniem skryptu.

---

## Logika skryptu

### Krok 1 — Konfiguracja (zmienne na górze skryptu)

```python
# === KONFIGURACJA ===
FOLDER = r"."  # folder roboczy (domyślnie ten sam co skrypt)
PROTOKOL_PLIK = "133; 148_LA_TH_2026 - protokół CC.xlsx"  # <- zmień jeśli nazwa protokołu się zmieni
SZABLON_PLIK  = "xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.11 z 29.08.2025 - 1 - RH (CC).xlsx"
ARKUSZ_STRONA2 = "Strona 2"   # arkusz z listą kopii do wygenerowania
ARKUSZ_STRONA3 = "Strona 3"   # arkusz z definicją zakładek dla każdej kopii
START_ROW_S2 = 11   # wiersz startowy w Strona 2 (kolumna A)
START_ROW_S3 = 20   # wiersz startowy w Strona 3 (kolumna A)
BLOK_S3 = 5         # liczba wierszy zajmowanych przez jedną zakładkę w Strona 3
ARKUSZ_CHRONIONY = "Wyniki"   # ta zakładka nigdy nie jest usuwana ani zmieniana
START_COL_E_S3 = 17  # kolumna Q (1-indexed) — źródło dla E15:E19, pierwsza kopia
START_COL_F_S3 = 18  # kolumna R (1-indexed) — źródło dla F15:F19, pierwsza kopia
KROK_COL_EF   = 2    # przesunięcie kolumny dla każdej kolejnej kopii (co 2 kolumny w prawo)
```

### Krok 2 — Odczyt danych z protokołu (Strona 2) — lista kopii

- Otwórz plik protokołu (`PROTOKOL_PLIK`) za pomocą `openpyxl` w trybie `read_only=True`
- Przejdź do arkusza o nazwie `Strona 2`
- Iteruj wiersze zaczynając od wiersza `START_ROW_S2` (wiersz 11) w dół
- **Warunek zatrzymania:** kolumna **A** jest pusta (`None` lub pusty string)
- Dla każdego wiersza (gdzie kolumna A nie jest pusta) odczytaj:
  - `wartosc_O` — wartość z kolumny **O** (15. kolumna) — zastępuje `xxx` w nazwie pliku
  - `wartosc_E` — wartość z kolumny **E** (5. kolumna) — zastępuje `RH (CC)` w nazwie pliku

### Krok 3 — Generowanie nazwy kopii

Szablon nazwy:
```
xxx_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.11 z 29.08.2025 - 1 - RH (CC).xlsx
```

Zasady zamiany w nazwie:
1. Zamień fragment `xxx` na wartość z kolumny **O** (jako string, bez spacji wiodących/końcowych)
2. Zamień fragment `RH (CC)` na wartość z kolumny **E** (jako string, bez spacji wiodących/końcowych)

Przykłady wynikowych nazw (dla danych z pliku przykładowego):
- Wiersz 11: `133_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.11 z 29.08.2025 - 1 - 1010223.xlsx`
- Wiersz 12: `133_LA_TH_2026 - ILAJ 5.4_11#21 - Wzór ark. obl.Wer.11 z 29.08.2025 - 1 - 290222.xlsx`

> **Uwaga:** Wartości z kolumn O i E mogą być liczbami całkowitymi (int) lub liczbami zmiennoprzecinkowymi — konwertuj je do `int` jeśli nie mają części ułamkowej, przed zamianą na string (aby uniknąć np. `133.0` zamiast `133`).

### Krok 4 — Tworzenie kopii pliku szablonu

- Użyj `shutil.copy2()` do skopiowania pliku szablonu pod nową nazwą
- Ścieżka docelowa: ten sam folder co szablon (`FOLDER`)
- Jeśli plik o danej nazwie już istnieje — **nadpisz go** (lub wyświetl ostrzeżenie i kontynuuj)

### Krok 5 — Raportowanie (Etap 1)

Po zakończeniu tworzenia kopii skrypt wypisuje w konsoli:
- Liczbę wierszy znalezionych w protokole
- Listę wszystkich utworzonych plików (nazwa każdego)

---

## Etap 2 — Automatyczne zarządzanie zakładkami w każdej kopii

Po stworzeniu kopii (Etap 1) skrypt musi wejść do każdego skopiowanego pliku i dostosować jego zakładki.

### Krok 6 — Odczyt definicji zakładek z protokołu (Strona 3)

Otwórz plik protokołu i przejdź do arkusza `Strona 3`.

**Logika odczytu — bloki 5-wierszowe:**

Każda wymagana zakładka jest opisana blokiem 5 wierszy w kolumnach A, B, C, zaczynając od wiersza `START_ROW_S3` (wiersz 20):

- Blok 1: wiersze 20–24
- Blok 2: wiersze 25–29
- Blok 3: wiersze 30–34
- itd.

**Warunek zatrzymania:** komórka A w pierwszym wierszu bloku jest pusta.

Dla każdego bloku odczytaj wartości z **pierwszego wiersza bloku**:
- Kolumna **B** → pierwsza część nazwy zakładki
- Kolumna **C** → druga część nazwy zakładki

Nazwa zakładki: `"{B}, {C}"` (z przecinkiem i spacją między wartościami)

**Przykład z pliku `133; 148_LA_TH_2026 - protokół CC`:**

| Blok | Wiersze | B      | C  | Nazwa zakładki |
|------|---------|--------|----|----------------|
| 1    | 20–24   | 20.3   | -  | `20.3, -`      |
| 2    | 25–29   | 5.3    | -  | `5.3, -`       |
| 3    | 30–34   | -19.9  | -  | `-19.9, -`     |
| 4    | 35–39   | *puste*|    | ← stop         |

Wynik: lista 3 nazw zakładek: `["20.3, -", "5.3, -", "-19.9, -"]`

**Uwaga:** Wartości B i C mogą być liczbami zmiennoprzecinkowymi lub stringami (np. `-`). Konwertuj do stringa:
- Jeśli wartość to float bez części ułamkowej (np. `20.0`) → zamień na `int` → string: `"20"`
- Jeśli float z częścią ułamkową (np. `20.3`) → zachowaj jako float → string: `"20.3"`
- Jeśli string lub None → użyj bezpośrednio jako string (None traktuj jako pusty `""`)

### Krok 7 — Modyfikacja zakładek w każdej skopiowanej kopii

> **Kluczowa zasada:** Zakładki w szablonie (`23,30`, `23,60`, `23,85`, `23,60 (2)`, `10,60`, `35,60`) **nie są puste** — zawierają dane obliczeniowe. Skrypt **nie dotyka zawartości komórek** żadnej zakładki. Jedyne dozwolone operacje to: zmiana nazwy zakładki oraz usunięcie całej zakładki (razem z jej zawartością, gdy jest nadmiarowa). Zawartość zachowanych zakładek pozostaje niezmieniona — pochodzi bezpośrednio z szablonu skopiowanego w Etapie 1.

Po odczycie listy zakładek z protokołu, dla każdego skopiowanego pliku:

1. Otwórz plik za pomocą `openpyxl` w trybie zapisu (nie `read_only`)
2. Pobierz listę wszystkich zakładek z pliku
3. Wydziel zakładkę chronioną (`Wyniki`) — **nigdy jej nie dotykaj ani nie przenoś**
4. Pobierz pozostałe zakładki (robocze) w ich oryginalnej kolejności z szablonu — np. `["23,30", "23,60", "23,85", "23,60 (2)", "10,60", "35,60"]`
5. Określ wymaganą liczbę zakładek roboczych: `N = len(lista_zakładek_z_protokołu)`
6. **Dopasuj liczbę zakładek roboczych do N:**
   - Jeśli roboczych jest **więcej niż N** → usuń nadmiarowe od końca listy (wraz z ich zawartością)
   - Jeśli roboczych jest **mniej niż N** → dla każdej brakującej zakładki wykonaj **kopię pierwszej zakładki roboczej** (np. `"23,30"`) używając `wb.copy_worksheet()`, wstaw kopię przed zakładką `Wyniki`. Kopiuj tyle razy ile brakuje.
   - Jeśli liczba jest **równa N** → bez zmian w liczbie
7. **Zmień nazwy wszystkich zakładek roboczych** (po dopasowaniu liczby): przemianuj każdą na odpowiadającą nazwę z listy protokołu, zachowując kolejność (pierwsza robocza → pierwsza nazwa z protokołu, itd.). Zawartość zakładek pozostaje bez zmian.
8. Zakładka `Wyniki` pozostaje bez jakichkolwiek zmian — jej nazwa, pozycja i zawartość są niemodyfikowane
9. Zapisz plik

**Przykład A — protokół wymaga MNIEJ zakładek niż jest w szablonie (nasz przypadek):**

Szablon ma zakładki: `["23,30", "23,60", "23,85", "23,60 (2)", "10,60", "35,60", "Wyniki"]`
Wymagane zakładki z protokołu: `["20.3, -", "5.3, -", "-19.9, -"]` (3 sztuki, szablon ma 6 roboczych)

Operacje:
- Usunięte (nadmiarowe, od końca): `"23,60 (2)"`, `"10,60"`, `"35,60"`
- Zachowane i przemianowane (zawartość nienaruszona): `"23,30"` → `"20.3, -"`, `"23,60"` → `"5.3, -"`, `"23,85"` → `"-19.9, -"`
- Bez zmian: `"Wyniki"`

Wynikowe zakładki: `["20.3, -", "5.3, -", "-19.9, -", "Wyniki"]`

**Przykład B — protokół wymaga WIĘCEJ zakładek niż jest w szablonie:**

Szablon ma zakładki: `["23,30", "23,60", "Wyniki"]` (2 robocze)
Wymagane zakładki z protokołu: `["20.3, -", "5.3, -", "-19.9, -", "1.0, -"]` (4 sztuki)

Operacje:
- Brakuje 2 zakładek → skopiuj `"23,30"` dwukrotnie przez `wb.copy_worksheet()`, wstaw kopie przed `"Wyniki"`
- Przemianuj wszystkie 4 robocze: `"23,30"` → `"20.3, -"`, `"23,60"` → `"5.3, -"`, kopia1 → `"-19.9, -"`, kopia2 → `"1.0, -"`
- Bez zmian: `"Wyniki"`

Wynikowe zakładki: `["20.3, -", "5.3, -", "-19.9, -", "1.0, -", "Wyniki"]`

### Krok 8 — Raportowanie końcowe

Po przetworzeniu wszystkich kopii skrypt wypisuje w konsoli:
- Liczbę zakładek odczytanych z protokołu (`Strona 3`)
- Dla każdej przetworzonej kopii: jej nazwę i listę zakładek po modyfikacji
- Komunikat końcowy: `"Zakończono. Utworzono X kopii, każda z Y zakładkami + Wyniki."`

---

## Etap 3 — Wypełnianie komórek w zakładkach roboczych każdej kopii

### Krok 9 — Odczyt danych do wypełnienia z protokołu (Strona 3, kolumny L i M)

Podczas odczytu danych z `Strona 3` (Krok 6), dla każdego bloku 5-wierszowego dodatkowo odczytaj:
- Kolumna **L** (12. kolumna), wiersze bloku (np. L20–L24 dla bloku 1) → lista 5 wartości dla komórek `C15:C19`
- Kolumna **M** (13. kolumna), wiersze bloku (np. M20–M24 dla bloku 1) → lista 5 wartości dla komórek `D15:D19`

Zaktualizuj funkcję `wczytaj_zakladki_z_protokolu_s3` tak, aby zwracała rozszerzoną strukturę:
```python
[
  {
    "nazwa":  "20.3, -",
    "L_dane": [L20, L21, L22, L23, L24],   # → C15:C19
    "M_dane": [M20, M21, M22, M23, M24],   # → D15:D19
  },
  {
    "nazwa":  "5.3, -",
    "L_dane": [L25, L26, L27, L28, L29],
    "M_dane": [M25, M26, M27, M28, M29],
  },
  ...
]
```

### Krok 10 — Zapis wartości do komórek C15:C19 i D15:D19 w każdej kopii

Po przemianowaniu zakładek (Krok 7), wciąż w tym samym otwartym pliku (`openpyxl`, tryb zapisu):

Dla każdej zakładki roboczej o indeksie `i` (0-bazowany):
1. Pobierz obiekt arkusza
2. Zapisz wartości z `L_dane[i]` do komórek `C15`, `C16`, `C17`, `C18`, `C19` (kolejno)
3. Zapisz wartości z `M_dane[i]` do komórek `D15`, `D16`, `D17`, `D18`, `D19` (kolejno)
4. Nie modyfikuj żadnych innych komórek w tym arkuszu

Zapisz plik po przetworzeniu wszystkich zakładek (jedno `wb.save()` na końcu, nie po każdej zakładce).

**Przykład dla naszego przypadku:**

| Zakładka w kopii | Źródło w Strona 3 | Cel w kopii   |
|------------------|-------------------|---------------|
| `"20.3, -"`      | L20:L24, M20:M24  | C15:C19, D15:D19 |
| `"5.3, -"`       | L25:L29, M25:M29  | C15:C19, D15:D19 |
| `"-19.9, -"`     | L30:L34, M30:M34  | C15:C19, D15:D19 |

> **Uwaga:** Każda zakładka robocza ma własne `C15:C19` i `D15:D19` — dane wpisywane są niezależnie do każdej zakładki. Zakładka `Wyniki` nie jest wypełniana.

---

## Etap 4 — Wypełnianie komórek E15:E19 i F15:F19 (dane zależne od numeru kopii)

### Krok 11 — Logika adresowania kolumn w Strona 3 dla E/F

W odróżnieniu od C/D (kolumny L i M — stałe dla wszystkich kopii), dane dla `E15:E19` i `F15:F19` **przesuwają się o 2 kolumny w prawo dla każdej kolejnej kopii**:

| Kopia (indeks `j`, 0-bazowany) | Źródło dla E15:E19 | Źródło dla F15:F19 |
|-------------------------------|---------------------|---------------------|
| 0 (kopia 1: `...1010223`)     | kolumna Q (17)      | kolumna R (18)      |
| 1 (kopia 2: `...290222`)      | kolumna S (19)      | kolumna T (20)      |
| 2 (kopia 3)                   | kolumna U (21)      | kolumna V (22)      |
| j                             | kolumna 17 + 2*j    | kolumna 18 + 2*j    |

Wiersze źródłowe zależą od bloku zakładki (tak samo jak dla L/M):
- Zakładka robocza `i` (0-bazowany) → wiersze `START_ROW_S3 + i * BLOK_S3` do `START_ROW_S3 + i * BLOK_S3 + 4`

**Przykład — kopia 1 (`j=0`), wszystkie zakładki:**

| Zakładka       | Źródło E15:E19 | Źródło F15:F19 |
|----------------|----------------|----------------|
| `"20.3, -"` (i=0) | Q20:Q24     | R20:R24        |
| `"5.3, -"`  (i=1) | Q25:Q29     | R25:R29        |
| `"-19.9, -"` (i=2)| Q30:Q34     | R30:R34        |

**Przykład — kopia 2 (`j=1`), wszystkie zakładki:**

| Zakładka       | Źródło E15:E19 | Źródło F15:F19 |
|----------------|----------------|----------------|
| `"20.3, -"` (i=0) | S20:S24     | T20:T24        |
| `"5.3, -"`  (i=1) | S25:S29     | T25:T29        |
| `"-19.9, -"` (i=2)| S30:S34     | T30:T34        |

### Krok 12 — Odczyt danych E/F z protokołu

Dane E/F należy wczytać **przed** przetwarzaniem kopii, ponieważ wymagają znajomości liczby kopii.

Dodaj funkcję `wczytaj_dane_ef_z_protokolu_s3(sciezka, arkusz, n_kopii, start_row_s3, blok, start_col_e, krok_col)` która zwraca strukturę 2D:

```python
# dane_ef[j][i] = {"E_dane": [5 wartości], "F_dane": [5 wartości]}
# j = indeks kopii (0-bazowany)
# i = indeks zakładki (0-bazowany)
dane_ef = [
  [  # kopia j=0
    {"E_dane": [Q20..Q24], "F_dane": [R20..R24]},  # zakładka i=0
    {"E_dane": [Q25..Q29], "F_dane": [R25..R29]},  # zakładka i=1
    ...
  ],
  [  # kopia j=1
    {"E_dane": [S20..S24], "F_dane": [T20..T24]},  # zakładka i=0
    ...
  ],
  ...
]
```

### Krok 13 — Zapis wartości do E15:E19 i F15:F19

Podczas przetwarzania kopii o indeksie `j`, w tej samej sesji `openpyxl` co Krok 10 (przed `wb.save()`):

Dla każdej zakładki roboczej o indeksie `i`:
1. Pobierz `dane_ef[j][i]["E_dane"]` → zapisz do `E15`, `E16`, `E17`, `E18`, `E19`
2. Pobierz `dane_ef[j][i]["F_dane"]` → zapisz do `F15`, `F16`, `F17`, `F18`, `F19`
3. Zakładka `Wyniki` nie jest wypełniana

Jedno `wb.save()` na końcu po zapisie C/D/E/F (nie po każdej zakładce).

---

## Obsługa błędów

- Jeśli plik protokołu nie istnieje — wyświetl czytelny błąd i zakończ skrypt
- Jeśli arkusz `Strona 2` lub `Strona 3` nie istnieje — wyświetl czytelny błąd i zakończ skrypt
- Jeśli szablon nie istnieje — wyświetl czytelny błąd i zakończ skrypt
- Jeśli wartość w kolumnie O lub E jest `None` dla danego wiersza (Strona 2) — pomiń ten wiersz i wypisz ostrzeżenie z numerem wiersza
- Jeśli lista zakładek z protokołu (Strona 3) jest pusta — wyświetl ostrzeżenie i nie modyfikuj zakładek w kopiach
- Jeśli w szablonie jest mniej zakładek roboczych niż wymagana liczba z protokołu — skopiuj pierwszą zakładkę roboczą (np. `"23,30"`) tyle razy ile brakuje, używając `wb.copy_worksheet()`

---

## Struktura pliku skryptu

Plik skryptu: **`generuj_arkusze.py`**

Struktura:
```
generuj_arkusze.py
├── Sekcja: KONFIGURACJA (zmienne na górze)
│
├── Funkcja: wczytaj_dane_z_protokolu_s2(sciezka, arkusz, start_row)
│   └── zwraca listę słowników: [{"O": "133", "E": "1010223"}, ...]
│
├── Funkcja: wczytaj_zakladki_z_protokolu_s3(sciezka, arkusz, start_row, blok)
│   └── zwraca listę słowników: [{"nazwa": "20.3, -", "L_dane": [...], "M_dane": [...]}, ...]
│
├── Funkcja: wczytaj_dane_ef_z_protokolu_s3(sciezka, arkusz, n_kopii, start_row_s3, blok, start_col_e, krok_col)
│   └── zwraca dane_ef[j][i] = {"E_dane": [...5...], "F_dane": [...5...]}
│
├── Funkcja: generuj_nazwe_pliku(szablon_nazwa, wartosc_O, wartosc_E)
│   └── zwraca string z nową nazwą pliku
│
├── Funkcja: dostosuj_zakladki_i_wypelnij(sciezka_pliku, dane_zakladek, dane_ef_kopia, arkusz_chroniony)
│   └── otwiera kopię, usuwa zbędne zakładki, zmienia nazwy,
│       wypełnia C15:C19 (L_dane), D15:D19 (M_dane), E15:E19 (E_dane), F15:F19 (F_dane), zapisuje
│
├── Funkcja: utworz_kopie(folder, szablon_plik, dane, dane_zakladek, dane_ef, arkusz_chroniony)
│   └── tworzy kopie (shutil.copy2) i wywołuje dostosuj_zakladki_i_wypelnij dla każdej (z dane_ef[j])
│
└── Blok główny: if __name__ == "__main__"
```

---

## Uwagi dodatkowe

- Skrypt powinien być uruchamiany z poziomu folderu, w którym znajdują się pliki Excel, **lub** ze ścieżką podaną w zmiennej `FOLDER`
- Nie modyfikuj zawartości komórek szablonu ani protokołu — dozwolone zmiany w **kopiach**: zmiana nazwy zakładki, usunięcie nadmiarowej zakładki, zapis wartości do komórek `C15:C19`, `D15:D19`, `E15:E19`, `F15:F19` w zakładkach roboczych
- Nie twórz żadnych podfolderów — wszystkie kopie lądują w tym samym folderze co szablon
- Encoding pliku skryptu: **UTF-8**
- Kolejność zakładek w kopii: zakładki robocze (przemianowane) na początku, `Wyniki` na końcu — tak jak w oryginale

---

## Przykładowe dane z protokołu (Strona 2)

| Wiersz | Kolumna A | Kolumna E | Kolumna O |
|--------|-----------|-----------|-----------|
| 11     | (liczba)  | 1010223   | 133       |
| 12     | (liczba)  | 290222    | 133       |
| ...    | ...       | ...       | ...       |
| 35     | (liczba)  | ...       | ...       |
| 36     | *puste*   | —         | —         |

Iteracja kończy się na wierszu 35 (A36 jest puste) → **25 kopii**.

## Przykładowe dane z protokołu (Strona 3)

Kolumny A, B, C — definicja zakładek (stałe dla wszystkich kopii):

| Wiersze bloku | Kol. A    | Kol. B | Kol. C | Nazwa zakładki | Kol. L → C15:C19 | Kol. M → D15:D19 |
|---------------|-----------|--------|--------|----------------|-------------------|-------------------|
| 20–24         | (wartość) | 20.3   | -      | `20.3, -`      | L20:L24           | M20:M24           |
| 25–29         | (wartość) | 5.3    | -      | `5.3, -`       | L25:L29           | M25:M29           |
| 30–34         | (wartość) | -19.9  | -      | `-19.9, -`     | L30:L34           | M30:M34           |

Kolumny Q, R, S, T, ... — dane E/F (przesuwają się per kopia, wiersze jak wyżej):

| Kopia j | Zakładka i | Kol. E → E15:E19 | Kol. F → F15:F19 |
|---------|-----------|-------------------|-------------------|
| 0       | 0         | Q20:Q24           | R20:R24           |
| 0       | 1         | Q25:Q29           | R25:R29           |
| 0       | 2         | Q30:Q34           | R30:R34           |
| 1       | 0         | S20:S24           | T20:T24           |
| 1       | 1         | S25:S29           | T25:T29           |
| 1       | 2         | S30:S34           | T30:T34           |
| 2       | 0         | U20:U24           | V20:V24           |
| 35–39         | *puste*    | —      | —      | ← stop         | —                 | —                 |

Wynik: 3 zakładki robocze + `Wyniki` = **4 zakładki** w każdej kopii, każda zakładka robocza ma wypełnione `C15:C19`, `D15:D19`, `E15:E19`, `F15:F19`.

---

## Etap 5 — Wypełnianie komórek nagłówkowych i stopkowych

### Kontekst

Etap 5 wypełnia dodatkowe komórki w każdej zakładce roboczej (poza `Wyniki` i `Wnioski`) każdej kopii. Dane pochodzą z:
- Nazwy pliku kopii (prefiks, rok)
- Numeru fabrycznego (ostatni segment nazwy pliku)
- Protokołu, arkusz `Strona 2` (kolumny B, D, K — ten sam wiersz co kolumny E i O już używane)
- Zmiennych konfiguracyjnych (imiona podpisujących)
- Aktualnej daty systemowej

### Nowe zmienne konfiguracyjne (dodane do sekcji KONFIGURACJA)

```python
ARKUSZ_WNIOSKI   = "Wnioski"          # ostatni arkusz — nie jest modyfikowany
PODPISUJACY_1    = "Artsiom Azhdzer"  # B230:C230 (scalona)
PODPISUJACY_2    = "Marek Szpakowski" # H230:I230 (scalona)
```

### Nowe dane odczytywane z protokołu (`Strona 2`)

Dla każdego wiersza (j-tego rekordu), oprócz kolumn O i E, odczytaj też:
- Kolumna **B** (2) → źródło dla `E5` kopii
- Kolumna **D** (4) → źródło dla `E6` kopii
- Kolumna **K** (11) → źródło dla `H57` kopii

Zaktualizowana struktura `dane_s2`:
```python
{"O": "133", "E": "1010223", "B": <wartość>, "D": <wartość>, "K": <wartość>}
```

### Funkcja pomocnicza `_parsuj_nazwe_pliku(nowa_nazwa)`

Parsuje nazwę pliku kopii (bez ścieżki) i zwraca `(prefiks, rok, numer_fabryczny)`.

Zasady:
- **Prefiks** — część przed `_LA_TH_`; np. `"133_LA_TH_2026 - ..."` → `"133"`
- **Rok** — 4 znaki po `_LA_TH_`; np. `"2026"`
- **Numer fabryczny** — ostatni segment po `" - "` (po usunięciu `.xlsx`); np. `"1010223"`

Jeśli `_LA_TH_` nie występuje w nazwie — prefiks to część przed pierwszym `_`, rok = `datetime.now().year`.

### Krok 14 — Zapis nowych komórek w każdej zakładce roboczej

Po wypełnieniu C/D/E/F 15–19 (Etapy 3 i 4), w tej samej pętli, dla każdej zakładki roboczej o indeksie `i` kopii `j`:

| Komórka         | Źródło danych                                     | Uwagi                          |
|-----------------|---------------------------------------------------|--------------------------------|
| `E4` (scal. F4) | `f"{prefiks}/LA/TH/{rok}"`                        | np. `"133/LA/TH/2026"`         |
| `G6`            | `numer_fabryczny`                                 | np. `"1010223"`                |
| `E5`            | `rekord["B"]` z protokołu Strona 2                | None → pomiń (nie nadpisuj)    |
| `E6`            | `rekord["D"]` z protokołu Strona 2                | None → pomiń (nie nadpisuj)    |
| `H57`           | `rekord["K"]` z protokołu Strona 2                | None → pomiń (nie nadpisuj)    |
| `B228`          | `datetime.datetime.now()` (obiekt Python datetime) | Excel formatuje wg komórki    |
| `H228`          | `datetime.datetime.now()` (obiekt Python datetime) | Excel formatuje wg komórki    |
| `B230` (scal. C230) | `PODPISUJACY_1`                               | imię z konfiguracji            |
| `H230` (scal. I230) | `PODPISUJACY_2`                               | imię z konfiguracji            |

Scalone komórki — pisz tylko do lewej-górnej komórki zakresu (xlwings automatycznie obsługuje scalone zakresy).

### Zakładki pomijane w Etapie 5

Tak samo jak w Etapach 2–4:
- `ARKUSZ_CHRONIONY` = `"Wyniki"` — nigdy nie modyfikowany
- `ARKUSZ_WNIOSKI` = `"Wnioski"` — nigdy nie modyfikowany

Zmienna `wykluczone = {chroniony, ARKUSZ_WNIOSKI}` używana we wszystkich filtrach `working`.

---

---

## Etap 6 — Wypełnianie arkusza Wyniki w każdej kopii

### Kontekst

Arkusz **Wyniki** (ARKUSZ_CHRONIONY) nie jest usuwany ani przemianowywany (jak w Etapach 2–5), ale w Etapie 6 niektóre jego komórki są zapisywane per kopia.

### Dane dla F24 — odczyt z protokołu (Strona 3, wiersz 17)

Wartość F24 pochodzi z Strona 3, wiersz **17** (powyżej bloków zakładek zaczynających się od wiersza 20). Kolumna przesuwa się tak samo jak dla E15:E19 i F15:F19:

| Kopia j | Kolumna w Strona 3 | Komórka scalona | Wartość → F24 |
|---------|-------------------|-----------------|----------------|
| 0       | Q (17)            | Q17:R17         | Q17            |
| 1       | S (19)            | S17:T17         | S17            |
| 2       | U (21)            | U17:V17         | U21            |
| j       | 17 + j×2          | —               | `ws3.cells(17, 17 + j*2).value` |

Odczyt odbywa się w `wczytaj_wszystko_xlwings` jako osobna pętla po zakończeniu odczytu E/F. Wynik: `f24_per_kopia` — lista wartości (jedna per kopia).

### Krok 15 — Zapis komórek w arkuszu Wyniki

Po wypełnieniu zakładek roboczych (Etapy 3–5), w tej samej sesji xlwings, dla każdej kopii:

| Komórka         | Źródło danych                                     | Uwagi                          |
|-----------------|---------------------------------------------------|--------------------------------|
| `F24`           | `f24_per_kopia[j]`                                | None → pomiń (nie nadpisuj)    |
| `C28`           | `datetime.datetime.now()`                         | data podpisu                   |
| `C32`           | `datetime.datetime.now()`                         | data podpisu                   |
| `E28` (scal. G28) | `PODPISUJACY_1`                                 | imię z konfiguracji            |
| `E32` (scal. G32) | `PODPISUJACY_2`                                 | imię z konfiguracji            |

Jeśli arkusz Wyniki nie istnieje w kopii — sekcja jest pomijana bez błędu.

### Przekazywanie danych

- `wczytaj_wszystko_xlwings` zwraca teraz 4-krotkę: `(dane_s2, dane_zakladek, dane_ef, f24_per_kopia)`
- `_dostosuj_xlwings` przyjmuje dodatkowy parametr `f24_val`
- `utworz_kopie` przyjmuje dodatkowy parametr `f24_per_kopia` i indeksuje go per kopia `j`

---

---

## Etap 7 — Tworzenie świadectw wzorcowania (dokumenty Word)

### Kontekst

Na podstawie każdej wygenerowanej kopii Excel skrypt tworzy odpowiadający jej plik Word (świadectwo wzorcowania). Plik Word powstaje przez skopiowanie szablonu i podmianę placeholderów tekstowych.

### Nowe zmienne konfiguracyjne

```python
SZABLON_WORD     = "xxx_yyy_LA_TH_2026 - tylko temp.docx"  # szablon świadectwa Word
NR_SW_POCZATKOWY = 770  # numer startowy świadectwa wzorcowania (inkrementowany per kopia)
```

Jeśli `SZABLON_WORD` jest pusty lub plik nie istnieje — Etap 7 jest pomijany z ostrzeżeniem.

### Wymagana biblioteka

```python
try:
    from docx import Document as DocxDocument
    _DOCX_OK = True
except ImportError:
    DocxDocument = None
    _DOCX_OK = False
```

Instalacja: `pip install python-docx`

### Placeholdery w szablonie Word

| Placeholder         | Źródło danych                                              |
|---------------------|------------------------------------------------------------|
| `[data]`            | bieżąca data systemowa, format: `"26 maja 2026 r."`       |
| `[nr_sw]`           | `NR_SW_POCZATKOWY + j` (j = indeks kopii, 0-bazowany)     |
| `[nr_zl]`           | `prefiks` z nazwy pliku Excel (część przed `_LA_TH_`), np. `"133"` |
| `[nr_fabr]`         | numer fabryczny z nazwy pliku (ostatni segment po ` - `)  |
| `[wytworca]`        | `rekord["B"]` z Strona 2 (= E5 kopii Excel)               |
| `[typ]`             | `rekord["D"]` z Strona 2 (= E6 kopii Excel)               |
| `[data_wzorcowania]`| daty K4 ze wszystkich zakładek roboczych, połączone        |
| `[Podpis]`          | `PODPISUJACY_2`                                            |

> **Uwaga:** `[typ]` pochodzi z **E6 pierwszej zakładki roboczej** kopii Excel, czyli `rekord["D"]` — NIE z nazwy pliku (nie parsować `_parsuj_nazwe_pliku` dla [typ]).

### Format daty (`[data]`)

Używa słownika `MIESIACE_GEN` z miesiącami w dopełniaczu (pl):
```python
MIESIACE_GEN = {
    1: "stycznia",  2: "lutego",       3: "marca",     4: "kwietnia",
    5: "maja",      6: "czerwca",      7: "lipca",     8: "sierpnia",
    9: "września", 10: "października", 11: "listopada", 12: "grudnia",
}
# Wynik: "26 maja 2026 r."
```

### Nazwa pliku Word

Generowana przez `generuj_nazwe_word(nr_sw, prefiks, rok)`:
```python
f"{nr_sw}_{prefiks}_LA_TH_{rok}.docx"
# Przykład: "770_133_LA_TH_2026.docx"
```

### Tabela kalibracyjna w Word

Szablon Word zawiera tabelę z wierszami:

| `[wartość_odn_1]` | `[zmierzona_1]` | `[poprawka_1]` | `[niepewność_1]` |
|-------------------|-----------------|----------------|-----------------|
| `[wartość_odn_2]` | `[zmierzona_2]` | `[poprawka_2]` | `[niepewność_2]` |
| `[wartość_odn_3]` | `[zmierzona_3]` | `[poprawka_3]` | `[niepewność_3]` |

Dane kalibracyjne czytane z komórek **D246:G246** każdej zakładki roboczej kopii Excel:
- D246 → `wartosc_odn`
- E246 → `zmierzona`
- F246 → `poprawka`
- G246 → `niepewnosc`

> **WAŻNE:** Odczyt D246:G246 musi następować **po** `wb.save()`, aby formuły Excela (D246, F246, G246) były w pełni przeliczone przed odczytem.

Jeśli punktów kalibracji jest więcej niż 3 — wiersz szablonu (indeks 2, z `_3]`) jest kopiowany i placeholder `_3]` zastępowany `_{k+1}]` dla każdego dodatkowego wiersza.

Separator dziesiętny: **przecinek** (polska konwencja). Funkcja `_fmt(v)` zwraca `str(v).replace(".", ",")` dla floatów z częścią ułamkową.

### Zastępowanie placeholderów — `_zastap_tekst_w_dok`

Funkcja iteruje przez:
1. Wszystkie paragrafy dokumentu (`doc.paragraphs`)
2. Wszystkie komórki tabel (`doc.tables`)
3. **Nagłówki i stopki wszystkich sekcji** (`doc.sections` → `section.header`, `section.footer`, `section.even_page_header`, `section.even_page_footer`, `section.first_page_header`, `section.first_page_footer`)

> **WAŻNE:** Bez iteracji po sekcjach — placeholdery w nagłówku (`[data]`, `[nr_sw]`, `[nr_zl]`) NIE są podmieniane.

Używa pomocniczej funkcji `_zastap_w_kontenerze(kontener)` wywoływanej dla `doc` i każdego `hdr_ftr`.

Dla placeholderów z polskimi znakami — próbuje **obie wersje**: z diakrytykami i bez (ASCII fallback), np. `[wartość_odn_1]` i `[wartosc_odn_1]`, `[niepewność_1]` i `[niepewnosc_1]`. Zależy to od tego, jak szablon Word koduje te znaki.

### Przepływ danych — modyfikacja `_dostosuj_xlwings`

```
_dostosuj_xlwings(app, sciezka_pliku, ...) 
    → po wb.save() odczytuje D246:G246
    → zwraca kalibracja = [{wartosc_odn, zmierzona, poprawka, niepewnosc}, ...]
```

Główna pętla (`_dostosuj_xlwings` tworząca kopie Excel) zbiera listy kalibracji per kopia:
```python
dane_kalibracji = []
for j, ...:
    kal = _dostosuj_xlwings(app, sciezka_kopii, ...)
    dane_kalibracji.append(kal)
nazwy = [...]
return nazwy, dane_kalibracji
```

### Nowa funkcja Word (`_dostosuj_xlwings` / w `main` wywołana jako `_dostosuj_xlwings`)

Parametry: `folder, szablon_word, dane_s2, kopie_excel, dane_zakladek, dane_kalibracji, nr_sw_poczatkowy`

Zwraca: lista nazw utworzonych plików Word.

---

## Znane błędy naprawione w trakcie implementacji

### Bug 1 — Nagłówek Word nie podmieniony
**Problem:** `_zastap_tekst_w_dok` iterowała tylko `doc.paragraphs` i `doc.tables` — pomijała nagłówki i stopki sekcji.  
**Naprawa:** Dodano iterację przez `doc.sections` → `section.header/footer/even_page_*/first_page_*`.

### Bug 2 — `[typ]` parsowany z nazwy pliku zamiast z E6
**Problem:** `[typ]` był ustawiany jako 4. element krotki z `_parsuj_nazwe_pliku` (segment po roku w nazwie pliku).  
**Naprawa:** `"[typ]": str(rekord.get("D") or "")` — ta sama wartość co E6 kopii i `[nr_zl]`.

### Bug 3 — `[wartość_odn]`, `[poprawka]`, `[niepewność]` nie były wypełniane
**Problem:** Dane D246, F246, G246 były odczytywane **przed** `wb.save()` — formuły Excela nie były jeszcze przeliczone, więc zwracały `None`. Samo przeniesienie odczytu po `wb.save()` nie wystarczyło — Excel COM nie gwarantuje przeliczenia formuł przed odczytem bez jawnego wywołania.  
**Naprawa:** Przed odczytem D246:G246 wywołać `app.api.CalculateFullRebuild()` (wymusza pełne przeliczenie wszystkich formuł w otwartym skoroszycie), następnie odczytać wartości. Dodano też diagnostyczny print wartości D/E/F/G246 per zakładka, by łatwo zweryfikować co jest czytane.

### Bug 4 — Separator dziesiętny jako kropka zamiast przecinka
**Problem:** `_fmt(v)` zwracało `str(v)` dla floatów z częścią ułamkową → `"1.23"`.  
**Naprawa:** `return str(v).replace(".", ",")` → `"1,23"` (polska konwencja).

---

## To jest Etap 1 + Etap 2 + Etap 3 + Etap 4 + Etap 5 + Etap 6 + Etap 7

Skrypt będzie w przyszłości rozszerzany. Zachowaj kod modularny i czytelny, aby łatwo było dodawać kolejne funkcje.

---

## Aktualizacja 2026-06-02 — tryby uruchamiania i F24

Dodane tryby rozdzielenia aktywnosci (sekcja KONFIGURACJA):

```python
GENERUJ_EXCEL = True
GENERUJ_WORD  = True
```

Zasady:
- `GENERUJ_EXCEL=True,  GENERUJ_WORD=True`  → pelny przebieg (Excel + Word)
- `GENERUJ_EXCEL=True,  GENERUJ_WORD=False` → tylko etapy Excel
- `GENERUJ_EXCEL=False, GENERUJ_WORD=True`  → bez tworzenia kopii Excel; Word powstaje z juz istniejacych kopii
- `GENERUJ_EXCEL=False, GENERUJ_WORD=False` → skrypt konczy prace (brak etapow)

### F24 (arkusz Wyniki) — zrodlo z komorki scalonej

Wartosc do `Wyniki!F24` jest czytana z **komorki scalonej** w wierszu 17 arkusza `Strona 3`:

- protokoly standardowe (`...protokół CC...`):
  - kopia 1: `Q:R17` (wartosc z lewej-gornej komorki obszaru scalonego)
  - kolejne kopie: przesuniecie o 2 kolumny w prawo (`S:T17`, `U:V17`, ...)

- protokoly `CC-04` (`...protokół CC-04...`):
  - kopia 1: `S:T17`
  - kolejne kopie: przesuniecie o 2 kolumny w prawo (`U:V17`, `W:X17`, ...)

Uwaga implementacyjna:
- dla `F24` pobierana jest wartosc z obszaru scalonego (MergeArea top-left),
- bez filtrowania po kolorze.
