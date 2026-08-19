# Testy

Bez zadnych dodatkowych zaleznosci — czysty `unittest` z biblioteki standardowej.

## Uruchomienie

```powershell
# Wszystko (z krokiem 3 — uruchamia Excel, ok. 3 minuty):
.venv\Scripts\python.exe -m unittest discover -s testy -t testy

# Szybko, bez Excela (ok. 1 minuty):
$env:CC_TESTY_SZYBKIE=1; .venv\Scripts\python.exe -m unittest discover -s testy -t testy

# Pojedynczy plik:
.venv\Scripts\python.exe -m unittest discover -s testy -t testy -p test_kontrakt.py
```

## Co gdzie

| Plik | Czego pilnuje |
|---|---|
| `test_config.py` | Rejestr ustawien: spojnosc definicji, konwersja do zmiennych srodowiskowych, zapis/odczyt `cc_ustawienia.json`, odpornosc helperow na smieciowe wartosci. |
| `test_kontrakt.py` | **Najwazniejszy.** Kazde ustawienie panelu naprawde dociera do skryptu i zmienia wlasciwa stala — a bez panelu skrypt ma wartosc domyslna. |
| `test_widgety.py` | Pola formularza (kazdy typ), lista plikow TXT, kolorowanie logu, wykrywanie utworzonych plikow, lista kontrolna swiezosci. |
| `test_panel.py` | Budowa okna, zbieranie ustawien, walidacja przed startem, przerwanie obiegu przy bledzie. |
| `test_obieg.py` | Pelny obieg 1 → 2 → 3 na prawdziwym pomiarze 188. Sprawdza gotowe dokumenty, nie tylko kod wyjscia. |

Plik `_podglad.py` nie jest testem — renderuje okno do PNG, gdy chcesz obejrzec
uklad bez klikania po aplikacji (`python testy\_podglad.py 2`).

## Zasada bezpieczenstwa

Zaden test nie zapisuje niczego w folderze projektu. Wszystko dzieje sie w
`testy/_piaskownica/<nazwa>/`, ktora dostaje **kopie** prawdziwych plikow
wejsciowych — testy pracuja na realnych danych, ale nie moga uszkodzic
protokolow ani swiadectw.

Folder `_piaskownica/` jest w `.gitignore` i mozna go skasowac w kazdej chwili.

## Dlaczego `test_kontrakt.py` jest najwazniejszy

Poprzedni panel psul sie dokladnie w jednym miejscu: pokazywal opcje, ktorych
skrypty nigdy nie czytaly. Uzytkownik "ustawial" cos, co nie mialo zadnego
wplywu na wynik wzorcowania — i dowiadywal sie o tym dopiero po otwarciu
gotowego swiadectwa.

Ten plik nie pozwala temu wrocic. Dla kazdego z ustawien importuje modul
w osobnym procesie dwa razy: raz z czysta lista zmiennych (ma wyjsc wartosc
domyslna) i raz z ustawiona zmienna (ma wyjsc nowa wartosc). Dodatkowo test
`test_kazde_ustawienie_jest_sprawdzane` pilnuje, zeby nowe ustawienie dodane
do rejestru nie moglo zostac bez podlaczenia — suite sie wywali, dopoki nie
dopiszesz go do tabeli `KONTRAKT`.

## Dodajesz nowe ustawienie?

1. Dopisz `_U(...)` w `cc_config.py`.
2. Odczytaj je w skrypcie przez `C.tekst(...)` / `C.liczba(...)` / `C.flaga(...)`.
3. Dopisz wiersz do `KONTRAKT` w `test_kontrakt.py`.

Krok 3 nie jest opcjonalny — bez niego `test_kazde_ustawienie_jest_sprawdzane`
zglosi brak pokrycia.
