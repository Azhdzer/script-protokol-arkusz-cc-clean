# -*- coding: utf-8 -*-
"""
cc_widok.py — jedno miejsce, w ktorym ustalamy WYGLAD OKNA gotowego pliku .xlsx.

Problem, ktory to rozwiazuje
----------------------------
Po otwarciu protokolu albo kopii arkusza obliczeniowego pasek zakladek na dole
bywal scisniety do jednej zakladki: zeby zobaczyc pozostale strony, trzeba bylo
recznie przeciagnac w lewo suwak miedzy paskiem zakladek a paskiem przewijania.

Odpowiadaja za to trzy atrybuty w 'xl/workbook.xml':

    tabRatio    — ile miejsca (0..1000) dostaje pasek zakladek kosztem paska
                  przewijania. Brak atrybutu = 600, czyli 60%.
    firstSheet  — od ktorej zakladki zaczyna sie widoczny pasek. 6 oznacza
                  „pokaz od siodmej", czyli reszta jest schowana z lewej.
    activeTab   — ktora zakladka jest aktywna po otwarciu.

Psuly sie one na dwa sposoby:

  1. Krok 3 otwiera protokol przez Excel COM (wpisanie F/G do Strony 3) i przy
     zapisie Excel przepisuje <workbookView> po swojemu — atrybut tabRatio
     znikal calkowicie.
  2. W kopiach ustawienie szlo przez zywe okno Excela
     (Windows(1).ScrollWorkbookTabs). To bywalo nieskuteczne: w tym samym
     przebiegu czesc kopii dostawala firstSheet=0, a czesc firstSheet=6.

Dlatego ustawiamy to NA GOTOWYM PLIKU, a nie na oknie Excela: plik .xlsx to
archiwum ZIP, wiec podmieniamy w nim jedna czesc — 'xl/workbook.xml'. Reszta
pliku zostaje bajt w bajt, razem z zapamietanymi wynikami formul (to wazne:
ponowny zapis przez openpyxl skasowalby ten cache, a swiadectwa Word czytaja
wlasnie policzone wartosci).

Zasada: KAZDY plik .xlsx, ktory oddajemy uzytkownikowi, przechodzi przez
`wymus_widok()` jako ostatni krok — juz po wszystkich zapisach, takze tych
przez COM. Jedno wywolanie na plik, zawsze na koncu.
"""

import os
import re
import shutil
import zipfile

CZESC = "xl/workbook.xml"

# 0.85 = pasek zakladek dostaje 85% szerokosci. Przy 7 zakladkach kopii i 3
# stronach protokolu wszystkie mieszcza sie bez przewijania.
TAB_RATIO_DOMYSLNY = 0.85

_RE_WORKBOOKVIEW = re.compile(r"<workbookView\b[^>]*?/?>")
_RE_BOOKVIEWS = re.compile(r"<bookViews\b")
_RE_SHEETS = re.compile(r"<sheets\b")


def _na_promile(tab_ratio):
    """
    Ustawienie panelu jest ulamkiem (0.85); w pliku stoi liczba 0..1000.

    Panel dopuszcza tylko 0.1..1.0, ale plik ustawien mozna poprawic recznie,
    wiec przyjmujemy takze zapis procentowy (85) i wprost w promilach (850).
    """
    try:
        wartosc = float(tab_ratio)
    except (TypeError, ValueError):
        wartosc = TAB_RATIO_DOMYSLNY
    if wartosc <= 1.0:                       # 0.85  -> 850
        wartosc *= 1000.0
    elif wartosc <= 100.0:                   # 85    -> 850
        wartosc *= 10.0
    return max(0, min(1000, int(round(wartosc))))


def _ustaw_atrybut(tag, nazwa, wartosc):
    """Wstawia albo nadpisuje atrybut w pojedynczym tagu XML."""
    wzorzec = re.compile(r'\s%s="[^"]*"' % re.escape(nazwa))
    if wzorzec.search(tag):
        return wzorzec.sub(' %s="%s"' % (nazwa, wartosc), tag, count=1)
    koniec = "/>" if tag.endswith("/>") else ">"
    return tag[:-len(koniec)] + ' %s="%s"' % (nazwa, wartosc) + koniec


def popraw_workbook_xml(xml, tab_ratio=TAB_RATIO_DOMYSLNY, pierwszy=0, aktywny=None):
    """
    Zwraca tresc 'xl/workbook.xml' z wymuszonym widokiem paska zakladek.

    `aktywny=None` zostawia aktywna zakladke bez zmian — czasem chcemy otworzyc
    plik na arkuszu zbiorczym, a nie na pierwszym.
    """
    tag = _RE_WORKBOOKVIEW.search(xml)
    if tag is None:
        # Plik bez <bookViews> w ogole (tak zapisuje Excel po niektorych edycjach)
        # — dokladamy caly blok tuz przed lista arkuszy.
        nowy = '<bookViews><workbookView tabRatio="%d" firstSheet="%d"%s/></bookViews>' % (
            _na_promile(tab_ratio), int(pierwszy),
            '' if aktywny is None else ' activeTab="%d"' % int(aktywny))
        m = _RE_BOOKVIEWS.search(xml) or _RE_SHEETS.search(xml)
        if m is None:
            return xml                        # nieznany uklad — nie ryzykujemy
        return xml[:m.start()] + nowy + xml[m.start():]

    stary = tag.group(0)
    poprawiony = _ustaw_atrybut(stary, "tabRatio", _na_promile(tab_ratio))
    poprawiony = _ustaw_atrybut(poprawiony, "firstSheet", int(pierwszy))
    if aktywny is not None:
        poprawiony = _ustaw_atrybut(poprawiony, "activeTab", int(aktywny))
    return xml[:tag.start()] + poprawiony + xml[tag.end():]


def wymus_widok(sciezka, tab_ratio=TAB_RATIO_DOMYSLNY, pierwszy=0, aktywny=None):
    """
    Ustawia widok paska zakladek w gotowym pliku .xlsx.

    Zwraca True, gdy plik zostal poprawiony. Bledy sa wylapywane — zly widok
    paska nie moze przerwac generowania dokumentow.
    """
    if not sciezka or not os.path.isfile(sciezka):
        return False
    tymczasowy = sciezka + ".widok.tmp"
    try:
        with zipfile.ZipFile(sciezka) as zrodlo:
            nazwy = zrodlo.namelist()
            if CZESC not in nazwy:
                return False
            xml = zrodlo.read(CZESC).decode("utf-8", "replace")
            nowy_xml = popraw_workbook_xml(xml, tab_ratio, pierwszy, aktywny)
            if nowy_xml == xml:
                return False
            with zipfile.ZipFile(tymczasowy, "w", zipfile.ZIP_DEFLATED) as cel:
                for info in zrodlo.infolist():
                    dane = (nowy_xml.encode("utf-8") if info.filename == CZESC
                            else zrodlo.read(info.filename))
                    cel.writestr(info, dane)
        shutil.move(tymczasowy, sciezka)
        return True
    except Exception as blad:
        print(f"    [UWAGA] Nie udalo sie ustawic widoku zakladek "
              f"({os.path.basename(sciezka)}): {blad}")
        return False
    finally:
        if os.path.exists(tymczasowy):
            try:
                os.remove(tymczasowy)
            except OSError:
                pass
