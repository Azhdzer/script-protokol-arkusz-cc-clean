# -*- coding: utf-8 -*-
"""
aktualizuj_dokumentacje.py — tabele dokumentacji generowane z kodu.

Sekcje miedzy znacznikami <!-- AUTO:nazwa --> ... <!-- /AUTO:nazwa --> w plikach
HTML z tego folderu sa zastepowane aktualna trescia:

    AUTO:statystyki  — liczba modulow, linii, funkcji, ustawien, testow,
    AUTO:moduly      — tabela modulow (linie, funkcje, rola z docstringu),
    AUTO:ustawienia  — wszystkie ustawienia panelu z rejestru cc_config.py,
    AUTO:funkcje     — spis funkcji kazdego modulu (pierwsza linia docstringu).

Uruchom po kazdej zmianie kodu, ktora dodaje ustawienie, modul albo funkcje:

    .venv\\Scripts\\python.exe DOKUMENTACJA\\aktualizuj_dokumentacje.py
"""

import ast
import datetime
import glob
import html
import os
import re
import sys

TU = os.path.dirname(os.path.abspath(__file__))
KORZEN = os.path.dirname(TU)
sys.path.insert(0, KORZEN)

import cc_config as C  # noqa: E402

PLIKI_HTML = ("instrukcja-uzytkownika.html", "infografika-projektu.html",
              os.path.join("..", "README.md"))
POMIN_MODULY = {"zrob_skrot.py"}


def e(tekst):
    return html.escape(str(tekst if tekst is not None else ""))


def moduly():
    """[(plik, linie, [(nazwa, nr_linii, opis)], opis_modulu)] — pliki .py projektu."""
    wynik = []
    for sciezka in sorted(glob.glob(os.path.join(KORZEN, "*.py"))):
        plik = os.path.basename(sciezka)
        if plik in POMIN_MODULY:
            continue
        zrodlo = open(sciezka, encoding="utf-8").read()
        drzewo = ast.parse(zrodlo)
        funkcje = []
        for wezel in ast.walk(drzewo):
            if isinstance(wezel, (ast.FunctionDef, ast.AsyncFunctionDef)):
                doc = (ast.get_docstring(wezel) or "").strip().splitlines()
                funkcje.append((wezel.name, wezel.lineno, doc[0] if doc else ""))
        funkcje.sort(key=lambda f: f[1])
        doc = [w.strip() for w in (ast.get_docstring(drzewo) or "").strip().splitlines() if w.strip()]
        opis = doc[0] if doc else ""
        opis = re.sub(rf"^{re.escape(plik)}\s*[—-]\s*", "", opis)
        if opis == plik and len(doc) > 1:
            opis = doc[1]
        wynik.append((plik, zrodlo.count("\n") + 1, funkcje, opis))
    return wynik


def liczba_testow():
    n = 0
    for sciezka in glob.glob(os.path.join(KORZEN, "testy", "test_*.py")):
        n += len(re.findall(r"^\s*def test_", open(sciezka, encoding="utf-8").read(), re.M))
    return n, len(glob.glob(os.path.join(KORZEN, "testy", "test_*.py")))


def html_statystyki():
    mod = moduly()
    testy, pliki = liczba_testow()
    kafle = [
        (len(mod), "modułów Python"),
        (f"{sum(m[1] for m in mod):,}".replace(",", " "), "linii kodu"),
        (sum(len(m[2]) for m in mod), "funkcji"),
        (len(C.USTAWIENIA), "ustawień w panelu"),
        (testy, f"testów ({pliki} plików)"),
        (len(C.KROKI) - 1, "kroków obiegu"),
    ]
    return ('<div class="stats">' + "".join(
        f'<div class="stat"><b>{e(w)}</b><span>{e(o)}</span></div>' for w, o in kafle)
        + f'</div><p class="muted small">Stan na {datetime.date.today():%d.%m.%Y} '
          f'(generowane z kodu: <code>aktualizuj_dokumentacje.py</code>).</p>')


def html_moduly():
    w = ['<table><thead><tr><th>Plik</th><th class="r">Linie</th><th class="r">Funkcje</th>'
         '<th>Rola</th></tr></thead><tbody>']
    for plik, linie, funkcje, opis in moduly():
        w.append(f"<tr><td><code>{e(plik)}</code></td><td class='r'>{linie}</td>"
                 f"<td class='r'>{len(funkcje)}</td><td>{e(opis)}</td></tr>")
    w.append("</tbody></table>")
    return "\n".join(w)


def md_moduly():
    w = ["| Plik | Linie | Funkcje | Rola |", "|---|---:|---:|---|"]
    for plik, linie, funkcje, opis in moduly():
        w.append(f"| `{plik}` | {linie} | {len(funkcje)} | {opis.replace('|', '/')} |")
    testy, pliki = liczba_testow()
    w.append(f"| `testy/` | — | {testy} testów | {pliki} plików `test_*.py` (unittest) |")
    return "\n".join(w)


def _domyslna(u):
    d = u.domyslna
    if isinstance(d, bool):
        return "wł." if d else "wył."
    if isinstance(d, list):
        return f"{len(d)} wierszy" if d else "—"
    if hasattr(d, "total_seconds"):
        return f"{int(d.total_seconds() // 60)} min"
    if d in (None, ""):
        return "—"
    d = str(d)
    if os.path.normcase(os.path.abspath(d)) == os.path.normcase(KORZEN):
        return "(folder projektu)"
    if u.typ == "minuty":
        return f"{d} min"
    return d


def html_ustawienia():
    nazwy = {k: n for k, n, _o in C.KROKI}
    w = []
    for krok, nazwa, opis in C.KROKI:
        lista = [u for u in C.USTAWIENIA if u.krok == krok]
        if not lista:
            continue
        w.append(f'<h3 id="ust-{e(krok)}">{e(nazwa)} <span class="muted">— {e(opis)}</span></h3>')
        w.append('<table class="ust"><thead><tr><th>Grupa / ustawienie</th><th>Zmienna</th>'
                 '<th>Domyślnie</th><th>Opis</th></tr></thead><tbody>')
        for u in sorted(lista, key=lambda u: (u.poziom != C.PODSTAWOWY, u.grupa)):
            poziom = ("" if u.poziom == C.PODSTAWOWY
                      else ' <span class="tag z">zaawansowane</span>')
            w.append(f"<tr><td><span class='muted'>{e(u.grupa)}</span><br><b>{e(u.etykieta)}</b>"
                     f"{poziom}</td><td><code>{e(u.env)}</code></td><td>{e(_domyslna(u))}</td>"
                     f"<td>{e(u.opis)}</td></tr>")
        w.append("</tbody></table>")
    _ = nazwy
    return "\n".join(w)


def html_funkcje():
    w = []
    for plik, linie, funkcje, opis in moduly():
        w.append(f"<details><summary><code>{e(plik)}</code> — {len(funkcje)} funkcji, "
                 f"{linie} linii <span class='muted'>· {e(opis)}</span></summary>")
        w.append('<table class="fn"><tbody>')
        for nazwa, nr, doc in funkcje:
            w.append(f"<tr><td class='r muted'>{nr}</td><td><code>{e(nazwa)}</code></td>"
                     f"<td>{e(doc)}</td></tr>")
        w.append("</tbody></table></details>")
    return "\n".join(w)


SEKCJE = {
    "statystyki": html_statystyki,
    "moduly": html_moduly,
    "ustawienia": html_ustawienia,
    "funkcje": html_funkcje,
    "moduly_md": md_moduly,
}


def aktualizuj(sciezka):
    tresc = open(sciezka, encoding="utf-8").read()
    zmienione = []
    for nazwa, funkcja in SEKCJE.items():
        wzor = re.compile(rf"(<!-- AUTO:{nazwa} -->)(.*?)(<!-- /AUTO:{nazwa} -->)", re.S)
        if wzor.search(tresc):
            nowe = funkcja()
            tresc = wzor.sub(lambda m: f"{m.group(1)}\n{nowe}\n{m.group(3)}", tresc)
            zmienione.append(nazwa)
    if sciezka.endswith(".md"):
        testy, _p = liczba_testow()
        tresc, n = re.subn(r"\*\*\d+ testów\*\*", f"**{testy} testów**", tresc)
        if n:
            zmienione.append("liczba testow")
    open(sciezka, "w", encoding="utf-8", newline="\n").write(tresc)
    return zmienione


if __name__ == "__main__":
    for plik in PLIKI_HTML:
        s = os.path.join(TU, plik)
        if os.path.exists(s):
            print(f"{plik}: {', '.join(aktualizuj(s)) or '(brak znacznikow AUTO)'}")
