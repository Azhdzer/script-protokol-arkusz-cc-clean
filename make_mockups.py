# -*- coding: utf-8 -*-
"""
make_mockups.py — generuje statyczne obrazki-makiety (PNG) dla podgladu w GUI.

Dla kazdego dokumentu tworzy dwa stany:
    *_pusto.png    — pusty szablon (tak wyglada PRZED),
    *_pelny.png    — wypelniony wynik (tak wyglada PO wygenerowaniu).

Obrazki sa renderowane w skali 2x (ostre na ekranach HiDPI) do folderu assets/.
Uruchom raz:   .venv\\Scripts\\python.exe make_mockups.py
"""

import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QImage, QPainter, QTextDocument
from PySide6.QtCore import Qt

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
os.makedirs(ASSETS, exist_ok=True)

SCALE = 2
WIDTH = 600

# Paleta jak w Excelu (spojna z arkuszami)
G_HEAD = "#C6E0B4"   # zielony naglowek
G_CELL = "#E2EFDA"   # jasny zielony
REJ    = "#FCA5A5"   # odrzucony punkt
EXT    = "#FFE0B2"   # kolumna zewnetrzna (tylko temperatura)
EMPTY  = "#F8FAFC"   # puste pole
BORDER = "#94A3B8"


def _cell(txt="", bg="#FFFFFF", bold=False, color="#1F2937"):
    w = "bold" if bold else "normal"
    return (f'<td bgcolor="{bg}" align="center" '
            f'style="padding:6px 10px; font-weight:{w}; color:{color};">{txt or "&nbsp;"}</td>')


def _doc_html(title, subtitle, header_html, rows_html):
    return f"""
    <div style="font-family:'Segoe UI'; color:#1F2937;">
      <div style="font-size:15px; font-weight:700; margin-bottom:2px;">{title}</div>
      <div style="font-size:11px; color:#6B7280; margin-bottom:10px;">{subtitle}</div>
      <table cellspacing="0" cellpadding="0" border="1" bordercolor="{BORDER}"
             width="100%" style="font-size:12px;">
        {header_html}
        {rows_html}
      </table>
    </div>
    """


def protokol(pelny: bool):
    # Naglowek: Nastawa | Przyrzad 1 (t, RH) | Przyrzad 2 (t) [zewn.]
    header = (
        "<tr>"
        + _cell("Nastawa", G_HEAD, True)
        + _cell("Przyrzad 1", G_HEAD, True) + _cell("", G_HEAD, True)
        + _cell("Przyrzad 2", EXT, True)
        + "</tr>"
        "<tr>"
        + _cell("t / RH", G_HEAD, True)
        + _cell("t [°C]", G_HEAD, True) + _cell("RH [%]", G_HEAD, True)
        + _cell("t [°C]", EXT, True)
        + "</tr>"
    )
    punkty = [("25°C / 51%", "25.0", "51", "25.1", False),
              ("40°C / 49%", "40.1", "49", "40.2", False),
              ("61°C / 50%", "61.0", "50", "61.1", False),
              ("9.7°C / 76%", "—", "—", "—", True)]  # odrzucony w wersji pelnej
    rows = ""
    for nast, t, rh, t2, rej in punkty:
        if not pelny:
            rows += ("<tr>" + _cell(nast, G_HEAD, True)
                     + _cell("", EMPTY) + _cell("", EMPTY) + _cell("", EMPTY) + "</tr>")
        elif rej:
            rows += ("<tr>" + _cell(nast, G_HEAD, True)
                     + _cell("odrzuc.", REJ) + _cell("", REJ) + _cell("odrzuc.", REJ) + "</tr>")
        else:
            rows += ("<tr>" + _cell(nast, G_HEAD, True)
                     + _cell(t, G_CELL) + _cell(rh, G_CELL) + _cell(t2, EXT) + "</tr>")
    sub = ("Wypelniony arkusz — filtr odrzucil punkt poza nastawa (czerwony)."
           if pelny else "Pusty szablon przed uruchomieniem generatora.")
    return _doc_html("Protokol · Strona 3", sub, header, rows)


def wyniki(pelny: bool):
    header = ("<tr>"
              + _cell("t odniesienia [°C]", G_HEAD, True)
              + _cell("t zmierzona [°C]", G_HEAD, True)
              + _cell("Poprawka Δt [°C]", G_HEAD, True)
              + _cell("Niepewnosc U [°C]", G_HEAD, True)
              + "</tr>")
    dane = [("5", "5.0", "-0.0", "0.05"), ("10", "10.1", "-0.1", "0.05"),
            ("25", "25.0", "0.0", "0.05"), ("40", "40.1", "-0.1", "0.06"),
            ("61", "61.0", "0.0", "0.07")]
    rows = ""
    for t, tm, d, u in dane:
        if pelny:
            rows += ("<tr>" + _cell(t, G_CELL, True) + _cell(tm, "#FFFFFF")
                     + _cell(d, "#FFFFFF") + _cell(u, "#FFFFFF") + "</tr>")
        else:
            rows += ("<tr>" + _cell(t, G_CELL, True) + _cell("", EMPTY)
                     + _cell("", EMPTY) + _cell("", EMPTY) + "</tr>")
    sub = ("Obliczone poprawki i niepewnosci." if pelny
           else "Szablon arkusza Wyniki (przed obliczeniem).")
    return _doc_html("Arkusz Wyniki", sub, header, rows)


def word(pelny: bool):
    header = ("<tr>"
              + _cell("t [°C]", G_HEAD, True) + _cell("t_m [°C]", G_HEAD, True)
              + _cell("Δt [°C]", G_HEAD, True) + _cell("U [°C]", G_HEAD, True) + "</tr>")
    dane = [("10", "10.1", "-0.1", "0.05"), ("25", "25.0", "0.0", "0.05"),
            ("40", "40.1", "-0.1", "0.06"), ("61", "61.0", "0.0", "0.07")]
    rows = ""
    for t, tm, d, u in dane:
        if pelny:
            rows += ("<tr>" + _cell(t, "#FFFFFF", True) + _cell(tm, "#FFFFFF")
                     + _cell(d, "#FFFFFF") + _cell(u, "#FFFFFF") + "</tr>")
        else:
            rows += ("<tr>" + _cell(t, EMPTY, True) + _cell("", EMPTY)
                     + _cell("", EMPTY) + _cell("", EMPTY) + "</tr>")
    sub = ("Tabela w swiadectwie wzorcowania (Word)." if pelny
           else "Szablon swiadectwa (przed wypelnieniem).")
    return _doc_html("Swiadectwo · Word", sub, header, rows)


def render(html, path):
    doc = QTextDocument()
    doc.setHtml(html)
    doc.setTextWidth(WIDTH)
    size = doc.size()
    img = QImage(int(size.width() * SCALE), int(size.height() * SCALE), QImage.Format_ARGB32)
    img.setDevicePixelRatio(SCALE)
    img.fill(Qt.white)
    p = QPainter(img)
    doc.drawContents(p)
    p.end()
    img.save(path)
    print(f"  -> {os.path.relpath(path, HERE)}  ({img.width()}x{img.height()} px)")


def main():
    QApplication([])
    jobs = {
        "protokol_pusto": protokol(False), "protokol_pelny": protokol(True),
        "wyniki_pusto": wyniki(False),     "wyniki_pelny": wyniki(True),
        "word_pusto": word(False),         "word_pelny": word(True),
    }
    print("Generuje makiety do assets/:")
    for name, html in jobs.items():
        render(html, os.path.join(ASSETS, f"{name}.png"))
    print("Gotowe.")


if __name__ == "__main__":
    main()
