# -*- coding: utf-8 -*-
"""
Weryfikacja wygenerowanego pliku przez xlwings - sprawdza rzeczywiste wartosci.
"""
import os
import glob
import xlwings as xw

FOLDER = os.path.dirname(os.path.abspath(__file__))

files = sorted(glob.glob(os.path.join(FOLDER, "133_*.xlsx")))
if not files:
    print("Brak plikow 133_*.xlsx")
    exit()

f = files[0]
print(f"Plik: {os.path.basename(f)}")

app = xw.App(visible=False, add_book=False)
try:
    wb = app.books.open(f)
    for ws in wb.sheets:
        print(f"\n  Zakladka: {ws.name}")
        for row in range(15, 20):
            c = ws.cells(row, 3).value
            d = ws.cells(row, 4).value
            e = ws.cells(row, 5).value
            fv = ws.cells(row, 6).value
            print(f"    row {row}: C={c!r}  D={d!r}  E={e!r}  F={fv!r}")
    wb.close()
finally:
    app.quit()

print("\nDone.")
