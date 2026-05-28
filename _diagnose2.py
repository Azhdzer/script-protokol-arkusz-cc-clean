# -*- coding: utf-8 -*-
"""
Diagnostyka kolumn M i R w protokole przez xlwings.
Sprawdza zarówno wartość jak i formułę komórek.
"""
import os
import xlwings as xw

FOLDER   = os.path.dirname(os.path.abspath(__file__))
PROTOKOL = os.path.join(FOLDER, "133; 148_LA_TH_2026 - protokół CC.xlsx")

app = xw.App(visible=False, add_book=False)
try:
    wb = app.books.open(PROTOKOL)
    ws3 = wb.sheets["Strona 3"]

    print("=== Kolumna M (13) i R (18) w bloku 1 (wiersze 20-24) ===")
    for row in range(20, 25):
        m_val     = ws3.cells(row, 13).value
        m_formula = ws3.cells(row, 13).formula
        r_val     = ws3.cells(row, 18).value
        r_formula = ws3.cells(row, 18).formula
        print(f"  row {row}: M.value={m_val!r}  M.formula={m_formula!r} | "
              f"R.value={r_val!r}  R.formula={r_formula!r}")

    print("\n=== Dla porównania: kolumna L (12) i Q (17) ===")
    for row in range(20, 25):
        l_val     = ws3.cells(row, 12).value
        l_formula = ws3.cells(row, 12).formula
        q_val     = ws3.cells(row, 17).value
        q_formula = ws3.cells(row, 17).formula
        print(f"  row {row}: L.value={l_val!r}  L.formula={l_formula!r} | "
              f"Q.value={q_val!r}  Q.formula={q_formula!r}")

    wb.close()
finally:
    app.quit()

print("\nDone.")
