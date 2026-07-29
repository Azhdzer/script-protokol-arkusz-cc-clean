# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — jeden plik .exe (ProtokolCC.exe) laczacy panel GUI oraz
oba workery (generuj_arkusze / generuj_obserwacje). Tryb wybiera app_entry.py
na podstawie zmiennej CC_WORKER.

Budowanie:   .venv\\Scripts\\pyinstaller ProtokolCC.spec --noconfirm --clean
Wynik:       dist\\ProtokolCC.exe
"""

from PyInstaller.utils.hooks import collect_all, collect_submodules

import os as _os

# Makiety podgladu (przed/po) oraz ikona — dolaczane do .exe.
datas = [("assets", "assets")]
if _os.path.exists("app.ico"):
    datas.append(("app.ico", "."))
binaries = []
hiddenimports = []

# xlwings + COM (Excel przez pywin32) — kompletny zestaw, inaczej worker pada.
for pkg in ("xlwings",):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hiddenimports += h

# python-docx potrzebuje swojego szablonu default.docx (import: "docx").
for pkg in ("docx",):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hiddenimports += h

hiddenimports += collect_submodules("openpyxl")
hiddenimports += [
    "win32com", "win32com.client", "pythoncom", "pywintypes",
    "win32api", "win32con", "win32gui", "win32process",
    # moduly aplikacji zaladowane leniwie w app_entry:
    "app_gui", "generuj_arkusze", "generuj_obserwacje",
]

excludes = ["tkinter", "unittest", "pydoc", "test", "pytest", "PyQt5", "PyQt6"]

a = Analysis(
    ["app_entry.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ProtokolCC",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX bywa wykrywany przez antywirusy — wylaczone.
    runtime_tmpdir=None,
    console=False,             # okno GUI bez konsoli; workery pisza do potoku QProcess.
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=["app.ico"] if __import__("os").path.exists("app.ico") else None,
)
