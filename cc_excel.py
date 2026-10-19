# -*- coding: utf-8 -*-
"""
cc_excel.py — wlasne instancje Excela: uruchomienie i PEWNE zamkniecie.

Zasada: skrypt zamyka WYLACZNIE Excela, ktorego sam uruchomil. Excel otwarty
przez uzytkownika (przed skryptem albo w trakcie) jest nietykalny.

Skad ta zasada
--------------
* win32com.client.Dispatch('Excel.Application') NIE tworzy nowego Excela —
  PODPINA SIE pod juz otwarty. Krok 2 robil tak przy usuwaniu blokow protokolu:
  ukrywal Excela uzytkownika (Visible=False), wylaczal pytania
  (DisplayAlerts=False) i na koniec wolal Quit() — zamykajac jego okna,
  a niezapisane zmiany przepadaly bez pytania. Dlatego zawsze DispatchEx
  (osobny proces), nigdy Dispatch.

* Duzy plik potrafi zawiesic Excela: Quit() nie dochodzi do skutku i w tle
  zostaje proces EXCEL.EXE, ktory blokuje pliki i kolejne przebiegi. Zamiast
  zamykac wszystkie Excele, ubijamy TYLKO proces o PID naszej instancji — i to
  dopiero wtedy, gdy po Quit() dalej zyje.
"""

import ctypes
import os
import time
from ctypes import wintypes

_SYNCHRONIZE = 0x00100000
_PROCESS_TERMINATE = 0x0001
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_STILL_ACTIVE = 259

CZEKAJ_NA_ZAMKNIECIE_S = 10.0


def _kernel32():
    k = ctypes.WinDLL("kernel32", use_last_error=True)
    k.OpenProcess.restype = wintypes.HANDLE
    k.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    k.CloseHandle.argtypes = (wintypes.HANDLE,)
    k.GetExitCodeProcess.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
    k.TerminateProcess.argtypes = (wintypes.HANDLE, wintypes.UINT)
    k.QueryFullProcessImageNameW.argtypes = (
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD))
    return k


def nazwa_procesu(pid):
    """Nazwa pliku procesu ('EXCEL.EXE') albo None, gdy procesu nie ma."""
    if not pid or os.name != "nt":
        return None
    k = _kernel32()
    h = k.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not h:
        return None
    try:
        bufor = ctypes.create_unicode_buffer(1024)
        rozmiar = wintypes.DWORD(len(bufor))
        if not k.QueryFullProcessImageNameW(h, 0, bufor, ctypes.byref(rozmiar)):
            return None
        return os.path.basename(bufor.value)
    finally:
        k.CloseHandle(h)


def proces_zyje(pid):
    if not pid or os.name != "nt":
        return False
    k = _kernel32()
    h = k.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION | _SYNCHRONIZE, False, int(pid))
    if not h:
        return False
    try:
        kod = wintypes.DWORD()
        if not k.GetExitCodeProcess(h, ctypes.byref(kod)):
            return False
        return kod.value == _STILL_ACTIVE
    finally:
        k.CloseHandle(h)


def ubij_excela(pid):
    """
    Konczy proces o podanym PID — ale TYLKO gdy to naprawde EXCEL.EXE. Chroni to
    przed sytuacja, w ktorej Windows nadal ten sam numer innemu programowi.
    """
    if not proces_zyje(pid):
        return False
    if (nazwa_procesu(pid) or "").lower() != "excel.exe":
        return False
    k = _kernel32()
    h = k.OpenProcess(_PROCESS_TERMINATE, False, int(pid))
    if not h:
        return False
    try:
        return bool(k.TerminateProcess(h, 1))
    finally:
        k.CloseHandle(h)


def _czekaj_na_koniec(pid, czekaj_s):
    koniec = time.time() + czekaj_s
    while time.time() < koniec:
        if not proces_zyje(pid):
            return True
        time.sleep(0.2)
    return not proces_zyje(pid)


def dopilnuj_zamkniecia(pid, czekaj_s=CZEKAJ_NA_ZAMKNIECIE_S, opis="Excel"):
    """
    Po Quit(): gdy NASZ proces Excela dalej zyje po `czekaj_s` — ubija go.
    Zwraca True, gdy proces jest zamkniety.
    """
    if not pid:
        return True
    if _czekaj_na_koniec(pid, czekaj_s):
        return True
    if ubij_excela(pid):
        print(f"    [Excel] {opis} (PID {pid}) nie zamknal sie po {czekaj_s:.0f} s — "
              f"zakonczony. Inne okna Excela nie zostaly ruszone.")
        return _czekaj_na_koniec(pid, 3.0)
    return False


# ── win32com ─────────────────────────────────────────────────────────────────

def nowy_excel_com():
    """
    NOWY, osobny proces Excela (DispatchEx). Nigdy Dispatch — ten podpina sie pod
    Excela uzytkownika i zamkniecie go zamykaloby jego okna.
    """
    import win32com.client
    return win32com.client.DispatchEx("Excel.Application")


def pid_excela_com(xl):
    try:
        import win32process
        _watek, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
        return pid or None
    except Exception:
        return None


def zamknij_excel_com(xl, wb=None, czekaj_s=CZEKAJ_NA_ZAMKNIECIE_S):
    """
    Zamyka skoroszyt i NASZA instancje; zawieszona jest konczona po PID.

    Pilnowanie idzie w TLE. Excel uruchomiony przez win32com nie konczy sie od
    razu po Quit(): czeka, az Python zwolni odwolania do skoroszytu i arkusza,
    a te trzyma jeszcze funkcja, ktora nas wywolala. Czekanie tutaj kosztowaloby
    pelne `czekaj_s` przy kazdym zamknieciu — w tle proces konczy sie sam chwile
    po powrocie z funkcji, a ubijany jest tylko wtedy, gdy naprawde wisi.
    Watek NIE jest demonem: przy wyjsciu skryptu interpreter na niego zaczeka,
    wiec zawieszony Excel nie zostanie w tle.
    """
    if xl is None:
        return
    pid = pid_excela_com(xl)
    if wb is not None:
        try:
            wb.Close(False)
        except Exception:
            pass
    try:
        xl.Quit()
    except Exception:
        pass
    if pid:
        import threading
        threading.Thread(target=dopilnuj_zamkniecia, args=(pid, czekaj_s),
                         name=f"pilnuj-excel-{pid}", daemon=False).start()


# ── xlwings ──────────────────────────────────────────────────────────────────

def pid_excela_xlwings(app):
    try:
        return int(app.pid)
    except Exception:
        return None


def zamknij_excel_xlwings(app, czekaj_s=CZEKAJ_NA_ZAMKNIECIE_S):
    """app.quit() i — gdy proces dalej zyje (zawieszony) — zakonczenie po PID."""
    if app is None:
        return
    pid = pid_excela_xlwings(app)
    try:
        app.quit()
    except Exception as e:
        print(f"    [Excel] app.quit() nie powiodlo sie (Excel prawdopodobnie juz padl): "
              f"{type(e).__name__}: {e}")
    dopilnuj_zamkniecia(pid, czekaj_s)
