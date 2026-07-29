# -*- coding: utf-8 -*-
"""
app_entry.py — punkt wejscia zamrozonej aplikacji (PyInstaller, jeden .exe).

Jeden plik wykonywalny pelni trzy role. O trybie decyduje zmienna srodowiskowa
CC_WORKER, ktora ustawia panel GUI przed uruchomieniem podprocesu:

    CC_WORKER=arkusze     -> generuj_arkusze.main()
    CC_WORKER=obserwacje  -> generuj_obserwacje.main()
    (brak zmiennej)       -> uruchom panel GUI (app_gui.main())

Dzieki temu GUI moze uruchamiac ten SAM exe jako worker, bez potrzeby
python.exe ani plikow .py na maszynie docelowej.
"""

import os
import sys
import traceback


def _log_crash(tag: str, exc: BaseException) -> None:
    """Zapisz pelny traceback obok .exe — zamrozony worker nie ma stdout w GUI-mode."""
    try:
        base = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__)
        with open(os.path.join(base, "crash.log"), "a", encoding="utf-8") as f:
            f.write(f"\n=== {tag} ===\n")
            f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    except Exception:
        pass


def _run_worker(mode: str) -> int:
    if mode == "arkusze":
        import generuj_arkusze as m
    elif mode == "obserwacje":
        import generuj_obserwacje as m
    else:
        print(f"[app_entry] Nieznany tryb workera: {mode!r}", file=sys.stderr)
        return 2
    m.main()
    return 0


def main() -> int:
    mode = os.environ.get("CC_WORKER", "").strip().lower()

    if mode:
        # Tryb workera (podproces uruchomiony przez GUI).
        # Windowed exe (--noconsole) potrafi miec sys.stdout == None. GUI czyta
        # stdout przez potok QProcess, wiec upewniamy sie, ze jest zapisywalny.
        for _name in ("stdout", "stderr"):
            if getattr(sys, _name, None) is None:
                try:
                    setattr(sys, _name, open(os.devnull if os.name != "nt" else "CONOUT$",
                                             "w", encoding="utf-8", errors="replace"))
                except Exception:
                    setattr(sys, _name, open(os.devnull, "w", encoding="utf-8"))
            else:
                try:
                    getattr(sys, _name).reconfigure(encoding="utf-8", errors="replace")
                except Exception:
                    pass
        try:
            return _run_worker(mode)
        except SystemExit as e:
            return int(e.code or 0)
        except BaseException as e:  # noqa: BLE001 — worker ma zawsze zwrocic kod bledu
            _log_crash(f"worker:{mode}", e)
            traceback.print_exc()
            return 1

    # Tryb GUI.
    try:
        import app_gui
        app_gui.main()
        return 0
    except BaseException as e:  # noqa: BLE001
        _log_crash("gui", e)
        raise


if __name__ == "__main__":
    sys.exit(main())
