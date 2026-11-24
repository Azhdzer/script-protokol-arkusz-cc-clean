# -*- coding: utf-8 -*-
"""
protokol_z_obserwacji.py — KROK 2b: protokol z poprawionej obserwacji.

Gdy krok 2 wybral punkty nie tak, jak trzeba (np. punkt na granicy kryterium, inna
nastawa niz zamowiona), operator poprawia wybor w arkuszu obserwacji: 5 wierszy
reprezentantow punktu oznacza wypelnieniem i pogrubieniem, a zbedny wybor skryptu
zdejmuje. Ten krok buduje protokol od nowa z TYCH punktow — razem z danymi
przyrzadow z plikow wynikow i oznaczeniami w wynikach i zestawieniu.

Pliki TXT — te same co w kroku 2 (lista w panelu). Logika: generuj_obserwacje.main_z_obserwacji.
"""

import sys

import generuj_obserwacje


def main():
    return generuj_obserwacje.main_z_obserwacji()


if __name__ == "__main__":
    sys.exit(main())
