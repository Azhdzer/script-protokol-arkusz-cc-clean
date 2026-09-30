# -*- coding: utf-8 -*-
"""
Testy parsowania pozycji 'Obiekty wzorcowania' z PZ.

PZ przychodzi w kilku ukladach. Trzeci z nich (PZ 197) wychodzil na jaw dopiero
w praktyce: gdy na te same punkty idzie DUZO przyrzadow, etykieta 'nr fabr.:'
stoi RAZ — na koncu naglowka pozycji — a numery sa dopiero w podpunktach:

    Termohigrometr (rejestrator, 9 szt.) typ: testo 174H, nr fabr.:
      • 83623973, nr wew.: UR00045;
      • 83617608, nr wew.: UR00052;
      ...
    wytworca: Testo.

Parser szukal 'nr fabr.:' wewnatrz podpunktu, wiec wszystkie 9 przyrzadow
dostawalo PUSTY numer fabryczny, a ostatni jeszcze nr ewidencyjny sklejony
z ogonem 'UR00044; wytworca: Testo'.

Testy pracuja na fragmentach tekstu, a nie na plikach PDF — PZ zawiera dane
zleceniodawcy i nie trafia do repozytorium.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pz_dane


def parsuj(tekst):
    return pz_dane._parsuj_wpis(tekst, "197")


class TestEtykietaWNaglowku(unittest.TestCase):
    """Uklad z PZ 197 — 'nr fabr.:' raz, numery w podpunktach."""

    WPIS = ("Termohigrometr (rejestrator, 9 szt.) typ: testo 174H, nr fabr.:\n"
            "• 83623973, nr wew.: UR00045;\n"
            "• 83617608, nr wew.: UR00052;\n"
            "• 83677891, nr wew.: UR00044;\n"
            "wytwórca: Testo.")

    def setUp(self):
        self.przyrzady = parsuj(self.WPIS)

    def test_wszystkie_przyrzady_rozpoznane(self):
        self.assertEqual(len(self.przyrzady), 3)

    def test_numery_fabryczne_wypelnione(self):
        self.assertEqual([p.nr_fabr for p in self.przyrzady],
                         ["83623973", "83617608", "83677891"])

    def test_numery_ewidencyjne_wypelnione(self):
        self.assertEqual([p.nr_ewid for p in self.przyrzady],
                         ["UR00045", "UR00052", "UR00044"])

    def test_ostatni_nie_zbiera_ogona_z_wytworca(self):
        """Regresja: 'UR00044; wytworca: Testo' zamiast samego 'UR00044'."""
        ostatni = self.przyrzady[-1]
        self.assertNotIn(";", ostatni.nr_ewid)
        self.assertNotIn("wytw", ostatni.nr_ewid.lower())

    def test_typ_i_wytworca_dziedziczone_z_naglowka(self):
        for p in self.przyrzady:
            with self.subTest(fabr=p.nr_fabr):
                self.assertEqual(p.typ, "testo 174H")
                self.assertEqual(p.wytworca, "Testo")


class TestEtykietaWKazdymPodpunkcie(unittest.TestCase):
    """Uklad A — kazdy podpunkt ma wlasna etykiete 'nr fabr.:'."""

    WPIS = ("Termometr (rejestrator, 2 szt.) typ: testo 175T2, "
            "• nr fabr.: 40118669, nr ewid.: Q/LOG/36, "
            "• nr fabr.: 40118614, nr ewid.: Q/LOG/37, "
            "wytwórca: Testo.")

    def test_nadal_dziala(self):
        p = parsuj(self.WPIS)
        self.assertEqual([x.nr_fabr for x in p], ["40118669", "40118614"])
        self.assertEqual([x.nr_ewid for x in p], ["Q/LOG/36", "Q/LOG/37"])


class TestJednaLinia(unittest.TestCase):
    """Uklad B — jeden przyrzad w jednej linii."""

    def test_nadal_dziala(self):
        p = parsuj("typ: M1, nr fabr.: TMM160500502, nr ewid.: Q/LOG/19, "
                   "wytwórca: Tempmate.")
        self.assertEqual(len(p), 1)
        self.assertEqual(p[0].nr_fabr, "TMM160500502")
        self.assertEqual(p[0].nr_ewid, "Q/LOG/19")
        self.assertEqual(p[0].wytworca, "Tempmate")

    def test_wariant_z_nr_wewn(self):
        p = parsuj("typ: 174H, nr fabr.: 123456, nr wewn.: CL-1318A, wytwórca: Testo.")
        self.assertEqual(p[0].nr_fabr, "123456")
        self.assertEqual(p[0].nr_ewid, "CL-1318A")


class TestOdpornoscNaFalszywePozytywy(unittest.TestCase):
    """
    Rozpoznanie numeru „z poczatku podpunktu" wolno wlaczac WYLACZNIE wtedy,
    gdy naglowek konczy sie sama etykieta — inaczej z opisow robilyby sie
    fikcyjne numery fabryczne.
    """

    def test_bez_etykiety_w_naglowku_nie_zgaduje(self):
        wpis = ("Termohigrometr (2 szt.) typ: testo 174H, "
                "• nr ewid.: UR00045, "
                "• nr ewid.: UR00052, "
                "wytwórca: Testo.")
        p = parsuj(wpis)
        self.assertEqual([x.nr_fabr for x in p], ["", ""])
        self.assertEqual([x.nr_ewid for x in p], ["UR00045", "UR00052"])

    def test_srednik_konczy_numer_ewidencyjny(self):
        p = parsuj("typ: X, nr fabr.: 111, nr wew.: AB-1; wytwórca: Testo.")
        self.assertEqual(p[0].nr_ewid, "AB-1")


class TestPunktyMieszane(unittest.TestCase):
    """
    Sekcja 'Zakres wzorcowania' laczy oba zapisy punktow: liste samych temperatur
    i punkty z wilgotnoscia. Zgloszenie z PZ 197: w protokole zabraklo punktow
    -20, 0 i 40 °C.

    Przyczyna: gdy trafil sie choc jeden punkt z wilgotnoscia, parser konczyl
    prace i listy '(-20; 0; 40) °C' juz nie szukal. Punkty do protokolu wybiera
    sie wg PZ, wiec te trzy po prostu z niego znikaly.
    """

    FRAGMENT = ("(-20; 0; 40) °C, (25 °C, 30 %rh); (25 °C, 60 %rh); "
                "(25 °C, 85 %rh); (25 °C, 60 %rh)")

    def punkty(self, frag=None):
        return pz_dane._punkty_z_fragmentu(frag if frag is not None else self.FRAGMENT)

    def test_wszystkie_punkty_rozpoznane(self):
        self.assertEqual(len(self.punkty()), 7)

    def test_punkty_samej_temperatury_nie_gina(self):
        tylko_temp = [t for t, rh in self.punkty() if rh is None]
        self.assertEqual(tylko_temp, [-20.0, 0.0, 40.0])

    def test_punkty_z_wilgotnoscia_zachowane(self):
        z_rh = [(t, rh) for t, rh in self.punkty() if rh is not None]
        self.assertEqual(z_rh, [(25.0, 30.0), (25.0, 60.0), (25.0, 85.0), (25.0, 60.0)])

    def test_kolejnosc_jak_w_zamowieniu(self):
        """Lista temperatur stoi w PZ pierwsza — ma byc pierwsza takze u nas."""
        self.assertEqual(self.punkty()[0], (-20.0, None))
        self.assertEqual(self.punkty()[3], (25.0, 30.0))

    def test_powtorzony_punkt_histerezy_zostaje(self):
        """Drugi raz 60 %rh to osobny punkt — nie wolno go scalic."""
        self.assertEqual(sum(1 for t, rh in self.punkty() if (t, rh) == (25.0, 60.0)), 2)

    def test_sama_lista_temperatur(self):
        self.assertEqual(self.punkty("(0; 10; 20) °C"),
                         [(0.0, None), (10.0, None), (20.0, None)])

    def test_same_punkty_z_wilgotnoscia(self):
        self.assertEqual(self.punkty("(25 °C, 30 %rh); (25 °C, 60 %rh)"),
                         [(25.0, 30.0), (25.0, 60.0)])

    def test_kilka_list_temperatur(self):
        self.assertEqual(self.punkty("(0; 10) °C oraz (30; 40) °C"),
                         [(0.0, None), (10.0, None), (30.0, None), (40.0, None)])

    def test_brak_punktow(self):
        self.assertEqual(self.punkty("brak danych o punktach"), [])


class TestNaglowekLiczbaPojedyncza(unittest.TestCase):
    """
    Zgloszenie z PZ 191: Strona 2 protokolu zostawala PUSTA.

    Gdy zlecenie obejmuje JEDEN przyrzad, PLUM pisze naglowek w liczbie
    pojedynczej — 'Obiekt wzorcowania:'. Parser wymagal liczby mnogiej, wiec
    cala sekcja nie byla znajdowana i lista przyrzadow wychodzila pusta.
    """

    OPIS = ("Termohigrometr zlozony ze wskaznika (rejestratora) typ: testo 176H1, "
            "nr fabr.: 40807872 oraz czujnika temperatury i wilgotnosci wzglednej "
            "typ: 0636 9735, nr fabr.: 21201805 (kanal pomiarowy nr: 1), "
            "wytworca: Testo.")

    METODA = ("Metoda wzorcowania: Metoda porownawcza w komorze klimatycznej, "
              "zgodnie z instrukcja ILAJ 5.4/11.")

    def przyrzady(self, naglowek):
        tekst = f"""{naglowek}
{self.OPIS}
{self.METODA}"""
        return pz_dane.parsuj_tekst(tekst)

    def test_liczba_pojedyncza_jest_rozpoznawana(self):
        self.assertEqual(len(self.przyrzady("Obiekt wzorcowania:")), 1)

    def test_liczba_mnoga_nadal_dziala(self):
        self.assertEqual(len(self.przyrzady("Obiekty wzorcowania:")), 1)

    def test_obiekt_i_czujnik_maja_swoje_dane(self):
        p = self.przyrzady("Obiekt wzorcowania:")[0]
        self.assertEqual((p.wytworca, p.typ, p.nr_fabr),
                         ("Testo", "testo 176H1", "40807872"))
        self.assertEqual((p.czuj_typ, p.czuj_nr_fabr), ("0636 9735", "21201805"))

    def test_dopisek_w_nawiasie_nie_wchodzi_do_numeru(self):
        """'21201805 (kanal pomiarowy nr: 1)' -> sam numer."""
        p = self.przyrzady("Obiekt wzorcowania:")[0]
        self.assertNotIn("(", p.czuj_nr_fabr)
        self.assertNotIn("kanal", p.czuj_nr_fabr.lower())


class TestWskaznikDopisanyZaObiektem(unittest.TestCase):
    """
    Zgloszenie z PZ 207: przyrzad zlozony, w ktorym panel odczytowy ma WLASNY
    typ i numer fabryczny, a w tekscie stoi ZA termohigrometrem:

        Termohigrometr typ: LB-701, nr fabr .: 2783 z panelem odczytowym
        (rejestratorem) typ: LB-706B, nr fabr.: 568, nr ewid.: DNW/PP/002/WS,
        wytworca: LAB-EL.

    Strona 2 protokolu wychodzila z tego bledna: typ brany byl z przodu
    (LB-701), numer fabryczny z konca (568), a kolumny czujnika zostawaly puste.
    W protokole OBIEKTEM jest panel odczytowy, a termohigrometr z przodu jest
    CZUJNIKIEM POMIAROWYM — kolejnosc odwrotna niz w tekscie PZ.

    Osobno: warstwa tekstowa tego PDF daje 'nr fabr .:' (odstep przed kropka).
    Wzorzec tego nie obejmowal, wiec numer fabryczny czujnika gubil sie cicho.
    """

    WPIS = ("Termohigrometr typ: LB-701, nr fabr .: 2783 z panelem odczytowym "
            "(rejestratorem) typ: LB-706B, nr fabr.: 568, "
            "nr ewid.: DNW/PP/002/WS, wytwórca: LAB-EL.")

    def setUp(self):
        self.p = parsuj(self.WPIS)[0]

    def test_jeden_przyrzad_a_nie_dwa(self):
        self.assertEqual(len(parsuj(self.WPIS)), 1)

    def test_obiektem_jest_panel_odczytowy(self):
        self.assertEqual((self.p.typ, self.p.nr_fabr), ("LB-706B", "568"))

    def test_czujnikiem_jest_termohigrometr(self):
        self.assertEqual((self.p.czuj_typ, self.p.czuj_nr_fabr), ("LB-701", "2783"))

    def test_nr_ewidencyjny_nalezy_do_panelu(self):
        self.assertEqual(self.p.nr_ewid, "DNW/PP/002/WS")

    def test_wytworca_wspolny_dla_obu_czesci(self):
        self.assertEqual((self.p.wytworca, self.p.czuj_wytworca), ("LAB-EL", "LAB-EL"))

    def test_odstep_przed_kropka_w_nr_fabr(self):
        """'nr fabr .:' to artefakt PDF — numer ma sie znalezc mimo niego."""
        self.assertEqual(self.p.czuj_nr_fabr, "2783")

    def test_numer_nie_zbiera_ogona_z_fraza(self):
        """Regresja: nr fabr wychodzil jako '2783 z panelem odczytowym'."""
        self.assertNotIn("panel", self.p.czuj_nr_fabr.lower())

    def test_wariant_wraz_z_rejestratorem(self):
        p = parsuj("Termohigrometr typ: T1, nr fabr.: 111 wraz z rejestratorem "
                   "typ: R2, nr fabr.: 222, wytwórca: Testo.")[0]
        self.assertEqual((p.typ, p.nr_fabr), ("R2", "222"))
        self.assertEqual((p.czuj_typ, p.czuj_nr_fabr), ("T1", "111"))

    def test_rozpoznanie_nie_zalezy_od_nazw_przyrzadow(self):
        """Liczy sie fraza, nie konkretny model — inaczej dzialaloby dla jednego PZ."""
        p = parsuj("Termohigrometr typ: HD-9817, nr fabr.: A1122 z wyświetlaczem "
                   "typ: HD-2301, nr fabr.: B7788, nr ewid.: Q/LOG/40, "
                   "wytwórca: Delta Ohm.")[0]
        self.assertEqual((p.wytworca, p.typ, p.nr_fabr), ("Delta Ohm", "HD-2301", "B7788"))
        self.assertEqual((p.czuj_typ, p.czuj_nr_fabr), ("HD-9817", "A1122"))


class TestWskaznikPrzyWieluPrzyrzadach(unittest.TestCase):
    """
    PZ rzadko ma jeden przyrzad — zwykle jest ich kilka albo kilkanascie, w
    roznych ukladach naraz. Uklad z panelem musi dzialac tak samo w kazdym
    z nich, a nie tylko w pojedynczej linii.
    """

    def test_kazda_pozycja_parsowana_osobno(self):
        """Rozne uklady w jednym PZ nie moga sobie nawzajem przeszkadzac."""
        tekst = (
            "Obiekty wzorcowania:\n"
            "1) Termohigrometr typ: LB-701, nr fabr .: 2783 z panelem odczytowym "
            "(rejestratorem) typ: LB-706B, nr fabr.: 568, nr ewid.: DNW/PP/002/WS, "
            "wytwórca: LAB-EL.\n"
            "2) Termometr typ: 175T2, nr fabr.: 40118669, nr ewid.: Q/LOG/36, "
            "wytwórca: Testo.\n"
            "3) Termohigrometr złożony ze wskaźnika (rejestratora) typ: testo 176H1, "
            "nr fabr.: 40807872 oraz czujnika typ: 0636 9735, nr fabr.: 21201805, "
            "wytwórca: Testo.\n"
            "Metoda wzorcowania:\n")
        p = pz_dane.parsuj_tekst(tekst)
        self.assertEqual(len(p), 3)
        self.assertEqual((p[0].typ, p[0].nr_fabr, p[0].czuj_typ), ("LB-706B", "568", "LB-701"))
        self.assertEqual((p[1].typ, p[1].nr_fabr, p[1].czuj_typ), ("175T2", "40118669", ""))
        self.assertEqual((p[2].typ, p[2].nr_fabr, p[2].czuj_typ),
                         ("testo 176H1", "40807872", "0636 9735"))

    def test_uklad_wypunktowany_z_panelem(self):
        """Kilka sztuk pod jednym naglowkiem, kazda z wlasnym panelem."""
        wpis = ("Termohigrometr (rejestrator, 2 szt.) typ: LB-701, "
                "• nr fabr.: 2783 z panelem odczytowym typ: LB-706B, "
                "nr fabr.: 568, nr ewid.: UR1; "
                "• nr fabr.: 2784 z panelem odczytowym typ: LB-706B, "
                "nr fabr.: 569, nr ewid.: UR2; wytwórca: LAB-EL.")
        p = parsuj(wpis)
        self.assertEqual(len(p), 2)
        self.assertEqual([(x.typ, x.nr_fabr, x.nr_ewid) for x in p],
                         [("LB-706B", "568", "UR1"), ("LB-706B", "569", "UR2")])
        self.assertEqual([(x.czuj_typ, x.czuj_nr_fabr) for x in p],
                         [("LB-701", "2783"), ("LB-701", "2784")])

    def test_typ_czujnika_dziedziczy_z_naglowka_pozycji(self):
        """W podpunkcie stoja same numery — typ czujnika jest wyzej."""
        wpis = ("Termohigrometr (rejestrator, 2 szt.) typ: LB-701, nr fabr.: "
                "• 2783 z panelem odczytowym typ: LB-706B, nr fabr.: 568, nr ewid.: UR1; "
                "• 2784 z panelem odczytowym typ: LB-706B, nr fabr.: 569, nr ewid.: UR2; "
                "wytwórca: LAB-EL.")
        p = parsuj(wpis)
        self.assertEqual([x.czuj_nr_fabr for x in p], ["2783", "2784"])
        self.assertEqual([x.czuj_typ for x in p], ["LB-701", "LB-701"])

    def test_wytworca_trafia_do_obu_czesci_kazdej_sztuki(self):
        wpis = ("Termohigrometr (rejestrator, 2 szt.) typ: LB-701, "
                "• nr fabr.: 2783 z panelem odczytowym typ: LB-706B, nr fabr.: 568; "
                "• nr fabr.: 2784 z panelem odczytowym typ: LB-706B, nr fabr.: 569; "
                "wytwórca: LAB-EL.")
        for x in parsuj(wpis):
            with self.subTest(nr=x.nr_fabr):
                self.assertEqual((x.wytworca, x.czuj_wytworca), ("LAB-EL", "LAB-EL"))

    def test_numerowana_lista_sztuk_jednej_pozycji(self):
        """
        Zgloszenie z PZ 211: cztery termohigrometry jednego typu, a numeracja
        '1) 2) 3) 4)' wylicza SZTUKI, nie osobne pozycje:

            Termohigrometr (rejestrator) typ: TERMIOPLUS,
            1) nr fabr.: 5000722, nr CLDK: CLDK/B-49;
            ...
            4) nr fabr.: 2950722, nr CLDK: CLDK/B-52, wytworca: Termoprodukt.

        Numeracja byla brana za podzial na pozycje, wiec naglowek z typem
        przepadal, a wytworca — stojacy raz, po ostatnim podpunkcie — trafial
        tylko do ostatniej sztuki. W protokole trzy z czterech wierszy mialy
        pusta kolumne 'Wytworca'.
        """
        tekst = ("Obiekty wzorcowania:\n"
                 "Termohigrometr (rejestrator) typ: TERMIOPLUS,\n"
                 "1) nr fabr.: 5000722, nr CLDK: CLDK/B-49;\n"
                 "2) nr fabr.: 4180722, nr CLDK: CLDK/B-50;\n"
                 "3) nr fabr.: 3780722, nr CLDK: CLDK/B-51;\n"
                 "4) nr fabr.: 2950722, nr CLDK: CLDK/B-52, wytwórca: Termoprodukt.\n"
                 "Metoda wzorcowania:\n")
        p = pz_dane.parsuj_tekst(tekst)
        self.assertEqual(len(p), 4)
        self.assertEqual([x.nr_fabr for x in p],
                         ["5000722", "4180722", "3780722", "2950722"])
        self.assertEqual([x.wytworca for x in p], ["Termoprodukt"] * 4)
        self.assertEqual([x.typ for x in p], ["TERMIOPLUS"] * 4)

    def test_wlasna_etykieta_numeru_nie_tworzy_widma(self):
        """
        'nr CLDK:' nie byl znany jako granica, wiec wartosc 'nr fabr.' siegala do
        konca podpunktu i po rozdzieleniu przecinkiem powstawal drugi, nieistniejacy
        przyrzad o numerze 'nr CLDK: CLDK/B-49'. Z czterech sztuk robilo sie osiem.
        """
        p = parsuj("typ: X, nr fabr.: 5000722, nr CLDK: CLDK/B-49, wytwórca: Termoprodukt.")
        self.assertEqual(len(p), 1)
        self.assertEqual(p[0].nr_fabr, "5000722")

    def test_dowolny_wlasny_skrot_konczy_numer(self):
        """Zasada jest ogolna — nie lista znanych skrotow."""
        p = parsuj("typ: X, nr fabr.: 111222, nr inw.: INW/7, wytwórca: Testo.")
        self.assertEqual(len(p), 1)
        self.assertEqual(p[0].nr_fabr, "111222")

    def test_wlasny_skrot_to_numer_ewidencyjny(self):
        """'nr CLDK:' pelni te sama role co 'nr ewid.:' — idzie do tej samej kolumny."""
        p = parsuj("typ: X, nr fabr.: 5000722, nr CLDK: CLDK/B-49, wytwórca: Termoprodukt.")
        self.assertEqual(p[0].nr_ewid, "CLDK/B-49")

    def test_numer_katalogowy_to_nie_ewidencyjny(self):
        """'nr kat.:' opisuje model, a nie egzemplarz — nie wolno go tu wpisac."""
        p = parsuj("typ: Y, nr fabr.: 999, nr kat.: KAT-5, wytwórca: Testo.")
        self.assertEqual((p[0].nr_fabr, p[0].nr_ewid), ("999", ""))

    def test_znane_zapisy_maja_pierwszenstwo(self):
        for wpis, oczekiwany in (
                ("typ: X, nr fabr.: 111, nr wew.: AB-1; wytwórca: Testo.", "AB-1"),
                ("typ: 174H, nr fabr.: 123456, nr wewn.: CL-1318A, wytwórca: Testo.", "CL-1318A"),
                ("typ: M1, nr fabr.: TMM1605, nr ewid.: Q/LOG/19, wytwórca: Tempmate.", "Q/LOG/19"),
        ):
            with self.subTest(wpis=wpis[:28]):
                self.assertEqual(parsuj(wpis)[0].nr_ewid, oczekiwany)

    def test_prawdziwe_pozycje_nadal_sa_pozycjami(self):
        """Gdy podpunkt zaczyna sie od NAZWY przyrzadu — to osobna pozycja PZ."""
        tekst = ("Obiekty wzorcowania:\n"
                 "1) Termometr typ: 175T2, nr fabr.: 40118669, wytwórca: Testo.\n"
                 "2) Termohigrometr typ: 174H, nr fabr.: 83623973, wytwórca: Testo.\n"
                 "Metoda wzorcowania:\n")
        p = pz_dane.parsuj_tekst(tekst)
        self.assertEqual([(x.pozycja, x.typ) for x in p],
                         [(1, "175T2"), (2, "174H")])

    def test_bez_typu_w_naglowku_numeracja_to_pozycje(self):
        """Sam naglowek 'Przyrzady:' nie opisuje przyrzadu — nie wolno scalac."""
        tekst = ("Obiekty wzorcowania: Przyrządy zleceniodawcy:\n"
                 "1) nr fabr.: 111, wytwórca: Testo.\n"
                 "2) nr fabr.: 222, wytwórca: Testo.\n"
                 "Metoda wzorcowania:\n")
        p = pz_dane.parsuj_tekst(tekst)
        self.assertEqual([x.pozycja for x in p], [1, 2])

    def test_wypunktowany_bez_panelu_bez_zmian(self):
        """Regresja: uklad z PZ 197 ma isc stara droga."""
        wpis = ("Termohigrometr (rejestrator, 2 szt.) typ: testo 174H, nr fabr.:\n"
                "• 83623973, nr wew.: UR00045;\n"
                "• 83617608, nr wew.: UR00052;\n"
                "wytwórca: Testo.")
        p = parsuj(wpis)
        self.assertEqual([x.nr_fabr for x in p], ["83623973", "83617608"])
        self.assertEqual([x.typ for x in p], ["testo 174H", "testo 174H"])
        self.assertEqual([x.czuj_typ for x in p], ["", ""])


class TestKolejnoscNieOdwraca_Sie_Bez_Powodu(unittest.TestCase):
    """
    Nowy uklad wolno wlaczyc TYLKO wtedy, gdy dopisana czesc ma wlasny 'typ:'.
    Inaczej zwykly opis ('z wyswietlaczem LCD') zamienialby role kolumn.
    """

    def test_uklad_ze_wskaznikiem_z_przodu_bez_zmian(self):
        """'zlozony ze wskaznika ... oraz czujnika ...' — kolejnosc jak dotad."""
        p = parsuj("Termohigrometr złożony ze wskaźnika (rejestratora) "
                   "typ: testo 176H1, nr fabr.: 40807872 oraz czujnika "
                   "temperatury typ: 0636 9735, nr fabr.: 21201805, "
                   "wytwórca: Testo.")[0]
        self.assertEqual((p.typ, p.nr_fabr), ("testo 176H1", "40807872"))
        self.assertEqual((p.czuj_typ, p.czuj_nr_fabr), ("0636 9735", "21201805"))

    def test_wzmianka_bez_typu_nie_przestawia_kolumn(self):
        p = parsuj("Termohigrometr z wyświetlaczem typ: 174H, nr fabr.: 123456, "
                   "nr ewid.: UR1, wytwórca: Testo.")[0]
        self.assertEqual((p.typ, p.nr_fabr, p.nr_ewid), ("174H", "123456", "UR1"))
        self.assertEqual((p.czuj_typ, p.czuj_nr_fabr), ("", ""))

    def test_zwykly_jednoliniowy_bez_zmian(self):
        p = parsuj("typ: M1, nr fabr.: TMM160500502, nr ewid.: Q/LOG/19, "
                   "wytwórca: Tempmate.")[0]
        self.assertEqual((p.typ, p.nr_fabr, p.nr_ewid),
                         ("M1", "TMM160500502", "Q/LOG/19"))


class TestPunktyListaWilgotnosci(unittest.TestCase):
    """
    Zgloszenie z PZ 218: skrypt pominal punkt 20 °C / 15 %rh (nastawa komory
    19,9 / 13). Zawinily DWIE rzeczy naraz i obie sprowadzaly sie do tego, ze
    lista punktow z PZ wychodzila PUSTA:

      1. Naglowek sekcji w liczbie mnogiej — 'Calibration ranges:'. Wzorzec
         wymagal 'Calibration range', wiec sekcja w ogole nie byla znajdowana.
      2. Trzeci zapis punktow, ktorego parser nie znal:
             - temperature: 20 °C; relative humidity: 15 %rh, 45 %rh, 95 %rh
         czyli jedna temperatura i LISTA wilgotnosci.

    Bez punktow z PZ segment nie mogl byc uznany za 'punkt zamowiony', wiec gdy
    komora nie dociagnela nastawy (13 %rh wobec ~15,9 %rh realnych), byl cicho
    odrzucany jako przejscie/suszenie — zamiast zostac punktem pomaranczowym.
    """

    SEKCJA_EN = (
        "Calibration ranges:\n"
        "1) Calibration in the range of temperature and relative humidity:\n"
        "- temperature: (-20; 5; 20; 40; 60; 70) °C,\n"
        "- temperature: 10 °C; relative humidity: 22 %rh, 45 %rh, 95 %rh,\n"
        "- temperature: 20 °C; relative humidity: 15 %rh, 45 %rh, 95 %rh, 45 %rh 1)\n"
        "1) Point (20 °C, 45 %rh) will be repeated in order to determine "
        "the hysteresis value of the calibrated instrument.\n"
        "2) Calibration in the range of temperature:\n"
        "- temperature: (-30; -25; 5) °C.\n"
        "Statement of conformity:\n"
        "The certificates will include a statement of conformity with requirements.\n")

    def punkty(self, tekst=None):
        return pz_dane.punkty_wzorcowania(tekst if tekst is not None else self.SEKCJA_EN)

    def test_sekcja_w_liczbie_mnogiej_jest_znajdowana(self):
        self.assertTrue(self.punkty())

    def test_punkt_ze_zgloszenia_jest_na_liscie(self):
        self.assertIn((20.0, 15.0), self.punkty())

    def test_cala_lista_wilgotnosci_jednej_temperatury(self):
        z_10 = [p for p in self.punkty() if p[0] == 10.0]
        self.assertEqual(z_10, [(10.0, 22.0), (10.0, 45.0), (10.0, 95.0)])

    def test_kolejnosc_jak_w_zamowieniu(self):
        """Najpierw lista samych temperatur, potem punkty z wilgotnoscia."""
        p = self.punkty()
        self.assertEqual(p[0], (-20.0, None))
        self.assertEqual(p[6], (10.0, 22.0))

    def test_powtorzony_punkt_histerezy_zostaje(self):
        """'45 %rh' pada w liscie 20 °C dwa razy — oba wystapienia sa znaczace."""
        self.assertEqual(sum(1 for p in self.punkty() if p == (20.0, 45.0)), 2)

    def test_przypis_nie_dodaje_trzeciego_punktu(self):
        """
        Przypis '1) Point (20 °C, 45 %rh) will be repeated...' opisuje punkt JUZ
        wymieniony. Policzony osobno dawalby punkt, ktorego nikt nie zamowil.
        """
        tekst = ("Calibration ranges:\n"
                 "- temperature: 20 °C; relative humidity: 45 %rh.\n"
                 "1) Point (20 °C, 45 %rh) will be repeated in order to determine "
                 "the hysteresis.\n")
        self.assertEqual(self.punkty(tekst), [(20.0, 45.0)])

    def test_przypis_po_polsku_tez_nie_dodaje_punktu(self):
        tekst = ("Zakres wzorcowania:\n"
                 "(25 °C, 60 %rh)\n"
                 "1) powtórzony punkt (25 °C, 60 %rh) w celu wyznaczenia histerezy.\n")
        self.assertEqual(self.punkty(tekst), [(25.0, 60.0)])

    def test_druga_pozycja_z_samymi_temperaturami(self):
        self.assertIn((-30.0, None), self.punkty())

    def test_sekcja_konczy_sie_na_oswiadczeniu_o_zgodnosci(self):
        """Dalszy tekst PZ nie moze dorzucac przypadkowych liczb."""
        self.assertNotIn((95.0, None), self.punkty())

    def test_wariant_polski_tego_ukladu(self):
        tekst = ("Zakresy wzorcowania:\n"
                 "- temperatura: 20 °C; wilgotność względna: 15 %rh, 45 %rh,\n")
        self.assertEqual(self.punkty(tekst), [(20.0, 15.0), (20.0, 45.0)])

    def test_dotychczasowe_zapisy_bez_zmian(self):
        """Regresja: nawiasowy zapis i lista temperatur dzialaja jak dotad."""
        tekst = ("Zakres wzorcowania: (-20; 0; 40) °C, (25 °C, 30 %rh); "
                 "(25 °C, 60 %rh); (25 °C, 85 %rh)\n")
        self.assertEqual(self.punkty(tekst),
                         [(-20.0, None), (0.0, None), (40.0, None),
                          (25.0, 30.0), (25.0, 60.0), (25.0, 85.0)])


if __name__ == "__main__":
    unittest.main(verbosity=2)
