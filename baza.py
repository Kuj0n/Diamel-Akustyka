"""
baza.py

Baza materiałów (przegrody, tłumiki), silników i chłodnic - jak w Kalkulatorze Rozruchu:
- baza globalna Diamel (panel administracyjny diamel.pl) pobierana w całości i zastępowana przy każdej aktualizacji
  (wpisy usunięte z bazy globalnej znikają także u użytkownika),
- wpisy własne użytkownika (nowe albo zmienione pozycje bazy globalnej) trzymane osobno - aktualizacja ich nie rusza.
Pliki: %APPDATA%\\Diamel\\Kalkulator Akustyki\\data: materials.json, engines.json, radiators.json (kopia bazy globalnej),
wlasne.json; pliki producentów z bazy (karty katalogowe) - folder pliki (pobierane po zalogowaniu, przy pierwszym użyciu).
"""
import json
import os
import shutil
import sys
from datetime import datetime

from konfiguracja_app import DB_VERSION, sciezka_do_zasobu

KATALOG_DANYCH = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "Diamel", "Kalkulator Akustyki")
KATALOG_BAZY = os.path.join(KATALOG_DANYCH, "data")
KATALOG_KOPII = os.path.join(KATALOG_DANYCH, "kopie_bazy")
PLIK_WLASNYCH = os.path.join(KATALOG_BAZY, "wlasne.json")
KATALOG_PLIKOW = os.path.join(KATALOG_DANYCH, "pliki")
PLIKI = {"materials": "materials.json", "engines": "engines.json", "radiators": "radiators.json"}
KOPIE_MAX = 10
TYPY_MATERIALU = ("Widmo pasmowe", "Rw / Ctr")

# Pola pozycji bazy: (klucz, opis, jednostka, rodzaj: True - liczba, False - tekst, "widmo" - 8 pasm 63 Hz - 8 kHz,
# "jedn" - dB / dB(A), "typ" - rodzaj materiału). Edytor zachowuje też pola spoza tej listy.
POLA_URZADZEN = {
    "materials": [
        ("name_en", "Nazwa angielska (raport EN)", "", False),
        ("type", "Rodzaj danych", "", "typ"),
        ("spec", "Izolacyjność w pasmach", "dB", "widmo"),
        ("rw", "Rw", "dB", True),
        ("ctr", "Ctr", "dB", True),
        ("d", "Grubość / długość", "mm", True),
        ("desc_pl", "Opis (raport PL)", "", False),
        ("desc_en", "Opis (raport EN)", "", False),
    ],
    "engines": [
        ("genset", "Agregat (model)", "", False),
        ("lw", "Widmo mocy akustycznej Lw", "", "widmo"),
        ("jedn", "Jednostka widma", "", "jedn"),
        ("notes", "Uwagi", "", False),
    ],
    "radiators": [
        ("lw", "Widmo mocy akustycznej Lw", "", "widmo"),
        ("jedn", "Jednostka widma", "", "jedn"),
        ("att", "Tłumienie rdzenia", "dB", "widmo"),
        ("notes", "Uwagi", "", False),
    ],
}
# Jednostka widma, gdy pozycja jej nie ma (jak liczyła wersja 1.x): silniki - dB(A), chłodnice - dB
JEDNOSTKA_DOMYSLNA = {"engines": "dBA", "radiators": "dB"}


def _czytaj(sciezka, domyslnie=None):
    try:
        with open(sciezka, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return domyslnie


def _zapisz(sciezka, dane):
    os.makedirs(os.path.dirname(sciezka), exist_ok=True)
    tymczasowy = sciezka + ".tmp"
    with open(tymczasowy, "w", encoding="utf-8") as f:
        json.dump(dane, f, indent=4, ensure_ascii=False)
    os.replace(tymczasowy, sciezka)


def _wbudowana(rodzaj):
    return (_czytaj(sciezka_do_zasobu(os.path.join("data", PLIKI[rodzaj])), {}) or {}).get(rodzaj, {})


def _stara_baza():
    """Baza wersji 1.x: acoustic_db.json obok programu (instalacja) albo w folderze roboczym."""
    program = os.path.dirname(sys.executable)
    # Program Files bez prawa zapisu: Windows zapisywał bazę 1.x w VirtualStore
    wirtualny = os.path.join(os.environ.get("LOCALAPPDATA", ""), "VirtualStore", os.path.splitdrive(program)[1].lstrip("\\/"))
    for folder in dict.fromkeys([wirtualny, program, os.path.abspath(".")]):
        dane = _czytaj(os.path.join(folder, "acoustic_db.json"))
        if isinstance(dane, dict) and all(isinstance(dane.get(r), dict) for r in PLIKI):
            return folder, dane
    return "", None


class Baza:
    def __init__(self):
        self.globalna = {r: {} for r in PLIKI}
        self.wlasne = {r: {} for r in PLIKI}
        self.wersja = DB_VERSION
        self.przeniesiono_z = ""
        self._przygotuj()
        self.wczytaj()

    def _plik(self, rodzaj):
        return os.path.join(KATALOG_BAZY, PLIKI[rodzaj])

    def _przygotuj(self):
        """
        Pierwsze uruchomienie tej wersji: wpisy własne wydzielone ze starej bazy (acoustic_db.json z wersji 1.x) - pozycje,
        których nie ma w bazie wbudowanej albo mają inne wartości. Kopia bazy globalnej = baza wbudowana.
        """
        if os.path.exists(PLIK_WLASNYCH):
            return
        os.makedirs(KATALOG_BAZY, exist_ok=True)
        wlasne = {r: {} for r in PLIKI}
        folder, stara = _stara_baza()
        if stara:
            for rodzaj in PLIKI:
                wbud = _wbudowana(rodzaj)
                wlasne[rodzaj] = {n: v for n, v in stara[rodzaj].items() if wbud.get(n) != v}
            if any(wlasne.values()):
                self.przeniesiono_z = folder
        for rodzaj, p in PLIKI.items():
            wbudowany = sciezka_do_zasobu(os.path.join("data", p))  # kopia bazy z serwera z dnia kompilacji
            if os.path.exists(wbudowany):
                shutil.copy(wbudowany, self._plik(rodzaj))
            else:  # bez kopii wbudowanej (uruchomienie ze źródeł): pusta kopia globalna, wersja 0 - pełna baza z serwera przy starcie
                _zapisz(self._plik(rodzaj), {"version": "0", rodzaj: {}})
        _zapisz(PLIK_WLASNYCH, wlasne)

    def wczytaj(self):
        for rodzaj in PLIKI:
            dane = _czytaj(self._plik(rodzaj)) or _czytaj(sciezka_do_zasobu(os.path.join("data", PLIKI[rodzaj])), {}) or {}
            self.globalna[rodzaj] = dane.get(rodzaj, {})
            if rodzaj == "materials":
                self.wersja = dane.get("version", self.wersja)
        wlasne = _czytaj(PLIK_WLASNYCH, {}) or {}
        self.wlasne = {r: wlasne.get(r, {}) for r in PLIKI}

    # -------------------------------------------------------------- odczyt

    def urzadzenia(self, rodzaj):
        """Baza globalna + wpisy własne (własny wpis o tej samej nazwie zastępuje globalny)."""
        wynik = {**self.globalna[rodzaj], **self.wlasne[rodzaj]}
        if rodzaj in JEDNOSTKA_DOMYSLNA:
            wynik = {n: {"jedn": JEDNOSTKA_DOMYSLNA[rodzaj], **v} for n, v in wynik.items()}
        return wynik

    def zrodla(self, rodzaj):
        """nazwa -> 'globalna' / 'wlasna' (nie ma w bazie globalnej) / 'zmieniona' (własna wersja pozycji z bazy globalnej)."""
        wynik = {n: "globalna" for n in self.globalna[rodzaj]}
        for n in self.wlasne[rodzaj]:
            wynik[n] = "zmieniona" if n in self.globalna[rodzaj] else "wlasna"
        return wynik

    # -------------------------------------------------------------- wpisy własne

    def _zapisz_wlasne(self):
        _zapisz(PLIK_WLASNYCH, self.wlasne)

    def zapisz_urzadzenie(self, rodzaj, stara_nazwa, nazwa, dane):
        """Zapis z edytora - zawsze jako wpis własny. Pola nieznane edytorowi (np. dodane w nowszej bazie) zostają."""
        nazwa = nazwa.strip()
        poprzednie = {**self.globalna[rodzaj], **self.wlasne[rodzaj]}.get(stara_nazwa, {}) if stara_nazwa else {}
        if stara_nazwa and stara_nazwa != nazwa:
            self.wlasne[rodzaj].pop(stara_nazwa, None)  # zmiana nazwy - pozycja z bazy globalnej pod starą nazwą zostaje
        nowe = {**poprzednie, **dane}
        globalna = self.globalna[rodzaj].get(nazwa)
        if globalna is not None and {"jedn": JEDNOSTKA_DOMYSLNA.get(rodzaj), **globalna} == {"jedn": JEDNOSTKA_DOMYSLNA.get(rodzaj), **nowe}:
            self.wlasne[rodzaj].pop(nazwa, None)  # bez różnic względem bazy globalnej - nie trzeba własnej wersji
        else:
            self.wlasne[rodzaj][nazwa] = nowe
        self._zapisz_wlasne()

    def usun_urzadzenie(self, rodzaj, nazwa):
        """Usuwa wpis własny. Pozycja z bazy globalnej wraca do wartości z bazy (tej nie da się usunąć)."""
        self.wlasne[rodzaj].pop(nazwa, None)
        self._zapisz_wlasne()

    # -------------------------------------------------------------- aktualizacja z serwera

    def kopia(self):
        """Kopia bazy przed aktualizacją (ostatnie KOPIE_MAX). Zwraca folder kopii."""
        cel = os.path.join(KATALOG_KOPII, datetime.now().strftime("%Y-%m-%d_%H%M%S"))
        os.makedirs(cel, exist_ok=True)
        for p in list(PLIKI.values()) + [os.path.basename(PLIK_WLASNYCH)]:
            if os.path.exists(os.path.join(KATALOG_BAZY, p)):
                shutil.copy(os.path.join(KATALOG_BAZY, p), cel)
        for stara in sorted(os.listdir(KATALOG_KOPII))[:-KOPIE_MAX]:
            shutil.rmtree(os.path.join(KATALOG_KOPII, stara), ignore_errors=True)
        return cel

    def zastap_globalna(self, nowe, nowa_wersja):
        """Pełna aktualizacja: baza globalna zastąpiona pobraną. Zwraca zmiany {rodzaj: {nowe, zmienione, usuniete}}."""
        zmiany = {}
        for rodzaj, wpisy in nowe.items():
            stare = self.globalna[rodzaj]
            zmiany[rodzaj] = {"nowe": sorted(set(wpisy) - set(stare)), "usuniete": sorted(set(stare) - set(wpisy)),
                              "zmienione": sorted(n for n in set(wpisy) & set(stare) if wpisy[n] != stare[n])}
            self.globalna[rodzaj] = wpisy
            _zapisz(self._plik(rodzaj), {"version": nowa_wersja, rodzaj: wpisy})
        self.wersja = nowa_wersja
        # własna wersja identyczna z nową globalną (np. materiał zgłoszony i dodany do bazy) - już niepotrzebna
        for rodzaj in PLIKI:
            for n in [n for n, v in self.wlasne[rodzaj].items() if self.globalna[rodzaj].get(n) == v]:
                del self.wlasne[rodzaj][n]
        self._zapisz_wlasne()
        return zmiany
