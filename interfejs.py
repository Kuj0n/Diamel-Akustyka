"""
interfejs.py

Funkcje Kalkulatora Akustyki wywoływane z okna (ui/app.js przez window.pywebview.api):
projekt, obliczenia na bieżąco, raport PDF / DOCX, baza materiałów, silników i chłodnic, aktualizacje, zgłoszenia.
"""
import base64
import copy
import json
import mimetypes
import os
import re
import subprocess
import tempfile
import threading
import urllib.error
import urllib.request
from datetime import datetime

import webview

import obliczenia
import raport_docx

from baza import Baza, KATALOG_DANYCH, KATALOG_PLIKOW, POLA_URZADZEN, PLIKI, TYPY_MATERIALU
from konto import Konto
import konto
from konfiguracja_app import VERSION, BUILD_DATE, CHANGELOG, GITHUB_REPO, NAZWA_PROGRAMU, URL_ZGLOSZEN, URL_BAZY

PLIK_USTAWIEN = os.path.join(KATALOG_DANYCH, "ustawienia.json")
NAGLOWKI = {"User-Agent": f"Diamel-{NAZWA_PROGRAMU}/{VERSION}"}
INSTALATOR = os.path.join(tempfile.gettempdir(), "Aktualizacja_Akustyka_Instalator.exe")
OSTATNIE_MAX = 8
OBRAZ_MAX_MB = 15
OBRAZ = "Obraz (PNG/JPG)"
DOKUMENT = "Dokument (PDF)"

USTAWIENIA_PROGRAMU = {
    "motyw": "system",       # jasny / ciemny / system
    "rozmiar": "normalny",   # maly / normalny / duzy
    "autor": "MK",           # inicjały w nowym projekcie
    "jezyk": "Polski",       # język raportu w nowym projekcie
    "otworz_raport": True,   # otwórz raport po zapisaniu
    "pdf_zalaczniki": True,  # karty producentów (PDF) dołączone na końcu raportu PDF
    "ostatnie": [],          # ostatnio otwierane projekty
    "folder_raportow": "",   # ostatni folder zapisu raportu (gdy projekt nie jest zapisany)
    "szablon": "",           # plik .docx z nagłówkiem / stopką raportu (pusty = stopka programu)
}

PUSTA_WARSTWA = {"name": "", "name_en": "", "desc_pl": "", "desc_en": "", "type": "Widmo pasmowe", "d": 0, "rw": 0, "ctr": 0, "spec": [0] * 8}

# Nowy projekt - wartości jak w wersji 1.x (format pliku .json bez zmian, nowe pola: jednostki widm, autor)
NOWY_PROJEKT = {
    "meta": {"project_no": "", "client_name": "", "po_number": "", "client_doc_no": "", "project_name": "",
             "silnik_typ": "", "agregat_typ": "", "typ_obudowy": "Obudowa kontenerowa", "jezyk_raportu": "Polski",
             "logo_path": "", "footer_text": "", "engine_wiz_path": "", "engine_pdf_path": "", "autor": ""},
    "revisions": [],
    "dimensions": {"kont_l": 12.0, "kont_w": 3.0, "kont_h": 3.0},
    "points": [{"name": "1 m", "d": 1.0, "limit": 85.0}],
    "sabine": {"alfa_podl": 0.02, "alfa_scian": 0.85, "sabine_q": 2.0, "sabine_r": 1.5},
    "radiator": {"active": False, "name": "", "lw": [0.0] * 8, "att": [0.0] * 8, "jedn": "dB", "wiz_path": "", "pdf_path": ""},
    "modules_active": {"sciana": True, "podloga": False, "sufit": False, "drzwi": True, "czerpnia": False, "wyrzut": False, "wydech": False},
    "mod_dims": {k: {"qty": 1, "w": 1.0, "h": 2.0, "d": 0.0, "s": 0.0, **({"kierunek": "Poziomy", "wysokosc": 0.0} if k in obliczenia.Z_KIERUNKIEM else {})}
                 for k in obliczenia.MODULY},
    "layers": {k: [] for k in obliczenia.MODULY},
    "mod_attachments": {k: [] for k in obliczenia.MODULY},
    "acoustics": {"lw_vars": [95.0] * 8, "room_abs_vars": [-3.0] * 8, "exhaust_lw_vars": [0.0] * 8, "jedn_silnika": "dBA", "jedn_wydechu": "dB"},
}


def wez_krotke(v):
    return tuple(map(int, (v.lstrip("vV").split(" ")[0]).split(".")))


def pobierz_json(url, timeout=6.0):
    req = urllib.request.Request(url, headers=NAGLOWKI)
    with urllib.request.urlopen(req, timeout=timeout) as odp:
        return json.loads(odp.read().decode("utf-8"))


def wyslij_json(url, dane, timeout=15.0):
    """POST z JSON; odpowiedź serwera (także z kodem błędu) jako słownik; brak sieci -> wyjątek."""
    req = urllib.request.Request(url, data=json.dumps(dane).encode("utf-8"), headers={**NAGLOWKI, "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as odp:
            return json.loads(odp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            return json.loads(e.read().decode("utf-8"))
        except Exception:
            raise e


ZALACZNIKI_MAX_MB = 5  # limit serwera zgłoszeń (diamel.pl/api/zgloszenie.php)
ZALACZNIKI_TYPY = ("Pliki (*.pdf;*.png;*.jpg;*.jpeg;*.doc;*.docx;*.xls;*.xlsx;*.txt;*.zip)",)


def _czytaj_json(sciezka, domyslnie):
    try:
        with open(sciezka, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return domyslnie


def _zapisz_json(sciezka, dane):
    os.makedirs(os.path.dirname(sciezka), exist_ok=True)
    tymczasowy = sciezka + ".tmp"
    with open(tymczasowy, "w", encoding="utf-8") as f:
        json.dump(dane, f, ensure_ascii=False, indent=4)
    os.replace(tymczasowy, sciezka)


def _widmo(v, domyslnie=0.0):
    lista = list(v) if isinstance(v, (list, tuple)) else []
    return [x if isinstance(x, (int, float, str)) else domyslnie for x in (lista + [domyslnie] * 8)[:8]]


def uzupelnij_projekt(stan):
    """Projekt z pliku (także z wersji 1.x): brakujące pola z nowego projektu, widma po 8 pasm, listy warstw i załączników."""
    p = copy.deepcopy(NOWY_PROJEKT)
    for klucz in ("meta", "dimensions", "sabine", "radiator", "modules_active", "acoustics"):
        if isinstance(stan.get(klucz), dict):
            p[klucz].update(stan[klucz])
    for klucz in ("revisions", "points"):
        if isinstance(stan.get(klucz), list):
            p[klucz] = [x for x in stan[klucz] if isinstance(x, dict)]
    for k in obliczenia.MODULY:
        if isinstance((stan.get("mod_dims") or {}).get(k), dict):
            p["mod_dims"][k].update(stan["mod_dims"][k])
        p["layers"][k] = [{**PUSTA_WARSTWA, **w, "spec": _widmo(w.get("spec"))} for w in (stan.get("layers") or {}).get(k, []) if isinstance(w, dict)]
        p["mod_attachments"][k] = [a for a in (stan.get("mod_attachments") or {}).get(k, []) if isinstance(a, dict) and a.get("path")]
    for klucz in ("lw_vars", "room_abs_vars", "exhaust_lw_vars"):
        p["acoustics"][klucz] = _widmo(p["acoustics"].get(klucz))
    p["radiator"]["lw"], p["radiator"]["att"] = _widmo(p["radiator"].get("lw")), _widmo(p["radiator"].get("att"))
    if not p["meta"].get("autor") and p["revisions"]:  # wersje 1.x: autor = wykonawca ostatniej rewizji
        p["meta"]["autor"] = str(p["revisions"][-1].get("drawn", ""))
    if p["meta"].get("jezyk_raportu") not in ("Polski", "English"):
        p["meta"]["jezyk_raportu"] = "Polski"
    return p


class Api:
    def __init__(self):
        self._okno = None
        self._baza = Baza()
        self._konto = Konto()
        self.program = {**USTAWIENIA_PROGRAMU, **_czytaj_json(PLIK_USTAWIEN, {}).get("program", {})}
        self.ui_gotowe = threading.Event()
        self._niezapisane = False
        self._zamykam = False

    # ------------------------------------------------------------------ start, ustawienia

    def informacje(self):
        self.ui_gotowe.set()  # okno działa - awaryjne okienko aktualizacji (aplikacja.py) niepotrzebne
        return {"wersja": VERSION, "data": BUILD_DATE[:10], "program": self.program, "changelog": CHANGELOG.strip(),
                "baza": self.baza(), "pola_urzadzen": POLA_URZADZEN, "typy_materialu": TYPY_MATERIALU, "projekt": self.nowy_projekt(),
                "przeniesiono_baze": self._baza.przeniesiono_z, "konto": self._konto.stan(),
                "pasma": obliczenia.PASMA, "korekta_a": obliczenia.KOREKTA_A, "moduly": obliczenia.MODULY}

    # ------------------------------------------------------------------ konto Diamel (nieobowiązkowe)

    def konto_odswiez(self):
        return self._konto.odswiez()

    def konto_zaloguj(self, login, haslo, zapamietaj=True):
        return self._konto.zaloguj(login, haslo, zapamietaj)

    def zakoncz_sesje(self):
        self._konto.zakoncz_sesje()

    def konto_przypomnij(self, login):
        return self._konto.przypomnij(login)

    def konto_wyloguj(self):
        return self._konto.wyloguj()

    def _szablon(self):
        """Szablon raportu: własny z Ustawień (zastępuje firmowy) albo szablon firmy zalogowanego użytkownika; '' = stopka programu."""
        return self.program["szablon"] or self._konto.szablon

    def ustaw_program(self, klucz, wartosc):
        if klucz in USTAWIENIA_PROGRAMU:
            self.program[klucz] = wartosc
            _zapisz_json(PLIK_USTAWIEN, {"program": self.program})
        return self.program

    def _dodaj_ostatni(self, sciezka):
        lista = [p for p in self.program["ostatnie"] if os.path.normcase(p) != os.path.normcase(sciezka)]
        self.ustaw_program("ostatnie", [sciezka] + lista[:OSTATNIE_MAX - 1])

    def ostatnie(self):
        return [{"sciezka": p, "nazwa": os.path.basename(p), "folder": os.path.dirname(p),
                 "data": datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m-%d %H:%M")}
                for p in self.program["ostatnie"] if os.path.isfile(p)]

    # ------------------------------------------------------------------ aktualizacje programu

    def sprawdz_aktualizacje(self):
        """Wersja zalecana = wydanie oznaczone na GitHub jako najnowsze; starsza od zainstalowanej = wycofanie wadliwej wersji."""
        try:
            dane = pobierz_json(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest", timeout=4.0)
        except Exception as e:
            return {"ok": False, "blad": str(e)}
        wersja = dane.get("tag_name", "0.0.0").lstrip("vV")
        link = next((a["browser_download_url"] for a in dane.get("assets", []) if a["name"].lower().endswith(".exe")), "")
        try:
            nowsza = bool(link) and wez_krotke(wersja) > wez_krotke(VERSION)
            wycofana = bool(link) and wez_krotke(wersja) < wez_krotke(VERSION)
        except ValueError:
            nowsza = wycofana = False
        return {"ok": True, "nowsza": nowsza, "wycofana": wycofana, "do_instalacji": nowsza or wycofana,
                "wersja": wersja, "opis": dane.get("body") or "", "link": link}

    def pobierz_aktualizacje(self, link):
        """Pobiera instalator w tle; postęp: window.postepAktualizacji(procent), koniec: window.aktualizacjaPobrana(ok, tekst)."""
        def postep(bloki, rozmiar_bloku, calosc):
            if calosc > 0 and self._okno:
                self._okno.evaluate_js(f"window.postepAktualizacji && window.postepAktualizacji({min(100, int(bloki * rozmiar_bloku * 100 / calosc))})")

        def watek():
            try:
                urllib.request.urlretrieve(link, INSTALATOR, postep)
                self._okno.evaluate_js("window.aktualizacjaPobrana(true, '')")
            except Exception as e:
                self._okno.evaluate_js(f"window.aktualizacjaPobrana(false, {json.dumps(str(e))})")
        threading.Thread(target=watek, daemon=True).start()
        return {"ok": True}

    def uruchom_instalator(self):
        subprocess.run(["powershell", "-NoProfile", "-Command", f"Unblock-File -Path '{INSTALATOR}'"], shell=True)
        subprocess.Popen([INSTALATOR, "/FORCECLOSEAPPLICATIONS", "/NOCANCEL"])
        os._exit(0)

    # ------------------------------------------------------------------ baza materiałów, silników i chłodnic

    def baza(self):
        b = self._baza
        return {"wersja": b.wersja, **{r: b.urzadzenia(r) for r in PLIKI}, "zrodla": {r: b.zrodla(r) for r in PLIKI}}

    def sprawdz_baze(self):
        """Wersja bazy w panelu diamel.pl (numer zwiększany przy każdej zmianie) - inna niż lokalna = do pobrania."""
        try:
            dane = wyslij_json(URL_BAZY, {"akcja": "wersja"}, timeout=6.0)
        except Exception as e:
            return {"ok": False, "blad": str(e)}
        if not dane.get("ok"):
            return {"ok": False, "blad": dane.get("blad", "Serwer bazy nie odpowiada.")}
        wersja = str(dane.get("wersja", ""))
        return {"ok": True, "nowsza": bool(wersja) and wersja != str(self._baza.wersja), "wersja": wersja, "lokalna": self._baza.wersja,
                "opis": f"Zmieniono {dane.get('zmieniono', '')}: {dane.get('opis', '')}".strip(": ")}

    def aktualizuj_baze(self):
        """Pełna aktualizacja bazy globalnej z serwera: pobrana baza zastępuje poprzednią, wpisy własne zostają. Przed zmianą kopia."""
        try:
            odp = wyslij_json(URL_BAZY, {"akcja": "baza", **({"token": konto.token()} if konto.token() else {})}, timeout=20.0)  # konto - do dziennika panelu
        except Exception as e:
            return {"ok": False, "blad": f"Nie udało się pobrać bazy: {e}"}
        wersja = str(odp.get("wersja", ""))
        nowe = {r: odp.get(r) for r in PLIKI}
        if not odp.get("ok") or not wersja or not all(isinstance(v, dict) for v in nowe.values()) or not nowe["materials"]:
            return {"ok": False, "blad": "Pobrana baza jest niepełna - zostaje dotychczasowa."}
        poprzednia = self._baza.wersja
        kopia = self._baza.kopia()
        zmiany = self._baza.zastap_globalna(nowe, wersja)
        return {"ok": True, "zmiany": zmiany, "kopia": kopia, "baza": self.baza(), "bylo": poprzednia, "jest": wersja}

    def zapisz_urzadzenie(self, rodzaj, stara_nazwa, nazwa, dane):
        if not nazwa.strip():
            return {"ok": False, "blad": "Podaj nazwę."}
        if nazwa.strip() != stara_nazwa and nazwa.strip() in self._baza.urzadzenia(rodzaj):
            return {"ok": False, "blad": "Pozycja o tej nazwie już jest w bazie."}
        wynik = {}
        for klucz, opis, _, rodzaj_pola in POLA_URZADZEN[rodzaj]:
            v = dane.get(klucz, "")
            try:
                if rodzaj_pola == "widmo":
                    wynik[klucz] = [obliczenia.liczba(x, 0.0) for x in _widmo(v)]
                elif rodzaj_pola is True:
                    wynik[klucz] = obliczenia.liczba(v, 0.0)
                elif rodzaj_pola == "jedn":
                    wynik[klucz] = v if v in obliczenia.JEDNOSTKI else "dBA"
                elif rodzaj_pola == "typ":
                    wynik[klucz] = v if v in TYPY_MATERIALU else TYPY_MATERIALU[0]
                else:
                    wynik[klucz] = str(v).strip()
            except ValueError:
                return {"ok": False, "blad": f"{opis}: wpisz liczbę."}
        self._baza.zapisz_urzadzenie(rodzaj, stara_nazwa, nazwa, wynik)
        return {"ok": True, "baza": self.baza()}

    def usun_urzadzenie(self, rodzaj, nazwa):
        self._baza.usun_urzadzenie(rodzaj, nazwa)
        return {"ok": True, "baza": self.baza()}

    def plik_bazy(self, id_):
        """Plik producenta z bazy (karta katalogowa): kopia lokalna albo pobranie z serwera (wymaga zalogowania)."""
        if not re.fullmatch(r"[0-9a-f]{16}\.(pdf|png|jpg|jpeg)", str(id_)):
            return {"ok": False, "blad": "Nieprawidłowy plik."}
        sciezka = os.path.join(KATALOG_PLIKOW, id_)
        if os.path.isfile(sciezka):
            return {"ok": True, "sciezka": sciezka}
        token = konto.token()
        if not token:
            return {"ok": False, "zaloguj": True, "blad": "Pliki producentów z bazy są dostępne po zalogowaniu (Ustawienia - Konto i dane)."}
        try:
            odp = wyslij_json(URL_BAZY, {"akcja": "plik", "id": id_, "token": token}, timeout=60.0)
        except Exception as e:
            return {"ok": False, "blad": f"Brak połączenia z serwerem: {e}"}
        if not odp.get("ok"):
            return {"ok": False, "zaloguj": bool(odp.get("zaloguj")), "blad": odp.get("blad", "Nie udało się pobrać pliku.")}
        os.makedirs(KATALOG_PLIKOW, exist_ok=True)
        with open(sciezka + ".tmp", "wb") as f:
            f.write(base64.b64decode(odp["plik"]))
        os.replace(sciezka + ".tmp", sciezka)
        return {"ok": True, "sciezka": sciezka}

    def otworz_plik_bazy(self, id_):
        wynik = self.plik_bazy(id_)
        if wynik["ok"]:
            os.startfile(wynik["sciezka"])
        return wynik

    def folder_bazy(self):
        os.startfile(os.path.join(KATALOG_DANYCH, "data"))
        return {"ok": True}

    # ------------------------------------------------------------------ projekt

    def nowy_projekt(self):
        autor = self._konto.stan()["inicjaly"] or self.program["autor"]  # inicjały z konta (imię i nazwisko w panelu)
        p = copy.deepcopy(NOWY_PROJEKT)
        p["meta"].update({"autor": autor, "jezyk_raportu": self.program["jezyk"]})
        p["revisions"] = [{"rev": "0", "desc": "Wydanie początkowe / Initial Issue", "drawn": autor, "check": "Appr.",
                           "date": datetime.now().strftime("%Y-%m-%d")}]
        p["_plik"] = ""
        return p

    def wczytaj_projekt(self, sciezka=""):
        if not sciezka:
            wybor = self._okno.create_file_dialog(webview.FileDialog.OPEN, file_types=("Projekt akustyki (*.json)", "Wszystkie pliki (*.*)"))
            if not wybor:
                return {"ok": False, "anulowano": True}
            sciezka = wybor[0]
        dane = _czytaj_json(sciezka, None)
        if not isinstance(dane, dict) or "dimensions" not in dane:
            return {"ok": False, "blad": "To nie jest projekt Kalkulatora Akustyki (wersja 1.0 lub nowsza)."}
        stan = uzupelnij_projekt(dane)
        stan["_plik"] = sciezka
        self._dodaj_ostatni(sciezka)
        brakujace = [os.path.basename(p) for p in self._sciezki_zalacznikow(stan) if not os.path.isfile(p)]
        return {"ok": True, "projekt": stan, "brakujace": brakujace}

    def zapisz_projekt(self, stan, jako=False):
        sciezka = stan.get("_plik") or ""
        if jako or not sciezka:
            nazwa = f"{stan.get('meta', {}).get('project_no') or 'Projekt'}_akustyka.json"
            wybor = self._okno.create_file_dialog(webview.FileDialog.SAVE, directory=os.path.dirname(sciezka) if sciezka else "",
                                                  save_filename=nazwa, file_types=("Projekt akustyki (*.json)",))
            if not wybor:
                return {"ok": False, "anulowano": True}
            sciezka = wybor if isinstance(wybor, str) else wybor[0]
            if not sciezka.lower().endswith(".json"):
                sciezka += ".json"
        dane = {k: v for k, v in stan.items() if not k.startswith("_")}
        dane["mod_dims"] = {k: {**v, "s": obliczenia.powierzchnie(dane)[k]} for k, v in dane.get("mod_dims", {}).items()}  # wersje 1.x czytają S z pliku
        try:
            _zapisz_json(sciezka, dane)
        except PermissionError:
            return {"ok": False, "blad": "Brak dostępu do pliku - wybierz inny folder."}
        self._dodaj_ostatni(sciezka)
        return {"ok": True, "plik": sciezka}

    def ustaw_niezapisane(self, wartosc):
        self._niezapisane = bool(wartosc)
        return True

    def zamknij(self):
        """Zamknięcie okna po decyzji użytkownika (zapisz / nie zapisuj)."""
        self._zamykam = True
        threading.Thread(target=self._okno.destroy, daemon=True).start()
        return True

    # ------------------------------------------------------------------ obliczenia i raport

    def oblicz(self, stan):
        try:
            return obliczenia.oblicz(stan)
        except Exception as e:
            return {"ok": False, "bledy": {"_": f"Błąd obliczeń: {e}"}}

    def sabine(self, l, w, h, alfa_podl, alfa_scian, q, r):
        try:
            return {"ok": True, **obliczenia.sabine(*(obliczenia.liczba(x) for x in (l, w, h, alfa_podl, alfa_scian, q, r)))}
        except ValueError as e:
            return {"ok": False, "blad": f"Pochłanianie wnętrza: {e}" if str(e) != "puste" else "Uzupełnij wymiary obudowy i parametry pochłaniania."}

    def wybierz_plik(self, rodzaj="obraz"):
        typy = ("Obrazy (*.png;*.jpg;*.jpeg;*.bmp;*.webp)",) if rodzaj == "obraz" else ("Dokument PDF (*.pdf)",)
        wybor = self._okno.create_file_dialog(webview.FileDialog.OPEN, file_types=typy)
        return wybor[0] if wybor else ""

    def obraz(self, sciezka):
        """Podgląd załącznika w oknie (data URI)."""
        if not sciezka or not os.path.isfile(sciezka):
            return {"ok": False, "blad": "Nie znaleziono pliku"}
        if os.path.getsize(sciezka) > OBRAZ_MAX_MB * 1024 * 1024:
            return {"ok": True, "sciezka": sciezka, "podglad": ""}
        typ = mimetypes.guess_type(sciezka)[0] or "image/png"
        return {"ok": True, "sciezka": sciezka, "podglad": f"data:{typ};base64,{base64.b64encode(open(sciezka, 'rb').read()).decode('ascii')}"}

    def wybierz_szablon(self):
        wybor = self._okno.create_file_dialog(webview.FileDialog.OPEN, file_types=("Dokument Word (*.docx)",))
        if wybor:
            self.ustaw_program("szablon", wybor[0])
        return self.program

    @staticmethod
    def _sciezki_zalacznikow(stan):
        m, rad = stan.get("meta", {}), stan.get("radiator", {})
        sciezki = [m.get("engine_wiz_path"), m.get("engine_pdf_path"), rad.get("wiz_path"), rad.get("pdf_path")]
        sciezki += [a.get("path") for lista in (stan.get("mod_attachments") or {}).values() for a in lista]
        return [p for p in sciezki if p]

    @staticmethod
    def _zalaczniki_raportu(stan, w):
        """Obrazy do treści raportu i karty PDF na koniec - tylko istniejące pliki, moduły / chłodnica - tylko aktywne."""
        m, rad = stan.get("meta", {}), stan.get("radiator", {})
        jest = lambda p: bool(p) and os.path.isfile(p)
        obrazy = {"engine": m.get("engine_wiz_path") if jest(m.get("engine_wiz_path")) else "",
                  "radiator": rad.get("wiz_path") if w["chlodnica"] and jest(rad.get("wiz_path")) else "", "moduly": {}}
        pdf = [p for p in (m.get("engine_pdf_path"), rad.get("pdf_path") if w["chlodnica"] else "") if jest(p)]
        for k in w["moduly"]:
            for a in (stan.get("mod_attachments") or {}).get(k, []):
                if jest(a.get("path")):
                    (obrazy["moduly"].setdefault(k, []) if a.get("type") != DOKUMENT else pdf).append(a["path"])
        return obrazy, pdf

    def generuj_raport(self, stan, format_="pdf"):
        """Raport PDF (przez Microsoft Word, z kartami producentów na końcu) albo DOCX; nagłówek i stopka z szablonu."""
        wyniki = self.oblicz(stan)
        if not wyniki["ok"]:
            return {"ok": False, "blad": "Uzupełnij dane: " + "; ".join(wyniki["bledy"].values())}
        m = stan.get("meta", {})
        rozszerzenie = ".pdf" if format_ == "pdf" else ".docx"
        plik_projektu = stan.get("_plik") or ""
        folder = os.path.dirname(plik_projektu) if plik_projektu else self.program["folder_raportow"]
        rew = (stan.get("revisions") or [{}])[-1]
        nazwa = re.sub(r'[\\/*?:"<>|]', "", f"Raport_{m.get('project_no') or 'Projekt'}_{m.get('agregat_typ') or 'Agregat'}_Rev{rew.get('rev', '0')}_"
                                              f"{(rew.get('date') or datetime.now().strftime('%Y-%m-%d')).replace('.', '-')}{rozszerzenie}".replace(" ", "_"))
        wybor = self._okno.create_file_dialog(webview.FileDialog.SAVE, directory=folder or "", save_filename=nazwa,
                                              file_types=("Dokument PDF (*.pdf)",) if format_ == "pdf" else ("Dokument Word (*.docx)",))
        if not wybor:
            return {"ok": False, "anulowano": True}
        sciezka = wybor if isinstance(wybor, str) else wybor[0]
        if not sciezka.lower().endswith(rozszerzenie):
            sciezka += rozszerzenie
        obrazy, karty = self._zalaczniki_raportu(stan, wyniki)
        szablon = self._szablon()
        if szablon and not os.path.isfile(szablon):
            return {"ok": False, "blad": f"Nie znaleziono szablonu raportu: {szablon} - wybierz go ponownie w Ustawieniach albo usuń."}
        pominiete = []
        try:
            if format_ == "pdf":
                tymczasowy = os.path.join(tempfile.gettempdir(), f"raport_akustyka_{os.getpid()}.docx")
                raport_docx.generuj(tymczasowy, stan, wyniki, obrazy, szablon)
                try:
                    raport_docx.zapisz_pdf(tymczasowy, sciezka)
                finally:
                    os.remove(tymczasowy)
                if karty and self.program["pdf_zalaczniki"]:
                    pominiete = raport_docx.dolacz_pdf(sciezka, karty)
            else:
                raport_docx.generuj(sciezka, stan, wyniki, obrazy, szablon)
        except PermissionError:
            return {"ok": False, "blad": "Plik raportu jest otwarty w innym programie - zamknij go i spróbuj ponownie."}
        except RuntimeError as e:
            return {"ok": False, "blad": str(e)}
        except Exception as e:
            return {"ok": False, "blad": f"Nie udało się utworzyć raportu: {e}"}
        if not plik_projektu:
            self.ustaw_program("folder_raportow", os.path.dirname(sciezka))
        if self.program["otworz_raport"]:
            os.startfile(sciezka)
        brak = [os.path.basename(p) for p in self._sciezki_zalacznikow(stan) if not os.path.isfile(p)]
        return {"ok": True, "plik": sciezka, "pominiete": brak + pominiete,
                "karty": len(karty) if format_ == "pdf" and self.program["pdf_zalaczniki"] else 0}

    # ------------------------------------------------------------------ pliki, strony, zgłoszenia

    def otworz_plik(self, sciezka):
        if os.path.isfile(sciezka):
            os.startfile(sciezka)
        return {"ok": True}

    def otworz_folder(self, sciezka):
        if os.path.isfile(sciezka):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(sciezka)])
        elif os.path.isdir(sciezka):
            os.startfile(sciezka)
        return {"ok": True}

    def otworz_strone(self, adres):
        import webbrowser
        webbrowser.open(adres)
        return {"ok": True}

    def wybierz_zalaczniki(self):
        wybor = self._okno.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=True, file_types=ZALACZNIKI_TYPY)
        return [{"sciezka": p, "nazwa": os.path.basename(p), "kb": max(1, round(os.path.getsize(p) / 1024))} for p in (wybor or []) if os.path.isfile(p)]

    def zglos(self, email, kategoria, nazwa, opis, projekt=None, zalaczniki=()):
        """
        Prośba o dodanie materiału / silnika / chłodnicy do bazy albo zgłoszenie błędu - formularz na diamel.pl.
        projekt: bieżący projekt dołączany jako plik .json (przy zgłoszeniu błędu).
        """
        zgloszenie = {"program": NAZWA_PROGRAMU, "wersja": VERSION, "email": email, "kategoria": kategoria, "nazwa": nazwa, "opis": opis}
        if projekt:
            tresc = json.dumps({k: v for k, v in projekt.items() if not k.startswith("_")}, ensure_ascii=False, indent=4).encode("utf-8")
            zgloszenie["zalaczniki"] = [{"nazwa": f"{projekt.get('meta', {}).get('project_no') or 'projekt'}_akustyka.json", "dane": base64.b64encode(tresc).decode("ascii")}]
        pliki = [p for p in dict.fromkeys(zalaczniki or []) if p and os.path.isfile(p)]
        if sum(os.path.getsize(p) for p in pliki) > ZALACZNIKI_MAX_MB * 1048576:
            return {"ok": False, "blad": f"Załączniki są większe niż {ZALACZNIKI_MAX_MB} MB razem - wybierz mniejsze pliki."}
        zgloszenie.setdefault("zalaczniki", []).extend(
            {"nazwa": os.path.basename(p), "dane": base64.b64encode(open(p, "rb").read()).decode("ascii")} for p in pliki)
        req = urllib.request.Request(URL_ZGLOSZEN, data=json.dumps(zgloszenie).encode("utf-8"), headers={**NAGLOWKI, "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=60) as odp:
                wynik = json.loads(odp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            try:
                wynik = json.loads(e.read().decode("utf-8"))
            except Exception:
                return {"ok": False, "blad": str(e)}
        except Exception as e:
            return {"ok": False, "blad": f"Brak połączenia z serwerem: {e}"}
        return {"ok": bool(wynik.get("ok")), "blad": wynik.get("blad", "")}
