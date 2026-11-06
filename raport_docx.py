"""
raport_docx.py

Raport akustyczny obudowy agregatu (PL / EN) w Wordzie, w stylu raportów Diamel (jak Kalkulator Rozruchu):
nagłówek z danymi projektu, kafelki z wynikiem każdego punktu pomiarowego, dane źródłowe z wykresem widma, budowa
obudowy, udział modułów w hałasie, opis metody, załączniki. Liczby z obliczenia.oblicz - te same, które pokazuje okno.

Szablon (opcjonalnie): plik .docx użytkownika - z niego zostają nagłówek, stopka, marginesy i rozmiar strony,
treść szablonu jest pomijana. PDF: ten sam dokument zapisany przez Microsoft Word (zapisz_pdf), opcjonalnie z kartami
producentów (PDF) dołączonymi na końcu (dolacz_pdf).
"""
import os
import tempfile
from datetime import datetime

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor, Emu

import obliczenia

CZCIONKA = "Segoe UI"
TEKST = RGBColor(0x11, 0x18, 0x27)
SZARY = RGBColor(0x6B, 0x72, 0x80)
AKCENT = RGBColor(0x03, 0x69, 0xA1)    # błękit Kalkulatora Akustyki (ciemny - czytelny po wydruku w odcieniach szarości)
ZIELONY = RGBColor(0x04, 0x78, 0x57)
CZERWONY = RGBColor(0xB9, 0x1C, 0x1C)
NIEBIESKI = RGBColor(0x1D, 0x4E, 0xD8)
TLO_SEKCJI = "EEF0F3"
TLO_NAGLOWKA = "F5F6F8"
LINIA = "D1D5DB"
KOLOR_WYKRESU = "#0369A1"

TEKSTY = {
    "Polski": {
        "eyebrow": "RAPORT AKUSTYCZNY", "tytul": "Analiza akustyczna obudowy agregatu",
        "agregat": "Agregat", "silnik": "silnik",
        "meta": ["Numer projektu", "Klient", "Data", "Przygotował", "Numer P.O.", "Dokument klienta"],
        "typy": {"Agregat otwarty": "Agregat otwarty (bez obudowy)", "Obudowa dźwiękochłonna": "Obudowa dźwiękochłonna (canopy)", "Obudowa kontenerowa": "Obudowa kontenerowa"},
        "moduly": {"sciana": "Ściany", "podloga": "Podłoga / rama", "sufit": "Sufit / dach", "drzwi": "Drzwi", "czerpnia": "Czerpnia powietrza (wlot)",
                   "wyrzut": "Wyrzut powietrza (wylot)", "wydech": "Wydech silnika (komin)"},
        "k_wymagane": "wymagane ≤ {l} dB(A) · {d} m", "k_zapas": "zapas {z} dB", "k_brak": "przekroczenie {z} dB",
        "ok": "Obudowa spełnia wymagania akustyczne we wszystkich punktach", "zle": "Obudowa nie spełnia wymagań akustycznych",
        "zle_pkt": "przekroczenie w punktach: {p}", "dominuje": "Największy udział w hałasie: {m} ({p} %).",
        "rewizje": "Historia zmian dokumentu", "rev_h": ["Rew.", "Opis zmiany", "Wykonał", "Sprawdził", "Data"],
        "s1": "Cel analizy i dane wejściowe", "s2": "Źródło hałasu - widmo silnika", "s3": "Pochłanianie i poziom ciśnienia wewnątrz obudowy",
        "s4": "Budowa obudowy - moduły i warstwy", "s5_chl": "Chłodnica", "s5_wyd": "Wydech silnika", "s6": "Wyniki w punktach pomiarowych", "s7": "Metoda obliczeń i założenia",
        "zal": "Załącznik",
        "cel": "Celem analizy jest określenie poziomu hałasu emitowanego przez agregat prądotwórczy w projektowanej obudowie i sprawdzenie wymagań akustycznych w wyznaczonych punktach pomiarowych. Pierwotnym źródłem jest hałas mechaniczny silnika; uwzględniono izolacyjność wszystkich przegród, czerpni i wyrzutu powietrza{dod}.",
        "cel_chl": ", hałas własny chłodnicy", "cel_wyd": ", wydech silnika",
        "parametr": "Parametr", "wartosc": "Wartość", "pasmo": "Pasmo [Hz]", "suma": "Suma",
        "typ": "Typ instalacji", "wymiary": "Wymiary obudowy (dł. × szer. × wys.)", "s_obudowy": "Powierzchnia promieniująca (ściany + dach)",
        "punkty": "Punkty pomiarowe", "punkt_w": "{n}: ≤ {l} dB(A) w odległości {d} m",
        "widmo_opis": "Widmo mocy akustycznej silnika Lw w pasmach oktawowych{z}. Całkowita moc akustyczna silnika:",
        "widmo_z_karty": " (karta producenta)", "lw_podane": "Lw [{j}]", "korekta_a": "Korekta A [dB]", "lwa": "LwA [dB(A)]",
        "lwa_suma": "Całkowita moc akustyczna silnika LwA", "lpa_suma": "Całkowity poziom ciśnienia wewnątrz obudowy LpA",
        "room_abs": "Pochłanianie wnętrza [dB]", "lpa": "LpA wewnątrz [dB(A)]",
        "sabine": "Pochłanianie wnętrza obliczono wzorem Sabine'a: Room Abs. = 10·lg(Q/(4πr²) + 4/R), gdzie R = S·ᾱ/(1−ᾱ) to stała pomieszczenia (S - powierzchnia wewnętrzna obudowy, ᾱ - średni współczynnik pochłaniania), Q - współczynnik kierunkowości źródła, r - średnia odległość od źródła.",
        "sabine_w": "Przyjęto: α podłogi {ap}, α ścian i sufitu {as_}, Q = {q}, r = {r} m.",
        "modul_w": "S = {s} m²", "modul_qty": "{q} szt.", "warstwy_brak": "bez warstw izolacyjnych (otwór)",
        "warstwa_rw": "Rw = {rw} dB, Ctr = {ctr} dB", "warstwa_widmo": "izolacyjność w pasmach", "grubosc": "grubość / długość {d} mm",
        "izol_tab": "Izolacyjność modułów w pasmach oktawowych [dB] (warstwy sumowane energetycznie)", "modul": "Moduł",
        "s_calk": "Całkowita powierzchnia promieniująca modułów: {s} m².",
        "chl": "Chłodnica na wyrzucie: {n}", "chl_lw": "Lw chłodnicy [{j}]", "chl_rdzen": "Tłumienie rdzenia [dB]", "chl_suma": "Moc akustyczna chłodnicy LwA",
        "wyd": "Liczba kominów: {q}, wylot {k}{h}", "wyd_lw": "Lw wydechu [{j}]", "wyd_suma": "Moc akustyczna wydechu LwA (1 szt.)",
        "pionowo": "pionowo w górę", "poziomo": "poziomo", "wysokosc": ", na wysokości {h} m",
        "wyd_lp": "Lp@1m od wylotu, bez tłumika [{j}]", "wyd_plus": "Przeliczenie na moc 10·lg(4π) [dB]", "wyd_lwj": "Lw wydechu [{j}]",
        "wyd_tl": "Tłumienie tłumika [dB]", "wyd_po": "LwA za tłumikiem [dB(A)]",
        "wyd_1m_wiersz": "LpA 1 m od wylotu, za tłumikiem [dB(A)]", "wyd_1m": "Hałas wydechu 1 m od wylotu (za tłumikiem, 1 komin): {v} dB(A)",
        "wyd_geo": "Wylot {nd} m nad dachem obudowy (koniec komina / kolana), {os} m od najbliższej ściany; punkt pomiaru {wp} m nad terenem.",
        "komin": "Po stronie komina (wydech jako źródło punktowe): wylot {h} m nad punktem, odległość {r} m - wydech {wy} dB(A), łącznie z pozostałymi modułami {wk} dB(A). Średnio wokół obudowy (ISO 3744): {ws} dB(A). O ocenie punktu decyduje wyższy poziom.",
        "pkt_tytul": "Punkt {n} - odległość {d} m", "s_pom": "Powierzchnia pomiarowa ISO 3744 (dł.+2d × szer.+2d × wys.+d, bez podłogi)",
        "korekta_odl": "Korekta odległości 10·lg(S_pom / S_obudowy)",
        "udzialy": ["Moduł", "S [m²]", "Lp przy powierzchni [dB(A)]", "Udział w punkcie [dB(A)]", "Udział [%]"],
        "wynik_pkt": "Poziom hałasu w punkcie", "wymagany": "Wymagany poziom",
        "pkt_ok": "Punkt {n}: {w} dB(A) - spełnia wymaganie ≤ {l} dB(A) (zapas {z} dB).",
        "pkt_zle": "Punkt {n}: {w} dB(A) - przekracza wymaganie ≤ {l} dB(A) o {z} dB. Największy udział: {m} - tę ścieżkę należy wytłumić w pierwszej kolejności.",
        "pasma_pkt": "Widmo hałasu w punkcie [dB(A)]",
        "rys_widmo": "Rys. {n}. Widmo: moc akustyczna silnika LwA, ciśnienie wewnątrz obudowy LpA i poziom w punkcie {p}",
        "rys_udzialy": "Rys. {n}. Udział modułów w hałasie w punkcie {p} ({w} dB(A))",
        "l_lwa": "LwA silnika", "l_lpa": "LpA wewnątrz", "l_pkt": "punkt {p}", "os_hz": "Pasmo oktawowe [Hz]",
        "metoda": [
            "Obliczenia prowadzono w 8 pasmach oktawowych 63 Hz - 8 kHz. Poziomy w decybelach sumowano energetycznie: L = 10·lg Σ 10^(Li/10).",
            "Korektę A (−26,2 / −16,1 / −8,6 / −3,2 / 0 / +1,2 / +1,0 / −1,1 dB) dodano tylko do widm podanych w dB; widma podane w dB(A) (zwykle karty silników) użyto bez zmian.",
            "Poziom ciśnienia wewnątrz obudowy: LpA = LwA + pochłanianie wnętrza (wzór Sabine'a).",
            "Moc akustyczna wypromieniowana przez moduł: W = S·10^((LpA − R)/10), gdzie S - powierzchnia modułu, R - izolacyjność modułu. Izolacyjność warstw jednej przegrody sumowano energetycznie (R = 10·lg Σ 10^(Ri/10)) - zachowawczo: wynik bliski najlepszej warstwie. Materiały podane jako Rw / Ctr przyjęto z wartością Rw + Ctr we wszystkich pasmach.",
            "Wydech: własne widmo mocy akustycznej wydechu pomniejszone o tłumienie tłumika, razy liczba kominów. Chłodnica na wyrzucie: hałas własny chłodnicy oraz hałas silnika przechodzący przez rdzeń chłodnicy, pomniejszone o izolacyjność wyrzutu.",
            "Wyloty skierowane pionowo w górę: −3 dB (kierunkowość). Wylot na wysokości h nad punktem pomiarowym: −20·lg(√(d² + h²)/d).",
            "Widmo wydechu podane jako poziom ciśnienia 1 m od wylotu (karta tłumika) przeliczono na moc akustyczną: Lw = Lp@1m + 10·lg(4π) ≈ Lp@1m + 11 dB (źródło punktowe w wolnym polu). Wydech oceniono dodatkowo po stronie komina jako źródło punktowe: Lp = LwA za tłumikiem − 10·lg(4π·r²) − 3 dB (wylot pionowo w górę), gdzie r to odległość wylotu od punktu; pozostałe moduły jak wyżej. O ocenie punktu decyduje wyższy z dwóch poziomów.",
            "Poziom w punkcie pomiarowym wg ISO 3744: Lp(d) = 10·lg(ΣW / S_pom), gdzie S_pom to powierzchnia prostopadłościanu pomiarowego (dł.+2d × szer.+2d × wys.+d, bez podłogi). Udział modułu w punkcie: 10·lg(W_modułu / S_pom) - udziały wszystkich modułów sumują się energetycznie do poziomu w punkcie.",
        ],
        "zalozenia": "Założenia zachowawcze: nie odjęto 6 dB przy przenikaniu przez przegrodę (EN 12354-4: na ścianę pada ćwierć energii pola rozproszonego) - wynik ma ok. 6 dB zapasu; warstwy przegrody sumowane energetycznie (w rzeczywistości izolacyjność przegrody wielowarstwowej jest zwykle większa); podłoga liczona jako promieniująca do punktu pomiarowego. Wynik nie obejmuje hałasu tła, odbić od pobliskich budynków ani drgań przenoszonych przez fundament.",
        "zal_nazwy": {"engine": "Widmo silnika", "radiator": "Chłodnica"},
        "strona": "Strona", "z_stron": "z", "stopka": "Raport akustyczny",
    },
    "English": {
        "eyebrow": "ACOUSTIC REPORT", "tytul": "Generator Enclosure Acoustic Analysis",
        "agregat": "Generating set", "silnik": "engine",
        "meta": ["Project no.", "Client", "Date", "Prepared by", "P.O. number", "Client doc. no."],
        "typy": {"Agregat otwarty": "Open set (unenclosed)", "Obudowa dźwiękochłonna": "Soundproof canopy", "Obudowa kontenerowa": "Containerised enclosure"},
        "moduly": {"sciana": "Walls", "podloga": "Floor / base frame", "sufit": "Roof", "drzwi": "Doors", "czerpnia": "Air intake (louvre)",
                   "wyrzut": "Air discharge (louvre)", "wydech": "Engine exhaust (stack)"},
        "k_wymagane": "required ≤ {l} dB(A) · {d} m", "k_zapas": "margin {z} dB", "k_brak": "exceeded by {z} dB",
        "ok": "The enclosure meets the acoustic requirements at all points", "zle": "The enclosure does not meet the acoustic requirements",
        "zle_pkt": "exceeded at: {p}", "dominuje": "Largest contribution to the noise: {m} ({p} %).",
        "rewizje": "Document revision history", "rev_h": ["Rev.", "Description of change", "Drawn", "Checked", "Date"],
        "s1": "Scope and input data", "s2": "Noise source - engine spectrum", "s3": "Absorption and sound pressure inside the enclosure",
        "s4": "Enclosure structure - modules and layers", "s5_chl": "Radiator", "s5_wyd": "Engine exhaust", "s6": "Results at the measurement points", "s7": "Calculation method and assumptions",
        "zal": "Appendix",
        "cel": "The aim of the analysis is to determine the noise level emitted by the generating set in the designed enclosure and to verify the acoustic requirements at the designated measurement points. The primary source is the mechanical noise of the engine; the sound insulation of all partitions, the air intake and the air discharge is included{dod}.",
        "cel_chl": ", radiator self-noise", "cel_wyd": ", engine exhaust",
        "parametr": "Parameter", "wartosc": "Value", "pasmo": "Band [Hz]", "suma": "Total",
        "typ": "Installation type", "wymiary": "Enclosure dimensions (L × W × H)", "s_obudowy": "Radiating surface (walls + roof)",
        "punkty": "Measurement points", "punkt_w": "{n}: ≤ {l} dB(A) at {d} m",
        "widmo_opis": "Engine sound power level spectrum Lw in octave bands{z}. Total engine sound power level:",
        "widmo_z_karty": " (manufacturer's data sheet)", "lw_podane": "Lw [{j}]", "korekta_a": "A-weighting [dB]", "lwa": "LwA [dB(A)]",
        "lwa_suma": "Total engine sound power level LwA", "lpa_suma": "Total sound pressure level inside the enclosure LpA",
        "room_abs": "Room absorption [dB]", "lpa": "LpA inside [dB(A)]",
        "sabine": "Room absorption was calculated with Sabine's formula: Room Abs. = 10·lg(Q/(4πr²) + 4/R), where R = S·ᾱ/(1−ᾱ) is the room constant (S - internal surface of the enclosure, ᾱ - mean absorption coefficient), Q - source directivity factor, r - mean distance from the source.",
        "sabine_w": "Assumed: floor α {ap}, walls and roof α {as_}, Q = {q}, r = {r} m.",
        "modul_w": "S = {s} m²", "modul_qty": "{q} pcs.", "warstwy_brak": "no insulating layers (opening)",
        "warstwa_rw": "Rw = {rw} dB, Ctr = {ctr} dB", "warstwa_widmo": "band insulation", "grubosc": "thickness / length {d} mm",
        "izol_tab": "Sound insulation of the modules in octave bands [dB] (layers added energetically)", "modul": "Module",
        "s_calk": "Total radiating surface of the modules: {s} m².",
        "chl": "Radiator on the air discharge: {n}", "chl_lw": "Radiator Lw [{j}]", "chl_rdzen": "Core attenuation [dB]", "chl_suma": "Radiator sound power level LwA",
        "wyd": "Number of stacks: {q}, outlet {k}{h}", "wyd_lw": "Exhaust Lw [{j}]", "wyd_suma": "Exhaust sound power level LwA (1 pc.)",
        "pionowo": "vertical (upwards)", "poziomo": "horizontal", "wysokosc": ", at {h} m height",
        "wyd_lp": "Lp@1m from the outlet, unsilenced [{j}]", "wyd_plus": "Conversion to sound power 10·lg(4π) [dB]", "wyd_lwj": "Exhaust Lw [{j}]",
        "wyd_tl": "Silencer attenuation [dB]", "wyd_po": "LwA after the silencer [dB(A)]",
        "wyd_1m_wiersz": "LpA 1 m from the outlet, after the silencer [dB(A)]", "wyd_1m": "Engine exhaust noise 1 m from the outlet (after the silencer, 1 stack): {v} dB(A)",
        "wyd_geo": "Outlet {nd} m above the enclosure roof (end of the stack / elbow), {os} m from the nearest wall; measurement point {wp} m above ground.",
        "komin": "On the stack side (exhaust as a point source): outlet {h} m above the point, distance {r} m - exhaust {wy} dB(A), together with the other modules {wk} dB(A). Average around the enclosure (ISO 3744): {ws} dB(A). The higher level governs the assessment of the point.",
        "pkt_tytul": "Point {n} - distance {d} m", "s_pom": "ISO 3744 measurement surface (L+2d × W+2d × H+d, without floor)",
        "korekta_odl": "Distance correction 10·lg(S_meas / S_enclosure)",
        "udzialy": ["Module", "S [m²]", "Lp at surface [dB(A)]", "Contribution at point [dB(A)]", "Share [%]"],
        "wynik_pkt": "Noise level at the point", "wymagany": "Required level",
        "pkt_ok": "Point {n}: {w} dB(A) - meets the requirement ≤ {l} dB(A) (margin {z} dB).",
        "pkt_zle": "Point {n}: {w} dB(A) - exceeds the requirement ≤ {l} dB(A) by {z} dB. Largest contribution: {m} - this path should be attenuated first.",
        "pasma_pkt": "Noise spectrum at the point [dB(A)]",
        "rys_widmo": "Fig. {n}. Spectrum: engine sound power LwA, sound pressure inside the enclosure LpA and level at point {p}",
        "rys_udzialy": "Fig. {n}. Contribution of the modules to the noise at point {p} ({w} dB(A))",
        "l_lwa": "Engine LwA", "l_lpa": "LpA inside", "l_pkt": "point {p}", "os_hz": "Octave band [Hz]",
        "metoda": [
            "The calculation was carried out in 8 octave bands 63 Hz - 8 kHz. Decibel levels were added energetically: L = 10·lg Σ 10^(Li/10).",
            "The A-weighting (−26.2 / −16.1 / −8.6 / −3.2 / 0 / +1.2 / +1.0 / −1.1 dB) was applied only to spectra given in dB; spectra given in dB(A) (usually engine data sheets) were used as they are.",
            "Sound pressure level inside the enclosure: LpA = LwA + room absorption (Sabine's formula).",
            "Sound power radiated by a module: W = S·10^((LpA − R)/10), where S is the module area and R its sound insulation. The insulation of the layers of one partition was added energetically (R = 10·lg Σ 10^(Ri/10)) - conservatively: the result is close to the best layer. Materials given as Rw / Ctr were taken with the value Rw + Ctr in all bands.",
            "Exhaust: the exhaust sound power spectrum less the silencer attenuation, times the number of stacks. Radiator on the air discharge: radiator self-noise plus the engine noise passing through the radiator core, less the insulation of the discharge.",
            "Outlets directed vertically upwards: −3 dB (directivity). Outlet at height h above the measurement point: −20·lg(√(d² + h²)/d).",
            "The exhaust spectrum given as sound pressure 1 m from the outlet (silencer data sheet) was converted to sound power: Lw = Lp@1m + 10·lg(4π) ≈ Lp@1m + 11 dB (point source, free field). The exhaust was additionally assessed on the stack side as a point source: Lp = LwA after the silencer − 10·lg(4π·r²) − 3 dB (vertical outlet), r being the distance from the outlet to the point; other modules as above. The higher of the two levels governs the assessment of the point.",
            "Level at the measurement point per ISO 3744: Lp(d) = 10·lg(ΣW / S_meas), where S_meas is the area of the measurement box (L+2d × W+2d × H+d, without floor). Module contribution at the point: 10·lg(W_module / S_meas) - the contributions of all modules add up energetically to the level at the point.",
        ],
        "zalozenia": "Conservative assumptions: the 6 dB transmission term (EN 12354-4: a quarter of the diffuse field energy is incident on the wall) was not subtracted - the result has a margin of about 6 dB; partition layers added energetically (in reality the insulation of a multi-layer partition is usually higher); the floor is treated as radiating to the measurement point. The result does not include background noise, reflections from nearby buildings or vibration transmitted through the foundation.",
        "zal_nazwy": {"engine": "Engine spectrum", "radiator": "Radiator"},
        "strona": "Page", "z_stron": "of", "stopka": "Acoustic report",
    },
}

# ------------------------------------------------------------------ pomocnicze: formatowanie i XML Worda

class Liczby:
    """Liczby w języku raportu: 1 014,23 (PL) / 1,014.23 (EN)."""

    def __init__(self, jezyk):
        self.pl = jezyk == "Polski"

    def _jezyk(self, t):
        return t.replace(",", " ").replace(".", ",") if self.pl else t

    def __call__(self, x, miejsca=2):
        return "—" if x is None else self._jezyk(f"{x:,.{miejsca}f}")

    def g(self, x):
        """Bez zbędnych zer: 12.5 -> 12,5; 95.0 -> 95."""
        t = f"{x:,.4f}"
        return self._jezyk(t.rstrip("0").rstrip(".") if "." in t else t)


def _tlo(komorka, kolor):
    tc = komorka._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), kolor)
    tc.append(shd)


def _ramki_komorki(komorka, **boki):
    """boki: top/bottom/left/right = (grubość w 1/8 pt, kolor) albo None (bez ramki)."""
    tcPr = komorka._tc.get_or_add_tcPr()
    ramki = OxmlElement("w:tcBorders")
    for bok in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{bok}")
        wartosc = boki.get(bok)
        if wartosc:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(wartosc[0]))
            el.set(qn("w:color"), wartosc[1])
        else:
            el.set(qn("w:val"), "nil")
        ramki.append(el)
    tcPr.append(ramki)


def _bez_ramek(tabela):
    tblPr = tabela._tbl.tblPr
    ramki = OxmlElement("w:tblBorders")
    for bok in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{bok}")
        el.set(qn("w:val"), "nil")
        ramki.append(el)
    tblPr.append(ramki)


def _marginesy_komorek(tabela, gora=40, dol=40, lewo=90, prawo=90):
    tblPr = tabela._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for bok, w in (("top", gora), ("left", lewo), ("bottom", dol), ("right", prawo)):
        el = OxmlElement(f"w:{bok}")
        el.set(qn("w:w"), str(w))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _ramka_akapitu(akapit, bok, grubosc, kolor, odstep=4):
    pPr = akapit._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    el = OxmlElement(f"w:{bok}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(grubosc))
    el.set(qn("w:space"), str(odstep))
    el.set(qn("w:color"), kolor)
    pBdr.append(el)


def _tlo_akapitu(akapit, kolor):
    pPr = akapit._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), kolor)
    pPr.append(shd)


def _pole(akapit, kod, rozmiar, kolor):
    """Pole Worda (PAGE, NUMPAGES) - numer strony w stopce."""
    run = akapit.add_run()
    _styl_runu(run, rozmiar, kolor=kolor)
    for typ, tekst in (("begin", None), (None, kod), ("separate", None), (None, "1"), ("end", None)):
        if typ:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), typ)
        elif tekst == kod:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = f" {kod} "
        else:
            el = OxmlElement("w:t")
            el.text = tekst
        run._r.append(el)


def _styl_runu(run, rozmiar=9.5, pogrubienie=False, kolor=TEKST, kursywa=False, czcionka=CZCIONKA):
    run.font.name = czcionka
    run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cs"), czcionka)
    run.font.size = Pt(rozmiar)
    run.font.bold = pogrubienie
    run.font.italic = kursywa
    run.font.color.rgb = kolor
    return run


def _odstepy(akapit, przed=0, po=0, interlinia=1.1):
    f = akapit.paragraph_format
    f.space_before = Pt(przed)
    f.space_after = Pt(po)
    f.line_spacing = interlinia




def meta(stan):
    return stan.get("meta") or {}


def autor(stan):
    m = meta(stan)
    rew = stan.get("revisions") or []
    return m.get("autor") or (rew[-1].get("drawn", "") if rew else "")


class Raport:
    def __init__(self, stan, w, obrazy, szablon=""):
        self.stan, self.w, self.obrazy = stan, w, obrazy
        jezyk = meta(stan).get("jezyk_raportu")
        self.jezyk = jezyk if jezyk in TEKSTY else "Polski"
        self.en = self.jezyk == "English"
        self.t = TEKSTY[self.jezyk]
        self.n = Liczby(self.jezyk)
        self.szablon = bool(szablon and os.path.isfile(szablon))
        self.doc = Document(szablon) if self.szablon else Document()
        self.nr_rys = 0
        self.tymczasowe = []
        self._przygotuj()

    # -------------------------------------------------------------- dokument, strona, stopka

    def _przygotuj(self):
        body = self.doc.element.body
        for el in list(body):
            if el.tag != qn("w:sectPr"):
                body.remove(el)  # treść szablonu pomijana - zostaje nagłówek, stopka i ustawienia strony
        normal = self.doc.styles["Normal"]
        normal.font.name = CZCIONKA
        normal.font.size = Pt(9.5)
        normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cs"), CZCIONKA)
        sekcja = self.doc.sections[0]
        if self.szablon:
            self._wypelnij_pola_szablonu()
        else:
            sekcja.page_width, sekcja.page_height = Cm(21), Cm(29.7)
            sekcja.left_margin = sekcja.right_margin = Cm(1.8)
            sekcja.top_margin, sekcja.bottom_margin = Cm(1.6), Cm(1.8)
            sekcja.footer_distance = Cm(0.9)
            self._stopka(sekcja)
        self.szer = sekcja.page_width - sekcja.left_margin - sekcja.right_margin

    def tytul(self):
        return meta(self.stan).get("project_name") or self.t["tytul"]

    def _wypelnij_pola_szablonu(self):
        """Pola w nagłówku / stopce szablonu: {PROJEKT} {KLIENT} {AGREGAT} {SILNIK} {AUTOR} {DATA} {TYTUL}."""
        m = meta(self.stan)
        pola = {"{PROJEKT}": m.get("project_no", ""), "{KLIENT}": m.get("client_name", ""), "{AGREGAT}": m.get("agregat_typ", ""),
                "{SILNIK}": m.get("silnik_typ", ""), "{AUTOR}": autor(self.stan), "{DATA}": datetime.now().strftime("%Y-%m-%d"),
                "{TYTUL}": self.tytul(), "{STRONA}": self.t["strona"], "{Z}": self.t["z_stron"]}

        def zamien(tekst):
            for k, v in pola.items():
                tekst = tekst.replace(k, v)
            return tekst
        for sekcja in self.doc.sections:
            for czesc in (sekcja.header, sekcja.footer, sekcja.first_page_header, sekcja.first_page_footer):
                akapity = list(czesc.paragraphs) + [p for tab in czesc.tables for w in tab.rows for k in w.cells for p in k.paragraphs]
                for p in akapity:
                    for r in p.runs:  # najpierw w obrębie fragmentów - zostaje formatowanie i pola Worda (numer strony)
                        if "{" in r.text:
                            r.text = zamien(r.text)
                    tekst = "".join(r.text for r in p.runs)
                    if not any(k in tekst for k in pola) or p._p.find(".//" + qn("w:fldChar")) is not None:
                        continue
                    p.runs[0].text = zamien(tekst)  # pole rozbite przez Worda na kilka fragmentów - formatowanie pierwszego
                    for r in p.runs[1:]:
                        r.text = ""

    def _stopka(self, sekcja):
        p = sekcja.footer.paragraphs[0]
        tabulatory = p.paragraph_format.tab_stops
        for poz in (Inches(3.25), Inches(6.5)):  # tabulatory stylu "Stopka" Worda (środek, prawo) - numer strony ma być przy prawym marginesie
            tabulatory.add_tab_stop(poz, WD_TAB_ALIGNMENT.CLEAR)
        tabulatory.add_tab_stop(sekcja.page_width - sekcja.left_margin - sekcja.right_margin, WD_TAB_ALIGNMENT.RIGHT)
        _ramka_akapitu(p, "top", 4, LINIA, 6)
        nazwa = " · ".join(x for x in (meta(self.stan).get("project_no", ""), self.t["stopka"]) if x)
        _styl_runu(p.add_run(f"{nazwa}\t{self.t['strona']} "), 7.5, kolor=SZARY)
        _pole(p, "PAGE", 7.5, SZARY)
        _styl_runu(p.add_run(f" {self.t['z_stron']} "), 7.5, kolor=SZARY)
        _pole(p, "NUMPAGES", 7.5, SZARY)

    def akapit(self, tekst="", rozmiar=9.5, pogrubienie=False, kolor=TEKST, przed=0, po=4, wyrownanie=None, kursywa=False):
        p = self.doc.add_paragraph()
        _odstepy(p, przed, po)
        if wyrownanie:
            p.alignment = wyrownanie
        if tekst:
            _styl_runu(p.add_run(tekst), rozmiar, pogrubienie, kolor, kursywa)
        return p

    def tabela(self, wiersze, kolumny):
        tab = self.doc.add_table(rows=wiersze, cols=kolumny)
        tab.alignment = WD_TABLE_ALIGNMENT.CENTER
        tab.autofit = False
        _bez_ramek(tab)
        _marginesy_komorek(tab)
        return tab

    @staticmethod
    def komorka(komorka, tekst, rozmiar=9, pogrubienie=False, kolor=TEKST, wyrownanie=None, kursywa=False):
        p = komorka.paragraphs[0]
        _odstepy(p, 0, 0, 1.05)
        if wyrownanie:
            p.alignment = wyrownanie
        _styl_runu(p.add_run(tekst), rozmiar, pogrubienie, kolor, kursywa)
        komorka.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        return p

    def szerokosci(self, tab, proporcje):
        suma = sum(proporcje)
        for wiersz in tab.rows:
            for k, p in zip(wiersz.cells, proporcje):
                k.width = Emu(int(self.szer * p / suma))

    def nazwa_modulu(self, k):
        return self.t["moduly"].get(k, k)

    def nazwa_warstwy(self, w):
        return (w.get("name_en") if self.en and w.get("name_en") else w.get("name")) or "—"

    def opis_warstwy(self, w):
        return (w.get("desc_en") if self.en and w.get("desc_en") else w.get("desc_pl")) or ""

    # -------------------------------------------------------------- elementy raportu

    def naglowek(self):
        t, m = self.t, meta(self.stan)
        p = self.akapit(t["eyebrow"], 7.5, kolor=AKCENT, po=2)
        p.runs[0].font.bold = True
        self.akapit(self.tytul(), 20, True, po=2)
        podtytul = " · ".join(x for x in (m.get("agregat_typ") and f"{t['agregat']} {m['agregat_typ']}", m.get("silnik_typ") and f"{t['silnik']} {m['silnik_typ']}",
                                          t["typy"].get(m.get("typ_obudowy"), m.get("typ_obudowy"))) if x)
        p = self.akapit(podtytul or " ", 10.5, kolor=SZARY, po=8)
        _ramka_akapitu(p, "bottom", 12, "111827", 6)

        tab = self.tabela(2, 3)
        wartosci = [m.get("project_no", ""), m.get("client_name", ""), datetime.now().strftime("%Y-%m-%d"), autor(self.stan), m.get("po_number", ""), m.get("client_doc_no", "")]
        for i, (etykieta, wartosc) in enumerate(zip(t["meta"], wartosci)):
            k = tab.cell(i // 3, i % 3)
            self.komorka(k, etykieta.upper(), 6.5, kolor=SZARY)
            p = k.add_paragraph()
            _odstepy(p, 1, 2)
            _styl_runu(p.add_run(wartosc or "—"), 10, True)
        self.szerokosci(tab, [1, 1, 1])
        self.akapit(po=6)

    def kafelki(self):
        t, w, n = self.t, self.w, self.n
        punkty = w["punkty"][:4]
        tab = self.tabela(1, len(punkty))
        _marginesy_komorek(tab, 90, 90, 130, 130)
        for k, pt in zip(tab.rows[0].cells, punkty):
            _ramki_komorki(k, top=(6, "9CA3AF"), bottom=(6, "9CA3AF"), left=(6, "9CA3AF"), right=(6, "9CA3AF"))
            self.komorka(k, (pt["name"] or f"{n.g(pt['d'])} m").upper(), 6.5, kolor=SZARY)
            p = k.add_paragraph()
            _odstepy(p, 2, 1)
            _styl_runu(p.add_run(f"{n(pt['wynik'], 1)} dB(A)"), 14, True)
            symbol, kolor = ("✓", ZIELONY) if pt["zgodny"] else ("✗", CZERWONY)
            _styl_runu(p.add_run(f"  {symbol}"), 11, True, kolor, czcionka="Segoe UI Symbol")
            p = k.add_paragraph()
            _odstepy(p, 0, 0)
            _styl_runu(p.add_run(t["k_wymagane"].format(l=n.g(pt["limit"]), d=n.g(pt["d"]))), 7.5, kolor=SZARY)
            p = k.add_paragraph()
            _odstepy(p, 0, 0)
            _styl_runu(p.add_run((t["k_zapas"] if pt["zgodny"] else t["k_brak"]).format(z=n(abs(pt["zapas"]), 1))), 7.5, kolor=ZIELONY if pt["zgodny"] else CZERWONY)
        self.szerokosci(tab, [1] * len(punkty))

        zle = [pt["name"] or f"{n.g(pt['d'])} m" for pt in w["punkty"] if not pt["zgodny"]]
        p = self.akapit(przed=10, po=2)
        _ramka_akapitu(p, "left", 24, "047857" if not zle else "B91C1C", 8)
        _styl_runu(p.add_run("✓  " if not zle else "✗  "), 11, True, ZIELONY if not zle else CZERWONY, czcionka="Segoe UI Symbol")
        _styl_runu(p.add_run(t["ok"] if not zle else t["zle"]), 11, True)
        if zle:
            _styl_runu(p.add_run(" - " + t["zle_pkt"].format(p=", ".join(zle))), 9.5, kolor=CZERWONY)
        pt0 = w["punkty"][0]
        if pt0["dominujacy"]:
            p = self.akapit(t["dominuje"].format(m=self.nazwa_modulu(pt0["dominujacy"]).lower() if not self.en else self.nazwa_modulu(pt0["dominujacy"]),
                                                 p=n(pt0["moduly"][pt0["dominujacy"]]["procent"], 0)), 9, kolor=SZARY, po=10)
            p.paragraph_format.left_indent = Cm(0.45)

    def rewizje(self):
        rew = self.stan.get("revisions") or []
        if not rew:
            return
        self.akapit(self.t["rewizje"].upper(), 7, True, SZARY, przed=2, po=3)
        tab = self.tabela(1 + len(rew), 5)
        for i, h in enumerate(self.t["rev_h"]):
            k = tab.cell(0, i)
            self.komorka(k, h.upper(), 6.5, True, SZARY)
            _ramki_komorki(k, bottom=(6, "111827"))
        for r_i, r in enumerate(rew, 1):
            for j, klucz in enumerate(("rev", "desc", "drawn", "check", "date")):
                k = tab.cell(r_i, j)
                self.komorka(k, str(r.get(klucz, "")), 8.5)
                _ramki_komorki(k, bottom=(4, LINIA))
        self.szerokosci(tab, [0.6, 4.4, 1, 1, 1.2])

    def sekcja(self, nr, tytul, nowa_strona=False):
        p = self.akapit(przed=4 if nowa_strona else 14, po=6)
        if nowa_strona:
            p.paragraph_format.page_break_before = True
        p.paragraph_format.keep_with_next = True
        _tlo_akapitu(p, TLO_SEKCJI)
        _ramka_akapitu(p, "left", 30, "111827", 6)
        _styl_runu(p.add_run(f"{nr}   "), 11, True, AKCENT)
        _styl_runu(p.add_run(tytul), 11, True)

    def podtytul(self, tekst):
        p = self.akapit(tekst, 10, True, przed=10, po=4)
        p.paragraph_format.keep_with_next = True
        return p

    def tabela_prosta(self, naglowki, wiersze, proporcje, pogrubiony_wiersz=None, wyrownaj_od=1):
        """Tabela z nagłówkiem i liczbami do prawej (od kolumny wyrownaj_od); pogrubiony_wiersz - indeks wiersza sumy."""
        tab = self.tabela(1 + len(wiersze), len(naglowki))
        for i, h in enumerate(naglowki):
            k = tab.cell(0, i)
            self.komorka(k, h.upper(), 6.5, True, SZARY, WD_ALIGN_PARAGRAPH.RIGHT if i >= wyrownaj_od else None)
            _ramki_komorki(k, bottom=(6, "111827"))
        for r, wiersz in enumerate(wiersze, 1):
            for i, tekst in enumerate(wiersz):
                k = tab.cell(r, i)
                self.komorka(k, str(tekst), 8.5, r - 1 == pogrubiony_wiersz, wyrownanie=WD_ALIGN_PARAGRAPH.RIGHT if i >= wyrownaj_od else None)
                _ramki_komorki(k, bottom=(4, LINIA))
                if r - 1 == pogrubiony_wiersz:
                    _tlo(k, TLO_NAGLOWKA)
        self.szerokosci(tab, proporcje)
        if len(tab.rows) <= 12:  # mała tabela nie dzieli się między strony
            for wiersz in tab.rows[:-1]:
                for k in wiersz.cells:
                    for p in k.paragraphs:
                        p.paragraph_format.keep_with_next = True
        return tab

    def tabela_pasm(self, wiersze, pierwsza="", suma=True):
        """Wiersze: (nazwa, [8 wartości], suma albo None) - tabela pasm oktawowych z kolumną sumy."""
        n = self.n
        nag = [pierwsza or self.t["pasmo"]] + obliczenia.PASMA + ([self.t["suma"]] if suma else [])
        dane = [[nazwa] + [n(x, 1) for x in wart] + ([n(s, 1) if s is not None else ""] if suma else []) for nazwa, wart, s in wiersze]
        return self.tabela_prosta(nag, dane, [2.6] + [1] * 8 + ([1.15] if suma else []))

    def ocena(self, ok, tekst):
        """Wiersz z oceną: ✓ zielony, ✗ czerwony, i niebieski (informacyjnie)."""
        symbol, kolor = {True: ("✓", ZIELONY), False: ("✗", CZERWONY), "info": ("i", NIEBIESKI), None: ("–", SZARY)}[ok]
        p = self.akapit(przed=6, po=4)
        p.paragraph_format.left_indent = Cm(0.1)
        _styl_runu(p.add_run(f"{symbol}  "), 10, True, kolor, czcionka="Segoe UI Symbol")
        _styl_runu(p.add_run(tekst), 9.5, ok is False, kolor if ok is False else TEKST)

    def opis(self, tekst, po=6):
        self.akapit(tekst, 9, przed=2, po=po)

    def rysunek(self, sciezka, opis, szerokosc=None):
        self.nr_rys += 1
        p = self.akapit(przed=4, po=2, wyrownanie=WD_ALIGN_PARAGRAPH.CENTER)
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(sciezka, width=szerokosc or self.szer)
        self.akapit(opis.format(n=self.nr_rys), 8, kolor=SZARY, po=8, wyrownanie=WD_ALIGN_PARAGRAPH.CENTER, kursywa=True)

    def obraz(self, sciezka, szerokosc=None):
        try:
            p = self.akapit(przed=4, po=6, wyrownanie=WD_ALIGN_PARAGRAPH.CENTER)
            p.add_run().add_picture(sciezka, width=szerokosc or Emu(int(self.szer * 0.85)))
        except Exception:  # format nieobsługiwany przez Worda - przez PIL do PNG
            from PIL import Image
            tymczasowy = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
            Image.open(sciezka).convert("RGB").save(tymczasowy)
            self.tymczasowe.append(tymczasowy)
            p.add_run().add_picture(tymczasowy, width=szerokosc or Emu(int(self.szer * 0.85)))

    # -------------------------------------------------------------- sekcje

    def cel(self):
        t, w, n, m = self.t, self.w, self.n, meta(self.stan)
        dod = (t["cel_chl"] if w["chlodnica"] else "") + (t["cel_wyd"] if "wydech" in w["moduly"] else "")
        self.opis(t["cel"].format(dod=dod))
        wym = self.stan["dimensions"]
        wiersze = [(t["typ"], t["typy"].get(m.get("typ_obudowy"), m.get("typ_obudowy") or "—")),
                   (t["wymiary"], f"{n.g(obliczenia.liczba(wym['kont_l']))} × {n.g(obliczenia.liczba(wym['kont_w']))} × {n.g(obliczenia.liczba(wym['kont_h']))} m"),
                   (t["s_obudowy"], f"{n(w['s_kont'])} m²")]
        wiersze += [(t["punkty"] if i == 0 else "", t["punkt_w"].format(n=pt["name"] or f"{n.g(pt['d'])} m", l=n.g(pt["limit"]), d=n.g(pt["d"])))
                    for i, pt in enumerate(w["punkty"])]
        self.tabela_prosta([t["parametr"], t["wartosc"]], wiersze, [2.2, 3.8], wyrownaj_od=9)

    def zrodlo(self):
        t, w, n = self.t, self.w, self.n
        dB = w["jedn_silnika"] == "dB"
        self.opis(t["widmo_opis"].format(z=t["widmo_z_karty"] if meta(self.stan).get("silnik_typ") else "") + f" {n(w['lw_a_suma'], 1)} dB(A).")
        wiersze = [(t["lw_podane"].format(j="dB" if dB else "dB(A)"), w["lw_silnika"], None)]
        if dB:
            wiersze.append((t["korekta_a"], obliczenia.KOREKTA_A, None))
        wiersze.append((t["lwa"], w["lw_a"], w["lw_a_suma"]))
        self.tabela_pasm(wiersze)
        self.rysunek(rysuj_widmo(w, t, n), t["rys_widmo"].replace("{p}", w["punkty"][0]["name"] or f"{n.g(w['punkty'][0]['d'])} m"))
        if self.obrazy.get("engine"):
            self.obraz(self.obrazy["engine"])

    def wnetrze(self):
        t, w, n = self.t, self.w, self.n
        sab = self.stan.get("sabine") or {}
        self.opis(t["sabine"])
        if sab:
            self.opis(t["sabine_w"].format(ap=n.g(obliczenia.liczba(sab.get("alfa_podl"), 0)), as_=n.g(obliczenia.liczba(sab.get("alfa_scian"), 0)),
                                           q=n.g(obliczenia.liczba(sab.get("sabine_q"), 0)), r=n.g(obliczenia.liczba(sab.get("sabine_r"), 0))), po=4)
        self.tabela_pasm([(t["lwa"], w["lw_a"], w["lw_a_suma"]), (t["room_abs"], w["room_abs"], None), (t["lpa"], w["lp_a"], w["lp_a_suma"])])
        self.ocena("info", f"{t['lpa_suma']}: {n(w['lp_a_suma'], 1)} dB(A)")

    def budowa(self):
        t, w, n = self.t, self.w, self.n
        warstwy, zalaczniki = self.stan.get("layers") or {}, self.obrazy.get("moduly") or {}
        for k, m in w["moduly"].items():
            opis = t["modul_qty"].format(q=m["qty"]) if k == "wydech" else t["modul_w"].format(s=n(m["s"])) + (f" · {t['modul_qty'].format(q=m['qty'])}" if m["qty"] > 1 and k not in obliczenia.AUTO_POWIERZCHNIA else "")
            p = self.podtytul(self.nazwa_modulu(k))
            _styl_runu(p.add_run(f"   {opis}"), 9, kolor=SZARY)
            lista = warstwy.get(k) or []
            if not lista:
                self.akapit(f"•  {t['warstwy_brak']}", 9, kolor=SZARY, po=2).paragraph_format.left_indent = Cm(0.4)
            for wa in lista:
                dane = t["warstwa_rw"].format(rw=n.g(obliczenia.liczba(wa.get("rw"), 0)), ctr=n.g(obliczenia.liczba(wa.get("ctr"), 0))) if wa.get("type") == "Rw / Ctr" else t["warstwa_widmo"]
                d = obliczenia.liczba(wa.get("d"), 0)
                tekst = " · ".join(x for x in (self.opis_warstwy(wa), dane, t["grubosc"].format(d=n.g(d)) if d else "") if x)
                p = self.akapit(po=1)
                p.paragraph_format.left_indent = Cm(0.4)
                _styl_runu(p.add_run(f"•  {self.nazwa_warstwy(wa)}"), 9, True)
                _styl_runu(p.add_run(f" - {tekst}"), 9, kolor=SZARY)
            for sciezka in zalaczniki.get(k, []):
                self.obraz(sciezka)
        self.podtytul(t["izol_tab"])
        self.tabela_pasm([(self.nazwa_modulu(k), m["r"], None) for k, m in w["moduly"].items()], t["modul"], suma=False)
        self.akapit(t["s_calk"].format(s=n(w["s_calkowita"])), 9, kolor=SZARY, przed=4)

    def chlodnica(self):
        t, w, n, s = self.t, self.w, self.n, self.stan
        if w["chlodnica"]:
            rad = s.get("radiator") or {}
            dB = rad.get("jedn", "dB") == "dB"
            self.podtytul(t["chl"].format(n=rad.get("name") or "—"))
            wiersze = [(t["chl_lw"].format(j="dB" if dB else "dB(A)"), [obliczenia.liczba(x, 0) for x in rad.get("lw", [0] * 8)], None)]
            if dB:
                wiersze.append((t["korekta_a"], obliczenia.KOREKTA_A, None))
            wiersze += [(t["lwa"], w["rad_lw_a"], w["rad_lw_a_suma"]), (t["chl_rdzen"], w["rad_rdzen"], None)]
            self.tabela_pasm(wiersze)
            if self.obrazy.get("radiator"):
                self.obraz(self.obrazy["radiator"])

    def wydech(self):
        t, w, n, s = self.t, self.w, self.n, self.stan
        if "wydech" in w["moduly"]:
            ak, md = s.get("acoustics") or {}, (s.get("mod_dims") or {}).get("wydech", {})
            dB = ak.get("jedn_wydechu", "dB") == "dB"
            geo = str(md.get("nad_dachem", "")).strip() != ""
            wys = 0 if geo else obliczenia.liczba(md.get("wysokosc"), 0)
            self.podtytul(t["wyd"].format(q=w["moduly"]["wydech"]["qty"], k=t["pionowo"] if md.get("kierunek") == obliczenia.PIONOWO else t["poziomo"],
                                          h=t["wysokosc"].format(h=n.g(wys)) if wys > 0 else ""))
            if geo:
                self.akapit(t["wyd_geo"].format(nd=n.g(obliczenia.liczba(md.get("nad_dachem"), 0)), os=n.g(obliczenia.liczba(md.get("od_sciany"), 0)),
                                                wp=n.g(obliczenia.liczba(md.get("wys_punktu"), obliczenia.WYS_PUNKTU))), 9, kolor=SZARY, po=4)
            j = "dB" if dB else "dB(A)"
            if w["wydech_lp1m"]:
                wiersze = [(t["wyd_lp"].format(j=j), w["wydech_wpisane"], None), (t["wyd_plus"], [obliczenia.LP_NA_LW] * 8, None),
                           (t["wyd_lwj"].format(j=j), w["wydech_lw"], None)]
            else:
                wiersze = [(t["wyd_lw"].format(j=j), w["wydech_lw"], None)]
            if dB:
                wiersze.append((t["korekta_a"], obliczenia.KOREKTA_A, None))
            wiersze += [(t["lwa"], w["wydech_lw_a"], w["wydech_lw_a_suma"]), (t["wyd_tl"], w["moduly"]["wydech"]["r"], None),
                        (t["wyd_po"], w["wydech_lw_a_po"], w["wydech_lw_a_po_suma"]),
                        (t["wyd_1m_wiersz"], w["wydech_lpa_1m"], w["wydech_lpa_1m_suma"])]
            self.tabela_pasm(wiersze)
            if w["wydech_lpa_1m_suma"]:
                self.ocena("info", t["wyd_1m"].format(v=n(w["wydech_lpa_1m_suma"], 1)))

    def wyniki(self):
        t, w, n = self.t, self.w, self.n
        for i, pt in enumerate(w["punkty"]):
            nazwa = pt["name"] or f"{n.g(pt['d'])} m"
            self.podtytul(t["pkt_tytul"].format(n=nazwa, d=n.g(pt["d"])))
            self.akapit(f"{t['s_pom']}: {n(pt['s_pom'])} m²  ·  {t['korekta_odl']}: {n(pt['korekta_odl'])} dB", 8.5, kolor=SZARY, po=4)
            wiersze = [[self.nazwa_modulu(k), n(w["moduly"][k]["s"]) if k != "wydech" else "—", n(m["lp_pow"], 1) if k != "wydech" else "—",
                        n(m["udzial"], 1), n(m["procent"], 1)] for k, m in sorted(pt["moduly"].items(), key=lambda x: -x[1]["udzial"])]
            wiersze.append([t["wynik_pkt"], n(w["s_calkowita"]), n(pt["lp_pow"], 1), n(pt["wynik"], 1), "100,0" if not self.en else "100.0"])
            self.tabela_prosta(t["udzialy"], wiersze, [2.6, 0.9, 1.4, 1.4, 0.9], pogrubiony_wiersz=len(wiersze) - 1)
            if i == 0 and len(pt["moduly"]) > 1:
                self.rysunek(rysuj_udzialy(pt, t, n, self), t["rys_udzialy"].replace("{p}", nazwa).replace("{w}", n(pt["wynik"], 1)), Emu(int(self.szer * 0.9)))
            self.tabela_pasm([(t["pasma_pkt"], pt["pasma"], pt["wynik"])])
            dom = self.nazwa_modulu(pt["dominujacy"]) if pt["dominujacy"] else "—"
            if pt.get("komin"):
                k = pt["komin"]
                self.akapit(t["komin"].format(h=n(k["wysokosc"], 1), r=n(k["r"], 1), wy=n(k["wydech"], 1), wk=n(k["wynik"], 1), ws=n(pt["wynik"], 1)), 9, przed=4, po=4)
                if k["wynik"] > pt["wynik"]:
                    dom = self.nazwa_modulu(k["dominujacy"])
            self.ocena(pt["zgodny"], (t["pkt_ok"] if pt["zgodny"] else t["pkt_zle"]).format(n=nazwa, w=n(pt.get("decyduje", pt["wynik"]), 1), l=n.g(pt["limit"]), z=n(abs(pt["zapas"]), 1), m=dom))

    def metoda(self):
        t = self.t
        for punkt in t["metoda"]:
            p = self.akapit(f"•  {punkt}", 9, po=3)
            p.paragraph_format.left_indent = Cm(0.4)
        self.akapit(t["zalozenia"], 9, przed=4, po=6)

    def zbuduj(self):
        t, w = self.t, self.w
        self.naglowek()
        self.kafelki()
        self.rewizje()
        self.sekcja(1, t["s1"])
        self.cel()
        self.sekcja(2, t["s2"])
        self.zrodlo()
        self.sekcja(3, t["s3"])
        self.wnetrze()
        self.sekcja(4, t["s4"], nowa_strona=True)
        self.budowa()
        nr = 5
        if w["chlodnica"]:
            self.sekcja(nr, t["s5_chl"])
            self.chlodnica()
            nr += 1
        if "wydech" in w["moduly"]:
            self.sekcja(nr, t["s5_wyd"])
            self.wydech()
            nr += 1
        self.sekcja(nr, t["s6"], nowa_strona=True)
        self.wyniki()
        self.sekcja(nr + 1, t["s7"])
        self.metoda()
        return self.doc


# ------------------------------------------------------------------ wykresy (obrazy PNG do raportu)

def _plotno(szer, wys):
    from PIL import Image, ImageDraw, ImageFont
    skala = 3
    obraz = Image.new("RGB", (szer * skala, wys * skala), "white")
    katalog = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")

    def czcionka(rozmiar, pogrubiona=False):
        try:
            return ImageFont.truetype(os.path.join(katalog, "segoeuib.ttf" if pogrubiona else "segoeui.ttf"), rozmiar * skala)
        except OSError:
            return ImageFont.load_default()
    return obraz, ImageDraw.Draw(obraz), czcionka, skala


def _zapisz_obraz(obraz):
    plik = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    plik.close()
    obraz.save(plik.name)
    return plik.name


def rysuj_widmo(w, t, n):
    """Widmo w pasmach: LwA silnika, LpA wewnątrz, poziom w pierwszym punkcie pomiarowym. Zwraca ścieżkę PNG."""
    obraz, rys, czcionka, sk = _plotno(900, 320)
    szer, wys = obraz.size
    l, p, g, d = 64 * sk, 20 * sk, 34 * sk, 46 * sk
    f_os, f_et = czcionka(12), czcionka(12, True)
    pt = w["punkty"][0]
    serie = [(t["l_lwa"], w["lw_a"], KOLOR_WYKRESU, False), (t["l_lpa"], w["lp_a"], "#6B7280", True),
             (t["l_pkt"].format(p=pt["name"] or f"{n.g(pt['d'])} m"), pt["pasma"], "#047857", False)]
    wartosci = [x for _, s, _, _ in serie for x in s if x]
    dol = int(min(wartosci) // 10 * 10) if wartosci else 0
    gora = int(-(-max(wartosci) // 10) * 10) if wartosci else 100
    gora = max(gora, dol + 10)
    x = lambda i: l + (i + 0.5) / 8 * (szer - l - p)
    y = lambda v: g + (1 - (v - dol) / (gora - dol)) * (wys - g - d)
    for v in range(dol, gora + 1, 10):
        rys.line([(l, y(v)), (szer - p, y(v))], fill="#E5E7EB", width=sk)
        rys.text((l - 8 * sk, y(v)), f"{v}", fill="#6B7280", font=f_os, anchor="rm")
    rys.text((l - 8 * sk, g - 18 * sk), "dB", fill="#6B7280", font=f_os, anchor="rm")
    for i, pasmo in enumerate(obliczenia.PASMA):
        rys.text((x(i), wys - d + 10 * sk), pasmo, fill="#6B7280", font=f_os, anchor="mt")
    rys.text(((l + szer - p) / 2, wys - 4 * sk), t["os_hz"], fill="#374151", font=f_os, anchor="mb")
    rys.line([(l, wys - d), (szer - p, wys - d)], fill="#6B7280", width=2 * sk)
    lx = l + 10 * sk
    for nazwa, s, kolor, przerywana in serie:
        punkty = [(x(i), y(v)) for i, v in enumerate(s)]
        if przerywana:
            for a, b in zip(punkty[:-1], punkty[1:]):
                for k in range(0, 10, 2):
                    rys.line([(a[0] + (b[0] - a[0]) * k / 10, a[1] + (b[1] - a[1]) * k / 10), (a[0] + (b[0] - a[0]) * (k + 1) / 10, a[1] + (b[1] - a[1]) * (k + 1) / 10)], fill=kolor, width=3 * sk)
        else:
            rys.line(punkty, fill=kolor, width=4 * sk, joint="curve")
        for px, py in punkty:
            rys.ellipse([px - 5 * sk, py - 5 * sk, px + 5 * sk, py + 5 * sk], fill=kolor, outline="white", width=2 * sk)
        rys.line([(lx, 14 * sk), (lx + 28 * sk, 14 * sk)], fill=kolor, width=4 * sk)
        rys.text((lx + 36 * sk, 14 * sk), nazwa, fill=kolor, font=f_et, anchor="lm")
        lx += 52 * sk + rys.textlength(nazwa, font=f_et)
    return _zapisz_obraz(obraz)


def rysuj_udzialy(pt, t, n, raport):
    """Poziome słupki: udział modułów w punkcie [dB(A)] i [%], od największego. Zwraca ścieżkę PNG."""
    moduly = sorted(pt["moduly"].items(), key=lambda x: -x[1]["udzial"])
    wiersz = 34
    obraz, rys, czcionka, sk = _plotno(820, 24 + wiersz * len(moduly))
    szer, _ = obraz.size
    f_et, f_b = czcionka(12), czcionka(12, True)
    l, p = 230 * sk, 130 * sk
    maks = max(m["udzial"] for _, m in moduly) or 1
    dol = max(0.0, min(m["udzial"] for _, m in moduly) - 15)
    for i, (k, m) in enumerate(moduly):
        yc = (12 + i * wiersz + wiersz / 2) * sk
        rys.text((l - 12 * sk, yc), raport.nazwa_modulu(k), fill="#111827", font=f_et, anchor="rm")
        dl = (m["udzial"] - dol) / ((maks - dol) or 1) * (szer - l - p)
        kolor = KOLOR_WYKRESU if i == 0 else "#7DB7D6"
        rys.rounded_rectangle([l, yc - 10 * sk, l + max(dl, 4 * sk), yc + 10 * sk], radius=4 * sk, fill=kolor)
        rys.text((l + max(dl, 4 * sk) + 10 * sk, yc), f"{n(m['udzial'], 1)} dB(A) · {n(m['procent'], 0)} %", fill="#111827" if i == 0 else "#374151",
                 font=f_b if i == 0 else f_et, anchor="lm")
    return _zapisz_obraz(obraz)


# ------------------------------------------------------------------ zapis

def generuj(sciezka, stan, w, obrazy, szablon=""):
    """
    stan: projekt (format pliku .json), w: wynik obliczenia.oblicz,
    obrazy: {'engine'|'radiator': ścieżka, 'moduly': {moduł: [ścieżki]}}, szablon: .docx z nagłówkiem / stopką.
    """
    r = Raport(stan, w, obrazy, szablon)
    try:
        r.zbuduj().save(sciezka)
    finally:
        for p in r.tymczasowe:
            try:
                os.remove(p)
            except OSError:
                pass
    return sciezka


def zapisz_pdf(docx, pdf):
    """PDF z dokumentu Word przez Microsoft Word (ta sama stopka i wygląd co w DOCX). RuntimeError, gdy Worda nie ma."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    word = None
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as e:
            raise RuntimeError("Zapis do PDF wymaga programu Microsoft Word na tym komputerze.") from e
        word.Visible = False
        word.DisplayAlerts = 0
        dokument = word.Documents.Open(os.path.abspath(docx), ReadOnly=True, AddToRecentFiles=False)
        try:
            dokument.ExportAsFixedFormat(os.path.abspath(pdf), 17)  # 17 = PDF
        finally:
            dokument.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
    return pdf


def dolacz_pdf(pdf, zalaczniki):
    """Karty producentów (PDF) dołączone na końcu raportu. Zwraca listę pominiętych (nieczytelnych) plików."""
    from pypdf import PdfReader, PdfWriter
    pisarz = PdfWriter()
    pisarz.append(PdfReader(pdf))
    pominiete = []
    for z in zalaczniki:
        try:
            pisarz.append(PdfReader(z))
        except Exception:
            pominiete.append(os.path.basename(z))
    tymczasowy = pdf + ".tmp"
    with open(tymczasowy, "wb") as f:
        pisarz.write(f)
    os.replace(tymczasowy, pdf)
    return pominiete
