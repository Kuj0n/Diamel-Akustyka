"""
Testy obliczeń Kalkulatora Akustyki (python testy/test_obliczenia.py).

- dane syntetyczne: wartości policzone niezależnie (wzory z obliczenia.py zapisane wprost),
- testy/dane_prywatne (poza repozytorium - projekty klientów): wyniki wersji 1.1.0 zapisane jako wzorzec
  (wzorzec_110.json + projekty .json) - wynik całkowity w każdym punkcie musi być taki sam jak w 1.1.0.
"""
import copy
import glob
import json
import math
import os
import sys

TU = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(TU))
import obliczenia  # noqa: E402

PRYWATNE = os.path.join(TU, "dane_prywatne")
db = lambda poziomy: 10 * math.log10(sum(10 ** (x / 10) for x in poziomy))


def projekt():
    """Kontener 12 × 3 × 3 m, ściany i dach, ta sama izolacja 30 dB, Lp wewnątrz 100 dB(A) w każdym paśmie."""
    return {
        "dimensions": {"kont_l": 12, "kont_w": 3, "kont_h": 3},
        "points": [{"name": "1 m", "d": 1, "limit": 85}],
        "modules_active": {"sciana": True, "sufit": True, "drzwi": True},
        "mod_dims": {"drzwi": {"qty": 1, "w": 1, "h": 2}},
        "layers": {k: [{"type": "Widmo pasmowe", "spec": [30] * 8}] for k in ("sciana", "sufit", "drzwi")},
        "acoustics": {"lw_vars": [103] * 8, "room_abs_vars": [-3] * 8, "jedn_silnika": "dBA"},
    }


def test_wynik_i_udzialy():
    p = projekt()
    w = obliczenia.oblicz(p)
    assert w["ok"], w
    s_sciana, s_sufit, s_drzwi = 2 * (12 * 3 + 3 * 3) - 2, 36, 2
    assert w["powierzchnie"]["sciana"] == s_sciana and w["powierzchnie"]["sufit"] == s_sufit and w["powierzchnie"]["drzwi"] == s_drzwi
    s_pom = 2 * (14 * 4 + 5 * 4) + 14 * 5
    lp_pasmo = 100 - 30
    w_suma = (s_sciana + s_sufit + s_drzwi) * 8 * 10 ** (lp_pasmo / 10)
    pt = w["punkty"][0]
    assert abs(pt["wynik"] - 10 * math.log10(w_suma / s_pom)) < 1e-9
    # udział modułu względem powierzchni pomiarowej - moduły sumują się do wyniku (1.1.0: każdy moduł = wynik całkowity)
    assert abs(pt["moduly"]["drzwi"]["udzial"] - 10 * math.log10(s_drzwi * 8 * 10 ** (lp_pasmo / 10) / s_pom)) < 1e-9
    assert abs(db([m["udzial"] for m in pt["moduly"].values()]) - pt["wynik"]) < 1e-9
    assert abs(sum(m["procent"] for m in pt["moduly"].values()) - 100) < 1e-9
    assert pt["dominujacy"] == "sciana" and pt["zgodny"] == (pt["wynik"] <= 85)
    assert abs(pt["korekta_odl"] - 10 * math.log10(s_pom / (2 * (12 * 3 + 3 * 3) + 12 * 3))) < 1e-9


def test_punkty_w_tej_samej_odleglosci():
    p = projekt()
    p["points"] = [{"name": "Front 1 m", "d": 1, "limit": 85}, {"name": "Bok 1 m", "d": 1, "limit": 60}, {"name": "7 m", "d": "7", "limit": "75"}]
    w = obliczenia.oblicz(p)
    assert [pt["name"] for pt in w["punkty"]] == ["Front 1 m", "Bok 1 m", "7 m"]
    assert w["punkty"][0]["wynik"] == w["punkty"][1]["wynik"] and w["punkty"][0]["zgodny"] != w["punkty"][1]["zgodny"]
    assert w["punkty"][2]["wynik"] < w["punkty"][0]["wynik"] and not w["zgodny"]


def test_jednostki_widm():
    p = projekt()
    w_dba = obliczenia.oblicz(p)
    p["acoustics"]["jedn_silnika"] = "dB"
    w_db = obliczenia.oblicz(p)
    assert [round(x, 9) for x in w_db["lw_a"]] == [round(103 + k, 9) for k in obliczenia.KOREKTA_A]
    assert w_db["punkty"][0]["wynik"] < w_dba["punkty"][0]["wynik"]
    # wydech: widmo w dB (jak liczyła 1.1.0) dostaje korektę A, w dB(A) - nie
    p = projekt()
    p["modules_active"]["wydech"] = True
    p["mod_dims"]["wydech"] = {"qty": 2}
    p["acoustics"]["exhaust_lw_vars"] = [110] * 8
    w = obliczenia.oblicz(p)
    assert w["moduly"]["wydech"]["s"] == 0 and abs(w["wydech_lw_a_suma"] - db([110 + k for k in obliczenia.KOREKTA_A])) < 1e-9
    s_pom = w["punkty"][0]["s_pom"]
    assert abs(w["punkty"][0]["moduly"]["wydech"]["udzial"] - 10 * math.log10(2 * sum(10 ** ((110 + k) / 10) for k in obliczenia.KOREKTA_A) / s_pom)) < 1e-9
    p["acoustics"]["jedn_wydechu"] = "dBA"
    assert abs(obliczenia.oblicz(p)["wydech_lw_a_suma"] - db([110] * 8)) < 1e-9


def test_izolacyjnosc_i_kierunek():
    assert [round(x, 6) for x in obliczenia.izolacyjnosc([{"type": "Widmo pasmowe", "spec": [30] * 8}] * 2)] == [round(30 + 10 * math.log10(2), 6)] * 8
    assert obliczenia.izolacyjnosc([{"type": "Rw / Ctr", "rw": 32, "ctr": -4}]) == [28.0] * 8
    assert obliczenia.izolacyjnosc([]) == [0.0] * 8
    p = projekt()
    p["modules_active"]["czerpnia"] = True
    p["mod_dims"]["czerpnia"] = {"qty": 1, "w": 1, "h": 1, "kierunek": "Pionowy w górę", "wysokosc": 3}
    w = obliczenia.oblicz(p)
    assert abs(w["punkty"][0]["moduly"]["czerpnia"]["korekta"] - (-3 - 20 * math.log10(math.sqrt(1 + 9) / 1))) < 1e-9
    assert w["powierzchnie"]["sciana"] == 2 * (12 * 3 + 3 * 3) - 2 - 1  # czerpnia zmniejsza ściany


def test_sabine_i_bledy():
    r = obliczenia.sabine(12, 3, 3, 0.02, 0.85, 2, 1.5)
    s = 2 * 36 + 2 * 45
    alfa = (0.02 * 36 + 0.85 * (90 + 36)) / s
    assert abs(r["room_abs"] - 10 * math.log10(2 / (4 * math.pi * 1.5 ** 2) + 4 / (s * alfa / (1 - alfa)))) < 1e-9
    p = projekt()
    p["dimensions"]["kont_w"] = ""
    p["points"] = []
    w = obliczenia.oblicz(p)
    assert not w["ok"] and "kont_w" in w["bledy"] and "points" in w["bledy"]


def test_wzorzec_110():
    """Projekty klientów (poza repozytorium): wynik całkowity jak w wersji 1.1.0."""
    plik = os.path.join(PRYWATNE, "wzorzec_110.json")
    if not os.path.isfile(plik):
        return 0
    wzorzec = json.load(open(plik, encoding="utf-8"))
    projekty = {os.path.basename(f): f for f in glob.glob(os.path.join(PRYWATNE, "**", "*.json"), recursive=True)}
    n = 0
    for nazwa, stary in wzorzec.items():
        w = obliczenia.oblicz(json.load(open(projekty[nazwa], encoding="utf-8")))
        for pt in w["punkty"]:
            st = stary["wyniki_dystans"].get(str(pt["d"]))
            if st and st["name"] == pt["name"]:  # 1.1.0: punkty w tej samej odległości nadpisane - zostaje ostatni
                assert abs(pt["wynik"] - st["wynik_dystans_wypadkowy"]) < 1e-9, (nazwa, pt["name"])
                assert abs(pt["lp_pow"] - st["halas_wypadkowy_pow"]) < 1e-9
                n += 1
        assert max(abs(a - b) for a, b in zip(w["lp_a"], stary["lp_a_wazone"])) < 1e-9
    return n


if __name__ == "__main__":
    test_wynik_i_udzialy()
    test_punkty_w_tej_samej_odleglosci()
    test_jednostki_widm()
    test_izolacyjnosc_i_kierunek()
    test_sabine_i_bledy()
    n = test_wzorzec_110()
    print(f"OK - obliczenia, udziały modułów, punkty, jednostki widm, izolacyjność, Sabine"
          + (f"; {n} punktów z projektów klientów zgodnych z wersją 1.1.0" if n else "; brak danych prywatnych (wzorzec 1.1.0 pominięty)"))
