"""
obliczenia.py

Obliczenia akustyczne obudowy agregatu - bez interfejsu (okno, raport i testy korzystają z tych samych funkcji).

Metoda (jak w wersji 1.1.0 - wynik całkowity w punkcie bez zmian):
- widmo mocy akustycznej silnika Lw w 8 pasmach oktawowych 63 Hz - 8 kHz; korekta A dodawana tylko do widm w dB
  (karty silników podają zwykle dB(A) - wtedy bez korekty),
- poziom ciśnienia wewnątrz obudowy: Lp = LwA + pochłanianie wnętrza (Room Abs., wzór Sabine'a),
- moduł obudowy (ściany, podłoga, sufit, drzwi, czerpnia, wyrzut): izolacyjność warstw sumowana energetycznie
  R = 10·lg Σ 10^(Ri/10) (zachowawczo - najlepsza warstwa + najwyżej kilka dB), moc wypromieniowana W = S·10^((Lp-R)/10),
- wydech (komin): własne widmo Lw wydechu minus tłumik, × liczba kominów; chłodnica na wyrzucie: własne Lw chłodnicy
  plus hałas silnika przez rdzeń chłodnicy,
- czerpnia / wyrzut / wydech skierowane pionowo w górę: −3 dB; wylot na wysokości h nad punktem: −20·lg(√(d²+h²)/d),
- wydech dodatkowo po stronie komina (od 2.0.2, gdy podane położenie wylotu): źródło punktowe Lp = LwA za tłumikiem −
  10·lg(4π·r²) + kierunek, r - odległość wylotu od punktu; pozostałe moduły jak w średniej. O zgodności decyduje gorszy wynik,
- widmo wydechu z karty tłumika (Lp@1m od wylotu): Lw = Lp + 10·lg(4π) ≈ Lp + 11 dB,
- poziom w punkcie pomiarowym wg ISO 3744 (prostopadłościenna powierzchnia pomiarowa l+2d × w+2d × h+d, bez podłogi):
  Lp(d) = 10·lg(ΣW / S_pom).

Poprawki względem 1.1.0 (wynik całkowity bez zmian):
- udział modułu w punkcie = 10·lg(W_modułu / S_pom) - moduły sumują się energetycznie do wyniku całkowitego
  (1.1.0 dzieliła przez powierzchnię modułu - małe moduły wychodziły tak głośne, jakby zajmowały całą obudowę),
- punkty pomiarowe są listą (1.1.0 trzymała je według odległości - dwa punkty w tej samej odległości się nadpisywały),
- jedno widmo silnika w oknie, obliczeniach i raporcie (z jednostką dB / dB(A)).
"""
import math

PASMA = ["63", "125", "250", "500", "1k", "2k", "4k", "8k"]
KOREKTA_A = [-26.2, -16.1, -8.6, -3.2, 0.0, 1.2, 1.0, -1.1]
MODULY = ["sciana", "podloga", "sufit", "drzwi", "czerpnia", "wyrzut", "wydech"]
AUTO_POWIERZCHNIA = {"sciana", "podloga", "sufit"}   # z wymiarów obudowy
OTWORY_W_SCIANIE = ("drzwi", "czerpnia", "wyrzut")   # zmniejszają powierzchnię ścian (jak w 1.1.0)
Z_KIERUNKIEM = ("wydech", "czerpnia", "wyrzut")       # kierunek wylotu i wysokość nad punktem
PIONOWO = "Pionowy w górę"
JEDNOSTKI = ("dBA", "dB")
LP_NA_LW = 10 * math.log10(4 * math.pi)  # ≈ 11 dB: Lw = Lp@1m + 10·lg(4π·1²) - źródło punktowe w wolnym polu (karty tłumików podają Lp@1m)
WYS_PUNKTU = 1.5  # pomiar standardowo 1,5 m nad terenem (odległość punktu - od obudowy, zwykle 1 m)


def wydech_geometria(md, h_obudowy, d):
    """
    Wydech jako źródło punktowe: (wysokość wylotu nad punktem, odległość wylotu od punktu po stronie komina) albo None,
    gdy projekt nie ma położenia wylotu (projekty sprzed 2.0.2 - tylko 'wysokosc' nad punktem).
    """
    nad_dachem = str(md.get("nad_dachem", "")).strip()
    if nad_dachem == "":
        return None
    v = max(0.0, h_obudowy + liczba(nad_dachem, 0.0) - WYS_PUNKTU)
    poziomo = max(0.0, liczba(md.get("od_sciany"), 0.0)) + d
    return v, max(0.5, math.sqrt(poziomo ** 2 + v ** 2))


def liczba(v, domyslnie=None):
    """Liczba z pola (przecinek albo kropka). Puste: domyslnie albo ValueError."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    t = str(v if v is not None else "").strip().replace(",", ".").replace(" ", "")
    if t == "":
        if domyslnie is None:
            raise ValueError("puste")
        return domyslnie
    return float(t)


def suma_db(poziomy):
    """Suma energetyczna poziomów [dB]; brak / same zera mocy -> 0."""
    w = sum(10 ** (x / 10) for x in poziomy)
    return 10 * math.log10(w) if w > 0 else 0.0


def widmo_a(widmo, jednostka):
    """Widmo w dB(A): korekta A dodawana tylko do widma w dB."""
    return [widmo[i] + (KOREKTA_A[i] if jednostka == "dB" else 0.0) for i in range(8)]


def izolacyjnosc(warstwy):
    """Izolacyjność modułu w pasmach: warstwy sumowane energetycznie; Rw/Ctr - ta sama wartość Rw+Ctr we wszystkich pasmach."""
    suma = [0.0] * 8
    for w in warstwy:
        if w.get("type") == "Rw / Ctr":
            r = [liczba(w.get("rw"), 0.0) + liczba(w.get("ctr"), 0.0)] * 8
        else:
            r = [liczba(x, 0.0) for x in (list(w.get("spec") or []) + [0] * 8)[:8]]
        for i in range(8):
            suma[i] += 10 ** (r[i] / 10)
    return [10 * math.log10(s) if s > 0 else 0.0 for s in suma]


def sabine(l, w, h, alfa_podlogi, alfa_scian, q, r):
    """Pochłanianie wnętrza (Room Abs.) [dB]: 10·lg(Q/(4πr²) + 4/R), R = S·α/(1-α) - stała pomieszczenia."""
    s_podl, s_sufit, s_scian = l * w, l * w, 2 * (l * h + w * h)
    s = s_podl + s_sufit + s_scian
    if s <= 0 or r <= 0:
        raise ValueError("wymiary obudowy i odległość r muszą być większe od zera")
    alfa = min(0.99, (alfa_podlogi * s_podl + alfa_scian * (s_scian + s_sufit)) / s)
    if alfa <= 0:
        raise ValueError("współczynniki pochłaniania muszą być większe od zera")
    stala = s * alfa / (1 - alfa)
    return {"alfa": alfa, "stala": stala, "s": s, "room_abs": 10 * math.log10(q / (4 * math.pi * r ** 2) + 4 / stala)}


def powierzchnie(p):
    """Powierzchnie modułów [m²]: ściany / podłoga / sufit z wymiarów obudowy, pozostałe - szerokość × wysokość × liczba."""
    wym = p.get("dimensions", {})
    l, w, h = (liczba(wym.get(k), 0.0) for k in ("kont_l", "kont_w", "kont_h"))
    aktywne, dims = p.get("modules_active", {}), p.get("mod_dims", {})
    s = {}
    for k in MODULY:
        if k not in AUTO_POWIERZCHNIA:
            m = dims.get(k, {})
            s[k] = round(liczba(m.get("w"), 0.0) * liczba(m.get("h"), 0.0) * liczba(m.get("qty"), 0.0), 2)
    s["wydech"] = 0.0  # komin: źródło o własnym widmie, bez powierzchni promieniującej
    s["podloga"] = s["sufit"] = round(l * w, 2)
    s["sciana"] = round(max(0.0, 2 * (l * h + w * h) - sum(s[k] for k in OTWORY_W_SCIANIE if aktywne.get(k))), 2)
    return s


def _sprawdz(p):
    bledy = {}
    wym = p.get("dimensions", {})
    for k, opis in (("kont_l", "Długość obudowy"), ("kont_w", "Szerokość obudowy"), ("kont_h", "Wysokość obudowy")):
        try:
            if liczba(wym.get(k)) <= 0:
                bledy[k] = f"{opis}: musi być większa od zera"
        except ValueError:
            bledy[k] = f"{opis}: wpisz liczbę"
    if not p.get("points"):
        bledy["points"] = "Dodaj co najmniej jeden punkt pomiarowy"
    for i, pt in enumerate(p.get("points", [])):
        try:
            if liczba(pt.get("d")) < 0:
                bledy[f"pt{i}"] = f"Punkt {pt.get('name') or i + 1}: odległość nie może być ujemna"
            liczba(pt.get("limit"))
        except ValueError:
            bledy[f"pt{i}"] = f"Punkt {pt.get('name') or i + 1}: wpisz odległość i wymagany poziom"
    if not any(p.get("modules_active", {}).values()):
        bledy["moduly"] = "Włącz co najmniej jeden moduł obudowy"
    return bledy


def oblicz(p):
    """Wyniki projektu (format pliku projektu .json). {'ok': False, 'bledy': {...}} przy brakujących / błędnych danych."""
    bledy = _sprawdz(p)
    if bledy:
        return {"ok": False, "bledy": bledy}
    ak, rad = p.get("acoustics", {}), p.get("radiator", {})
    lw = [liczba(x, 0.0) for x in ak.get("lw_vars", [0.0] * 8)]
    room_abs = [liczba(x, 0.0) for x in ak.get("room_abs_vars", [0.0] * 8)]
    jedn_silnika = ak.get("jedn_silnika", "dBA")
    lw_a = widmo_a(lw, jedn_silnika)
    lp_a = [lw_a[i] + room_abs[i] for i in range(8)]
    rad_aktywna = bool(rad.get("active"))
    rad_lw_a = widmo_a([liczba(x, 0.0) for x in rad.get("lw", [0.0] * 8)], rad.get("jedn", "dB"))
    rad_rdzen = [liczba(x, 0.0) for x in rad.get("att", [0.0] * 8)]
    lp_wydechu = ak.get("wielkosc_wydechu") == "Lp1m"  # widmo z karty tłumika: Lp@1m od wylotu -> Lw = Lp + 11 dB
    wydech_wpisane = [liczba(x, 0.0) for x in ak.get("exhaust_lw_vars", [0.0] * 8)]
    wydech_lw = [x + (LP_NA_LW if lp_wydechu else 0.0) for x in wydech_wpisane]
    wydech_lw_a = widmo_a(wydech_lw, ak.get("jedn_wydechu", "dB"))

    s_mod = powierzchnie(p)
    aktywne, dims, warstwy = p.get("modules_active", {}), p.get("mod_dims", {}), p.get("layers", {})
    moduly = {}
    for k in MODULY:
        if not aktywne.get(k):
            continue
        qty = int(liczba(dims.get(k, {}).get("qty"), 0.0))
        s = 0.0 if (k == "wydech" and qty > 0) else s_mod[k]
        if s <= 0 and k != "wydech":
            continue
        moduly[k] = {"s": s, "qty": qty, "r": izolacyjnosc(warstwy.get(k, [])), "warstwy": len(warstwy.get(k, []))}
    s_calkowita = sum(m["s"] for m in moduly.values())

    wym = p["dimensions"]
    l, w, h = (liczba(wym[k]) for k in ("kont_l", "kont_w", "kont_h"))
    s_kont = 2 * (l * h + w * h) + l * w
    punkty = []
    for pt in p["points"]:
        d, limit = liczba(pt.get("d")), liczba(pt.get("limit"))
        s_pom = 2 * ((l + 2 * d) * (h + d) + (w + 2 * d) * (h + d)) + (l + 2 * d) * (w + 2 * d)
        korekta_odl = 10 * math.log10(s_pom / s_kont)
        w_emisja_pasma = [0.0] * 8
        w_bez_wydechu = [0.0] * 8
        w_pow_suma = 0.0
        wyniki_mod = {}
        for k, m in moduly.items():
            s, r = m["s"], m["r"]
            korekta = 0.0  # kierunek i wysokość wylotu
            if k in Z_KIERUNKIEM:
                md = dims.get(k, {})
                if md.get("kierunek") == PIONOWO:
                    korekta -= 3.0
                geo = wydech_geometria(md, h, d) if k == "wydech" else None
                wys = geo[0] if geo else liczba(md.get("wysokosc"), 0.0)
                if wys > 0:
                    r_baza = max(d, 1.0)
                    korekta -= 20 * math.log10(math.sqrt(r_baza ** 2 + wys ** 2) / r_baza)
            if k == "wyrzut" and rad_aktywna:
                pasma = [s * 10 ** ((lp_a[i] - rad_rdzen[i] - r[i]) / 10) + 10 ** ((rad_lw_a[i] - r[i]) / 10) for i in range(8)]
            elif k == "wydech":
                pasma = [m["qty"] * 10 ** ((wydech_lw_a[i] - r[i]) / 10) for i in range(8)]
            else:
                pasma = [s * 10 ** ((lp_a[i] - r[i]) / 10) for i in range(8)]
            w_mod = sum(pasma)
            w_emisja = w_mod * 10 ** (korekta / 10)
            for i in range(8):
                w_emisja_pasma[i] += pasma[i] * 10 ** (korekta / 10)
                if k != "wydech":
                    w_bez_wydechu[i] += pasma[i] * 10 ** (korekta / 10)
            if k != "wydech":
                w_pow_suma += w_mod
            wyniki_mod[k] = {
                "lw": 10 * math.log10(w_mod) if w_mod > 0 else 0.0,                       # moc wypromieniowana przez moduł [dB(A)]
                "lp_pow": (10 * math.log10(w_mod / s) if s > 0 and w_mod > 0 else 0.0),  # poziom przy powierzchni modułu
                "udzial": 10 * math.log10(w_emisja / s_pom) if w_emisja > 0 else 0.0,     # udział w punkcie [dB(A)]
                "korekta": korekta,
            }
        w_emisja = sum(w_emisja_pasma)
        wynik = 10 * math.log10(w_emisja / s_pom) if w_emisja > 0 else 0.0
        procent = {k: (10 ** (v["udzial"] / 10) / 10 ** (wynik / 10) * 100 if w_emisja > 0 and v["udzial"] else 0.0) for k, v in wyniki_mod.items()}
        for k in wyniki_mod:
            wyniki_mod[k]["procent"] = procent[k]
        # Po stronie komina: wydech jako źródło punktowe w rzeczywistej odległości od wylotu (pozostałe moduły - jak wyżej, ISO 3744):
        # Lp = Lw za tłumikiem − 10·lg(4π·r²) + kierunek. Średnia po powierzchni pomiarowej zaniża wydech przy kominie.
        komin = None
        geo = wydech_geometria(dims.get("wydech", {}), h, d) if "wydech" in moduly else None
        if geo:
            m, md = moduly["wydech"], dims.get("wydech", {})
            kier = -3.0 if md.get("kierunek") == PIONOWO else 0.0
            pasma_wyd = [m["qty"] * 10 ** ((wydech_lw_a[i] - m["r"][i] - 10 * math.log10(4 * math.pi * geo[1] ** 2) + kier) / 10) for i in range(8)]
            pasma_k = [w_bez_wydechu[i] / s_pom + pasma_wyd[i] for i in range(8)]  # wydech w średniej zastąpiony źródłem punktowym
            suma_wyd = sum(pasma_wyd)
            komin = {"r": geo[1], "wysokosc": geo[0], "wydech": 10 * math.log10(suma_wyd) if suma_wyd > 0 else 0.0,
                     "wynik": 10 * math.log10(sum(pasma_k)) if sum(pasma_k) > 0 else 0.0,
                     "pasma": [10 * math.log10(x) if x > 0 else 0.0 for x in pasma_k]}
            inne = {kk: v["udzial"] for kk, v in wyniki_mod.items() if kk != "wydech"}
            komin["dominujacy"] = "wydech" if not inne or komin["wydech"] >= max(inne.values()) else max(inne, key=inne.get)
        decyduje = max(wynik, komin["wynik"]) if komin else wynik
        punkty.append({
            "name": pt.get("name", ""), "d": d, "limit": limit, "s_pom": s_pom, "korekta_odl": korekta_odl,
            "moduly": wyniki_mod, "wynik": wynik, "komin": komin, "decyduje": decyduje,
            "zapas": limit - decyduje, "zgodny": decyduje <= limit,
            "pasma": [10 * math.log10(x / s_pom) if x > 0 else 0.0 for x in w_emisja_pasma],
            "lp_pow": 10 * math.log10(w_pow_suma / s_calkowita) if s_calkowita > 0 and w_pow_suma > 0 else 0.0,
            "dominujacy": max(wyniki_mod, key=lambda k: wyniki_mod[k]["udzial"]) if wyniki_mod else "",
        })

    return {
        "ok": True, "pasma": PASMA, "korekta_a": KOREKTA_A,
        "lw_silnika": lw, "jedn_silnika": jedn_silnika, "lw_a": lw_a, "lw_a_suma": suma_db(lw_a),
        "room_abs": room_abs, "lp_a": lp_a, "lp_a_suma": suma_db(lp_a),
        "chlodnica": rad_aktywna, "rad_lw_a": rad_lw_a, "rad_lw_a_suma": suma_db(rad_lw_a), "rad_rdzen": rad_rdzen,
        "wydech_lw_a": wydech_lw_a, "wydech_lw_a_suma": suma_db(wydech_lw_a) if "wydech" in moduly else 0.0,
        "wydech_wpisane": wydech_wpisane, "wydech_lw": wydech_lw, "wydech_lp1m": lp_wydechu,
        "wydech_pusty": "wydech" in moduly and not any(wydech_wpisane),
        # hałas wydechu 1 m od wylotu, za tłumikiem (jak "Reduced Sound Pressure Level: Lp at 1 m" w kartach tłumików), 1 komin
        "wydech_lpa_1m": [wydech_lw_a[i] - moduly["wydech"]["r"][i] - LP_NA_LW for i in range(8)] if "wydech" in moduly else [0.0] * 8,
        "wydech_lpa_1m_suma": suma_db([wydech_lw_a[i] - moduly["wydech"]["r"][i] for i in range(8)]) - LP_NA_LW if "wydech" in moduly and any(wydech_wpisane) else 0.0,
        "wydech_lw_a_po": [wydech_lw_a[i] - moduly["wydech"]["r"][i] for i in range(8)] if "wydech" in moduly else [0.0] * 8,
        "wydech_lw_a_po_suma": suma_db([wydech_lw_a[i] - moduly["wydech"]["r"][i] for i in range(8)]) if "wydech" in moduly and any(wydech_wpisane) else 0.0,
        "powierzchnie": s_mod, "moduly": moduly, "s_calkowita": s_calkowita, "s_kont": s_kont,
        "punkty": punkty, "zgodny": all(pt["zgodny"] for pt in punkty),
    }
