import os
import json
import urllib.request
import threading
import subprocess
import tempfile
import urllib.error
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

from konfiguracja_app import VERSION, NAZWA_PROGRAMU, GITHUB_REPO, URL_BAZY, URL_ZGLOSZEN

NAGLOWKI = {'User-Agent': f'Diamel-{NAZWA_PROGRAMU}/{VERSION}'}

def wez_krotke(v): return tuple(map(int, (v.lstrip("vV").split(" ")[0]).split(".")))

def pobierz_json(url, timeout=4.0):
    req = urllib.request.Request(url, headers=NAGLOWKI)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode('utf-8'))

def pobierz_wydanie():
    """Najnowsze wydanie z GitHub Releases. Zwraca (dane wydania, {nazwa załącznika: link})."""
    dane = pobierz_json(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest")
    zalaczniki = {a["name"]: a["browser_download_url"] for a in dane.get("assets", [])}
    return dane, zalaczniki

# ----------------- AKTUALIZACJE PROGRAMU -----------------
def sprawdz_aktualizacje(root, tryb="recznie"):
    def task():
        try:
            dane, zalaczniki = pobierz_wydanie()
            wersja_serwera = dane.get("tag_name", "0.0.0").lstrip("vV")
            changelog = dane.get("body") or "Brak informacji o nowościach."
            link = next((url for nazwa, url in zalaczniki.items() if nazwa.lower().endswith(".exe")), "")

            if wez_krotke(wersja_serwera) > wez_krotke(VERSION) and link:
                root.after(0, lambda: pokaz_monit_aktualizacji(root, wersja_serwera, link, changelog))
            elif tryb == "recznie":
                msg = f"Posiadasz najnowszą dostępną wersję programu!\n\nTwoja wersja: {VERSION}\nWersja na serwerze: {wersja_serwera}"
                root.after(0, lambda: messagebox.showinfo("Aktualizacje Programu", msg, parent=root))
        except Exception as e:
            if tryb == "recznie":
                root.after(0, lambda: messagebox.showerror("Błąd połączenia", f"Nie udało się połączyć z serwerem.\nSzczegóły: {e}", parent=root))
    threading.Thread(target=task, daemon=True).start()

def pokaz_monit_aktualizacji(root, nowa_wersja, link, changelog):
    odp = messagebox.askyesno("Nowa wersja programu!", f"Dostępna jest nowa wersja programu: {nowa_wersja} (Twoja: {VERSION})\n\nCO NOWEGO:\n{changelog}\n\nCzy chcesz pobrać i zainstalować aktualizację?", parent=root)
    if odp: pobierz_i_zainstaluj(root, link)

def pobierz_i_zainstaluj(root, link):
    okno_pob = tk.Toplevel(root)
    okno_pob.title("Aktualizacja Akustyki")
    okno_pob.geometry("400x180")
    okno_pob.grab_set()

    lbl = tk.Label(okno_pob, text="Trwa pobieranie aktualizacji...", font=("Arial", 10, "bold"))
    lbl.pack(pady=15)
    
    pasek = ttk.Progressbar(okno_pob, mode='indeterminate', length=300)
    pasek.pack(pady=5)
    pasek.start()

    def start_update():
        sciezka_temp = os.path.join(tempfile.gettempdir(), "Aktualizacja_Akustyka_Instalator.exe")
        try:
            subprocess.run(['powershell', '-Command', f'Unblock-File -Path \'{sciezka_temp}\''], shell=True)
            subprocess.Popen([sciezka_temp, "/FORCECLOSEAPPLICATIONS", "/NOCANCEL"])
            os._exit(0)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się uruchomić: {e}", parent=okno_pob)

    btn_start = tk.Button(okno_pob, text="URUCHOM INSTALATOR", command=start_update, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), height=2)

    def download_thread():
        try:
            sciezka_temp = os.path.join(tempfile.gettempdir(), "Aktualizacja_Akustyka_Instalator.exe")
            urllib.request.urlretrieve(link, sciezka_temp)
            def pokaz_przycisk():
                pasek.stop()
                pasek.pack_forget()
                lbl.config(text="Aktualizacja pobrana pomyślnie!")
                btn_start.pack(pady=15, fill="x", padx=40)
            okno_pob.after(0, pokaz_przycisk)
        except Exception as e:
            okno_pob.after(0, lambda: messagebox.showerror("Błąd", f"Nie udało się pobrać: {e}", parent=okno_pob))
            okno_pob.after(0, okno_pob.destroy)

    threading.Thread(target=download_thread, daemon=True).start()


# ----------------- AKTUALIZACJE BAZY DANYCH -----------------
def sprawdz_aktualizacje_bazy(root, app_instance, tryb="recznie"):
    def task():
        try:
            # Baza to plik acoustic_db.json w repozytorium - wersja i historia zmian są w samym pliku
            dane = pobierz_json(URL_BAZY)
            wersja_bazy_serwera = dane.get("version", "0.0.0")
            changelog_bazy = dane.get("changelog") or "Brak informacji."

            if wez_krotke(wersja_bazy_serwera) > wez_krotke(app_instance.db_version):
                root.after(0, lambda: pokaz_monit_bazy(root, app_instance, wersja_bazy_serwera, URL_BAZY, changelog_bazy))
            elif tryb == "recznie":
                msg = f"Posiadasz najnowszą bazę danych!\n\nTwoja wersja: {app_instance.db_version}\nWersja na serwerze: {wersja_bazy_serwera}"
                root.after(0, lambda: messagebox.showinfo("Aktualizacje Bazy", msg, parent=root))
        except Exception as e:
            if tryb == "recznie":
                root.after(0, lambda: messagebox.showerror("Błąd połączenia", f"Nie udało się połączyć z serwerem.\nSzczegóły: {e}", parent=root))
    threading.Thread(target=task, daemon=True).start()

def pokaz_monit_bazy(root, app_instance, nowa_wersja, link, changelog):
    najnowsze_zmiany = changelog.strip().split("\n\n")[0]
    odp = messagebox.askyesno("Nowa wersja Bazy Danych!", f"Wydano nową bazę danych: {nowa_wersja} (Twoja: {app_instance.db_version})\n\nZMIANY W BAZIE:\n{najnowsze_zmiany}\n\nCzy chcesz pobrać i scalić nową bazę z Twoimi danymi?", parent=root)
    if odp: 
        chce_kopie = messagebox.askyesno("Kopia zapasowa", "Zanim zaktualizuję dane, czy chcesz utworzyć kopię zapasową swojej obecnej bazy danych?", parent=root)
        # Przekazujemy changelog aby zapisał się w historii aplikacji
        pobierz_i_scal_baze(root, app_instance, link, chce_kopie, nowa_wersja, changelog)

def pobierz_i_scal_baze(root, app_instance, link, zrob_kopie, nowa_wersja, nowy_changelog):
    def task():
        try:
            req = urllib.request.Request(link, headers=NAGLOWKI)
            with urllib.request.urlopen(req, timeout=10.0) as response:
                nowa_baza = json.loads(response.read().decode('utf-8'))
                
                db_path = app_instance.db_file
                bkp_name = ""
                
                if zrob_kopie and os.path.exists(db_path):
                    bkp_name = f"acoustic_db_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                    bkp_path = os.path.join(os.path.dirname(os.path.abspath(db_path)), bkp_name)
                    shutil.copy(db_path, bkp_path)
                
                if "materials" in nowa_baza: app_instance.materials_db.update(nowa_baza["materials"])
                if "engines" in nowa_baza: app_instance.engines_db.update(nowa_baza["engines"])
                if "radiators" in nowa_baza: app_instance.radiators_db.update(nowa_baza["radiators"])
                    
                app_instance.db_version = nowa_wersja
                
                # Historia zmian bazy - pełna lista z pliku w repozytorium
                if nowy_changelog:
                    app_instance.db_changelog = nowy_changelog
                
                app_instance.save_db()
                
                root.after(0, app_instance.refresh_ui_dropdowns)
                
                komunikat = "Baza danych została pomyślnie zaktualizowana i scalona!"
                if zrob_kopie:
                    komunikat += f"\n\nUtworzono plik z kopią zapasową:\n{bkp_name}"
                    
                root.after(0, lambda: messagebox.showinfo("Sukces", komunikat, parent=root))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("Błąd", f"Aktualizacja bazy nie powiodła się:\n{e}", parent=root))
    threading.Thread(target=task, daemon=True).start()


# ----------------- PROŚBA O NOWY MATERIAŁ (Z PARAMETRAMI) -----------------
def pokaz_okno_zgloszenia_materialu(root):
    pop = tk.Toplevel(root)
    pop.title("Zgłoś prośbę o dodanie materiału / maszyny do Bazy Danych")
    pop.geometry("550x700")
    pop.grab_set()

    canvas = tk.Canvas(pop, highlightthickness=0)
    scroll = ttk.Scrollbar(pop, orient="vertical", command=canvas.yview)
    frame = ttk.Frame(canvas)
    
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scroll.set)
    canvas.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")

    tk.Label(frame, text="Twój e-mail (DO KONTAKTU):", font=("Arial", 9, "bold")).pack(pady=(15, 2), anchor="w", padx=20)
    v_email = tk.StringVar()
    tk.Entry(frame, textvariable=v_email, width=45).pack(padx=20, anchor="w")

    tk.Label(frame, text="Kategoria obiektu:", font=("Arial", 9, "bold")).pack(pady=(10, 2), anchor="w", padx=20)
    v_typ = tk.StringVar(value="Materiał Tłumiący")
    ttk.Combobox(frame, textvariable=v_typ, values=["Materiał Tłumiący", "Silnik/Agregat", "Chłodnica", "Tłumik Kominowy"], state="readonly", width=42).pack(padx=20, anchor="w")

    tk.Label(frame, text="Nazwa / Producent / Model:", font=("Arial", 9, "bold")).pack(pady=(10, 2), anchor="w", padx=20)
    v_nazwa = tk.StringVar()
    tk.Entry(frame, textvariable=v_nazwa, width=45).pack(padx=20, anchor="w")
    
    tk.Label(frame, text="Grubość / Długość [mm]:", font=("Arial", 9, "bold")).pack(pady=(10, 2), anchor="w", padx=20)
    v_d = tk.DoubleVar(value=0.0)
    tk.Entry(frame, textvariable=v_d, width=15).pack(padx=20, anchor="w")

    f_param = ttk.LabelFrame(frame, text=" Parametry akustyczne (uzupełnij jeśli znasz) ")
    f_param.pack(fill="x", padx=20, pady=15)

    tk.Label(f_param, text="Typ podawanych danych:").pack(anchor="w", padx=10, pady=(10, 2))
    v_format = tk.StringVar(value="Widmo pasmowe")
    cb_format = ttk.Combobox(f_param, textvariable=v_format, values=["Rw / Ctr", "Widmo pasmowe"], state="readonly", width=25)
    cb_format.pack(anchor="w", padx=10, pady=2)

    f_rw = ttk.Frame(f_param)
    v_rw = tk.DoubleVar(value=0.0)
    v_ctr = tk.DoubleVar(value=0.0)
    tk.Label(f_rw, text="Rw [dB]:").pack(side="left", padx=2)
    tk.Entry(f_rw, textvariable=v_rw, width=8).pack(side="left", padx=2)
    tk.Label(f_rw, text="Ctr [dB]:").pack(side="left", padx=15)
    tk.Entry(f_rw, textvariable=v_ctr, width=8).pack(side="left", padx=2)

    f_spec = ttk.Frame(f_param)
    freqs = ["63", "125", "250", "500", "1k", "2k", "4k", "8k"]
    v_spec = [tk.DoubleVar(value=0.0) for _ in range(8)]
    for i, f in enumerate(freqs):
        tk.Label(f_spec, text=f, font=("", 8)).grid(row=0, column=i, padx=5, pady=2)
        tk.Entry(f_spec, textvariable=v_spec[i], width=5).grid(row=1, column=i, padx=5, pady=(0,10))

    def update_format(*args):
        if v_format.get() == "Rw / Ctr":
            f_spec.pack_forget()
            f_rw.pack(anchor="w", padx=10, pady=10)
        else:
            f_rw.pack_forget()
            f_spec.pack(anchor="w", padx=10, pady=10)
            
    cb_format.bind("<<ComboboxSelected>>", update_format)
    update_format()

    tk.Label(frame, text="Dodatkowy opis (link do karty produktu, uwagi):", font=("Arial", 9, "bold")).pack(pady=(10, 2), anchor="w", padx=20)
    txt_opis = tk.Text(frame, width=50, height=6, font=("Arial", 9))
    txt_opis.pack(padx=20, anchor="w")

    def wyslij():
        e, n, t, o = v_email.get().strip(), v_nazwa.get().strip(), v_typ.get(), txt_opis.get("1.0", tk.END).strip()
        if not e or "@" not in e: return messagebox.showerror("Błąd", "Podaj poprawny e-mail!", parent=pop)
        if not n: return messagebox.showerror("Błąd", "Podaj nazwę!", parent=pop)

        param_str = f"Grubość/Długość: {v_d.get()} mm\nTyp danych: {v_format.get()}\n"
        if v_format.get() == "Rw / Ctr":
            param_str += f"Rw: {v_rw.get()} dB\nCtr: {v_ctr.get()} dB\n"
        else:
            pasmowe = [str(x.get()) for x in v_spec]
            param_str += f"Widmo (63Hz - 8kHz): {', '.join(pasmowe)} dB\n"

        btn_wyslij.config(state="disabled", text="Wysyłanie zgłoszenia. Proszę czekać...")
        
        def task():
            try:
                opis = f"{o}\n\n--- PARAMETRY TECHNICZNE ---\n{param_str}"
                zgloszenie = {"program": NAZWA_PROGRAMU, "wersja": VERSION, "email": e, "kategoria": t, "nazwa": n, "opis": opis}
                req = urllib.request.Request(URL_ZGLOSZEN, data=json.dumps(zgloszenie).encode('utf-8'),
                                             headers={**NAGLOWKI, 'Content-Type': 'application/json'})
                try:
                    with urllib.request.urlopen(req, timeout=15.0) as response:
                        wynik = json.loads(response.read().decode('utf-8'))
                except urllib.error.HTTPError as http_err:
                    # Serwer odrzucił zgłoszenie - powód jest w treści odpowiedzi
                    try: wynik = json.loads(http_err.read().decode('utf-8'))
                    except Exception: raise http_err
                if not wynik.get("ok"):
                    raise Exception(wynik.get("blad", "Nieznany błąd serwera."))

                pop.after(0, lambda: messagebox.showinfo("Wysłano", "Prośba o dodanie materiału z wprowadzonymi parametrami została przekazana. Odpowiemy na podany adres e-mail.\n\nDziękujemy!", parent=pop))
                pop.after(0, pop.destroy)
            except Exception as exc:
                pop.after(0, lambda: messagebox.showerror("Błąd", f"Nie udało się wysłać:\n{exc}", parent=pop))
                pop.after(0, lambda: btn_wyslij.config(state="normal", text="Wyślij do weryfikacji przez inżyniera"))
                
        threading.Thread(target=task, daemon=True).start()

    btn_wyslij = tk.Button(frame, text="Wyślij do weryfikacji przez inżyniera", command=wyslij, bg="#2196F3", fg="white", font=("Arial", 10, "bold"), height=2)
    btn_wyslij.pack(pady=25, fill="x", padx=20)