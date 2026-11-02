import sys
import os

# --- FIX DLA BRAKU KONSOLI (--noconsole) PRZY KONWERSJI PDF ---
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")
# --------------------------------------------------------------

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import math
from datetime import datetime
import io
import re

# Nasze moduły MVC
from konfiguracja_app import *
from narzedzia import sprawdz_aktualizacje, sprawdz_aktualizacje_bazy, pokaz_okno_zgloszenia_materialu

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    pass

try:
    from docx2pdf import convert as docx_to_pdf
    from PyPDF2 import PdfMerger
    PDF_MERGE_AVAILABLE = True
except ImportError:
    PDF_MERGE_AVAILABLE = False

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


class AcousticReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Acoustic Report Generator - Projektowanie Warstwowe v{VERSION}")
        self.root.geometry("1200x850")
        
        self.freqs = ["63", "125", "250", "500", "1k", "2k", "4k", "8k"]
        self.k_dba_default = [-26.2, -16.1, -8.6, -3.2, 0.0, 1.2, 1.0, -1.1]
        
        self.alfa_podl = tk.DoubleVar(value=0.02)
        self.alfa_scian = tk.DoubleVar(value=0.85)
        self.sabine_q = tk.DoubleVar(value=2.0)
        self.sabine_r = tk.DoubleVar(value=1.5)

        self.rad_active = tk.BooleanVar(value=False)
        self.rad_name = tk.StringVar(value="")
        self.rad_lw_vars = [tk.DoubleVar(value=0.0) for _ in range(8)]
        self.rad_att_vars = [tk.DoubleVar(value=0.0) for _ in range(8)]
        self.rad_wiz_path = tk.StringVar(value="")
        self.rad_pdf_path = tk.StringVar(value="")

        self.engine_wiz_path = tk.StringVar(value="")
        self.engine_pdf_path = tk.StringVar(value="")
        self.exhaust_lw_vars = [tk.DoubleVar(value=0.0) for _ in range(8)]
        
        self.points_data = []

        self.translations = {
            'Polski': {
                'title': "RAPORT AKUSTYCZNY\nANALIZA ZŁOŻONEJ OBUDOWY",
                'meta': ["Nr Projektu:", "Tytuł Projektu:", "Dla Klienta:", "Numer P.O.:", "Dokument Klienta (Nr):", "Data:"],
                'inst_type_label': "Typ instalacji:",
                'types': {'Agregat otwarty': "Agregat otwarty (bez obudowy)", 'Obudowa dźwiękochłonna': "Obudowa dźwiękochłonna (Canopy)", 'Obudowa kontenerowa': "Obudowa kontenerowa"},
                'rev_th': ["Rew.", "Opis zmiany", "Wykonał", "Sprawdził", "Data"],
                'sec1': "1. CEL ANALIZY",
                'sec1_desc': "Celem niniejszej analizy jest określenie danych obliczeniowych dla projektowanej obudowy, aby spełnić wymagania akustyczne w wyznaczonych punktach pomiarowych. Jako pierwotne źródło przyjęto poziom hałasu mechanicznego generowanego przez silnik.",
                'sec2': "2. DANE ŹRÓDŁOWE",
                'sec2_desc': "Wewnątrz obudowy umieszczony zostanie agregat prądotwórczy (Model: {genset}) z silnikiem {engine}. Widmo źródła odniesienia:",
                'th1_1': "Pasmo oktawowe", 'th1_2': "Poprawka A [dB]", 'th1_3': "LwA [dB(A)]",
                'calc_lwa': "Całkowita moc akustyczna silnika (LwA): ",
                'sec3': "3. CHŁONNOŚĆ AKUSTYCZNA I CIŚNIENIE WEWNĘTRZNE",
                'th2_2': "Moc LwA [dB(A)]", 'th2_3': "Absorpcja [dB]", 'th2_4': "Ciśnienie LpA [dB(A)]",
                'calc_lpa': "Całkowite ciśnienie akustyczne we wnętrzu (LpA): ",
                'sec4': "4. SPECYFIKACJA WARSTWOWA I AKTYWNE ŚCIEŻKI",
                'sec4_desc': "Do obliczeń uwzględniono następujące moduły wraz z ich budową warstwową (tłumienie widmowe sumowane pasmowo):",
                'mod_names': {'sciana': "Ściany", 'podloga': "Podłoga / Baza", 'sufit': "Sufit", 'drzwi': "Drzwi Akustyczne", 'czerpnia': "Czerpnia (Wlot)", 'wyrzut': "Wyrzut (Wylot)", 'wydech': "Wydech silnika (Komin)"},
                'depth': "Grub./Dług.", 'spec_att': "Tłumienie Widmowe", 'tot_area': "Całkowita czynna powierzchnia obudowy", 'qty_pcs': "szt.",
                'sec5': "5. WNIOSKI I ZGODNOŚĆ Z WYMAGANIAMI", 'dist_corr': "Korekcja na odległość ISO 3744", 'ind_noise': "Indywidualna transmisja hałasu z modułów (na podanym dystansie):",
                'tot_noise': "CAŁKOWITY WYPADKOWY POZIOM HAŁASU", 'proj_stat': "Status punktu: ", 'pass': "ZGODNY (SPEŁNIA WYMAGANIA).", 'fail': "NIEZGODNY (NIE SPEŁNIA WYMAGAŃ).",
                'err_in_use_title': "Błąd zapisu / Plik w użyciu", 'err_in_use_msg': "Plik:\n{}\njest obecnie otwarty.\n\n• 'Tak' = ponów próbę\n• 'Nie' = zapisz pod nową nazwą\n• 'Anuluj' = przerwij."
            },
            'English': {
                'title': "ACOUSTIC REPORT\nCOMPOSITE ENCLOSURE ANALYSIS",
                'meta': ["Project no:", "Project Title:", "Intended for:", "P.O. Number:", "Client doc.no:", "Date:"],
                'inst_type_label': "Installation Type:",
                'types': {
                    'Agregat otwarty': "Open Set (Unenclosed)",
                    'Obudowa dźwiękochłonna': "Soundproof Canopy",
                    'Obudowa kontenerowa': "Containerized Enclosure"
                },
                'rev_th': ["Rev.", "Description of Change", "Drawn", "Check", "Date"],
                'sec1': "1. INITIAL SUBJECT",
                'sec1_desc': "The aim of this analysis is to establish the calculation data for the designed enclosure to meet the acoustic requirements at the designated measurement points. The mechanical noise generated by the engine was taken as the original source.",
                'sec2': "2. SOURCE DATA",
                'sec2_desc': "A power generator model {genset} with engine {engine} will be placed inside. The source spectrum is presented below:",
                'th1_1': "Octave Band", 'th1_2': "A-Weight k [dB]", 'th1_3': "Provided LwA [dB(A)]",
                'calc_lwa': "Calculated Total Engine Sound Power Level (LwA): ",
                'sec3': "3. ROOM ABSORPTION & OPERATIONAL SOUND PRESSURE",
                'th2_2': "Power LwA [dB(A)]", 'th2_3': "Room Abs. [dB]", 'th2_4': "Pressure LpA [dB(A)]",
                'calc_lpa': "Calculated Total Internal Sound Pressure Level (LpA): ",
                'sec4': "4. ENCLOSURE STRUCTURE & LAYER SPECIFICATIONS",
                'sec4_desc': "The following modules and their detailed layer structures have been included in the band-by-band transmission calculation:",
                'mod_names': {
                    'sciana': "Enclosure Walls", 'podloga': "Floor / Base", 'sufit': "Ceiling", 
                    'drzwi': "Acoustic Door(s)", 'czerpnia': "Air Intake (Louver)", 
                    'wyrzut': "Air Discharge (Louver)", 'wydech': "Engine Exhaust (Silencer)"
                },
                'depth': "Thick./Length",
                'spec_att': "Spectral Attenuation",
                'tot_area': "Total Composite Assembly Area",
                'qty_pcs': "pcs.",
                'sec5': "5. CONCLUSION & COMPLIANCE",
                'dist_corr': "ISO 3744 Distance Correction Factor",
                'ind_noise': "Individual Noise Transmission at Measurement Distance:",
                'tot_noise': "TOTAL COMBINED COMPOSITE NOISE LEVEL",
                'proj_stat': "Point Status: ",
                'pass': "COMPLIANT (PASS).",
                'fail': "NON-COMPLIANT (FAIL).",
                'err_in_use_title': "Save Error / File in Use",
                'err_in_use_msg': "The file:\n{}\nis currently open (e.g., in Word or PDF reader).\n\n• Click 'Yes' to retry (after closing it)\n• Click 'No' to save as a new file\n• Click 'Cancel' to abort saving."
            }
        }

        self.modules_config = {
            'sciana':   {"label": "Ściany", "auto_dim": True, "req": True},
            'podloga':  {"label": "Podłoga", "auto_dim": True, "req": False},
            'sufit':    {"label": "Sufit", "auto_dim": True, "req": False},
            'drzwi':    {"label": "Drzwi", "auto_dim": False, "req": False},
            'czerpnia': {"label": "Czerpnia (Wlot)", "auto_dim": False, "req": False},
            'wyrzut':   {"label": "Wyrzut (Wylot)", "auto_dim": False, "req": False},
            'wydech':   {"label": "Wydech silnika (Komin)", "auto_dim": False, "req": False}
        }

        self.global_config_file = "akustyka_config.json"
        self.db_file = "acoustic_db.json"
        
        if os.path.exists(self.global_config_file):
            try:
                with open(self.global_config_file, 'r', encoding='utf-8') as f:
                    gcfg = json.load(f)
                    if "custom_db_path" in gcfg and os.path.exists(gcfg["custom_db_path"]):
                        self.db_file = gcfg["custom_db_path"]
            except Exception: pass

        self.db_version = DB_VERSION
        self.load_db()

        self.init_vars()
        self.create_menu()
        self.create_widgets()
        self.calc_kontener_surfaces()
        
        self.last_saved_state = self.get_state_json()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        sprawdz_aktualizacje(self.root, tryb="w_tle")
        sprawdz_aktualizacje_bazy(self.root, self, tryb="w_tle")

    def zapisz_globalny_config(self):
        try:
            with open(self.global_config_file, 'w', encoding='utf-8') as f:
                json.dump({"custom_db_path": self.db_file}, f, ensure_ascii=False, indent=4)
        except Exception: pass

    def load_db(self):
        if os.path.exists(self.db_file):
            try:
                with open(self.db_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.db_version = data.get("version", DB_VERSION)
                    self.db_changelog = data.get("changelog", CHANGELOG_DB)
                    self.materials_db = data.get("materials", {})
                    self.engines_db = data.get("engines", {})
                    self.radiators_db = data.get("radiators", {})
            except Exception: pass
        else:
            # Pierwsze uruchomienie: baza startowa z wersją 0.0.0, żeby aktualizacja od razu zaproponowała pełną bazę
            self.db_version = "0.0.0"
            self.db_changelog = CHANGELOG_DB
            self.materials_db = {"Blacha stalowa płaska 2.0mm": {"type": "Widmo pasmowe", "d": 2.0, "spec": [13, 19, 25, 31, 37, 43, 45, 50]}}
            self.engines_db = {}
            self.radiators_db = {}
            if self.db_file == "acoustic_db.json":
                self.save_db()

    def save_db(self):
        try:
            with open(self.db_file, 'w', encoding='utf-8') as f:
                json.dump({
                    "version": self.db_version,
                    "changelog": getattr(self, 'db_changelog', CHANGELOG_DB),
                    "materials": self.materials_db, 
                    "engines": self.engines_db,
                    "radiators": self.radiators_db
                }, f, indent=4, ensure_ascii=False)
        except Exception as e: messagebox.showerror("Błąd", f"Błąd zapisu bazy: {e}")

    def init_vars(self):
        self.current_project_dir = ""
        self.project_no = tk.StringVar(value="PRJ-2026-001")
        self.client_name = tk.StringVar(value="Nazwa Klienta")
        self.po_number = tk.StringVar(value="PO-123456")
        self.client_doc_no = tk.StringVar(value="DOC-001-REV0")
        self.project_name = tk.StringVar(value="Raport Akustyczny")
        self.silnik_typ = tk.StringVar(value="")
        self.agregat_typ = tk.StringVar(value="")
        self.typ_obudowy = tk.StringVar(value="Obudowa kontenerowa")
        self.jezyk_raportu = tk.StringVar(value="Polski")
        self.logo_path = tk.StringVar()
        self.footer_text = tk.StringVar(value="Diamel | diamel.pl")
        self.rev_data = [] 
        self.kont_l = tk.DoubleVar(value=12.0)
        self.kont_w = tk.DoubleVar(value=3.0)
        self.kont_h = tk.DoubleVar(value=3.0)
        self.s_kont = tk.DoubleVar(value=0.0)
        
        self.aktywne = {k: tk.BooleanVar(value=v['req'] or k in ['drzwi']) for k, v in self.modules_config.items()}
        self.mod_dims = {}
        self.mod_layers = {k: [] for k in self.modules_config}
        self.mod_attachments = {k: [] for k in self.modules_config}
        self.points_data = []

        for k in self.modules_config:
            self.mod_dims[k] = {
                'qty': tk.IntVar(value=1), 'w': tk.DoubleVar(value=1.0), 'h': tk.DoubleVar(value=2.0), 
                'd': tk.DoubleVar(value=0.0), 's': tk.DoubleVar(value=0.0)
            }
            if k in ['wydech', 'czerpnia', 'wyrzut']:
                self.mod_dims[k]['kierunek'] = tk.StringVar(value="Poziomy")
                self.mod_dims[k]['wysokosc'] = tk.DoubleVar(value=0.0)
            if not self.modules_config[k]['auto_dim']:
                self.mod_dims[k]['qty'].trace_add("write", self.make_updater(k))
                self.mod_dims[k]['w'].trace_add("write", self.make_updater(k))
                self.mod_dims[k]['h'].trace_add("write", self.make_updater(k))

        self.lw_vars = [tk.DoubleVar(value=95.0) for _ in range(8)]
        self.room_abs_vars = [tk.DoubleVar(value=-3.0) for _ in range(8)]
        
        for var in [self.kont_l, self.kont_w, self.kont_h]:
            var.trace_add("write", self.calc_kontener_surfaces)

        self.layer_ui_frames = {} 
        self.attachment_ui_frames = {}

    def on_closing(self):
        if self.get_state_json() != self.last_saved_state:
            ans = messagebox.askyesnocancel("Niezapisane zmiany", "Masz niezapisane zmiany.\nCzy zapisać projekt przed zamknięciem?")
            if ans is True:
                self.zapisz_projekt()
                self.root.destroy()
            elif ans is False: self.root.destroy()
        else: self.root.destroy()

    def calc_kontener_surfaces(self, *args):
        try:
            l, w, h = self.kont_l.get(), self.kont_w.get(), self.kont_h.get()
            self.s_kont.set(round(2 * (l * h + w * h) + l * w, 2))
            self.mod_dims['podloga']['w'].set(l); self.mod_dims['podloga']['h'].set(w); self.mod_dims['podloga']['s'].set(round(l * w, 2))
            self.mod_dims['sufit']['w'].set(l); self.mod_dims['sufit']['h'].set(w); self.mod_dims['sufit']['s'].set(round(l * w, 2))
            self.recalc_wall_area()
        except tk.TclError: pass

    def recalc_wall_area(self, *args):
        try:
            l, w, h = self.kont_l.get(), self.kont_w.get(), self.kont_h.get()
            wall_total = 2 * (l * h + w * h)
            used_s = sum(self.mod_dims[k]['s'].get() for k in ['drzwi', 'czerpnia', 'wyrzut'] if self.aktywne[k].get())
            wall_s = max(0.0, wall_total - used_s)
            
            self.mod_dims['sciana']['w'].set(round(2*(l+w), 2))
            self.mod_dims['sciana']['h'].set(h)
            self.mod_dims['sciana']['s'].set(round(wall_s, 2))
        except tk.TclError: pass

    def make_updater(self, key):
        def updater(*args):
            try:
                self.mod_dims[key]['s'].set(round(self.mod_dims[key]['w'].get() * self.mod_dims[key]['h'].get() * self.mod_dims[key]['qty'].get(), 2))
                self.recalc_wall_area()
            except tk.TclError: pass
        return updater

    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Nowy Projekt", command=self.nowy_projekt)
        file_menu.add_command(label="Wczytaj Projekt (JSON)...", command=self.wczytaj_projekt)
        file_menu.add_command(label="Zapisz Projekt (JSON)...", command=self.zapisz_projekt)
        file_menu.add_separator()
        file_menu.add_command(label="Zakończ", command=self.on_closing)
        menubar.add_cascade(label="Plik", menu=file_menu)
        
        db_menu = tk.Menu(menubar, tearoff=0)
        db_menu.add_command(label="Edytuj Bazę Silników", command=self.open_engines_db_editor)
        db_menu.add_command(label="Edytuj Bazę Chłodnic", command=self.open_radiators_db_editor)
        db_menu.add_separator()
        db_menu.add_command(label="Edytuj Bazę Materiałów / Warstw", command=self.open_materials_db_editor)
        menubar.add_cascade(label="Baza Danych (Rozszerzenia)", menu=db_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Zgłoś prośbę o dodanie materiału do bazy...", command=lambda: pokaz_okno_zgloszenia_materialu(self.root))
        help_menu.add_separator()
        help_menu.add_command(label="Sprawdź aktualizacje programu", command=lambda: sprawdz_aktualizacje(self.root, "recznie"))
        help_menu.add_command(label="Sprawdź aktualizacje bazy danych", command=lambda: sprawdz_aktualizacje_bazy(self.root, self, "recznie"))
        help_menu.add_separator()
        help_menu.add_command(label="O Programie / Historia zmian", command=self.pokaz_o_programie)
        menubar.add_cascade(label="Pomoc", menu=help_menu)

        self.root.config(menu=menubar)

    def pokaz_o_programie(self):
        pop = tk.Toplevel(self.root)
        pop.title("O Programie / Historia Zmian")
        pop.geometry("800x600")
        
        ttk.Label(pop, text=f"Akustyka - Analiza Obudów v{VERSION}", font=("Arial", 12, "bold")).pack(pady=5)
        ttk.Label(pop, text="Profesjonalne środowisko symulacji zjawisk akustycznych.\n© Diamel / Mateusz Kujawa").pack(pady=5)
        
        paned = ttk.PanedWindow(pop, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=5)
        
        f_app = ttk.LabelFrame(paned, text="Historia zmian programu")
        paned.add(f_app, weight=1)
        txt_app = tk.Text(f_app, wrap="word", font=("Consolas", 9))
        txt_app.pack(fill="both", expand=True, padx=5, pady=5)
        txt_app.insert("1.0", CHANGELOG.strip())
        txt_app.config(state="disabled")
        
        f_db = ttk.LabelFrame(paned, text="Historia zmian bazy danych")
        paned.add(f_db, weight=1)
        txt_db = tk.Text(f_db, wrap="word", font=("Consolas", 9))
        txt_db.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Wczytujemy historię bezpośrednio z zaktualizowanej bazy!
        aktualna_historia = getattr(self, 'db_changelog', CHANGELOG_DB)
        txt_db.insert("1.0", aktualna_historia.strip())
        txt_db.config(state="disabled")

    # --- OKNA EDYCJI BAZY DANYCH ---
    def open_engines_db_editor(self):
        win = tk.Toplevel(self.root)
        win.title("Baza Silników")
        win.geometry("650x500")
        
        frame_list = ttk.Frame(win)
        frame_list.pack(side="left", fill="y", padx=5, pady=5)
        lst = tk.Listbox(frame_list, width=28)
        lst.pack(side="left", fill="y", expand=True)
        sorted_engines = sorted(self.engines_db.keys())
        for e in sorted_engines: lst.insert(tk.END, e)
        
        frame_det = ttk.Frame(win)
        frame_det.pack(side="right", fill="both", expand=True, padx=10, pady=5)
        
        v_eng = tk.StringVar()
        v_gen = tk.StringVar()
        v_lws = [tk.DoubleVar() for _ in range(8)]
        v_exh_lws = [tk.DoubleVar() for _ in range(8)]
        
        ttk.Label(frame_det, text="Nazwa Silnika:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(frame_det, textvariable=v_eng, width=30).grid(row=0, column=1, columnspan=3, sticky="w", pady=2)
        ttk.Label(frame_det, text="Domyślny Model Agregatu:").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(frame_det, textvariable=v_gen, width=30).grid(row=1, column=1, columnspan=3, sticky="w", pady=2)
        
        ttk.Label(frame_det, text="Widmo Lw źródła (Obudowa) [dB]:", font=("", 9, "bold")).grid(row=2, column=0, columnspan=4, sticky="w", pady=(10,2))
        for i, f in enumerate(self.freqs):
            ttk.Label(frame_det, text=f).grid(row=3+i//4, column=(i%4)*2, sticky="e")
            ttk.Entry(frame_det, textvariable=v_lws[i], width=6).grid(row=3+i//4, column=(i%4)*2+1, padx=5, pady=2)

        ttk.Label(frame_det, text="Widmo Lw źródła (Wydech silnika) [dB]:", font=("", 9, "bold")).grid(row=6, column=0, columnspan=4, sticky="w", pady=(10,2))
        for i, f in enumerate(self.freqs):
            ttk.Label(frame_det, text=f).grid(row=7+i//4, column=(i%4)*2, sticky="e")
            ttk.Entry(frame_det, textvariable=v_exh_lws[i], width=6).grid(row=7+i//4, column=(i%4)*2+1, padx=5, pady=2)
            
        def on_select(evt):
            sel = lst.curselection()
            if sel:
                name = lst.get(sel[0])
                data = self.engines_db[name]
                v_eng.set(name)
                v_gen.set(data.get('genset', ''))
                for i in range(8): 
                    v_lws[i].set(data.get('lw', [0]*8)[i])
                    v_exh_lws[i].set(data.get('lw_exh', [0]*8)[i])
        lst.bind("<<ListboxSelect>>", on_select)
        
        def save_eng():
            name = v_eng.get().strip()
            if not name: return
            self.engines_db[name] = {
                "genset": v_gen.get(),
                "lw": [v_lws[i].get() for i in range(8)],
                "lw_exh": [v_exh_lws[i].get() for i in range(8)]
            }
            self.save_db()
            self.refresh_ui_dropdowns()
            lst.delete(0, tk.END)
            for e in sorted(self.engines_db.keys()): lst.insert(tk.END, e)
            messagebox.showinfo("Baza Silników", f"Zapisano silnik: {name}", parent=win)
            
        def del_eng():
            sel = lst.curselection()
            if sel:
                name = lst.get(sel[0])
                if messagebox.askyesno("Usuń", f"Czy na pewno usunąć silnik {name}?", parent=win):
                    del self.engines_db[name]
                    self.save_db()
                    self.refresh_ui_dropdowns()
                    lst.delete(0, tk.END)
                    for e in sorted(self.engines_db.keys()): lst.insert(tk.END, e)
                    v_eng.set(""); v_gen.set("")
                    for i in range(8): 
                        v_lws[i].set(0.0)
                        v_exh_lws[i].set(0.0)

        frame_btns = ttk.Frame(frame_det)
        frame_btns.grid(row=10, column=0, columnspan=4, pady=20)
        ttk.Button(frame_btns, text="Zapisz / Aktualizuj", command=save_eng).pack(side="left", padx=5)
        ttk.Button(frame_btns, text="Usuń zaznaczony", command=del_eng).pack(side="left", padx=5)

    def open_radiators_db_editor(self):
        win = tk.Toplevel(self.root)
        win.title("Baza Chłodnic")
        win.geometry("650x450")
        
        frame_list = ttk.Frame(win)
        frame_list.pack(side="left", fill="y", padx=5, pady=5)
        lst = tk.Listbox(frame_list, width=28)
        lst.pack(side="left", fill="y", expand=True)
        sorted_rads = sorted(self.radiators_db.keys())
        for e in sorted_rads: lst.insert(tk.END, e)
        
        frame_det = ttk.Frame(win)
        frame_det.pack(side="right", fill="both", expand=True, padx=10, pady=5)
        
        v_name = tk.StringVar()
        v_lws = [tk.DoubleVar() for _ in range(8)]
        v_atts = [tk.DoubleVar() for _ in range(8)]
        
        ttk.Label(frame_det, text="Nazwa Chłodnicy:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(frame_det, textvariable=v_name, width=40).grid(row=0, column=1, columnspan=3, sticky="w", pady=2)
        
        ttk.Label(frame_det, text="Generowany Hałas Własny (Lw) [dB] bez wał. A:", font=("", 9, "bold")).grid(row=1, column=0, columnspan=4, sticky="w", pady=(10,2))
        for i, f in enumerate(self.freqs):
            ttk.Label(frame_det, text=f).grid(row=2+i//4, column=(i%4)*2, sticky="e")
            ttk.Entry(frame_det, textvariable=v_lws[i], width=6).grid(row=2+i//4, column=(i%4)*2+1, padx=5, pady=2)
            
        ttk.Label(frame_det, text="Własne Tłumienie Rdzenia [dB]:", font=("", 9, "bold")).grid(row=4, column=0, columnspan=4, sticky="w", pady=(10,2))
        for i, f in enumerate(self.freqs):
            ttk.Label(frame_det, text=f).grid(row=5+i//4, column=(i%4)*2, sticky="e")
            ttk.Entry(frame_det, textvariable=v_atts[i], width=6).grid(row=5+i//4, column=(i%4)*2+1, padx=5, pady=2)
            
        def on_select(evt):
            sel = lst.curselection()
            if sel:
                name = lst.get(sel[0])
                data = self.radiators_db[name]
                v_name.set(name)
                for i in range(8): 
                    v_lws[i].set(data.get('lw', [0]*8)[i])
                    v_atts[i].set(data.get('att', [0]*8)[i])
        lst.bind("<<ListboxSelect>>", on_select)
        
        def save_rad():
            name = v_name.get().strip()
            if not name: return
            self.radiators_db[name] = {
                "lw": [v_lws[i].get() for i in range(8)],
                "att": [v_atts[i].get() for i in range(8)]
            }
            self.save_db()
            self.refresh_ui_dropdowns()
            lst.delete(0, tk.END)
            for e in sorted(self.radiators_db.keys()): lst.insert(tk.END, e)
            messagebox.showinfo("Baza Chłodnic", f"Zapisano chłodnicę: {name}", parent=win)
            
        def del_rad():
            sel = lst.curselection()
            if sel:
                name = lst.get(sel[0])
                if messagebox.askyesno("Usuń", f"Czy usunąć chłodnicę {name}?", parent=win):
                    del self.radiators_db[name]
                    self.save_db()
                    self.refresh_ui_dropdowns()
                    lst.delete(0, tk.END)
                    for e in sorted(self.radiators_db.keys()): lst.insert(tk.END, e)
                    v_name.set("")
                    for i in range(8): v_lws[i].set(0.0); v_atts[i].set(0.0)

        frame_btns = ttk.Frame(frame_det)
        frame_btns.grid(row=8, column=0, columnspan=4, pady=20)
        ttk.Button(frame_btns, text="Zapisz / Aktualizuj", command=save_rad).pack(side="left", padx=5)
        ttk.Button(frame_btns, text="Usuń zaznaczony", command=del_rad).pack(side="left", padx=5)

    def open_materials_db_editor(self):
        win = tk.Toplevel(self.root)
        win.title("Baza Materiałów i Warstw")
        win.geometry("750x500")
        
        frame_list = ttk.Frame(win)
        frame_list.pack(side="left", fill="y", padx=5, pady=5)
        lst = tk.Listbox(frame_list, width=35)
        lst.pack(side="left", fill="y", expand=True)
        sorted_mats = sorted(self.materials_db.keys())
        for m in sorted_mats: lst.insert(tk.END, m)
        
        frame_det = ttk.Frame(win)
        frame_det.pack(side="right", fill="both", expand=True, padx=10, pady=5)
        
        v_name_pl = tk.StringVar()
        v_name_en = tk.StringVar()
        v_desc_pl = tk.StringVar()
        v_desc_en = tk.StringVar()
        v_type = tk.StringVar(value="Rw / Ctr")
        v_d = tk.DoubleVar()
        v_rw = tk.DoubleVar()
        v_ctr = tk.DoubleVar()
        v_spec = [tk.DoubleVar() for _ in range(8)]
        
        ttk.Label(frame_det, text="Nazwa PL:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(frame_det, textvariable=v_name_pl, width=30).grid(row=0, column=1, columnspan=3, sticky="w", pady=2)
        ttk.Label(frame_det, text="Nazwa EN:").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(frame_det, textvariable=v_name_en, width=30).grid(row=1, column=1, columnspan=3, sticky="w", pady=2)
        
        ttk.Label(frame_det, text="Opis PL (Typ/Detale):").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(frame_det, textvariable=v_desc_pl, width=30).grid(row=2, column=1, columnspan=3, sticky="w", pady=2)
        ttk.Label(frame_det, text="Opis EN (Type/Details):").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(frame_det, textvariable=v_desc_en, width=30).grid(row=3, column=1, columnspan=3, sticky="w", pady=2)
        
        ttk.Label(frame_det, text="Typ Tłumienia:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Combobox(frame_det, textvariable=v_type, values=["Rw / Ctr", "Widmo pasmowe"], state="readonly", width=27).grid(row=4, column=1, columnspan=3, sticky="w", pady=2)
        ttk.Label(frame_det, text="Grub./Dług. domyślna [mm]:").grid(row=5, column=0, sticky="w", pady=2)
        ttk.Entry(frame_det, textvariable=v_d, width=10).grid(row=5, column=1, sticky="w", pady=2)
        
        ttk.Label(frame_det, text="Dla typu Rw/Ctr:", font=("", 9, "bold")).grid(row=6, column=0, columnspan=4, sticky="w", pady=(10,2))
        ttk.Label(frame_det, text="Rw [dB]:").grid(row=7, column=0, sticky="w")
        ttk.Entry(frame_det, textvariable=v_rw, width=8).grid(row=7, column=1, sticky="w")
        ttk.Label(frame_det, text="Ctr [dB]:").grid(row=7, column=2, sticky="e")
        ttk.Entry(frame_det, textvariable=v_ctr, width=8).grid(row=7, column=3, sticky="w")
        
        ttk.Label(frame_det, text="Dla typu Widmo pasmowe [dB]:", font=("", 9, "bold")).grid(row=8, column=0, columnspan=4, sticky="w", pady=(10,2))
        for i, f in enumerate(self.freqs):
            ttk.Label(frame_det, text=f).grid(row=9+i//4, column=(i%4)*2, sticky="e")
            ttk.Entry(frame_det, textvariable=v_spec[i], width=6).grid(row=9+i//4, column=(i%4)*2+1, padx=2, pady=2)
            
        def on_select(evt):
            sel = lst.curselection()
            if sel:
                name_pl = lst.get(sel[0])
                data = self.materials_db[name_pl]
                v_name_pl.set(name_pl)
                v_name_en.set(data.get('name_en', ''))
                v_desc_pl.set(data.get('desc_pl', ''))
                v_desc_en.set(data.get('desc_en', ''))
                v_type.set(data.get('type', 'Rw / Ctr'))
                v_d.set(data.get('d', 0.0))
                v_rw.set(data.get('rw', 0.0))
                v_ctr.set(data.get('ctr', 0.0))
                for i in range(8): v_spec[i].set(data.get('spec', [0]*8)[i])
        lst.bind("<<ListboxSelect>>", on_select)
        
        def save_mat():
            name_pl = v_name_pl.get().strip()
            if not name_pl: return
            self.materials_db[name_pl] = {
                "name_en": v_name_en.get().strip(),
                "desc_pl": v_desc_pl.get().strip(),
                "desc_en": v_desc_en.get().strip(),
                "type": v_type.get(), "d": v_d.get(), "rw": v_rw.get(), "ctr": v_ctr.get(),
                "spec": [v_spec[i].get() for i in range(8)]
            }
            self.save_db()
            self.refresh_ui_dropdowns()
            lst.delete(0, tk.END)
            for m in sorted(self.materials_db.keys()): lst.insert(tk.END, m)
            messagebox.showinfo("Baza Materiałów", f"Zapisano materiał: {name_pl}", parent=win)
            
        def del_mat():
            sel = lst.curselection()
            if sel:
                name_pl = lst.get(sel[0])
                if messagebox.askyesno("Usuń", f"Czy usunąć {name_pl}?", parent=win):
                    del self.materials_db[name_pl]
                    self.save_db()
                    self.refresh_ui_dropdowns()
                    lst.delete(0, tk.END)
                    for m in sorted(self.materials_db.keys()): lst.insert(tk.END, m)
                    v_name_pl.set(""); v_name_en.set(""); v_desc_pl.set(""); v_desc_en.set("")

        frame_btns = ttk.Frame(frame_det)
        frame_btns.grid(row=13, column=0, columnspan=4, pady=15)
        ttk.Button(frame_btns, text="Zapisz / Aktualizuj", command=save_mat).pack(side="left", padx=5)
        ttk.Button(frame_btns, text="Usuń zaznaczony", command=del_mat).pack(side="left", padx=5)

    def refresh_ui_dropdowns(self):
        if hasattr(self, 'cb_engine'):
            self.cb_engine['values'] = sorted(list(self.engines_db.keys()))
        if hasattr(self, 'cb_radiator'):
            self.cb_radiator['values'] = sorted(list(self.radiators_db.keys()))
        
        mat_keys = ["[Własna warstwa]"] + sorted(list(self.materials_db.keys()))
        for k in self.mod_layers:
            for layer in self.mod_layers[k]:
                if 'cb_mat' in layer:
                    layer['cb_mat']['values'] = mat_keys

    # --- BUDOWA GŁÓWNEGO GUI ---
    def create_widgets(self):
        frame_bottom = ttk.Frame(self.root)
        frame_bottom.pack(fill="x", pady=5, side="bottom")
        ttk.Button(frame_bottom, text="OBLICZ WYNIKI (PODGLĄD PASMOWY)", command=self.pokaz_obliczenia, style="Accent.TButton").pack(side="top", fill="x", padx=20, pady=2, ipady=5)
        
        frame_gen = ttk.Frame(frame_bottom)
        frame_gen.pack(fill="x", padx=20, pady=2)
        ttk.Button(frame_gen, text="GENERUJ RAPORT (DOCX)", command=lambda: self.generuj_raport(mode='docx')).pack(side="left", expand=True, fill="x", padx=5, ipady=8)
        if PDF_MERGE_AVAILABLE:
            ttk.Button(frame_gen, text="GENERUJ PDF (Bez załączników)", command=lambda: self.generuj_raport(mode='pdf_no_att')).pack(side="left", expand=True, fill="x", padx=5, ipady=8)
            ttk.Button(frame_gen, text="GENERUJ PDF (Kompletny z zał.)", command=lambda: self.generuj_raport(mode='pdf_full')).pack(side="left", expand=True, fill="x", padx=5, ipady=8)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=5, expand=True, fill='both')

        t_main = ttk.Frame(self.notebook); self.notebook.add(t_main, text="1. Dane Ogólne")
        s_main = ScrollableFrame(t_main); s_main.pack(fill="both", expand=True)
        self.build_main_tab(s_main.scrollable_frame)

        t_ac = ttk.Frame(self.notebook); self.notebook.add(t_ac, text="2. Widmo Silnika")
        self.build_acoustics_tab(t_ac)

        t_rad = ttk.Frame(self.notebook); self.notebook.add(t_rad, text="3. Chłodnica")
        self.build_radiator_tab(t_rad)
        
        t_pts = ttk.Frame(self.notebook); self.notebook.add(t_pts, text="4. Punkty Pomiarowe")
        self.build_points_tab(t_pts)

        for key, cfg in self.modules_config.items():
            t_m = ttk.Frame(self.notebook); self.notebook.add(t_m, text=cfg['label'])
            self.build_module_tab(t_m, key, cfg)
            
        t_settings = ttk.Frame(self.notebook); self.notebook.add(t_settings, text="Ustawienia / Baza")
        self.build_settings_tab(t_settings)

    def build_settings_tab(self, parent):
        frame_db = ttk.LabelFrame(parent, text="Zarządzanie plikiem Bazy Danych (Silniki, Chłodnice, Materiały)")
        frame_db.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(frame_db, text="Aktualna ścieżka do bazy:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.lbl_db_path = ttk.Label(frame_db, text=os.path.abspath(self.db_file), foreground="blue", wraplength=800)
        self.lbl_db_path.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Button(frame_db, text="Wskaż inny plik bazy (JSON)...", command=self.zmien_plik_bazy).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(frame_db, text="Przywróć domyślną bazę (acoustic_db.json)", command=self.przywroc_domyslna_baze).grid(row=1, column=1, padx=5, pady=5, sticky="w")

    def zmien_plik_bazy(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filepath:
            self.db_file = filepath
            self.lbl_db_path.config(text=os.path.abspath(self.db_file))
            self.zapisz_globalny_config()
            self.load_db()
            self.refresh_ui_dropdowns()
            messagebox.showinfo("Baza zmieniona", f"Pomyślnie załadowano bazę z:\n{filepath}")

    def przywroc_domyslna_baze(self):
        self.db_file = "acoustic_db.json"
        self.lbl_db_path.config(text=os.path.abspath(self.db_file))
        self.zapisz_globalny_config()
        self.load_db()
        self.refresh_ui_dropdowns()
        messagebox.showinfo("Baza zresetowana", "Przywrócono domyślny plik bazy danych.")

    def browse_logo(self):
        file = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.webp")])
        if file: self.logo_path.set(file)

    def build_main_tab(self, parent):
        frame_proj = ttk.LabelFrame(parent, text="Metadane Dokumentu (Strona Tytułowa)")
        frame_proj.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(frame_proj, text="Numer Projektu:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_proj, textvariable=self.project_no, width=35).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(frame_proj, text="Tytuł Projektu:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_proj, textvariable=self.project_name, width=35).grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(frame_proj, text="Dla Klienta:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_proj, textvariable=self.client_name, width=35).grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(frame_proj, text="Numer P.O.:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_proj, textvariable=self.po_number, width=35).grid(row=1, column=3, padx=5, pady=5)
        ttk.Label(frame_proj, text="Dokument Klienta (Nr):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_proj, textvariable=self.client_doc_no, width=35).grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(frame_proj, text="Typ Silnika (Baza lub Wpisz):").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.cb_engine = ttk.Combobox(frame_proj, textvariable=self.silnik_typ, values=sorted(list(self.engines_db.keys())), width=32)
        self.cb_engine.grid(row=3, column=1, padx=5, pady=5)
        
        def on_engine_select(event):
            eng = self.silnik_typ.get()
            if eng in self.engines_db:
                self.agregat_typ.set(self.engines_db[eng].get("genset", ""))
                lws = self.engines_db[eng].get("lw", [0]*8)
                exh_lws = self.engines_db[eng].get("lw_exh", [0]*8)
                for i in range(8): 
                    self.lw_vars[i].set(lws[i])
                    self.exhaust_lw_vars[i].set(exh_lws[i])
        self.cb_engine.bind("<<ComboboxSelected>>", on_engine_select)

        ttk.Label(frame_proj, text="Typ Agregatu:").grid(row=3, column=2, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_proj, textvariable=self.agregat_typ, width=35).grid(row=3, column=3, padx=5, pady=5)

        ttk.Label(frame_proj, text="Typ Instalacji:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        cb_typ = ttk.Combobox(frame_proj, textvariable=self.typ_obudowy, state="readonly", width=32)
        cb_typ['values'] = ('Agregat otwarty', 'Obudowa dźwiękochłonna', 'Obudowa kontenerowa')
        cb_typ.grid(row=4, column=1, padx=5, pady=5)
        
        ttk.Label(frame_proj, text="Język Raportu:").grid(row=4, column=2, padx=5, pady=5, sticky="w")
        cb_lang = ttk.Combobox(frame_proj, textvariable=self.jezyk_raportu, state="readonly", width=32)
        cb_lang['values'] = ('Polski', 'English')
        cb_lang.grid(row=4, column=3, padx=5, pady=5)

        frame_revs = ttk.LabelFrame(parent, text="Tabela Rewizji (Historia zmian dokumentu)")
        frame_revs.pack(fill="x", padx=10, pady=5)
        ttk.Button(frame_revs, text="+ Dodaj Rewizję", command=self.ui_add_rev_row).pack(anchor="nw", padx=5, pady=5)
        self.rev_container = ttk.Frame(frame_revs)
        self.rev_container.pack(fill="x", expand=True, padx=5, pady=5)
        self.ui_add_rev_row(rev="1", desc="Wydanie 1 / Initial Issue", drawn="MK", check="Appr.")

        frame_dim = ttk.LabelFrame(parent, text="Główne Gabaryty Kontenera (Wyznaczają bazę referencyjną do dystansu)")
        frame_dim.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(frame_dim, text="Długość L [m]:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(frame_dim, textvariable=self.kont_l, width=10).grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(frame_dim, text="Szerokość W [m]:").grid(row=0, column=2, padx=5, pady=2, sticky="w")
        ttk.Entry(frame_dim, textvariable=self.kont_w, width=10).grid(row=0, column=3, padx=5, pady=2)
        ttk.Label(frame_dim, text="Wysokość H [m]:").grid(row=0, column=4, padx=5, pady=2, sticky="w")
        ttk.Entry(frame_dim, textvariable=self.kont_h, width=10).grid(row=0, column=5, padx=5, pady=2)

        frame_style = ttk.LabelFrame(parent, text="Stylizacja i Szybkie Wczytywanie Logo")
        frame_style.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame_style, text="Logo firmy (ścieżka):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_style, textvariable=self.logo_path, width=40).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_style, text="Wstaw Genpower", command=lambda: self.logo_path.set(sciezka_do_zasobu(os.path.join("Zasoby", "logo_genpower.jpg")))).grid(row=0, column=2, padx=2)
        ttk.Button(frame_style, text="Wstaw Diamel", command=lambda: self.logo_path.set(sciezka_do_zasobu(os.path.join("Zasoby", "logo_duze_diamel.png")))).grid(row=0, column=3, padx=2)
        ttk.Button(frame_style, text="Przeglądaj (Własne)...", command=lambda: self.logo_path.set(filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp")]) or self.logo_path.get())).grid(row=0, column=4, padx=5)

        ttk.Label(frame_style, text="Tekst stopki:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_style, textvariable=self.footer_text, width=70).grid(row=1, column=1, columnspan=4, padx=5, pady=5, sticky="w")

    def ui_add_rev_row(self, rev="", desc="", drawn="", check="", date=None):
        if date is None: date = datetime.now().strftime('%Y-%m-%d')
        row_frame = ttk.Frame(self.rev_container)
        row_frame.pack(fill="x", pady=2)
        
        v_rev, v_desc, v_drawn, v_check, v_date = tk.StringVar(value=rev), tk.StringVar(value=desc), tk.StringVar(value=drawn), tk.StringVar(value=check), tk.StringVar(value=date)
        rev_dict = {'rev': v_rev, 'desc': v_desc, 'drawn': v_drawn, 'check': v_check, 'date': v_date, 'frame': row_frame}
        self.rev_data.append(rev_dict)
        
        ttk.Label(row_frame, text="Rew:").pack(side="left")
        ttk.Entry(row_frame, textvariable=v_rev, width=4).pack(side="left", padx=2)
        ttk.Label(row_frame, text="Opis:").pack(side="left")
        ttk.Entry(row_frame, textvariable=v_desc, width=25).pack(side="left", padx=2)
        ttk.Label(row_frame, text="Wyk:").pack(side="left")
        ttk.Entry(row_frame, textvariable=v_drawn, width=5).pack(side="left", padx=2)
        ttk.Label(row_frame, text="Spr:").pack(side="left")
        ttk.Entry(row_frame, textvariable=v_check, width=5).pack(side="left", padx=2)
        ttk.Label(row_frame, text="Data:").pack(side="left")
        ttk.Entry(row_frame, textvariable=v_date, width=12).pack(side="left", padx=2)
        ttk.Button(row_frame, text="Usuń", width=5, command=lambda: self.ui_remove_rev_row(rev_dict)).pack(side="right", padx=5)

    def ui_remove_rev_row(self, rev_dict):
        rev_dict['frame'].destroy()
        self.rev_data.remove(rev_dict)

    def build_points_tab(self, parent):
        frame_container = ttk.LabelFrame(parent, text="Zdefiniowane Punkty Pomiarowe (Różne dystanse i limity)")
        frame_container.pack(fill="both", expand=True, padx=10, pady=10)

        btn_add = ttk.Button(frame_container, text="+ Dodaj Punkt Pomiarowy", command=self.ui_add_point_row)
        btn_add.pack(anchor="nw", padx=5, pady=5)

        scroll_frame = ScrollableFrame(frame_container)
        scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.points_ui_frame = scroll_frame.scrollable_frame
        
        self.ui_add_point_row(name="Standardowy", d=1.0, limit=85.0)

    def ui_add_point_row(self, name="Punkt Pomiarowy", d=1.0, limit=85.0):
        row_frame = ttk.Frame(self.points_ui_frame, relief="groove", borderwidth=1)
        row_frame.pack(fill="x", pady=2, ipadx=5, ipady=5)

        v_name = tk.StringVar(value=name)
        v_d = tk.DoubleVar(value=d)
        v_limit = tk.DoubleVar(value=limit)

        pt_data = {'name': v_name, 'd': v_d, 'limit': v_limit, 'frame': row_frame}
        self.points_data.append(pt_data)

        ttk.Label(row_frame, text="Nazwa punktu:").pack(side="left", padx=5)
        ttk.Entry(row_frame, textvariable=v_name, width=25).pack(side="left", padx=5)

        ttk.Label(row_frame, text="Odległość 'd' od kontenera [m]:").pack(side="left", padx=10)
        ttk.Entry(row_frame, textvariable=v_d, width=8).pack(side="left", padx=5)

        ttk.Label(row_frame, text="Wymagany Limit [dB(A)]:").pack(side="left", padx=10)
        ttk.Entry(row_frame, textvariable=v_limit, width=8).pack(side="left", padx=5)

        ttk.Button(row_frame, text="Usuń", width=6, command=lambda: self.ui_remove_point(pt_data)).pack(side="right", padx=5)

    def ui_remove_point(self, pt_data):
        pt_data['frame'].destroy()
        self.points_data.remove(pt_data)

    def build_acoustics_tab(self, parent):
        frame_main = ScrollableFrame(parent)
        frame_main.pack(fill="both", expand=True, padx=5, pady=5)
        content = frame_main.scrollable_frame

        frame_top = ttk.Frame(content)
        frame_top.pack(fill="x", padx=5, pady=5)

        # TABELA
        frame_table = ttk.LabelFrame(frame_top, text="Widmo mocy akustycznej silnika [dB] oraz obliczone LwA")
        frame_table.pack(side="left", fill="y", padx=(0, 10))
        
        ttk.Label(frame_table, text="Częstotliwość").grid(row=0, column=0, padx=10, pady=5)
        ttk.Label(frame_table, text="Lw źródła [dB]").grid(row=0, column=1, padx=10, pady=5)
        ttk.Label(frame_table, text="Room Abs. [dB]").grid(row=0, column=2, padx=10, pady=5)
        ttk.Label(frame_table, text="LwA pasmo [dB(A)]").grid(row=0, column=3, padx=10, pady=5)
        
        self.engine_lwa_labels = []
        for i, freq in enumerate(self.freqs):
            ttk.Label(frame_table, text=freq+"Hz").grid(row=i+1, column=0, padx=10, pady=2)
            ttk.Entry(frame_table, textvariable=self.lw_vars[i], width=10).grid(row=i+1, column=1, padx=10, pady=2)
            ttk.Entry(frame_table, textvariable=self.room_abs_vars[i], width=10).grid(row=i+1, column=2, padx=10, pady=2)
            
            lbl_band_lwa = ttk.Label(frame_table, text="0.0", foreground="blue", font=("", 9, "bold"))
            lbl_band_lwa.grid(row=i+1, column=3, padx=10, pady=2)
            self.engine_lwa_labels.append(lbl_band_lwa)

        self.engine_total_lwa_lbl = ttk.Label(frame_table, text="Całkowite LwA silnika: 0.0 dB(A)", font=("", 10, "bold"), foreground="darkgreen")
        self.engine_total_lwa_lbl.grid(row=len(self.freqs)+1, column=0, columnspan=4, pady=15)

        # KALKULATOR SABINE'A
        frame_calc = ttk.LabelFrame(frame_top, text="Kalkulator Absorpcji Wnętrza (Wzór Sabine'a)")
        frame_calc.pack(side="left", fill="both", expand=True)

        ttk.Label(frame_calc, text="Średni wsp. pochłaniania podłogi (α):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        ttk.Entry(frame_calc, textvariable=self.alfa_podl, width=10).grid(row=0, column=1, padx=10, pady=10)

        ttk.Label(frame_calc, text="Średni wsp. pochłaniania ścian i sufitu (α):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        ttk.Entry(frame_calc, textvariable=self.alfa_scian, width=10).grid(row=1, column=1, padx=10, pady=10)

        ttk.Label(frame_calc, text="Współczynnik kierunkowości źródła Q:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        ttk.Entry(frame_calc, textvariable=self.sabine_q, width=10).grid(row=2, column=1, padx=10, pady=10)

        ttk.Label(frame_calc, text="Odległość śr. pomiaru od źródła r [m]:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        ttk.Entry(frame_calc, textvariable=self.sabine_r, width=10).grid(row=3, column=1, padx=10, pady=10)

        ttk.Button(frame_calc, text="Oblicz i wpisz do tabeli obok", command=self.calculate_sabine, style="Accent.TButton").grid(row=4, column=0, columnspan=2, pady=15, ipadx=10, ipady=5)
        
        self.sabine_result_lbl = ttk.Label(frame_calc, text="Ostatni wynik Room Abs.: --- dB", font=("", 10, "bold"), foreground="blue")
        self.sabine_result_lbl.grid(row=5, column=0, columnspan=2, pady=5)

        # Dolna sekcja: Wizualizacja / Karta Silnika
        frame_files_engine = ttk.LabelFrame(content, text="Załączniki dla Silnika (Obrazki PNG do raportu + Karty PDF na koniec)")
        frame_files_engine.pack(fill="x", padx=5, pady=10)
        
        frame_img_eng = ttk.Frame(frame_files_engine)
        frame_img_eng.pack(fill="x", pady=2)
        ttk.Label(frame_img_eng, text="Plik graficzny (PNG/JPG):").pack(side="left", padx=5)
        ttk.Entry(frame_img_eng, textvariable=self.engine_wiz_path, width=50, state="readonly").pack(side="left", padx=5)
        ttk.Button(frame_img_eng, text="Wybierz...", command=lambda: self.browse_general_image(self.engine_wiz_path)).pack(side="left", padx=5)
        ttk.Button(frame_img_eng, text="X", width=3, command=lambda: self.engine_wiz_path.set("")).pack(side="left", padx=2)

        frame_pdf_eng = ttk.Frame(frame_files_engine)
        frame_pdf_eng.pack(fill="x", pady=2)
        ttk.Label(frame_pdf_eng, text="Karta Techniczna (PDF):").pack(side="left", padx=5)
        ttk.Entry(frame_pdf_eng, textvariable=self.engine_pdf_path, width=50, state="readonly").pack(side="left", padx=5)
        ttk.Button(frame_pdf_eng, text="Wybierz...", command=lambda: self.browse_pdf(self.engine_pdf_path)).pack(side="left", padx=5)
        ttk.Button(frame_pdf_eng, text="X", width=3, command=lambda: self.engine_pdf_path.set("")).pack(side="left", padx=2)

        for var in self.lw_vars:
            var.trace_add("write", self.update_engine_lwa_labels)
        self.update_engine_lwa_labels()

    def update_engine_lwa_labels(self, *args):
        try:
            lws = [v.get() for v in self.lw_vars]
            band_lwas = [lws[i] + self.k_dba_default[i] for i in range(8)]
            for i in range(8):
                self.engine_lwa_labels[i].config(text=f"{band_lwas[i]:.1f}")
            tot = 10 * math.log10(sum(10**(v/10) for v in band_lwas))
            self.engine_total_lwa_lbl.config(text=f"Całkowite LwA silnika: {tot:.1f} dB(A)")
        except: pass

    def calculate_sabine(self):
        try:
            l = self.kont_l.get()
            w = self.kont_w.get()
            h = self.kont_h.get()
            
            s_podl = l * w
            s_sufit = l * w
            s_sciany = 2 * (l * h + w * h)
            s_tot = s_podl + s_sufit + s_sciany
            
            if s_tot <= 0: return

            a_p = self.alfa_podl.get()
            a_s = self.alfa_scian.get()
            q = self.sabine_q.get()
            r = self.sabine_r.get()
            
            x_med = (a_p * s_podl + a_s * (s_sciany + s_sufit)) / s_tot
            if x_med >= 1.0: x_med = 0.99
            
            room_const = (s_tot * x_med) / (1 - x_med)
            if room_const <= 0: return
            
            direct = q / (4 * math.pi * (r**2))
            reverb = 4 / room_const
            
            room_abs = 10 * math.log10(direct + reverb)
            room_abs = round(room_abs, 2)
            
            for i in range(8):
                self.room_abs_vars[i].set(room_abs)
                
            self.sabine_result_lbl.config(text=f"Ostatni wynik Room Abs.: {room_abs} dB")
        except Exception as e:
            messagebox.showerror("Błąd", f"Błąd w obliczeniach Sabine: {e}")

    def build_radiator_tab(self, parent):
        frame_top = ttk.Frame(parent)
        frame_top.pack(fill="x", padx=10, pady=10)
        
        ttk.Checkbutton(frame_top, text="Włącz CHŁODNICĘ NA WYRZUCIE (Integruje hałas chłodnicy z tłumikami wyrzutni)", variable=self.rad_active, command=self.toggle_radiator_state).grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
        
        self.rad_frame = ttk.LabelFrame(parent, text="Parametry Chłodnicy (Baza Danych)")
        self.rad_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        ttk.Label(self.rad_frame, text="Wybierz z bazy (auto-wypełnienie):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.cb_radiator = ttk.Combobox(self.rad_frame, values=sorted(list(self.radiators_db.keys())), state="readonly", width=40)
        self.cb_radiator.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        ttk.Label(self.rad_frame, text="Nazwa / Model w raporcie:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.rad_name_ent = ttk.Entry(self.rad_frame, textvariable=self.rad_name, width=43)
        self.rad_name_ent.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        
        def on_rad_select(event):
            rad = self.cb_radiator.get()
            if rad in self.radiators_db:
                self.rad_name.set(rad)
                lws = self.radiators_db[rad].get("lw", [0]*8)
                atts = self.radiators_db[rad].get("att", [0]*8)
                for i in range(8):
                    self.rad_lw_vars[i].set(lws[i])
                    self.rad_att_vars[i].set(atts[i])
        self.cb_radiator.bind("<<ComboboxSelected>>", on_rad_select)

        frame_tables = ttk.Frame(self.rad_frame)
        frame_tables.grid(row=2, column=0, columnspan=2, pady=15, sticky="w")
        
        ttk.Label(frame_tables, text="Częstotliwość [Hz]").grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(frame_tables, text="Hałas Własny (Lw) [dB]").grid(row=0, column=1, padx=15, pady=5)
        ttk.Label(frame_tables, text="Poprawka A [dB]").grid(row=0, column=2, padx=15, pady=5)
        ttk.Label(frame_tables, text="LwA chłodnicy [dB(A)]").grid(row=0, column=3, padx=15, pady=5)
        ttk.Label(frame_tables, text="Tłumienie Rdzenia [dB]").grid(row=0, column=4, padx=15, pady=5)
        
        self.rad_lwa_labels = []
        for i, freq in enumerate(self.freqs):
            ttk.Label(frame_tables, text=freq).grid(row=i+1, column=0, padx=5, pady=2)
            ttk.Entry(frame_tables, textvariable=self.rad_lw_vars[i], width=10).grid(row=i+1, column=1, padx=15, pady=2)
            ttk.Label(frame_tables, text=str(self.k_dba_default[i])).grid(row=i+1, column=2, padx=15, pady=2)
            
            lbl_lwa = ttk.Label(frame_tables, text="0.0", foreground="blue", font=("", 9, "bold"))
            lbl_lwa.grid(row=i+1, column=3, padx=15, pady=2)
            self.rad_lwa_labels.append(lbl_lwa)
            
            ttk.Entry(frame_tables, textvariable=self.rad_att_vars[i], width=10).grid(row=i+1, column=4, padx=15, pady=2)

        self.rad_total_lwa_lbl = ttk.Label(frame_tables, text="Całkowite LwA chłodnicy: 0.0 dB(A)", font=("", 10, "bold"), foreground="darkgreen")
        self.rad_total_lwa_lbl.grid(row=len(self.freqs)+1, column=0, columnspan=5, pady=10)

        frame_wiz_rad = ttk.LabelFrame(self.rad_frame, text="Załączniki dla Chłodnicy (Obrazki PNG do raportu + Karty PDF na koniec)")
        frame_wiz_rad.grid(row=3, column=0, columnspan=2, padx=5, pady=10, sticky="ew")
        
        frame_img_rad = ttk.Frame(frame_wiz_rad)
        frame_img_rad.pack(fill="x", pady=2)
        ttk.Label(frame_img_rad, text="Plik graficzny (PNG/JPG):").pack(side="left", padx=5)
        ttk.Entry(frame_img_rad, textvariable=self.rad_wiz_path, width=40, state="readonly").pack(side="left", padx=5)
        ttk.Button(frame_img_rad, text="Wybierz...", command=lambda: self.browse_general_image(self.rad_wiz_path)).pack(side="left", padx=5)
        ttk.Button(frame_img_rad, text="X", width=3, command=lambda: self.rad_wiz_path.set("")).pack(side="left", padx=2)

        frame_pdf_rad = ttk.Frame(frame_wiz_rad)
        frame_pdf_rad.pack(fill="x", pady=2)
        ttk.Label(frame_pdf_rad, text="Karta Techniczna (PDF):").pack(side="left", padx=5)
        ttk.Entry(frame_pdf_rad, textvariable=self.rad_pdf_path, width=40, state="readonly").pack(side="left", padx=5)
        ttk.Button(frame_pdf_rad, text="Wybierz...", command=lambda: self.browse_pdf(self.rad_pdf_path)).pack(side="left", padx=5)
        ttk.Button(frame_pdf_rad, text="X", width=3, command=lambda: self.rad_pdf_path.set("")).pack(side="left", padx=2)

        def update_rad_lwa(*args):
            try:
                lws = [v.get() for v in self.rad_lw_vars]
                band_lwas = [lws[i] + self.k_dba_default[i] for i in range(8)]
                for i in range(8):
                    self.rad_lwa_labels[i].config(text=f"{band_lwas[i]:.1f}")
                tot = 10 * math.log10(sum(10**(v/10) for v in band_lwas))
                self.rad_total_lwa_lbl.config(text=f"Całkowite LwA chłodnicy: {tot:.1f} dB(A)")
            except: pass

        for var in self.rad_lw_vars:
            var.trace_add("write", update_rad_lwa)
        update_rad_lwa()
        self.toggle_radiator_state()

    def browse_general_image(self, var):
        file = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.webp")])
        if file: var.set(file)

    def browse_pdf(self, var):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file: var.set(file)

    def toggle_radiator_state(self):
        state = 'normal' if self.rad_active.get() else 'disabled'
        self.cb_radiator.config(state='readonly' if self.rad_active.get() else 'disabled')
        self.rad_name_ent.config(state=state)
        for child in self.rad_frame.winfo_children():
            if isinstance(child, ttk.Frame):
                for subchild in child.winfo_children():
                    if isinstance(subchild, ttk.Entry):
                        subchild.config(state=state)

    def build_module_tab(self, parent, key, cfg):
        frame_top = ttk.Frame(parent)
        frame_top.pack(fill="x", padx=10, pady=5)
        
        if not cfg['req']:
            ttk.Checkbutton(frame_top, text=f"Włącz moduł: {cfg['label']}", variable=self.aktywne[key], command=lambda k=key: self.toggle_tab_state(k)).pack(side="left", padx=5)
        else:
            ttk.Label(frame_top, text=f"Moduł podstawowy (Zawsze aktywny)").pack(side="left", padx=5)
            
        if key == 'wydech':
            frame_exh = ttk.LabelFrame(parent, text="Widmo Mocy Akustycznej Wydechu (Lw) [dB] bez wał. A (DŹWIĘK PRZED TŁUMIKIEM)")
            frame_exh.pack(fill="x", padx=10, pady=5)
            for i, f in enumerate(self.freqs):
                ttk.Label(frame_exh, text=f).grid(row=0, column=i*2, padx=(10,2), pady=5, sticky="e")
                ttk.Entry(frame_exh, textvariable=self.exhaust_lw_vars[i], width=6).grid(row=0, column=i*2+1, padx=2, pady=5)

        frame_dim = ttk.LabelFrame(parent, text="Wymiary Całkowite Modułu / Opcje geometryczne")
        frame_dim.pack(fill="x", padx=10, pady=5)
        
        row = 0
        ttk.Label(frame_dim, text="Sztuk:").grid(row=row, column=0, padx=5, pady=5, sticky="w")
        state_q = "readonly" if cfg['auto_dim'] else "normal"
        ent_q = ttk.Entry(frame_dim, textvariable=self.mod_dims[key]['qty'], width=5, state=state_q)
        ent_q.grid(row=row, column=1, padx=5, pady=5)

        if key != 'wydech':
            ttk.Label(frame_dim, text="Szer. W [m]:").grid(row=row, column=2, padx=5, pady=5, sticky="w")
            state_w = "readonly" if cfg['auto_dim'] else "normal"
            ent_w = ttk.Entry(frame_dim, textvariable=self.mod_dims[key]['w'], width=8, state=state_w)
            ent_w.grid(row=row, column=3, padx=5, pady=5)
            
            ttk.Label(frame_dim, text="Wys. H [m]:").grid(row=row, column=4, padx=5, pady=5, sticky="w")
            ent_h = ttk.Entry(frame_dim, textvariable=self.mod_dims[key]['h'], width=8, state=state_w)
            ent_h.grid(row=row, column=5, padx=5, pady=5)
                
            ttk.Label(frame_dim, text="Grub./Dług. [mm]:").grid(row=row, column=6, padx=5, pady=5, sticky="w")
            ent_d = ttk.Entry(frame_dim, textvariable=self.mod_dims[key]['d'], width=8, state="readonly", foreground="blue")
            ent_d.grid(row=row, column=7, padx=5, pady=5)

            ttk.Label(frame_dim, text="Pole S [m2]:").grid(row=row, column=8, padx=5, pady=5, sticky="w")
            ent_s = ttk.Entry(frame_dim, textvariable=self.mod_dims[key]['s'], width=8, state="readonly", foreground="blue")
            ent_s.grid(row=row, column=9, padx=5, pady=5)
            
            self.modules_config[key]['ui_entries'] = [ent_q, ent_w, ent_h]

            # Dodatkowe opcje kierunku i wysokości dla czerpni i wyrzutu
            if key in ['czerpnia', 'wyrzut']:
                row_special = 1
                ttk.Label(frame_dim, text="Kierunek emisji hałasu:").grid(row=row_special, column=2, padx=5, pady=5, sticky="w")
                cb_dir = ttk.Combobox(frame_dim, textvariable=self.mod_dims[key]['kierunek'], values=["Poziomy", "Pionowy w górę"], state="readonly", width=15)
                cb_dir.grid(row=row_special, column=3, padx=5, pady=5)
                
                ttk.Label(frame_dim, text="Lokalizacja/Wysokość nad kontener [m]:").grid(row=row_special, column=4, padx=5, pady=5, sticky="w")
                ent_wys = ttk.Entry(frame_dim, textvariable=self.mod_dims[key]['wysokosc'], width=8)
                ent_wys.grid(row=row_special, column=5, padx=5, pady=5)
                self.modules_config[key]['ui_entries'].extend([cb_dir, ent_wys])

        else:
            # Dla wydechu ukrywamy W i H, pokazujemy od razu kierunek
            ttk.Label(frame_dim, text="Kierunek wylotu:").grid(row=row, column=2, padx=5, pady=5, sticky="w")
            cb_dir = ttk.Combobox(frame_dim, textvariable=self.mod_dims[key]['kierunek'], values=["Poziomy", "Pionowy w górę"], state="readonly", width=15)
            cb_dir.grid(row=row, column=3, padx=5, pady=5)
            
            ttk.Label(frame_dim, text="Dodatkowa odległość od obwiedni [m]:").grid(row=row, column=4, padx=5, pady=5, sticky="w")
            ent_wys = ttk.Entry(frame_dim, textvariable=self.mod_dims[key]['wysokosc'], width=8)
            ent_wys.grid(row=row, column=5, padx=5, pady=5)
            
            ttk.Label(frame_dim, text="Dług. tłumika [mm]:").grid(row=row, column=6, padx=5, pady=5, sticky="w")
            ent_d = ttk.Entry(frame_dim, textvariable=self.mod_dims[key]['d'], width=8, state="readonly", foreground="blue")
            ent_d.grid(row=row, column=7, padx=5, pady=5)
            self.modules_config[key]['ui_entries'] = [ent_q, cb_dir, ent_wys]

        frame_layers_container = ttk.LabelFrame(parent, text="Budowa Warstwowa / Konfiguracja (Wybierz materiał z bazy lub wpisz własny)")
        frame_layers_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        btn_add = ttk.Button(frame_layers_container, text="+ Dodaj Warstwę / Element", command=lambda k=key: self.ui_add_layer_row(k))
        btn_add.pack(anchor="nw", padx=5, pady=5)
        
        scroll_frame = ScrollableFrame(frame_layers_container)
        scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
        scroll_frame.canvas.config(height=200) 
        self.layer_ui_frames[key] = scroll_frame.scrollable_frame
        
        frame_files = ttk.LabelFrame(parent, text="Załączniki dla modułu (Wizualizacje PNG wklejane w tekście / Karty PDF dodawane na końcu)")
        frame_files.pack(fill="x", padx=10, pady=5)
        btn_add_att = ttk.Button(frame_files, text="+ Dodaj załącznik", command=lambda k=key: self.ui_add_attachment_row(k))
        btn_add_att.pack(anchor="nw", padx=5, pady=5)
        
        att_scroll = ScrollableFrame(frame_files)
        att_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        att_scroll.canvas.config(height=80) 
        self.attachment_ui_frames[key] = att_scroll.scrollable_frame

        if key != 'wydech':
            self.modules_config[key]['ui_entries'].extend([btn_add, btn_add_att])
        else:
            self.modules_config[key]['ui_entries'].extend([btn_add, btn_add_att])
            
        self.toggle_tab_state(key)

    def ui_add_attachment_row(self, mod_key, type_val="Obraz (PNG/JPG)", path_val=""):
        parent = self.attachment_ui_frames[mod_key]
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill="x", pady=2)

        v_type = tk.StringVar(value=type_val)
        v_path = tk.StringVar(value=path_val)

        att_data = {'type': v_type, 'path': v_path, 'frame': row_frame}
        self.mod_attachments[mod_key].append(att_data)

        cb = ttk.Combobox(row_frame, textvariable=v_type, values=["Obraz (PNG/JPG)", "Dokument (PDF)"], state="readonly", width=18)
        cb.pack(side="left", padx=2)

        ent = ttk.Entry(row_frame, textvariable=v_path, width=50, state="readonly")
        ent.pack(side="left", padx=2)

        def browse():
            if v_type.get() == "Obraz (PNG/JPG)":
                f = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.webp")])
            else:
                f = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if f:
                v_path.set(f)

        ttk.Button(row_frame, text="Wybierz...", command=browse).pack(side="left", padx=2)
        ttk.Button(row_frame, text="Usuń", width=5, command=lambda: self.ui_remove_attachment(mod_key, att_data)).pack(side="right", padx=2)
        ttk.Button(row_frame, text="▼", width=3, command=lambda: self.ui_move_att_down(mod_key, att_data)).pack(side="right", padx=1)
        ttk.Button(row_frame, text="▲", width=3, command=lambda: self.ui_move_att_up(mod_key, att_data)).pack(side="right", padx=1)

    def ui_remove_attachment(self, mod_key, att_data):
        att_data['frame'].destroy()
        self.mod_attachments[mod_key].remove(att_data)

    def ui_move_att_up(self, mod_key, att_data):
        lst = self.mod_attachments[mod_key]
        idx = lst.index(att_data)
        if idx > 0:
            lst[idx - 1], lst[idx] = lst[idx], lst[idx - 1]
            self.redraw_attachments(mod_key)

    def ui_move_att_down(self, mod_key, att_data):
        lst = self.mod_attachments[mod_key]
        idx = lst.index(att_data)
        if idx < len(lst) - 1:
            lst[idx + 1], lst[idx] = lst[idx], lst[idx + 1]
            self.redraw_attachments(mod_key)

    def redraw_attachments(self, mod_key):
        for att in self.mod_attachments[mod_key]:
            att['frame'].pack_forget()
        for att in self.mod_attachments[mod_key]:
            att['frame'].pack(fill="x", pady=2)

    def ui_add_layer_row(self, mod_key, name="Nowa warstwa", name_en="", desc_pl="", desc_en="", t="Rw / Ctr", rw=0, ctr=0, spec=None, d=0):
        parent = self.layer_ui_frames[mod_key]
        row_frame = ttk.Frame(parent, relief="groove", borderwidth=1)
        row_frame.pack(fill="x", pady=2, ipadx=5, ipady=5)
        
        v_name, v_name_en, v_type, v_d = tk.StringVar(value=name), tk.StringVar(value=name_en), tk.StringVar(value=t), tk.DoubleVar(value=d)
        v_desc_pl, v_desc_en = tk.StringVar(value=desc_pl), tk.StringVar(value=desc_en)
        v_rw, v_ctr = tk.DoubleVar(value=rw), tk.DoubleVar(value=ctr)
        v_spec = [tk.DoubleVar(value=spec[i] if spec else 0.0) for i in range(8)]
        v_mat_choice = tk.StringVar(value="[Własna warstwa]")
        
        layer_data = {
            'name': v_name, 'name_en': v_name_en, 'desc_pl': v_desc_pl, 'desc_en': v_desc_en,
            'type': v_type, 'd': v_d, 'rw': v_rw, 'ctr': v_ctr, 'spec': v_spec, 'frame': row_frame
        }

        top_db = ttk.Frame(row_frame)
        top_db.pack(fill="x", pady=(2, 3))
        
        ttk.Button(top_db, text="Usuń", width=6, command=lambda: self.ui_remove_layer(mod_key, layer_data)).pack(side="right", padx=5)
        ttk.Button(top_db, text="▼", width=3, command=lambda: self.ui_move_layer_down(mod_key, layer_data)).pack(side="right", padx=2)
        ttk.Button(top_db, text="▲", width=3, command=lambda: self.ui_move_layer_up(mod_key, layer_data)).pack(side="right", padx=2)
        
        ttk.Label(top_db, text="Wybierz z bazy (auto-wypełnienie):").pack(side="left", padx=2)
        cb_mat = ttk.Combobox(top_db, textvariable=v_mat_choice, values=["[Własna warstwa]"] + sorted(list(self.materials_db.keys())), state="readonly", width=45)
        cb_mat.pack(side="left", padx=5)
        layer_data['cb_mat'] = cb_mat
        
        self.mod_layers[mod_key].append(layer_data)

        mid1 = ttk.Frame(row_frame)
        mid1.pack(fill="x", pady=2)
        ttk.Label(mid1, text="Nazwa PL:").pack(side="left", padx=2)
        ttk.Entry(mid1, textvariable=v_name, width=22).pack(side="left", padx=2)
        ttk.Label(mid1, text="Nazwa EN:").pack(side="left", padx=(5,2))
        ttk.Entry(mid1, textvariable=v_name_en, width=22).pack(side="left", padx=2)
        ttk.Label(mid1, text="Grub./Dł. [mm]:").pack(side="left", padx=(5,2))
        ttk.Entry(mid1, textvariable=v_d, width=6).pack(side="left", padx=2)
        ttk.Label(mid1, text="Typ:").pack(side="left", padx=(5,2))
        cb_type = ttk.Combobox(mid1, textvariable=v_type, values=["Rw / Ctr", "Widmo pasmowe"], state="readonly", width=14)
        cb_type.pack(side="left", padx=2)

        mid2 = ttk.Frame(row_frame)
        mid2.pack(fill="x", pady=2)
        ttk.Label(mid2, text="Opis PL:").pack(side="left", padx=2)
        ttk.Entry(mid2, textvariable=v_desc_pl, width=32).pack(side="left", padx=2)
        ttk.Label(mid2, text="Opis EN:").pack(side="left", padx=(5,2))
        ttk.Entry(mid2, textvariable=v_desc_en, width=32).pack(side="left", padx=2)

        bot = ttk.Frame(row_frame)
        bot.pack(fill="x", pady=2)
        frame_rw = ttk.Frame(bot)
        frame_spec = ttk.Frame(bot)
        
        ttk.Label(frame_rw, text="Rw [dB]:").pack(side="left", padx=2)
        ttk.Entry(frame_rw, textvariable=v_rw, width=8).pack(side="left", padx=2)
        ttk.Label(frame_rw, text="Ctr [dB]:").pack(side="left", padx=10)
        ttk.Entry(frame_rw, textvariable=v_ctr, width=8).pack(side="left", padx=2)
        
        for i, f in enumerate(self.freqs):
            ttk.Label(frame_spec, text=f, font=("", 8)).pack(side="left", padx=2)
            ttk.Entry(frame_spec, textvariable=v_spec[i], width=5).pack(side="left", padx=1)
            
        def update_bot_frame(*args):
            if v_type.get() == "Rw / Ctr":
                frame_spec.pack_forget()
                frame_rw.pack(side="left", padx=10)
            else:
                frame_rw.pack_forget()
                frame_spec.pack(side="left", padx=10)
                
        cb_type.bind("<<ComboboxSelected>>", update_bot_frame)
        update_bot_frame()
        
        def on_mat_change(*args):
            sel = v_mat_choice.get()
            if sel in self.materials_db:
                mat = self.materials_db[sel]
                v_name.set(sel)
                v_name_en.set(mat.get('name_en', ''))
                v_desc_pl.set(mat.get('desc_pl', ''))
                v_desc_en.set(mat.get('desc_en', ''))
                v_type.set(mat['type'])
                v_d.set(mat['d'])
                if mat['type'] == "Rw / Ctr":
                    v_rw.set(mat.get('rw', 0))
                    v_ctr.set(mat.get('ctr', 0))
                else:
                    for i in range(8): v_spec[i].set(mat.get('spec', [0]*8)[i])
                update_bot_frame()

        cb_mat.bind("<<ComboboxSelected>>", on_mat_change)
        v_d.trace_add("write", lambda *a: self.calc_module_depth(mod_key))
        self.calc_module_depth(mod_key)

    def ui_remove_layer(self, mod_key, layer_data):
        layer_data['frame'].destroy()
        self.mod_layers[mod_key].remove(layer_data)
        self.calc_module_depth(mod_key)

    def ui_move_layer_up(self, mod_key, layer_data):
        lst = self.mod_layers[mod_key]
        idx = lst.index(layer_data)
        if idx > 0:
            lst[idx - 1], lst[idx] = lst[idx], lst[idx - 1]
            self.redraw_layers(mod_key)

    def ui_move_layer_down(self, mod_key, layer_data):
        lst = self.mod_layers[mod_key]
        idx = lst.index(layer_data)
        if idx < len(lst) - 1:
            lst[idx + 1], lst[idx] = lst[idx], lst[idx + 1]
            self.redraw_layers(mod_key)

    def redraw_layers(self, mod_key):
        for layer in self.mod_layers[mod_key]:
            layer['frame'].pack_forget()
        for layer in self.mod_layers[mod_key]:
            layer['frame'].pack(fill="x", pady=2, ipadx=5, ipady=5)

    def calc_module_depth(self, mod_key):
        try:
            total_d = sum(l['d'].get() for l in self.mod_layers[mod_key])
            self.mod_dims[mod_key]['d'].set(round(total_d, 2))
        except tk.TclError: pass

    def toggle_tab_state(self, key):
        state = 'normal' if self.aktywne[key].get() else 'disabled'
        for entry in self.modules_config[key].get('ui_entries', []):
            if isinstance(entry, ttk.Entry) and entry.cget('state') == 'readonly': continue
            entry.config(state=state)
        self.recalc_wall_area()

    # --- ZAPIS I ODCZYT JSON ---
    def get_state_dict(self):
        saved_layers = {}
        for k, layers in self.mod_layers.items():
            saved_layers[k] = []
            for l in layers:
                saved_layers[k].append({
                    'name': l['name'].get(), 'name_en': l['name_en'].get(),
                    'desc_pl': l['desc_pl'].get(), 'desc_en': l['desc_en'].get(),
                    'type': l['type'].get(), 'd': l['d'].get(),
                    'rw': l['rw'].get(), 'ctr': l['ctr'].get(), 'spec': [v.get() for v in l['spec']]
                })
                
        saved_attachments = {}
        for k, atts in self.mod_attachments.items():
            saved_attachments[k] = [{'type': a['type'].get(), 'path': a['path'].get()} for a in atts]
        
        return {
            "meta": {
                "project_no": self.project_no.get(), "client_name": self.client_name.get(),
                "po_number": self.po_number.get(), "client_doc_no": self.client_doc_no.get(),
                "project_name": self.project_name.get(), "silnik_typ": self.silnik_typ.get(),
                "agregat_typ": self.agregat_typ.get(), "typ_obudowy": self.typ_obudowy.get(),
                "jezyk_raportu": self.jezyk_raportu.get(), "logo_path": self.logo_path.get(),
                "footer_text": self.footer_text.get(),
                "engine_wiz_path": self.engine_wiz_path.get(),
                "engine_pdf_path": self.engine_pdf_path.get()
            },
            "revisions": [
                {"rev": r['rev'].get(), "desc": r['desc'].get(), "drawn": r['drawn'].get(), "check": r['check'].get(), "date": r['date'].get()} 
                for r in self.rev_data
            ],
            "dimensions": {
                "kont_l": self.kont_l.get(), "kont_w": self.kont_w.get(), "kont_h": self.kont_h.get()
            },
            "points": [{"name": p['name'].get(), "d": p['d'].get(), "limit": p['limit'].get()} for p in self.points_data],
            "sabine": {
                "alfa_podl": self.alfa_podl.get(),
                "alfa_scian": self.alfa_scian.get(),
                "sabine_q": self.sabine_q.get(),
                "sabine_r": self.sabine_r.get()
            },
            "radiator": {
                "active": self.rad_active.get(),
                "name": self.rad_name.get(),
                "lw": [v.get() for v in self.rad_lw_vars],
                "att": [v.get() for v in self.rad_att_vars],
                "wiz_path": self.rad_wiz_path.get(),
                "pdf_path": self.rad_pdf_path.get()
            },
            "modules_active": {k: v.get() for k, v in self.aktywne.items()},
            "mod_dims": {
                k: {
                    'qty': v['qty'].get(), 'w': v['w'].get(), 'h': v['h'].get(), 'd': v['d'].get(), 's': v['s'].get(),
                    'kierunek': v.get('kierunek', tk.StringVar(value="Poziomy")).get(),
                    'wysokosc': v.get('wysokosc', tk.DoubleVar(value=0.0)).get()
                } 
                for k, v in self.mod_dims.items()
            },
            "layers": saved_layers,
            "mod_attachments": saved_attachments,
            "acoustics": {
                "lw_vars": [v.get() for v in self.lw_vars], 
                "room_abs_vars": [v.get() for v in self.room_abs_vars],
                "exhaust_lw_vars": [v.get() for v in self.exhaust_lw_vars]
            }
        }

    def get_state_json(self):
        return json.dumps(self.get_state_dict(), sort_keys=True)

    def nowy_projekt(self):
        if messagebox.askyesno("Nowy", "Czy wyczyścić wszystkie dane i rozpocząć nowy projekt?"):
            self.current_project_dir = ""
            self.project_no.set("PRJ-2026-001")
            self.client_name.set("Nazwa Klienta")
            self.po_number.set("PO-123456")
            self.client_doc_no.set("DOC-001-REV0")
            self.project_name.set("Raport Akustyczny Obudowy")
            self.silnik_typ.set("")
            self.agregat_typ.set("")
            self.typ_obudowy.set("Obudowa kontenerowa")
            self.jezyk_raportu.set("Polski")
            self.logo_path.set("")
            self.footer_text.set("Diamel | diamel.pl")

            self.kont_l.set(12.0)
            self.kont_w.set(3.0)
            self.kont_h.set(3.0)

            for p in list(self.points_data): self.ui_remove_point(p)
            self.ui_add_point_row(name="Standardowy", d=1.0, limit=85.0)

            for i in range(8):
                self.lw_vars[i].set(95.0)
                self.room_abs_vars[i].set(-3.0)
                self.rad_lw_vars[i].set(0.0)
                self.rad_att_vars[i].set(0.0)
                self.exhaust_lw_vars[i].set(0.0)

            self.rad_active.set(False)
            self.rad_name.set("")
            self.rad_wiz_path.set("")
            self.rad_pdf_path.set("")
            self.engine_wiz_path.set("")
            self.engine_pdf_path.set("")

            self.alfa_podl.set(0.02)
            self.alfa_scian.set(0.85)
            self.sabine_q.set(2.0)
            self.sabine_r.set(1.5)

            for k in self.mod_layers:
                for l in list(self.mod_layers[k]): self.ui_remove_layer(k, l)
            
            for k in self.mod_attachments:
                for a in list(self.mod_attachments[k]): self.ui_remove_attachment(k, a)

            for r in list(self.rev_data): self.ui_remove_rev_row(r)
            self.ui_add_rev_row(rev="1", desc="Wydanie 1 / Initial Issue", drawn="MK", check="Appr.")

            for k in self.modules_config:
                self.aktywne[k].set(self.modules_config[k]['req'] or k in ['drzwi'])
                self.mod_dims[k]['qty'].set(1)
                self.mod_dims[k]['w'].set(1.0)
                self.mod_dims[k]['h'].set(2.0)
                self.mod_dims[k]['d'].set(0.0)
                self.mod_dims[k]['s'].set(0.0)
                if k in ['wydech', 'czerpnia', 'wyrzut']:
                    self.mod_dims[k]['kierunek'].set("Poziomy")
                    self.mod_dims[k]['wysokosc'].set(0.0)
                self.toggle_tab_state(k)
                
            self.toggle_radiator_state()
            self.calc_kontener_surfaces()
            self.last_saved_state = self.get_state_json()

    def zapisz_projekt(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            self.current_project_dir = os.path.dirname(file_path)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.get_state_dict(), f, indent=4, ensure_ascii=False)
            self.last_saved_state = self.get_state_json()
            messagebox.showinfo("Zapisano", "Projekt zapisany!")

    def wczytaj_projekt(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                self.current_project_dir = os.path.dirname(file_path)
                with open(file_path, 'r', encoding='utf-8') as f: state = json.load(f)
                
                meta = state.get("meta", {})
                self.project_no.set(meta.get("project_no", ""))
                self.client_name.set(meta.get("client_name", ""))
                self.po_number.set(meta.get("po_number", ""))
                self.client_doc_no.set(meta.get("client_doc_no", ""))
                self.project_name.set(meta.get("project_name", "Wczytany_Projekt"))
                self.silnik_typ.set(meta.get("silnik_typ", ""))
                self.agregat_typ.set(meta.get("agregat_typ", ""))
                self.typ_obudowy.set(meta.get("typ_obudowy", "Obudowa kontenerowa"))
                self.jezyk_raportu.set(meta.get("jezyk_raportu", "Polski"))
                self.logo_path.set(meta.get("logo_path", ""))
                self.footer_text.set(meta.get("footer_text", ""))
                self.engine_wiz_path.set(meta.get("engine_wiz_path", ""))
                self.engine_pdf_path.set(meta.get("engine_pdf_path", ""))
                
                sab = state.get("sabine", {})
                self.alfa_podl.set(sab.get("alfa_podl", 0.02))
                self.alfa_scian.set(sab.get("alfa_scian", 0.85))
                self.sabine_q.set(sab.get("sabine_q", 2.0))
                self.sabine_r.set(sab.get("sabine_r", 1.5))
                
                rad = state.get("radiator", {})
                self.rad_active.set(rad.get("active", False))
                self.rad_name.set(rad.get("name", ""))
                self.rad_wiz_path.set(rad.get("wiz_path", ""))
                self.rad_pdf_path.set(rad.get("pdf_path", ""))
                for i in range(8):
                    self.rad_lw_vars[i].set(rad.get("lw", [0.0]*8)[i])
                    self.rad_att_vars[i].set(rad.get("att", [0.0]*8)[i])
                self.toggle_radiator_state()

                for r in list(self.rev_data): self.ui_remove_rev_row(r)
                revs = state.get("revisions", [])
                for rev in revs:
                    self.ui_add_rev_row(rev=rev.get('rev',''), desc=rev.get('desc',''), drawn=rev.get('drawn',''), check=rev.get('check',''), date=rev.get('date',''))
                if not revs: self.ui_add_rev_row(rev="1", desc="Initial Issue", drawn="MK", check="Appr.")

                dim = state.get("dimensions", {})
                self.kont_l.set(dim.get("kont_l", 12.0))
                self.kont_w.set(dim.get("kont_w", 3.0))
                self.kont_h.set(dim.get("kont_h", 3.0))
                
                for p in list(self.points_data): self.ui_remove_point(p)
                pts = state.get("points", [])
                if pts:
                    for p in pts: self.ui_add_point_row(p['name'], p['d'], p['limit'])
                else:
                    self.ui_add_point_row("Standardowy", 1.0, 85.0)

                mod_act = state.get("modules_active", {})
                mod_dims = state.get("mod_dims", {})
                
                for k in self.mod_layers:
                    for l in list(self.mod_layers[k]): self.ui_remove_layer(k, l)

                layers_data = state.get("layers", {})
                for k, lst in layers_data.items():
                    for l_data in lst:
                        self.ui_add_layer_row(
                            k, name=l_data.get('name',''), name_en=l_data.get('name_en',''),
                            desc_pl=l_data.get('desc_pl',''), desc_en=l_data.get('desc_en',''),
                            t=l_data.get('type','Rw / Ctr'), rw=l_data.get('rw',0), ctr=l_data.get('ctr',0), 
                            spec=l_data.get('spec'), d=l_data.get('d',0)
                        )

                for k in self.mod_attachments:
                    for a in list(self.mod_attachments[k]): self.ui_remove_attachment(k, a)
                
                atts_data = state.get("mod_attachments", {})
                for k, lst in atts_data.items():
                    for a_data in lst:
                        self.ui_add_attachment_row(k, type_val=a_data.get('type', 'Obraz (PNG/JPG)'), path_val=a_data.get('path', ''))

                for k in self.aktywne.keys():
                    if k in mod_act:
                        self.aktywne[k].set(mod_act[k])
                        self.toggle_tab_state(k)
                    if k in mod_dims:
                        if not self.modules_config[k]['auto_dim']:
                            self.mod_dims[k]['qty'].set(mod_dims[k].get('qty', 1))
                            self.mod_dims[k]['w'].set(mod_dims[k].get('w', 0.0))
                            self.mod_dims[k]['h'].set(mod_dims[k].get('h', 0.0))
                        self.mod_dims[k]['d'].set(mod_dims[k].get('d', 0.0))
                        self.mod_dims[k]['s'].set(mod_dims[k].get('s', 0.0))
                        if k in ['wydech', 'czerpnia', 'wyrzut']:
                            self.mod_dims[k]['kierunek'].set(mod_dims[k].get('kierunek', 'Poziomy'))
                            self.mod_dims[k]['wysokosc'].set(mod_dims[k].get('wysokosc', 0.0))

                aco = state.get("acoustics", {})
                for i in range(8):
                    self.lw_vars[i].set(aco.get("lw_vars", [95.0]*8)[i])
                    self.room_abs_vars[i].set(aco.get("room_abs_vars", [-3.0]*8)[i])
                    self.exhaust_lw_vars[i].set(aco.get("exhaust_lw_vars", [0.0]*8)[i])

                self.calc_kontener_surfaces() 
                self.last_saved_state = self.get_state_json()
                messagebox.showinfo("Wczytano", "Projekt wczytany pomyślnie.")
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się wczytać pliku:\n{e}")

    # --- MATEMATYKA ---
    def wykonaj_obliczenia(self, t=None):
        if not t: t = self.translations.get(self.jezyk_raportu.get(), self.translations['Polski'])
        is_eng = (self.jezyk_raportu.get() == 'English')
        
        try:
            lw_a_wazone = [v.get() for v in self.lw_vars]
            room_abs = [v.get() for v in self.room_abs_vars]
            lp_a_wazone = [lw_a_wazone[i] + room_abs[i] for i in range(8)]
            
            calkowite_lw_a = 10 * math.log10(sum([10 ** (val / 10) for val in lw_a_wazone]))
            calkowite_lp_wewn = 10 * math.log10(sum([10 ** (val / 10) for val in lp_a_wazone]))

            rad_is_active = self.rad_active.get()

            # 1. Obliczenie zjawisk niezależnych od dystansu (tłumienia modułów)
            modules_base_data = {}
            s_calkowita = 0.0

            for key in self.aktywne.keys():
                s = self.mod_dims[key]['s'].get()
                qty = self.mod_dims[key]['qty'].get()
                
                if key == 'wydech' and self.aktywne[key].get() and qty > 0:
                    s = 0.0 
                
                if self.aktywne[key].get() and (s > 0 or key == 'wydech'):
                    sum_10_R = [0.0] * 8
                    layer_infos = []
                    
                    for layer in self.mod_layers[key]:
                        l_name_pl = layer['name'].get()
                        l_name_en = layer['name_en'].get()
                        l_name = l_name_en if (is_eng and l_name_en) else l_name_pl
                        
                        l_desc_pl = layer['desc_pl'].get()
                        l_desc_en = layer['desc_en'].get()
                        l_desc = l_desc_en if (is_eng and l_desc_en) else l_desc_pl
                        
                        l_type, l_d = layer['type'].get(), layer['d'].get()
                        
                        desc_str = f" - {l_desc}" if l_desc else ""
                        if l_type == "Rw / Ctr":
                            rw, ctr = layer['rw'].get(), layer['ctr'].get()
                            for i in range(8): sum_10_R[i] += 10**((rw + ctr)/10)
                            layer_infos.append(f"{l_name}{desc_str} (Rw: {rw}dB, Ctr: {ctr}dB, {t['depth']}: {l_d}mm)")
                        else:
                            spec = [v.get() for v in layer['spec']]
                            for i in range(8): sum_10_R[i] += 10**(spec[i]/10)
                            layer_infos.append(f"{l_name}{desc_str} ({t['spec_att']}, {t['depth']}: {l_d}mm)")

                    mod_attenuation = [0.0] * 8
                    for i in range(8):
                        if sum_10_R[i] > 0:
                            mod_attenuation[i] = 10 * math.log10(sum_10_R[i])
                            
                    modules_base_data[key] = {
                        's': s, 'qty': qty, 'mod_attenuation': mod_attenuation, 'layer_infos': layer_infos
                    }
                    s_calkowita += s

            wyniki_dystans = {}

            # 2. Pętla przeliczająca dla każdego zadanego dystansu
            l, w, h = self.kont_l.get(), self.kont_w.get(), self.kont_h.get()

            for pt in self.points_data:
                d = pt['d'].get()
                limit = pt['limit'].get()
                name = pt['name'].get()
                
                s_kont_val = 2 * (l * h + w * h) + l * w
                l_pom, w_pom, h_pom = l + 2*d, w + 2*d, h + d
                s_pom_val = 2 * (l_pom * h_pom + w_pom * h_pom) + l_pom * w_pom
                korekta_dystansu = 10 * math.log10(s_pom_val / s_kont_val) if s_kont_val > 0 and s_pom_val > 0 else 0.0

                sum_W_emisja_d = 0.0
                sum_W_sur_d = 0.0
                results_d = {}

                for key, b_data in modules_base_data.items():
                    s = b_data['s']
                    qty = b_data['qty']
                    mod_attenuation = b_data['mod_attenuation']
                    layer_infos = list(b_data['layer_infos'])

                    halas_pow_log = 0.0
                    wynik_dystans = 0.0
                    W_total_module = 0.0
                    W_total_module_emisja = 0.0

                    # Specjalne logiki i korekty geometrii (kierunek, odległość Pitagorasa)
                    korekta_kierunku = 0.0
                    korekta_wysokosci = 0.0
                    if key in ['wydech', 'czerpnia', 'wyrzut']:
                        kierunek = self.mod_dims[key]['kierunek'].get()
                        wysokosc = self.mod_dims[key]['wysokosc'].get()
                        if kierunek == "Pionowy w górę":
                            korekta_kierunku = -3.0
                        if wysokosc > 0:
                            r_base = max(d, 1.0)
                            r_actual = math.sqrt(r_base**2 + wysokosc**2)
                            korekta_wysokosci = -20 * math.log10(r_actual / r_base) if r_actual > r_base else 0.0
                            
                    korekta_dodatkowa = korekta_kierunku + korekta_wysokosci

                    if key == 'wyrzut' and rad_is_active:
                        rad_lw_vals = [v.get() for v in self.rad_lw_vars]
                        rad_att_vals = [v.get() for v in self.rad_att_vars]

                        for i in range(8):
                            Lp_eng_out = lp_a_wazone[i] - rad_att_vals[i] - mod_attenuation[i]
                            W_eng = s * (10 ** (Lp_eng_out / 10))

                            LwA_rad = rad_lw_vals[i] + self.k_dba_default[i]
                            LwA_rad_out = LwA_rad - mod_attenuation[i]
                            W_rad = 10 ** (LwA_rad_out / 10)

                            W_band = W_eng + W_rad
                            W_total_module += W_band
                            W_total_module_emisja += W_band * (10 ** (korekta_dodatkowa / 10))

                        halas_pow_log = 10 * math.log10(W_total_module / s) if s > 0 and W_total_module > 0 else 0
                        
                        if "[CHŁODNICA AKTYWNA]" not in str(layer_infos):
                            layer_infos.insert(0, f"[CHŁODNICA AKTYWNA] {self.rad_name.get()} (Własne Lw + Tłumienie rdzenia)")
                        if korekta_dodatkowa != 0.0:
                            layer_infos.insert(1, f"Fizyczna korekta kierunku/odległości (Wyrzutnia dachowa/nadstawka): {korekta_dodatkowa:.1f} dB")
                        
                        wynik_dystans = (10 * math.log10(W_total_module_emisja / s) - korekta_dystansu) if s > 0 and W_total_module_emisja > 0 else 0
                        
                        sum_W_emisja_d += W_total_module_emisja
                        sum_W_sur_d += W_total_module
                        
                    elif key == 'wydech':
                        exh_lw_vals = [v.get() for v in self.exhaust_lw_vars]
                        
                        LwA_out_bands = []
                        for i in range(8):
                            LwA_exh = exh_lw_vals[i] + self.k_dba_default[i]
                            LwA_out = LwA_exh - mod_attenuation[i]
                            
                            W_band = qty * (10 ** (LwA_out / 10))
                            LwA_out_bands.append(W_band)
                            W_total_module += W_band
                            W_total_module_emisja += W_band * (10 ** (korekta_dodatkowa / 10))

                        halas_pow_log = 10 * math.log10(sum(LwA_out_bands)) if sum(LwA_out_bands) > 0 else 0
                        
                        if s_pom_val > 0 and W_total_module_emisja > 0:
                            wynik_dystans = 10 * math.log10(W_total_module_emisja / s_pom_val)
                        else:
                            wynik_dystans = 0
                            
                        if "[ŹRÓDŁO ZEWNĘTRZNE]" not in str(layer_infos):
                            layer_infos.insert(0, f"[ŹRÓDŁO ZEWNĘTRZNE] Niezależne widmo Lw wydechu silnika")
                            if korekta_dodatkowa != 0.0:
                                layer_infos.insert(1, f"Fizyczna korekta kierunku/odległości komina: {korekta_dodatkowa:.1f} dB")
                        
                        sum_W_emisja_d += W_total_module_emisja

                    else:
                        for i in range(8):
                            Lp_out = lp_a_wazone[i] - mod_attenuation[i]
                            W_band = s * (10 ** (Lp_out / 10))
                            W_total_module += W_band
                            W_total_module_emisja += W_band * (10 ** (korekta_dodatkowa / 10))

                        halas_pow_log = 10 * math.log10(W_total_module / s) if s > 0 and W_total_module > 0 else 0
                        wynik_dystans = (10 * math.log10(W_total_module_emisja / s) - korekta_dystansu) if s > 0 and W_total_module_emisja > 0 else 0
                        
                        if korekta_dodatkowa != 0.0 and key in ['czerpnia', 'wyrzut']:
                            if not any("Fizyczna korekta" in info for info in layer_infos):
                                layer_infos.insert(0, f"Fizyczna korekta kierunku/odległości (Nadstawka): {korekta_dodatkowa:.1f} dB")

                        sum_W_emisja_d += W_total_module_emisja
                        sum_W_sur_d += W_total_module

                    results_d[key] = {
                        's': s, 'd': self.mod_dims[key]['d'].get(), 'qty': qty,
                        'halas_pow': halas_pow_log, 'wynik_dystans': wynik_dystans, 'layer_infos': layer_infos,
                        'is_exhaust': True if key == 'wydech' else False
                    }

                halas_wypadkowy_pow_d = 10 * math.log10(sum_W_sur_d / s_calkowita) if s_calkowita > 0 and sum_W_sur_d > 0 else 0
                wynik_dystans_wypadkowy_d = 10 * math.log10(sum_W_emisja_d / s_pom_val) if s_pom_val > 0 and sum_W_emisja_d > 0 else 0

                wyniki_dystans[d] = {
                    'name': name,
                    'd': d,
                    'limit': limit,
                    'korekta_dystansu': korekta_dystansu,
                    'moduly': results_d,
                    'halas_wypadkowy_pow': halas_wypadkowy_pow_d,
                    'wynik_dystans_wypadkowy': wynik_dystans_wypadkowy_d
                }

            return {
                "lw_a_wazone": lw_a_wazone, "room_abs": room_abs, "lp_a_wazone": lp_a_wazone,
                "calkowite_lw_a": calkowite_lw_a, "calkowite_lp_wewn": calkowite_lp_wewn,
                "s_calkowita": s_calkowita, 
                "wyniki_dystans": wyniki_dystans
            }
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Błąd Obliczeń", f"Sprawdź wpisane dane! Wystąpił błąd:\n{e}")
            return None

    def pokaz_obliczenia(self):
        t = self.translations.get(self.jezyk_raportu.get(), self.translations['Polski'])
        dane = self.wykonaj_obliczenia(t)
        if not dane: return

        lines = [
            "--- WSTĘPNE WYNIKI OBLICZEŃ (METODA PASMOWA) ---\n",
            f"LwA silnika: {dane['calkowite_lw_a']:.1f} dB(A)",
            f"LpA wewnątrz: {dane['calkowite_lp_wewn']:.1f} dB(A)"
        ]

        for d, res_d in dane['wyniki_dystans'].items():
            lines.append(f"\n======================================")
            lines.append(f"WYNIKI DLA PUNKTU: {res_d['name']} ({d} m)")
            lines.append(f"Korekta obwiedni pomiarowej: {res_d['korekta_dystansu']:.1f} dB\n")
            lines.append("Indywidualny udział modułów:")
            
            for key, res_mod in res_d['moduly'].items(): 
                mod_name = t['mod_names'].get(key, key)
                lines.append(f" - {mod_name}:")
                if res_mod.get('is_exhaust'):
                    lines.append(f"     Emitowane LwA wydechu: {res_mod['halas_pow']:.1f} dB(A)")
                else:
                    lines.append(f"     Tuż przy obudowie (LpA): {res_mod['halas_pow']:.1f} dB(A)")
                lines.append(f"     W podanej odległości:        {res_mod['wynik_dystans']:.1f} dB(A)")

            wynik = res_d['wynik_dystans_wypadkowy']
            limit = res_d['limit']
            lines.append(f"\n---")
            lines.append(f"WYNIK CAŁKOWITY ({d}m): {wynik:.1f} dB(A) (LIMIT: {limit:.1f})")
            lines.append("✅ PASS" if wynik <= limit else "❌ FAIL")

        top = tk.Toplevel(self.root)
        top.title("Podgląd Obliczeń")
        top.geometry("500x650")
        txt = tk.Text(top, wrap="word", font=("Consolas", 10))
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", "\n".join(lines))
        txt.config(state="disabled")
        ttk.Button(top, text="OK", command=top.destroy).pack(pady=10)

    def get_report_filepath(self, ext=".docx"):
        proj_no = self.project_no.get().strip().replace(" ", "_")
        genset = self.agregat_typ.get().strip().replace(" ", "_")
        if not genset: genset = "BrakAgregatu"

        rev_no = "1"
        rev_date = datetime.now().strftime('%Y-%m-%d')
        if self.rev_data:
            last_rev = self.rev_data[-1]
            rev_no = last_rev['rev'].get().strip()
            rev_date = last_rev['date'].get().strip().replace(".", "-")

        filename = f"Raport_{proj_no}_{genset}_Rev{rev_no}_{rev_date}{ext}"
        filename = re.sub(r'[\\/*?:"<>|]', "", filename)

        if self.current_project_dir and os.path.exists(self.current_project_dir):
            return os.path.join(self.current_project_dir, filename), False
        else:
            filetypes = [("Word Document", "*.docx")] if ext == ".docx" else [("PDF Document", "*.pdf")]
            path = filedialog.asksaveasfilename(
                defaultextension=ext,
                initialfile=filename,
                filetypes=filetypes
            )
            return path, True

    def insert_safe_picture(self, parent_obj, img_path, width):
        try:
            parent_obj.add_picture(img_path, width=width)
        except Exception as e:
            print(f"Direct picture insert failed for {img_path}: {e}")
            if PIL_AVAILABLE:
                try:
                    with PILImage.open(img_path) as img:
                        b = io.BytesIO()
                        if img.mode not in ('RGB', 'RGBA'): img = img.convert('RGBA')
                        img.save(b, format='PNG')
                        b.seek(0)
                        parent_obj.add_picture(b, width=width)
                except Exception as e2: 
                    print(f"PIL fallback failed: {e2}")

    # --- GENERATOR WORD ---
    def generuj_raport(self, mode='docx'):
        if not self.points_data:
            messagebox.showwarning("Brak punktów", "Dodaj najpierw punkt pomiarowy w zakładce 'Punkty Pomiarowe'!")
            return
            
        lang = self.jezyk_raportu.get()
        t = self.translations.get(lang, self.translations['Polski'])
        
        dane = self.wykonaj_obliczenia(t)
        if not dane: return

        output_docx, user_asked = self.get_report_filepath(ext=".docx")
        if not output_docx: return

        try:
            doc = Document()
            doc.core_properties.author = "Mateusz Kujawa"
            doc.core_properties.title = t['title'].replace("\n", " - ")

            for section in doc.sections:
                section.top_margin, section.bottom_margin = Inches(0.8), Inches(0.8)
                section.left_margin, section.right_margin = Inches(1), Inches(1)
                
                if self.logo_path.get() and os.path.exists(self.logo_path.get()):
                    header = section.header
                    h_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                    h_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    h_para.text = "" 
                    run_logo = h_para.add_run()
                    self.insert_safe_picture(run_logo, self.logo_path.get(), Inches(1.8))
                    
                if self.footer_text.get():
                    footer = section.footer
                    f_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                    f_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    f_para.text = ""
                    f_para.add_run(self.footer_text.get()).font.size = Pt(9)

            # STRONA 1: TYTUŁ I METADANE
            title = doc.add_heading(t['title'], 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("\n")

            table_meta = doc.add_table(rows=6, cols=2)
            table_meta.style = 'Light Grid Accent 1'
            meta_data = [
                (t['meta'][0], self.project_no.get()), (t['meta'][1], self.project_name.get()),
                (t['meta'][2], self.client_name.get()), (t['meta'][3], self.po_number.get()),
                (t['meta'][4], self.client_doc_no.get()), (t['meta'][5], datetime.now().strftime('%Y-%m-%d'))
            ]
            for i, (label, val) in enumerate(meta_data):
                table_meta.rows[i].cells[0].text = label
                table_meta.rows[i].cells[0].paragraphs[0].runs[0].bold = True
                table_meta.rows[i].cells[1].text = val

            doc.add_paragraph("\n")
            
            # TABELA REWIZJI
            if self.rev_data:
                table_rev = doc.add_table(rows=1 + len(self.rev_data), cols=5)
                table_rev.style = 'Light List Accent 1'
                for i, h in enumerate(t['rev_th']):
                    table_rev.rows[0].cells[i].text = h
                    table_rev.rows[0].cells[i].paragraphs[0].runs[0].bold = True
                    
                for i, rev in enumerate(self.rev_data):
                    row_cells = table_rev.rows[i+1].cells
                    row_cells[0].text = rev['rev'].get()
                    row_cells[1].text = rev['desc'].get()
                    row_cells[2].text = rev['drawn'].get()
                    row_cells[3].text = rev['check'].get()
                    row_cells[4].text = rev['date'].get()

            # STRONA 2: SEKCJE OPISOWE
            doc.add_page_break()
            doc.add_heading(t['sec1'], level=1)
            doc.add_paragraph(t['sec1_desc'])
            
            for d, res_d in dane['wyniki_dystans'].items():
                p_cel_pt = doc.add_paragraph(f"    - {res_d['name']}: ≤ {res_d['limit']:.1f} dB(A) w odl. {d} m")
                p_cel_pt.style = 'List Bullet 2'

            wybrany_typ = self.typ_obudowy.get()
            tlumaczony_typ = t['types'].get(wybrany_typ, wybrany_typ)
            
            doc.add_heading(t['sec2'], level=1)
            p_desc2 = doc.add_paragraph()
            p_desc2.add_run(f"{t['inst_type_label']} {tlumaczony_typ}\n").bold = True
            p_desc2.add_run(t['sec2_desc'].format(genset=self.agregat_typ.get(), engine=self.silnik_typ.get()))

            table1 = doc.add_table(rows=1, cols=3)
            table1.style = 'Light List Accent 1'
            h1 = table1.rows[0].cells
            h1[0].text, h1[1].text, h1[2].text = t['th1_1'], t['th1_2'], t['th1_3']
            for i in range(8):
                row = table1.add_row().cells
                row[0].text = self.freqs[i] + "Hz"
                row[1].text = f"{self.k_dba_default[i]:.1f}"
                row[2].text = f"{dane['lw_a_wazone'][i]:.1f}"

            p_lw = doc.add_paragraph(f"\n{t['calc_lwa']}")
            p_lw.add_run(f"{dane['calkowite_lw_a']:.1f} dB(A)").bold = True

            # Obrazek silnika w raporcie
            if self.engine_wiz_path.get() and os.path.exists(self.engine_wiz_path.get()):
                doc.add_paragraph()
                p_img_eng = doc.add_paragraph()
                p_img_eng.alignment = WD_ALIGN_PARAGRAPH.CENTER
                self.insert_safe_picture(p_img_eng.runs[0] if p_img_eng.runs else p_img_eng.add_run(), self.engine_wiz_path.get(), Inches(5.5))
                doc.add_paragraph()

            if self.rad_active.get():
                p_rad = doc.add_paragraph()
                p_rad.add_run(f"Chłodnica aktywna na wyrzucie: {self.rad_name.get()}").bold = True
                
                table_rad = doc.add_table(rows=1, cols=3)
                table_rad.style = 'Light List Accent 1'
                hr = table_rad.rows[0].cells
                hr[0].text, hr[1].text, hr[2].text = t['th1_1'], "Lw [dB]", "Tłumienie Rdzenia [dB]"
                rad_lws = [v.get() for v in self.rad_lw_vars]
                rad_atts = [v.get() for v in self.rad_att_vars]
                for i in range(8):
                    rowr = table_rad.add_row().cells
                    rowr[0].text = self.freqs[i] + "Hz"
                    rowr[1].text = f"{rad_lws[i]:.1f}"
                    rowr[2].text = f"{rad_atts[i]:.1f}"

                # Obrazek chłodnicy w raporcie
                if self.rad_wiz_path.get() and os.path.exists(self.rad_wiz_path.get()):
                    doc.add_paragraph()
                    p_img_rad = doc.add_paragraph()
                    p_img_rad.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    self.insert_safe_picture(p_img_rad.runs[0] if p_img_rad.runs else p_img_rad.add_run(), self.rad_wiz_path.get(), Inches(5.5))
                    doc.add_paragraph()

            doc.add_heading(t['sec3'], level=1)
            table2 = doc.add_table(rows=1, cols=4)
            table2.style = 'Light List Accent 1'
            h2 = table2.rows[0].cells
            h2[0].text, h2[1].text, h2[2].text, h2[3].text = t['th1_1'], t['th2_2'], t['th2_3'], t['th2_4']
            for i in range(8):
                row2 = table2.add_row().cells
                row2[0].text = self.freqs[i] + "Hz"
                row2[1].text = f"{dane['lw_a_wazone'][i]:.1f}"
                row2[2].text = f"{dane['room_abs'][i]:.1f}"
                row2[3].text = f"{dane['lp_a_wazone'][i]:.1f}"

            p_lp = doc.add_paragraph(f"\n{t['calc_lpa']}")
            p_lp.add_run(f"{dane['calkowite_lp_wewn']:.1f} dB(A)").bold = True
            doc.add_page_break()

            # STRONA 3: SPECYFIKACJA WARSTW (Z DYNAMICZNYCH DANYCH)
            doc.add_heading(t['sec4'], level=1)
            doc.add_paragraph(t['sec4_desc'])
            
            first_distance_results = list(dane['wyniki_dystans'].values())[0]['moduly']

            for key, res in first_distance_results.items():
                mod_name = t['mod_names'].get(key, key)
                p_mod = doc.add_paragraph()
                p_mod.add_run(f"• {mod_name} ").bold = True
                
                qty_str = f"{res['qty']} {t['qty_pcs']}, " if res['qty'] > 1 else ""
                if res.get('is_exhaust'):
                    p_mod.add_run(f"({qty_str}Tłumiki spalin komina)")
                else:
                    p_mod.add_run(f"({qty_str}S = {res['s']:.2f} m2, {t['depth']} = {res['d']:.1f} mm)")
                
                if res['layer_infos']:
                    for info in res['layer_infos']:
                        if "Fizyczna korekta kierunku" in info: continue
                        p_lay = doc.add_paragraph(f"    - {info}")
                        p_lay.style = 'List Bullet 2'

                # Wklejenie wielu obrazków wizualizacji modułu
                for att in self.mod_attachments.get(key, []):
                    if att['type'].get() == "Obraz (PNG/JPG)" and os.path.exists(att['path'].get()):
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        self.insert_safe_picture(p_img.runs[0] if p_img.runs else p_img.add_run(), att['path'].get(), Inches(5.5))
                        doc.add_paragraph()
            
            doc.add_paragraph(f"\n{t['tot_area']}: {dane['s_calkowita']:.2f} m2")

            # STRONA 4: PODSUMOWANIE WYNIKÓW
            doc.add_page_break()
            doc.add_heading(t['sec5'], level=1)
            
            for d, res_d in dane['wyniki_dystans'].items():
                doc.add_heading(f"Wyniki dla punktu: {res_d['name']} (odl. {d} m)", level=2)
                doc.add_paragraph(f"{t['dist_corr']} ({d}m): {res_d['korekta_dystansu']:.2f} dB\n")
                
                doc.add_paragraph(t['ind_noise'], style='Heading 3')
                for key, res_mod in res_d['moduly'].items(): 
                    mod_name = t['mod_names'].get(key, key)
                    doc.add_paragraph(f"- {mod_name}: {res_mod['wynik_dystans']:.1f} dB(A)")

                koncowy_wynik = res_d['wynik_dystans_wypadkowy']
                limit = res_d['limit']
                p_final = doc.add_paragraph(f"\n{t['tot_noise']} {d}m: ")
                p_final.add_run(f"{koncowy_wynik:.1f} dB(A)").bold = True
                p_final.runs[0].font.size = Pt(14)

                conclusion = doc.add_paragraph(f"{t['proj_stat']}")
                if koncowy_wynik <= limit:
                    run_status = conclusion.add_run(t['pass'])
                    run_status.font.color.rgb = RGBColor(0, 150, 0)
                else:
                    run_status = conclusion.add_run(t['fail'])
                    run_status.font.color.rgb = RGBColor(200, 0, 0)
                run_status.bold = True

            # Zapis i PDF Merge
            saved_docx = False
            while not saved_docx:
                try:
                    doc.save(output_docx)
                    saved_docx = True
                except PermissionError:
                    if not user_asked:
                        choice = messagebox.askyesnocancel(
                            t['err_in_use_title'],
                            t['err_in_use_msg'].format(output_docx)
                        )
                        if choice is True:  
                            continue
                        elif choice is False:  
                            new_path = filedialog.asksaveasfilename(
                                defaultextension=".docx", 
                                filetypes=[("Word Document", "*.docx")],
                                initialfile=os.path.basename(output_docx)
                            )
                            if new_path:
                                output_docx = new_path
                                self.current_project_dir = os.path.dirname(new_path)
                                continue
                            else: return
                        else: return
                except Exception as e:
                    messagebox.showerror("Błąd / Error", str(e))
                    return

            if mode in ['pdf_no_att', 'pdf_full'] and PDF_MERGE_AVAILABLE:
                output_pdf = output_docx.replace(".docx", ".pdf")
                final_pdf = output_pdf.replace(".pdf", "_KOMPLET.pdf") if mode == 'pdf_full' else output_pdf
                
                all_pdfs_to_merge = []
                
                if mode == 'pdf_full':
                    if self.engine_pdf_path.get() and os.path.exists(self.engine_pdf_path.get()):
                        all_pdfs_to_merge.append(self.engine_pdf_path.get())
                        
                    if self.rad_active.get() and self.rad_pdf_path.get() and os.path.exists(self.rad_pdf_path.get()):
                        all_pdfs_to_merge.append(self.rad_pdf_path.get())
                        
                    for k in self.aktywne:
                        if self.aktywne[k].get():
                            for att in self.mod_attachments.get(k, []):
                                if att['type'].get() == "Dokument (PDF)" and os.path.exists(att['path'].get()):
                                    all_pdfs_to_merge.append(att['path'].get())
                
                saved_pdf = False
                while not saved_pdf:
                    try:
                        # Przechwytujemy standardowe wyjście, aby tqdm nie zablokował biblioteki
                        with open(os.devnull, "w") as devnull:
                            old_stdout, old_stderr = sys.stdout, sys.stderr
                            sys.stdout, sys.stderr = devnull, devnull
                            try:
                                docx_to_pdf(output_docx, output_pdf)
                            except Exception as pdf_e:
                                # Ignorujemy błąd "Quit" w Word.Application, o ile plik powstał
                                if "Quit" in str(pdf_e) and os.path.exists(output_pdf):
                                    pass 
                                else:
                                    raise pdf_e
                            finally:
                                sys.stdout, sys.stderr = old_stdout, old_stderr
                        
                        if mode == 'pdf_full' and all_pdfs_to_merge:
                            merger = PdfMerger()
                            merger.append(output_pdf)
                            for att in all_pdfs_to_merge:
                                merger.append(att)
                                
                            merger.write(final_pdf)
                            merger.close()
                            if os.path.exists(output_pdf): os.remove(output_pdf)
                            messagebox.showinfo("Sukces", f"Zapisano KOMPLETNY raport PDF z załącznikami:\n{final_pdf}")
                        else:
                            messagebox.showinfo("Sukces", f"Zapisano raport PDF:\n{output_pdf}")
                        saved_pdf = True
                        
                    except PermissionError:
                        choice = messagebox.askyesnocancel(
                            t['err_in_use_title'],
                            t['err_in_use_msg'].format(final_pdf)
                        )
                        if choice is True: continue
                        elif choice is False:
                            new_path = filedialog.asksaveasfilename(
                                defaultextension=".pdf", 
                                filetypes=[("PDF Document", "*.pdf")],
                                initialfile=os.path.basename(final_pdf)
                            )
                            if new_path:
                                final_pdf = new_path
                                continue
                            else: return
                        else: return
                    except Exception as e:
                        messagebox.showerror("Błąd PDF", f"Word zapisał się, ale wystąpił błąd przy konwersji PDF:\n{e}")
                        return
            else:
                messagebox.showinfo("Sukces", f"Zapisano raport:\n{output_docx}")
                
        except Exception as global_e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Błąd Generowania", f"Wystąpił nieoczekiwany problem podczas generowania raportu:\n{str(global_e)}")

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    if 'clam' in style.theme_names(): style.theme_use('clam')
    app = AcousticReportApp(root)
    root.mainloop()