import os
import ctypes
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from konfiguracja_app import sciezka_do_zasobu
from gui_akustyka import AcousticReportApp

def pokaz_splash_screen(root, po_zakonczeniu_callback):
    root.withdraw()
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg="white")
    
    splash_w, splash_h = 600, 350
    screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
    x = (screen_w // 2) - (splash_w // 2)
    y = (screen_h // 2) - (splash_h // 2)
    splash.geometry(f"{splash_w}x{splash_h}+{x}+{y}")
    
    try:
        # Możesz wstawić ładny render agregatu w zasobach jako "splash.jpg"
        img_path = sciezka_do_zasobu(os.path.join("Zasoby", "logo_duze_diamel.png"))
        oryginalny_obraz = Image.open(img_path)
        oryginalny_obraz.thumbnail((560, 240), Image.Resampling.LANCZOS)
        bg_image = ImageTk.PhotoImage(oryginalny_obraz)
        splash.bg_image = bg_image
        lbl_img = tk.Label(splash, image=bg_image, bg="white")
        lbl_img.pack(pady=(20, 10))
    except Exception:
        lbl_title = tk.Label(splash, text="SYMULATOR AKUSTYKI\nDIAMEL", font=("Arial", 22, "bold"), bg="white", fg="#2196F3")
        lbl_title.pack(expand=True)
        
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TProgressbar", thickness=6)
    
    progress = ttk.Progressbar(splash, orient="horizontal", length=500, mode="determinate", style="TProgressbar")
    progress.pack(side="bottom", pady=(15, 20))
    
    lbl_status = tk.Label(splash, text="Inicjalizacja środowiska bazowego...", font=("Arial", 10, "italic"), bg="white", fg="gray")
    lbl_status.pack(side="bottom")
    
    komunikaty = [
        (0, "Inicjalizacja środowiska bazowego..."),
        (25, "Ładowanie silnika akustycznego ISO 3744..."),
        (50, "Wczytywanie bazy maszyn i materiałów..."),
        (75, "Integracja modułów DOCX/PDF..."),
        (95, "Uruchamianie interfejsu graficznego...")
    ]
    
    def aktualizuj_postep(krok=0):
        if krok <= 100:
            progress['value'] = krok
            for prog, text in reversed(komunikaty):
                if krok >= prog:
                    lbl_status.config(text=text)
                    break
            splash.after(40, aktualizuj_postep, krok + 1) # ~4 sekundy
        else:
            splash.destroy()
            root.deiconify()
            po_zakonczeniu_callback()

    aktualizuj_postep(0)


if __name__ == "__main__":
    try:
        myappid = "diamel.akustyka.main"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception: pass

    root = tk.Tk()
    
    try:
        icon_path = sciezka_do_zasobu(os.path.join("Zasoby", "ikona.ico"))
        if os.path.exists(icon_path): root.iconbitmap(default=icon_path)
    except Exception: pass

    def start_aplikacji(): app = AcousticReportApp(root)
    pokaz_splash_screen(root, start_aplikacji)
    root.mainloop()