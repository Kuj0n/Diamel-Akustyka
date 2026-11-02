import os
import sys

# ==========================================
# KONFIGURACJA I METADANE
# ==========================================
VERSION = "1.1.0"
DB_VERSION = "1.0.0"
BUILD_DATE = "2026-11-02 23:00"

CHANGELOG = """
v1.1.0 (2026-11-02)
- Aktualizacje programu pobierane z GitHub (github.com/Kuj0n/Diamel-Akustyka).
- Baza danych aktualizowana z repozytorium - nowe materiały bez instalowania nowej wersji programu.
- Zgłoszenia materiałów wysyłane przez formularz na diamel.pl - program nie zawiera już hasła do poczty.
- Domyślna stopka raportu: Diamel.

v1.0.0 (2026-09-16)
- Wdrożenie aplikacji w profesjonalnej architekturze MVC.
- Dodano mechanizm auto-aktualizacji (OTA) oraz instalator Inno Setup.
- Wbudowano zakładkę Pomoc i wysyłanie próśb o nowe materiały.
- Pełne obliczenia ISO 3744 z uwzględnieniem wielu dystansów i wpływu wysokości komina.
"""

CHANGELOG_DB = """
v1.0.0 (2026-09-16)
- Dodano wbudowaną bazę materiałów tłumiących.
- Zaimplementowano bazę silników Cummins (w tym widma wydechowe).
- Dodano parametry chłodnic Berry Radiateurs.
"""

# Aktualizacje (GitHub: wydania programu, baza z gałęzi master) i zgłoszenia (formularz na diamel.pl)
NAZWA_PROGRAMU = "akustyka"
GITHUB_REPO = "Kuj0n/Diamel-Akustyka"
URL_BAZY = f"https://raw.githubusercontent.com/{GITHUB_REPO}/master/acoustic_db.json"
URL_ZGLOSZEN = "https://diamel.pl/api/zgloszenie.php"

def sciezka_do_zasobu(relatywna_sciezka):
    """ Pancerna funkcja szukająca plików zasobów w trybie .exe i developerskim """
    katalogi = [
        os.path.abspath("."),
        os.path.join(os.path.abspath("."), "_internal")
    ]
    if hasattr(sys, '_MEIPASS'):
        katalogi.insert(0, sys._MEIPASS)
        katalogi.insert(1, os.path.join(sys._MEIPASS, "_internal"))
        
    if hasattr(sys, 'executable'):
        baza_exe = os.path.dirname(sys.executable)
        katalogi.append(baza_exe)
        katalogi.append(os.path.join(baza_exe, "_internal"))
        
    for k in katalogi:
        pelna = os.path.join(k, relatywna_sciezka)
        if os.path.exists(pelna):
            return pelna
            
    return relatywna_sciezka