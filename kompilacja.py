import os
import re
import json
import subprocess
import shutil

# --- KONFIGURACJA ---
folder_zasoby = "Zasoby"
plik_glowny = "start_akustyka.py"
pliki_modulow = ["gui_akustyka.py", "narzedzia.py", "konfiguracja_app.py"]
plik_ikony = os.path.join(folder_zasoby, "ikona.ico")

inno_setup_compiler = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not os.path.exists(inno_setup_compiler):
    print("❌ Błąd: Nie znaleziono kompilatora Inno Setup!")
    exit(1)

print("🔍 Analizowanie konfiguracji (wersja MVC)...")
with open("konfiguracja_app.py", "r", encoding="utf-8") as f:
    kod_konf = f.read()

wersja = re.search(r'VERSION\s*=\s*"([^"]+)"', kod_konf).group(1)
data_wydania = re.search(r'BUILD_DATE\s*=\s*"([^"]+)"', kod_konf).group(1)[:10]
changelog_full = re.search(r'CHANGELOG\s*=\s*"""(.*?)"""', kod_konf, re.DOTALL).group(1).strip()
github_repo = re.search(r'GITHUB_REPO\s*=\s*"([^"]+)"', kod_konf).group(1)

# Paczka = pliki wydania na GitHub (instalator). Baza nie jest częścią wydania -
# program pobiera acoustic_db.json z gałęzi master repozytorium.
folder_archiwum = os.path.join(wersja, "Archiwum")
folder_paczka = os.path.join(wersja, "Paczka")

os.makedirs(folder_archiwum, exist_ok=True)
os.makedirs(folder_paczka, exist_ok=True)

print("📂 Kopiowanie rozdzielonych modułów i zasobów...")
for m_plik in [plik_glowny] + pliki_modulow:
    if os.path.exists(m_plik): shutil.copy(m_plik, os.path.join(folder_archiwum, m_plik))
if os.path.exists(folder_zasoby):
    shutil.copytree(folder_zasoby, os.path.join(folder_archiwum, folder_zasoby), dirs_exist_ok=True)
if os.path.exists(plik_ikony): shutil.copy(plik_ikony, os.path.join(folder_archiwum, "ikona.ico"))

v_parts = wersja.split('.') + ['0'] * (4 - len(wersja.split('.')))
v_str = ".".join(v_parts)
nazwa_stala = "Akustyka_Kalkulator"
nazwa_instalatora_exe = f"Instalator_Akustyka_v{wersja}.exe"

with open(os.path.join(folder_archiwum, "wersja_info.txt"), "w", encoding="utf-8") as f:
    f.write(f"VSVersionInfo(ffi=FixedFileInfo(filevers=({v_parts[0]},{v_parts[1]},{v_parts[2]},{v_parts[3]}), prodvers=({v_parts[0]},{v_parts[1]},{v_parts[2]},{v_parts[3]}), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)), kids=[StringFileInfo([StringTable(u'041504b0', [StringStruct(u'CompanyName', u'Diamel'), StringStruct(u'FileDescription', u'Akustyka Raporty'), StringStruct(u'FileVersion', u'{v_str}'), StringStruct(u'OriginalFilename', u'{nazwa_stala}.exe'), StringStruct(u'ProductName', u'Akustyka Kalkulator'), StringStruct(u'ProductVersion', u'{v_str}')])]), VarFileInfo([VarStruct(u'Translation', [1045, 1200])])])")

# Lista zmian tej wersji = opis wydania na GitHub
wzorzec = rf'(v{wersja}.*?)(?=\nv\d+\.\d+|$)'
match = re.search(wzorzec, changelog_full, re.DOTALL)
log_wersji = match.group(1).strip() if match else f"v{wersja}\nZaktualizowano."

print(f"⏳ Kompilacja struktury programu dla wersji {wersja}...")
cmd = (f'python -m PyInstaller --onedir --noconsole --clean --icon="ikona.ico" '
       f'--add-data "ikona.ico;." --add-data "{folder_zasoby};{folder_zasoby}" '
       f'--version-file="wersja_info.txt" -n "{nazwa_stala}" "{plik_glowny}"')
subprocess.run(cmd, shell=True, cwd=folder_archiwum, check=True)

sciezka_folderu_dist = os.path.join(folder_archiwum, "dist", nazwa_stala)

plik_info_txt = os.path.join(folder_archiwum, "info_instalatora.txt")
with open(plik_info_txt, "w", encoding="utf-8") as f: f.write(f"KALKULATOR AKUSTYKI - DIAMEL\nWersja: {wersja}\n\n{changelog_full}")

inno_skrypt = f"""
[Setup]
AppId=Akustyka_Kalkulator_Diamel
AppName=Kalkulator Akustyki
AppVersion={wersja}
AppPublisher=Diamel
DefaultDirName={{autopf}}\\Kalkulator Akustyki
DefaultGroupName=Kalkulator Akustyki
DisableWelcomePage=no
DisableProgramGroupPage=yes
OutputDir={os.path.abspath(folder_paczka)}
OutputBaseFilename=Instalator_Akustyka_v{wersja}
SetupIconFile={os.path.abspath(os.path.join(folder_archiwum, "ikona.ico"))}
Compression=lzma2/ultra
SolidCompression=yes
PrivilegesRequired=lowest
InfoBeforeFile={os.path.abspath(plik_info_txt)}

[Languages]
Name: "pl"; MessagesFile: "compiler:Languages\\Polish.isl"

[Tasks]
Name: "desktopicon"; Description: "Utwórz skrót na pulpicie"

[Files]
Source: "{os.path.abspath(sciezka_folderu_dist)}\\*"; DestDir: "{{app}}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{{group}}\\Kalkulator Akustyki"; Filename: "{{app}}\\{nazwa_stala}.exe"
Name: "{{autodesktop}}\\Kalkulator Akustyki"; Filename: "{{app}}\\{nazwa_stala}.exe"; Tasks: desktopicon

[Run]
Filename: "{{app}}\\{nazwa_stala}.exe"; Description: "Uruchom Kalkulator"; Flags: nowait postinstall skipifsilent
"""
plik_iss = os.path.join(folder_archiwum, "skrypt_instalatora.iss")
with open(plik_iss, "w", encoding="utf-8") as f: f.write(inno_skrypt)

print("📦 Kompresowanie instalatora...")
subprocess.run(f'"{inno_setup_compiler}" "{plik_iss}"', shell=True, check=True)
print(f"🚀 SUKCES! Wersja: {wersja} gotowa w folderze Paczka.")

# --- PUBLIKACJA NA GITHUB ---
odp = input(f"\n🌐 Opublikować wydanie v{wersja} na GitHub ({github_repo})? [t/N]: ").strip().lower()
if odp == "t":
    gh = shutil.which("gh") or r"C:\Program Files\GitHub CLI\gh.exe"
    if subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True).stdout.strip():
        print("❌ Niezatwierdzone zmiany w repozytorium - najpierw commit, potem ponów publikację.")
        exit(1)
    subprocess.run(["git", "push", "origin", "HEAD"], check=True)
    subprocess.run([gh, "release", "create", f"v{wersja}", os.path.join(folder_paczka, nazwa_instalatora_exe),
                    "--repo", github_repo, "--target", "master", "--title", f"v{wersja}", "--notes", log_wersji], check=True)
    print(f"✅ Opublikowano: https://github.com/{github_repo}/releases/tag/v{wersja}")
else:
    print("Publikację pominięto. Instalator jest w folderze Paczka.")