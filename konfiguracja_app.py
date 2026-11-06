import os
import sys

# ==========================================
# KONFIGURACJA I METADANE
# ==========================================
VERSION = "2.0.8"
DB_VERSION = "0"  # wersja bazy, gdy lokalny plik jej nie zawiera (baza z panelu diamel.pl ma numer wersji 1, 2, 3...)
BUILD_DATE = "2026-11-06 19:00"

CHANGELOG = """
v2.0.8 (2026-11-06)
- Raport, pierwsza strona: jedno podsumowanie wyników zamiast kafelków i powtórzonej ramki - werdykt w ramce i tabela
  punktów (odległość, wymagany, obliczony, zapas, ocena); przypis, gdy decyduje strona komina.

v2.0.7 (2026-11-06)
- Raport PDF: karty producentów (PDF) w swoich rozdziałach - karta silnika po rozdziale silnika, chłodnicy po chłodnicy,
  modułów po budowie obudowy, wydechu po rozdziale wydechu (wcześniej wszystkie na końcu raportu).
- Raport: wyraźny wynik końcowy na pierwszej stronie (ramka: spełnia / nie spełnia, punkt krytyczny, zapas) i ramka
  z wynikiem przy każdym punkcie pomiarowym; kafelki punktów w kolorze wyniku.

v2.0.6 (2026-11-06)
- Raport: kafelek punktu pokazuje wynik decydujący (gorszy z dwóch: po stronie komina / średnio) - zgodny z zapasem;
  gdy decyduje strona komina, pod spodem średnia wokół obudowy.

v2.0.5 (2026-11-06)
- Raport: osobne rozdziały Chłodnica i Wydech silnika (chłodnica tylko gdy włączona).
- Wydech: hałas 1 m od wylotu za tłumikiem w dB(A) - jak w kartach tłumików (np. AXCES 85 dB(A)@1m) - w raporcie
  (wiersz tabeli i wynik) oraz w zakładce Wydech.

v2.0.4 (2026-11-06)
- Wydech: na zakładce i przy module udział po stronie komina (wcześniej pokazywał 0% ze średniej wokół obudowy).
- Kierunek wylotu: krótsze opisy przycisków (Poziomo / W górę) - nie są już ucinane.
- Test aktualizacji bez klikania instalatora.

v2.0.3 (2026-11-06)
- Aktualizacja bez klikania instalatora: program pobiera nową wersję, instaluje ją w tle i uruchamia się ponownie sam
  (działa przy aktualizacjach z tej wersji na kolejne).

v2.0.2 (2026-11-06)
- Wydech: wynik po stronie komina. Po podaniu położenia wylotu (wysokość nad dachem, odległość od ściany, wysokość
  punktu pomiaru) wydech liczony jest także jako źródło punktowe w rzeczywistej odległości od wylotu. Przy punkcie
  widać oba wyniki - po stronie komina i średnio wokół obudowy; o zgodności decyduje gorszy. Opis w raporcie.
- Wydech: widmo można wpisać prosto z karty tłumika jako Lp@1m od wylotu - program przelicza je na moc (+11 dB).
  Podpowiedź przy pasmach (pomiń 31,5 Hz; kolumna dB czy dB(A)) i ostrzeżenie, gdy widmo wydechu jest puste.
- Raport: przeliczenie widma wydechu, tłumienie tłumika i LwA za tłumikiem w tabeli.

v2.0.1 (2026-11-05)
- Ekran obliczeń w zakładkach: Projekt, Obudowa i punkty, Silnik, osobno każdy moduł (ściany, podłoga, sufit, drzwi,
  czerpnia, wyrzut, wydech) i Chłodnica. Przy module - jego udział w hałasie; moduł, który hałasuje najbardziej, wyróżniony.

v2.0.0 (2026-11-05)
- Nowy program: okno w stylu Korektora HR i Kalkulatora Rozruchu, wyniki liczone na bieżąco podczas wpisywania.
- Logowanie kontem Diamel przy uruchomieniu (jak w Korektorze HR): raporty na papierze firmowym Twojej firmy,
  inicjały w raporcie, pliki producentów z bazy; można pracować bez logowania.
- Baza materiałów, silników i chłodnic na serwerze diamel.pl (edytowana w panelu administracyjnym), pobierana przy
  starcie; wpisy własne zostają przy aktualizacji. Karty producentów - po zalogowaniu.
- Naprawa: udział modułów w raporcie - każdy moduł liczony względem powierzchni pomiarowej, moduły sumują się do
  wyniku całkowitego (wcześniej małe moduły, np. czerpnia, wychodziły głośniejsze niż cała obudowa). Wynik całkowity
  w punkcie bez zmian.
- Naprawa: dwa punkty pomiarowe w tej samej odległości (np. front i bok 1 m) - oba w wynikach i raporcie.
- Naprawa: jedno widmo silnika na ekranie, w obliczeniach i w raporcie. Przy każdym widmie (silnik, wydech, chłodnica)
  wybór jednostki dB / dB(A) - korekta A dodawana tylko do widm w dB.
- Raport PDF / DOCX w nowym układzie: wykres widma, udział modułów, ocena każdego punktu, opis metody; szablon
  z nagłówkiem i stopką (Ustawienia albo papier firmowy z konta).
- Projekty z wersji 1.x otwierają się bez zmian wyników.

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

# Aktualizacje (GitHub: wydania programu), baza (panel diamel.pl), zgłoszenia i konto (diamel.pl)
NAZWA_PROGRAMU = "akustyka"
GITHUB_REPO = "Kuj0n/Diamel-Akustyka"
URL_BAZY = "https://diamel.pl/api/akustyka.php"  # baza materiałów, silników i chłodnic edytowana w panelu administracyjnym
URL_ZGLOSZEN = "https://diamel.pl/api/zgloszenie.php"
URL_KONTA = "https://diamel.pl/api/konto.php"  # logowanie kontem Diamel (szablon raportu firmy, inicjały)


def sciezka_do_zasobu(relatywna_sciezka):
    """ Pancerna funkcja szukająca plików zasobów w trybie .exe i developerskim """
    katalogi = [
        os.path.dirname(os.path.abspath(__file__)),
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
