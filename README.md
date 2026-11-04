# Kalkulator Akustyki - raporty akustyczne obudów agregatów

Program na Windows do obliczeń i raportów akustycznych obudów agregatów prądotwórczych: poziom hałasu wg ISO 3744 w wielu punktach pomiarowych, pochłanianie wnętrza (wzór Sabine'a), izolacyjność przegród z bazy materiałów, widmo silnika, wydech i chłodnica. Raport PDF / DOCX (PL / EN).

> Status: **wersja 2.0.0** (2026-11-05). Instalator: [Releases](../../releases/latest)

## Możliwości

- **Jeden ekran: dane i wyniki** - wyniki liczone na bieżąco podczas wpisywania, jasny / ciemny motyw (jak diamel.pl, Korektor HR i Kalkulator Rozruchu)
- **Widma w 8 pasmach oktawowych** (63 Hz - 8 kHz) z wyborem jednostki dB / dB(A) - korekta A tylko dla widm w dB
- **Moduły obudowy** - ściany, podłoga, sufit, drzwi, czerpnia, wyrzut, wydech; warstwy z bazy materiałów albo własne (widmo albo Rw / Ctr)
- **Punkty pomiarowe** - dowolna liczba, także kilka w tej samej odległości; udział każdego modułu w hałasie w punkcie
- **Raport PDF / DOCX** - wykres widma, udziały modułów, ocena każdego punktu, opis metody; karty producentów PDF dołączane na końcu; szablon z nagłówkiem i stopką (Ustawienia albo papier firmowy z konta Diamel)
- **Logowanie kontem Diamel** przy uruchomieniu (nieobowiązkowe) - papier firmowy, inicjały, karty producentów z bazy
- **Baza materiałów, silników i chłodnic** z serwera diamel.pl (edytowana w panelu administracyjnym) + własne wpisy, zgłaszanie nowych pozycji z załącznikami
- **Automatyczna aktualizacja** programu z GitHub Releases i bazy z serwera

## Uruchomienie ze źródeł

Python 3 (Windows), zależności:

```
pip install pywebview python-docx pillow pypdf pywin32
```

```
python aplikacja.py
```

`AKUSTYKA_DEBUG=1` - narzędzia programisty (F12) w oknie. Raport PDF zapisuje Microsoft Word.

## Testy

```
python testy/test_obliczenia.py
```

Testy na danych syntetycznych; z folderem `testy/dane_prywatne` (projekty klientów, poza repozytorium) dodatkowo porównanie z wynikami wersji 1.1.0.

## Baza danych

Baza jest na serwerze diamel.pl i edytuje się ją w panelu administracyjnym (zakładka Kalkulator Akustyki) - każdy zapis to nowa wersja, programy pobierają ją przy starcie. Pliki producentów (karty katalogowe) są dostępne tylko po zalogowaniu kontem z dostępem do programu.

## Wydanie nowej wersji programu

1. Podnieś `VERSION` i dopisz zmiany do `CHANGELOG` w `konfiguracja_app.py`, zrób commit.
2. `python kompilacja.py` pobiera kopię bazy z serwera i buduje program (PyInstaller) i instalator (Inno Setup 6) do `<wersja>/Paczka/`.
3. Na pytanie o publikację odpowiedz `t` - skrypt wypchnie commit i utworzy wydanie `v<wersja>` na GitHub ([GitHub CLI](https://cli.github.com/)).

Program nie zawiera haseł - zgłoszenia przyjmuje `diamel.pl/api/zgloszenie.php`, konto `diamel.pl/api/konto.php`.

## Licencja

**Wszelkie prawa zastrzeżone** - © 2026 Diamel ([diamel.pl](https://diamel.pl)). Kod jest udostępniony wyłącznie do wglądu, szczegóły w [LICENSE](LICENSE).

Autor: **Diamel** - [diamel.pl](https://diamel.pl)
