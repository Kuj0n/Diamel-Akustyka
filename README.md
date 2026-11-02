# Kalkulator Akustyki - raporty akustyczne obudów agregatów

Program na Windows do obliczeń i raportów akustycznych obudów kontenerowych agregatów prądotwórczych: poziomy hałasu wg ISO 3744 w wielu punktach pomiarowych, pochłanianie wnętrza (Sabine), izolacyjność przegród z bazy materiałów, widmo silnika i wydechu, chłodnica. Raport DOCX / PDF (PL / EN).

> Status: **wersja 1.1.0** (2026-11-02). Instalator: [Releases](../../releases/latest)

## Możliwości

- **Dane ogólne** - projekt, klient, obudowa, tabela rewizji, logo i stopka raportu
- **Widmo silnika** - pasma 63 Hz - 8 kHz, wydech z uwzględnieniem wysokości komina
- **Chłodnica**, czerpnia, wyrzut i inne moduły obudowy z warstwami materiałów
- **Punkty pomiarowe** - wiele dystansów, obliczenia ISO 3744
- **Raport DOCX**, opcjonalnie PDF z załącznikami
- **Baza materiałów, silników i chłodnic** (`acoustic_db.json`) - własne wpisy, scalanie z bazą globalną, zgłaszanie nowych materiałów przez formularz na diamel.pl

## Uruchomienie ze źródeł

Python 3 (Windows), zależności:

```
pip install python-docx docx2pdf PyPDF2 pillow
```

```
python start_akustyka.py
```

## Baza danych

Baza globalna to plik [`acoustic_db.json`](acoustic_db.json) w tym repozytorium. Program porównuje jego pole `version` z wersją lokalnej bazy i proponuje scalenie nowych wpisów z własnymi.

Aktualizacja bazy **bez wydawania nowej wersji programu**:

1. Edytuj `acoustic_db.json` (ręcznie albo w programie uruchomionym ze źródeł - zapisuje do tego pliku).
2. Podnieś `version` i dopisz wpis na początku `changelog`, np. `"v1.0.3 (2026-11-10)\n- Dodano ...\n\n"` + poprzednia historia.
3. Commit + push do `master`. Programy zaproponują aktualizację przy następnym uruchomieniu (GitHub odświeża plik do 5 minut).

## Wydanie nowej wersji programu

1. Podnieś `VERSION` i dopisz zmiany do `CHANGELOG` w `konfiguracja_app.py`, zrób commit.
2. `python kompilacja.py` buduje program (PyInstaller) i instalator (Inno Setup 6) do `<wersja>/Paczka/`.
3. Na pytanie o publikację odpowiedz `t` - skrypt wypchnie commit i utworzy wydanie `v<wersja>` na GitHub (wymaga [GitHub CLI](https://cli.github.com/) zalogowanego przez `gh auth login`).

Program nie zawiera haseł - zgłoszenia materiałów przyjmuje `diamel.pl/api/zgloszenie.php`.

## Licencja

**Wszelkie prawa zastrzeżone** - © 2026 Diamel ([diamel.pl](https://diamel.pl)). Kod jest udostępniony wyłącznie do wglądu, szczegóły w [LICENSE](LICENSE).

Autor: **Diamel** - [diamel.pl](https://diamel.pl)
