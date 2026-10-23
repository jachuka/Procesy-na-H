# Procesy na sztukę

Aplikacja do przeliczania wyceny Excel na godziny dziesiętne na sztukę do Subiekta Nexo. Bez tabeli przestawnej.

## Uruchomienie

Python 3.11 lub nowszy:

```sh
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py --port 8000
```

Otwórz w przeglądarce adres lokalnego serwera na porcie 8000. Domyślnie serwer jest dostępny wyłącznie na tym komputerze. Opcja `--host 0.0.0.0` służy do uruchamiania za kontrolowanym proxy; aplikacja nie ma własnego logowania i nie jest przygotowana do publicznego hostingu.

Wgraj XLSX, wybierz arkusz, jeśli plik zawiera kilka wycen. Kopiuj samą liczbę przy procesie lub całą tabelę z nazwami. Domyślnie wynik ma sześć miejsc po przecinku; ustaw precyzję zgodnie z polem w Subiekcie. Aplikacja nie integruje się bezpośrednio z Subiektem.

## Układ wyceny

- Nakład: K4, dodatnia całkowita liczba sztuk; nazwa projektu: H4.
- Nagłówek sekcji: `OPIS PROCESU` w H i nagłówek zawierający `RBH` w M, w pierwszych 100 wierszach.
- Kolejne wiersze: opis w H, czas Excel (ułamek doby) w M. Pierwszy pusty opis kończy sekcję.
- Grupa przed myślnikiem: FREZ, LASER, MANUALNE, SZTANCA, GILOTYNA, FORMIERKA. Nieznane grupy blokują wynik zamiast znikać z podsumowania.
- `RBH NA ROZRUCH PRODUKCJI` doliczany do MANUALNE.
- Wynik: suma czasów w M × 24 / K4. Zaokrąglenie dopiero do prezentacji i kopiowania. Czasy powyżej 24 godzin są obsługiwane.

Arkusz TAB jest pomijany. Odczytywane są zapisane wyniki formuł, bez ich wykonywania. Przed wgraniem przelicz i zapisz plik w Excelu. Nieaktualnych zapisanych wyników nie można automatycznie rozpoznać. Brakujący wynik lub błąd w czasie blokuje obliczenie. Pliki są przetwarzane w pamięci i nie są zapisywane na serwerze. Limit przesłania: 10 MB, rozpakowanej zawartości: 50 MB. Stary format XLS i pliki szyfrowane nie są obsługiwane.

## Testy

```sh
python -m unittest discover -s tests -v
```
