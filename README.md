# Procesy na sztukę

Aplikacja do przeliczania wyceny Excel na godziny dziesiętne na sztukę do Subiekta Nexo. Bez tabeli przestawnej.

## Wersja Windows EXE

W sekcji **Releases** repozytorium pobierz `Procesy-na-sztuke.exe` (Windows 10/11, 64-bit). Uruchom dwuklikiem; Python ani Excel nie muszą być zainstalowane. Otworzy się przeglądarka i małe okno sterujące. Zamknięcie okna sterującego wyłącza aplikację. Zamkniętą kartę możesz otworzyć ponownie przyciskiem „Otwórz aplikację”. Dane pozostają na Twoim komputerze; po pobraniu program działa offline. Plik nie jest podpisany cyfrowo.

Plik EXE jest budowany na Windows przez GitHub Actions po wysłaniu tagu `v*` lub zmianie numeru w `VERSION` na gałęzi `main`. Publikacja w Releases następuje dopiero po testach obliczeń i uruchomienia gotowego EXE. Ręczny workflow „Aplikacja Windows” udostępnia wynik jako artefakt. Lokalne budowanie na Windows:

```powershell
python -m pip install -r requirements.txt pyinstaller==6.16.0
python -m PyInstaller --noconfirm --clean --onefile --windowed --name Procesy-na-sztuke --add-data "static:static" desktop.py
```

## Uruchomienie

Python 3.11 lub nowszy:

```sh
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py --port 8000
```

Otwórz w przeglądarce adres lokalnego serwera na porcie 8000. Domyślnie serwer jest dostępny wyłącznie na tym komputerze. Opcja `--host 0.0.0.0` służy do uruchamiania za kontrolowanym proxy; aplikacja nie ma własnego logowania i nie jest przygotowana do publicznego hostingu.

Wgraj XLSX, wybierz arkusz, jeśli plik zawiera kilka wycen. Kopiuj samą liczbę przy procesie lub całą tabelę z nazwami. Domyślnie wynik ma sześć miejsc po przecinku; ustaw precyzję zgodnie z polem w Subiekcie. Aplikacja nie integruje się bezpośrednio z Subiektem.

Obok wyników RBH aplikacja pokazuje zużycie tworzyw: opis, m² łącznie oraz m² / szt. Przycisk „Kopiuj” przy tworzywie kopiuje samą liczbę m² / szt. z przecinkiem. Możesz też skopiować całą tabelę tworzyw. Wybrana precyzja dotyczy wyświetlania i kopiowania obu tabel.

### Dodatki

Przed tworzywami znajduje się tabela „Dodatki” z opisem i specyfikacją, sekcją, jednostką i ilością na sztukę wyrobu. Przycisk przy pozycji kopiuje samą ilość; można też skopiować całą tabelę. Koszty nie są odczytywane do tej tabeli.

- Komponenty: ilość z C / liczba sztuk wyrobu z K w danym wierszu. Przykłady: karton na 12 sztuk = 1/12; 6 nitów na 1 wyrób = 6.
- Pianka: zapisane m² z E / liczba sztuk wyrobu z K; jednostka m²/szt. Opis z F.
- Dodatki na cały nakład i dostawy tworzywa: ilość z C / nakład z K4.
- Wiersze komponentów i dodatków mają specyfikację w E i opis w H. Puste pozycje są pomijane. Brak poprawnej ilości lub dodatniego dzielnika pokazuje ostrzeżenie z adresem komórki i nie blokuje innych wyników.
- Nagłówki tabel są wyszukiwane w wybranym arkuszu; transport do klienta oraz tworzywa pozostają poza sekcją dodatków. Wyniki są liczone osobno dla każdego arkusza, bez zaokrąglania ilości do całych opakowań.

### Bloki tworzyw

- Każdy blok zaczyna nagłówek `TWORZYWO I OPIS`; opis znajduje się bezpośrednio pod nim (także pod scalonym nagłówkiem).
- Odpowiadający nagłówek `m2 łącznie` lub `m² łącznie` znajduje się w tym samym bloku, na prawo od opisu, w wierszu nagłówka bloku lub opisu. Wartość jest bezpośrednio pod nagłówkiem. Przykład: opis w H76, nagłówek w J76, wartość w J77; scalenia F:H są obsługiwane.
- Bloki mogą występować jeden pod drugim lub obok siebie. Pusty opis pomija blok. Oryginalne pozycje pozostają widoczne, a tabela „Tworzywa łącznie” sumuje m² dla tego samego tworzywa i dzieli sumę przez K4.
- Grupowanie porównuje nazwę przed pierwszym myślnikiem (`-`, `–`, `—`), pomija wielkość liter i nadmiarowe spacje, ujednolica grubości (`3mm`, `3 mm`, `3,0 mm`) oraz podstawowe określenia białego, czarnego i bezbarwnego koloru (np. `WH`, `biała`, `white`). Kolejność słów nie wpływa na dopasowanie. Pozostałe oznaczenia, w tym typ tworzywa, kolor, wykończenie i model, muszą się zgadzać; program nie zgaduje nieznanych zamienników. Nazwa w podsumowaniu pochodzi z pierwszej pozycji.
- Podsumowanie ma przyciski do kopiowania zużycia na sztukę i całej tabeli z sumą m². Sumy liczone są przed zaokrągleniem. Gdy blok zgłasza błąd odczytu, podsumowanie obejmuje wyłącznie poprawne pozycje.
- Zużycie na sztukę = zapisany wynik m² łącznie / K4. Wartość zero jest poprawna; zaokrąglenie następuje przy prezentacji.
- Brak lub niepoprawny wynik m² zgłasza ostrzeżenie z adresem komórki, bez zastępowania go zerem. Wyniki procesów nadal są dostępne.

### Tworzywo na druk OFFSET TORRO

Po „Tworzywa łącznie” znajduje się osobna tabela z nazwą tworzywa, formatem arkusza, liczbą arkuszy łącznie i Ark. / 1 szt. Przycisk przy pozycji kopiuje samą liczbę Ark./szt.; kopiowanie całej tabeli obejmuje także nazwę, format i łączną liczbę arkuszy.

- Odczyt dotyczy sekcji `TWORZYWO NA DRUK OFFSET TORRO`, do początku następnej sekcji druku. Nie pobiera ilości z tabel kosztów druku ani OFFSET VIPER.
- Nazwa znajduje się pod nagłówkiem `TWORZYWO NA OFFSET` (starszy wariant `TWORZYWO 1 NA OFFSET` również działa), arkusze pod `ARK łącznie`, a format pod `ARKUSZ - ARK - FORMAT`. Obsługiwane są scalone komórki, np. nazwa F126:H126, format D127:E127 oraz ilość J127.
- Ark./szt. = arkusze łącznie / K4 wybranego arkusza. Przykład: 1400 arkuszy / 1000 szt. = 1,4 Ark./szt. Uwzględniana jest wskazana łączna liczba arkuszy, wraz z zapisanym w niej zapasem.
- Brak opisu pomija pozycję. Brak poprawnej ilości zgłasza ostrzeżenie i nie jest zastępowany zerem. Brak formatu jest oznaczany przy zachowaniu poprawnej ilości. Dotychczasowe wyniki m² są liczone oddzielnie.

## Układ wyceny

- Nakład: K4, dodatnia całkowita liczba sztuk; nazwa projektu: H4.
- Nagłówek sekcji: `OPIS PROCESU` w H i nagłówek zawierający `RBH` w M, w pierwszych 100 wierszach.
- Kolejne wiersze: opis w H, czas Excel (ułamek doby) w M. Pierwszy pusty opis kończy sekcję.
- Grupa przed myślnikiem: FREZ, LASER, MANUALNE, SZTANCA, GILOTYNA, FORMIERKA. Nieznane grupy blokują wynik zamiast znikać z podsumowania.
- `RBH NA ROZRUCH PRODUKCJI` doliczany do MANUALNE.
- `TERMOFORMIERKA + ROLSZTANCA` ma osobną grupę wynikową, pokazywaną, gdy taki proces występuje w wycenie.
- Wynik: suma czasów w M × 24 / K4. Zaokrąglenie dopiero do prezentacji i kopiowania. Czasy powyżej 24 godzin są obsługiwane.

Arkusz TAB jest pomijany. Odczytywane są zapisane wyniki formuł, bez ich wykonywania. Przed wgraniem przelicz i zapisz plik w Excelu. Nieaktualnych zapisanych wyników nie można automatycznie rozpoznać. Brakujący wynik lub błąd w czasie blokuje obliczenie. Pliki są przetwarzane w pamięci i nie są zapisywane na serwerze. Limit przesłania: 10 MB, rozpakowanej zawartości: 50 MB. Stary format XLS i pliki szyfrowane nie są obsługiwane.

## Testy

```sh
python -m unittest discover -s tests -v
```
