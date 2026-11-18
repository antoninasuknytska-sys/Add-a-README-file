# 🦷 Tablica spraw recepcji: przekazanie zmiany bez gubienia pacjentów

Bezpłatna aplikacja internetowa dla gabinetu stomatologicznego, zbudowana na **Arkuszu Google + Google Apps Script**.
Wszystkie „ogony” recepcji są w jednym miejscu: telefony po zabiegach, tomografie, laboratorium, potwierdzenia wizyt, wizyty kontrolne.
Każda sprawa ma etap, termin i pełną historię, więc druga zmiana zawsze wie, co zrobiła pierwsza.

## Jak to rozwiązuje problemy

| Problem | Rozwiązanie |
|---|---|
| Karteczki, czat WhatsApp i dwie tabele w różnych miejscach | Jedna tablica. Każda sprawa to karta z **etapem**, **terminem** i **historią** („kto, kiedy, co zrobił”) |
| Tomografia: recepcja raz powiedziała lekarzowi i sprawa utknęła | Sprawa tomografii przechodzi etapy *Skierowanie wysłane → Tomografia gotowa → Czekamy na odpowiedź lekarza → Lekarz odpowiedział → Pacjent umówiony*. Na etapie „Czekamy na odpowiedź lekarza” karta po 3–4 godz. robi się **czerwona**. Przycisk **🔔 Przypomnij lekarzowi** otwiera WhatsApp lekarza z gotowym tekstem i liczy przypomnienia („przypomniano 3×”). Na etap „Lekarz odpowiedział” można przejść tylko po wpisaniu, **co dokładnie powiedział lekarz** |
| Mail z Diagdent „Nowe badanie w systemie Portal Lekarza” ginie w skrzynce | **Automatycznie** co 15 min skrypt sprawdza Gmail i odczytuje z maila **imię i nazwisko pacjenta, datę urodzenia i rodzaj badania**. Jeśli pacjent ma otwartą sprawę tomografii, sprawa sama przechodzi na „Tomografia gotowa — pokazać lekarzowi”. Jeśli nie ma, powstaje nowa sprawa z nazwiskiem z maila. Kolejność słów, wielkie litery i polskie znaki nie mają znaczenia |
| Nadia N dzwoniła, pacjent nie odebrał, Tosia nic o tym nie wie | Przycisk **📵 Nie odebrał** zapisuje próbę z godziną i imieniem i przypomina o oddzwonieniu za 3 godz. Gdy pacjent oddzwania, wystarczy wpisać w **wyszukiwarkę** nazwisko lub kilka cyfr telefonu, żeby od razu zobaczyć, kto, kiedy i **po co** dzwonił. Pole „O co chodzi” jest obowiązkowe |
| Przekazanie zmiany ustnie | Zakładka **🔄 Przekazanie zmiany**: „Przekazuję zmianę” → koleżanka widzi żółty baner i listę spraw, którymi zajmowała się poprzednia zmiana, i klika **„Przejmuję”**. Wszystko trafia do dziennika |
| Właściciel nie widzi, że sprawa wisi 2–3 dni | Zakładka **📊 Statystyki**: zaległe, „zawieszone” (po terminie i nikt nie ruszał ponad 48 godz.), na odpowiedź których lekarzy czekamy, kto ile zrobił w tygodniu. Codziennie rano właściciel dostaje **raport e-mail** |

Szybkie szablony w oknie „Nowa sprawa”: *„Ekstrakcja / implant dziś → telefon jutro”*, *„Wysłano skierowanie na tomografię”*, *„Wysłano do laboratorium”*: termin i tekst wstawiają się same.

## Instalacja (ok. 15 minut, jednorazowo)

Wszystko rób na **koncie Google gabinetu**, czyli tym, na które przychodzą maile z Diagdent.

1. Otwórz [sheets.new](https://sheets.new) i utwórz nowy arkusz, np. „Tablica spraw”.
2. Menu **Rozszerzenia → Apps Script**.
3. W pliku `Code.gs` wklej zawartość pliku [`apps-script/Code.gs`](apps-script/Code.gs) (zastąp wszystko, co tam jest).
4. Kliknij **＋ → HTML**, nazwij plik `Index` (dokładnie tak) i wklej zawartość [`apps-script/Index.html`](apps-script/Index.html).
5. Kliknij 💾 **Zapisz**. Na liście funkcji u góry wybierz `setup` → **Uruchom**. Google poprosi o uprawnienia (arkusz, Gmail, wysyłanie maili, wyzwalacze): zatwierdź. Jeśli pojawi się „Aplikacja niezweryfikowana”, kliknij **Zaawansowane → Przejdź do projektu** (to Twój własny skrypt).
6. Wróć do arkusza. Pojawią się arkusze:
   - **Ustawienia**: osoby `admins` (domyślnie *Nadia N, Nadia R, Tosia, Właściciel*), e-mail właściciela `ownerEmail`. Zapytanie `ktEmailQuery` jest już ustawione na maile Diagdent (`"Portal Lekarza" "Nowe badanie"`). Wpisz je w wyszukiwarkę Gmaila i sprawdź, że znajduje tylko te maile.
   - **Lekarze**: imiona lekarzy i numery WhatsApp w formacie `48XXXXXXXXX`.
7. W Apps Script: **Wdróż → Nowe wdrożenie → Typ: Aplikacja internetowa**.
   - *Wykonaj jako*: **Ja**
   - *Kto ma dostęp*: **Tylko ja** (najbezpieczniej: dane pacjentów widzi tylko osoba zalogowana na konto gabinetu)
   - **Wdróż** → skopiuj link.
8. Otwórz link na komputerze w recepcji i **dodaj do zakładek**. Właściciel otwiera ten sam link w telefonie, zalogowany w przeglądarce na konto gabinetu. Można wybrać „Dodaj do ekranu głównego”, a pojawi się ikona jak aplikacja.

> Po zmianach w kodzie: **Wdróż → Zarządzaj wdrożeniami → ✏️ → Wersja: nowa → Wdróż**. Link się nie zmienia.

Menu **🦷 Tablica spraw** w arkuszu pozwala ręcznie sprawdzić maile z tomografią albo wysłać raport do właściciela, żeby to przetestować.

## Zasady pracy (regulamin, 5 punktów)

1. **Wszystko, czego nie da się zrobić „tu i teraz”, staje się kartą.** Żadnych karteczek. Zapisanie sprawy zajmuje 15 sekund.
2. **„O co chodzi” piszemy tak, żeby zrozumiała koleżanka, która nic nie wie**: *kto, po co, co dalej*. Np.: „Telefon: jak się czuje po ekstrakcji 36 (wczoraj, dr Nowak)”.
3. **Każde działanie z pacjentem czy lekarzem zapisujemy przyciskiem**: 📵 nie odebrał, 🔔 przypomnienie dla lekarza, 💬 komentarz, ➡️ następny etap.
4. **Przed przekazaniem zmiany**: w zakładce „Teraz” nie ma czerwonych kart bez komentarza → kliknij „📤 Przekazuję zmianę”.
5. **Początek zmiany**: przeczytaj baner → „✅ Przejmuję zmianę” → pracuj z zakładki „⏰ Teraz” od góry do dołu.

Właściciel raz dziennie zagląda do porannego maila lub zakładki „Statystyki” i pyta tylko o czerwone i „🧊 zawieszone”.

## Co gdzie jest zapisane

- Arkusz **Sprawy**: wszystkie sprawy. Można go filtrować i sortować jak zwykłą tabelę, ale **nie edytuj ręcznie kolumny „Historia (JSON)”**.
- Arkusz **Zmiany**: dziennik przekazań zmian.
- Zamknięte sprawy starsze niż 30 dni nie są pokazywane na tablicy, ale zostają w arkuszu na zawsze.

## Jak zmienić typy spraw i etapy

Na początku `Code.gs` jest obiekt `TYPES`. Można tam zmienić nazwy etapów, po ilu godzinach sprawa staje się zaległa (`due`), albo dodać nowy typ według wzoru istniejących.

## Dla programisty: test lokalny

```bash
cd apps-script
node dev/server.js   # http://localhost:8787, prawdziwy Code.gs z imitacją Arkusza Google
```
`dev/seed.js` zawiera dane demo. Folderu `dev/` nie trzeba kopiować do Apps Script.
