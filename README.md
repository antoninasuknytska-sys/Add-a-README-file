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
| Zapytania z Booksy, ZnanyLekarz, strony, maili, SMS-ów i nieodebranych połączeń giną | Zakładka **📥 Skrzynka**: co 5 min wszystko trafia w jedno miejsce (szczegóły niżej). Wiadomość, na którą nikt nie zareagował przez 30 min, robi się czerwona |
| Właściciel nie widzi, że sprawa wisi 2–3 dni | Zakładka **📊 Statystyki**: zaległe, „zawieszone” (po terminie i nikt nie ruszał ponad 48 godz.), na odpowiedź których lekarzy czekamy, kto ile zrobił w tygodniu. Codziennie rano właściciel dostaje **raport e-mail** |

Szybkie szablony w oknie „Nowa sprawa”: *„Ekstrakcja / implant dziś → telefon jutro”*, *„Wysłano skierowanie na tomografię”*, *„Wysłano do laboratorium”*: termin i tekst wstawiają się same.

## 📥 Skrzynka: wszystkie zapytania w jednym miejscu

| Źródło | Co tablica odczytuje | Co robi recepcja |
|---|---|---|
| 📅 **Booksy** | pacjent, telefon, e-mail, termin, usługa, pracownik | „⚡ nowa rezerwacja — POTWIERDŹ” → potwierdź w Booksy → ✔ Załatwione |
| 🩺 **ZnanyLekarz** | pacjent, telefon, e-mail, termin, usługa; rozpoznaje odwołania, nowe wizyty, zmiany i wiadomości | Odwołanie → „➕ Utwórz sprawę” (gotowa sprawa „zaproponować nowy termin”) |
| 🌐 **Formularz na stronie** („Twoje imię / Twój numer telefonu”) | imię, telefon | Oddzwonić |
| 💬 **Wiadomość ze strony** („New message from …”) | imię, telefon, treść | Oddzwonić |
| ✉️ **Maile od pacjentów** (np. „Zapytanie o …”) | nadawca, telefon z treści, treść bez cytowanej historii | Odpisać / utworzyć sprawę. Newsletter? → „🚫 Ignoruj nadawcę” |
| 📱 **SMS-y** z telefonu gabinetu | numer, nazwa z kontaktów, treść | Odpisać / oddzwonić |
| 📞 **Nieodebrane połączenia** | numer, nazwa z kontaktów | Oddzwonić |

Przyciski przy każdej wiadomości:
- **🔗 Dodaj do sprawy**, jeśli pacjent (po telefonie lub nazwisku) ma już otwartą sprawę. Przykład: pacjent oddzwania po nieodebranym telefonie, a tablica od razu pokazuje: *„Ten pacjent ma otwartą sprawę: Telefon po zabiegu — jak się czuje po ekstrakcji 38”*.
- **➕ Utwórz sprawę**: formularz wypełnia się danymi z wiadomości.
- **✔ Załatwione** z krótką notatką, co zrobiono.

Pomijane automatycznie: maile WeTransfer i Diagdent (tomografie mają własną automatyzację), maile wysłane z adresów gabinetu, SMS-y serwisowe (Play, kody). Listę można zmienić w arkuszu „Ustawienia” → `inboxIgnore`.

Wyszukiwarka u góry przeszukuje też skrzynkę z ostatnich 14 dni.

## Instalacja (ok. 15 minut, jednorazowo)

Wszystko rób na koncie **PolishDent.pl@gmail.com**: tablica czyta pocztę tego konta.

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

9. **Przekierowanie z PolishDent.wro → PolishDent.pl.** Skrypt widzi tylko jedną skrzynkę, więc w Gmailu **PolishDent.wro@gmail.com**: ⚙️ **Ustawienia → Zobacz wszystkie ustawienia → Przekazywanie i POP/IMAP → Dodaj adres przekazywania** → `polishdent.pl@gmail.com`. Na skrzynkę PL przyjdzie kod potwierdzający: wpisz go, wybierz „Przekazuj kopię poczty przychodzącej do…” i „zachowaj kopię w Odebranych”. Na tablicy każda wiadomość ma znaczek 📧 PL lub 📧 WRO.

> Po zmianach w kodzie: **Wdróż → Zarządzaj wdrożeniami → ✏️ → Wersja: nowa → Wdróż**. Link się nie zmienia.

Menu **🦷 Tablica spraw** w arkuszu pozwala ręcznie sprawdzić maile z tomografią albo wysłać raport do właściciela, żeby to przetestować.

## SMS-y i nieodebrane połączenia z telefonu (Android, karta Play)

Darmowa aplikacja **MacroDroid** (Google Play) przesyła SMS-y i nieodebrane połączenia mailem na `polishdent.pl@gmail.com`, a tablica je odczytuje. W darmowej wersji można mieć 5 makr, potrzebne są 2.

1. Zainstaluj **MacroDroid** na telefonie gabinetu i daj mu wymagane uprawnienia (SMS, rejestr połączeń, kontakty, powiadomienia).
2. **Makro 1 — „SMS do tablicy”**
   - Wyzwalacz: **SMS odebrany** → Dowolny numer.
   - Akcja: **Wyślij e-mail** → Do: `polishdent.pl@gmail.com`
   - Temat: `CRM SMS ` + (przycisk **…** / *Magiczny tekst*) **numer nadawcy SMS** + spacja + **nazwa kontaktu SMS**
   - Treść: (*Magiczny tekst*) **treść SMS**
   - Wysyłka: przez konto Gmail (zaloguj się w MacroDroid kontem PolishDent.pl).
3. **Makro 2 — „Nieodebrane do tablicy”**
   - Wyzwalacz: **Połączenie nieodebrane** → Dowolny numer.
   - Akcja: **Wyślij e-mail** → Do: `polishdent.pl@gmail.com`, Temat: `CRM NIEODEBRANE ` + (*Magiczny tekst*) **numer dzwoniącego** + spacja + **nazwa kontaktu**. Treść może być pusta.
4. W ustawieniach Androida wyłącz **optymalizację baterii** dla MacroDroid, inaczej telefon może go usypiać.
5. Test: wyślij SMS na numer gabinetu i zadzwoń, nie odbierając. Po maks. 5 min obie wiadomości będą w zakładce „📥 Skrzynka”.

Temat musi zaczynać się dokładnie od `CRM SMS` lub `CRM NIEODEBRANE`. Po tym tablica je rozpoznaje.

## WhatsApp i Instagram

- **Na razie ręcznie:** „＋ Nowa sprawa” → kanał *WhatsApp* lub *Instagram*.
- **Instagram:** wiadomości z Instagrama, Facebooka i WhatsApp Business widać w jednym miejscu w darmowym **Meta Business Suite** (skrzynka odbiorcza) na komputerze w recepcji. Konto Instagram musi być firmowe.
- **WhatsApp → WhatsApp Business:** przy przejściu na tym samym telefonie aplikacja proponuje przeniesienie historii czatów. Przed zmianą zrób kopię zapasową (WhatsApp → Ustawienia → Czaty → Kopia zapasowa czatów) i traktuj przejście jako jednokierunkowe. Business daje etykiety czatów, szybkie odpowiedzi i automatyczne wiadomości poza godzinami pracy.
- **Pełna automatyzacja** (wiadomości z WhatsApp/Instagrama same w skrzynce) wymaga oficjalnego API Meta: weryfikacji firmy w Meta, płatnego dostawcy i konfiguracji. To możliwy następny etap.

## Zasady pracy (regulamin, 5 punktów)

1. **Wszystko, czego nie da się zrobić „tu i teraz”, staje się kartą.** Żadnych karteczek. Zapisanie sprawy zajmuje 15 sekund.
2. **„O co chodzi” piszemy tak, żeby zrozumiała koleżanka, która nic nie wie**: *kto, po co, co dalej*. Np.: „Telefon: jak się czuje po ekstrakcji 36 (wczoraj, dr Nowak)”.
3. **Każde działanie z pacjentem czy lekarzem zapisujemy przyciskiem**: 📵 nie odebrał, 🔔 przypomnienie dla lekarza, 💬 komentarz, ➡️ następny etap.
4. **Przed przekazaniem zmiany**: w zakładce „Teraz” nie ma czerwonych kart bez komentarza, a w „📥 Skrzynce” nie ma nieobsłużonych wiadomości → kliknij „📤 Przekazuję zmianę”.
5. **Początek zmiany**: przeczytaj baner → „✅ Przejmuję zmianę” → pracuj z zakładki „⏰ Teraz” od góry do dołu.

Właściciel raz dziennie zagląda do porannego maila lub zakładki „Statystyki” i pyta tylko o czerwone i „🧊 zawieszone”.

## Co gdzie jest zapisane

- Arkusz **Sprawy**: wszystkie sprawy. Można go filtrować i sortować jak zwykłą tabelę, ale **nie edytuj ręcznie kolumny „Historia (JSON)”**.
- Arkusz **Zmiany**: dziennik przekazań zmian.
- Arkusz **Skrzynka**: wszystkie odebrane wiadomości, kto i kiedy je obsłużył (statystyka czasu reakcji).
- Zamknięte sprawy starsze niż 30 dni nie są pokazywane na tablicy, ale zostają w arkuszu na zawsze.

## Jak zmienić typy spraw i etapy

Na początku `Code.gs` jest obiekt `TYPES`. Można tam zmienić nazwy etapów, po ilu godzinach sprawa staje się zaległa (`due`), albo dodać nowy typ według wzoru istniejących.

## Dla programisty: test lokalny

```bash
cd apps-script
node dev/server.js   # http://localhost:8787, prawdziwy Code.gs z imitacją Arkusza Google
```
`dev/seed.js` zawiera dane demo, a `dev/sample-mails.js` przykładowe maile (Booksy, ZnanyLekarz, strona, SMS, połączenie, Diagdent). Folderu `dev/` nie trzeba kopiować do Apps Script.
