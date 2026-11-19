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
| Właściciel nie widzi, że sprawa wisi 2–3 dni | Zakładka **📊 Statystyki**: zaległe, „zawieszone” (po terminie i nikt nie ruszał ponad 48 godz.), na odpowiedź których lekarzy czekamy, kto ile zrobił w tygodniu. Codziennie o 20:30 właściciel dostaje **raport e-mail** |

Szybkie szablony w oknie „Nowa sprawa”: *„Ekstrakcja / implant dziś → telefon jutro”*, *„Wysłano skierowanie na tomografię”*, *„Wysłano do laboratorium”*: termin i tekst wstawiają się same.

## 📥 Skrzynka: wszystkie zapytania w jednym miejscu

| Źródło | Co tablica odczytuje | Co robi recepcja |
|---|---|---|
| 📅 **Booksy** | pacjent, telefon, e-mail, **termin, usługa (z ceną), do kogo** (pracownik) | „⚡ nowa rezerwacja — POTWIERDŹ” → potwierdź w Booksy → ✔ Załatwione |
| 🩺 **ZnanyLekarz** | pacjent, telefon, e-mail, **termin, usługa, do kogo** (jeśli w mailu jest „lek. dent. …”); rozpoznaje odwołania, nowe wizyty, zmiany i wiadomości | Odwołanie → „➕ Utwórz sprawę” (gotowa sprawa „zaproponować nowy termin”) |
| 🌐 **Formularz na stronie** („Twoje imię / Twój numer telefonu”) | imię, telefon (+ termin, usługa i lekarz, jeśli dodacie takie pola do formularza: „Usługa:”, „Termin:”, „Lekarz:”) | Oddzwonić |
| 💬 **Wiadomość ze strony** („New message from …”) | imię, telefon, treść | Oddzwonić |
| 📦 **WeTransfer** (pacjent przesyła RTG / tomografię) | nazwa przesyłki, e-mail pacjenta, link, **do kiedy link jest ważny**, rozmiar | „➕ Utwórz sprawę” → sprawa *Tomografia gotowa — pokazać lekarzowi*. Pobierz pliki, zanim link wygaśnie |
| ✉️ **Maile od pacjentów** (np. „Zapytanie o …”) | nadawca, telefon z treści, treść bez cytowanej historii | Odpisać / utworzyć sprawę. Newsletter? → „🚫 Ignoruj nadawcę” |
| 📱 **SMS-y** z telefonu gabinetu | numer, nazwa z kontaktów, treść | Odpisać / oddzwonić |
| 📞 **Nieodebrane połączenia** | numer, nazwa z kontaktów | Oddzwonić |

Przyciski przy każdej wiadomości:
- **🔗 Dodaj do sprawy**, jeśli pacjent (po telefonie lub nazwisku) ma już otwartą sprawę. Przykład: pacjent oddzwania po nieodebranym telefonie, a tablica od razu pokazuje: *„Ten pacjent ma otwartą sprawę: Telefon po zabiegu — jak się czuje po ekstrakcji 38”*.
- **➕ Utwórz sprawę**: formularz wypełnia się danymi z wiadomości.
- **✔ Załatwione** z krótką notatką, co zrobiono.

Pomijane automatycznie: maile Diagdent (tomografie mają własną automatyzację), potwierdzenia pobrania z WeTransfer, maile wysłane z adresów gabinetu, SMS-y serwisowe (Play, kody). Listę można zmienić w arkuszu „Ustawienia” → `inboxIgnore`.

Wyszukiwarka u góry przeszukuje też skrzynkę z ostatnich 14 dni.

**Nie trzeba odświeżać strony.** Otwarta tablica co minutę sama sprawdza pocztę i telefon (gdy otwiera ją konto recepcji). Nowa wiadomość pokazuje się sama, z krótkim powiadomieniem i dźwiękiem, a na karcie przeglądarki widać licznik, np. „(3) Gabinet…”. Obok „Na zmianie” jest godzina ostatniego odświeżenia (🔄 14:32). Gdy tablica jest zamknięta, poczta i tak jest sprawdzana co 5 minut w tle.

## Instalacja (ok. 20 minut, jednorazowo)

**Gdzie co jest:**
- **dentonia.pl@gmail.com** (właściciel) jest **właścicielem pliku tablicy**: zarządza nim, widzi historię zmian i może w każdej chwili odebrać dostęp. Tu przychodzą raporty.
- **PolishDent.pl@gmail.com** (recepcja) ma dostęp jako **Edytor**. Na tym koncie przychodzą wszystkie maile i to z tego konta tablica je czyta.

Czy plik na koncie właściciela zmienia czytanie poczty recepcji? **Nie**, pod jednym warunkiem: kroki 3–7 robi się **zalogowanym na PolishDent.pl**. Automaty (sprawdzanie poczty co 5 min, raporty) działają na koncie, które uruchomiło `setup`. Jeśli spróbuje go uruchomić inne konto, tablica pokaże błąd.

**Na koncie dentonia.pl (właściciel):**
1. Otwórz [sheets.new](https://sheets.new) i utwórz arkusz, np. „Tablica spraw — PolishDent”.
2. **Udostępnij** → dodaj `polishdent.pl@gmail.com` jako **Edytor** (odznacz „Powiadom” albo zostaw, bez znaczenia).

**Na koncie PolishDent.pl (recepcja):**

3. Otwórz ten arkusz (Dysk Google → „Udostępnione dla mnie”) → menu **Rozszerzenia → Apps Script**.
4. W pliku `Code.gs` wklej zawartość [`apps-script/Code.gs`](apps-script/Code.gs) (zastąp wszystko). Kliknij **＋ → HTML**, nazwij plik `Index` (dokładnie tak) i wklej [`apps-script/Index.html`](apps-script/Index.html).
5. ⚙️ **Ustawienia projektu** → **Strefa czasowa: (GMT+01:00) Warszawa**. Bez tego raporty o 20:30 przyjdą o złej godzinie.
6. 💾 **Zapisz** → na liście funkcji wybierz `setup` → **Uruchom**. Google poprosi o uprawnienia (arkusz, Gmail, wysyłanie maili, wyzwalacze): zatwierdź. Przy „Aplikacja niezweryfikowana” kliknij **Zaawansowane → Przejdź do projektu** (to Wasz własny skrypt).
7. **Wdróż → Nowe wdrożenie → Typ: Aplikacja internetowa**:
   - *Wykonaj jako*: **Użytkownik uzyskujący dostęp do aplikacji internetowej**
   - *Kto ma dostęp*: **Każdy, kto ma konto Google**
   - **Wdróż** → skopiuj link.

   Tablicę otworzy tylko ktoś, kto ma dostęp do arkusza (właściciel i konto recepcji). Komuś obcemu z linkiem pokaże się błąd braku dostępu. Dzięki temu tablica wie, kto ją otworzył: właściciel wchodzi **swoim kontem dentonia.pl**, a tylko on może oznaczać potrącenia jako rozliczone.

**Uzupełnij arkusze** (pojawią się po `setup`):
- **Ustawienia**: `admins` (domyślnie *Nadia N, Nadia R, Tosia, Właściciel*); `ownerEmail` = **dentonia.pl@gmail.com** (tam przychodzą tylko raport dzienny o 20:30 i miesięczny ostatniego dnia miesiąca o 20:30, godzina: `reportTime`); ewentualnie `staff`, czyli inni pracownicy do potrąceń (np. asystentki). Zapytanie `ktEmailQuery` jest już ustawione na maile Diagdent (`"Portal Lekarza" "Nowe badanie"`): wpisz je w wyszukiwarkę Gmaila i sprawdź, że znajduje tylko te maile.
- **Lekarze**: lista jest już wpisana: Krzysztof Dragun, Andrzej Kulpa, Stanisław Tsydzik, Ewa Winiarska, Piotr Wereski, Vasyl Kuz, Kornelia Borska, Artur oraz higienistki Oliwia i Antonina. Tosia (rejestracja) to ta sama osoba co higienistka Antonina: w kolumnie C Antoniny jest „Tosia”, więc w potrąceniach widnieje raz, a „pracownik: Tosia” z Booksy zostanie rozpoznany jako Antonina. **Wszystko można w każdej chwili dopisać lub poprawić ręcznie** w arkuszu. Tablica bierze zmiany przy następnym odświeżeniu, bez ponownego wdrażania.
  - Kolumna A: nazwa widoczna na tablicy.
  - Kolumna B (opcjonalnie): **prywatny numer WhatsApp lekarza** w formacie `48XXXXXXXXX`. Służy tylko do przycisku „🔔 Przypomnij lekarzowi”: otwiera czat z tym lekarzem z gotowym tekstem. Bez numeru przycisk tylko zapisuje, że przypomniano ustnie.
  - Kolumna C: **jak ta osoba jest zapisana w Booksy i ZnanyLekarz**, po przecinku (np. *Oliwia, Oliwia Kowalska*). Po tym tablica rozpoznaje, do kogo jest wizyta. „Dowolny lekarz” z Booksy pokaże się jako „Dowolny lekarz”.

**Otwieranie tablicy:** na komputerze w recepcji (przeglądarka zalogowana na PolishDent.pl) dodaj link do zakładek. Właściciel otwiera ten sam link w telefonie zalogowany na dentonia.pl, wtedy w „Na zmianie” sam ustawi się „Właściciel”. Można wybrać „Dodaj do ekranu głównego”, a pojawi się ikona jak aplikacja. Przy pierwszym otwarciu każde konto raz zatwierdza uprawnienia.

> Po zmianach w kodzie (na koncie PolishDent.pl): **Wdróż → Zarządzaj wdrożeniami → ✏️ → Wersja: nowa → Wdróż**. Link się nie zmienia.

Menu **🦷 Tablica spraw** w arkuszu pozwala ręcznie sprawdzić pocztę albo wysłać raport dzienny / miesięczny do właściciela, żeby to przetestować.

> **Tablica ≠ arkusz.** Arkusz Google to tylko „magazyn” danych. Z tablicą pracuje się przez **link aplikacji** (Wdróż → Zarządzaj wdrożeniami → „Adres URL aplikacji internetowej”, kończy się na `/exec`). Arkusza na co dzień nie trzeba otwierać.

> **Pierwsze uruchomienie:** maile Diagdent starsze niż `ktMaxAgeDays` (domyślnie 3 dni) nie tworzą spraw. Jeśli wcześniej powstało dużo spraw ze starych maili, użyj menu **🦷 Tablica spraw → 🧹 Porządki: zamknij stare tomografie z maili**. Zamyka tylko sprawy ze starych maili, których nikt nie ruszał; trafiają do „Archiwum”.

## SMS-y i nieodebrane połączenia z telefonu (Android, karta Play)

Darmowa aplikacja **MacroDroid — Device Automation** (Google Play, wydawca ArloSoft) przesyła SMS-y i nieodebrane połączenia mailem na `polishdent.pl@gmail.com`, a tablica je odczytuje. Konfiguruje się ją **raz**, potem działa **automatycznie w tle**: nikt nic nie klika. Warunki: telefon włączony, z internetem, a MacroDroid niewyłączony przez oszczędzanie baterii. W darmowej wersji można mieć 5 makr, potrzebne są 4.

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
4. **Makro 3 — „CRM Rozmowa”**
   - Wyzwalacz: **Rozmowy/SMS → Koniec rozmowy** → Dowolny numer.
   - Akcja: **Wyślij e-mail** → Do: `polishdent.pl@gmail.com`, Temat: `CRM ROZMOWA ` + (*Magiczny tekst*) **Numer telefonu** + spacja + **Nazwa dzwoniącego**.
5. **Makro 4 — „CRM Wychodzące”**
   - Wyzwalacz: **Rozmowy/SMS → Połączenie wychodzące** → Dowolny numer.
   - Akcja: **Wyślij e-mail** → Temat: `CRM WYCHODZĄCE ` + **Numer telefonu** + spacja + **Nazwa dzwoniącego**.

   „Koniec rozmowy” w MacroDroid nie mówi, czy rozmowa była przychodząca, czy wychodząca. Tablica rozpoznaje to sama: koniec rozmowy po „wychodzącym” na ten sam numer = koniec tego telefonu, każdy inny = odebrana rozmowa przychodząca. Każde makro: jeden wyzwalacz + jedna akcja, nazwa zaczynająca się od „CRM”.
6. W ustawieniach Androida wyłącz **optymalizację baterii** dla MacroDroid, inaczej telefon może go usypiać.
7. Test: wyślij SMS na numer gabinetu, zadzwoń bez odbierania, potem odbierz drugi telefon i sama zadzwoń na dowolny numer. Po maks. 5 min wszystko będzie w „📥 Skrzynce”.

Temat musi zaczynać się dokładnie od `CRM SMS`, `CRM NIEODEBRANE`, `CRM ROZMOWA` lub `CRM WYCHODZĄCE` (działa też `CRM ODEBRANE`). Po tym tablica je rozpoznaje.

**Co tablica robi z telefonami i SMS-ami:**
- Wszystko z telefonu gabinetu trafia do **„📞 Rejestr telefonu: rozmowy i SMS”** w „📥 Skrzynce” (dziś i wczoraj): nieodebrane, odebrane, wychodzące i SMS-y, z godziną, numerem i nazwą z kontaktów.
- **Kilka nieodebranych z jednego numeru** (zanim ktoś oddzwonił) = **jedna pozycja** „Dzwonił(a) 4×: 08:37, 09:01, 09:14, 09:32”. Opisuje się tylko rozmowę (oddzwonienie), raz.
- **Nieodebrane** → dodatkowo nowa wiadomość „⚡ oddzwonić”. Gdy recepcja **oddzwoni** albo pacjent **dodzwoni się ponownie**, zamyka się samo, a w rejestrze jest **jeden wiersz** *„Nieodebrane → oddzwoniono 14:32”* z przyciskiem **🏷 Opisz** (opis dotyczy rozmowy przy oddzwonieniu).
- **Każdą rozmowę opisujemy** przyciskiem **🏷 Opisz**: kto to był, **powód / wynik (można zaznaczyć kilka naraz)** i krótka notatka. Lista powodów jest w „Ustawieniach” → `callTags` (można dopisywać własne): *Umówiony / zapisał się, Zastanowi się, Oddzwonić, Za drogo, Szuka NFZ, Pomyłka (inna klinika), Pytanie o cenę, Promocja / akcja, Zmiana lub odwołanie wizyty, Pytanie o wizytę / wynik, Nie odebrał, Firma / dostawca / reklama, Inne*.
- Nieopisane rozmowy są na samej górze skrzynki (**„🏷 Rozmowy do opisania”**) i w liczniku zakładki, a przy przekazaniu zmiany tablica o nich przypomina.
- Zaznaczenie **„Oddzwonić”** lub **„Zastanowi się”** od razu włącza **📅 Przypomnij**: powstaje sprawa „Nowe zapytanie” z terminem (domyślnie jutro 10:00), a przy pacjencie z otwartą sprawą przesuwa się jej termin.
- Numer pacjenta z otwartą sprawą: rozmowa sama zapisuje się w historii sprawy („📤 Połączenie wychodzące o 14:32”), a opis z 🏷 też.
- SMS-y, formularze ze strony i maile pacjentów: „✔ Załatwione” otwiera ten sam opis z powodami.
- **📊 Statystyki** i **raport miesięczny**: ile rozmów przychodzących, wychodzących, nieodebranych (i ile oddzwoniono, średnio po ilu minutach), SMS-ów, rozmów bez opisu oraz **ile razy który powód** (np. „Za drogo” 14×, „Szuka NFZ” 9×), a także kto opisywał rozmowy.

## 💰 Potrącenia: reklamacje i koszty do pokrycia

Zakładka **💰 Potrącenia** zbiera wszystko, co na koniec miesiąca trzeba potrącić pracownikom: reklamacje lekarzy (poprawka pracy, koszt laboratorium) i pacjentów, za których lekarz pokrywa koszty (zwrot, rabat).

- **Dodawanie**: przycisk „➕ Dodaj potrącenie” albo „💰 Potrącenie” w karcie sprawy (pacjent i lekarz wpiszą się same). Pola: kogo dotyczy, kwota (np. `450,50`), data, kategoria, pacjent, opis.
- **Kto co widzi i może:** recepcja **widzi listę i sumy** (żeby mogła przypomnieć właścicielowi), dodaje wpisy i poprawia swoje. **✅ Potrącono** i usuwanie: tylko właściciel, zalogowany swoim kontem dentonia.pl (sprawdza to serwer, nie da się tego obejść, wybierając „Właściciel” na liście).
- **Widok:** wybór miesiąca, **podsumowanie na osobę** (razem / już potrącono / do potrącenia), lista wpisów z datą, pacjentem i opisem oraz osobna lista nierozliczonych z wcześniejszych miesięcy.
- Osoby do wyboru: lekarze i higienistki z arkusza „Lekarze” + pracownicy z `staff` + rejestratorki.
- Ostatniego dnia miesiąca podsumowanie potrąceń (na osobę i ze szczegółami) jest w raporcie miesięcznym na dentonia.pl.

Potrącenia leżą w arkuszu „Potrącenia” w pliku tablicy, czyli na koncie właściciela. Historię zmian (kto, kiedy, co zmienił) właściciel sprawdzi w **Plik → Historia wersji**. Opcjonalnie można trzymać je w osobnym pliku: link wpisz w „Ustawieniach” → `debtsSheetUrl` (plik musi być udostępniony dla polishdent.pl jako Edytor).

## 📧 Raporty na e-mail właściciela (dentonia.pl)

**Codziennie o 20:30 (raport dzienny):** ile spraw otwartych, zaległych i zawieszonych, lista zaległych i zawieszonych z ostatnim ruchem i osobą, pytania bez odpowiedzi.

**Ostatniego dnia miesiąca o 20:30 (raport miesięczny za ten miesiąc):**
- sprawy: nowe, zrobione, zamknięte bez realizacji, średni czas realizacji, podział na typy;
- **osobno listy pacjentów: 🦷 Tomografia i 🩻 RTG (razem), 📋 Plan leczenia, 📞 Telefony po zabiegu, 🧪 Laboratorium, ✅ Potwierdzenia wizyt, 📌 Bieżące zadania, 📣 Ogłoszenia FB**: pacjent, lekarz, o co chodzi, status (✅ umówiony / ⏳ etap / ✖️ bez realizacji, zaległe na czerwono), data utworzenia;
- skrzynka: ile wiadomości z każdego źródła (Booksy, ZnanyLekarz, strona, maile, SMS, telefony, WeTransfer), ile obsłużono i jak szybko;
- praca recepcji na osobę: działania, utworzone i zamknięte sprawy, nieodebrane telefony, przypomnienia lekarzom, odpowiedzi na pytania;
- ile razy recepcja musiała przypominać każdemu lekarzowi;
- pytania: zadane, odpowiedziane, średni czas odpowiedzi;
- 💰 potrącenia: suma na osobę (razem / do potrącenia) i lista szczegółowa.

Raporty można wysłać ręcznie z menu „🦷 Tablica spraw” w arkuszu (raport miesięczny wysłany ręcznie w inny dzień niż ostatni dotyczy poprzedniego miesiąca).

## Typy spraw

🙋 Nowe zapytanie · 📣 Ogłoszenia FB (pacjenci z reklam na Facebooku/Instagramie) · 🦷 Tomografia · 🩻 RTG · 📋 Plan leczenia · 📞 Telefon po zabiegu · 🧪 Laboratorium · ✅ Potwierdzenie wizyty · 🔁 Wizyta kontrolna · 📌 Bieżące zadania.

RTG przysłane przez WeTransfer → „➕ Utwórz sprawę” zakłada sprawę 🩻 RTG (lub 🦷 Tomografia, gdy w nazwie nie ma „rtg”) od razu na etapie „pokazać lekarzowi”.

## 💬 Komentarze i ❓ pytania właściciela

- **💬 Komentarz / termin**: każda recepcjonistka może dopisać do sprawy komentarz (i zmienić termin). Wszystko widać w historii karty z imieniem i godziną.
- **❓ Pytanie**: właściciel (lub ktokolwiek) wybiera na karcie „❓ Pytanie”, pisze np. *„Dlaczego p. Kowalska jeszcze nieumówiona?”* i wybiera adresata: cała recepcja albo konkretna osoba.
  - U recepcji pytanie pojawia się na samej górze zakładki „⏰ Teraz” (żółty blok z przyciskiem **✍️ Odpowiedz**), a na innych zakładkach jako baner.
  - Przed przekazaniem zmiany tablica przypomina o pytaniach bez odpowiedzi.
  - Gdy recepcja odpowie, odpowiedź jest pod pytaniem w historii karty i w „📊 Statystyki → Pytania i odpowiedzi” (z czasem odpowiedzi). Żeby nie zaśmiecać poczty, e-maili o odpowiedziach domyślnie **nie ma** (można włączyć: „Ustawienia” → `notifyAnswers` = `tak`).
  - Pytania bez odpowiedzi trafiają też do porannego raportu.

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
