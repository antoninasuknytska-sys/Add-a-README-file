// Przykładowe maile (format jak na zrzutach ekranu, dane zmyślone)
module.exports = function (ctx) {
  const H = 36e5, now = Date.now();
  ctx.__inbox.push(
    { id: 'bk1', date: now - 50 * 6e4, from: '"Anna Rafalska" <no-reply@booksy.com>', subject: 'Anna Rafalska: nowa rezerwacja piątek, 27 listopada 2026 16:15',
      body: 'booksy\nTwój klient oczekuje na potwierdzenie rezerwacji!\nAnna Rafalska\n731 222 361\nanna.rafalska93@gmail.com\npiątek, 27 listopada 2026, 16:15 - 17:00\nPROFILAKTYKA: Pakiet higienizacji\n330,00 zł, 16:15 - 17:00\npracownik: Oliwia\nZarządzaj swoimi rezerwacjami' },
    { id: 'zl1', date: now - 3 * H, from: 'ZnanyLekarz <powiadomienia@znanylekarz.pl>', subject: '❌ Julia Kowalczyk odwołał(a) wizytę umówioną na sobota 14.11.2026 o 10:00',
      body: 'Julia Kowalczyk odwołał(a) wizytę\nPacjent\nJulia Kowalczyk (+48 888 350 204 julia.kowalczyk@gmail.com)\nSobota, 14 listopada 2026 o 10:00\nBadanie stomatologiczne/Przegląd stomatologiczny (30 min)\nlek. dent. Anna Nowak' },
    { id: 'www1', date: now - 20 * 6e4, from: 'Klinika stomatologiczna <wordpress@polishdent.pl>', subject: 'Submission Confirmation',
      body: 'Twoje imię:          Anna\nTwój numer telefonu: 888 123 449' },
    { id: 'www2', date: now - 10 * 6e4, from: 'Klinika stomatologiczna <wordpress@polishdent.pl>', subject: 'New message from Anna (Nowa wiadomość od Anny)', body: '-Anna ( 888 777 116 )' },
    { id: 'em1', date: now - 2 * H, from: 'Elena Nowak <elena.nowak@gmail.com>', subject: 'Zapytanie o koszt re-endo (ząb trzykanałowy)',
      body: 'Dzień dobry,\nWykonałam zdjęcie RTG, ile kosztuje re-endo zęba 26? Mój tel. 601 202 303\n\nOn Mon, 10 Nov 2026 at 10:00 Gabinet <polishdent.pl@gmail.com> wrote:\n> poprzednia wiadomość' },
    { id: 'wt1', date: now - 1 * H, from: 'WeTransfer <noreply@wetransfer.com>', subject: 'Marzena Kowal rtg sent successfully to marzena.kowal@gmail.com (Marzena Kowal rtg sent successfully to marzena.kowal@gmail.com)',
      body: 'Marzena Kowal rtg przesłana na adres marzena.kowal@gmail.com\n1 element, łącznie 137 MB · Wygasa 19 listopada 2026 r.\nDziękujemy za skorzystanie z WeTransfer.\nOdbiorcy marzena.kowal@gmail.com\nLink do pobrania https://we.tl/t-X8OR0DsHWuRSzX4B' },
    { id: 'wt2', date: now - 1 * H, from: 'WeTransfer <noreply@wetransfer.com>', subject: 'Your files were downloaded', body: 'pobrane' },
    { id: 'n1', date: now - 1 * H, from: 'ChatGPT <noreply@email.openai.com>', subject: '1 nowy styl obrazu do wypróbowania', body: 'Tapeta z naturą 108370056' },
    { id: 'n2', date: now - 1 * H, from: 'DentalTree <sklep@dentaltree.pl>', subject: '👉 Darmowy webinar', body: 'x', headers: { 'List-Unsubscribe': '<mailto:u@x>' } },
    { id: 'n3', date: now - 1 * H, from: 'FELG App <noreply@email.felg.pl>', subject: 'Alert bezpieczeństwa FELG', body: 'Dzień dobry' },
    { id: 'n4', date: now - 1 * H, from: 'WordPress <wordpress@polishdent.pl>', subject: '[Klinika stomatologiczna] Automatyczna aktualizacja', body: 'Cześć! Automatyczne' },
    { id: 'n5', date: now - 1 * H, from: 'Dentonia (przez Arkusze Google) <drive-shares-dm-noreply@google.com>', subject: 'Udostępniono Ci arkusz kalkulacyjny', body: 'x' },
    { id: 'wt3', date: now - 1 * H, from: 'WeTransfer <noreply@wetransfer.com>', subject: 'A transfer you sent is about to expire', body: 'Link: https://we.tl/t-abc' },
    { id: 'own', date: now - 1 * H, from: 'PolishDent <polishdent.pl@gmail.com>', subject: 'Re: Zapytanie', body: 'Odpowiedź gabinetu' },
    { id: 'sms1', date: now - 5 * 6e4, from: 'polishdent.pl@gmail.com', subject: 'CRM SMS +48671112233 Piotr Wójcik', body: 'Dzień dobry, czy mogę przesunąć wizytę na piątek?' },
    { id: 'sms2', date: now - 5 * 6e4, from: 'polishdent.pl@gmail.com', subject: 'CRM SMS PLAY', body: 'Twój pakiet internetu się kończy' },
    { id: 'cl1', date: now - 15 * 6e4, from: 'polishdent.pl@gmail.com', subject: 'CRM WYCHODZĄCE +48 501 234 567 Kowalska', body: '' },
    { id: 'cl2', date: now - 12 * 6e4, from: 'polishdent.pl@gmail.com', subject: 'CRM ODEBRANE +48 600 700 800 Nowy pacjent', body: '' },
    { id: 'call1', date: now - 40 * 6e4, from: 'polishdent.pl@gmail.com', subject: 'CRM NIEODEBRANE +48 671 112 233', body: '' },
    { id: 'kt1', kt: true, date: now - 30 * 6e4, from: 'Diagdent - Portal Lekarza <portal@diagdent.pl>', subject: 'Nowe badanie w systemie Portal Lekarza DIAGDENT',
      body: 'Na Państwa koncie w Portalu Lekarza pojawiło się już badanie Tomografia szczęki i żuchwy dla pacjenta KOWALSKA ANNA (02.10.1969).' }
  );
};
