# Звіти стоматологічної клініки

| Файл | Що це |
|---|---|
| `Zvity_kliniki.xlsx` | Шаблон з демо-даними: звіти польською «Raport» і «Rok», аркуші даних, довідники. Кольори: зелений = добре, оранжевий = середньо, червоний = погано. |
| `apps-script/` | Скрипт для Google Таблиць: меню «🦷 Звіти клініки», завантаження файлів, імпорт з папки Drive, PDF (польською) і кольоровий лист 15-го о 21:00. |
| `build_template.py` | Збирає шаблон заново: `python3 build_template.py Zvity_kliniki.xlsx tests/defaults.json` |
| `tests/` | Тести логіки імпорту на зразках файлів (mBank, PKO, Google Ads, KSeF, Felgdent, Booksy). |

## Встановлення скрипта
1. Google Drive → завантажити `Zvity_kliniki.xlsx` → Файл → Зберегти як Google Таблиці.
2. Розширення → Apps Script → створити файли `Import.gs`, `Kod.gs`, `Raport.gs`, `Dialog.html` і вставити код з `apps-script/`.
3. Налаштування проєкту → «Показувати файл маніфесту» → вставити `appsscript.json`.
4. У `Raport.gs` вписати пошту (`EMAIL`), зберегти, оновити таблицю.
5. Меню «🦷 Звіти клініки» → «Увімкнути автоматику».

## Тести
```
node tests/test_import.js
python3 tests/export_wb.py Zvity_kliniki.xlsx /tmp/wb.json && node tests/test_kod_mock.js /tmp/wb.json
```
`test_kod_mock.js` перевіряє `Kod.gs` на імітації Google Таблиць, а не в справжньому Google.
