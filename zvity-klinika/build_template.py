import random, datetime as dt, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule

OUT = sys.argv[1]
random.seed(7)

F = "Arial"
def font(**k): return Font(name=F, size=k.pop("size", 10), **k)
HDR = PatternFill("solid", fgColor="1F4E5A")
SUB = PatternFill("solid", fgColor="D9E7EA")
INP = PatternFill("solid", fgColor="FFF6C8")
CALC = PatternFill("solid", fgColor="EFEFEF")
TOT = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
ZL = '#,##0 "zł";-#,##0 "zł";"-"'
PCT = '0.0%;-0.0%;"-"'
NUM = '#,##0;-#,##0;"-"'
H1 = '0.0;-0.0;"-"'
MON = 'mm.yyyy'
DATE = 'dd.mm.yyyy'
N = 3001  # last data row on entry sheets

wb = Workbook()

def header(ws, row, labels, col=1, fill=HDR):
    for i, l in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=l)
        c.font = font(bold=True, color="FFFFFF" if fill is HDR else "000000")
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BOX

def widths(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v

# ---------------- Налаштування ----------------
st = wb.active
st.title = "Налаштування"
S = "'Налаштування'!"
st["A1"] = "НАЛАШТУВАННЯ"; st["A1"].font = font(bold=True, size=14)
rows = [
    ("Звітний місяць (будь-яка дата в місяці)", dt.date(2026, 10, 1), DATE, "Від цієї клітинки рахується весь Дашборд. Скрипт розсилки 15-го числа ставить сюди попередній місяць автоматично."),
    ("Кількість крісел", 2, "0", "Дані від власника"),
    ("Годин роботи клініки на день", 12, "0", "Графік 8:00–20:00 (дані від власника)"),
    ("Робочих субот у місяці", 0, "0", "Будні рахуються автоматично (пн–пт, без урахування свят). Якщо працюєте в суботу, вкажіть кількість субот."),
    ("Цільова рентабельність (чистий прибуток / виручка)", 0.20, PCT, "Орієнтир для приватної стоматології: 15–25%"),
    ("Цільове завантаження крісел", 0.80, PCT, "Орієнтир: 75–85%"),
]
for i, (lbl, val, fmt, note) in enumerate(rows):
    r = 3 + i
    st.cell(row=r, column=1, value=lbl).font = font()
    c = st.cell(row=r, column=2, value=val); c.fill = INP; c.font = font(color="0000FF"); c.number_format = fmt; c.border = BOX
    st.cell(row=r, column=3, value=note).font = font(italic=True, color="666666")
st["A10"] = "Попередній місяць (авто)"; st["A10"].font = font()
st["B10"] = "=DATE(YEAR(B3),MONTH(B3)-1,1)"; st["B10"].number_format = MON
st["A11"] = "Звітний місяць, 1-ше число (авто)"; st["A11"].font = font()
st["B11"] = "=DATE(YEAR(B3),MONTH(B3),1)"; st["B11"].number_format = MON
widths(st, {"A": 48, "B": 16, "C": 90})
MCUR, MPREV = f"{S}$B$11", f"{S}$B$10"

# ---------------- Довідники ----------------
dv = wb.create_sheet("Довідники")
D = "'Довідники'!"
dv["A1"] = "ДОВІДНИКИ: заповніть один раз, далі вони підтягуються у випадаючі списки"; dv["A1"].font = font(bold=True, size=12)
dv["A2"] = "Жовті клітинки можна змінювати. Імена та % нижче — приклад, замініть на реальні."; dv["A2"].font = font(italic=True, color="666666")

docs = [("Лікар 1", "Лікар", 0.35, "Zachowawcza, Endodoncja"),
        ("Лікар 2", "Лікар", 0.35, "Zachowawcza, Dziecięca"),
        ("Лікар 3", "Лікар", 0.40, "Endodoncja"),
        ("Лікар 4", "Лікар", 0.40, "Chirurgia, Implantologia"),
        ("Лікар 5", "Лікар", 0.35, "Protetyka"),
        ("Лікар 6", "Лікар", 0.40, "Ortodoncja"),
        ("Лікар 7", "Лікар", 0.30, "Zachowawcza, Protetyka"),
        ("Гігієністка 1", "Гігієністка", 0.25, "Profilaktyka"),
        ("Гігієністка 2", "Гігієністка", 0.25, "Profilaktyka")]
header(dv, 3, ["Лікар / гігієністка", "Роль", "% від (вартість − технік − спецпрепарати)", "Спеціалізація (довідково)"], 1)
for i in range(17):
    r = 4 + i
    vals = docs[i] if i < len(docs) else ("", "", None, "")
    for j, v in enumerate(vals):
        c = dv.cell(row=r, column=1 + j, value=v if v != "" else None)
        c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
        if j == 2: c.number_format = "0%"
dv["C3"].comment = Comment("Відсоток, який отримує лікар від (вартість послуги − технік − спецпрепарати). Значення-приклади, впишіть ваші реальні.", "Шаблон")

dirs = ["Zachowawcza", "Endodoncja", "Ortodoncja", "Chirurgia", "Implantologia", "Profilaktyka", "Protetyka", "Dziecięca"]
header(dv, 3, ["Напрям"], 6)
for i in range(12):
    c = dv.cell(row=4 + i, column=6, value=dirs[i] if i < len(dirs) else None); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")

pays = ["Karta", "Gotówka", "Przelew", "BLIK", "Raty PayU", "Mediraty", "Bon podarunkowy", "Przedpłata (wcześniej)"]
header(dv, 3, ["Спосіб оплати"], 8)
for i in range(10):
    c = dv.cell(row=4 + i, column=8, value=pays[i] if i < len(pays) else None); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")

srcs = [("Google", "Google Ads"), ("Instagram/Facebook", "Meta Ads (Instagram/Facebook)"), ("ZnanyLekarz", "ZnanyLekarz"),
        ("Booksy", "Booksy"), ("Voucher (spotkania)", "Vouchery i materiały promocyjne"), ("Polecenie (сарафан)", None),
        ("Pacjent lekarza", None), ("Z ulicy", None), ("Inne", None)]
header(dv, 3, ["Джерело пацієнта", "Рекламна категорія витрат для цього джерела"], 10)
for i in range(12):
    s = srcs[i] if i < len(srcs) else (None, None)
    for j in range(2):
        c = dv.cell(row=4 + i, column=10 + j, value=s[j]); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")

groups = [("Лабораторія (технік)", 0.08, 0.12), ("Спецпрепарати / імпланти", None, None),
          ("Лікарі і гігієністки", 0.25, 0.35), ("Персонал погодинно", 0.08, 0.12),
          ("Матеріали", 0.05, 0.08), ("Приміщення", 0.05, 0.10), ("Обладнання", 0.03, 0.06),
          ("Маркетинг", 0.03, 0.07), ("Адміністрування", 0.02, 0.05), ("Фінансові витрати", 0.0, 0.02),
          ("Податки і ZUS", 0.05, 0.12), ("Власник (виплата собі)", None, None)]
header(dv, 3, ["Група витрат", "Норма min, % виручки", "Норма max, % виручки"], 13)
for i, g in enumerate(groups):
    for j, v in enumerate(g):
        c = dv.cell(row=4 + i, column=13 + j, value=v); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
        if j: c.number_format = "0%"
dv["N3"].comment = Comment("Орієнтовні галузеві норми для приватної стоматології (емпіричні, не з офіційного джерела). Скоригуйте під себе.", "Шаблон")

cats = [("Materiały i narzędzia (zakupy)", "Матеріали"),
        ("Czynsz", "Приміщення"), ("Media (prąd, woda, ogrzewanie)", "Приміщення"), ("Sprzątanie / pranie", "Приміщення"),
        ("Leasing sprzętu", "Обладнання"), ("Serwis i naprawy sprzętu", "Обладнання"), ("Sterylizacja / odpady medyczne", "Обладнання"),
        ("Google Ads", "Маркетинг"), ("Meta Ads (Instagram/Facebook)", "Маркетинг"), ("ZnanyLekarz", "Маркетинг"),
        ("Booksy", "Маркетинг"), ("Vouchery i materiały promocyjne", "Маркетинг"), ("Inny marketing (strona, foto, SMM)", "Маркетинг"),
        ("Księgowość", "Адміністрування"), ("Oprogramowanie (Felgdent i in.)", "Адміністрування"), ("Telefon / internet", "Адміністрування"),
        ("Ubezpieczenie OC / mienia", "Адміністрування"), ("Szkolenia", "Адміністрування"), ("Biuro i drobne", "Адміністрування"),
        ("Prowizje (terminal, PayU, Mediraty, bank)", "Фінансові витрати"), ("Kredyty / odsetki", "Фінансові витрати"),
        ("Podatki (PIT/CIT, VAT)", "Податки і ZUS"), ("ZUS", "Податки і ZUS"),
        ("Wypłata właściciela", "Власник (виплата собі)")]
header(dv, 3, ["Категорія витрат", "Група"], 17)
for i in range(40):
    cv = cats[i] if i < len(cats) else (None, None)
    for j in range(2):
        c = dv.cell(row=4 + i, column=17 + j, value=cv[j]); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")

header(dv, 3, ["Статус візиту"], 20)
for i, s in enumerate(["Виконано", "Неявка", "Скасовано"]):
    dv.cell(row=4 + i, column=20, value=s).border = BOX
header(dv, 3, ["Так/Ні"], 22)
dv["V4"] = "Так"; dv["V5"] = "Ні"
widths(dv, {"A": 20, "B": 13, "C": 18, "D": 26, "E": 3, "F": 16, "G": 3, "H": 22, "I": 3, "J": 22, "K": 32, "L": 3,
            "M": 26, "N": 13, "O": 13, "P": 3, "Q": 40, "R": 24, "S": 3, "T": 14, "U": 3, "V": 8})

DOC_N = f"{D}$A$4:$A$20"; DOC_P = f"{D}$C$4:$C$20"
CAT_N = f"{D}$Q$4:$Q$43"; CAT_G = f"{D}$R$4:$R$43"
GRP_N = f"{D}$M$4:$M$15"

def add_list(ws, rng, src):
    v = DataValidation(type="list", formula1=src, allow_blank=True)
    v.error = "Виберіть значення зі списку (див. аркуш Довідники)"; v.showErrorMessage = True
    ws.add_data_validation(v); v.add(rng)

# ---------------- Візити ----------------
vs = wb.create_sheet("Візити")
V = "'Візити'!"
cols = ["Дата візиту", "Місяць (авто)", "Пацієнт (ID / прізвище)", "Новий пацієнт? (Так/Ні)", "Джерело пацієнта",
        "Лікар", "Напрям", "Послуга", "Вартість, zł", "Технік (лабораторія), zł", "Спецпрепарати / імплант, zł",
        "Спосіб оплати", "Статус", "Тривалість, год", "% лікаря (авто)", "Виплата лікарю (авто), zł", "Маржа клініки (авто), zł"]
header(vs, 1, cols)
vs.freeze_panes = "C2"
widths(vs, {"A": 12, "B": 10, "C": 18, "D": 10, "E": 20, "F": 15, "G": 15, "H": 34, "I": 11, "J": 11, "K": 12,
            "L": 18, "M": 11, "N": 9, "O": 9, "P": 13, "Q": 13})
vs.row_dimensions[1].height = 45

services = {
    "Zachowawcza": [("Wypełnienie kompozytowe", 350, 550, 0.75, 1.0, 0, 0, 70), ("Przegląd (bezpłatny voucher)", 0, 0, 0.5, 0.5, 0, 0, 10), ("Konsultacja", 150, 200, 0.5, 0.5, 0, 0, 20)],
    "Endodoncja": [("Leczenie kanałowe pod mikroskopem", 1200, 1900, 1.5, 2.0, 0, 120, 100)],
    "Ortodoncja": [("Wizyta kontrolna – aparat stały", 250, 350, 0.5, 0.5, 0, 0, 80), ("Założenie aparatu stałego", 3500, 5000, 1.5, 1.5, 900, 0, 20)],
    "Chirurgia": [("Ekstrakcja zęba", 300, 600, 0.5, 0.75, 0, 0, 70), ("Ekstrakcja ósemki", 800, 1200, 1.0, 1.0, 0, 50, 30)],
    "Implantologia": [("Implant (wszczepienie)", 4000, 5000, 1.5, 1.5, 0, 1100, 70), ("Konsultacja implantologiczna", 200, 250, 0.5, 0.5, 0, 0, 30)],
    "Profilaktyka": [("Higienizacja (skaling + piaskowanie)", 350, 450, 1.0, 1.0, 0, 0, 100)],
    "Protetyka": [("Korona pełnoceramiczna", 2200, 2800, 1.0, 1.0, 700, 0, 60), ("Proteza", 2500, 3500, 1.0, 1.0, 900, 0, 20), ("Wizyta protetyczna (wycisk/przymiarka)", 200, 300, 0.5, 0.5, 0, 0, 20)],
    "Dziecięca": [("Leczenie zęba mlecznego", 200, 350, 0.5, 0.5, 0, 0, 100)],
}
doc_dirs = {"Лікар 1": ["Zachowawcza", "Endodoncja"], "Лікар 2": ["Zachowawcza", "Dziecięca"], "Лікар 3": ["Endodoncja"],
            "Лікар 4": ["Chirurgia", "Implantologia"], "Лікар 5": ["Protetyka"], "Лікар 6": ["Ortodoncja"],
            "Лікар 7": ["Zachowawcza", "Protetyka"], "Гігієністка 1": ["Profilaktyka"], "Гігієністка 2": ["Profilaktyka"]}
doc_weight = {"Лікар 1": 16, "Лікар 2": 14, "Лікар 3": 7, "Лікар 4": 9, "Лікар 5": 8, "Лікар 6": 12, "Лікар 7": 9, "Гігієністка 1": 13, "Гігієністка 2": 12}
src_w = [("Google", 30), ("Instagram/Facebook", 12), ("ZnanyLekarz", 12), ("Booksy", 4), ("Voucher (spotkania)", 8),
         ("Polecenie (сарафан)", 18), ("Pacjent lekarza", 8), ("Z ulicy", 5), ("Inne", 3)]
pay_w = [("Karta", 45), ("BLIK", 15), ("Gotówka", 15), ("Przelew", 10), ("Raty PayU", 5), ("Mediraty", 4), ("Bon podarunkowy", 2), ("Przedpłata (wcześniej)", 4)]
def wchoice(pairs): return random.choices([p[0] for p in pairs], weights=[p[1] for p in pairs])[0]

visits = []
pid = 0
patients = []
for (y, m, nvis) in [(2026, 9, 215), (2026, 10, 225)]:
    d0 = dt.date(y, m, 1)
    days = []
    d = d0
    while d.month == m:
        if d.weekday() < 5: days.append(d)
        d += dt.timedelta(days=1)
    for _ in range(nvis):
        doc = wchoice(list(doc_weight.items()))
        direc = random.choice(doc_dirs[doc])
        svc = random.choices(services[direc], weights=[s[7] for s in services[direc]])[0]
        name, pmin, pmax, dmin, dmax, lab, prep, _w = svc
        new = random.random() < 0.24 or not patients
        if name.startswith("Przegląd"): new = True
        if new:
            pid += 1
            src = "Voucher (spotkania)" if name.startswith("Przegląd") else wchoice(src_w)
            p = (f"P{pid:04d}", src); patients.append(p)
        else:
            p = random.choice(patients)
        price = round(random.uniform(pmin, pmax) / 10) * 10
        dur = random.choice([dmin, dmax])
        r = random.random()
        status = "Виконано" if r < 0.92 else ("Неявка" if r < 0.97 else "Скасовано")
        visits.append([random.choice(days), p[0], "Так" if new else "Ні", p[1], doc, direc, name, price,
                       lab, prep, wchoice(pay_w) if status == "Виконано" and price > 0 else None, status, dur])
visits.sort(key=lambda x: x[0])

for i in range(2, N + 1):
    if i - 2 < len(visits):
        v = visits[i - 2]
        vals = [v[0], None, v[1], v[2], v[3], v[4], v[5], v[6], v[7], v[8] or None, v[9] or None, v[10], v[11], v[12]]
        for j, val in enumerate(vals):
            if j == 1: continue
            c = vs.cell(row=i, column=1 + j, value=val)
            if j == 0: c.number_format = DATE
    vs.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    vs.cell(row=i, column=15, value=f'=IF(F{i}="","",IFERROR(INDEX({DOC_P},MATCH(F{i},{DOC_N},0)),0))').number_format = "0%"
    vs.cell(row=i, column=16, value=f'=IF(M{i}="Виконано",MAX(0,I{i}-J{i}-K{i})*O{i},0)').number_format = ZL
    vs.cell(row=i, column=17, value=f'=IF(M{i}="Виконано",I{i}-J{i}-K{i}-P{i},0)').number_format = ZL
    for col in (9, 10, 11): vs.cell(row=i, column=col).number_format = ZL
    for col in (2, 15, 16, 17): vs.cell(row=i, column=col).fill = CALC
    if i > 2 and i - 2 >= len(visits): vs.cell(row=i, column=1).number_format = DATE
add_list(vs, f"D2:D{N}", f"{D}$V$4:$V$5")
add_list(vs, f"E2:E{N}", f"{D}$J$4:$J$15")
add_list(vs, f"F2:F{N}", DOC_N)
add_list(vs, f"G2:G{N}", f"{D}$F$4:$F$15")
add_list(vs, f"L2:L{N}", f"{D}$H$4:$H$13")
add_list(vs, f"M2:M{N}", f"{D}$T$4:$T$6")

# ---------------- Витрати ----------------
ex = wb.create_sheet("Витрати")
E = "'Витрати'!"
header(ex, 1, ["Дата", "Місяць (авто)", "Категорія витрат", "Група (авто)", "Опис / постачальник", "Сума, zł", "Спосіб оплати"])
ex.freeze_panes = "A2"
widths(ex, {"A": 12, "B": 10, "C": 40, "D": 22, "E": 34, "F": 13, "G": 16})
ex["I1"] = "ВАЖЛИВО: виплати лікарям і гігієністкам рахуються автоматично з аркуша «Візити», а погодинний персонал іде з аркуша «Персонал». Тут їх НЕ вносимо, інакше буде подвійний облік."
ex["I1"].font = font(bold=True, color="C00000"); ex["I1"].alignment = Alignment(wrap_text=True); ex.column_dimensions["I"].width = 60
demo_ex = [("Materiały i narzędzia (zakupy)", 9200, "Hurtownia"), ("Czynsz", 9000, "Wynajmujący"), ("Media (prąd, woda, ogrzewanie)", 1800, ""),
           ("Sprzątanie / pranie", 1200, ""), ("Leasing sprzętu", 6500, "Leasing unit + RTG"), ("Serwis i naprawy sprzętu", 800, ""),
           ("Sterylizacja / odpady medyczne", 600, ""), ("Google Ads", 6500, "Google"), ("Meta Ads (Instagram/Facebook)", 1500, "Meta"),
           ("ZnanyLekarz", 900, ""), ("Booksy", 300, ""), ("Vouchery i materiały promocyjne", 400, "Druk voucherów"),
           ("Księgowość", 1500, "Biuro rachunkowe"), ("Oprogramowanie (Felgdent i in.)", 450, "Felgdent"), ("Telefon / internet", 300, ""),
           ("Ubezpieczenie OC / mienia", 350, ""), ("Szkolenia", 500, ""), ("Prowizje (terminal, PayU, Mediraty, bank)", 950, ""),
           ("ZUS", 1800, ""), ("Podatki (PIT/CIT, VAT)", 4200, ""), ("Wypłata właściciela", 10000, "")]
r = 2
for m in (9, 10):
    for cat, amt, desc in demo_ex:
        amt2 = amt if amt in (9000, 6500, 1500, 450, 10000) else round(amt * random.uniform(0.85, 1.15) / 10) * 10
        if cat == "Google Ads": amt2 = 6200 if m == 9 else 6900
        ex.cell(row=r, column=1, value=dt.date(2026, m, 10)).number_format = DATE
        ex.cell(row=r, column=3, value=cat); ex.cell(row=r, column=5, value=desc or None)
        ex.cell(row=r, column=6, value=amt2); ex.cell(row=r, column=7, value="Przelew")
        r += 1
for i in range(2, N + 1):
    ex.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    ex.cell(row=i, column=4, value=f'=IF(C{i}="","",IFERROR(INDEX({CAT_G},MATCH(C{i},{CAT_N},0)),"⚠ немає в довіднику"))')
    ex.cell(row=i, column=6).number_format = ZL
    ex.cell(row=i, column=1).number_format = DATE
    for col in (2, 4): ex.cell(row=i, column=col).fill = CALC
add_list(ex, f"C2:C{N}", CAT_N)
add_list(ex, f"G2:G{N}", f"{D}$H$4:$H$13")

# ---------------- Персонал ----------------
pe = wb.create_sheet("Персонал")
P = "'Персонал'!"
header(pe, 1, ["Дата (будь-який день місяця)", "Місяць (авто)", "Працівник", "Роль", "Годин за місяць", "Ставка, zł/год", "Сума, zł (авто)"])
pe.freeze_panes = "A2"
widths(pe, {"A": 14, "B": 10, "C": 20, "D": 22, "E": 12, "F": 12, "G": 14})
pe.row_dimensions[1].height = 32
pe["I1"] = "Погодинний персонал на umowa zlecenie: асистентки та рецепція. Один рядок на людину на місяць. Якщо гігієністка також працює асистенткою чи адміністратором, її погодинні години вносимо сюди, а % за гігієну рахується з «Візитів»."
pe["I1"].font = font(italic=True, color="444444"); pe["I1"].alignment = Alignment(wrap_text=True); pe.column_dimensions["I"].width = 60
staff = [("Asystentka 1", "Asystentka", 160, 35), ("Asystentka 2", "Asystentka", 140, 35), ("Asystentka 3 (doraźnie)", "Asystentka", 40, 32),
         ("Recepcja 1", "Recepcja", 160, 32), ("Recepcja 2 / Higienistka 1", "Recepcja", 60, 32)]
r = 2
for m in (9, 10):
    for s in staff:
        pe.cell(row=r, column=1, value=dt.date(2026, m, 1)).number_format = DATE
        pe.cell(row=r, column=3, value=s[0]); pe.cell(row=r, column=4, value=s[1])
        pe.cell(row=r, column=5, value=s[2] + random.choice([-8, 0, 0, 8])); pe.cell(row=r, column=6, value=s[3])
        r += 1
for i in range(2, 501):
    pe.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    pe.cell(row=i, column=7, value=f'=IF(A{i}="","",E{i}*F{i})').number_format = ZL
    pe.cell(row=i, column=6).number_format = ZL
    pe.cell(row=i, column=1).number_format = DATE
    for col in (2, 7): pe.cell(row=i, column=col).fill = CALC

# --------------- formula helpers ---------------
def vr(col): return f"{V}${col}$2:${col}${N}"
VB, VD, VE, VF, VG, VI, VJ, VK, VL, VM, VN, VP, VQ = [vr(c) for c in "BDEFGIJKLMNPQ"]
EB, ED, EC, EF = f"{E}$B$2:$B${N}", f"{E}$D$2:$D${N}", f"{E}$C$2:$C${N}", f"{E}$F$2:$F${N}"
PB, PG = f"{P}$B$2:$B$500", f"{P}$G$2:$G$500"
OK = '"Виконано"'

def rev(m, extra=""): return f'SUMIFS({VI},{VB},{m},{VM},{OK}{extra})'
def lab(m, extra=""): return f'SUMIFS({VJ},{VB},{m},{VM},{OK}{extra})'
def prep(m, extra=""): return f'SUMIFS({VK},{VB},{m},{VM},{OK}{extra})'
def payout(m, extra=""): return f'SUMIFS({VP},{VB},{m},{VM},{OK}{extra})'
def marg(m, extra=""): return f'SUMIFS({VQ},{VB},{m},{VM},{OK}{extra})'
def hrs(m, extra=""): return f'SUMIFS({VN},{VB},{m},{VM},{OK}{extra})'
def cnt(m, extra=""): return f'COUNTIFS({VB},{m},{VM},{OK}{extra})'
def newp(m, extra=""): return f'COUNTIFS({VB},{m},{VM},{OK},{VD},"Так"{extra})'
def staff_cost(m): return f'SUMIFS({PG},{PB},{m})'
def grp(m, g): return f'SUMIFS({EF},{EB},{m},{ED},{g})'
def opex(m): return f'SUMIFS({EF},{EB},{m})'
def cap(m): return f'({S}$B$4*{S}$B$5*(NETWORKDAYS({m},DATE(YEAR({m}),MONTH({m})+1,0))+{S}$B$6))'

def put(ws, ref, val, fmt=None, bold=False, fill=None, color=None, border=True, align=None):
    c = ws[ref]; c.value = val
    c.font = font(bold=bold, color=color) if color else font(bold=bold)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if border: c.border = BOX
    if align: c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    return c

def section(ws, row, title, ncols=11):
    c = ws.cell(row=row, column=1, value=title)
    c.font = font(bold=True, size=12, color="1F4E5A")
    ws.cell(row=row, column=1).border = Border(bottom=Side(style="medium", color="1F4E5A"))
    for k in range(2, ncols + 1):
        ws.cell(row=row, column=k).border = Border(bottom=Side(style="medium", color="1F4E5A"))

# ---------------- Дашборд ----------------
db = wb.create_sheet("Дашборд")
widths(db, {"A": 34, "B": 14, "C": 14, "D": 14, "E": 13, "F": 13, "G": 14, "H": 14, "I": 14, "J": 14, "K": 13, "L": 11})
db.sheet_view.showGridLines = False
db["A1"] = "МІСЯЧНИЙ ЗВІТ КЛІНІКИ"; db["A1"].font = font(bold=True, size=16, color="1F4E5A")
db["A2"] = "Звітний місяць:"; db["A2"].font = font(bold=True)
db["B2"] = f"={MCUR}"; db["B2"].number_format = MON; db["B2"].font = font(bold=True, size=12)
db["C2"] = "порівняння з:"; db["C2"].font = font(italic=True, color="666666"); db["C2"].alignment = Alignment(horizontal="right")
db["D2"] = f"={MPREV}"; db["D2"].number_format = MON; db["D2"].font = font(italic=True, color="666666")
db["F2"] = "Місяць змінюється на аркуші «Налаштування»"; db["F2"].font = font(italic=True, color="999999", size=9)
M, Pm = "$B$2", "$D$2"

# KPI block rows 4-7
section(db, 4, "1. ГОЛОВНЕ ЗА 5 СЕКУНД")
kpi_lbl = ["Виручка", "Чистий прибуток", "Рентабельність", "Візитів (виконано)", "Нових пацієнтів", "Середній чек", "Завантаження крісел", "Неявки + скасування"]
header(db, 5, ["Показник"] + kpi_lbl, 1, SUB)
db.row_dimensions[5].height = 30
put(db, "A6", "Цей місяць", bold=True); put(db, "A7", "Минулий місяць"); put(db, "A8", "Зміна")
# will reference P&L cells (filled below) — rows defined now
PL0 = 11  # P&L header row
pl_rows = {}  # name -> row

def kpi_formulas(m, plcol):
    return [f"={rev(m)}",
            f"={plcol}{PL0 + 17}",
            f"=IFERROR({plcol}{PL0 + 17}/{plcol}{PL0 + 1},0)",
            f"={cnt(m)}",
            f"={newp(m)}",
            f"=IFERROR({rev(m)}/{cnt(m)},0)",
            f"=IFERROR({hrs(m)}/{cap(m)},0)",
            f'=IFERROR((COUNTIFS({VB},{m},{VM},"Неявка")+COUNTIFS({VB},{m},{VM},"Скасовано"))/COUNTIFS({VB},{m},{VM},"<>"),0)']
kfmt = [ZL, ZL, PCT, NUM, NUM, ZL, PCT, PCT]
for j, (fc, fp) in enumerate(zip(kpi_formulas(M, "B"), kpi_formulas(Pm, "C"))):
    col = chr(ord("B") + j)
    put(db, f"{col}6", fc, kfmt[j], bold=True, align="center").font = font(bold=True, size=12)
    put(db, f"{col}7", fp, kfmt[j], align="center").font = font(color="666666")
    if kfmt[j] == PCT:
        put(db, f"{col}8", f"={col}6-{col}7", '+0.0" п.п.";-0.0" п.п.";"-"', align="center")
    else:
        put(db, f"{col}8", f"=IFERROR({col}6/{col}7-1,0)", '+0.0%;-0.0%;"-"', align="center")
db.conditional_formatting.add("B8:F8", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000", bold=True)))
db.conditional_formatting.add("B8:F8", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="00703C", bold=True)))
db.conditional_formatting.add("H8", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="00703C", bold=True)))
db.conditional_formatting.add("I8", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="C00000", bold=True)))
db.conditional_formatting.add("D6", CellIsRule(operator="lessThan", formula=[f"{S}$B$7"], fill=PatternFill("solid", fgColor="F8CBAD")))
db.conditional_formatting.add("H6", CellIsRule(operator="lessThan", formula=[f"{S}$B$8"], fill=PatternFill("solid", fgColor="F8CBAD")))

# P&L
section(db, PL0 - 1, "2. ДЕ ГРОШІ: ПРИБУТКИ ТА ЗБИТКИ (P&L)")
header(db, PL0, ["Стаття", "Цей місяць", "Минулий місяць", "Зміна, zł", "% від виручки", "Норма, %", "Статус"], 1, SUB)
lines = [("Виручка (виконані послуги)", lambda m: rev(m), None, "rev"),
         ("− Лабораторія (технік)", lambda m: lab(m), "Лабораторія (технік)", "c"),
         ("− Спецпрепарати / імпланти", lambda m: prep(m), "Спецпрепарати / імпланти", "c"),
         ("− Виплати лікарям і гігієністкам", lambda m: payout(m), "Лікарі і гігієністки", "c"),
         ("МАРЖА ПІСЛЯ ПРЯМИХ ВИТРАТ", None, None, "gm"),
         ("− Персонал погодинно (асистентки, рецепція)", lambda m: staff_cost(m), "Персонал погодинно", "c")]
for g in ["Матеріали", "Приміщення", "Обладнання", "Маркетинг", "Адміністрування", "Фінансові витрати", "Податки і ZUS", "Власник (виплата собі)"]:
    lines.append((f"− {g}", (lambda gg: (lambda m: grp(m, f'"{gg}"')))(g), g, "c"))
lines.append(("− Інше / не розподілено", lambda m: f'{opex(m)}-SUMPRODUCT(SUMIFS({EF},{EB},{m},{ED},{GRP_N}))', None, "c"))
lines.append(("Усього витрат", None, None, "tc"))
lines.append(("ЧИСТИЙ ПРИБУТОК", None, None, "np"))
lines.append(("Рентабельність", None, None, "pct"))
r0 = PL0 + 1
cost_rows = []
for i, (lbl, fn, g, kind) in enumerate(lines):
    r = r0 + i
    pl_rows[kind if kind != "c" else lbl] = r
    bold = kind in ("rev", "gm", "np", "tc", "pct")
    fill = TOT if kind in ("gm", "np") else None
    put(db, f"A{r}", lbl, bold=bold, fill=fill)
    for col, m in (("B", M), ("C", Pm)):
        if kind in ("rev", "c"):
            f = f"={fn(m)}"
            if kind == "c": cost_rows.append(r) if col == "B" else None
        elif kind == "gm":
            f = f"={col}{r0}-{col}{r0+1}-{col}{r0+2}-{col}{r0+3}"
        elif kind == "tc":
            f = "=" + "+".join(f"{col}{x}" for x in cost_rows)
        elif kind == "np":
            f = f"={col}{r0}-{col}{r-1}"
        elif kind == "pct":
            f = f"=IFERROR({col}{r-1}/{col}{r0},0)"
        put(db, f"{col}{r}", f, PCT if kind == "pct" else ZL, bold=bold, fill=fill)
    if kind == "pct":
        put(db, f"D{r}", f"=B{r}-C{r}", '+0.0" п.п.";-0.0" п.п.";"-"', fill=fill)
        put(db, f"E{r}", None); put(db, f"F{r}", f'="ціль "&ROUND({S}$B$7*100,0)&"%"', align="center")
        put(db, f"G{r}", f'=IF(B{r}>={S}$B$7,"✓ ціль досягнута","⚠ нижче цілі")', align="center")
    else:
        put(db, f"D{r}", f"=B{r}-C{r}", ZL, fill=fill)
        put(db, f"E{r}", f"=IFERROR(B{r}/$B${r0},0)", PCT, fill=fill)
        if g and g not in ("Спецпрепарати / імпланти", "Власник (виплата собі)"):
            gi = f'MATCH("{g}",{GRP_N},0)'
            mn = f"INDEX({D}$N$4:$N$15,{gi})"; mx = f"INDEX({D}$O$4:$O$15,{gi})"
            put(db, f"F{r}", f'=IFERROR(ROUND({mn}*100,0)&"–"&ROUND({mx}*100,0)&"%","")', align="center")
            put(db, f"G{r}", f'=IFERROR(IF($B${r0}=0,"",IF(E{r}>{mx},"⚠ вище норми","✓")),"")', align="center")
        else:
            put(db, f"F{r}", None, fill=fill); put(db, f"G{r}", None, fill=fill)
NP_ROW = pl_rows["np"]
assert NP_ROW == PL0 + 17, NP_ROW
db.conditional_formatting.add(f"G{r0}:G{r0+len(lines)}", FormulaRule(formula=[f'LEFT(G{r0},1)="⚠"'], font=Font(color="C00000", bold=True), fill=PatternFill("solid", fgColor="FCE4D6")))
db.conditional_formatting.add(f"B{NP_ROW}:C{NP_ROW}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000", bold=True)))
PL_END = r0 + len(lines) - 1

# Doctors
DR0 = PL_END + 3
section(db, DR0 - 1, "3. ХТО З ЛІКАРІВ ВИГІДНИЙ (крісел лише 2, тому головний показник — прибуток клініки за годину крісла)", 12)
header(db, DR0, ["Лікар / гігієністка", "Візитів", "Нових пацієнтів", "Виручка", "Середній чек", "Годин у кріслі",
                 "Виручка за год", "Виплата лікарю", "Прибуток клініки*", "Прибуток клініки за год", "Неявки / скас.", "Ранг"], 1, SUB)
db.row_dimensions[DR0].height = 42
NDOC = 12
for i in range(NDOC):
    r = DR0 + 1 + i
    nm = f"$A{r}"
    x = f",{VF},{nm}"
    put(db, f"A{r}", f'=IF({D}A{4+i}="","",{D}A{4+i})')
    fs = [f"={cnt(M, x)}", f"={newp(M, x)}", f"={rev(M, x)}", f"=IFERROR(D{r}/B{r},0)", f"={hrs(M, x)}",
          f"=IFERROR(D{r}/F{r},0)", f"={payout(M, x)}", f"={marg(M, x)}", f'=IF(F{r}=0,"",I{r}/F{r})',
          f'=COUNTIFS({VB},{M},{VF},{nm},{VM},"Неявка")+COUNTIFS({VB},{M},{VF},{nm},{VM},"Скасовано")',
          f'=IF(J{r}="","",RANK(J{r},$J${DR0+1}:$J${DR0+NDOC}))']
    fm = [NUM, NUM, ZL, ZL, H1, ZL, ZL, ZL, ZL, NUM, "0"]
    for j, (f, fmt) in enumerate(zip(fs, fm)):
        col = chr(ord("B") + j)
        put(db, f"{col}{r}", f'=IF($A{r}="","",{f[1:]})', fmt, align="center" if col in "BCFKL" else None)
DRT = DR0 + NDOC + 1
put(db, f"A{DRT}", "РАЗОМ", bold=True, fill=TOT)
for col, fmt in zip("BCDEFGHIJKL", [NUM, NUM, ZL, ZL, H1, ZL, ZL, ZL, ZL, NUM, "0"]):
    rng = f"{col}{DR0+1}:{col}{DR0+NDOC}"
    f = {"E": f"=IFERROR(D{DRT}/B{DRT},0)", "G": f"=IFERROR(D{DRT}/F{DRT},0)", "J": f"=IFERROR(I{DRT}/F{DRT},0)", "L": None}.get(col, f"=SUM({rng})")
    put(db, f"{col}{DRT}", f, fmt, bold=True, fill=TOT)
db.conditional_formatting.add(f"J{DR0+1}:J{DR0+NDOC}", CellIsRule(operator="lessThan", formula=[f"$J${DRT}*0.8"], fill=PatternFill("solid", fgColor="FCE4D6")))
db.conditional_formatting.add(f"J{DR0+1}:J{DR0+NDOC}", CellIsRule(operator="greaterThan", formula=[f"$J${DRT}*1.2"], fill=PatternFill("solid", fgColor="E2EFDA")))
db[f"A{DRT+1}"] = "* Прибуток клініки = вартість − технік − спецпрепарати − виплата лікарю (до загальних витрат). Зелене: вище середнього на 20%+, червоне: нижче на 20%+."
db[f"A{DRT+1}"].font = font(italic=True, size=8, color="666666")

# Directions
DI0 = DRT + 4
section(db, DI0 - 1, "4. ЯКІ ПОСЛУГИ (НАПРЯМИ) ПРИНОСЯТЬ ГРОШІ")
header(db, DI0, ["Напрям", "Візитів", "Виручка", "Технік + препарати", "Виплата лікарям", "Маржа клініки", "Маржа, %",
                 "Частка виручки", "Годин у кріслі", "Маржа за год", "Ранг"], 1, SUB)
db.row_dimensions[DI0].height = 32
NDIR = 10
for i in range(NDIR):
    r = DI0 + 1 + i
    x = f",{VG},$A{r}"
    put(db, f"A{r}", f'=IF({D}F{4+i}="","",{D}F{4+i})')
    fs = [cnt(M, x), rev(M, x), f"{lab(M, x)}+{prep(M, x)}", payout(M, x), marg(M, x), f"IFERROR(F{r}/C{r},0)",
          f"IFERROR(C{r}/$C${DI0+NDIR+1},0)", hrs(M, x), f'IF(I{r}=0,"",F{r}/I{r})', f'IF(J{r}="","",RANK(J{r},$J${DI0+1}:$J${DI0+NDIR}))']
    fm = [NUM, ZL, ZL, ZL, ZL, PCT, PCT, H1, ZL, "0"]
    for j, (f, fmt) in enumerate(zip(fs, fm)):
        col = chr(ord("B") + j)
        put(db, f"{col}{r}", f'=IF($A{r}="","",{f})', fmt, align="center" if col in "BIK" else None)
DIT = DI0 + NDIR + 1
put(db, f"A{DIT}", "РАЗОМ", bold=True, fill=TOT)
for col, fmt in zip("BCDEFGHIJK", [NUM, ZL, ZL, ZL, ZL, PCT, PCT, H1, ZL, "0"]):
    rng = f"{col}{DI0+1}:{col}{DI0+NDIR}"
    f = {"G": f"=IFERROR(F{DIT}/C{DIT},0)", "H": f"=IFERROR(C{DIT}/C{DIT},0)", "J": f"=IFERROR(F{DIT}/I{DIT},0)", "K": None}.get(col, f"=SUM({rng})")
    put(db, f"{col}{DIT}", f, fmt, bold=True, fill=TOT)

# Sources
SO0 = DIT + 4
section(db, SO0 - 1, "5. ЗВІДКИ ПАЦІЄНТИ І ЧИ ОКУПАЄТЬСЯ РЕКЛАМА")
header(db, SO0, ["Джерело", "Нових пацієнтів", "Візитів (усі)", "Виручка (усі візити)", "Рекламні витрати", "Вартість 1 нового пацієнта",
                 "Виручка на 1 zł реклами"], 1, SUB)
db.row_dimensions[SO0].height = 32
NSRC = 10
for i in range(NSRC):
    r = SO0 + 1 + i
    x = f",{VE},$A{r}"
    put(db, f"A{r}", f'=IF({D}J{4+i}="","",{D}J{4+i})')
    adv = f'IF({D}K{4+i}="",0,SUMIFS({EF},{EB},{M},{EC},{D}K{4+i}))'
    fs = [newp(M, x), cnt(M, x), rev(M, x), adv, f'IF(OR(E{r}=0,B{r}=0),"",E{r}/B{r})', f'IF(E{r}=0,"",D{r}/E{r})']
    fm = [NUM, NUM, ZL, ZL, ZL, '0.0"x"']
    for j, (f, fmt) in enumerate(zip(fs, fm)):
        col = chr(ord("B") + j)
        put(db, f"{col}{r}", f'=IF($A{r}="","",{f})', fmt, align="center" if col in "BCG" else None)
SOT = SO0 + NSRC + 1
put(db, f"A{SOT}", "РАЗОМ", bold=True, fill=TOT)
for col, fmt in zip("BCDEFG", [NUM, NUM, ZL, ZL, ZL, '0.0"x"']):
    rng = f"{col}{SO0+1}:{col}{SO0+NSRC}"
    f = {"F": f'=IF(OR(E{SOT}=0,B{SOT}=0),"",E{SOT}/B{SOT})', "G": f'=IF(E{SOT}=0,"",D{SOT}/E{SOT})'}.get(col, f"=SUM({rng})")
    put(db, f"{col}{SOT}", f, fmt, bold=True, fill=TOT)
db[f"A{SOT+1}"] = "«Виручка на 1 zł реклами» нижче 3x — сигнал переглянути кампанію. Джерело потрібно вказувати для кожного візиту пацієнта (у Felgdent воно зберігається в картці пацієнта)."
db[f"A{SOT+1}"].font = font(italic=True, size=8, color="666666")

# Chairs & patients
CH0 = SOT + 4
section(db, CH0 - 1, "6. КРІСЛА ПРОСТОЮЮТЬ? ПАЦІЄНТИ ПОВЕРТАЮТЬСЯ?")
header(db, CH0, ["Показник", "Цей місяць", "Минулий місяць"], 1, SUB)
ch = [("Доступно годин крісел (крісла × години × робочі дні)", lambda m: cap(m), H1),
      ("Зайнято годин (виконані візити)", lambda m: hrs(m), H1),
      ("Завантаження крісел, %", lambda m: f"IFERROR({hrs(m)}/{cap(m)},0)", PCT),
      ("Вільних годин крісел", lambda m: f"{cap(m)}-{hrs(m)}", H1),
      ("Втрачено годин через неявки/скасування", lambda m: f'SUMIFS({VN},{VB},{m},{VM},"Неявка")+SUMIFS({VN},{VB},{m},{VM},"Скасовано")', H1),
      ("Втрачена виручка через неявки/скасування", lambda m: f'SUMIFS({VI},{VB},{m},{VM},"Неявка")+SUMIFS({VI},{VB},{m},{VM},"Скасовано")', ZL),
      ("ПОТЕНЦІАЛ: +маржа, якщо заповнити 25% вільних годин", lambda m: f"IFERROR(({cap(m)}-{hrs(m)})*0.25*{marg(m)}/{hrs(m)},0)", ZL),
      ("Візитів повторних пацієнтів", lambda m: f'COUNTIFS({VB},{m},{VM},{OK},{VD},"Ні")', NUM),
      ("Частка візитів повторних пацієнтів, %", lambda m: f'IFERROR(COUNTIFS({VB},{m},{VM},{OK},{VD},"Ні")/{cnt(m)},0)', PCT)]
for i, (lbl, fn, fmt) in enumerate(ch):
    r = CH0 + 1 + i
    put(db, f"A{r}", lbl, bold=lbl.startswith("ПОТЕНЦІАЛ"), fill=TOT if lbl.startswith("ПОТЕНЦІАЛ") else None)
    put(db, f"B{r}", f"={fn(M)}", fmt, align="center", bold=lbl.startswith("ПОТЕНЦІАЛ"), fill=TOT if lbl.startswith("ПОТЕНЦІАЛ") else None)
    put(db, f"C{r}", f"={fn(Pm)}", fmt, align="center")
for rr in range(CH0 + 1, CH0 + 1 + len(ch)):
    db.row_dimensions[rr].height = 26
    db[f"A{rr}"].alignment = Alignment(wrap_text=True, vertical="center")
CHE = CH0 + len(ch)

# Payments (next to chairs, cols E-G)
header(db, CH0, ["Спосіб оплати", "Сума", "Частка"], 5, SUB)
for i in range(10):
    r = CH0 + 1 + i
    put(db, f"E{r}", f'=IF({D}H{4+i}="","",{D}H{4+i})')
    put(db, f"F{r}", f'=IF(E{r}="","",SUMIFS({VI},{VB},{M},{VM},{OK},{VL},E{r}))', ZL)
    put(db, f"G{r}", f'=IF(E{r}="","",IFERROR(F{r}/$B${r0},0))', PCT)

# Hints
HI0 = max(CHE, CH0 + 10) + 3
section(db, HI0 - 1, "7. АВТОМАТИЧНІ ПІДКАЗКИ: ДЕ ОПТИМІЗУВАТИ")
dJ = f"$J${DR0+1}:$J${DR0+NDOC}"; dA = f"$A${DR0+1}:$A${DR0+NDOC}"
iJ = f"$J${DI0+1}:$J${DI0+NDIR}"; iA = f"$A${DI0+1}:$A${DI0+NDIR}"
hints = [
    f'=IFERROR("Найвигідніший лікар (прибуток/год крісла): "&INDEX({dA},MATCH(MAX({dJ}),{dJ},0))&" — "&TEXT(MAX({dJ}),"0")&" zł/год","")',
    f'=IFERROR("Найменш вигідний: "&INDEX({dA},MATCH(MIN({dJ}),{dJ},0))&" — "&TEXT(MIN({dJ}),"0")&" zł/год. Перевірте ціни послуг, витрату матеріалів і % виплати.","")',
    f'=IFERROR("Найвигідніший напрям (маржа/год): "&INDEX({iA},MATCH(MAX({iJ}),{iJ},0))&"; найменш вигідний: "&INDEX({iA},MATCH(MIN({iJ}),{iJ},0))&". Віддавайте кращі години крісла вигіднішим послугам.","")',
    f'=IF(H6<{S}$B$8,"Крісла завантажені на "&TEXT(H6*100,"0")&"% (ціль "&TEXT({S}$B$8*100,"0")&"%). Вільних годин: "&TEXT(B{CH0+4},"0")&". Потенціал +"&TEXT(B{CH0+7},"0")&" zł маржі.","Завантаження крісел в нормі.")',
    f'=IF(D6<{S}$B$7,"Рентабельність "&TEXT(D6*100,"0")&"% нижче цілі "&TEXT({S}$B$7*100,"0")&"%. Дивіться рядки з ⚠ у розділі 2.","Рентабельність в нормі.")',
    f'=IF(I6>0.08,"Неявок і скасувань "&TEXT(I6*100,"0")&"%: налаштуйте SMS-нагадування за 24 год і передоплату за довгі візити.","Неявки в нормі (до 8%).")',
]
for i, h in enumerate(hints):
    r = HI0 + i
    db.merge_cells(f"A{r}:K{r}")
    c = db[f"A{r}"]; c.value = h; c.font = font(); c.alignment = Alignment(wrap_text=True, vertical="center")
    db.row_dimensions[r].height = 22
LAST = HI0 + len(hints)

# chart: profit per hour by doctor
chA = BarChart(); chA.type = "bar"; chA.title = "Прибуток клініки за годину крісла, zł"
chA.add_data(Reference(db, min_col=10, min_row=DR0, max_row=DR0 + len(docs)), titles_from_data=True)
chA.set_categories(Reference(db, min_col=1, min_row=DR0 + 1, max_row=DR0 + len(docs)))
chA.legend = None; chA.roundedCorners = False; chA.height = 8; chA.width = 18
chA.x_axis.delete = False; chA.y_axis.delete = False; chA.title.overlay = False
db.add_chart(chA, f"A{LAST + 2}")
chB = BarChart(); chB.type = "bar"; chB.title = "Маржа за годину крісла по напрямах, zł"
chB.add_data(Reference(db, min_col=10, min_row=DI0, max_row=DI0 + len(dirs)), titles_from_data=True)
chB.set_categories(Reference(db, min_col=1, min_row=DI0 + 1, max_row=DI0 + len(dirs)))
chB.legend = None; chB.roundedCorners = False; chB.height = 8; chB.width = 16
chB.x_axis.delete = False; chB.y_axis.delete = False; chB.title.overlay = False
db.add_chart(chB, f"F{LAST + 2}")

db.print_area = f"A1:L{LAST + 19}"
db.page_setup.orientation = "portrait"; db.page_setup.paperSize = 9
db.page_setup.fitToWidth = 1; db.page_setup.fitToHeight = 0
db.sheet_properties.pageSetUpPr.fitToPage = True
db.print_options.horizontalCentered = True
db.page_margins.left = db.page_margins.right = 0.4

# ---------------- Рік ----------------
yr = wb.create_sheet("Рік")
widths(yr, {"A": 34, **{chr(ord("B") + i): 11 for i in range(13)}})
yr.sheet_view.showGridLines = False
yr["A1"] = "ДИНАМІКА ЗА РІК"; yr["A1"].font = font(bold=True, size=16, color="1F4E5A")
yr["A2"] = "Рік:"; yr["B2"] = f"=YEAR({MCUR})"; yr["B2"].font = font(bold=True, size=12)
header(yr, 4, ["Показник"] + [""] * 12 + ["Разом / сер."], 1, SUB)
for k in range(12):
    col = chr(ord("B") + k)
    c = yr[f"{col}4"]; c.value = f"=DATE($B$2,{k+1},1)"; c.number_format = MON
yl = [("Виручка", lambda m: rev(m), ZL, "sum"),
      ("Прямі витрати (технік, препарати, лікарі)", lambda m: f"{lab(m)}+{prep(m)}+{payout(m)}", ZL, "sum"),
      ("Персонал погодинно", lambda m: staff_cost(m), ZL, "sum"),
      ("Інші витрати (аркуш Витрати)", lambda m: opex(m), ZL, "sum"),
      ("Чистий прибуток", None, ZL, "sum"),
      ("Рентабельність", None, PCT, "avg"),
      ("Візитів (виконано)", lambda m: cnt(m), NUM, "sum"),
      ("Нових пацієнтів", lambda m: newp(m), NUM, "sum"),
      ("Середній чек", None, ZL, "avg"),
      ("Завантаження крісел", lambda m: f"IFERROR({hrs(m)}/{cap(m)},0)", PCT, "avg"),
      ("Маркетинг", lambda m: grp(m, '"Маркетинг"'), ZL, "sum")]
for i, (lbl, fn, fmt, agg) in enumerate(yl):
    r = 5 + i
    put(yr, f"A{r}", lbl, bold=lbl in ("Виручка", "Чистий прибуток"), fill=TOT if lbl == "Чистий прибуток" else None)
    for k in range(12):
        col = chr(ord("B") + k); m = f"{col}$4"
        if lbl == "Чистий прибуток": f = f"={col}5-{col}6-{col}7-{col}8"
        elif lbl == "Рентабельність": f = f"=IFERROR({col}9/{col}5,0)"
        elif lbl == "Середній чек": f = f"=IFERROR({col}5/{col}11,0)"
        else: f = f"={fn(m)}"
        put(yr, f"{col}{r}", f, fmt, fill=TOT if lbl == "Чистий прибуток" else None)
    if lbl == "Рентабельність": f = "=IFERROR(N9/N5,0)"
    elif lbl == "Середній чек": f = "=IFERROR(N5/N11,0)"
    elif agg == "sum": f = f"=SUM(B{r}:M{r})"
    else: f = f'=IFERROR(AVERAGEIF(B{r}:M{r},">0"),0)'
    put(yr, f"N{r}", f, fmt, bold=True, fill=TOT)
yr.conditional_formatting.add("B9:N9", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000", bold=True)))
lc = BarChart(); lc.title = "Виручка і чистий прибуток по місяцях, zł"
lc.add_data(Reference(yr, min_col=1, max_col=13, min_row=5), titles_from_data=True, from_rows=True)
lc.add_data(Reference(yr, min_col=1, max_col=13, min_row=9), titles_from_data=True, from_rows=True)
lc.set_categories(Reference(yr, min_col=2, max_col=13, min_row=4))
lc.roundedCorners = False; lc.legend.overlay = False; lc.title.overlay = False
lc.x_axis.delete = False; lc.y_axis.delete = False; lc.x_axis.number_format = MON
lc.height = 9; lc.width = 30
yr.add_chart(lc, "A18")
yr.print_area = "A1:N36"; yr.page_setup.orientation = "landscape"; yr.page_setup.fitToWidth = 1; yr.page_setup.fitToHeight = 0
yr.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- Інструкція ----------------
ins = wb.create_sheet("Інструкція", 0)
ins.sheet_view.showGridLines = False
ins.column_dimensions["A"].width = 4; ins.column_dimensions["B"].width = 120
txt = [
    ("ЯК КОРИСТУВАТИСЯ ЗВІТАМИ КЛІНІКИ", "h1"),
    ("Мета: щомісяця за 5 хвилин бачити, скільки клініка заробила, на що пішли гроші, які лікарі й послуги вигідні, чи простоюють крісла і чи окупається реклама.", ""),
    ("", ""),
    ("ЛЕГЕНДА КОЛЬОРІВ", "h2"),
    ("Жовті клітинки / сині цифри: тут вносимо дані.  Сірі колонки «(авто)»: формули, НЕ чіпати.  Зелені рядки: підсумки.", ""),
    ("", ""),
    ("ЩО ДЕ ЛЕЖИТЬ", "h2"),
    ("1. Налаштування: звітний місяць, кількість крісел, години роботи, цілі.", ""),
    ("2. Довідники: лікарі та їхні %, напрями, способи оплати, джерела пацієнтів, категорії витрат і норми. Заповнити один раз.", ""),
    ("3. Візити: один рядок = один візит (вивантаження з Felgdent). Звідси рахуються виручка, виплати лікарям, маржа, завантаження крісел.", ""),
    ("4. Витрати: усі рахунки клініки (оренда, лізинг, реклама, податки, матеріали...). БЕЗ виплат лікарям і погодинного персоналу.", ""),
    ("5. Персонал: погодинні асистентки та рецепція на umowa zlecenie (години × ставка). Один рядок на людину на місяць.", ""),
    ("6. Дашборд: місячний звіт (саме він іде в PDF на пошту).   7. Рік: динаміка по місяцях і графік.", ""),
    ("", ""),
    ("РЕГЛАМЕНТ (до 2 годин на тиждень)", "h2"),
    ("Щопонеділка (адміністратор, ~1 год): вивантажити з Felgdent візити за минулий тиждень і вставити на аркуш «Візити» (колонки A, C–N). Перевірити, що вказані лікар, напрям, статус, тривалість, технік / спецпрепарати і джерело пацієнта.", ""),
    ("Щопонеділка (власник, ~30 хв): внести рахунки й оплати за тиждень на аркуш «Витрати».", ""),
    ("До 10-го числа: внести години погодинного персоналу за минулий місяць («Персонал»), рахунок лабораторії звірити із сумою колонки «Технік», внести ZUS і податки.", ""),
    ("15-го о 21:00: скрипт Google автоматично ставить попередній місяць і надсилає Дашборд у PDF на пошту (див. нижче).", ""),
    ("", ""),
    ("ВАЖЛИВІ ПРАВИЛА ОБЛІКУ", "h2"),
    ("• Виручка = вартість ВИКОНАНИХ послуг у місяці візиту. Ваучер (bon podarunkowy) або передоплату за імплант у день продажу НЕ вносимо як виручку: виручка з'являється, коли послугу виконано.", ""),
    ("• Протетика в 2 етапи: вносимо суму на тому візиті, на якому її фактично розраховано з пацієнтом/лікарем (або ділимо на 2 рядки).", ""),
    ("• Raty PayU / Mediraty: у «Візитах» повна вартість, а комісію банку вносимо у «Витрати → Prowizje».", ""),
    ("• Технік і спецпрепарати / імплант під конкретного пацієнта: вносимо у відповідні колонки візиту. Тоді % лікаря рахується правильно, а маржа видно чесно.", ""),
    ("• Звичайні матеріали (композит, анестезія, рукавички): закупівлі вносимо у «Витрати → Materiały». Норма: 5–8% виручки.", ""),
    ("• Безкоштовний огляд за ваучером: вносимо візит з вартістю 0 і джерелом «Voucher (spotkania)». Так видно, скільки таких пацієнтів потім лікуються.", ""),
    ("• Свою зарплату власник вносить у «Витрати → Wypłata właściciela». Інакше прибуток виглядає більшим, ніж він є.", ""),
    ("", ""),
    ("ПЕРЕД ПОЧАТКОМ РОБОТИ", "h2"),
    ("Файл містить ДЕМО-дані (вересень–жовтень 2026), щоб було видно, як виглядає звіт. Щоб їх видалити, виділіть колонки вводу (Візити: A, C–N; Витрати: A, C, E–G; Персонал: A, C–F) і натисніть Delete. Цілі рядки не видаляйте, бо в них формули. Потім впишіть своїх лікарів і % в «Довідники».", ""),
    ("Місця під дані: Візити — 3000 рядків (≈ 1 рік), Витрати — 3000, Персонал — 500. На новий рік краще зробити копію файлу.", ""),
    ("", ""),
    ("АВТОМАТИЧНИЙ PDF НА ПОШТУ 15-ГО О 21:00", "h2"),
    ("1. Завантажте файл у Google Drive → Відкрити за допомогою Google Таблиць (Файл → Зберегти як Google Таблиці).", ""),
    ("2. Розширення → Apps Script → вставте код із файлу wyslij_raport.gs → впишіть свою пошту в рядку EMAIL → Зберегти.", ""),
    ("3. Виберіть функцію ustawHarmonogram → Запустити → дозволити доступ. Готово: щомісяця 15-го між 21:00 і 22:00 (Google запускає в межах години) прийде PDF Дашборду за попередній місяць.", ""),
    ("4. Для перевірки одразу запустіть функцію wyslijRaport.", ""),
]
for i, (t, kind) in enumerate(txt):
    c = ins.cell(row=1 + i, column=2, value=t)
    if kind == "h1": c.font = font(bold=True, size=16, color="1F4E5A")
    elif kind == "h2": c.font = font(bold=True, size=11, color="1F4E5A")
    else: c.font = font()
    c.alignment = Alignment(wrap_text=True, vertical="top")

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font and c.font.name != F:
                c.font = Font(name=F, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic, color=c.font.color)
wb.active = wb.sheetnames.index("Дашборд")
wb.save(OUT)
print("saved", OUT, "visits", len(visits))
