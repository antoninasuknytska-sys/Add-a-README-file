"""Збирає шаблон Zvity_kliniki.xlsx (з демо-даними).  python3 build_template.py <out.xlsx> [defaults.json]"""
import random, datetime as dt, sys, json, unicodedata, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1]
DEFAULTS_JSON = sys.argv[2] if len(sys.argv) > 2 else None
random.seed(7)

F = "Arial"
def font(**k): return Font(name=F, size=k.pop("size", 10), **k)
HDR = PatternFill("solid", fgColor="1F4E5A")
SUB = PatternFill("solid", fgColor="D9E7EA")
INP = PatternFill("solid", fgColor="FFF7CC")   # блідо-жовтий: вписувати
CALC = PatternFill("solid", fgColor="E7E7E7")  # світло-сірий: рахується само
TOT = PatternFill("solid", fgColor="E7EEF0")
# світлофор
GOOD_F, GOOD_T = "9FD8A6", "0B4A17"
MID_F, MID_T = "FFC56E", "6B3A00"
BAD_F, BAD_T = "F28B82", "6E0B0B"
def fill(c): return PatternFill("solid", fgColor=c, bgColor=c)
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
ZL = '#,##0 "zł";-#,##0 "zł";"-"'
ZL2 = '#,##0.00 "zł";-#,##0.00 "zł";"-"'
PCT = '0.0%;-0.0%;"-"'
NUM = '#,##0;-#,##0;"-"'
H1 = '0.0;-0.0;"-"'
MON = 'mm.yyyy'
DATE = 'dd.mm.yyyy'
N = 3001

wb = Workbook()

def header(ws, row, labels, col=1, fl=HDR):
    for i, l in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=l)
        c.font = font(bold=True, color="FFFFFF" if fl is HDR else "000000")
        c.fill = fl
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BOX

def widths(ws, w):
    for k, v in w.items(): ws.column_dimensions[k].width = v

def add_list(ws, rng, src):
    v = DataValidation(type="list", formula1=src, allow_blank=True)
    v.error = "Виберіть значення зі списку"; v.showErrorMessage = True
    ws.add_data_validation(v); v.add(rng)

def put(ws, ref, val, fmt=None, bold=False, fl=None, color=None, border=True, align=None, size=10):
    c = ws[ref]; c.value = val
    c.font = font(bold=bold, color=color, size=size) if color else font(bold=bold, size=size)
    if fmt: c.number_format = fmt
    if fl: c.fill = fl
    if border: c.border = BOX
    if align: c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    return c

def section(ws, row, title, ncols=13):
    c = ws.cell(row=row, column=1, value=title)
    c.font = font(bold=True, size=12, color="1F4E5A")
    for k in range(1, ncols + 1):
        ws.cell(row=row, column=k).border = Border(bottom=Side(style="medium", color="1F4E5A"))

# --- умовне форматування «світлофор» ---
def cf_rules(ws, rng, good, mid, bad):
    for f, (bg, tx) in ((good, (GOOD_F, GOOD_T)), (mid, (MID_F, MID_T)), (bad, (BAD_F, BAD_T))):
        if f:
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f], fill=fill(bg), font=Font(color=tx, bold=True), stopIfTrue=True))

def cf_text(ws, rng, ref):
    """Колір за текстом оцінки в клітинці ref: «добре» / «середньо» / «погано»."""
    cf_rules(ws, rng, f'{ref}="добре"', f'{ref}="середньо"', f'{ref}="погано"')

def cf_num(ws, rng, x, g, m, higher=True):
    if higher:
        cf_rules(ws, rng, f"AND(ISNUMBER({x}),{x}>={g})", f"AND(ISNUMBER({x}),{x}<{g},{x}>={m})", f"AND(ISNUMBER({x}),{x}<{m})")
    else:
        cf_rules(ws, rng, f"AND(ISNUMBER({x}),{x}<={g})", f"AND(ISNUMBER({x}),{x}>{g},{x}<={m})", f"AND(ISNUMBER({x}),{x}>{m})")

def cf_zero_good(ws, rng, x, bad=True):
    """0 → зелений, >0 → червоний (або оранжевий)."""
    cf_rules(ws, rng, f"AND(ISNUMBER({x}),{x}=0)", None if bad else f"AND(ISNUMBER({x}),{x}>0)", f"AND(ISNUMBER({x}),{x}>0)" if bad else None)

def st_hi(x, g, m): return f'IF(NOT(ISNUMBER({x})),"",IF({x}>={g},"добре",IF({x}>={m},"середньо","погано")))'
def st_lo(x, g, m): return f'IF(NOT(ISNUMBER({x})),"",IF({x}<={g},"добре",IF({x}<={m},"середньо","погано")))'

def nrm(s):
    s = (s or "").lower().replace("ł", "l")
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).replace('"', "").replace("'", "").strip()

# =============== Налаштування ===============
st = wb.active; st.title = "Налаштування"
S = "'Налаштування'!"
st["A1"] = "НАЛАШТУВАННЯ"; st["A1"].font = font(bold=True, size=14)
for i, (lbl, val, fmt, note) in enumerate([
    ("Звітний місяць (будь-яка дата в місяці)", dt.date(2026, 10, 1), DATE, "Від цієї клітинки рахується весь звіт «Raport». Скрипт розсилки 15-го числа сам ставить сюди попередній місяць."),
    ("Кількість крісел", 2, "0", "Дані від власника"),
    ("Годин роботи клініки на день", 12, "0", "Графік 8:00–20:00 (дані від власника)"),
    ("Робочих субот у місяці", 0, "0", "Будні рахуються автоматично (пн–пт, без свят)."),
    ("Тривалість візиту за замовчуванням, год", 0.7, "0.00", "Якщо у файлі немає часу візиту і лікар невідомий. 0,7 год = 42 хв (середнє з Felgdent)."),
]):
    r = 3 + i
    st.cell(row=r, column=1, value=lbl).font = font()
    c = st.cell(row=r, column=2, value=val); c.fill = INP; c.font = font(color="0000FF"); c.number_format = fmt; c.border = BOX
    st.cell(row=r, column=5, value=note).font = font(italic=True, color="666666")
st["A8"] = "Звітний місяць, 1-ше число (авто)"; st["B8"] = "=DATE(YEAR(B3),MONTH(B3),1)"; st["B8"].number_format = MON
st["A9"] = "Попередній місяць (авто)"; st["B9"] = "=DATE(YEAR(B3),MONTH(B3)-1,1)"; st["B9"].number_format = MON
MCUR, MPREV = f"{S}$B$8", f"{S}$B$9"

st["A11"] = "ПОРОГИ КОЛЬОРІВ (зелений / оранжевий / червоний)"; st["A11"].font = font(bold=True, size=12)
header(st, 12, ["Показник", "ЗЕЛЕНИЙ (добре), якщо", "ОРАНЖЕВИЙ (середньо), якщо", "Напрям", "Пояснення"])
TH_DEF = [
    ("margin", "Рентабельність (чистий прибуток / виручка)", 0.20, 0.10, "hi", PCT, "Ціль для приватної стоматології 15–25%"),
    ("util", "Завантаження крісел", 0.75, 0.55, "hi", PCT, "Ціль 75–85%"),
    ("noshow", "Неявки + скасування", 0.05, 0.10, "lo", PCT, ""),
    ("repeat", "Частка візитів повторних пацієнтів", 0.60, 0.40, "hi", PCT, ""),
    ("roas", "Виручка на 1 zł реклами", 5, 3, "hi", '0.0"x"', ""),
    ("cac", "Вартість 1 нового пацієнта з реклами, zł", 300, 500, "lo", ZL, ""),
    ("chg", "Зміна до минулого місяця, % (виручка, прибуток, візити, чек)", 0.03, -0.03, "hi", PCT, "зелений: зростання понад 3%; оранжевий: майже без змін"),
    ("chgpp", "Зміна в процентних пунктах (рентабельність, завантаження)", 0.01, -0.01, "hi", PCT, "0,01 = 1 п.п."),
    ("rel", "Лікар / напрям: прибуток за год ÷ середнє по клініці", 1.10, 0.90, "hi", '0.00"x"', "1,10 = на 10% краще за середнє"),
    ("dirm", "Маржа напряму, %", 0.50, 0.35, "hi", PCT, ""),
    ("costx", "Витрати: частка виручки ÷ верхня норма", 1.0, 1.2, "lo", '0.00"x"', "зелений: у межах норми; оранжевий: до +20% понад норму"),
    ("recon", "Звірка банку з візитами (розбіжність, %)", 0.05, 0.15, "lo", PCT, "лабораторія, імпланти, зарплати, безготівкова виручка"),
]
TH_ROW = {}
for i, (k, lbl, g, m, d, fmt, note) in enumerate(TH_DEF):
    r = 13 + i; TH_ROW[k] = r
    st.cell(row=r, column=1, value=lbl).font = font()
    for col, v in ((2, g), (3, m)):
        c = st.cell(row=r, column=col, value=v); c.fill = INP; c.font = font(color="0000FF"); c.number_format = fmt; c.border = BOX
    st.cell(row=r, column=4, value="більше = краще" if d == "hi" else "менше = краще").font = font(color="666666")
    st.cell(row=r, column=5, value=note).font = font(italic=True, color="666666")
    st.cell(row=r, column=2).fill = fill(GOOD_F); st.cell(row=r, column=3).fill = fill(MID_F)
widths(st, {"A": 58, "B": 18, "C": 18, "D": 16, "E": 70})

def th_helper(ws, col0=16, row0=4):
    """Копія порогів на тому ж аркуші (Google Таблиці не дозволяють умовному форматуванню посилатися на інші аркуші)."""
    refs = {}
    cl, cg, cm = L(col0), L(col0 + 1), L(col0 + 2)
    ws[f"{cl}{row0-1}"] = "Пороги (копія з Налаштувань)"
    for i, (k, *_r) in enumerate(TH_DEF):
        r = row0 + i
        ws[f"{cl}{r}"] = k
        ws[f"{cg}{r}"] = f"={S}B{TH_ROW[k]}"
        ws[f"{cm}{r}"] = f"={S}C{TH_ROW[k]}"
        refs[k] = (f"${cg}${r}", f"${cm}${r}")
    for c in (cl, cg, cm): ws.column_dimensions[c].hidden = True
    return refs

# =============== Довідники ===============
dv = wb.create_sheet("Довідники")
D = "'Довідники'!"
dv["A1"] = "ДОВІДНИКИ: заповніть один раз, далі вони підтягуються у випадаючі списки та звіти"; dv["A1"].font = font(bold=True, size=12)
dv["A2"] = "Жовті клітинки можна змінювати. Імена та % — приклад. Нових лікарів і послуги з Felgdent скрипт додає сам (порожній % підсвічується червоним)."
dv["A2"].font = font(italic=True, color="666666")
docs = [("Lekarz 1", "Лікар", 0.35, 0.9), ("Lekarz 2", "Лікар", 0.35, 0.75), ("Lekarz 3", "Лікар", 0.40, 1.05), ("Lekarz 4", "Лікар", 0.40, 1.0),
        ("Lekarz 5", "Лікар", 0.35, 1.0), ("Lekarz 6", "Лікар", 0.40, 0.4), ("Lekarz 7", "Лікар", 0.30, 1.2),
        ("Higienistka 1", "Гігієністка", 0.20, 0.95), ("Higienistka 2", "Гігієністка", 0.20, 0.95)]
header(dv, 3, ["Лікар / гігієністка (як у Felgdent)", "Роль", "% від (вартість − технік − спецпрепарати)", "Сер. тривалість візиту, год (з Felgdent)"], 1)
for i in range(17):
    vals = docs[i] if i < len(docs) else (None, None, None, None)
    for j, v in enumerate(vals):
        c = dv.cell(row=4 + i, column=1 + j, value=v); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
        if j == 2: c.number_format = "0%"
        if j == 3: c.number_format = "0.00"
cf_rules(dv, "C4:C20", None, None, 'AND($A4<>"",$C4="")')
dirs = ["Zachowawcza", "Endodoncja", "Ortodoncja", "Chirurgia", "Implantologia", "Profilaktyka", "Protetyka", "Dziecięca"]
header(dv, 3, ["Напрям"], 6)
for i in range(12):
    c = dv.cell(row=4 + i, column=6, value=dirs[i] if i < len(dirs) else None); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
pays = ["Karta", "Gotówka", "Przelew", "BLIK", "Raty PayU", "Mediraty", "Bon podarunkowy", "Przedpłata (wcześniej)", "Czek", "Wpłata mieszana"]
NONCASH = ["Karta", "Przelew", "BLIK", "Raty PayU", "Mediraty"]
header(dv, 3, ["Спосіб оплати"], 8)
for i in range(10):
    c = dv.cell(row=4 + i, column=8, value=pays[i] if i < len(pays) else None); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
srcs = [("Google", "Google Ads"), ("Instagram/Facebook", "Meta Ads (Instagram/Facebook)"), ("Z sieci", "Inny marketing (strona, foto, SMM)"),
        ("ZnanyLekarz", "ZnanyLekarz"), ("Booksy", "Booksy"), ("Ulotka (promocja)", "Vouchery i materiały promocyjne"),
        ("Znajomy", None), ("Inny pacjent", None), ("Mailing", None), ("Z ulicy", None), ("Inne", None)]
header(dv, 3, ["Джерело пацієнта (як «Skąd pacjent» у Felgdent)", "Рекламна категорія витрат для джерела"], 10)
for i in range(16):
    s = srcs[i] if i < len(srcs) else (None, None)
    for j in range(2):
        c = dv.cell(row=4 + i, column=10 + j, value=s[j]); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
groups = [("Лабораторія (технік)", 0.08, 0.12), ("Спецпрепарати / імпланти", None, None), ("Лікарі і гігієністки", 0.25, 0.35),
          ("Персонал погодинно", 0.08, 0.12), ("Матеріали", 0.05, 0.08), ("Приміщення", 0.05, 0.10), ("Обладнання", 0.03, 0.06),
          ("Маркетинг", 0.03, 0.07), ("Адміністрування", 0.02, 0.05), ("Фінансові витрати", 0.0, 0.02),
          ("Податки і ZUS", 0.05, 0.12), ("Власник (виплата собі)", None, None)]
header(dv, 3, ["Група витрат (P&L)", "Норма min, % виручки", "Норма max, % виручки"], 13)
for i, gr in enumerate(groups):
    for j, v in enumerate(gr):
        c = dv.cell(row=4 + i, column=13 + j, value=v); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
        if j: c.number_format = "0%"
dv["N3"].comment = Comment("Орієнтовні галузеві норми (емпіричні, не з офіційного джерела). Скоригуйте під себе.", "Шаблон")
cats = [("Materiały i narzędzia (zakupy)", "Матеріали"),
        ("Czynsz", "Приміщення"), ("Media (prąd, woda, ogrzewanie)", "Приміщення"), ("Sprzątanie / pranie", "Приміщення"),
        ("Leasing sprzętu", "Обладнання"), ("Serwis i naprawy sprzętu", "Обладнання"), ("Sterylizacja / odpady medyczne", "Обладнання"),
        ("Google Ads", "Маркетинг"), ("Meta Ads (Instagram/Facebook)", "Маркетинг"), ("ZnanyLekarz", "Маркетинг"), ("Booksy", "Маркетинг"),
        ("Vouchery i materiały promocyjne", "Маркетинг"), ("Inny marketing (strona, foto, SMM)", "Маркетинг"),
        ("Księgowość", "Адміністрування"), ("Oprogramowanie (Felgdent i in.)", "Адміністрування"), ("Telefon / internet", "Адміністрування"),
        ("Ubezpieczenie OC / mienia", "Адміністрування"), ("Szkolenia", "Адміністрування"), ("Biuro i drobne", "Адміністрування"),
        ("Prowizje (terminal, PayU, Mediraty, bank)", "Фінансові витрати"), ("Kredyty / odsetki", "Фінансові витрати"),
        ("Podatki (PIT/CIT, VAT)", "Податки і ZUS"), ("ZUS", "Податки і ZUS"), ("Wypłata właściciela", "Власник (виплата собі)"),
        ("Technik / laboratorium (koszt z Felgdent)", "Лабораторія (технік)"), ("Inne koszty gabinetu (Felgdent)", "Адміністрування"),
        ("Laboratorium – płatność (звірка)", "Не рахувати: лабораторія"),
        ("Implanty / preparaty – płatność (звірка)", "Не рахувати: препарати"),
        ("Wypłaty lekarzy i personelu (звірка)", "Не рахувати: зарплати"),
        ("Przelew między własnymi rachunkami", "Не рахувати: перекази"),
        ("Надходження: термінал", "Надходження"), ("Надходження: PayU raty", "Надходження"), ("Надходження: Mediraty", "Надходження"),
        ("Надходження: BLIK / перекази пацієнтів", "Надходження"), ("Надходження: готівка на рахунок", "Надходження"),
        ("Надходження: інше", "Надходження")]
header(dv, 3, ["Категорія витрат / операцій", "Група"], 17)
for i in range(60):
    cv = cats[i] if i < len(cats) else (None, None)
    for j in range(2):
        c = dv.cell(row=4 + i, column=17 + j, value=cv[j]); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
dv["Q2"] = "«Не рахувати: …» — платежі, які вже пораховані з Візитів/Персоналу; вони йдуть лише у звірку, щоб не було подвійного обліку."
dv["Q2"].font = font(italic=True, size=8, color="666666")
header(dv, 3, ["Статус візиту"], 20)
for i, s in enumerate(["Виконано", "Неявка", "Скасовано"]): dv.cell(row=4 + i, column=20, value=s).border = BOX
header(dv, 3, ["Так/Ні"], 22); dv["V4"] = "Так"; dv["V5"] = "Ні"
header(dv, 3, ["Тип операції банку"], 24)
for i, s in enumerate(["витрата", "надходження", "будь-яка"]): dv.cell(row=4 + i, column=24, value=s).border = BOX
widths(dv, {"A": 22, "B": 13, "C": 18, "D": 26, "E": 3, "F": 16, "G": 3, "H": 22, "I": 3, "J": 22, "K": 32, "L": 3,
            "M": 26, "N": 12, "O": 12, "P": 3, "Q": 42, "R": 26, "S": 3, "T": 14, "U": 3, "V": 8, "W": 3, "X": 18})
DOC_N, DOC_P = f"{D}$A$4:$A$20", f"{D}$C$4:$C$20"
CAT_N, CAT_G = f"{D}$Q$4:$Q$63", f"{D}$R$4:$R$63"
GRP_N = f"{D}$M$4:$M$15"
CAT_GROUP = dict(cats)

# =============== Правила банку ===============
RULES = [
    ("GOOGLE", "витрата", "Google Ads"), ("FACEBK", "витрата", "Meta Ads (Instagram/Facebook)"), ("FACEBOOK", "витрата", "Meta Ads (Instagram/Facebook)"),
    ("META PLATFORMS", "витрата", "Meta Ads (Instagram/Facebook)"), ("INSTAGRAM", "витрата", "Meta Ads (Instagram/Facebook)"),
    ("ZNANYLEKARZ", "витрата", "ZnanyLekarz"), ("ZNANY LEKARZ", "витрата", "ZnanyLekarz"), ("DOCPLANNER", "витрата", "ZnanyLekarz"),
    ("BOOKSY", "витрата", "Booksy"), ("VOUCHER", "витрата", "Vouchery i materiały promocyjne"),
    ("ZUS", "витрата", "ZUS"), ("URZAD SKARBOWY", "витрата", "Podatki (PIT/CIT, VAT)"), ("MIKRORACHUNEK", "витрата", "Podatki (PIT/CIT, VAT)"),
    ("PIT-", "витрата", "Podatki (PIT/CIT, VAT)"), ("VAT-7", "витрата", "Podatki (PIT/CIT, VAT)"), ("CIT-8", "витрата", "Podatki (PIT/CIT, VAT)"),
    ("FELGDENT", "витрата", "Oprogramowanie (Felgdent i in.)"), ("LEASING", "витрата", "Leasing sprzętu"),
    ("PGE", "витрата", "Media (prąd, woda, ogrzewanie)"), ("TAURON", "витрата", "Media (prąd, woda, ogrzewanie)"), ("ENEA", "витрата", "Media (prąd, woda, ogrzewanie)"),
    ("ENERGA", "витрата", "Media (prąd, woda, ogrzewanie)"), ("E.ON", "витрата", "Media (prąd, woda, ogrzewanie)"), ("PGNIG", "витрата", "Media (prąd, woda, ogrzewanie)"),
    ("FORTUM", "витрата", "Media (prąd, woda, ogrzewanie)"), ("WODOCIAGI", "витрата", "Media (prąd, woda, ogrzewanie)"), ("MPWIK", "витрата", "Media (prąd, woda, ogrzewanie)"),
    ("VEOLIA", "витрата", "Media (prąd, woda, ogrzewanie)"),
    ("ORANGE", "витрата", "Telefon / internet"), ("T-MOBILE", "витрата", "Telefon / internet"), ("POLKOMTEL", "витрата", "Telefon / internet"),
    ("P4 SP", "витрата", "Telefon / internet"), ("NETIA", "витрата", "Telefon / internet"), ("VECTRA", "витрата", "Telefon / internet"),
    ("PZU", "витрата", "Ubezpieczenie OC / mienia"), ("WARTA", "витрата", "Ubezpieczenie OC / mienia"), ("ALLIANZ", "витрата", "Ubezpieczenie OC / mienia"),
    ("HESTIA", "витрата", "Ubezpieczenie OC / mienia"), ("GENERALI", "витрата", "Ubezpieczenie OC / mienia"),
    ("CZYNSZ", "витрата", "Czynsz"), ("NAJEM", "витрата", "Czynsz"), ("SPRZATANIE", "витрата", "Sprzątanie / pranie"), ("PRALNIA", "витрата", "Sprzątanie / pranie"),
    ("KSIEGOW", "витрата", "Księgowość"), ("BIURO RACHUNKOWE", "витрата", "Księgowość"),
    ("REMONDIS", "витрата", "Sterylizacja / odpady medyczne"), ("ODPADY", "витрата", "Sterylizacja / odpady medyczne"),
    ("UTYLIZ", "витрата", "Sterylizacja / odpady medyczne"), ("STERYLIZ", "витрата", "Sterylizacja / odpady medyczne"),
    ("SERWIS", "витрата", "Serwis i naprawy sprzętu"), ("SZKOLENIE", "витрата", "Szkolenia"), ("KURS", "витрата", "Szkolenia"),
    ("HENRY SCHEIN", "витрата", "Materiały i narzędzia (zakupy)"), ("SIGMA DENTAL", "витрата", "Materiały i narzędzia (zakupy)"),
    ("PRO-DENT", "витрата", "Materiały i narzędzia (zakupy)"), ("DENTAL PARTNER", "витрата", "Materiały i narzędzia (zakupy)"),
    ("MEDENT", "витрата", "Materiały i narzędzia (zakupy)"), ("KOL-DENT", "витрата", "Materiały i narzędzia (zakupy)"),
    ("ARTDENT", "витрата", "Materiały i narzędzia (zakupy)"), ("DENTSPLY", "витрата", "Materiały i narzędzia (zakupy)"),
    ("IVOCLAR", "витрата", "Materiały i narzędzia (zakupy)"), ("HURTOWNIA", "витрата", "Materiały i narzędzia (zakupy)"),
    ("STRAUMANN", "витрата", "Implanty / preparaty – płatność (звірка)"), ("NOBEL BIOCARE", "витрата", "Implanty / preparaty – płatność (звірка)"),
    ("OSSTEM", "витрата", "Implanty / preparaty – płatność (звірка)"), ("MEGAGEN", "витрата", "Implanty / preparaty – płatność (звірка)"),
    ("NEOSS", "витрата", "Implanty / preparaty – płatność (звірка)"), ("BIOHORIZONS", "витрата", "Implanty / preparaty – płatność (звірка)"),
    ("LABORATORIUM", "витрата", "Laboratorium – płatność (звірка)"), ("PROTETYK", "витрата", "Laboratorium – płatność (звірка)"),
    ("PRACOWNIA PROTET", "витрата", "Laboratorium – płatność (звірка)"),
    ("WYNAGRODZENIE", "витрата", "Wypłaty lekarzy i personelu (звірка)"), ("UMOWA ZLECENIE", "витрата", "Wypłaty lekarzy i personelu (звірка)"),
    ("WYPLATA WLASCICIELA", "витрата", "Wypłata właściciela"),
    ("PRZELEW WLASNY", "будь-яка", "Przelew między własnymi rachunkami"), ("PRZELEW MIEDZY", "будь-яка", "Przelew między własnymi rachunkami"),
    ("RATA KREDYTU", "витрата", "Kredyty / odsetki"), ("ODSETKI", "витрата", "Kredyty / odsetki"),
    ("PROWIZJA", "витрата", "Prowizje (terminal, PayU, Mediraty, bank)"), ("OPLATA ZA", "витрата", "Prowizje (terminal, PayU, Mediraty, bank)"),
    ("PAYU", "витрата", "Prowizje (terminal, PayU, Mediraty, bank)"),
    ("PAYU", "надходження", "Надходження: PayU raty"), ("MEDIRATY", "надходження", "Надходження: Mediraty"), ("MEDI RATY", "надходження", "Надходження: Mediraty"),
    ("ELAVON", "надходження", "Надходження: термінал"), ("PAYTEL", "надходження", "Надходження: термінал"), ("FISERV", "надходження", "Надходження: термінал"),
    ("POLCARD", "надходження", "Надходження: термінал"), ("GLOBAL PAYMENTS", "надходження", "Надходження: термінал"), ("ESERVICE", "надходження", "Надходження: термінал"),
    ("SUMUP", "надходження", "Надходження: термінал"), ("TERMINAL", "надходження", "Надходження: термінал"), ("ROZLICZENIE KART", "надходження", "Надходження: термінал"),
    ("WPLATA WLASNA", "надходження", "Надходження: готівка на рахунок"), ("WPLATA GOTOWKI", "надходження", "Надходження: готівка на рахунок"),
    ("WPLATOMAT", "надходження", "Надходження: готівка на рахунок"),
    ("BLIK", "надходження", "Надходження: BLIK / перекази пацієнтів"), ("LECZENIE", "надходження", "Надходження: BLIK / перекази пацієнтів"),
    ("WIZYTA", "надходження", "Надходження: BLIK / перекази пацієнтів"), ("ZAPLATA ZA", "надходження", "Надходження: BLIK / перекази пацієнтів"),
]
rb = wb.create_sheet("Правила банку")
rb["A1"] = "ПРАВИЛА АВТОКАТЕГОРІЙ ДЛЯ ВИПИСКИ З БАНКУ"; rb["A1"].font = font(bold=True, size=12)
rb["A2"] = ("Скрипт шукає ключове слово в полях «Контрагент» та «Опис» (без різниці великі/малі літери та польські літери). Спрацьовує ПЕРШЕ правило зверху. "
            "Додайте своїх постачальників, лікарів (ім'я → «Wypłaty lekarzy…»), лабораторію. Після змін: меню 🦷 → «Перерахувати категорії банку».")
rb["A2"].font = font(italic=True, color="444444"); rb["A2"].alignment = Alignment(wrap_text=True); rb.merge_cells("A2:D2"); rb.row_dimensions[2].height = 42
header(rb, 3, ["Ключове слово", "Тип операції", "Категорія"])
for i in range(300):
    rr = RULES[i] if i < len(RULES) else (None, None, None)
    for j, v in enumerate(rr):
        c = rb.cell(row=4 + i, column=1 + j, value=v); c.fill = INP; c.border = BOX; c.font = font(color="0000FF")
add_list(rb, "B4:B303", f"{D}$X$4:$X$6"); add_list(rb, "C4:C303", CAT_N)
widths(rb, {"A": 26, "B": 16, "C": 44, "D": 40})

# =============== Імпорт-налаштування ===============
FIELDS = {
 "BANK": [("date", "так", "data operacji; data transakcji; data ksiegowania; data; booking date; transaction date; date; дата"),
          ("amount", "один із", "kwota; kwota operacji; kwota transakcji; amount; сума"),
          ("debit", "один із", "obciazenia; obciazenie; wydatki; kwota obciazenia; debit; winien"),
          ("credit", "один із", "uznania; uznanie; wplywy; kwota uznania; credit; ma"),
          ("counterparty", "ні", "nadawca/odbiorca; nadawca / odbiorca; odbiorca/nadawca; kontrahent; dane kontrahenta; nazwa kontrahenta; odbiorca; nadawca; counterparty; beneficiary"),
          ("description", "так", "opis operacji; opis transakcji; tytul operacji; tytul; opis; szczegoly; description; title; details"),
          ("account", "ні", "numer rachunku; rachunek; konto; account")],
 "INVOICES": [("number", "так", "numer faktury; nr faktury; numer dokumentu; nr dokumentu; numer; invoice number"),
          ("date", "так", "data wystawienia; data faktury; data dokumentu; data; issue date"),
          ("seller", "ні", "sprzedawca; nazwa sprzedawcy; dostawca; wystawca; kontrahent; nazwa kontrahenta; seller"),
          ("nip", "ні", "nip sprzedawcy; nip; nip kontrahenta"),
          ("gross", "так", "kwota brutto; brutto; wartosc brutto; razem brutto; do zaplaty; kwota; gross"),
          ("due", "ні", "termin platnosci; termin zaplaty; data platnosci; termin; due date"),
          ("category", "ні", "kategoria")],
 "VISITS": [("date", "так", "data wizyty; data; dzien; termin wizyty; data zabiegu"),
          ("patient", "так", "pacjent; imie i nazwisko; nazwisko i imie; nr karty; numer karty; id pacjenta"),
          ("doctor", "так", "lekarz; lekarz prowadzacy; wykonujacy; wykonal; personel; higienistka"),
          ("service", "ні", "usluga; procedura; zabieg; nazwa uslugi; czynnosc; swiadczenie"),
          ("amount", "так", "kwota; wartosc; cena; naleznosc; do zaplaty; oplata; suma; wartosc brutto"),
          ("status", "ні", "status; status wizyty; stan"),
          ("duration", "ні", "czas trwania; czas; dlugosc wizyty; minuty; czas wizyty"),
          ("start", "ні", "godzina od; od; poczatek; godzina rozpoczecia; godzina"),
          ("end", "ні", "godzina do; do; koniec; godzina zakonczenia"),
          ("lab", "ні", "koszt laboratorium; laboratorium; technik; protetyk"),
          ("prep", "ні", "koszt materialow; materialy; implant; preparaty"),
          ("payment", "ні", "forma platnosci; sposob platnosci; metoda platnosci; platnosc"),
          ("isNew", "ні", "nowy pacjent; pierwsza wizyta; nowy"),
          ("source", "ні", "zrodlo pacjenta; zrodlo; skad pacjent; skad"),
          ("direction", "ні", "specjalizacja; dziedzina; kategoria uslugi")],
 "PATIENTS": [("patient", "так", "pacjent; imie i nazwisko; nazwisko i imie; nr karty; numer karty; id pacjenta"),
          ("regDate", "так", "data rejestracji; data dodania; data utworzenia; data pierwszej wizyty; pierwsza wizyta"),
          ("source", "ні", "zrodlo pacjenta; zrodlo; skad pacjent; skad")],
 "ZL": [("date", "так", "data wizyty; termin wizyty; data i godzina; data; termin; appointment date; date"),
          ("patient", "так", "pacjent; imie i nazwisko; nazwisko i imie; patient"),
          ("doctor", "ні", "lekarz; specjalista; personel"),
          ("service", "ні", "usluga; nazwa uslugi; service"),
          ("status", "ні", "status wizyty; status; stan")],
 "BOOKSY": [("date", "так", "data wizyty; termin wizyty; data i godzina; data; termin; appointment date; date"),
          ("patient", "так", "klient; nazwa klienta; imie i nazwisko; pacjent; customer; client"),
          ("doctor", "ні", "pracownik; specjalista; personel; staff; lekarz"),
          ("service", "ні", "usluga; nazwa uslugi; service"),
          ("status", "ні", "status wizyty; status; stan")],
 "ADS": [("date", "ні", "dzien; data; day; date; miesiac; month"),
          ("campaign", "так", "kampania; campaign"),
          ("impressions", "ні", "wyswietlenia; impr.; impressions"),
          ("clicks", "ні", "klikniecia; clicks"),
          ("cost", "так", "koszt; cost"),
          ("conversions", "ні", "konwersje; conversions")],
 "WPLATY": [("date", "так", "data; data wplaty"), ("patient", "так", "pacjent"), ("description", "так", "opis"),
          ("payment", "так", "metoda platnosci"), ("amount", "так", "kwota"), ("addedBy", "ні", "dodane przez")],
 "FKOSZTY": [("date", "так", "data"), ("addedBy", "так", "dodane przez"), ("description", "так", "opis"),
          ("payment", "ні", "metoda platnosci"), ("amount", "так", "kwota")],
 "FSOURCES": [("firstName", "так", "imie"), ("lastName", "так", "nazwisko"), ("source", "так", "pozyskanie pacjenta; skad pacjent; zrodlo")],
 "FSTATUS": [("status", "так", "status wizyty"), ("count", "так", "ilosc wizyt; liczba wizyt")],
 "FDUR": [("doctor", "так", "lekarz"), ("avg", "так", "sredni czas wizyty; srednia dlugosc wizyty")],
 "FZAPISY": [("patient", "так", "pacjent"), ("added", "ні", "data dodania wizyty"), ("orig", "так", "termin wizyty (pierwotny); termin wizyty"),
          ("newTerm", "ні", "nowy termin wizyty (przy przesunieciu); nowy termin wizyty"), ("status", "так", "status wizyt; status"),
          ("changedBy", "ні", "dokonal zmiany")],
}
TYPE_LBL = {"BANK": "Виписка з банку", "INVOICES": "Фактури (реєстр)", "VISITS": "Felgdent – візити", "PATIENTS": "Felgdent – пацієнти",
            "ZL": "ZnanyLekarz – записи", "BOOKSY": "Booksy – записи", "ADS": "Google Ads – звіт",
            "WPLATY": "Felgdent – wpłaty", "FKOSZTY": "Felgdent – koszty", "FSOURCES": "Felgdent – pozyskanie pacjenta",
            "FSTATUS": "Felgdent – status wizyt", "FDUR": "Felgdent – średnia długość wizyty", "FZAPISY": "Felgdent – przesunięcia wizyt"}
STATUSES = [("nieobecn; nie przysz; nie stawil; niestawien; no-show; no show; noshow", "Неявка"),
            ("odwolan; anulowan; rezygnac; cancel", "Скасовано"),
            ("przesuni; zaplanowan; potwierdzon; umowion; oczekuj; planned; confirmed; booked; nowa", "Пропустити"),
            ("zrealizowan; odbyt; wykonan; zakonczon; obecny; obecna; przyszedl; completed; done; finished", "Виконано")]
HINTS = [("WPLATY", "wplaty"), ("FKOSZTY", "koszty"), ("FSOURCES", "pozyskani"), ("FSTATUS", "status wizyt"),
         ("FDUR", "dlugosc wizyty; czas wizyty"), ("FZAPISY", "przesuni"), ("BOOKSY", "booksy"), ("ZL", "znany; docplanner"), ("ADS", "google ads; googleads; adwords; kampani; campaign"),
         ("INVOICES", "faktur; ksef; invoice"), ("PATIENTS", "pacjenci; pacjentow; patients"),
         ("VISITS", "felg; wizyty; wizyt; zabiegi"), ("BANK", "wyciag; historia; pko; mbank; santander; ing; pekao; millennium; alior; bnp; credit agricole; bank; operacj")]
ic = wb.create_sheet("Імпорт-налаштування")
ic["A1"] = "НАЛАШТУВАННЯ ІМПОРТУ: як скрипт впізнає колонки у файлах"; ic["A1"].font = font(bold=True, size=12)
ic["A2"] = ("Якщо файл не розпізнається, додайте точну назву колонки з вашого файлу в колонку D (через крапку з комою). Регістр і польські літери не важливі. "
            "«так» — колонка обов'язкова; «один із» — потрібна хоча б одна з таких.")
ic["A2"].font = font(italic=True, color="444444"); ic["A2"].alignment = Alignment(wrap_text=True); ic.merge_cells("A2:E2"); ic.row_dimensions[2].height = 32
header(ic, 3, ["Тип файлу", "Поле", "Обов'язкове", "Назви колонок у файлі (через ;)", "Тип (опис)"])
r = 4
for t, fl in FIELDS.items():
    for name, req, al in fl:
        for j, v in enumerate([t, name, req, al, TYPE_LBL[t]]):
            c = ic.cell(row=r, column=1 + j, value=v); c.border = BOX; c.font = font(color="0000FF" if j == 3 else "000000")
            if j == 3: c.fill = INP
        r += 1
header(ic, 3, ["Фрагменти статусу у файлі (через ;) — перевіряються зверху вниз", "Статус у звіті"], 7)
for i, (fr, s) in enumerate(STATUSES):
    ic.cell(row=4 + i, column=7, value=fr).fill = INP; ic.cell(row=4 + i, column=8, value=s)
    for col in (7, 8): ic.cell(row=4 + i, column=col).border = BOX
header(ic, 3, ["Тип файлу", "Фрагменти назви файлу (через ;)"], 10)
for i, (t, fr) in enumerate(HINTS):
    ic.cell(row=4 + i, column=10, value=t); ic.cell(row=4 + i, column=11, value=fr).fill = INP
    for col in (10, 11): ic.cell(row=4 + i, column=col).border = BOX
widths(ic, {"A": 11, "B": 13, "C": 11, "D": 90, "E": 22, "F": 3, "G": 60, "H": 14, "I": 3, "J": 11, "K": 60})

if DEFAULTS_JSON:
    json.dump({"fields": {t: [{"name": n, "required": rq, "aliases": [nrm(a) for a in al.split(";") if a.strip()]} for n, rq, al in fl] for t, fl in FIELDS.items()},
               "statuses": [{"fragments": [nrm(a) for a in fr.split(";") if a.strip()], "status": s} for fr, s in STATUSES],
               "hints": [{"type": t, "fragments": [nrm(a) for a in fr.split(";") if a.strip()]} for t, fr in HINTS],
               "rules": [{"kw": nrm(k), "dir": {"витрата": "out", "надходження": "in"}.get(d, "any"), "category": c} for k, d, c in RULES]},
              open(DEFAULTS_JSON, "w"), ensure_ascii=False, indent=1)

def categorize(cp, desc, amt):
    text = " " + nrm(cp + " " + desc) + " "
    for k, d, c in RULES:
        kw = nrm(k)
        if kw in text and (d == "будь-яка" or (d == "витрата" and amt < 0) or (d == "надходження" and amt > 0)):
            return c
    return "Надходження: інше" if amt > 0 else ""

# =============== Послуги ===============
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

# =============== Візити ===============
vs = wb.create_sheet("Візити")
V = "'Візити'!"
header(vs, 1, ["Дата візиту", "Місяць (авто)", "Пацієнт (ID / прізвище)", "Новий пацієнт? (Так/Ні)", "Джерело пацієнта", "Лікар", "Напрям",
               "Послуга", "Вартість, zł", "Технік (лабораторія), zł", "Спецпрепарати / імплант, zł", "Спосіб оплати", "Статус", "Тривалість, год",
               "% лікаря (авто)", "Виплата лікарю (авто), zł", "Маржа клініки (авто), zł", "Ключ імпорту"])
vs.freeze_panes = "C2"; vs.row_dimensions[1].height = 45
widths(vs, {"A": 12, "B": 10, "C": 18, "D": 10, "E": 20, "F": 15, "G": 15, "H": 34, "I": 11, "J": 11, "K": 12, "L": 18, "M": 11, "N": 9, "O": 9, "P": 13, "Q": 13, "R": 30})
vs.column_dimensions["R"].hidden = True
doc_dirs = {"Lekarz 1": ["Zachowawcza", "Endodoncja"], "Lekarz 2": ["Zachowawcza", "Dziecięca"], "Lekarz 3": ["Endodoncja"],
            "Lekarz 4": ["Chirurgia", "Implantologia"], "Lekarz 5": ["Protetyka"], "Lekarz 6": ["Ortodoncja"],
            "Lekarz 7": ["Zachowawcza", "Protetyka"], "Higienistka 1": ["Profilaktyka"], "Higienistka 2": ["Profilaktyka"]}
doc_weight = {"Lekarz 1": 16, "Lekarz 2": 14, "Lekarz 3": 7, "Lekarz 4": 9, "Lekarz 5": 8, "Lekarz 6": 12, "Lekarz 7": 9, "Higienistka 1": 13, "Higienistka 2": 12}
src_w = [("Google", 26), ("Instagram/Facebook", 10), ("Z sieci", 6), ("ZnanyLekarz", 12), ("Booksy", 4), ("Ulotka (promocja)", 4),
         ("Znajomy", 14), ("Inny pacjent", 10), ("Mailing", 2), ("Z ulicy", 5), ("Inne", 3)]
pay_w = [("Karta", 45), ("BLIK", 15), ("Gotówka", 15), ("Przelew", 10), ("Raty PayU", 5), ("Mediraty", 4), ("Bon podarunkowy", 2), ("Przedpłata (wcześniej)", 4)]
def wchoice(pairs): return random.choices([p[0] for p in pairs], weights=[p[1] for p in pairs])[0]
def workdays(y, m):
    d = dt.date(y, m, 1); out = []
    while d.month == m:
        if d.weekday() < 5: out.append(d)
        d += dt.timedelta(days=1)
    return out
visits, patients = [], []
pid = 0
for (y, m, nvis) in [(2026, 9, 215), (2026, 10, 225)]:
    days = workdays(y, m)
    for _ in range(nvis):
        doc = wchoice(list(doc_weight.items())); direc = random.choice(doc_dirs[doc])
        svc = random.choices(services[direc], weights=[s[7] for s in services[direc]])[0]
        name, pmin, pmax, dmin, dmax, lab, prep, _w = svc
        new = random.random() < 0.24 or not patients
        if name.startswith("Przegląd"): new = True
        if new:
            pid += 1
            src = "Ulotka (promocja)" if name.startswith("Przegląd") else wchoice(src_w)
            p = (f"P{pid:04d}", src); patients.append(p)
        else:
            p = random.choice(patients)
        price = round(random.uniform(pmin, pmax) / 10) * 10
        r_ = random.random()
        status = "Виконано" if r_ < 0.92 else ("Неявка" if r_ < 0.97 else "Скасовано")
        visits.append([random.choice(days), p[0], "Так" if new else "Ні", p[1], doc, direc, name, price, lab, prep,
                       wchoice(pay_w) if status == "Виконано" and price > 0 else None, status, random.choice([dmin, dmax])])
visits.sort(key=lambda x: x[0])
for i in range(2, N + 1):
    if i - 2 < len(visits):
        v = visits[i - 2]
        for j, val in enumerate([v[0], None, v[1], v[2], v[3], v[4], v[5], v[6], v[7], v[8] or None, v[9] or None, v[10], v[11], v[12]]):
            if j != 1: vs.cell(row=i, column=1 + j, value=val)
    vs.cell(row=i, column=1).number_format = DATE
    vs.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    vs.cell(row=i, column=15, value=f'=IF(F{i}="","",IFERROR(INDEX({DOC_P},MATCH(F{i},{DOC_N},0))*1,0))').number_format = "0%"
    vs.cell(row=i, column=16, value=f'=IF(M{i}="Виконано",MAX(0,I{i}-J{i}-K{i})*N(O{i}),0)').number_format = ZL
    vs.cell(row=i, column=17, value=f'=IF(M{i}="Виконано",I{i}-J{i}-K{i}-P{i},0)').number_format = ZL
    for col in (9, 10, 11): vs.cell(row=i, column=col).number_format = ZL
    for col in (2, 15, 16, 17): vs.cell(row=i, column=col).fill = CALC
add_list(vs, f"D2:D{N}", f"{D}$V$4:$V$5"); add_list(vs, f"E2:E{N}", f"{D}$J$4:$J$19"); add_list(vs, f"F2:F{N}", DOC_N)
add_list(vs, f"G2:G{N}", f"{D}$F$4:$F$15"); add_list(vs, f"L2:L{N}", f"{D}$H$4:$H$13"); add_list(vs, f"M2:M{N}", f"{D}$T$4:$T$6")
cf_text(vs, f"M2:M{N}", "$M2".replace("$M2", 'IF($M2="Виконано","добре",IF($M2="Скасовано","середньо",IF($M2="Неявка","погано","")))'))
cf_rules(vs, f"G2:G{N}", None, None, 'AND($A2<>"",$G2="")')
cf_rules(vs, f"F2:F{N}", None, None, 'AND($A2<>"",OR($F2="",N($O2)=0))')
cf_rules(vs, f"E2:E{N}", None, None, 'AND($A2<>"",$D2="Так",$E2="")')

# =============== Банк ===============
bk = wb.create_sheet("Банк")
B = "'Банк'!"
header(bk, 1, ["Дата", "Місяць (авто)", "Сума, zł (− списання, + надходження)", "Контрагент", "Опис / tytuł", "Авто-категорія (скрипт)",
               "Ручна категорія (якщо треба виправити)", "Категорія (авто)", "Група (авто)", "Рахунок / файл", "Ключ імпорту"])
bk.freeze_panes = "A2"; bk.row_dimensions[1].height = 45
widths(bk, {"A": 12, "B": 10, "C": 14, "D": 30, "E": 40, "F": 30, "G": 30, "H": 30, "I": 24, "J": 16, "K": 30})
bk.column_dimensions["K"].hidden = True
bank_rows = []
def bank(d, amt, cp, desc, manual=None): bank_rows.append((d, round(amt, 2), cp, desc, manual))
def month_sum(m, key, val):
    return sum(v[7] for v in visits if v[0].month == m and v[11] == "Виконано" and v[key] == val)
for m in (9, 10):
    nxt = lambda d: dt.date(2026, m, d)
    mv = [v for v in visits if v[0].month == m and v[11] == "Виконано"]
    lab_sum = sum(v[8] for v in mv); prep_sum = sum(v[9] for v in mv)
    bank(nxt(2), -9000, "WYNAJMUJĄCY SP. Z O.O.", f"Czynsz {m:02d}/2026")
    bank(nxt(5), -random.choice([1080, 1150, 1220]), "PGE OBRÓT S.A.", "Energia elektryczna")
    bank(nxt(6), -260, "MPWiK", "Woda i ścieki"); bank(nxt(6), -430, "PGNiG OD", "Gaz")
    bank(nxt(8), -1200, "CZYSTY DOM S.C.", "Usługi porządkowe", "Sprzątanie / pranie" if m == 9 else None)
    bank(nxt(10), -4200, "mLeasing Sp. z o.o.", "Rata leasingu unit"); bank(nxt(10), -2300, "PKO Leasing SA", "Rata leasingu RTG")
    bank(nxt(12), -800, "SERWIS DENTAL", "Przegląd autoklawu"); bank(nxt(12), -600, "REMONDIS MEDISON", "Odbiór odpadów medycznych")
    bank(nxt(3), -(6200 if m == 9 else 6900), "GOOGLE IRELAND LTD", "Google Ads")
    bank(nxt(3), -1500, "FACEBK *ADS", "Meta ads"); bank(nxt(4), -940, "DOCPLANNER SP. Z O.O.", "Abonament ZnanyLekarz")
    bank(nxt(4), -270, "BOOKSY INTERNATIONAL", "Subskrypcja"); bank(nxt(15), -380, "DRUKARNIA FOTO-PRINT", "Druk voucherów")
    bank(nxt(14), -1500, "BIURO RACHUNKOWE ABC", f"Księgowość {m:02d}/2026"); bank(nxt(1), -450, "FELGDENT", "Abonament")
    bank(nxt(9), -300, "ORANGE POLSKA", "Telefon i internet"); bank(nxt(9), -350, "PZU SA", "Polisa OC")
    bank(nxt(16), -5200, "HENRY SCHEIN POLSKA", f"FV 2026/{m:02d}/118"); bank(nxt(22), -4400, "SIGMA DENTAL", f"FV {400+m}/2026")
    bank(nxt(15), -1800, "ZUS", "Składki"); bank(nxt(20), -4590 + (m - 9) * 200, "URZĄD SKARBOWY", f"PIT-5L {m:02d}/2026")
    bank(nxt(28), -780, "ELAVON", "Prowizja terminal"); bank(nxt(28), -40, "BANK", "Opłata za prowadzenie rachunku")
    bank(nxt(25), -10000, "WŁAŚCICIEL", "Wypłata właściciela")
    bank(nxt(18), -round(lab_sum * 0.9, -1), "LAB PROTETYK SP. Z O.O.", f"FV LAB/{m:02d}/2026")
    bank(nxt(19), -round(prep_sum * 0.8, -1), "STRAUMANN SP. Z O.O.", f"FV ST/{m:02d}/77")
    for doc in docs:
        pay = sum(max(0, v[7] - v[8] - v[9]) * doc[2] for v in mv if v[4] == doc[0])
        if pay: bank(nxt(27), -round(pay, -1), doc[0].upper(), f"Wynagrodzenie {m:02d}/2026")
    staff_total = 160 * 35 + 140 * 35 + 40 * 32 + 160 * 32 + 60 * 32
    bank(nxt(27), -staff_total, "ASYSTENTKI I RECEPCJA", f"Umowa zlecenie {m:02d}/2026")
    card = month_sum(m, 10, "Karta"); blik = month_sum(m, 10, "BLIK") + month_sum(m, 10, "Przelew")
    for k in range(4): bank(nxt(7 + k * 7), round(card / 4 * 0.99, 2), "ELAVON FINANCIAL SERVICES", "Rozliczenie transakcji kartowych")
    for k in range(5): bank(nxt(5 + k * 5), round(blik / 5, 2), f"PACJENT {k+1}", "Zapłata za leczenie")
    bank(nxt(21), month_sum(m, 10, "Raty PayU"), "PAYU S.A.", "Wypłata raty")
    bank(nxt(23), month_sum(m, 10, "Mediraty"), "MEDIRATY SP. Z O.O.", "Finansowanie leczenia")
    bank(nxt(26), round(month_sum(m, 10, "Gotówka") * 0.6, -2), "WPŁATA WŁASNA", "Wpłatomat")
bank_rows.sort(key=lambda x: x[0])
for i in range(2, N + 1):
    if i - 2 < len(bank_rows):
        d, amt, cp, desc, manual = bank_rows[i - 2]
        bk.cell(row=i, column=1, value=d); bk.cell(row=i, column=3, value=amt); bk.cell(row=i, column=4, value=cp)
        bk.cell(row=i, column=5, value=desc); bk.cell(row=i, column=6, value=categorize(cp, desc, amt) or None)
        bk.cell(row=i, column=7, value=manual); bk.cell(row=i, column=10, value="Demo")
        bk.cell(row=i, column=11, value=f"B|demo|{i}")
    bk.cell(row=i, column=1).number_format = DATE
    bk.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    bk.cell(row=i, column=3).number_format = ZL2
    bk.cell(row=i, column=8, value=f'=IF(A{i}="","",IF(G{i}<>"",G{i},F{i}))')
    bk.cell(row=i, column=9, value=f'=IF(A{i}="","",IF(H{i}="","⚠ без категорії",IFERROR(INDEX({CAT_G},MATCH(H{i},{CAT_N},0)),"⚠ без категорії")))')
    for col in (2, 6, 8, 9): bk.cell(row=i, column=col).fill = CALC
add_list(bk, f"G2:G{N}", CAT_N)
cf_rules(bk, f"H2:I{N}", None, None, '$I2="⚠ без категорії"')
cf_rules(bk, f"C2:C{N}", 'AND(ISNUMBER($C2),$C2>0)', None, None)

# =============== Фактури ===============
fv = wb.create_sheet("Фактури")
FV = "'Фактури'!"
header(fv, 1, ["Номер фактури", "Дата виставлення", "Контрагент (продавець)", "NIP", "Сума брутто, zł", "Термін оплати", "Категорія",
               "Дата оплати (скрипт / вручну)", "Статус (авто)", "Днів до терміну (авто)", "Ключ імпорту"])
fv.freeze_panes = "A2"; fv.row_dimensions[1].height = 45
widths(fv, {"A": 20, "B": 12, "C": 30, "D": 13, "E": 13, "F": 12, "G": 32, "H": 14, "I": 16, "J": 12, "K": 30})
fv.column_dimensions["K"].hidden = True
inv = []
for m in (9, 10):
    inv += [(f"2026/{m:02d}/118", dt.date(2026, m, 3), "HENRY SCHEIN POLSKA", "5260001111", 5200, dt.date(2026, m, 17), "Materiały i narzędzia (zakupy)", dt.date(2026, m, 16)),
            (f"{400+m}/2026", dt.date(2026, m, 8), "SIGMA DENTAL", "5270002222", 4400, dt.date(2026, m, 22), "Materiały i narzędzia (zakupy)", dt.date(2026, m, 22)),
            (f"LAB/{m:02d}/2026", dt.date(2026, m, 10), "LAB PROTETYK SP. Z O.O.", "5280003333", None, dt.date(2026, m, 24), "Laboratorium – płatność (звірка)", dt.date(2026, m, 18)),
            (f"FV/{m:02d}/2026/15", dt.date(2026, m, 1), "BIURO RACHUNKOWE ABC", "5290004444", 1500, dt.date(2026, m, 14), "Księgowość", dt.date(2026, m, 14))]
inv[2] = inv[2][:4] + (next(b for b in bank_rows if b[2].startswith("LAB") and b[0].month == 9)[1] * -1,) + inv[2][5:]
inv[6] = inv[6][:4] + (next(b for b in bank_rows if b[2].startswith("LAB") and b[0].month == 10)[1] * -1,) + inv[6][5:]
inv += [("FV 2026/11/031", dt.date(2026, 11, 2), "SERWIS DENTAL", "5300005555", 1850, dt.date(2026, 11, 30), "Serwis i naprawy sprzętu", None),
        ("FV/10/889", dt.date(2026, 10, 20), "PRO-DENT", "5310006666", 2340, dt.date(2026, 11, 3), "Materiały i narzędzia (zakupy)", None)]
for i in range(2, N + 1):
    if i - 2 < len(inv):
        for j, v in enumerate(inv[i - 2]): fv.cell(row=i, column=1 + j, value=v)
        fv.cell(row=i, column=11, value=f"F|demo|{i}")
    for col in (2, 6, 8): fv.cell(row=i, column=col).number_format = DATE
    fv.cell(row=i, column=5).number_format = ZL2
    fv.cell(row=i, column=9, value=f'=IF(A{i}="","",IF(H{i}<>"","Оплачено",IF(F{i}="","Не оплачено",IF(F{i}<TODAY(),"Прострочено","Очікує оплати"))))')
    fv.cell(row=i, column=10, value=f'=IF(OR(A{i}="",H{i}<>"",F{i}=""),"",F{i}-TODAY())').number_format = '0;-0'
    for col in (9, 10): fv.cell(row=i, column=col).fill = CALC
add_list(fv, f"G2:G{N}", CAT_N)
cf_text(fv, f"I2:J{N}", 'IF($I2="Оплачено","добре",IF($I2="Прострочено","погано",IF($I2="","","середньо")))')

# =============== Витрати (вручну) ===============
ex = wb.create_sheet("Витрати")
E = "'Витрати'!"
header(ex, 1, ["Дата", "Місяць (авто)", "Категорія витрат", "Група (авто)", "Опис / постачальник", "Сума, zł", "Спосіб оплати",
               "Лікар (для техніка)", "Ключ імпорту"])
ex.freeze_panes = "A2"; ex.row_dimensions[1].height = 32
widths(ex, {"A": 12, "B": 10, "C": 40, "D": 24, "E": 34, "F": 13, "G": 16, "H": 22, "I": 30, "K": 62}); ex.column_dimensions["I"].hidden = True
ex["K1"] = ("ТІЛЬКИ витрати, яких НЕМАЄ у виписці з банку (готівка, інша картка). Банківські витрати приходять автоматично з аркуша «Банк». "
            "Виплати лікарям рахуються з «Візитів», погодинний персонал — з «Персоналу». Витрати на техніка з Felgdent (Koszty) імпортуються сюди з лікарем.")
ex["K1"].font = font(bold=True, color="9C0006"); ex["K1"].alignment = Alignment(wrap_text=True)
for i, (d, c_, desc, a) in enumerate([(dt.date(2026, 9, 12), "Biuro i drobne", "Kawa, woda, drobne (gotówka)", 230),
                                       (dt.date(2026, 10, 9), "Szkolenia", "Kurs endodontyczny (gotówka)", 500),
                                       (dt.date(2026, 10, 14), "Biuro i drobne", "Drobne zakupy (gotówka)", 180)]):
    ex.cell(row=2 + i, column=1, value=d); ex.cell(row=2 + i, column=3, value=c_); ex.cell(row=2 + i, column=5, value=desc)
    ex.cell(row=2 + i, column=6, value=a); ex.cell(row=2 + i, column=7, value="Gotówka")
for i in range(2, N + 1):
    ex.cell(row=i, column=1).number_format = DATE
    ex.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    ex.cell(row=i, column=4, value=f'=IF(C{i}="","",IFERROR(INDEX({CAT_G},MATCH(C{i},{CAT_N},0)),"⚠ немає в довіднику"))')
    ex.cell(row=i, column=6).number_format = ZL
    for col in (2, 4): ex.cell(row=i, column=col).fill = CALC
add_list(ex, f"C2:C{N}", CAT_N); add_list(ex, f"G2:G{N}", f"{D}$H$4:$H$13"); add_list(ex, f"H2:H{N}", DOC_N)
cf_rules(ex, f"D2:D{N}", None, None, '$D2="⚠ немає в довіднику"')

# =============== Персонал ===============
pe = wb.create_sheet("Персонал")
P = "'Персонал'!"
header(pe, 1, ["Дата (будь-який день місяця)", "Місяць (авто)", "Працівник", "Роль", "Годин за місяць", "Ставка, zł/год", "Сума, zł (авто)"])
pe.freeze_panes = "A2"; pe.row_dimensions[1].height = 32
widths(pe, {"A": 14, "B": 10, "C": 26, "D": 16, "E": 12, "F": 12, "G": 14, "I": 60})
pe["I1"] = "Погодинний персонал (umowa zlecenie). Один рядок на людину на місяць. % гігієністок рахується з «Візитів»."
pe["I1"].font = font(italic=True, color="444444"); pe["I1"].alignment = Alignment(wrap_text=True)
staff = [("Asystentka 1", "Asystentka", 160, 35), ("Asystentka 2", "Asystentka", 140, 35), ("Asystentka 3 (doraźnie)", "Asystentka", 40, 32),
         ("Recepcja 1", "Recepcja", 160, 32), ("Recepcja 2 / Higienistka 1", "Recepcja", 60, 32)]
r = 2
for m in (9, 10):
    for s in staff:
        pe.cell(row=r, column=1, value=dt.date(2026, m, 1)); pe.cell(row=r, column=3, value=s[0]); pe.cell(row=r, column=4, value=s[1])
        pe.cell(row=r, column=5, value=s[2]); pe.cell(row=r, column=6, value=s[3]); r += 1
for i in range(2, 501):
    pe.cell(row=i, column=1).number_format = DATE
    pe.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    pe.cell(row=i, column=7, value=f'=IF(A{i}="","",E{i}*F{i})').number_format = ZL
    pe.cell(row=i, column=6).number_format = ZL
    for col in (2, 7): pe.cell(row=i, column=col).fill = CALC

# =============== Пацієнти, Онлайн-записи, Google Ads, Послуги, Журнал ===============
pa = wb.create_sheet("Пацієнти")
header(pa, 1, ["Пацієнт (як у Felgdent)", "Дата реєстрації / першого візиту", "Джерело", "Ключ імпорту"])
widths(pa, {"A": 26, "B": 16, "C": 22, "D": 30, "F": 70}); pa.column_dimensions["D"].hidden = True; pa.freeze_panes = "A2"
pa["F1"] = ("Необов'язково, але дуже корисно: список пацієнтів з Felgdent з датою реєстрації. Тоді «новий пацієнт» визначається точно, "
            "а не за першим візитом у цій таблиці.")
pa["F1"].font = font(italic=True, color="444444"); pa["F1"].alignment = Alignment(wrap_text=True)
for i in range(2, N + 1): pa.cell(row=i, column=2).number_format = DATE

fsw = wb.create_sheet("Felgdent статуси")
header(fsw, 1, ["Місяць (1-ше число)", "Статус у Felgdent", "Статус у звіті", "Кількість візитів", "Ключ імпорту"])
widths(fsw, {"A": 14, "B": 36, "C": 16, "D": 12, "E": 30, "G": 70}); fsw.column_dimensions["E"].hidden = True; fsw.freeze_panes = "A2"
fsw.row_dimensions[1].height = 30
fsw["G1"] = ("Звіт Felgdent «Status wizyt» за місяць (при завантаженні вкажіть місяць). Звідси звіт «Raport» бере неявки та скасування, "
             "якщо візити імпортовано з оплат (у них неявок немає).")
fsw["G1"].font = font(italic=True, color="444444"); fsw["G1"].alignment = Alignment(wrap_text=True)
for i in range(2, 501): fsw.cell(row=i, column=1).number_format = MON
zp = wb.create_sheet("Записи Felgdent")
header(zp, 1, ["Термін візиту", "Місяць (авто)", "Пацієнт", "Статус", "Статус у Felgdent", "Перенесено?", "Хто змінив", "Ключ імпорту"])
widths(zp, {"A": 12, "B": 10, "C": 26, "D": 14, "E": 28, "F": 11, "G": 22, "H": 30, "J": 70}); zp.column_dimensions["H"].hidden = True; zp.freeze_panes = "A2"
zp["J1"] = "Звіт Felgdent «Raport przesunięć wizyt»: перенесення, скасування й неявки за датою візиту."
zp["J1"].font = font(italic=True, color="444444"); zp["J1"].alignment = Alignment(wrap_text=True)
NZ = 5001
for i in range(2, NZ + 1):
    zp.cell(row=i, column=1).number_format = DATE
    zp.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
cf_text(zp, f"D2:D{NZ}", 'IF($D2="Виконано","добре",IF(OR($D2="Скасовано",$D2="Перенесено"),"середньо",IF($D2="Неявка","погано","")))')

on = wb.create_sheet("Онлайн-записи")
O_ = "'Онлайн-записи'!"
header(on, 1, ["Дата візиту", "Місяць (авто)", "Платформа", "Пацієнт", "Лікар", "Послуга", "Статус", "Ключ імпорту"])
widths(on, {"A": 12, "B": 10, "C": 14, "D": 22, "E": 16, "F": 30, "G": 14, "H": 30}); on.column_dimensions["H"].hidden = True; on.freeze_panes = "A2"
onl = []
for v in visits:
    if v[3] in ("ZnanyLekarz", "Booksy") and v[2] == "Так":
        onl.append((v[0], v[3], v[1], v[4], v[6], "Виконано" if v[11] == "Виконано" else v[11]))
for m in (9, 10):
    for k in range(random.randint(3, 6)):
        onl.append((random.choice(workdays(2026, m)), random.choice(["ZnanyLekarz", "ZnanyLekarz", "Booksy"]), f"Online {m}-{k}", None, "Konsultacja", "Скасовано"))
onl.sort(key=lambda x: x[0])
for i in range(2, N + 1):
    if i - 2 < len(onl):
        d, pl, pt, dc, sv, s_ = onl[i - 2]
        for col, val in ((1, d), (3, pl), (4, pt), (5, dc), (6, sv), (7, s_)): on.cell(row=i, column=col, value=val)
        on.cell(row=i, column=8, value=f"O|demo|{i}")
    on.cell(row=i, column=1).number_format = DATE
    on.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    on.cell(row=i, column=2).fill = CALC
cf_text(on, f"G2:G{N}", 'IF($G2="Виконано","добре",IF($G2="Скасовано","середньо",IF($G2="Неявка","погано","")))')

ga = wb.create_sheet("Google Ads")
GA = "'Google Ads'!"
header(ga, 1, ["Дата (або 1-ше число місяця)", "Місяць (авто)", "Кампанія", "Покази", "Кліки", "Вартість, zł", "Конверсії (дзвінки / заявки)", "Ключ імпорту"])
widths(ga, {"A": 14, "B": 10, "C": 30, "D": 11, "E": 10, "F": 13, "G": 14, "H": 30}); ga.column_dimensions["H"].hidden = True
ga.freeze_panes = "A2"; ga.row_dimensions[1].height = 32
ads = []
for m, tot in ((9, 6200), (10, 6900)):
    for camp, share, cpc, cr in (("Implanty", 0.45, 6.5, 0.05), ("Leczenie kanałowe", 0.30, 4.8, 0.06), ("Marka – nazwa kliniki", 0.25, 1.9, 0.12)):
        cost = round(tot * share, 2); clicks = int(cost / cpc)
        ads.append((dt.date(2026, m, 1), camp, clicks * 18, clicks, cost, round(clicks * cr)))
for i in range(2, N + 1):
    if i - 2 < len(ads):
        for col, val in zip((1, 3, 4, 5, 6, 7), ads[i - 2]): ga.cell(row=i, column=col, value=val)
        ga.cell(row=i, column=8, value=f"A|demo|{i}")
    ga.cell(row=i, column=1).number_format = DATE
    ga.cell(row=i, column=2, value=f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))').number_format = MON
    ga.cell(row=i, column=2).fill = CALC; ga.cell(row=i, column=6).number_format = ZL2

sv = wb.create_sheet("Послуги")
header(sv, 1, ["Послуга (точно як у Felgdent)", "Напрям", "Тривалість за замовчуванням, год"])
widths(sv, {"A": 44, "B": 18, "C": 16, "E": 70}); sv.freeze_panes = "A2"; sv.row_dimensions[1].height = 32
sv["E1"] = "Скрипт додає сюди нові послуги з Felgdent. Червоне = вкажіть напрям (один раз), далі він підставлятиметься автоматично."
sv["E1"].font = font(italic=True, color="444444"); sv["E1"].alignment = Alignment(wrap_text=True)
r = 2
for d_, lst in services.items():
    for s in lst:
        sv.cell(row=r, column=1, value=s[0]); sv.cell(row=r, column=2, value=d_); sv.cell(row=r, column=3, value=s[4]); r += 1
add_list(sv, "B2:B500", f"{D}$F$4:$F$15")
cf_rules(sv, "B2:B500", None, None, 'AND($A2<>"",$B2="")')

jl = wb.create_sheet("Журнал імпорту")
header(jl, 1, ["Дата і час", "Файл", "Тип", "Додано", "Оновлено", "Пропущено", "Попередження / помилки"])
widths(jl, {"A": 18, "B": 36, "C": 22, "D": 9, "E": 9, "F": 10, "G": 90})

# ========================= формули для звітів =========================
def vr(col): return f"{V}${col}$2:${col}${N}"
VB, VD, VE, VF, VG, VI, VJ, VK, VL, VM, VN, VO, VP, VQ = [vr(c) for c in "BDEFGIJKLMNOPQ"]
EB, ED, EC, EF = f"{E}$B$2:$B${N}", f"{E}$D$2:$D${N}", f"{E}$C$2:$C${N}", f"{E}$F$2:$F${N}"
PB, PG = f"{P}$B$2:$B$500", f"{P}$G$2:$G$500"
BB, BC, BH, BI = f"{B}$B$2:$B${N}", f"{B}$C$2:$C${N}", f"{B}$H$2:$H${N}", f"{B}$I$2:$I${N}"
OB, OC, OG = f"{O_}$B$2:$B${N}", f"{O_}$C$2:$C${N}", f"{O_}$G$2:$G${N}"
GB, GD, GE, GF, GG = [f"{GA}${c}$2:${c}${N}" for c in "BDEFG"]
FI, FE, FH = f"{FV}$I$2:$I${N}", f"{FV}$E$2:$E${N}", f"{FV}$H$2:$H${N}"
OK = '"Виконано"'

def rev(m, x=""): return f'SUMIFS({VI},{VB},{m},{VM},{OK}{x})'
def lab(m, x=""): return f'SUMIFS({VJ},{VB},{m},{VM},{OK}{x})'
def prep(m, x=""): return f'SUMIFS({VK},{VB},{m},{VM},{OK}{x})'
def payout(m, x=""): return f'SUMIFS({VP},{VB},{m},{VM},{OK}{x})'
def marg(m, x=""): return f'SUMIFS({VQ},{VB},{m},{VM},{OK}{x})'
def hrs(m, x=""): return f'SUMIFS({VN},{VB},{m},{VM},{OK}{x})'
def cnt(m, x=""): return f'COUNTIFS({VB},{m},{VM},{OK}{x})'
def newp(m, x=""): return f'COUNTIFS({VB},{m},{VM},{OK},{VD},"Так"{x})'
def staff_cost(m): return f'SUMIFS({PG},{PB},{m})'
FSA, FSC, FSD = "'Felgdent статуси'!$A$2:$A$500", "'Felgdent статуси'!$C$2:$C$500", "'Felgdent статуси'!$D$2:$D$500"
ZB, ZD = "'Записи Felgdent'!$B$2:$B$5001", "'Записи Felgdent'!$D$2:$D$5001"
EH = f"{E}$H$2:$H${N}"
LABF = '"Technik / laboratorium (koszt z Felgdent)"'
def labF(m, doc=None): return f'SUMIFS({EF},{EB},{m},{EC},{LABF}' + (f',{EH},{doc})' if doc else ')')
def labF_pct(m): return f'SUMPRODUCT(SUMIFS({EF},{EB},{m},{EC},{LABF},{EH},{DOC_N})*{DOC_P})'
def noshow(m):
    fs_ = lambda st_: f'SUMIFS({FSD},{FSA},{m},{FSC},"{st_}")'
    felg = f'({fs_("Неявка")}+{fs_("Скасовано")})/({fs_("Виконано")}+{fs_("Неявка")}+{fs_("Скасовано")})'
    vis = f'(COUNTIFS({VB},{m},{VM},"Неявка")+COUNTIFS({VB},{m},{VM},"Скасовано"))/COUNTIFS({VB},{m},{VM},"<>")'
    return f'IFERROR(IF(SUMIFS({FSD},{FSA},{m})>0,{felg},{vis}),0)'
def grp(m, g): return f'(SUMIFS({EF},{EB},{m},{ED},{g})-SUMIFS({BC},{BB},{m},{BI},{g}))'
def cat_cost(m, c): return f'(SUMIFS({EF},{EB},{m},{EC},{c})-SUMIFS({BC},{BB},{m},{BH},{c}))'
def bank_cat(m, c): return f'SUMIFS({BC},{BB},{m},{BH},{c})'
def unalloc(m): return f'(SUMIFS({EF},{EB},{m},{ED},"⚠ немає в довіднику")-SUMIFS({BC},{BB},{m},{BI},"⚠ без категорії",{BC},"<0"))'
def opex_all(m): return f'(SUMPRODUCT(SUMIFS({EF},{EB},{m},{ED},{GRP_N}))-SUMPRODUCT(SUMIFS({BC},{BB},{m},{BI},{GRP_N}))+{unalloc(m)})'
def cap(m): return f'({S}$B$4*{S}$B$5*(NETWORKDAYS({m},DATE(YEAR({m}),MONTH({m})+1,0))+{S}$B$6))'
def ads_cost(m): return f'SUMIFS({GF},{GB},{m})'
def noncash(m): return "(" + "+".join(f'SUMIFS({VI},{VB},{m},{VM},{OK},{VL},"{p}")' for p in NONCASH) + ")"

# ========================= ДАШБОРД =========================
db = wb.create_sheet("Дашборд", 0)
widths(db, {"A": 36, "B": 14, "C": 14, "D": 14, "E": 24, "F": 13, "G": 14, "H": 14, "I": 14, "J": 14, "K": 12, "L": 10, "M": 12})
db.sheet_view.showGridLines = False
TH = th_helper(db, 16, 4)
def T(k): return TH[k]
db["A1"] = "МІСЯЧНИЙ ЗВІТ КЛІНІКИ"; db["A1"].font = font(bold=True, size=16, color="1F4E5A")
db["A2"] = "Звітний місяць:"; db["A2"].font = font(bold=True)
db["B2"] = f"={MCUR}"; db["B2"].number_format = MON; db["B2"].font = font(bold=True, size=12)
db["C2"] = "порівняння з:"; db["C2"].font = font(italic=True, color="666666"); db["C2"].alignment = Alignment(horizontal="right")
db["D2"] = f"={MPREV}"; db["D2"].number_format = MON; db["D2"].font = font(italic=True, color="666666")
db["F2"] = "Місяць змінюється на аркуші «Налаштування»"; db["F2"].font = font(italic=True, color="999999", size=9)
db["A3"] = "Кольори:"; db["A3"].font = font(bold=True)
for col, txt, (bg, tx) in (("B", "ДОБРЕ", (GOOD_F, GOOD_T)), ("C", "СЕРЕДНЬО", (MID_F, MID_T)), ("D", "ПОГАНО", (BAD_F, BAD_T))):
    c = put(db, f"{col}3", txt, bold=True, fl=fill(bg), color=tx, align="center")
db["E3"] = "Пороги кольорів — на аркуші «Налаштування»"; db["E3"].font = font(italic=True, color="999999", size=9)
M, Pm = "$B$2", "$D$2"

# --- розміщення секцій ---
K0 = 5            # 1. KPI
C0 = K0 + 7       # 2. перевірки
CHECKS = 8
PL0 = C0 + CHECKS + 4   # заголовок P&L
NP = PL0 + 17     # рядок чистого прибутку

# 1. KPI
section(db, K0, "1. ГОЛОВНЕ ЗА 5 СЕКУНД")
header(db, K0 + 1, ["Показник", "Виручка", "Чистий прибуток", "Рентабельність", "Візитів (виконано)", "Нових пацієнтів", "Середній чек",
                    "Завантаження крісел", "Неявки + скасування"], 1, SUB)
db.row_dimensions[K0 + 1].height = 30
rT, rP, rC, rS = K0 + 2, K0 + 3, K0 + 4, K0 + 5
for rr, lbl in ((rT, "Цей місяць"), (rP, "Минулий місяць"), (rC, "Зміна"), (rS, "ОЦІНКА")):
    put(db, f"A{rr}", lbl, bold=rr in (rT, rS))
def kpi(m, col):
    return [f"={rev(m)}", f"={col}{NP}", f"=IFERROR({col}{NP}/{col}{PL0+1},0)", f"={cnt(m)}", f"={newp(m)}", f"=IFERROR({rev(m)}/{cnt(m)},0)",
            f"=IFERROR({hrs(m)}/{cap(m)},0)",
            f'={noshow(m)}']
kfmt = [ZL, ZL, PCT, NUM, NUM, ZL, PCT, PCT]
PPF = '+0.0" п.п.";-0.0" п.п.";"0"'
for j, (fc, fp) in enumerate(zip(kpi(M, "B"), kpi(Pm, "C"))):
    col = L(2 + j)
    put(db, f"{col}{rT}", fc, kfmt[j], bold=True, align="center", size=13)
    put(db, f"{col}{rP}", fp, kfmt[j], align="center", color="666666")
    if kfmt[j] == PCT: put(db, f"{col}{rC}", f"=({col}{rT}*100-{col}{rP}*100)/100", PPF.replace('"0"', '"0 п.п."'), align="center")
    else: put(db, f"{col}{rC}", f"=IFERROR({col}{rT}/{col}{rP}-1,0)", '+0.0%;-0.0%;"0%"', align="center")
db[f"D{rC}"].number_format = '+0.0%;-0.0%;"0"'  # п.п. як частка (0,01 = 1 п.п.)
for col in "DHI": db[f"{col}{rC}"].number_format = '[=0]"0";+0.0%;-0.0%'
g, mm = T("chg"); gp, mp = T("chgpp")
stat = {
    "B": st_hi(f"B{rC}", g, mm), "C": f'IF(C{rT}<0,"погано",{st_hi(f"D{rT}", *T("margin"))})', "D": st_hi(f"D{rT}", *T("margin")),
    "E": st_hi(f"E{rC}", g, mm), "F": st_hi(f"F{rC}", g, mm), "G": st_hi(f"G{rC}", g, mm),
    "H": st_hi(f"H{rT}", *T("util")), "I": st_lo(f"I{rT}", *T("noshow"))}
for col, f in stat.items():
    put(db, f"{col}{rS}", "=" + f, bold=True, align="center")
cf_text(db, f"B{rT}:I{rT}", f"B${rS}")
cf_text(db, f"B{rS}:I{rS}", f"B${rS}")
for col in "BCEFG": cf_num(db, f"{col}{rC}", f"{col}{rC}", g, mm)
for col in "DH": cf_num(db, f"{col}{rC}", f"{col}{rC}", gp, mp)
cf_num(db, f"I{rC}", f"I{rC}", f"-{gp}", f"-{mp}", higher=False)

# 2. Перевірки
section(db, C0, "2. ЩО ПОТРІБНО ДОПОВНИТИ (зелений = все гаразд, червоний = треба виправити)")
header(db, C0 + 1, ["Перевірка", "Кількість", "Що зробити"], 1, SUB)
db.merge_cells(start_row=C0 + 1, start_column=3, end_row=C0 + 1, end_column=11)
checks = [
    ("Банк: списання без категорії", f'=COUNTIFS({BB},{M},{BI},"⚠ без категорії",{BC},"<0")', "Аркуш «Банк»: вибрати «Ручну категорію» або додати правило в «Правила банку»"),
    ("Візити без напряму", f'=COUNTIFS({VB},{M},{VG},"")', "Аркуш «Послуги»: вказати напрям для нових послуг (червоні), потім імпорт оновить візити"),
    ("Візити: лікар без % у довіднику", f'=COUNTIFS({VB},{M},{VO},0)+COUNTIFS({VB},{M},{VF},"")', "Аркуш «Довідники»: вписати лікаря і його % (червоні клітинки)"),
    ("Візити з невідомим статусом", f'=COUNTIFS({VB},{M})-COUNTIFS({VB},{M},{VM},"Виконано")-COUNTIFS({VB},{M},{VM},"Неявка")-COUNTIFS({VB},{M},{VM},"Скасовано")', "Аркуш «Імпорт-налаштування»: додати статус з Felgdent"),
    ("Нові пацієнти без джерела", f'=COUNTIFS({VB},{M},{VD},"Так",{VE},"")', "Аркуш «Візити»: вказати джерело (червоні клітинки) — тоді видно окупність реклами"),
    ("Протерміновані фактури (усі)", f'=COUNTIF({FI},"Прострочено")', "Аркуш «Фактури»: оплатити або вписати дату оплати"),
    ("Ручні витрати з категорією не з довідника", f'=COUNTIFS({EB},{M},{ED},"⚠ немає в довіднику")', "Аркуш «Витрати»: вибрати категорію зі списку"),
    ("Послуги без напряму (аркуш «Послуги»)", "=COUNTIFS('Послуги'!$A$2:$A$500,\"<>\",'Послуги'!$B$2:$B$500,\"\")", "Аркуш «Послуги»: вибрати напрям"),
]
for i, (lbl, f, todo) in enumerate(checks):
    rr = C0 + 2 + i
    put(db, f"A{rr}", lbl); put(db, f"B{rr}", f, NUM, bold=True, align="center")
    put(db, f"C{rr}", todo); db.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=11)
    db[f"C{rr}"].font = font(size=9, color="444444")
cf_zero_good(db, f"A{C0+2}:B{C0+1+CHECKS}", f"$B{C0+2}")

# 3. P&L
section(db, PL0 - 1, "3. ДЕ ГРОШІ: ПРИБУТКИ ТА ЗБИТКИ (P&L)")
header(db, PL0, ["Стаття", "Цей місяць", "Минулий місяць", "Зміна, zł", "% від виручки", "Норма, %", "Оцінка"], 1, SUB)
lines = [("Виручка (виконані послуги)", lambda m: rev(m), None, "rev"),
         ("− Лабораторія (технік)", lambda m: f"{lab(m)}+{labF(m)}", "Лабораторія (технік)", "c"),
         ("− Спецпрепарати / імпланти", lambda m: prep(m), None, "c"),
         ("− Виплати лікарям і гігієністкам", lambda m: f"{payout(m)}-{labF_pct(m)}", "Лікарі і гігієністки", "c"),
         ("МАРЖА ПІСЛЯ ПРЯМИХ ВИТРАТ", None, None, "gm"),
         ("− Персонал погодинно (асистентки, рецепція)", lambda m: staff_cost(m), "Персонал погодинно", "c")]
for gname in ["Матеріали", "Приміщення", "Обладнання", "Маркетинг", "Адміністрування", "Фінансові витрати", "Податки і ZUS", "Власник (виплата собі)"]:
    lines.append((f"− {gname}", (lambda gg: (lambda m: grp(m, f'"{gg}"')))(gname), None if gname.startswith("Власник") else gname, "c"))
lines += [("− Не розподілено (⚠ без категорії)", lambda m: unalloc(m), None, "u"), ("Усього витрат", None, None, "tc"),
          ("ЧИСТИЙ ПРИБУТОК", None, None, "np"), ("Рентабельність", None, None, "pct")]
r0 = PL0 + 1
cost_rows = [r0 + i for i, ln in enumerate(lines) if ln[3] in ("c", "u")]
cg, cm = T("costx")
for i, (lbl, fn, gname, kind) in enumerate(lines):
    rr = r0 + i
    bold = kind in ("rev", "gm", "np", "tc", "pct")
    fl_ = TOT if kind in ("gm", "np") else None
    put(db, f"A{rr}", lbl, bold=bold, fl=fl_)
    for col, m in (("B", M), ("C", Pm)):
        if kind in ("rev", "c", "u"): f = f"={fn(m)}"
        elif kind == "gm": f = f"={col}{r0}-{col}{r0+1}-{col}{r0+2}-{col}{r0+3}"
        elif kind == "tc": f = "=" + "+".join(f"{col}{x}" for x in cost_rows)
        elif kind == "np": f = f"={col}{r0}-{col}{rr-1}"
        else: f = f"=IFERROR({col}{rr-1}/{col}{r0},0)"
        put(db, f"{col}{rr}", f, PCT if kind == "pct" else ZL, bold=bold, fl=fl_)
    if kind == "pct":
        put(db, f"D{rr}", f"=B{rr}-C{rr}", '+0.0%;-0.0%;"0"', fl=fl_); put(db, f"E{rr}", None)
        put(db, f"F{rr}", f'="ціль "&ROUND({T("margin")[0]}*100,0)&"%"', align="center")
        put(db, f"G{rr}", "=" + st_hi(f"B{rr}", *T("margin")), bold=True, align="center")
    else:
        put(db, f"D{rr}", f"=B{rr}-C{rr}", ZL, fl=fl_)
        put(db, f"E{rr}", f"=IFERROR(B{rr}/$B${r0},0)", PCT, fl=fl_)
        if gname:
            gi = f'MATCH("{gname}",{GRP_N},0)'
            mn, mx = f"INDEX({D}$N$4:$N$15,{gi})", f"INDEX({D}$O$4:$O$15,{gi})"
            put(db, f"F{rr}", f'=IFERROR(ROUND({mn}*100,0)&"–"&ROUND({mx}*100,0)&"%","")', align="center")
            put(db, f"G{rr}", f'=IFERROR(IF($B${r0}=0,"",IF(E{rr}<={mx}*{cg},"добре",IF(E{rr}<={mx}*{cm},"середньо","погано"))),"")', bold=True, align="center")
        elif kind == "u":
            put(db, f"F{rr}", "0 zł", align="center"); put(db, f"G{rr}", f'=IF(B{rr}>0,"погано","добре")', bold=True, align="center")
        elif kind == "np":
            put(db, f"F{rr}", None, fl=fl_); put(db, f"G{rr}", f'=IF(B{rr}<0,"погано",G{rr+1})', bold=True, align="center")
        elif kind in ("rev", "gm"):
            put(db, f"F{rr}", "до мин. міс.", align="center", fl=fl_)
            put(db, f"G{rr}", "=" + st_hi(f"IFERROR(B{rr}/C{rr}-1,0)", g, mm), bold=True, align="center")
        else:
            put(db, f"F{rr}", None, fl=fl_); put(db, f"G{rr}", None, fl=fl_)
PL_END = r0 + len(lines) - 1
assert r0 + [ln[3] for ln in lines].index("np") == NP
cf_text(db, f"E{r0}:G{PL_END}", f"$G{r0}")
cf_text(db, f"A{NP}:B{NP+1}", f"$G{NP}")

# 4. Лікарі
DR0 = PL_END + 3
section(db, DR0 - 1, "4. ХТО З ЛІКАРІВ ВИГІДНИЙ (крісел лише 2 → головне: прибуток клініки за годину крісла)")
header(db, DR0, ["Лікар / гігієністка", "Візитів", "Нових пацієнтів", "Виручка", "Середній чек", "Годин у кріслі", "Виручка за год",
                 "Виплата лікарю", "Прибуток клініки*", "Прибуток клініки за год", "Неявки / скас.", "Ранг", "Оцінка"], 1, SUB)
db.row_dimensions[DR0].height = 42
NDOC = 12
DRT = DR0 + NDOC + 1
rg, rm = T("rel")
for i in range(NDOC):
    rr = DR0 + 1 + i; nm = f"$A{rr}"; x = f",{VF},{nm}"
    put(db, f"A{rr}", f'=IF({D}A{4+i}="","",{D}A{4+i})', bold=True)
    pct = f"N(IFERROR(INDEX({DOC_P},MATCH({nm},{DOC_N},0))*1,0))"
    fs = [cnt(M, x), newp(M, x), rev(M, x), f"IFERROR(D{rr}/B{rr},0)", hrs(M, x), f"IFERROR(D{rr}/F{rr},0)",
          f"{payout(M, x)}-{labF(M, nm)}*{pct}", f"{marg(M, x)}-{labF(M, nm)}*(1-{pct})",
          f'IF(F{rr}=0,"",I{rr}/F{rr})', f'COUNTIFS({VB},{M},{VF},{nm},{VM},"Неявка")+COUNTIFS({VB},{M},{VF},{nm},{VM},"Скасовано")',
          f'IF(J{rr}="","",RANK(J{rr},$J${DR0+1}:$J${DR0+NDOC}))', st_hi(f"IFERROR(J{rr}/$J${DRT},\"\")", rg, rm)]
    fm = [NUM, NUM, ZL, ZL, H1, ZL, ZL, ZL, ZL, NUM, "0", None]
    for j, (f, fmt) in enumerate(zip(fs, fm)):
        col = L(2 + j)
        put(db, f"{col}{rr}", f'=IF($A{rr}="","",{f})', fmt, bold=col == "M", align="center" if col in "BCFKLM" else None)
put(db, f"A{DRT}", "РАЗОМ / СЕРЕДНЄ", bold=True, fl=TOT)
for col, fmt in zip("BCDEFGHIJKLM", [NUM, NUM, ZL, ZL, H1, ZL, ZL, ZL, ZL, NUM, None, None]):
    rng = f"{col}{DR0+1}:{col}{DR0+NDOC}"
    f = {"E": f"=IFERROR(D{DRT}/B{DRT},0)", "G": f"=IFERROR(D{DRT}/F{DRT},0)", "J": f"=IFERROR(I{DRT}/F{DRT},0)", "L": None, "M": None}.get(col, f"=SUM({rng})")
    put(db, f"{col}{DRT}", f, fmt, bold=True, fl=TOT)
cf_text(db, f"A{DR0+1}:A{DR0+NDOC}", f"$M{DR0+1}"); cf_text(db, f"I{DR0+1}:J{DR0+NDOC}", f"$M{DR0+1}"); cf_text(db, f"M{DR0+1}:M{DR0+NDOC}", f"$M{DR0+1}")
db[f"A{DRT+1}"] = "* Прибуток клініки = вартість − технік − спецпрепарати − виплата лікарю (до загальних витрат). Оцінка: порівняно із середнім по клініці."
db[f"A{DRT+1}"].font = font(italic=True, size=8, color="666666")

# 5. Напрями
DI0 = DRT + 4
section(db, DI0 - 1, "5. ЯКІ ПОСЛУГИ (НАПРЯМИ) ПРИНОСЯТЬ ГРОШІ")
header(db, DI0, ["Напрям", "Візитів", "Виручка", "Технік + препарати", "Виплата лікарям", "Маржа клініки", "Маржа, %", "Частка виручки",
                 "Годин у кріслі", "Маржа за год", "Ранг", "Оцінка"], 1, SUB)
db.row_dimensions[DI0].height = 32
NDIR = 10
DIT = DI0 + NDIR + 1
for i in range(NDIR):
    rr = DI0 + 1 + i; x = f",{VG},$A{rr}"
    put(db, f"A{rr}", f'=IF({D}F{4+i}="","",{D}F{4+i})', bold=True)
    fs = [cnt(M, x), rev(M, x), f"{lab(M, x)}+{prep(M, x)}", payout(M, x), marg(M, x), f"IFERROR(F{rr}/C{rr},0)",
          f"IFERROR(C{rr}/$C${DIT},0)", hrs(M, x), f'IF(I{rr}=0,"",F{rr}/I{rr})', f'IF(J{rr}="","",RANK(J{rr},$J${DI0+1}:$J${DI0+NDIR}))',
          st_hi(f"IFERROR(J{rr}/$J${DIT},\"\")", rg, rm)]
    fm = [NUM, ZL, ZL, ZL, ZL, PCT, PCT, H1, ZL, "0", None]
    for j, (f, fmt) in enumerate(zip(fs, fm)):
        col = L(2 + j)
        put(db, f"{col}{rr}", f'=IF($A{rr}="","",{f})', fmt, bold=col == "L", align="center" if col in "BIKL" else None)
put(db, f"A{DIT}", "РАЗОМ / СЕРЕДНЄ", bold=True, fl=TOT)
for col, fmt in zip("BCDEFGHIJKL", [NUM, ZL, ZL, ZL, ZL, PCT, PCT, H1, ZL, None, None]):
    rng = f"{col}{DI0+1}:{col}{DI0+NDIR}"
    f = {"G": f"=IFERROR(F{DIT}/C{DIT},0)", "H": f"=IFERROR(C{DIT}/C{DIT},0)", "J": f"=IFERROR(F{DIT}/I{DIT},0)", "K": None, "L": None}.get(col, f"=SUM({rng})")
    put(db, f"{col}{DIT}", f, fmt, bold=True, fl=TOT)
cf_text(db, f"A{DI0+1}:A{DI0+NDIR}", f"$L{DI0+1}"); cf_text(db, f"J{DI0+1}:J{DI0+NDIR}", f"$L{DI0+1}"); cf_text(db, f"L{DI0+1}:L{DI0+NDIR}", f"$L{DI0+1}")
cf_num(db, f"G{DI0+1}:G{DI0+NDIR}", f"G{DI0+1}", *T("dirm"))

# 6. Джерела і реклама
SO0 = DIT + 4
section(db, SO0 - 1, "6. ЗВІДКИ ПАЦІЄНТИ І ЧИ ОКУПАЄТЬСЯ РЕКЛАМА")
header(db, SO0, ["Джерело", "Нових пацієнтів", "Візитів (усі)", "Виручка (усі візити)", "Рекламні витрати", "Вартість 1 нового пацієнта",
                 "Виручка на 1 zł реклами", "Онлайн-записів"], 1, SUB)
db.row_dimensions[SO0].height = 32
NSRC = 12
SOT = SO0 + NSRC + 1
for i in range(NSRC):
    rr = SO0 + 1 + i; x = f",{VE},$A{rr}"
    put(db, f"A{rr}", f'=IF({D}J{4+i}="","",{D}J{4+i})', bold=True)
    adv = f'IF(AND($A{rr}="Google",{ads_cost(M)}>0),{ads_cost(M)},IF({D}K{4+i}="",0,{cat_cost(M, f"{D}K{4+i}")}))'
    fs = [newp(M, x), cnt(M, x), rev(M, x), adv, f'IF(OR(E{rr}=0,B{rr}=0),"",E{rr}/B{rr})', f'IF(E{rr}=0,"",D{rr}/E{rr})',
          f'COUNTIFS({OB},{M},{OC},$A{rr})']
    fm = [NUM, NUM, ZL, ZL, ZL, '0.0"x"', NUM]
    for j, (f, fmt) in enumerate(zip(fs, fm)):
        col = L(2 + j)
        put(db, f"{col}{rr}", f'=IF($A{rr}="","",{f})', fmt, align="center" if col in "BCGH" else None)
put(db, f"A{SOT}", "РАЗОМ", bold=True, fl=TOT)
for col, fmt in zip("BCDEFGH", [NUM, NUM, ZL, ZL, ZL, '0.0"x"', NUM]):
    rng = f"{col}{SO0+1}:{col}{SO0+NSRC}"
    f = {"F": f'=IF(OR(E{SOT}=0,B{SOT}=0),"",E{SOT}/B{SOT})', "G": f'=IF(E{SOT}=0,"",D{SOT}/E{SOT})'}.get(col, f"=SUM({rng})")
    put(db, f"{col}{SOT}", f, fmt, bold=True, fl=TOT)
cf_num(db, f"F{SO0+1}:F{SOT}", f"F{SO0+1}", *T("cac"), higher=False)
cf_num(db, f"G{SO0+1}:G{SOT}", f"G{SO0+1}", *T("roas"))
db[f"A{SOT+1}"] = "Витрати на Google беруться зі звіту Google Ads (якщо завантажено), інші — з банку та ручних витрат за категорією з «Довідників»."
db[f"A{SOT+1}"].font = font(italic=True, size=8, color="666666")
# Google Ads + платформи
GA0 = SOT + 3
header(db, GA0, ["Google Ads", "Цей місяць", "Минулий місяць"], 1, SUB)
header(db, GA0, ["Онлайн-платформи", "Записів", "Скасовано / неявки", "Нових пацієнтів", "Вартість платформи", "Ціна нового пацієнта"], 5, SUB)
db.row_dimensions[GA0].height = 30
def gnew(m): return f'COUNTIFS({VB},{m},{VM},{OK},{VD},"Так",{VE},"Google")'
def grev(m): return f'SUMIFS({VI},{VB},{m},{VM},{OK},{VE},"Google")'
ga_rows = [("Витрати (звіт Google Ads)", ads_cost, ZL), ("Кліки", lambda m: f"SUMIFS({GE},{GB},{m})", NUM),
           ("Ціна кліку", lambda m: f"IFERROR({ads_cost(m)}/SUMIFS({GE},{GB},{m}),0)", ZL2),
           ("Конверсії (дзвінки / заявки)", lambda m: f"SUMIFS({GG},{GB},{m})", NUM),
           ("Ціна конверсії", lambda m: f"IFERROR({ads_cost(m)}/SUMIFS({GG},{GB},{m}),0)", ZL),
           ("Нових пацієнтів з Google", gnew, NUM),
           ("Вартість нового пацієнта", lambda m: f'IFERROR({ads_cost(m)}/{gnew(m)},"")', ZL),
           ("Виручка на 1 zł реклами", lambda m: f'IFERROR({grev(m)}/{ads_cost(m)},"")', '0.0"x"')]
for i, (lbl, fn, fmt) in enumerate(ga_rows):
    rr = GA0 + 1 + i
    put(db, f"A{rr}", lbl); put(db, f"B{rr}", f"={fn(M)}", fmt, align="center"); put(db, f"C{rr}", f"={fn(Pm)}", fmt, align="center")
cf_num(db, f"B{GA0+7}:C{GA0+7}", f"B{GA0+7}", *T("cac"), higher=False)
cf_num(db, f"B{GA0+8}:C{GA0+8}", f"B{GA0+8}", *T("roas"))
for i, plat in enumerate(["ZnanyLekarz", "Booksy"]):
    rr = GA0 + 1 + i
    put(db, f"E{rr}", plat, bold=True)
    put(db, f"F{rr}", f'=COUNTIFS({OB},{M},{OC},E{rr})', NUM, align="center")
    put(db, f"G{rr}", f'=COUNTIFS({OB},{M},{OC},E{rr},{OG},"Скасовано")+COUNTIFS({OB},{M},{OC},E{rr},{OG},"Неявка")', NUM, align="center")
    put(db, f"H{rr}", f'=COUNTIFS({VB},{M},{VM},{OK},{VD},"Так",{VE},E{rr})', NUM, align="center")
    put(db, f"I{rr}", f'={cat_cost(M, f"E{rr}")}', ZL)
    put(db, f"J{rr}", f'=IF(OR(I{rr}=0,H{rr}=0),"",I{rr}/H{rr})', ZL)
cf_num(db, f"J{GA0+1}:J{GA0+2}", f"J{GA0+1}", *T("cac"), higher=False)
cf_rules(db, f"G{GA0+1}:G{GA0+2}", f"AND(ISNUMBER(G{GA0+1}),G{GA0+1}<=F{GA0+1}*{T('noshow')[0]})",
         f"AND(G{GA0+1}>F{GA0+1}*{T('noshow')[0]},G{GA0+1}<=F{GA0+1}*{T('noshow')[1]})", f"G{GA0+1}>F{GA0+1}*{T('noshow')[1]}")
GA_END = GA0 + len(ga_rows)

# 7. Крісла і пацієнти
CH0 = GA_END + 3
section(db, CH0 - 1, "7. КРІСЛА ПРОСТОЮЮТЬ? ПАЦІЄНТИ ПОВЕРТАЮТЬСЯ?")
header(db, CH0, ["Показник", "Цей місяць", "Минулий місяць"], 1, SUB)
ch = [("Доступно годин крісел (крісла × години × дні)", lambda m: cap(m), H1),
      ("Зайнято годин (виконані візити)", lambda m: hrs(m), H1),
      ("Завантаження крісел, %", lambda m: f"IFERROR({hrs(m)}/{cap(m)},0)", PCT),
      ("Вільних годин крісел", lambda m: f"{cap(m)}-{hrs(m)}", H1),
      ("Неявки + скасування, %", lambda m: noshow(m), PCT),
      ("Втрачена виручка через неявки/скасування", lambda m: f'SUMIFS({VI},{VB},{m},{VM},"Неявка")+SUMIFS({VI},{VB},{m},{VM},"Скасовано")', ZL),
      ("ПОТЕНЦІАЛ: +маржа, якщо заповнити 25% вільних годин", lambda m: f"IFERROR(({cap(m)}-{hrs(m)})*0.25*{marg(m)}/{hrs(m)},0)", ZL),
      ("Частка візитів повторних пацієнтів, %", lambda m: f'IFERROR(COUNTIFS({VB},{m},{VM},{OK},{VD},"Ні")/{cnt(m)},0)', PCT),
      ("Перенесено візитів (звіт Felgdent)", lambda m: f'COUNTIFS({ZB},{m},{ZD},"Перенесено")', NUM),
      ("Неявки (звіт Felgdent «Status wizyt»)", lambda m: f'SUMIFS({FSD},{FSA},{m},{FSC},"Неявка")', NUM)]
for i, (lbl, fn, fmt) in enumerate(ch):
    rr = CH0 + 1 + i
    pot = lbl.startswith("ПОТЕНЦІАЛ")
    put(db, f"A{rr}", lbl, bold=pot, fl=TOT if pot else None).alignment = Alignment(wrap_text=True, vertical="center")
    put(db, f"B{rr}", f"={fn(M)}", fmt, align="center", bold=pot, fl=TOT if pot else None)
    put(db, f"C{rr}", f"={fn(Pm)}", fmt, align="center")
    db.row_dimensions[rr].height = 26
cf_num(db, f"B{CH0+3}:C{CH0+3}", f"B{CH0+3}", *T("util"))
cf_num(db, f"B{CH0+5}:C{CH0+5}", f"B{CH0+5}", *T("noshow"), higher=False)
cf_num(db, f"B{CH0+8}:C{CH0+8}", f"B{CH0+8}", *T("repeat"))
header(db, CH0, ["Спосіб оплати", "Сума", "Частка"], 5, SUB)
for i in range(10):
    rr = CH0 + 1 + i
    put(db, f"E{rr}", f'=IF({D}H{4+i}="","",{D}H{4+i})')
    put(db, f"F{rr}", f'=IF(E{rr}="","",SUMIFS({VI},{VB},{M},{VM},{OK},{VL},E{rr}))', ZL)
    put(db, f"G{rr}", f'=IF(E{rr}="","",IFERROR(F{rr}/$B${r0},0))', PCT)
CHE = CH0 + 10

# 8. Банк і звірки
BK0 = CHE + 3
section(db, BK0 - 1, "8. ГРОШІ НА РАХУНКУ, ФАКТУРИ І ЗВІРКИ")
header(db, BK0, ["Рух грошей по банку", "Цей місяць", "Минулий місяць"], 1, SUB)
inc = ["Надходження: термінал", "Надходження: PayU raty", "Надходження: Mediraty", "Надходження: BLIK / перекази пацієнтів",
       "Надходження: готівка на рахунок", "Надходження: інше"]
for i, c_ in enumerate(inc):
    rr = BK0 + 1 + i
    put(db, f"A{rr}", c_.replace("Надходження: ", "+ ")); put(db, f"B{rr}", f'={bank_cat(M, f""""{c_}" """.strip())}', ZL); put(db, f"C{rr}", f'={bank_cat(Pm, f""""{c_}" """.strip())}', ZL)
rIn = BK0 + 1 + len(inc)
put(db, f"A{rIn}", "Усього надходжень", bold=True, fl=TOT)
for col in "BC": put(db, f"{col}{rIn}", f"=SUM({col}{BK0+1}:{col}{rIn-1})", ZL, bold=True, fl=TOT)
put(db, f"A{rIn+1}", "Усього списань", bold=True)
put(db, f"B{rIn+1}", f'=-SUMIFS({BC},{BB},{M},{BC},"<0")', ZL, bold=True); put(db, f"C{rIn+1}", f'=-SUMIFS({BC},{BB},{Pm},{BC},"<0")', ZL, bold=True)
put(db, f"A{rIn+2}", "САЛЬДО МІСЯЦЯ (надходження − списання)", bold=True)
for col in "BC": put(db, f"{col}{rIn+2}", f"={col}{rIn}-{col}{rIn+1}", ZL, bold=True)
cf_rules(db, f"A{rIn+2}:C{rIn+2}", f"$B{rIn+2}>=0", None, f"$B{rIn+2}<0")
# звірки
header(db, BK0, ["Звірка", "Нараховано (візити)", "Оплачено з банку", "Розбіжність", "Оцінка"], 5, SUB)
recon = [("Лабораторія", f"{lab(M)}+{labF(M)}", f'-{bank_cat(M, chr(34)+"Laboratorium – płatność (звірка)"+chr(34))}'),
         ("Імпланти / препарати", prep(M), f'-{bank_cat(M, chr(34)+"Implanty / preparaty – płatność (звірка)"+chr(34))}'),
         ("Лікарі + персонал", f"{payout(M)}+{staff_cost(M)}", f'-{bank_cat(M, chr(34)+"Wypłaty lekarzy i personelu (звірка)"+chr(34))}'),
         ("Безготівкова виручка → рахунок", noncash(M), f"B{rIn}-B{BK0+5}-B{BK0+6}")]
rcg, rcm = T("recon")
for i, (lbl, a, b_) in enumerate(recon):
    rr = BK0 + 1 + i
    put(db, f"E{rr}", lbl, bold=True); put(db, f"F{rr}", f"={a}", ZL); put(db, f"G{rr}", f"={b_}", ZL)
    put(db, f"H{rr}", f"=IFERROR(G{rr}/F{rr}-1,0)", '+0.0%;-0.0%;"0%"', align="center")
    put(db, f"I{rr}", f'=IF(AND(F{rr}=0,G{rr}=0),"",{st_lo(f"ABS(H{rr})", rcg, rcm)})', bold=True, align="center")
cf_text(db, f"E{BK0+1}:I{BK0+4}", f"$I{BK0+1}")
db[f"E{BK0+5}"] = "Невелика розбіжність нормальна: лабораторія та зарплати часто оплачуються наступного місяця."
db[f"E{BK0+5}"].font = font(italic=True, size=8, color="666666")
# фактури
FA0 = BK0 + 7
header(db, FA0, ["Фактури", "Кількість", "Сума"], 5, SUB)
fa = [("Прострочені", "Прострочено"), ("Очікують оплати", "Очікує оплати"), ("Без терміну, не оплачені", "Не оплачено")]
for i, (lbl, s_) in enumerate(fa):
    rr = FA0 + 1 + i
    put(db, f"E{rr}", lbl, bold=True); put(db, f"F{rr}", f'=COUNTIF({FI},"{s_}")', NUM, align="center"); put(db, f"G{rr}", f'=SUMIFS({FE},{FI},"{s_}")', ZL)
cf_rules(db, f"E{FA0+1}:G{FA0+1}", f"$F{FA0+1}=0", None, f"$F{FA0+1}>0")
cf_rules(db, f"E{FA0+2}:G{FA0+3}", f"$F{FA0+2}=0", f"$F{FA0+2}>0", None)
BKE = max(rIn + 2, FA0 + 3)

# 9. Підказки
HI0 = BKE + 3
section(db, HI0 - 1, "9. АВТОМАТИЧНІ ПІДКАЗКИ: ДЕ ОПТИМІЗУВАТИ")
dJ, dA = f"$J${DR0+1}:$J${DR0+NDOC}", f"$A${DR0+1}:$A${DR0+NDOC}"
iJ, iA = f"$J${DI0+1}:$J${DI0+NDIR}", f"$A${DI0+1}:$A${DI0+NDIR}"
hints = [
    (f'=IFERROR("Найвигідніший лікар (прибуток / год крісла): "&INDEX({dA},MATCH(MAX({dJ}),{dJ},0))&" — "&TEXT(MAX({dJ}),"0")&" zł/год","")', '"добре"'),
    (f'=IFERROR("Найменш вигідний: "&INDEX({dA},MATCH(MIN({dJ}),{dJ},0))&" — "&TEXT(MIN({dJ}),"0")&" zł/год. Перевірте ціни послуг, витрату матеріалів і % виплати.","")', '"погано"'),
    (f'=IFERROR("Найвигідніший напрям: "&INDEX({iA},MATCH(MAX({iJ}),{iJ},0))&"; найменш вигідний: "&INDEX({iA},MATCH(MIN({iJ}),{iJ},0))&". Віддавайте кращі години крісел вигіднішим послугам.","")', '"середньо"'),
    (f'=IF(H{rT}<{T("util")[0]},"Крісла завантажені на "&TEXT(H{rT}*100,"0")&"%. Вільних годин: "&TEXT(B{CH0+4},"0")&". Якщо заповнити чверть з них: +"&TEXT(B{CH0+7},"0")&" zł маржі.","Завантаження крісел в нормі.")', f'H{rS}'),
    (f'=IF(D{rT}<{T("margin")[0]},"Рентабельність "&TEXT(D{rT}*100,"0")&"% нижче цілі "&TEXT({T("margin")[0]}*100,"0")&"%. Дивіться червоні й оранжеві рядки в розділі 3.","Рентабельність в нормі.")', f'D{rS}'),
    (f'=IF(I{rT}>{T("noshow")[0]},"Неявок і скасувань "&TEXT(I{rT}*100,"0")&"%: SMS-нагадування за 24 год і передоплата за довгі візити.","Неявки в нормі.")', f'I{rS}'),
    (f'=IF(B{C0+2}+B{C0+3}+B{C0+4}+B{C0+5}>0,"Дані неповні: виправте червоні рядки в розділі 2, інакше цифри неточні.","Дані повні.")', f'IF(B{C0+2}+B{C0+3}+B{C0+4}+B{C0+5}>0,"погано","добре")'),
]
for i, (h, s_) in enumerate(hints):
    rr = HI0 + i
    put(db, f"L{rr}", "=" + s_, border=False).font = font(color="FFFFFF", size=6)
    db.merge_cells(f"A{rr}:K{rr}")
    c = db[f"A{rr}"]; c.value = h; c.font = font(); c.alignment = Alignment(wrap_text=True, vertical="center")
    db.row_dimensions[rr].height = 22
    cf_text(db, f"A{rr}:K{rr}", f"$L{rr}")
LAST = HI0 + len(hints)

def bar(title, col, r1, n, anchor, w):
    ch_ = BarChart(); ch_.type = "bar"; ch_.title = title
    ch_.add_data(Reference(db, min_col=col, min_row=r1, max_row=r1 + n), titles_from_data=True)
    ch_.set_categories(Reference(db, min_col=1, min_row=r1 + 1, max_row=r1 + n))
    ch_.legend = None; ch_.roundedCorners = False; ch_.height = 8; ch_.width = w
    ch_.x_axis.delete = False; ch_.y_axis.delete = False; ch_.title.overlay = False
    db.add_chart(ch_, anchor)
bar("Прибуток клініки за годину крісла, zł", 10, DR0, len(docs), f"A{LAST+4}", 18)
bar("Маржа за годину крісла по напрямах, zł", 10, DI0, len(dirs), f"F{LAST+4}", 17)
db.print_area = f"A1:M{LAST+21}"
db.page_setup.orientation = "portrait"; db.page_setup.paperSize = 9
db.page_setup.fitToWidth = 1; db.page_setup.fitToHeight = 0
db.sheet_properties.pageSetUpPr.fitToPage = True
db.page_margins.left = db.page_margins.right = 0.4

# ========================= РІК =========================
yr = wb.create_sheet("Рік", 1)
widths(yr, {"A": 36, **{L(2 + i): 11 for i in range(13)}})
yr.sheet_view.showGridLines = False
YTH = th_helper(yr, 17, 4)
yr["A1"] = "ДИНАМІКА ЗА РІК"; yr["A1"].font = font(bold=True, size=16, color="1F4E5A")
yr["A2"] = "Рік:"; yr["B2"] = f"=YEAR({MCUR})"; yr["B2"].font = font(bold=True, size=12)
header(yr, 4, ["Показник"] + [""] * 12 + ["Разом / сер."], 1, SUB)
for k in range(12):
    c = yr[f"{L(2+k)}4"]; c.value = f"=DATE($B$2,{k+1},1)"; c.number_format = MON
yl = [("Виручка", lambda m: rev(m), ZL, "sum"), ("Прямі витрати (технік, препарати, лікарі)", lambda m: f"{lab(m)}+{prep(m)}+{payout(m)}-{labF_pct(m)}", ZL, "sum"),
      ("Персонал погодинно", lambda m: staff_cost(m), ZL, "sum"), ("Інші витрати (банк + ручні)", lambda m: opex_all(m), ZL, "sum"),
      ("Чистий прибуток", None, ZL, "sum"), ("Рентабельність", None, PCT, "avg"), ("Візитів (виконано)", lambda m: cnt(m), NUM, "sum"),
      ("Нових пацієнтів", lambda m: newp(m), NUM, "sum"), ("Середній чек", None, ZL, "avg"),
      ("Завантаження крісел", lambda m: f"IFERROR({hrs(m)}/{cap(m)},0)", PCT, "avg"), ("Маркетинг", lambda m: grp(m, '"Маркетинг"'), ZL, "sum")]
for i, (lbl, fn, fmt, agg) in enumerate(yl):
    rr = 5 + i
    put(yr, f"A{rr}", lbl, bold=lbl in ("Виручка", "Чистий прибуток"), fl=TOT if lbl == "Чистий прибуток" else None)
    for k in range(12):
        col = L(2 + k); m = f"{col}$4"
        if lbl == "Чистий прибуток": f = f"={col}5-{col}6-{col}7-{col}8"
        elif lbl == "Рентабельність": f = f"=IF({col}5=0,\"\",{col}9/{col}5)"
        elif lbl == "Середній чек": f = f"=IFERROR({col}5/{col}11,0)"
        elif lbl == "Завантаження крісел": f = f'=IF({col}5=0,"",{fn(m)})'
        else: f = f"={fn(m)}"
        put(yr, f"{col}{rr}", f, fmt)
    if lbl == "Рентабельність": f = "=IFERROR(N9/N5,0)"
    elif lbl == "Середній чек": f = "=IFERROR(N5/N11,0)"
    elif agg == "sum": f = f"=SUM(B{rr}:M{rr})"
    else: f = f'=IFERROR(AVERAGEIF(B{rr}:M{rr},">0"),0)'
    put(yr, f"N{rr}", f, fmt, bold=True, fl=TOT)
cf_rules(yr, "B9:N9", 'AND(B$5<>0,B9>=0)', None, 'AND(B$5<>0,B9<0)')
cf_num(yr, "B10:N10", "B10", *YTH["margin"])
cf_num(yr, "B14:N14", "B14", *YTH["util"])
bc = BarChart(); bc.title = "Виручка і чистий прибуток по місяцях, zł"
bc.add_data(Reference(yr, min_col=1, max_col=13, min_row=5), titles_from_data=True, from_rows=True)
bc.add_data(Reference(yr, min_col=1, max_col=13, min_row=9), titles_from_data=True, from_rows=True)
bc.set_categories(Reference(yr, min_col=2, max_col=13, min_row=4))
bc.roundedCorners = False; bc.legend.overlay = False; bc.title.overlay = False
bc.x_axis.delete = False; bc.y_axis.delete = False; bc.x_axis.number_format = MON; bc.height = 9; bc.width = 30
yr.add_chart(bc, "A18")
yr.print_area = "A1:N36"; yr.page_setup.orientation = "landscape"; yr.page_setup.fitToWidth = 1; yr.page_setup.fitToHeight = 0
yr.sheet_properties.pageSetUpPr.fitToPage = True

# ========================= Інструкція =========================
ins = wb.create_sheet("Інструкція", 0)
ins.sheet_view.showGridLines = False
ins.column_dimensions["A"].width = 4; ins.column_dimensions["B"].width = 125
txt = [
    ("ЯК КОРИСТУВАТИСЯ ЗВІТАМИ КЛІНІКИ", "h1"),
    ("Щомісяця за 5 хвилин: скільки заробили, на що пішли гроші, які лікарі й послуги вигідні, чи простоюють крісла, чи окупається реклама.", ""),
    ("", ""),
    ("КОЛЬОРИ", "h2"),
    ("ЗЕЛЕНИЙ = добре", "c_good"), ("ОРАНЖЕВИЙ = середньо, варто придивитися", "c_mid"), ("ЧЕРВОНИЙ = погано, потрібна дія", "c_bad"),
    ("Пороги кольорів змінюються на аркуші «Налаштування» (таблиця «Пороги кольорів»).", ""),
    ("БЛІДО-ЖОВТІ клітинки — сюди вписуємо дані", "c_inp"), ("СВІТЛО-СІРІ клітинки — рахуються самі, захищені від випадкового стирання", "c_calc"),
    ("Захист: у Excel — Рецензування → Зняти захист аркуша (без пароля). У Google Таблицях сірі клітинки може змінювати лише власник; адміністратор їх не зітре.", ""),
    ("", ""),
    ("ЯК ЗАВАНТАЖУВАТИ ДАНІ (Google Таблиці, меню «🦷 Звіти клініки»)", "h2"),
    ("Варіант А: меню 🦷 Звіти клініки → «Завантажити файли…» → виберіть один або кілька файлів (CSV, XLSX) → «Імпортувати». Тип файлу визначається автоматично.", ""),
    ("Варіант Б: скиньте файли в папку Google Drive «Звіти клініки – Імпорт» (можна з телефону). Щоночі скрипт їх забирає, а оброблені переносить у підпапку «Оброблено».", ""),
    ("Повторне завантаження того самого файлу НЕ створює дублів: скрипт впізнає рядки, які вже є, і лише оновлює їх (наприклад, статус візиту).", ""),
    ("Кожне завантаження записується на аркуш «Журнал імпорту» разом із попередженнями.", ""),
    ("", ""),
    ("ЯКІ ФАЙЛИ ЗАВАНТАЖУВАТИ", "h2"),
    ("1. Felgdent → Raporty → «Koszty i wpłaty w gabinecie» (за тиждень або місяць): оплати → «Візити» (виручка, послуги, напрям, спосіб оплати); витрати на техніка → «Витрати» з лікарем.", ""),
    ("   ⚠ В оплатах Felgdent немає лікаря. Поки не буде звіту з лікарями, впишіть лікаря в «Візитах» (червоні клітинки) або завантажте звіт «Rozliczenia z lekarzami».", ""),
    ("   Felgdent → «Status wizyt» (за місяць, вкажіть місяць при завантаженні) → неявки та скасування.  «Średnia długość wizyty» → час візиту кожного лікаря в «Довідниках».", ""),
    ("   Felgdent → «Raport pozyskania pacjenta» → джерела пацієнтів.  «Raport przesunięć wizyt» → перенесення й скасування.", ""),
    ("2. Якщо є: список візитів з лікарем і сумою (CSV/XLSX) → «Візити» напряму; список пацієнтів з датою реєстрації → точно видно нових.", ""),
    ("ВАЖЛИВО у Felgdent: при записі НОВОГО пацієнта завжди заповнювати «Skąd pacjent». «Brak informacji» = невідомо, чи окупається реклама.", "warn"),
    ("3. Виписка з банку CSV/XLSX (раз на тиждень) → аркуш «Банк». Категорії підставляються за «Правилами банку»; невпізнані — червоні, їх вибирають вручну один раз.", ""),
    ("4. Реєстр фактур (з KSeF або програми бухгалтерії): номер, дата, продавець, NIP, сума, термін → аркуш «Фактури». Скрипт сам знаходить оплату в банку й ставить «Оплачено».", ""),
    ("5. ZnanyLekarz і Booksy – список записів → аркуш «Онлайн-записи». Пацієнтам з цих записів автоматично ставиться джерело.", ""),
    ("6. Google Ads – звіт кампаній (бажано по днях, або вкажіть місяць при завантаженні) → аркуш «Google Ads».", ""),
    ("ВРУЧНУ: «Витрати» (тільки готівка й те, чого немає в банку), «Персонал» (години погодинних працівників), лабораторія та імпланти у «Візитах», якщо Felgdent їх не віддає.", ""),
    ("", ""),
    ("РЕГЛАМЕНТ (до 2 годин на тиждень)", "h2"),
    ("Щопонеділка, адміністратор (~30 хв): завантажити візити з Felgdent і записи ZnanyLekarz / Booksy; виправити червоні клітинки в «Послугах» і «Візитах».", ""),
    ("Щопонеділка, власник (~30 хв): завантажити виписку з банку і реєстр фактур; вибрати категорію для червоних рядків у «Банку».", ""),
    ("До 10-го числа: години персоналу, звіт Google Ads за минулий місяць. Подивитися розділ 2 «Co trzeba uzupełnić» у звіті «Raport» — усе має бути зеленим.", ""),
    ("15-го о 21:00: скрипт сам ставить попередній місяць і надсилає звіт «Raport» (польською) у PDF плюс кольоровий лист на пошту.", ""),
    ("", ""),
    ("ПРАВИЛА ОБЛІКУ (щоб не рахувати двічі)", "h2"),
    ("• Виручка = вартість ВИКОНАНИХ послуг з «Візитів». Ваучер чи передоплата за імплант стають виручкою, коли послугу виконано.", ""),
    ("• Банк: надходження (термінал, PayU, Mediraty, BLIK) НЕ є виручкою, вони йдуть лише в «Рух грошей» і звірку.", ""),
    ("• Банк: оплати лабораторії, імплантів, зарплати лікарів і персоналу мають категорії «… (звірка)». Вони не додаються до витрат (бо вже пораховані з Візитів / Персоналу), а показуються у звірці.", ""),
    ("• Власна зарплата власника — категорія «Wypłata właściciela». Без неї прибуток здається більшим, ніж є.", ""),
    ("", ""),
    ("ПЕРЕД ПОЧАТКОМ", "h2"),
    ("Файл містить ДЕМО-дані (вересень–жовтень 2026). Щоб очистити: меню 🦷 → «Очистити демо-дані» (або вручну: виділіть колонки вводу й натисніть Delete; цілі рядки не видаляйте).", ""),
    ("Потім: «Довідники» (лікарі й %), «Правила банку» (ваші постачальники, лабораторія, імена лікарів), «Послуги» (послуги з Felgdent → напрям).", ""),
    ("", ""),
    ("ВСТАНОВЛЕННЯ СКРИПТА (один раз, ~10 хв)", "h2"),
    ("1. Google Drive → завантажити цей файл → відкрити → Файл → Зберегти як Google Таблиці.", ""),
    ("2. Розширення → Apps Script. Створіть файли з такими самими назвами, як у папці apps-script (Import.gs, Kod.gs, Raport.gs, Dialog.html), і вставте код.", ""),
    ("3. Шестерня «Налаштування проєкту» → «Показувати файл маніфесту appsscript.json» → вставте вміст appsscript.json (вмикає перетворення XLSX).", ""),
    ("4. У Raport.gs впишіть свою пошту в EMAIL → Зберегти → оновіть таблицю: з'явиться меню «🦷 Звіти клініки».", ""),
    ("5. Меню → «Увімкнути автоматику» → дозволити доступ. Готово: імпорт з папки щоночі, звіт 15-го між 21:00 і 22:00.", ""),
]
for i, (t, kind) in enumerate(txt):
    c = ins.cell(row=1 + i, column=2, value=t)
    c.font = {"h1": font(bold=True, size=16, color="1F4E5A"), "h2": font(bold=True, size=11, color="1F4E5A"), "warn": font(bold=True, color=BAD_T)}.get(kind, font())
    if kind == "warn": c.fill = fill(BAD_F)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    sw = {"c_good": (GOOD_F, GOOD_T), "c_mid": (MID_F, MID_T), "c_bad": (BAD_F, BAD_T), "c_inp": ("FFF7CC", "000000"), "c_calc": ("E7E7E7", "000000")}
    if kind in sw:
        c.fill = fill(sw[kind][0]); c.font = font(bold=True, color=sw[kind][1], size=11); c.border = BOX
ins.page_setup.fitToWidth = 1; ins.page_setup.fitToHeight = 0; ins.sheet_properties.pageSetUpPr.fitToPage = True

# ========================= ПОЛЬСЬКА ВЕРСІЯ ЗВІТІВ =========================
import translations_pl as TPL
from copy import copy as _copy
def tr_formula(f):
    return re.sub(r'"([^"]*)"', lambda m_: '"' + TPL.FORMULA.get(m_.group(1), m_.group(1)) + '"', f)
def clone_pl(src, title):
    dst = wb.copy_worksheet(src); dst.title = title
    for row in dst.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str):
                c.value = tr_formula(v) if v.startswith("=") else TPL.PLAIN.get(v, v)
    for cf in src.conditional_formatting:
        for rule in cf.rules:
            r2 = _copy(rule); r2.formula = [tr_formula(f) for f in rule.formula]
            dst.conditional_formatting.add(str(cf.sqref), r2)
    dst.sheet_view.showGridLines = False
    dst.print_area = src.print_area.split('!')[-1].replace('$', '') if src.print_area else None
    dst.page_setup.orientation = src.page_setup.orientation; dst.page_setup.paperSize = 9
    dst.page_setup.fitToWidth = 1; dst.page_setup.fitToHeight = 0
    dst.sheet_properties.pageSetUpPr.fitToPage = True
    dst.page_margins.left = dst.page_margins.right = 0.4
    dst.sheet_properties.tabColor = "C0392B"
    return dst
rpl = clone_pl(db, "Raport")
for title, col, r1, n, anchor_, w in ((TPL.CHARTS["Прибуток клініки за годину крісла, zł"], 10, DR0, len(docs), f"A{LAST+4}", 18),
                                      (TPL.CHARTS["Маржа за годину крісла по напрямах, zł"], 10, DI0, len(dirs), f"F{LAST+4}", 17)):
    ch_ = BarChart(); ch_.type = "bar"; ch_.title = title
    ch_.add_data(Reference(rpl, min_col=col, min_row=r1, max_row=r1 + n), titles_from_data=True)
    ch_.set_categories(Reference(rpl, min_col=1, min_row=r1 + 1, max_row=r1 + n))
    ch_.legend = None; ch_.roundedCorners = False; ch_.height = 8; ch_.width = w
    ch_.x_axis.delete = False; ch_.y_axis.delete = False; ch_.title.overlay = False
    rpl.add_chart(ch_, anchor_)
ypl = clone_pl(yr, "Rok")
bc2 = BarChart(); bc2.title = TPL.CHARTS["Виручка і чистий прибуток по місяцях, zł"]
bc2.add_data(Reference(ypl, min_col=1, max_col=13, min_row=5), titles_from_data=True, from_rows=True)
bc2.add_data(Reference(ypl, min_col=1, max_col=13, min_row=9), titles_from_data=True, from_rows=True)
bc2.set_categories(Reference(ypl, min_col=2, max_col=13, min_row=4))
bc2.roundedCorners = False; bc2.legend.overlay = False; bc2.title.overlay = False
bc2.x_axis.delete = False; bc2.y_axis.delete = False; bc2.x_axis.number_format = MON; bc2.height = 9; bc2.width = 30
ypl.add_chart(bc2, "A18")
ypl["A5"] = "Przychód"; ypl["A9"] = "Zysk netto"   # підписи рядів графіка
# звіти — лише польською: українські робочі копії видаляємо
for _ws in (rpl, ypl): _ws.sheet_properties.tabColor = "1F4E5A"
wb.remove(db); wb.remove(yr)
# польська інструкція
ipl = wb.create_sheet("Instrukcja")
ipl.sheet_view.showGridLines = False
ipl.column_dimensions["A"].width = 4; ipl.column_dimensions["B"].width = 125
for i, (t, kind) in enumerate(TPL.INSTRUKCJA):
    c = ipl.cell(row=1 + i, column=2, value=t)
    c.font = {"h1": font(bold=True, size=16, color="1F4E5A"), "h2": font(bold=True, size=11, color="1F4E5A"), "warn": font(bold=True, color=BAD_T)}.get(kind, font())
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "warn": c.fill = fill(BAD_F)
    if kind in sw:
        c.fill = fill(sw[kind][0]); c.font = font(bold=True, color=sw[kind][1], size=11); c.border = BOX
ipl.page_setup.fitToWidth = 1; ipl.page_setup.fitToHeight = 0; ipl.sheet_properties.pageSetUpPr.fitToPage = True
ipl.sheet_properties.tabColor = "C0392B"
# двомовні заголовки аркушів введення
for name in ("Візити", "Банк", "Фактури", "Витрати", "Персонал"):
    ws = wb[name]
    for c in ws[1]:
        if isinstance(c.value, str) and c.value in TPL.HEADERS:
            c.value = c.value + "\n" + TPL.HEADERS[c.value]
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[1].height = 58

# ========================= ЖОВТЕ / СІРЕ + ЗАХИСТ =========================
UNLOCK = Protection(locked=False)
INPUT_COLS = {"Візити": ("ACDEFGHIJKLMN", N), "Банк": ("ACDEGJ", N), "Фактури": ("ABCDEFGH", N), "Витрати": ("ACEFGH", N), "Felgdent статуси": ("ABCD", 500), "Записи Felgdent": ("ACDEFG", 5001),
              "Персонал": ("ACDEF", 500), "Онлайн-записи": ("ACDEFG", N), "Google Ads": ("ACDEFG", N), "Пацієнти": ("ABC", N), "Послуги": ("ABC", 500)}
CALC_COLS = {"Візити": ("BOPQ", "R"), "Банк": ("BHI", "FK"), "Фактури": ("IJ", "K"), "Витрати": ("BD", "I"), "Felgdent статуси": ("", "E"), "Записи Felgdent": ("B", "H"), "Персонал": ("BG", ""),
             "Онлайн-записи": ("B", "H"), "Google Ads": ("B", "H"), "Пацієнти": ("", "D")}
def protect(ws):
    pr = ws.protection
    pr.sheet = True
    pr.formatColumns = pr.formatRows = pr.formatCells = pr.autoFilter = pr.sort = False
for name, (cols, last) in INPUT_COLS.items():
    ws = wb[name]
    calc, script = CALC_COLS.get(name, ("", ""))
    for r in range(2, last + 1):
        for col in cols:
            c = ws[f"{col}{r}"]; c.fill = INP; c.protection = UNLOCK
        for col in calc:
            ws[f"{col}{r}"].fill = CALC
        for col in script:   # заповнює скрипт: сіре, але не заблоковане (інакше імпорт від імені адміністратора не спрацює)
            c = ws[f"{col}{r}"]; c.fill = CALC; c.protection = UNLOCK
    protect(ws)
for name in ("Довідники", "Правила банку", "Імпорт-налаштування", "Налаштування"):
    ws = wb[name]
    for row in ws.iter_rows():
        for c in row:
            if c.fill is not None and c.fill.fgColor is not None and c.fill.fgColor.rgb in ("00FFF7CC", "FFFFF7CC", "00" + GOOD_F, "00" + MID_F):
                c.protection = UNLOCK
    protect(ws)
for row in wb["Налаштування"].iter_rows(min_row=13, max_row=12 + len(TH_DEF), min_col=2, max_col=3):
    for c in row: c.protection = UNLOCK
# Дашборд і Рік: усі розрахункові клітинки світло-сірі (кольори світлофора лягають поверх), аркуш повністю захищений
GREY_DASH = PatternFill("solid", fgColor="F0F0F0")
for ws in (wb["Raport"], wb["Rok"]):
    for row in ws.iter_rows(min_row=4, max_col=14):
        for c in row:
            v = c.value
            if v is None or c.coordinate in ws.merged_cells: continue
            if (isinstance(v, str) and v.startswith("=")) or isinstance(v, (int, float)):
                if (c.fill is None or c.fill.fgColor.rgb in ("00000000", None)) and not (c.font and c.font.color and c.font.color.rgb == "00FFFFFF"):
                    c.fill = GREY_DASH
    protect(ws)
protect(wb["Інструкція"]); protect(wb["Instrukcja"])

# порядок аркушів
order = ["Raport", "Rok", "Instrukcja", "Інструкція", "Візити", "Банк", "Фактури", "Витрати", "Персонал", "Онлайн-записи", "Google Ads",
         "Пацієнти", "Felgdent статуси", "Записи Felgdent", "Послуги", "Довідники", "Правила банку", "Імпорт-налаштування", "Налаштування", "Журнал імпорту"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font and c.font.name != F:
                c.font = Font(name=F, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic, color=c.font.color)
for name, color in (("Raport", "1F4E5A"), ("Rok", "1F4E5A"), ("Інструкція", "999999")):
    wb[name].sheet_properties.tabColor = color
for name in ("Візити", "Банк", "Фактури", "Витрати", "Персонал", "Онлайн-записи", "Google Ads", "Пацієнти", "Felgdent статуси", "Записи Felgdent"):
    wb[name].sheet_properties.tabColor = "F2C744"
wb.active = 1
wb.save(OUT)
print("saved", OUT, "visits", len(visits), "bank", len(bank_rows))
