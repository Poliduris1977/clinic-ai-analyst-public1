"""
Шаг 11: аналитика врачей, сравнение врачей и грейдов по выручке и марже за 2025 год в Excel.

Вход:  data/corrected/ - visits, visit_items, services, materials_1c, payroll_1c, doctors
Выход: data/reports/Doctors_2025.xlsx

Две ступени маржи, как в факторном анализе:
  прямая маржа = выручка - материалы - лаборатории (затраты 1С разнесены по нормативу услуг, как в build_factor)
  вклад врача  = прямая маржа - оплата врача (gross, payroll_1c) - страховые взносы
Главный показатель сравнения: вклад на врача-месяц (месяц с записью в payroll), в тыс. ₽.
Процент маржи показан вторым: он зависит от ставки оплаты и сам по себе грейды сравнивать не помогает.

Книга живая: листы считаются формулами SUMIFS по листу «Данные».
Листы:
  Описание  - определения, правила, ограничения, контрольные суммы с источниками
  Врачи     - 43 врача + строка «Без врача» (фильтр по месяцу)
  Грейды    - грейды по сети и внутри каждого направления (фильтр по клинике; месяц берётся с листа «Врачи»)
  Калькулятор - схема оплаты «что если»: ставки по грейдам, порог и надбавка, фикс, сценарий роста выручки
  Расчёт    - построчный расчёт калькулятора (врач x месяц), логика зеркальна common/payscheme.py
  Данные    - врач x месяц, источник для формул

Запуск:  python src/reports/build_doctors.py
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.recalc import recalculate  # noqa: E402
from common import payscheme  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
YEAR = 2025
CELL = ["period", "clinic_id", "direction_code"]
DIRS = {"DENT": "Стоматология", "FAM": "Семейная медицина", "COSM": "Косметология",
        "PLAS": "Пластическая хирургия", "LASER": "Лазерная коррекция зрения"}
GRADES = ["junior", "middle", "senior", "lead"]
NO_DOC = "—"

F = "Arial"
BLUE = Font(name=F, size=10, color="0000FF")
BLACK = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=14, bold=True)
GREY = Font(name=F, size=8, color="808080")
HINT = Font(name=F, size=9, italic=True, color="808080")
HDR_FONT = Font(name=F, size=10, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SEC_FILL = PatternFill("solid", fgColor="D9E2F3")
MAIN_FILL = PatternFill("solid", fgColor="E2EFDA")
INPUT_FILL = PatternFill("solid", fgColor="FFFF99")
TOP = Border(top=Side(style="thin"))
NUM = "#,##0,;(#,##0,);-"
PCT = "0.0%;(0.0%);-"
INT = "#,##0;(#,##0);-"


# ------------------------------------------------------------------ данные
def prepare():
    visits = pd.read_csv(CORR / "visits.csv")
    items = pd.read_csv(CORR / "visit_items.csv")
    services = pd.read_csv(CORR / "services.csv")
    mat1c = pd.read_csv(CORR / "materials_1c.csv")
    pay = pd.read_csv(CORR / "payroll_1c.csv")
    doctors = pd.read_csv(CORR / "doctors.csv")

    im = items.merge(visits[["visit_id", "visit_date", "clinic_id", "doctor_id", "direction_code"]], on="visit_id") \
              .merge(services[["service_id", "material_cost", "external_cost"]], on="service_id")
    im["period"] = im.visit_date.str[:7]
    im["n_mat"] = im.qty * im.material_cost
    im["n_ext"] = im.qty * im.external_cost
    # факт затрат 1С (по направлениям) -> по позициям пропорционально нормативу (как в build_factor)
    cn = im.groupby(CELL)[["n_mat", "n_ext"]].transform("sum")
    m = im[CELL].merge(mat1c[CELL + ["material_fact", "external_fact"]], on=CELL, how="left")
    im["mat"] = np.where(cn.n_mat > 0, m.material_fact * im.n_mat / cn.n_mat.replace(0, 1), 0.0)
    im["ext"] = np.where(cn.n_ext > 0, m.external_fact * im.n_ext / cn.n_ext.replace(0, 1), 0.0)
    im["doctor_id"] = im.doctor_id.fillna(NO_DOC)

    g = im.groupby(["doctor_id", "period"], as_index=False).agg(
        revenue=("revenue", "sum"), materials=("mat", "sum"), labs=("ext", "sum"))
    v = visits[visits.status == "completed"].copy()
    v["period"] = v.visit_date.str[:7]
    v["doctor_id"] = v.doctor_id.fillna(NO_DOC)
    vc = v.groupby(["doctor_id", "period"], as_index=False).size().rename(columns={"size": "visits"})
    p = pay[pay.role == "doctor"].rename(columns={"employee_id": "doctor_id", "gross": "pay_gross",
                                                  "contributions": "pay_contrib"})
    p = p[["doctor_id", "period", "pay_gross", "pay_contrib"]].assign(worked=1)

    d = g.merge(vc, on=["doctor_id", "period"], how="outer").merge(p, on=["doctor_id", "period"], how="outer")
    d = d.fillna({"revenue": 0, "materials": 0, "labs": 0, "visits": 0, "pay_gross": 0, "pay_contrib": 0, "worked": 0})
    attrs = doctors[["doctor_id", "clinic_id", "direction_code", "grade"]]
    d = d.merge(attrs, on="doctor_id", how="left")
    for c in ("clinic_id", "direction_code", "grade"):
        d[c] = d[c].fillna(NO_DOC)
    d = d[["period", "doctor_id", "clinic_id", "direction_code", "grade", "visits", "revenue", "materials", "labs",
           "pay_gross", "pay_contrib", "worked"]].sort_values(["doctor_id", "period"]).reset_index(drop=True)
    for c in ("revenue", "materials", "labs", "pay_gross", "pay_contrib"):
        d[c] = d[c].round(2)
    d[["visits", "worked"]] = d[["visits", "worked"]].astype(int)

    order = {k: i for i, k in enumerate(DIRS)}
    dl = doctors.assign(o=doctors.direction_code.map(order)).sort_values(["o", "clinic_id", "doctor_id"])
    dl = dl[["doctor_id", "doctor_name", "clinic_id", "direction_code", "grade", "pay_pct"]].reset_index(drop=True)

    ctrl = {"revenue": items.revenue.sum(), "materials": mat1c.material_fact.sum(),
            "labs": mat1c.external_fact.sum(), "gross": p.pay_gross.sum(), "contrib": p.pay_contrib.sum()}
    rates0 = doctors.groupby("grade").pay_pct.agg(["min", "max"])
    assert (rates0["min"] == rates0["max"]).all(), "ставка оплаты не единая внутри грейда"
    return d, dl, ctrl, rates0["min"].to_dict()


# ------------------------------------------------------------------ оформление
def head(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build(data, dl, ctrl, rates0):
    wb = Workbook()

    # ---------- Данные ----------
    wd = wb.active
    wd.title = "Данные"
    wd.append(list(data.columns))
    for r in data.itertuples(index=False):
        wd.append(list(r))
    nd = wd.max_row
    head(wd, 1, 1, 12)
    for row in wd.iter_rows(min_row=2, max_row=nd):
        for c in row:
            c.font = BLUE if c.column >= 6 else BLACK
            if c.column >= 7 and c.column <= 11:
                c.number_format = "#,##0.00"
    for i, w in enumerate([10, 10, 9, 14, 9, 8, 14, 14, 14, 14, 14, 8], 1):
        wd.column_dimensions[col(i)].width = w
    wd.freeze_panes = "A2"
    D = lambda c: f"'Данные'!${c}$2:${c}${nd}"          # noqa: E731

    periods = [f"{YEAR}-{m:02d}" for m in range(1, 13)]

    # ======================= Врачи =======================
    wv = wb.create_sheet("Врачи", 0)
    wv["A1"], wv["A1"].font = "Врачи: выручка, прямая маржа и вклад за 2025 год (тыс. ₽)", TITLE
    wv["A2"], wv["A2"].font = "Месяц (фильтр):", BOLD
    wv["B2"], wv["B2"].fill, wv["B2"].font = "Все", INPUT_FILL, BLUE
    wv["C2"] = '=IF($B$2="Все","*",$B$2)'
    wv["C2"].font = GREY
    wv["D2"] = "← «Все» или месяц вида 2025-05"
    wv["D2"].font = HINT
    dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + periods) + '"', allow_blank=False)
    wv.add_data_validation(dv)
    dv.add("B2")
    wv["A3"] = ("Главный показатель: вклад на врача-месяц (колонка T). Врач-месяц = месяц с записью в payroll. "
                "Затраты на материалы и лаборатории разнесены по нормативу услуг.")
    wv["A3"].font = Font(name=F, size=9, italic=True)

    heads = ["Врач", "ID", "Клиника", "Направление", "Грейд", "Ставка оплаты", "Месяцев работы", "Визитов",
             "Выручка", "Материалы", "Лаборатории", "Прямая маржа", "Прямая маржа, %", "Оплата врача",
             "Взносы", "Вклад врача", "Вклад, %", "Выручка на врача-месяц", "Прямая маржа на врача-месяц",
             "Вклад на врача-месяц", "Ранг по вкладу на месяц"]
    HR = 5
    for i, h in enumerate(heads, 1):
        wv.cell(HR, i, h)
    head(wv, HR, 1, len(heads))
    wv.row_dimensions[HR].height = 42
    r1 = HR + 1
    r2 = r1 + len(dl) - 1

    def doctor_row(r, key):
        # агрегаты по врачу и выбранному месяцу
        def s(c):
            return f"=SUMIFS({D(c)},{D('B')},$B{r},{D('A')},$C$2)"
        wv.cell(r, 7, s("L")); wv.cell(r, 8, s("F")); wv.cell(r, 9, s("G"))
        wv.cell(r, 10, s("H")); wv.cell(r, 11, s("I")); wv.cell(r, 14, s("J")); wv.cell(r, 15, s("K"))

    for k, row in enumerate(dl.itertuples(index=False)):
        r = r1 + k
        wv.cell(r, 1, row.doctor_name); wv.cell(r, 2, row.doctor_id); wv.cell(r, 3, row.clinic_id)
        wv.cell(r, 4, row.direction_code); wv.cell(r, 5, row.grade); wv.cell(r, 6, float(row.pay_pct))
        doctor_row(r, row.doctor_id)
    rn = r2 + 1                       # «Без врача»
    wv.cell(rn, 1, "Без врача"); wv.cell(rn, 2, NO_DOC); wv.cell(rn, 3, NO_DOC)
    wv.cell(rn, 4, NO_DOC); wv.cell(rn, 5, NO_DOC)
    doctor_row(rn, NO_DOC)
    rd, rt = rn + 1, rn + 2           # итого врачи, итого сеть
    wv.cell(rd, 1, "Врачи итого"); wv.cell(rt, 1, "Итого по сети")
    for c in (7, 8, 9, 10, 11, 14, 15):
        L = col(c)
        wv.cell(rd, c, f"=SUM({L}{r1}:{L}{r2})")
        wv.cell(rt, c, f"={L}{rd}+{L}{rn}")
    for r in list(range(r1, r2 + 1)) + [rn, rd, rt]:
        wv.cell(r, 12, f"=I{r}-J{r}-K{r}")
        wv.cell(r, 13, f"=IF(I{r}=0,0,L{r}/I{r})")
        wv.cell(r, 16, f"=L{r}-N{r}-O{r}")
        wv.cell(r, 17, f"=IF(I{r}=0,0,P{r}/I{r})")
        if r != rt and r != rn:
            wv.cell(r, 18, f"=IF(G{r}=0,0,I{r}/G{r})")
            wv.cell(r, 19, f"=IF(G{r}=0,0,L{r}/G{r})")
            wv.cell(r, 20, f"=IF(G{r}=0,0,P{r}/G{r})")
    for r in range(r1, r2 + 1):
        wv.cell(r, 21, f'=IF(G{r}=0,"—",COUNTIFS($T${r1}:$T${r2},">"&T{r},$G${r1}:$G${r2},">0")+1)')

    for r in range(r1, rt + 1):
        bold = r >= rd
        for c in range(1, 22):
            cell = wv.cell(r, c)
            cell.font = BOLD if bold else (BLUE if c in (6,) else BLACK)
            if bold:
                cell.border = TOP
            if c == 6:
                cell.number_format = "0%"
            elif c in (7, 8, 21):
                cell.number_format = INT
            elif c in (13, 17):
                cell.number_format = PCT
            elif c >= 9:
                cell.number_format = NUM
        wv.cell(r, 20).fill = MAIN_FILL
    wv.conditional_formatting.add(f"T{r1}:T{r2}", ColorScaleRule(
        start_type="min", start_color="F8CBAD", mid_type="percentile", mid_value=50, mid_color="FFFFFF",
        end_type="max", end_color="C6E0B4"))
    wv.cell(rt + 2, 1, "Итого по сети включает строку «Без врача» (визиты без указанного врача, оплаты нет). "
                      "Показатели на врача-месяц считаются только по врачам.").font = HINT
    for i, w in enumerate([18, 7, 8, 12, 9, 9, 9, 9, 12, 12, 12, 12, 10, 12, 11, 12, 9, 12, 12, 12, 9], 1):
        wv.column_dimensions[col(i)].width = w
    wv.freeze_panes = wv.cell(r1, 3)
    wv.auto_filter.ref = f"A{HR}:{col(len(heads))}{r2}"

    V = lambda c: f"'Врачи'!${c}${r1}:${c}${r2}"        # noqa: E731

    # ======================= Грейды =======================
    wg = wb.create_sheet("Грейды", 1)
    wg["A1"], wg["A1"].font = "Грейды врачей: вклад на врача-месяц по сети и внутри направлений (тыс. ₽)", TITLE
    wg["A2"], wg["A2"].font = "Месяц:", BOLD
    wg["B2"] = "=Врачи!$B$2"
    wg["B2"].font = BOLD
    wg["C2"] = "← меняется на листе «Врачи»"
    wg["C2"].font = HINT
    wg["A3"], wg["A3"].font = "Клиника (фильтр):", BOLD
    wg["B3"], wg["B3"].fill, wg["B3"].font = "Все", INPUT_FILL, BLUE
    wg["C3"] = '=IF($B$3="Все","*",$B$3)'
    wg["C3"].font = GREY
    wg["D3"] = "← «Все», C01, C02 или C03"
    wg["D3"].font = HINT
    dv2 = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    wg.add_data_validation(dv2)
    dv2.add("B3")
    wg["A4"] = ("Сравнивать грейды корректнее внутри направления: маржа услуг в направлениях разная. "
                "Пометка «мало врачей» там, где врачей меньше трёх.")
    wg["A4"].font = Font(name=F, size=9, italic=True)

    gh = ["Направление", "Грейд", "Врачей", "Врач-месяцев", "Выручка", "Прямая маржа", "Оплата и взносы", "Вклад",
          "Выручка на врача-месяц", "Прямая маржа на врача-месяц", "Вклад на врача-месяц", "Прямая маржа, %",
          "Оплата и взносы, % выручки", "Вклад, %", "Вклад на месяц к среднему по направлению", "Примечание",
          "код напр.", "критерий грейда"]
    NC = len(gh)

    def agg(c, r):
        return (f"SUMIFS({V(c)},{V('E')},$R{r},{V('D')},$Q{r},{V('C')},$C$3)")

    def grade_row(r, dname, dcode, gcrit, gname, total_row=None, bold=False):
        wg.cell(r, 1, dname); wg.cell(r, 2, gname); wg.cell(r, 17, dcode); wg.cell(r, 18, gcrit)
        wg.cell(r, 3, f'=COUNTIFS({V("E")},$R{r},{V("D")},$Q{r},{V("C")},$C$3,{V("G")},">0")')
        wg.cell(r, 4, "=" + agg("G", r))
        wg.cell(r, 5, "=" + agg("I", r))
        wg.cell(r, 6, "=" + agg("L", r))
        wg.cell(r, 7, f"={agg('N', r)}+{agg('O', r)}")
        wg.cell(r, 8, "=" + agg("P", r))
        wg.cell(r, 9, f"=IF(D{r}=0,0,E{r}/D{r})")
        wg.cell(r, 10, f"=IF(D{r}=0,0,F{r}/D{r})")
        wg.cell(r, 11, f"=IF(D{r}=0,0,H{r}/D{r})")
        wg.cell(r, 12, f"=IF(E{r}=0,0,F{r}/E{r})")
        wg.cell(r, 13, f"=IF(E{r}=0,0,G{r}/E{r})")
        wg.cell(r, 14, f"=IF(E{r}=0,0,H{r}/E{r})")
        if total_row:
            wg.cell(r, 15, f'=IF(OR(D{r}=0,K{total_row}=0),"—",K{r}/K{total_row}-1)')
        wg.cell(r, 16, f'=IF(C{r}=0,"нет врачей",IF(C{r}<3,"мало врачей (n<3): вывод ненадёжен","—"))')
        for c in range(1, NC + 1):
            cell = wg.cell(r, c)
            cell.font = BOLD if bold else BLACK
            if bold:
                cell.border = TOP
            if c in (3, 4):
                cell.number_format = INT
            elif c in (12, 13, 14):
                cell.number_format = PCT
            elif c == 15:
                cell.number_format = "+0.0%;-0.0%;0.0%"
                cell.alignment = Alignment(horizontal="right")
            elif 5 <= c <= 11:
                cell.number_format = NUM
            if c >= 17:
                cell.font = GREY
        wg.cell(r, 11).fill = MAIN_FILL
        wg.cell(r, 16).font = Font(name=F, size=9, italic=True, color="C00000")

    # блок A: по сети
    wg.cell(6, 1, "По сети (все направления)").font = BOLD
    for c in range(1, NC + 1):
        wg.cell(6, c).fill = SEC_FILL
    for i, h in enumerate(gh, 1):
        wg.cell(7, i, h)
    head(wg, 7, 1, NC)
    wg.row_dimensions[7].height = 54
    ra1 = 8
    for k, g in enumerate(GRADES):
        grade_row(ra1 + k, "Сеть", "*", g, g, total_row=ra1 + len(GRADES))
    grade_row(ra1 + len(GRADES), "Сеть", "*", "*", "Все грейды", bold=True)
    ra_t = ra1 + len(GRADES)
    wg.conditional_formatting.add(f"K{ra1}:K{ra_t - 1}", ColorScaleRule(
        start_type="min", start_color="F8CBAD", end_type="max", end_color="C6E0B4"))

    # блок B: внутри направлений
    rb0 = ra_t + 3
    wg.cell(rb0 - 1, 1, "Внутри направлений").font = BOLD
    for c in range(1, NC + 1):
        wg.cell(rb0 - 1, c).fill = SEC_FILL
    for i, h in enumerate(gh, 1):
        wg.cell(rb0, i, h)
    head(wg, rb0, 1, NC)
    wg.row_dimensions[rb0].height = 54
    r = rb0 + 1
    for code, name in DIRS.items():
        tot = r + len(GRADES)
        for g in GRADES:
            grade_row(r, name, code, g, g, total_row=tot)
            r += 1
        grade_row(r, name, code, "*", "Все грейды", bold=True)
        r += 1
    wg.conditional_formatting.add(f"O{rb0 + 1}:O{r - 1}", ColorScaleRule(
        start_type="num", start_value=-0.5, start_color="F8CBAD", mid_type="num", mid_value=0, mid_color="FFFFFF",
        end_type="num", end_value=0.5, end_color="C6E0B4"))
    for i, w in enumerate([27, 11, 8, 10, 12, 12, 12, 12, 12, 12, 12, 10, 12, 9, 14, 34, 8, 9], 1):
        wg.column_dimensions[col(i)].width = w
    wg.freeze_panes = "C8"

    # ======================= Расчёт (построчно, врач x месяц) =======================
    wc = wb.create_sheet("Расчёт")
    ch = ["ID врача", "Грейд", "Месяц работы", "Выручка факт", "Материалы и лаборатории факт", "Оплата факт",
          "Взносы факт", "Выше порога", "Выручка сценарий", "Материалы и лаборатории сценарий", "Ставка",
          "Оплата сценарий", "Взносы сценарий", "Вклад факт", "Вклад сценарий", "Оплата по исходной схеме",
          "Взносы по исходной схеме"]
    for i, h in enumerate(ch, 1):
        wc.cell(1, i, h)
    head(wc, 1, 1, len(ch))
    wc.row_dimensions[1].height = 42
    K = "'Калькулятор'!"
    TH = f"{K}$B$9*1000"
    for r in range(2, nd + 1):
        wc.cell(r, 1, f"='Данные'!B{r}")
        wc.cell(r, 2, f"='Данные'!E{r}")
        wc.cell(r, 3, f"='Данные'!L{r}")
        wc.cell(r, 4, f"='Данные'!G{r}")
        wc.cell(r, 5, f"='Данные'!H{r}+'Данные'!I{r}")
        wc.cell(r, 6, f"='Данные'!J{r}")
        wc.cell(r, 7, f"='Данные'!K{r}")
        wc.cell(r, 8, f"=IF(AND(C{r}=1,D{r}>{TH}),1,0)")
        wc.cell(r, 9, f"=D{r}*(1+H{r}*{K}$B$12)")
        wc.cell(r, 10, f"=IF(D{r}=0,0,E{r}*I{r}/D{r})")
        wc.cell(r, 11, f"=IFERROR(INDEX({K}$B$5:$B$8,MATCH(B{r},{K}$A$5:$A$8,0)),0)")
        wc.cell(r, 12, f"=IF(C{r}=1,ROUND(K{r}*I{r}+{K}$B$10*MAX(0,I{r}-{TH})+{K}$B$11*1000,0),0)")
        wc.cell(r, 13, f"=ROUND(L{r}*{K}$B$13,0)")
        wc.cell(r, 14, f"=D{r}-E{r}-F{r}-G{r}")
        wc.cell(r, 15, f"=I{r}-J{r}-L{r}-M{r}")
        wc.cell(r, 16, f"=IF(C{r}=1,ROUND(IFERROR(INDEX({K}$C$5:$C$8,MATCH(B{r},{K}$A$5:$A$8,0)),0)*D{r},0),0)")
        wc.cell(r, 17, f"=ROUND(P{r}*{K}$C$13,0)")
        for c in range(1, 18):
            cell = wc.cell(r, c)
            cell.font = BLACK
            if c == 11:
                cell.number_format = "0%"
            elif c in (3, 8):
                cell.number_format = "0"
            elif c >= 4:
                cell.number_format = "#,##0"
    for i in range(1, 18):
        wc.column_dimensions[col(i)].width = 13
    wc.freeze_panes = "A2"
    R = lambda c: f"'Расчёт'!${c}$2:${c}${nd}"          # noqa: E731

    # ======================= Калькулятор =======================
    wk = wb.create_sheet("Калькулятор")
    wk["A1"], wk["A1"].font = "Калькулятор схемы оплаты врачей: что если (2025, вся сеть, тыс. ₽)", TITLE
    wk["A2"] = ("Жёлтые ячейки меняются. Расчёт идёт на фактических данных 2025: как выглядели бы ФОТ и вклад, "
                "если бы врачам платили по новой схеме.")
    wk["A2"].font = Font(name=F, size=9, italic=True)
    for i, h in enumerate(["Параметр", "Значение", "Исходная схема", "Пояснение"], 1):
        wk.cell(4, i, h)
    head(wk, 4, 1, 4)
    params = [
        (5, "junior", rates0["junior"], rates0["junior"], "0%", "Ставка оплаты, доля выручки врача"),
        (6, "middle", rates0["middle"], rates0["middle"], "0%", "Ставка оплаты, доля выручки врача"),
        (7, "senior", rates0["senior"], rates0["senior"], "0%", "Ставка оплаты, доля выручки врача"),
        (8, "lead", rates0["lead"], rates0["lead"], "0%", "Ставка оплаты, доля выручки врача"),
        (9, "Порог месячной выручки врача, тыс. ₽", 3000, "—", "#,##0",
         "Надбавка и сценарий роста действуют на врач-месяцы выше порога (медиана врач-месяца около 1 500 тыс. ₽)"),
        (10, "Надбавка к ставке выше порога, п.п.", 0.0, 0.0, "0.0%",
         "Только на выручку сверх порога (ступенчатая шкала): оплата += надбавка × (выручка − порог)"),
        (11, "Фиксированная часть, тыс. ₽ на врача-месяц", 0, 0, "#,##0.0", "Платится за каждый врач-месяц с записью в payroll"),
        (12, "Изменение выручки врач-месяцев выше порога, %", 0.0, 0.0, "0.0%",
         "Сценарное допущение, не прогноз: выручка всего врач-месяца меняется на этот процент, материалы и лаборатории пропорционально"),
        (13, "Ставка страховых взносов", 0.22, 0.22, "0.0%", "Из payroll; менять не нужно"),
        (14, "Минимальный вклад на врача-месяц, тыс. ₽", 500, "—", "#,##0",
         "Врачи с вкладом ниже минимума попадают в список внимания"),
    ]
    for r, name, val, base, fmt, note in params:
        wk.cell(r, 1, name).font = BOLD if r <= 8 else BLACK
        c = wk.cell(r, 2, val)
        c.number_format = fmt
        c.font = BLACK if r == 13 else BLUE
        if r != 13:
            c.fill = INPUT_FILL
        b = wk.cell(r, 3, base)
        b.number_format = fmt
        b.font = GREY
        wk.cell(r, 4, note).font = HINT
    wk["A15"] = ('=IF(AND(B5=C5,B6=C6,B7=C7,B8=C8,B10=C10,B11=C11,B12=C12,B13=C13),'
                 '"Параметры совпадают с исходной схемой: расчёт равен факту",'
                 '"Схема изменена: сравнение с фактом в таблицах ниже")')
    wk["A15"].font = Font(name=F, size=10, bold=True, color="C00000")

    # результат по грейдам
    wk["A17"], wk["A17"].font = "Результат по грейдам за 2025 год (только врачи, без строки «Без врача»)", BOLD
    for c in range(1, 15):
        wk.cell(17, c).fill = SEC_FILL
    kh = ["Грейд", "Врачей", "Врач-месяцев", "Выручка факт", "Выручка сценарий", "Оплата и взносы факт",
          "Оплата и взносы сценарий", "Изменение оплаты и взносов", "Вклад факт", "Вклад сценарий",
          "Изменение вклада", "Вклад на врача-месяц факт", "Вклад на врача-месяц сценарий",
          "Вклад на врача-месяц, изменение %"]
    for i, h in enumerate(kh, 1):
        wk.cell(18, i, h)
    head(wk, 18, 1, 14)
    wk.row_dimensions[18].height = 54
    dr1, dr2 = 29, 29 + len(dl) - 1            # таблица врачей ниже
    for k, g in enumerate(GRADES):
        r = 19 + k
        wk.cell(r, 1, g)
        wk.cell(r, 2, f'=COUNTIFS($D${dr1}:$D${dr2},$A{r},$E${dr1}:$E${dr2},">0")')
        wk.cell(r, 3, f"=SUMIFS({R('C')},{R('B')},$A{r})")
        wk.cell(r, 4, f"=SUMIFS({R('D')},{R('B')},$A{r})")
        wk.cell(r, 5, f"=SUMIFS({R('I')},{R('B')},$A{r})")
        wk.cell(r, 6, f"=SUMIFS({R('F')},{R('B')},$A{r})+SUMIFS({R('G')},{R('B')},$A{r})")
        wk.cell(r, 7, f"=SUMIFS({R('L')},{R('B')},$A{r})+SUMIFS({R('M')},{R('B')},$A{r})")
        wk.cell(r, 9, f"=SUMIFS({R('N')},{R('B')},$A{r})")
        wk.cell(r, 10, f"=SUMIFS({R('O')},{R('B')},$A{r})")
    rtot = 19 + len(GRADES)
    wk.cell(rtot, 1, "Все врачи")
    for c in (2, 3, 4, 5, 6, 7, 9, 10):
        L = col(c)
        wk.cell(rtot, c, f"=SUM({L}19:{L}{rtot - 1})")
    for r in range(19, rtot + 1):
        wk.cell(r, 8, f"=G{r}-F{r}")
        wk.cell(r, 11, f"=J{r}-I{r}")
        wk.cell(r, 12, f"=IF(C{r}=0,0,I{r}/C{r})")
        wk.cell(r, 13, f"=IF(C{r}=0,0,J{r}/C{r})")
        wk.cell(r, 14, f"=IF(L{r}=0,0,M{r}/L{r}-1)")
        for c in range(1, 15):
            cell = wk.cell(r, c)
            cell.font = BOLD if r == rtot else BLACK
            if r == rtot:
                cell.border = TOP
            if c in (2, 3):
                cell.number_format = INT
            elif c == 14:
                cell.number_format = "+0.0%;-0.0%;0.0%"
            elif c >= 4:
                cell.number_format = NUM
        wk.cell(r, 13).fill = MAIN_FILL
    wk.cell(rtot + 2, 1, "Врачей ниже минимального вклада:").font = BOLD
    wk.cell(rtot + 2, 3, f'=COUNTIF($I${dr1}:$I${dr2},"ниже минимума")').font = BOLD
    wk.cell(rtot + 2, 4, "по сценарию (по факту):").font = HINT
    wk.cell(rtot + 2, 6, f'=COUNTIFS($F${dr1}:$F${dr2},"<"&$B$14*1000,$E${dr1}:$E${dr2},">0")').font = BOLD

    # врачи в сценарии
    wk.cell(dr1 - 2, 1, "Врачи: вклад на врача-месяц по факту и по сценарию").font = BOLD
    for c in range(1, 10):
        wk.cell(dr1 - 2, c).fill = SEC_FILL
    for i, h in enumerate(["Врач", "ID", "Направление", "Грейд", "Месяцев работы", "Вклад на месяц факт",
                           "Вклад на месяц сценарий", "Изменение", "Статус"], 1):
        wk.cell(dr1 - 1, i, h)
    head(wk, dr1 - 1, 1, 9)
    wk.row_dimensions[dr1 - 1].height = 42
    for k, row in enumerate(dl.itertuples(index=False)):
        r = dr1 + k
        wk.cell(r, 1, row.doctor_name)
        wk.cell(r, 2, row.doctor_id)
        wk.cell(r, 3, row.direction_code)
        wk.cell(r, 4, row.grade)
        wk.cell(r, 5, f"=SUMIFS({R('C')},{R('A')},$B{r})")
        wk.cell(r, 6, f"=IF(E{r}=0,0,SUMIFS({R('N')},{R('A')},$B{r})/E{r})")
        wk.cell(r, 7, f"=IF(E{r}=0,0,SUMIFS({R('O')},{R('A')},$B{r})/E{r})")
        wk.cell(r, 8, f"=G{r}-F{r}")
        wk.cell(r, 9, f'=IF(E{r}=0,"—",IF(G{r}<$B$14*1000,"ниже минимума","ок"))')
        for c in range(1, 10):
            cell = wk.cell(r, c)
            cell.font = BLACK
            if c == 5:
                cell.number_format = INT
            elif c in (6, 7, 8):
                cell.number_format = NUM
    wk.conditional_formatting.add(f"I{dr1}:I{dr2}", FormulaRule(
        formula=[f'$I{dr1}="ниже минимума"'], font=Font(name=F, size=10, bold=True, color="C00000"),
        fill=PatternFill("solid", bgColor="F8CBAD")))
    for i, w in enumerate([44, 12, 14, 12, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14], 1):
        wk.column_dimensions[col(i)].width = w
    wk.freeze_panes = "A5"

    # ======================= Описание =======================
    wo = wb.create_sheet("Описание", 0)
    wo["A1"], wo["A1"].font = "Аналитика врачей, 2025: описание и контроль", TITLE
    lines = [
        ("Что считаем", None),
        ("Выручка", "Сумма услуг завершённых визитов врача (visit_items). Визиты без врача выделены в строку «Без врача»."),
        ("Материалы, лаборатории", "Факт 1С по клинике, направлению и месяцу, разнесённый на позиции пропорционально нормативу "
                                   "(количество × норма услуги), как в факторном анализе. Сумма по врачам равна 1С до рубля."),
        ("Прямая маржа", "Выручка − материалы − лаборатории."),
        ("Оплата и взносы", "payroll_1c: начислено врачу (gross, до НДФЛ) и страховые взносы работодателя."),
        ("Вклад врача", "Прямая маржа − оплата врача − взносы. Это то, что врач оставляет клинике до распределения "
                        "медперсонала, амортизации, управленческих и коммерческих расходов (по врачам не распределяются)."),
        ("Врач-месяц", "Месяц, по которому у врача есть запись в payroll. Главный показатель: вклад на врача-месяц; "
                       "годовые суммы между врачами несопоставимы, потому что у части врачей в payroll 4-8 месяцев."),
        ("Как читать", None),
        ("Процент маржи", "Вклад в % выручки почти полностью определяется ставкой оплаты (у младших 28%, у старших до 40%), "
                          "поэтому по нему младшие грейды всегда «лучше». Для сравнения используется вклад в рублях на врача-месяц."),
        ("Грейды", "Сравнивать корректнее внутри направления: маржа услуг в направлениях разная (около 92% у лазера и семейной "
                   "медицины, около 80% у стоматологии). Колонка «к среднему по направлению» убирает эффект состава."),
        ("Малые группы", "Если в группе меньше трёх врачей, на листе «Грейды» стоит пометка: вывод по такой группе ненадёжен."),
        ("Фильтры", "Месяц задаётся на листе «Врачи» и действует на оба листа. Клиника задаётся на листе «Грейды»."),
        ("Калькулятор", None),
        ("Что считает", "Схема оплаты «что если» на фактических данных 2025 по врач-месяцам: ставки по грейдам, надбавка к ставке на "
                        "выручку сверх порога, фиксированная часть, сценарный рост выручки. Показывает ФОТ, взносы, вклад по грейдам и "
                        "врачам и список врачей ниже заданного минимума вклада."),
        ("Проверка", "При исходных параметрах расчёт равен оплате и взносам из payroll до рубля (контроль ниже). Логика зеркальна "
                     "функции src/common/payscheme.py; формулы проверены на изменённых параметрах против pandas."),
        ("Ограничения", None),
        ("Визиты без врача", "В ERP было два визита без врача (55,0 тыс. ₽, март, C03 и 65,7 тыс. ₽, август, C02, обе стоматология). Сверка "
                             "привязала их к врачам D008 и D007: payroll начислил этим врачам выручку выше учтённой в ERP ровно на эти суммы. "
                             "Исправление средней уверенности, помечено «подтвердить бухгалтером»."),
        ("Калькулятор не прогноз", "Выручка врачей остаётся фактической, кроме явно заданного сценарного роста. Реакция врачей на схему, "
                                   "влияние на загрузку кабинетов, цены и других врачей не моделируется; сценарный рост это допущение "
                                   "пользователя, а не предсказание."),
        ("Затраты по нормативу", "Расход материалов и лабораторий врача считается по норме услуги: перерасход конкретного врача "
                                 "не виден, различия между врачами возникают только из-за набора услуг."),
        ("Данные синтетические", "Это модель клиники. Цифры демонстрируют метод и его внутреннюю согласованность, "
                                 "а не реальную эффективность врачей."),
    ]
    r = 3
    for k, v in lines:
        if v is None:
            wo.cell(r, 1, k).font = BOLD
            wo.cell(r, 1).fill = SEC_FILL
            wo.cell(r, 2).fill = SEC_FILL
        else:
            wo.cell(r, 1, k).font = BOLD
            wo.cell(r, 2, v).font = BLACK
            wo.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
            wo.cell(r, 1).alignment = Alignment(vertical="top")
        r += 1
    r += 1
    wo.cell(r, 1, "Контроль (при фильтре месяца «Все»), тыс. ₽").font = BOLD
    for c in (1, 2, 3, 4):
        wo.cell(r, c).fill = SEC_FILL
    r += 1
    for i, h in enumerate(["Показатель", "По книге", "Источник", "Разница"], 1):
        wo.cell(r, i, h)
    head(wo, r, 1, 4)
    ctrl_rows = [("Выручка (visit_items)", f"='Врачи'!I{rt}", ctrl["revenue"]),
                 ("Материалы (materials_1c)", f"='Врачи'!J{rt}", ctrl["materials"]),
                 ("Лаборатории (materials_1c)", f"='Врачи'!K{rt}", ctrl["labs"]),
                 ("Оплата врачей (payroll_1c)", f"='Врачи'!N{rt}", ctrl["gross"]),
                 ("Взносы с врачей (payroll_1c)", f"='Врачи'!O{rt}", ctrl["contrib"]),
                 ("Оплата по исходной схеме, расчёт калькулятора", f"=SUM('Расчёт'!$P$2:$P${nd})", ctrl["gross"]),
                 ("Взносы по исходной схеме, расчёт калькулятора", f"=SUM('Расчёт'!$Q$2:$Q${nd})", ctrl["contrib"])]
    for name, f, src in ctrl_rows:
        r += 1
        wo.cell(r, 1, name).font = BLACK
        wo.cell(r, 2, f).number_format = "#,##0.00,;(#,##0.00,);-"
        if isinstance(src, str):
            wo.cell(r, 3, src).number_format = "#,##0.00;(#,##0.00);-"
        else:
            wo.cell(r, 3, round(float(src), 2)).font = BLUE
            wo.cell(r, 3).number_format = "#,##0.00,;(#,##0.00,);-"
        wo.cell(r, 4, f"=B{r}-C{r}").number_format = "#,##0.00;(#,##0.00);-"
    wo.column_dimensions["A"].width = 30
    wo.column_dimensions["B"].width = 110
    wo.column_dimensions["C"].width = 16
    wo.column_dimensions["D"].width = 14

    wb._sheets = [wb["Описание"], wb["Врачи"], wb["Грейды"], wb["Калькулятор"], wb["Расчёт"], wb["Данные"]]
    wb.active = 1
    return wb, (r1, r2, rn, rd, rt)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    data, dl, ctrl, rates0 = prepare()
    wb, rows = build(data, dl, ctrl, rates0)
    path = OUT / "Doctors_2025.xlsx"
    wb.save(path)
    recalculate(path)
    print(f"Сохранено: {path}  (строк данных: {len(data)}, врачей: {len(dl)})")

    doc = data[data.doctor_id != NO_DOC]
    t = doc.groupby("doctor_id").agg(rev=("revenue", "sum"), mat=("materials", "sum"), lab=("labs", "sum"),
                                     pay=("pay_gross", "sum"), ctr=("pay_contrib", "sum"), mon=("worked", "sum"))
    t["direct"] = t.rev - t.mat - t.lab
    t["contr"] = t.direct - t.pay - t.ctr
    grade = dl.set_index("doctor_id").grade
    gr = t.join(grade).groupby("grade").agg(n=("rev", "size"), mon=("mon", "sum"), contr=("contr", "sum"))
    gr["вклад на врача-месяц, тыс. ₽"] = (gr.contr / gr.mon / 1e3).round(0)
    print("Контроль в pandas: выручка", f"{data.revenue.sum():,.0f}", "| вклад врачей", f"{t.contr.sum():,.0f}")
    print(gr.loc[GRADES])


if __name__ == "__main__":
    main()
