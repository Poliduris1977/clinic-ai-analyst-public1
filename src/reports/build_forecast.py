"""
Шаг 14: прогноз выручки по направлениям и врачам на январь-март 2026 в Excel.

Вход:  data/corrected/ (visits, visit_items, doctors, payroll_1c через build_doctors.prepare)
Выход: data/reports/Forecast_2026Q1.xlsx

Позиция: рост 2025 года считается ростом бизнеса, сезонность по одному году оценить нельзя, поэтому базовый прогноз
это тренд с затуханием (метод Хольта) без сезонности; сезонная поправка на январь-март задаётся вручную (1,00).
Прогноз врача = прогноз направления x доля врача (среднемесячная выручка за последние три месяца), сумма врачей равна направлению.
Погрешность и надёжность по направлениям берутся из отката по 2025 году (таблица «Откат»).

Листы:
  Описание            - метод, допущения, ограничения, контроль
  Прогноз             - направления: факт IV квартала, прогноз по месяцам и за квартал, коридор, надёжность
  Врачи               - прогноз по врачам, доля в направлении, пометки по стажу
  Откат               - проверка методов на 2025 годе, погрешность по направлениям, детальная таблица
  Расчёт Хольта       - формулы метода по месяцам (источник для листа «Прогноз»)
  Данные направлений, Данные врачей - источники

Запуск:  python src/reports/build_forecast.py
"""
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from common.recalc import recalculate  # noqa: E402
from common import forecast as fc  # noqa: E402
import build_doctors  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
DIRS = {"DENT": "Стоматология", "FAM": "Семейная медицина", "COSM": "Косметология",
        "PLAS": "Пластическая хирургия", "LASER": "Лазерная коррекция зрения"}
CODES = list(DIRS)
MONTHS26 = ["Янв 2026", "Фев 2026", "Мар 2026"]

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
INPUT_FILL = PatternFill("solid", fgColor="FFFF99")
MAIN_FILL = PatternFill("solid", fgColor="E2EFDA")
TOP = Border(top=Side(style="thin"))
NUM = "#,##0.0,,;(#,##0.0,,);-"        # ₽ -> млн ₽
NUMK = "#,##0,;(#,##0,);-"             # ₽ -> тыс. ₽
PCT = "0.0%;(0.0%);-"
INT = "#,##0;(#,##0);-"


# ------------------------------------------------------------------ данные
def prepare():
    v = pd.read_csv(CORR / "visits.csv")
    it = pd.read_csv(CORR / "visit_items.csv")
    r = it.merge(v[["visit_id", "visit_date", "direction_code"]], on="visit_id")
    r["m"] = r.visit_date.str[5:7].astype(int)
    S = r.pivot_table(index="m", columns="direction_code", values="revenue", aggfunc="sum")[CODES]

    dd, dl, _, _ = build_doctors.prepare()
    dd = dd[dd.doctor_id != build_doctors.NO_DOC].copy()
    dd["m"] = dd.period.str[5:7].astype(int)
    rev = dd.pivot_table(index="doctor_id", columns="m", values="revenue", aggfunc="sum", fill_value=0.0) \
            .reindex(columns=range(1, 13), fill_value=0.0)
    wk = dd.pivot_table(index="doctor_id", columns="m", values="worked", aggfunc="sum", fill_value=0) \
           .reindex(columns=range(1, 13), fill_value=0)
    dl = dl.set_index("doctor_id")
    rev = rev.reindex(dl.index).fillna(0.0)
    wk = wk.reindex(dl.index).fillna(0).astype(int)

    methods = {}
    for name, m in fc.METHODS.items():
        bt = fc.backtest(S, m)
        methods[name] = fc.summary(bt)
    bt = fc.backtest(S, fc.holt)
    by_series = {}
    for s, g in bt.groupby("series"):
        by_series[s] = {"mape": float(g.err.abs().mean()), "bias": float(g.err.mean()), "n": int(len(g))}
    pyf = {c: fc.holt(S[c].values) for c in CODES}
    ctrl = {"rev_total": float(S.values.sum()), "q1_total": float(sum(x.sum() for x in pyf.values()))}
    return S, rev, wk, dl, bt, methods, by_series, ctrl


# ------------------------------------------------------------------ оформление
def head(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build(S, rev, wk, dl, bt, methods, by_series, ctrl):
    wb = Workbook()

    # ---------- Данные направлений ----------
    wn = wb.active
    wn.title = "Данные направлений"
    wn.cell(1, 1, "Месяц 2025")
    for j, c in enumerate(CODES):
        wn.cell(1, 2 + j, c)
    head(wn, 1, 1, 6)
    for m in range(1, 13):
        wn.cell(1 + m, 1, m).font = BLACK
        for j, c in enumerate(CODES):
            cell = wn.cell(1 + m, 2 + j, round(float(S.loc[m, c]), 2))
            cell.font = BLUE
            cell.number_format = "#,##0.00"
    for i in range(1, 7):
        wn.column_dimensions[col(i)].width = 16
    DN = lambda j, r1=2, r2=13: f"'Данные направлений'!${col(2 + j)}${r1}:${col(2 + j)}${r2}"      # noqa: E731

    # ---------- Данные врачей ----------
    wr = wb.create_sheet("Данные врачей")
    hd = ["ID", "Врач", "Направление", "Грейд"] + [f"Выручка {m}" for m in range(1, 13)] + \
         [f"Месяц работы {m}" for m in range(1, 13)]
    for i, h in enumerate(hd, 1):
        wr.cell(1, i, h)
    head(wr, 1, 1, len(hd))
    names = pd.read_csv(CORR / "doctors.csv").set_index("doctor_id").doctor_name
    for k, (did, row) in enumerate(dl.iterrows()):
        r = 2 + k
        wr.cell(r, 1, did)
        wr.cell(r, 2, names[did])
        wr.cell(r, 3, row.direction_code)
        wr.cell(r, 4, row.grade)
        for m in range(1, 13):
            c1 = wr.cell(r, 4 + m, round(float(rev.loc[did, m]), 2))
            c1.font = BLUE
            c1.number_format = "#,##0.00"
            c2 = wr.cell(r, 16 + m, int(wk.loc[did, m]))
            c2.font = BLUE
    nr = len(dl) + 1
    for i in range(1, len(hd) + 1):
        wr.column_dimensions[col(i)].width = 14 if i > 4 else 12
    wr.freeze_panes = "E2"

    # ======================= Откат =======================
    wo = wb.create_sheet("Откат", 0)
    wo["A1"], wo["A1"].font = "Откат: проверка методов на 2025 годе (обучение по первым 6-9 месяцам, прогноз трёх следующих)", TITLE
    wo["A2"] = ("Погрешность = факт / прогноз − 1. Положительная значит, что прогноз занизил факт. Итог по сети это сумма прогнозов "
                "направлений. Всего 72 наблюдения (6 рядов × 4 момента × 3 горизонта), они зависимы: оценка ориентировочная.")
    wo["A2"].font = Font(name=F, size=9, italic=True)
    for i, h in enumerate(["Метод", "Средняя абсолютная ошибка", "Среднее смещение", "10-й перцентиль", "90-й перцентиль"], 1):
        wo.cell(4, i, h)
    head(wo, 4, 1, 5)
    wo.row_dimensions[4].height = 30
    for k, (name, s) in enumerate(methods.items()):
        r = 5 + k
        for c, v in enumerate([name, s["mape"], s["bias"], s["q10"], s["q90"]], 1):
            cell = wo.cell(r, c, v)
            cell.font = BOLD if name == "Хольт с затуханием" else BLACK
            if c > 1:
                cell.number_format = PCT
        if name == "Хольт с затуханием":
            for c in range(1, 6):
                wo.cell(r, c).fill = MAIN_FILL
    r0 = 5 + len(methods) + 2
    wo.cell(r0 - 1, 1, "Хольт с затуханием: погрешность по рядам").font = BOLD
    for c in range(1, 6):
        wo.cell(r0 - 1, c).fill = SEC_FILL
    for i, h in enumerate(["Ряд (код)", "Название", "Средняя абсолютная ошибка", "Среднее смещение", "Наблюдений"], 1):
        wo.cell(r0, i, h)
    head(wo, r0, 1, 5)
    wo.row_dimensions[r0].height = 30
    series = CODES + ["Итого"]
    for k, s in enumerate(series):
        r = r0 + 1 + k
        wo.cell(r, 1, s).font = BLACK
        wo.cell(r, 2, DIRS.get(s, "Итого (сумма направлений)")).font = BLACK
        for c, v in ((3, by_series[s]["mape"]), (4, by_series[s]["bias"]), (5, by_series[s]["n"])):
            cell = wo.cell(r, c, v)
            cell.font = BLACK
            cell.number_format = INT if c == 5 else PCT
    BT_S = (f"'Откат'!$A${r0 + 1}:$A${r0 + len(series)}", f"'Откат'!$C${r0 + 1}:$C${r0 + len(series)}")
    rd = r0 + len(series) + 3
    wo.cell(rd - 1, 1, "Детальная таблица отката (Хольт с затуханием)").font = BOLD
    for c in range(1, 7):
        wo.cell(rd - 1, c).fill = SEC_FILL
    for i, h in enumerate(["Ряд", "Обучение по месяцам 1..", "Горизонт, мес.", "Прогноз", "Факт", "Погрешность"], 1):
        wo.cell(rd, i, h)
    head(wo, rd, 1, 6)
    wo.row_dimensions[rd].height = 30
    for k, row in enumerate(bt.sort_values(["series", "origin", "horizon"]).itertuples(index=False)):
        r = rd + 1 + k
        for c, v in enumerate([row.series, row.origin, row.horizon, row.forecast, row.actual, row.err], 1):
            cell = wo.cell(r, c, float(v) if isinstance(v, (np.floating, float)) else (int(v) if isinstance(v, (np.integer, int)) else v))
            cell.font = BLACK
            if c in (4, 5):
                cell.number_format = NUMK
            elif c == 6:
                cell.number_format = PCT
    for i, w in enumerate([34, 26, 24, 18, 18, 14], 1):
        wo.column_dimensions[col(i)].width = w

    # ======================= Расчёт Хольта =======================
    wh = wb.create_sheet("Расчёт Хольта")
    wh["A1"] = "Месяц 2025"
    for j, c in enumerate(CODES):
        wh.cell(1, 2 + 3 * j, f"{c}: выручка")
        wh.cell(1, 3 + 3 * j, f"{c}: уровень")
        wh.cell(1, 4 + 3 * j, f"{c}: тренд")
    head(wh, 1, 1, 1 + 3 * len(CODES))
    wh.row_dimensions[1].height = 30
    P = "'Прогноз'!"
    for m in range(1, 13):
        r = 1 + m
        wh.cell(r, 1, m).font = BLACK
        for j in range(len(CODES)):
            y, l, b = col(2 + 3 * j), col(3 + 3 * j), col(4 + 3 * j)
            wh.cell(r, 2 + 3 * j, f"='Данные направлений'!{col(2 + j)}{r}")
            if m == 1:
                wh.cell(r, 3 + 3 * j, f"={y}{r}")
                wh.cell(r, 4 + 3 * j, f"={y}{r + 1}-{y}{r}")
            else:
                wh.cell(r, 3 + 3 * j, f"={P}$B$4*{y}{r}+(1-{P}$B$4)*({l}{r - 1}+{P}$B$6*{b}{r - 1})")
                wh.cell(r, 4 + 3 * j, f"={P}$B$5*({l}{r}-{l}{r - 1})+(1-{P}$B$5)*{P}$B$6*{b}{r - 1}")
            for c in range(2 + 3 * j, 5 + 3 * j):
                wh.cell(r, c).number_format = "#,##0"
                wh.cell(r, c).font = BLACK
    for i in range(1, 2 + 3 * len(CODES)):
        wh.column_dimensions[col(i)].width = 14

    # ======================= Прогноз =======================
    wp = wb.create_sheet("Прогноз", 0)
    wp["A1"], wp["A1"].font = "Прогноз выручки по направлениям на январь-март 2026 (млн ₽)", TITLE
    wp["A3"], wp["A3"].font = "Параметры метода и сезонная поправка", BOLD
    for c in range(1, 5):
        wp.cell(3, c).fill = SEC_FILL
    params = [(4, "α, сглаживание уровня", fc.ALPHA, "0.00"), (5, "β, сглаживание тренда", fc.BETA, "0.00"),
              (6, "φ, затухание тренда", fc.PHI, "0.00"),
              (7, "Сезонная поправка, январь 2026", 1.0, "0.00"), (8, "Сезонная поправка, февраль 2026", 1.0, "0.00"),
              (9, "Сезонная поправка, март 2026", 1.0, "0.00")]
    for r, name, val, fmt in params:
        wp.cell(r, 1, name).font = BLACK
        c = wp.cell(r, 2, val)
        c.number_format = fmt
        c.fill, c.font = INPUT_FILL, BLUE
        d = wp.cell(r, 3, val)
        d.number_format = fmt
        d.font = GREY
    wp["D4"] = "← исходные значения справа; параметры не подбирались по данным"
    wp["D7"] = "← множитель к прогнозу месяца: 1,00 без поправки, 0,85 это −15%"
    wp["D4"].font = wp["D7"].font = HINT
    wp["A10"] = ('=IF(AND(B4=C4,B5=C5,B6=C6),"Параметры метода исходные: таблица «Откат» актуальна",'
                 '"Параметры метода изменены: «Откат» посчитан при исходных, перезапустите скрипт")')
    wp["A11"] = ('=IF(AND(B7=1,B8=1,B9=1),"Сезонная поправка не применяется",'
                 '"Сезонная поправка применена вручную: прогноз скорректирован пользователем")')
    wp["A10"].font = wp["A11"].font = Font(name=F, size=10, bold=True, color="C00000")
    hh = ["Направление", "IV кв. 2025 факт", *MONTHS26, "I кв. 2026 прогноз", "К IV кварталу, %", "Погрешность отката",
          "Нижняя граница I кв.", "Верхняя граница I кв.", "Надёжность", "код"]
    HR = 13
    for i, h in enumerate(hh, 1):
        wp.cell(HR, i, h)
    head(wp, HR, 1, len(hh))
    wp.row_dimensions[HR].height = 44
    for j, code in enumerate(CODES):
        r = HR + 1 + j
        Lc, Bc = col(3 + 3 * j), col(4 + 3 * j)
        wp.cell(r, 1, DIRS[code])
        wp.cell(r, 2, f"=SUM({DN(j, 11, 13)})")
        for h in range(1, 4):
            shift = f"IF($B$6=1,{h},$B$6*(1-$B$6^{h})/(1-$B$6))"
            wp.cell(r, 2 + h, f"=('Расчёт Хольта'!{Lc}13+{shift}*'Расчёт Хольта'!{Bc}13)*$B${6 + h}")
        wp.cell(r, 6, f"=SUM(C{r}:E{r})")
        wp.cell(r, 7, f"=IF(B{r}=0,0,F{r}/B{r}-1)")
        wp.cell(r, 8, f"=INDEX({BT_S[1]},MATCH($M{r},{BT_S[0]},0))")
        wp.cell(r, 9, f"=F{r}*MAX(0,1-H{r})")
        wp.cell(r, 10, f"=F{r}*(1+H{r})")
        wp.cell(r, 11, f'=IF(H{r}<0.15,"высокая",IF(H{r}<0.3,"средняя","низкая"))')
        wp.cell(r, 13, code)
    rt = HR + 1 + len(CODES)
    wp.cell(rt, 1, "Итого (сумма направлений)")
    for c in range(2, 7):
        L = col(c)
        wp.cell(rt, c, f"=SUM({L}{HR + 1}:{L}{rt - 1})")
    wp.cell(rt, 7, f"=IF(B{rt}=0,0,F{rt}/B{rt}-1)")
    wp.cell(rt, 8, f"=INDEX({BT_S[1]},MATCH(\"Итого\",{BT_S[0]},0))")
    wp.cell(rt, 9, f"=F{rt}*MAX(0,1-H{rt})")
    wp.cell(rt, 10, f"=F{rt}*(1+H{rt})")
    wp.cell(rt, 11, f'=IF(H{rt}<0.15,"высокая",IF(H{rt}<0.3,"средняя","низкая"))')
    for r in range(HR + 1, rt + 1):
        for c in range(1, 12):
            cell = wp.cell(r, c)
            cell.font = BOLD if r == rt else BLACK
            if r == rt:
                cell.border = TOP
            if c in (2, 3, 4, 5, 6, 9, 10):
                cell.number_format = NUM
            elif c in (7, 8):
                cell.number_format = PCT
        wp.cell(r, 6).fill = MAIN_FILL
        wp.cell(r, 13).font = GREY
    for rng_ in (f"K{HR + 1}:K{rt}",):
        wp.conditional_formatting.add(rng_, FormulaRule(
            formula=[f'K{HR + 1}="низкая"'], font=Font(name=F, bold=True, color="C00000"),
            fill=PatternFill("solid", bgColor="F8CBAD")))
        wp.conditional_formatting.add(rng_, FormulaRule(
            formula=[f'K{HR + 1}="средняя"'], font=Font(name=F, bold=True, color="7F6000"),
            fill=PatternFill("solid", bgColor="FFE699")))
    notes = [
        "Тренд с затуханием без сезонности: рост 2025 года принят за рост бизнеса, сезонность по одному году оценить нельзя.",
        "Коридор = прогноз ± средняя абсолютная ошибка отката по направлению (12 наблюдений на направление). Надёжность: ошибка "
        "меньше 15% высокая, 15-30% средняя, больше 30% низкая.",
        "В откате прогноз в среднем занижал факт (бизнес рос быстрее тренда), поэтому риск по верхней границе больше, чем по нижней.",
        "Скачок сентября 2025 (косметология, пластика, лазер) метод не предсказывает: по этим направлениям надёжность низкая.",
    ]
    for k, t in enumerate(notes):
        wp.cell(rt + 2 + k, 1, t).font = HINT
    wp.column_dimensions["A"].width = 38
    for c in range(2, 12):
        wp.column_dimensions[col(c)].width = 14
    wp.column_dimensions["M"].width = 7
    wp.freeze_panes = "B14"

    # ======================= Врачи =======================
    wv = wb.create_sheet("Врачи", 1)
    wv["A1"], wv["A1"].font = "Прогноз выручки по врачам на январь-март 2026 (тыс. ₽)", TITLE
    wv["A2"] = ("Прогноз врача = прогноз направления × доля врача в направлении по среднемесячной выручке за октябрь-декабрь 2025 "
                "(выручка за эти месяцы / число месяцев с записью в payroll). Сумма врачей равна направлению.")
    wv["A2"].font = Font(name=F, size=9, italic=True)
    vh = ["Врач", "ID", "Направление", "Грейд", "Месяцев в IV кв.", "Выручка IV кв.", "Среднемесячная выручка",
          "Доля в направлении", *MONTHS26, "I кв. 2026 прогноз", "К IV кварталу, %", "Надёжность направления", "Примечание"]
    HRV = 4
    for i, h in enumerate(vh, 1):
        wv.cell(HRV, i, h)
    head(wv, HRV, 1, len(vh))
    wv.row_dimensions[HRV].height = 44
    r1, r2 = HRV + 1, HRV + len(dl)
    DR = lambda c: f"'Данные врачей'!${c}$2:${c}${nr}"                        # noqa: E731
    for k, (did, row) in enumerate(dl.iterrows()):
        r = r1 + k
        dr = 2 + k
        wv.cell(r, 1, names[did])
        wv.cell(r, 2, did)
        wv.cell(r, 3, row.direction_code)
        wv.cell(r, 4, row.grade)
        wv.cell(r, 5, f"=SUM('Данные врачей'!{col(16 + 10)}{dr}:{col(16 + 12)}{dr})")
        wv.cell(r, 6, f"=SUM('Данные врачей'!{col(4 + 10)}{dr}:{col(4 + 12)}{dr})")
        wv.cell(r, 7, f"=IF(E{r}=0,0,F{r}/E{r})")
        wv.cell(r, 8, f"=IF(SUMIFS($G${r1}:$G${r2},$C${r1}:$C${r2},C{r})=0,0,G{r}/SUMIFS($G${r1}:$G${r2},$C${r1}:$C${r2},C{r}))")
        for h in range(3):
            pc = col(3 + h)
            wv.cell(r, 9 + h, f"=$H{r}*INDEX('Прогноз'!{pc}${HR + 1}:{pc}${HR + len(CODES)},MATCH($C{r},'Прогноз'!$M${HR + 1}:$M${HR + len(CODES)},0))")
        wv.cell(r, 12, f"=SUM(I{r}:K{r})")
        wv.cell(r, 13, f'=IF(F{r}=0,"—",L{r}/F{r}-1)')
        wv.cell(r, 14, f"=INDEX('Прогноз'!$K${HR + 1}:$K${HR + len(CODES)},MATCH($C{r},'Прогноз'!$M${HR + 1}:$M${HR + len(CODES)},0))")
        wv.cell(r, 15, f'=IF(E{r}=0,"нет выручки в IV квартале",IF(E{r}<3,"стаж в IV кв. меньше 3 мес.: доля по неполному периоду","—"))')
        for c in range(1, len(vh) + 1):
            cell = wv.cell(r, c)
            cell.font = BLACK
            if c == 5:
                cell.number_format = INT
            elif c in (6, 7, 9, 10, 11, 12):
                cell.number_format = NUMK
            elif c == 8:
                cell.number_format = "0.0%"
            elif c == 13:
                cell.number_format = PCT
                cell.alignment = Alignment(horizontal="right")
        wv.cell(r, 12).fill = MAIN_FILL
        wv.cell(r, 15).font = Font(name=F, size=9, italic=True, color="C00000")
    wv.conditional_formatting.add(f"N{r1}:N{r2}", FormulaRule(
        formula=[f'N{r1}="низкая"'], font=Font(name=F, bold=True, color="C00000"),
        fill=PatternFill("solid", bgColor="F8CBAD")))
    rs = r2 + 2
    wv.cell(rs, 1, "Контроль: сумма врачей и прогноз направления (тыс. ₽)").font = BOLD
    for c in range(1, 5):
        wv.cell(rs, c).fill = SEC_FILL
    for i, h in enumerate(["Направление", "Сумма врачей, I кв.", "Прогноз направления, I кв.", "Разница"], 1):
        wv.cell(rs + 1, i, h)
    head(wv, rs + 1, 1, 4)
    for j, code in enumerate(CODES):
        r = rs + 2 + j
        wv.cell(r, 1, DIRS[code])
        wv.cell(r, 2, f"=SUMIFS($L${r1}:$L${r2},$C${r1}:$C${r2},\"{code}\")")
        wv.cell(r, 3, f"='Прогноз'!F{HR + 1 + j}")
        wv.cell(r, 4, f"=B{r}-C{r}")
        for c in range(1, 5):
            wv.cell(r, c).font = BLACK
            if c > 1:
                wv.cell(r, c).number_format = "#,##0.00,;(#,##0.00,);-" if c < 4 else "#,##0.00;(#,##0.00);-"
    for i, w in enumerate([18, 7, 12, 9, 10, 12, 13, 11, 11, 11, 11, 13, 11, 14, 54], 1):
        wv.column_dimensions[col(i)].width = w
    wv.freeze_panes = wv.cell(r1, 3)
    wv.auto_filter.ref = f"A{HRV}:{col(len(vh))}{r2}"
    ctrl_sum = f"=SUM('Врачи'!$L${r1}:$L${r2})"

    # ======================= Описание =======================
    wi = wb.create_sheet("Описание", 0)
    wi["A1"], wi["A1"].font = "Прогноз выручки 2026 (январь-март): описание и контроль", TITLE
    lines = [
        ("Метод", None),
        ("Тренд с затуханием", "Метод Хольта по месяцам 2025 года для каждого направления: уровень и тренд сглаживаются, тренд затухает "
                               "с коэффициентом φ. Параметры α=0,5, β=0,3, φ=0,9 не подбирались по данным. Формулы на листе «Расчёт Хольта»; "
                               "итог по сети это сумма направлений."),
        ("Позиция по сезонности", "Рост 2025 года (34,3 млн ₽ в январе, 102,3 в декабре) принят за рост бизнеса: по одному году сезонность "
                                  "нельзя отделить от разгона, а форма года у плана и у факта одинаковая, то есть план независимой "
                                  "информации не даёт. Сезонная поправка задаётся пользователем вручную (по умолчанию 1,00)."),
        ("Врачи", "Прогноз врача равен прогнозу направления, умноженному на долю врача по среднемесячной выручке за октябрь-декабрь. "
                  "Сумма по врачам равна прогнозу направления. У врачей со стажем в IV квартале меньше трёх месяцев доля считается "
                  "по неполному периоду."),
        ("Проверка и надёжность", None),
        ("Откат", "На 2025 году проверены четыре метода (последний месяц, среднее трёх месяцев, линейный тренд, Хольт): обучение по "
                  "первым 6-9 месяцам, прогноз трёх следующих. Хольт показал наименьшую ошибку. Погрешность сильно зависит от "
                  "направления: стоматология 5%, семейная 12%, итог 20%, косметология 39%, лазер 50%, пластика 78% (скачок сентября)."),
        ("Надёжность", "Коридор равен прогнозу ± ошибке отката по направлению. Ошибка меньше 15% даёт «высокую» надёжность, 15-30% «среднюю», "
                       "больше 30% «низкую». Прогноз по низконадёжным направлениям нельзя использовать как план."),
        ("Ограничения", None),
        ("Один год данных", "Всего 12 месяцев: нет второго года для сезонности, откат содержит 72 зависимых наблюдения, оценка погрешности "
                            "ориентировочная. В откате прогноз в среднем занижал факт: бизнес рос быстрее тренда."),
        ("Не планируется", "Прогноз не учитывает каналы привлечения, загрузку кабинетов, найм и уход врачей, цены, решения директоров. "
                           "Это прогноз по тренду, а не бюджет."),
        ("Данные синтетические", "Модель клиники: цифры демонстрируют метод."),
    ]
    r = 3
    for k, v in lines:
        wi.cell(r, 1, k).font = BOLD
        if v is None:
            wi.cell(r, 1).fill = SEC_FILL
            wi.cell(r, 2).fill = SEC_FILL
        else:
            wi.cell(r, 2, v).font = BLACK
            wi.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
            wi.cell(r, 1).alignment = Alignment(vertical="top")
        r += 1
    r += 1
    wi.cell(r, 1, "Контроль, тыс. ₽").font = BOLD
    for c in (1, 2, 3, 4):
        wi.cell(r, c).fill = SEC_FILL
    r += 1
    for i, h in enumerate(["Показатель", "По книге", "Источник", "Разница"], 1):
        wi.cell(r, i, h)
    head(wi, r, 1, 4)
    ctrl_rows = [("Выручка 2025 по направлениям (визиты и услуги)", "=SUM('Данные направлений'!B2:F13)", ctrl["rev_total"]),
                 ("Выручка 2025 по врачам равна выручке направлений", f"=SUM('Данные врачей'!E2:P{nr})", ctrl["rev_total"]),
                 ("Прогноз I кв. по сети (при исходных параметрах и поправках 1,00)", f"='Прогноз'!F{rt}", ctrl["q1_total"]),
                 ("Прогноз I кв.: сумма врачей и сеть (при любых параметрах)", ctrl_sum, f"='Прогноз'!F{rt}")]
    for name, f, src in ctrl_rows:
        r += 1
        wi.cell(r, 1, name).font = BLACK
        wi.cell(r, 2, f).number_format = "#,##0.00,;(#,##0.00,);-"
        if isinstance(src, str):
            wi.cell(r, 3, src).number_format = "#,##0.00,;(#,##0.00,);-"
        else:
            wi.cell(r, 3, round(float(src), 2)).font = BLUE
            wi.cell(r, 3).number_format = "#,##0.00,;(#,##0.00,);-"
        wi.cell(r, 4, f"=B{r}-C{r}").number_format = "#,##0.00;(#,##0.00);-"
    wi.column_dimensions["A"].width = 56
    wi.column_dimensions["B"].width = 110
    wi.column_dimensions["C"].width = 16
    wi.column_dimensions["D"].width = 14

    wb._sheets = [wb["Описание"], wb["Прогноз"], wb["Врачи"], wb["Откат"], wb["Расчёт Хольта"],
                  wb["Данные направлений"], wb["Данные врачей"]]
    wb.active = 1
    return wb


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    S, rev, wk, dl, bt, methods, by_series, ctrl = prepare()
    wb = build(S, rev, wk, dl, bt, methods, by_series, ctrl)
    path = OUT / "Forecast_2026Q1.xlsx"
    wb.save(path)
    recalculate(path)
    print(f"Сохранено: {path}  (врачей: {len(dl)})")
    print("Прогноз I кв. 2026 по сети:", f"{ctrl['q1_total']:,.0f}", "| выручка 2025:", f"{ctrl['rev_total']:,.0f}")


if __name__ == "__main__":
    main()
