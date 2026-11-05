"""
Шаг 12: ДДС прямым методом, БДДС план-факт и ликвидность за 2025 год в Excel.

Вход:  data/corrected/ - cashflow_1c, osv_1c, budget_revenue, budget_opex, visit_payments, visits
Выход: data/reports/CashFlow_2025.xlsx

Листы:
  Описание     - определения, график платежей, допущения, ограничения, контроль
  ДДС          - факт по месяцам и статьям (прямой метод), остаток на начало и конец (фильтр по клинике)
  БДДС         - план и факт по статьям, отклонения (фильтры: клиника, месяц), остатки по месяцам
  Ликвидность  - остаток на конец и минимальный дневной остаток по месяцам, покрытие в днях выплат,
                 порог резерва (вход пользователя), стресс-сценарий задержки страховых платежей
  Данные, Данные дн, Остатки, Ликвидность дн - источники для формул

План денежного бюджета считается в src/common/cashplan.py (плановые начисления x график платежей).
Остаток показан до налога на прибыль, НДС и дивидендов: в модели они не учтены.

Запуск:  python src/reports/build_cashflow.py
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.recalc import recalculate  # noqa: E402
from common import cashplan  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
YEAR = 2025
CLINICS = ["C01", "C02", "C03"]
MONTHS = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
OUTS = [a for act, a in cashplan.ARTICLES if act == "operating"][2:]
INS = [a for act, a in cashplan.ARTICLES if act == "operating"][:2]
INV = [a for act, a in cashplan.ARTICLES if act == "investing"]

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
NUM = "#,##0,;(#,##0,);-"
PCT = "0.0%;(0.0%);-"


# ------------------------------------------------------------------ данные
def prepare():
    cash = pd.read_csv(CORR / "cashflow_1c.csv")
    osv = pd.read_csv(CORR / "osv_1c.csv", dtype={"account": str})
    brev = pd.read_csv(CORR / "budget_revenue.csv")
    box = pd.read_csv(CORR / "budget_opex.csv")
    pay = pd.read_csv(CORR / "visit_payments.csv")
    visits = pd.read_csv(CORR / "visits.csv")

    cash["period"] = cash.date.str[:7]
    opening = cashplan.opening_balances(osv)
    share = cashplan.dms_share(pay, visits)
    plan = cashplan.plan_cash(brev, box, share, opening)
    fact = cash.groupby(["period", "clinic_id", "activity", "article"], as_index=False).signed_amount.sum() \
               .rename(columns={"signed_amount": "fact"})
    m = fact.merge(plan.rename(columns={"plan": "plan"}), on=["period", "clinic_id", "activity", "article"], how="outer")
    m[["fact", "plan"]] = m[["fact", "plan"]].fillna(0.0)
    m["month"] = m.period.str[5:7].astype(int)
    m = m[m.month.between(1, 12) & (m.period.str[:4] == str(YEAR))]
    m = m[["period", "month", "clinic_id", "activity", "article", "fact", "plan"]] \
        .sort_values(["month", "clinic_id", "activity", "article"]).reset_index(drop=True)
    m[["fact", "plan"]] = m[["fact", "plan"]].round(2)

    days = pd.date_range(f"{YEAR}-01-01", f"{YEAR}-12-31")
    cash["d"] = pd.to_datetime(cash.date)
    cash["inflow"] = cash.signed_amount.clip(lower=0)
    cash["outflow"] = cash.signed_amount.clip(upper=0)
    cash["ins"] = cash.signed_amount.where(cash.article == "Поступления от страховых компаний", 0.0)
    daily = cash.groupby(["d", "clinic_id"])[["inflow", "outflow", "ins"]].sum()
    idx = pd.MultiIndex.from_product([days, CLINICS], names=["d", "clinic_id"])
    daily = daily.reindex(idx, fill_value=0.0).reset_index()

    close_dec = osv[(osv.period == f"{YEAR}-12") & (osv.account == "51")].closing_debit.sum()
    ctrl = {"net_fact": cash.signed_amount.sum(), "close_dec": float(close_dec), "net_plan": plan.plan.sum()}
    return m, daily, opening["51"], ctrl


# ------------------------------------------------------------------ оформление
def head(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put_filter(ws, row, label, value, options, hint, crit_formula=None):
    ws.cell(row, 1, label).font = BOLD
    c = ws.cell(row, 2, value)
    c.fill, c.font = INPUT_FILL, BLUE
    if crit_formula:
        ws.cell(row, 3, crit_formula).font = GREY
    ws.cell(row, 4, hint).font = HINT
    if options:
        dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(c)


def build(m, daily, open51, ctrl):
    wb = Workbook()

    # ---------- Данные (месяц x клиника x статья) ----------
    wd = wb.active
    wd.title = "Данные"
    for i, h in enumerate(["Период", "Месяц", "Клиника", "Вид деятельности", "Статья ДДС", "Факт", "План"], 1):
        wd.cell(1, i, h)
    head(wd, 1, 1, 7)
    for r, row in enumerate(m.itertuples(index=False), 2):
        for c, v in enumerate(row, 1):
            cell = wd.cell(r, c, v)
            cell.font = BLUE if c >= 6 else BLACK
            if c >= 6:
                cell.number_format = "#,##0.00"
    nm = len(m) + 1
    for i, w in enumerate([10, 8, 9, 14, 36, 16, 16], 1):
        wd.column_dimensions[col(i)].width = w
    wd.freeze_panes = "A2"
    M = lambda c: f"'Данные'!${c}$2:${c}${nm}"                                  # noqa: E731

    # ---------- Остатки ----------
    wo = wb.create_sheet("Остатки")
    wo["A1"], wo["B1"] = "Клиника", "Остаток на 01.01, ₽ (ОСВ, счёт 51)"
    head(wo, 1, 1, 2)
    for r, c in enumerate(CLINICS, 2):
        wo.cell(r, 1, c).font = BLACK
        wo.cell(r, 2, float(open51[c])).font = BLUE
        wo.cell(r, 2).number_format = "#,##0"
    wo.column_dimensions["A"].width = 10
    wo.column_dimensions["B"].width = 34
    OP = lambda c: f"'Остатки'!${c}$2:${c}${len(CLINICS) + 1}"                  # noqa: E731

    # ---------- Данные дн (день x клиника) ----------
    wn = wb.create_sheet("Данные дн")
    for i, h in enumerate(["Дата", "Клиника", "Поступления", "Выплаты", "из них страховые"], 1):
        wn.cell(1, i, h)
    head(wn, 1, 1, 5)
    for r, row in enumerate(daily.itertuples(index=False), 2):
        wn.cell(r, 1, row.d.to_pydatetime()).number_format = "dd.mm.yyyy"
        wn.cell(r, 2, row.clinic_id)
        for c, v in ((3, row.inflow), (4, row.outflow), (5, row.ins)):
            cell = wn.cell(r, c, round(float(v), 2))
            cell.font = BLUE
            cell.number_format = "#,##0.00"
    nn = len(daily) + 1
    for i, w in enumerate([12, 9, 16, 16, 16], 1):
        wn.column_dimensions[col(i)].width = w
    wn.freeze_panes = "A2"
    DD = lambda c: f"'Данные дн'!${c}$2:${c}${nn}"                              # noqa: E731

    # ======================= ДДС =======================
    wc = wb.create_sheet("ДДС", 0)
    wc["A1"], wc["A1"].font = "ДДС прямым методом за 2025 год (тыс. ₽, до налога на прибыль, НДС и дивидендов)", TITLE
    put_filter(wc, 2, "Клиника (фильтр):", "Все", ["Все"] + CLINICS, "← «Все», C01, C02 или C03",
               '=IF($B$2="Все","*",$B$2)')
    wc.column_dimensions["A"].width = 58

    def cash_rows(ws, start, per_formula, first_open):
        """Общая структура строк отчёта; per_formula(r, article_cell) -> формула суммы по статье."""
        rows = {}
        r = start
        ws.cell(r, 1, "Операционная деятельность").font = BOLD
        r += 1
        for a in INS:
            rows[a] = r; r += 1
        rows["tin"] = r; r += 1
        for a in OUTS:
            rows[a] = r; r += 1
        rows["tout"] = r; r += 1
        rows["op"] = r; r += 1
        ws.cell(r, 1, "Инвестиционная деятельность").font = BOLD
        r += 1
        for a in INV:
            rows[a] = r; r += 1
        rows["inv"] = r; r += 1
        rows["net"] = r; r += 1
        rows["open"] = r; r += 1
        rows["close"] = r
        return rows

    # структура строк ДДС
    R = cash_rows(wc, 5, None, None)
    for c in range(1, 15):
        wc.cell(5, c).fill = SEC_FILL
        wc.cell(R["inv"] - 2, c).fill = SEC_FILL
    for i, mn in enumerate(MONTHS):
        wc.cell(4, 2 + i, mn)
        wc.cell(3, 2 + i, i + 1).font = GREY
    wc.cell(4, 1, "Статья")
    wc.cell(4, 14, "Итого")
    head(wc, 4, 1, 14)
    first_m, last_m = 2, 13
    for name in INS + OUTS + INV:
        r = R[name]
        wc.cell(r, 1, name)
        for c in range(first_m, last_m + 1):
            L = col(c)
            wc.cell(r, c, f"=SUMIFS({M('F')},{M('B')},{L}$3,{M('C')},$C$2,{M('E')},$A{r})")
        wc.cell(r, 14, f"=SUM(B{r}:M{r})")
    labels = {"tin": "Итого поступления", "tout": "Итого выплаты",
              "op": "Чистый денежный поток от операционной деятельности",
              "inv": "Чистый денежный поток от инвестиционной деятельности", "net": "Чистое изменение денежных средств",
              "open": "Остаток на начало периода", "close": "Остаток на конец периода"}
    for k, t in labels.items():
        wc.cell(R[k], 1, t)
    for c in range(first_m, 15):
        L = col(c)
        wc.cell(R["tin"], c, f"=SUM({L}{R[INS[0]]}:{L}{R[INS[-1]]})")
        wc.cell(R["tout"], c, f"=SUM({L}{R[OUTS[0]]}:{L}{R[OUTS[-1]]})")
        wc.cell(R["op"], c, f"={L}{R['tin']}+{L}{R['tout']}")
        wc.cell(R["inv"], c, f"=SUM({L}{R[INV[0]]}:{L}{R[INV[-1]]})")
        wc.cell(R["net"], c, f"={L}{R['op']}+{L}{R['inv']}")
    wc.cell(R["open"], 2, f"=SUMIFS({OP('B')},{OP('A')},$C$2)")
    for c in range(3, 14):
        wc.cell(R["open"], c, f"={col(c - 1)}{R['close']}")
    wc.cell(R["open"], 14, f"=B{R['open']}")
    for c in range(2, 14):
        wc.cell(R["close"], c, f"={col(c)}{R['open']}+{col(c)}{R['net']}")
    wc.cell(R["close"], 14, f"=M{R['close']}")
    for r in range(R[INS[0]], R["close"] + 1):
        bold = r in (R["tin"], R["tout"], R["op"], R["inv"], R["net"], R["close"])
        for c in range(1, 15):
            cell = wc.cell(r, c)
            cell.font = BOLD if bold else BLACK
            if c > 1:
                cell.number_format = NUM
            if bold and wc.cell(r, 1).value:
                cell.border = TOP
    for c in range(2, 15):
        wc.column_dimensions[col(c)].width = 10
    wc.freeze_panes = "B5"
    wc.cell(R["close"] + 2, 1, "Остаток до налога на прибыль, НДС и дивидендов: в модели они не учтены, "
                               "реальный остаток ниже.").font = HINT

    # ======================= БДДС =======================
    wb_ = wb.create_sheet("БДДС", 1)
    wb_["A1"], wb_["A1"].font = "БДДС: план и факт денежных потоков за 2025 год (тыс. ₽)", TITLE
    put_filter(wb_, 2, "Клиника (фильтр):", "Все", ["Все"] + CLINICS, "← «Все», C01, C02 или C03",
               '=IF($B$2="Все","*",$B$2)')
    put_filter(wb_, 3, "Месяц (фильтр):", "Все", ["Все"] + [str(i) for i in range(1, 13)],
               "← «Все» или номер месяца 1-12")
    wb_["C3"] = '=IF($B$3="Все",1,VALUE($B$3))'
    wb_["D3"] = '=IF($B$3="Все",12,VALUE($B$3))'
    wb_["C3"].font = wb_["D3"].font = GREY
    wb_["E3"] = "← месяц от и до"
    wb_["E3"].font = HINT
    wb_.column_dimensions["A"].width = 58
    RB = cash_rows(wb_, 6, None, None)
    for c in range(1, 6):
        wb_.cell(6, c).fill = SEC_FILL
        wb_.cell(RB["inv"] - 2, c).fill = SEC_FILL
    for i, h in enumerate(["Статья", "План", "Факт", "Отклонение (факт − план)", "Отклонение, %"], 1):
        wb_.cell(5, i, h)
    head(wb_, 5, 1, 5)
    wb_.row_dimensions[5].height = 30
    crit = f'{M("B")},">="&$C$3,{M("B")},"<="&$D$3,{M("C")},$C$2'
    for name in INS + OUTS + INV:
        r = RB[name]
        wb_.cell(r, 1, name)
        wb_.cell(r, 2, f"=SUMIFS({M('G')},{crit},{M('E')},$A{r})")
        wb_.cell(r, 3, f"=SUMIFS({M('F')},{crit},{M('E')},$A{r})")
    for k, t in labels.items():
        wb_.cell(RB[k], 1, t)
    for c in (2, 3):
        L = col(c)
        wb_.cell(RB["tin"], c, f"=SUM({L}{RB[INS[0]]}:{L}{RB[INS[-1]]})")
        wb_.cell(RB["tout"], c, f"=SUM({L}{RB[OUTS[0]]}:{L}{RB[OUTS[-1]]})")
        wb_.cell(RB["op"], c, f"={L}{RB['tin']}+{L}{RB['tout']}")
        wb_.cell(RB["inv"], c, f"=SUM({L}{RB[INV[0]]}:{L}{RB[INV[-1]]})")
        wb_.cell(RB["net"], c, f"={L}{RB['op']}+{L}{RB['inv']}")
    for c, src in ((2, "G"), (3, "F")):
        wb_.cell(RB["open"], c, f'=SUMIFS({OP("B")},{OP("A")},$C$2)+SUMIFS({M(src)},{M("B")},"<"&$C$3,{M("C")},$C$2)')
        wb_.cell(RB["close"], c, f"={col(c)}{RB['open']}+{col(c)}{RB['net']}")
    for r in range(RB[INS[0]], RB["close"] + 1):
        if wb_.cell(r, 1).value is None or r in (6, RB["inv"] - 2):
            continue
        wb_.cell(r, 4, f"=C{r}-B{r}")
        wb_.cell(r, 5, f"=IF(B{r}=0,0,D{r}/ABS(B{r}))")
        bold = r in (RB["tin"], RB["tout"], RB["op"], RB["inv"], RB["net"], RB["close"])
        for c in range(1, 6):
            cell = wb_.cell(r, c)
            cell.font = BOLD if bold else BLACK
            cell.number_format = PCT if c == 5 else (NUM if c > 1 else "General")
            if bold:
                cell.border = TOP
    wb_.cell(RB["close"] + 1, 1, "Для выплат положительное отклонение значит, что заплатили меньше плана.").font = HINT

    rm = RB["close"] + 4
    wb_.cell(rm - 1, 1, "Остаток на конец месяца: план и факт").font = BOLD
    for c in range(1, 6):
        wb_.cell(rm - 1, c).fill = SEC_FILL
    for i, h in enumerate(["Месяц", "План", "Факт", "Отклонение (факт − план)", ""], 1):
        wb_.cell(rm, i, h)
    head(wb_, rm, 1, 4)
    wb_.row_dimensions[rm].height = 30
    for i, mn in enumerate(MONTHS):
        r = rm + 1 + i
        wb_.cell(r, 1, f"{i + 1}. {mn}")
        wb_.cell(r, 2, f'=SUMIFS({OP("B")},{OP("A")},$C$2)+SUMIFS({M("G")},{M("B")},"<={i + 1}",{M("C")},$C$2)')
        wb_.cell(r, 3, f'=SUMIFS({OP("B")},{OP("A")},$C$2)+SUMIFS({M("F")},{M("B")},"<={i + 1}",{M("C")},$C$2)')
        wb_.cell(r, 4, f"=C{r}-B{r}")
        for c in range(1, 5):
            cell = wb_.cell(r, c)
            cell.font = BLACK
            if c > 1:
                cell.number_format = NUM
    for c in range(2, 6):
        wb_.column_dimensions[col(c)].width = 16
    wb_.freeze_panes = "B6"

    # ======================= Ликвидность дн =======================
    wl = wb.create_sheet("Ликвидность дн")
    for i, h in enumerate(["Дата", "Месяц", "Поступления", "Выплаты", "из них страховые", "Остаток на конец дня",
                           "Остаток при задержке страховых"], 1):
        wl.cell(1, i, h)
    head(wl, 1, 1, 7)
    days = pd.date_range(f"{YEAR}-01-01", f"{YEAR}-12-31")
    LQ = "'Ликвидность'!"
    last = len(days) + 1
    for k, d in enumerate(days):
        r = 2 + k
        wl.cell(r, 1, d.to_pydatetime()).number_format = "dd.mm.yyyy"
        wl.cell(r, 2, d.month)
        for c, src in ((3, "C"), (4, "D"), (5, "E")):
            wl.cell(r, c, f"=SUMIFS({DD(src)},{DD('A')},$A{r},{DD('B')},{LQ}$C$2)")
        wl.cell(r, 6, (f"=SUMIFS({OP('B')},{OP('A')},{LQ}$C$2)+C{r}+D{r}" if k == 0 else f"=F{r - 1}+C{r}+D{r}"))
        wl.cell(r, 7, f'=F{r}-SUMIFS($E$2:$E${last},$A$2:$A${last},">"&($A{r}-{LQ}$B$4),$A$2:$A${last},"<="&$A{r})')
        for c in range(3, 8):
            wl.cell(r, c).number_format = "#,##0"
            wl.cell(r, c).font = BLACK
    for i, w in enumerate([12, 8, 14, 14, 14, 18, 20], 1):
        wl.column_dimensions[col(i)].width = w
    wl.freeze_panes = "A2"
    bounds = {}
    for k, d in enumerate(days):
        lo, hi = bounds.get(d.month, (2 + k, 2 + k))
        bounds[d.month] = (lo, 2 + k)

    # ======================= Ликвидность =======================
    wq = wb.create_sheet("Ликвидность", 2)
    wq["A1"], wq["A1"].font = "Ликвидность 2025: остаток, запас в днях выплат и стресс по страховым (тыс. ₽)", TITLE
    put_filter(wq, 2, "Клиника (фильтр):", "Все", ["Все"] + CLINICS, "← «Все», C01, C02 или C03",
               '=IF($B$2="Все","*",$B$2)')
    put_filter(wq, 3, "Минимальный резерв, дней выплат:", 10, None,
               "← запас должен покрывать столько дней среднедневных выплат месяца")
    put_filter(wq, 4, "Стресс: задержка страховых платежей, дней:", 30, None,
               "← каждое поступление от страховых сдвигается на столько дней (0 = без стресса)")
    wq.column_dimensions["A"].width = 54
    hh = ["Месяц", "Остаток на начало", "Поступления", "Выплаты", "Остаток на конец", "Минимум дня",
          "Дата минимума", "Среднедневные выплаты", "Запас на конец, дней", "Запас в минимуме, дней", "Статус",
          "Стресс: остаток на конец", "Стресс: минимум дня", "Стресс: запас в минимуме, дней", "Стресс: статус"]
    HR = 6
    for i, h in enumerate(hh, 1):
        wq.cell(HR, i, h)
    head(wq, HR, 1, len(hh))
    wq.row_dimensions[HR].height = 44
    L2 = "'Ликвидность дн'!"
    for i, mn in enumerate(MONTHS):
        r = HR + 1 + i
        lo, hi = bounds[i + 1]
        rng = lambda c: f"{L2}${c}${lo}:${c}${hi}"                              # noqa: E731
        wq.cell(r, 1, f"{i + 1}. {mn}")
        wq.cell(r, 2, f"=SUMIFS({OP('B')},{OP('A')},$C$2)" if i == 0 else f"=E{r - 1}")
        wq.cell(r, 3, f"=SUM({rng('C')})")
        wq.cell(r, 4, f"=SUM({rng('D')})")
        wq.cell(r, 5, f"=B{r}+C{r}+D{r}")
        wq.cell(r, 6, f"=MIN({rng('F')})")
        wq.cell(r, 7, f"=INDEX({rng('A')},MATCH(F{r},{rng('F')},0))")
        wq.cell(r, 8, f"=-D{r}/COUNT({rng('A')})")
        wq.cell(r, 9, f"=IF(H{r}=0,0,E{r}/H{r})")
        wq.cell(r, 10, f"=IF(H{r}=0,0,F{r}/H{r})")
        wq.cell(r, 11, f'=IF(F{r}<0,"разрыв",IF(J{r}<$B$3,"ниже резерва","ок"))')
        wq.cell(r, 12, f"={L2}$G${hi}")
        wq.cell(r, 13, f"=MIN({rng('G')})")
        wq.cell(r, 14, f"=IF(H{r}=0,0,M{r}/H{r})")
        wq.cell(r, 15, f'=IF(M{r}<0,"разрыв",IF(N{r}<$B$3,"ниже резерва","ок"))')
        for c in range(1, len(hh) + 1):
            cell = wq.cell(r, c)
            cell.font = BLACK
            if c in (2, 3, 4, 5, 6, 8, 12, 13):
                cell.number_format = NUM
            elif c == 7:
                cell.number_format = "dd.mm"
            elif c in (9, 10, 14):
                cell.number_format = "0.0"
    r1, r2 = HR + 1, HR + 12
    for rng_, col_ in ((f"K{r1}:K{r2}", "K"), (f"O{r1}:O{r2}", "O")):
        wq.conditional_formatting.add(rng_, FormulaRule(
            formula=[f'{col_}{r1}="разрыв"'], font=Font(name=F, bold=True, color="FFFFFF"),
            fill=PatternFill("solid", bgColor="C00000")))
        wq.conditional_formatting.add(rng_, FormulaRule(
            formula=[f'{col_}{r1}="ниже резерва"'], font=Font(name=F, bold=True, color="C00000"),
            fill=PatternFill("solid", bgColor="F8CBAD")))
    rs = r2 + 2
    wq.cell(rs, 1, "Месяцев со статусом не «ок»: факт / стресс").font = BOLD
    wq.cell(rs, 2, f'=COUNTIF(K{r1}:K{r2},"<>ок")').font = BOLD
    wq.cell(rs, 3, f'=COUNTIF(O{r1}:O{r2},"<>ок")').font = BOLD
    wq.cell(rs + 1, 1, "Минимальный дневной остаток за год: факт / стресс").font = BOLD
    wq.cell(rs + 1, 2, f"=MIN(F{r1}:F{r2})").number_format = NUM
    wq.cell(rs + 1, 3, f"=MIN(M{r1}:M{r2})").number_format = NUM
    wq.cell(rs + 3, 1, "Статусы: «разрыв» остаток ниже нуля; «ниже резерва» запас в минимуме меньше заданного числа дней "
                       "среднедневных выплат месяца. Остаток до налога на прибыль, НДС и дивидендов.").font = HINT
    for c in range(2, len(hh) + 1):
        wq.column_dimensions[col(c)].width = 13
    wq.freeze_panes = "B7"

    # ======================= Описание =======================
    wi = wb.create_sheet("Описание", 0)
    wi["A1"], wi["A1"].font = "ДДС, БДДС и ликвидность 2025: описание и контроль", TITLE
    lines = [
        ("Что считаем", None),
        ("ДДС факт", "Прямой метод: фактические поступления и выплаты по статьям из cashflow_1c, по месяцам; остаток на начало берётся "
                     "из ОСВ (счёт 51), на конец равен остатку на начало плюс чистое изменение."),
        ("БДДС план", "Плановые начисления (budget_revenue, budget_opex) x график платежей: поступления от пациентов в месяце услуги, "
                      "от страховых через месяц; зарплата, НДФЛ, взносы, материалы и маркетинг в следующем месяце; аренда, "
                      "коммунальные и эквайринг в месяце начисления. Входящие сальдо на 01.01 (ОСВ) гасятся в январе, "
                      "дебиторка ДМС поровну в январе и феврале."),
        ("Ликвидность", "Остаток на конец месяца и минимальный дневной остаток в месяце, запас в днях среднедневных выплат месяца. "
                        "Статус «разрыв» остаток ниже нуля, «ниже резерва» запас меньше заданного резерва. Стресс сдвигает "
                        "поступления от страховых на заданное число дней."),
        ("Допущения", None),
        ("Доли ДМС и инвестиции", "Доли ДМС по направлениям в плане приняты равными фактическим долям 2025, инвестпрограмма равна "
                                  "факту (закупки оборудования C01 март, C03 июнь, C02 сентябрь), потому что в исходном бюджете "
                                  "их нет. Отклонения плана от факта поэтому объясняются выручкой, расходами и сроками платежей."),
        ("График платежей", "Принят по фактическим датам платежей: зарплата 10-го, НДФЛ и взносы 28-го, материалы 20-го, "
                            "маркетинг 15-го числа следующего месяца, аренда 5-го текущего."),
        ("Ограничения", None),
        ("Налоги и дивиденды", "Налог на прибыль, НДС и дивиденды в модели не учтены ни в плане, ни в факте, поэтому остаток завышен "
                               "и показан «до налога на прибыль, НДС и дивидендов». Финансовой деятельности (кредиты) нет."),
        ("Разрывы", "В фактических данных остаток ни разу не ниже нуля. Показатель запаса и стресс-сценарий показывают, насколько "
                    "близко к риску была сеть, а не выдуманный разрыв."),
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
    wi.cell(r, 1, "Контроль (при фильтрах «Все»), тыс. ₽").font = BOLD
    for c in (1, 2, 3, 4):
        wi.cell(r, c).fill = SEC_FILL
    r += 1
    for i, h in enumerate(["Показатель", "По книге", "Источник", "Разница"], 1):
        wi.cell(r, i, h)
    head(wi, r, 1, 4)
    ctrl_rows = [("Факт: чистое изменение денежных средств", f"='ДДС'!N{R['net']}", ctrl["net_fact"]),
                 ("Факт: остаток на конец года (ОСВ, счёт 51)", f"='ДДС'!N{R['close']}", ctrl["close_dec"]),
                 ("План: чистое изменение денежных средств (cashplan.py)", f"='БДДС'!B{RB['net']}", ctrl["net_plan"]),
                 ("Ликвидность: остаток на конец года", f"='Ликвидность'!E{HR + 12}", ctrl["close_dec"])]
    for name, f, src in ctrl_rows:
        r += 1
        wi.cell(r, 1, name).font = BLACK
        wi.cell(r, 2, f).number_format = "#,##0.00,;(#,##0.00,);-"
        wi.cell(r, 3, round(float(src), 2)).font = BLUE
        wi.cell(r, 3).number_format = "#,##0.00,;(#,##0.00,);-"
        wi.cell(r, 4, f"=B{r}-C{r}").number_format = "#,##0.00;(#,##0.00);-"
    wi.column_dimensions["A"].width = 52
    wi.column_dimensions["B"].width = 110
    wi.column_dimensions["C"].width = 16
    wi.column_dimensions["D"].width = 14

    wb._sheets = [wb["Описание"], wb["ДДС"], wb["БДДС"], wb["Ликвидность"], wb["Данные"], wb["Данные дн"],
                  wb["Ликвидность дн"], wb["Остатки"]]
    wb.active = 1
    return wb


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    m, daily, open51, ctrl = prepare()
    wb = build(m, daily, open51, ctrl)
    path = OUT / "CashFlow_2025.xlsx"
    wb.save(path)
    recalculate(path)
    print(f"Сохранено: {path}  (строк данных: {len(m)}, дней x клиник: {len(daily)})")
    print("Факт: чистое изменение", f"{ctrl['net_fact']:,.0f}", "| остаток на конец года", f"{ctrl['close_dec']:,.0f}",
          "| план: чистое изменение", f"{ctrl['net_plan']:,.0f}")


if __name__ == "__main__":
    main()
