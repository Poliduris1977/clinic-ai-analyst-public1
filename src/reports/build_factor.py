"""
Шаг 11: факторный анализ прямой маржи (план -> факт) за 2025 год в Excel.

Прямая маржа = выручка - материалы - лаборатории и внешние услуги: ровно те статьи,
которые в плане заданы на уровне услуг (количество, цена, себестоимость единицы).

Достройка до валовой маржи (лист «Валовая маржа»): к шести факторам добавляются
  Врачи: оплата следует за выручкой = плановая оплата - (фактическая база выплат x плановая доля направления)
         (гибкий план: оплата врачей - процент от выручки, поэтому при недоборе выручки она падает сама)
  Врачи: ставка и грейды = -(фактическая оплата - гибкий план); раскладывается по грейдам (лист «Врачи по грейдам»)
  Медперсонал, амортизация = план - факт (фиксированные, заданы по клинике, не по направлению)
  Взносы = ставка x (эффекты врачей и медперсонала)
Плановая доля врачей направления = простое среднее pay_pct врачей клиники и направления
(так построен бюджет), поэтому «грейды» отчасти отражают упрощение плана, а не операционную проблему.

Вход:  data/corrected/ - visits, visit_items, services (факт по услугам),
                         budget_revenue (план по услугам), materials_1c (факт затрат из 1С),
                         payroll_1c, doctors (оплата врачей по грейдам), budget_opex (через build_planfact)
Выход: data/reports/FactorAnalysis_2025.xlsx

Метод. Для каждой строки «месяц x клиника x услуга» маржа единицы m = цена - материалы - лаборатории.
Отклонение факта от плана раскладывается точно, без остатка, на 6 факторов:

  Объём             = плановая маржа ячейки x (индекс объёма - 1); индекс объёма считается
                      в плановых ценах по ячейке «месяц x клиника x направление»:
                      V = сумма(факт. кол-во x плановая цена) / сумма(плановое кол-во x плановая цена)
  Структура услуг   = сумма(факт. кол-во x плановая маржа единицы) - V x плановая маржа ячейки
                      (сдвиг состава услуг внутри направления в сторону более или менее маржинальных)
  Цена: прайс       = сумма(факт. кол-во x действовавший прайс) - сумма(факт. кол-во x плановая цена)
                      (план закладывал одну индексацию, факт - другую)
  Цена: скидки      = факт. выручка - сумма(факт. кол-во x действовавший прайс)
  Материалы         = -(факт. материалы - факт. кол-во x плановые материалы единицы)
  Лаборатории       = -(факт. лаборатории - факт. кол-во x плановые лаборатории единицы)

Плюс = выгодно для маржи, минус = невыгодно. Книга живая: «Данные» содержат входные числа и
формулы эффектов по строкам, отчётные листы считаются SUMIFS.

Факт затрат: 1С хранит материалы и лаборатории по направлениям. На услуги они разнесены
пропорционально нормативу (количество x норма единицы), поэтому итоги совпадают с 1С.

Запуск:  python src/reports/build_factor.py
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.recalc import recalculate  # noqa: E402
import build_planfact as bpf  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
YEAR = 2025
CLINICS = ["C01", "C02", "C03"]
DIRS = {"DENT": "Стоматология", "FAM": "Семейная медицина", "COSM": "Косметология",
        "PLAS": "Пластическая хирургия", "LASER": "Лазерная коррекция зрения"}
CELL = ["period", "clinic_id", "direction_code"]

FACTORS = [("Объём", "W"), ("Структура услуг", "X"), ("Цена: прайс (индексация)", "Y"),
           ("Цена: скидки", "Z"), ("Материалы", "AA"), ("Лаборатории и внешние услуги", "AB")]

F = "Arial"
BLACK, BOLD = Font(name=F, size=10), Font(name=F, size=10, bold=True)
BLUE, GREEN_LINK = Font(name=F, size=10, color="0000FF"), Font(name=F, size=10, color="008000")
RED, GREEN = Font(name=F, size=10, color="C00000"), Font(name=F, size=10, color="00803C")
TITLE, NOTE = Font(name=F, size=14, bold=True), Font(name=F, size=9, italic=True, color="595959")
HDR_FONT = Font(name=F, size=10, bold=True, color="FFFFFF")
HDR_FILL, SEC_FILL = PatternFill("solid", fgColor="1F3864"), PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFFF99")
TOP = Border(top=Side(style="thin"))
NUM = "#,##0,;(#,##0,);-"
PCT = "0.0%;(0.0%);-"
PP = '0.0" п.п.";(0.0" п.п.");-'


# ------------------------------------------------------------------ данные
def prepare():
    visits = pd.read_csv(CORR / "visits.csv")
    items = pd.read_csv(CORR / "visit_items.csv")
    services = pd.read_csv(CORR / "services.csv")
    plan = pd.read_csv(CORR / "budget_revenue.csv")
    mat1c = pd.read_csv(CORR / "materials_1c.csv")

    im = items.merge(visits[["visit_id", "visit_date", "clinic_id"]], on="visit_id").merge(
        services[["service_id", "service_name", "material_cost", "external_cost"]], on="service_id")
    im["period"] = im.visit_date.str[:7]
    im["list_total"] = im.list_price_at_date * im.qty
    im["norm_mat"] = im.qty * im.material_cost
    im["norm_ext"] = im.qty * im.external_cost
    fact = im.groupby(["period", "clinic_id", "service_id"], as_index=False).agg(
        fact_qty=("qty", "sum"), fact_revenue=("revenue", "sum"), list_total=("list_total", "sum"),
        norm_mat=("norm_mat", "sum"), norm_ext=("norm_ext", "sum"))

    d = plan.merge(fact, on=["period", "clinic_id", "service_id"], how="outer", validate="1:1")
    if d[["plan_qty", "fact_qty"]].isna().any().any():
        raise ValueError("Ячейки плана и факта не совпадают: нужен разбор")
    d = d.merge(services[["service_id", "service_name"]], on="service_id")
    d["direction_code"] = d.direction_code.fillna("")

    # факт затрат 1С (по направлениям) -> по услугам пропорционально нормативу
    cell_norm = d.groupby(CELL)[["norm_mat", "norm_ext"]].transform("sum")
    m = mat1c[CELL + ["material_fact", "external_fact"]]
    d = d.merge(m, on=CELL, how="left")
    d["mat_fact"] = np.where(cell_norm.norm_mat > 0, d.material_fact * d.norm_mat / cell_norm.norm_mat.replace(0, 1), 0.0)
    d["ext_fact"] = np.where(cell_norm.norm_ext > 0, d.external_fact * d.norm_ext / cell_norm.norm_ext.replace(0, 1), 0.0)

    d = d.sort_values(["period", "clinic_id", "direction_code", "service_id"]).reset_index(drop=True)
    d = d[["period", "clinic_id", "direction_code", "service_id", "service_name", "plan_qty", "fact_qty",
           "plan_price", "list_total", "fact_revenue", "plan_unit_material", "plan_unit_external",
           "mat_fact", "ext_fact"]]
    for c in ("list_total", "fact_revenue", "mat_fact", "ext_fact"):
        d[c] = d[c].round(2)
    return d, mat1c


def effects(d):
    """То же разложение в pandas: контроль формул Excel и таблицы «Топ факторов»."""
    e = d.copy()
    e["q_f_pp"] = e.fact_qty * e.plan_price
    e["q_p_pp"] = e.plan_qty * e.plan_price
    e["m_unit"] = e.plan_price - e.plan_unit_material - e.plan_unit_external
    e["plan_margin"] = e.plan_qty * e.m_unit
    e["vol_idx"] = e.groupby(CELL).q_f_pp.transform("sum") / e.groupby(CELL).q_p_pp.transform("sum")
    e["volume"] = e.plan_margin * (e.vol_idx - 1)
    e["mix"] = e.fact_qty * e.m_unit - e.vol_idx * e.plan_margin
    e["price_list"] = e.list_total - e.q_f_pp
    e["discount"] = e.fact_revenue - e.list_total
    e["materials"] = -(e.mat_fact - e.fact_qty * e.plan_unit_material)
    e["labs"] = -(e.ext_fact - e.fact_qty * e.plan_unit_external)
    e["fact_margin"] = e.fact_revenue - e.mat_fact - e.ext_fact
    e["check"] = e.fact_margin - (e.plan_margin + e.volume + e["mix"] + e.price_list + e.discount + e.materials + e.labs)
    e["plan_revenue"] = e.q_p_pp
    return e


MEANING = {
    "Объём": "Больше или меньше услуг, чем в плане, при плановых ценах и составе",
    "Структура услуг": "Состав услуг внутри направления сместился к более или менее маржинальным",
    "Цена: прайс (индексация)": "Действовавший прайс отличается от плановой цены (другая индексация)",
    "Цена: скидки": "Скидки пациентам снизили факт. цену относительно прайса",
    "Материалы": "Материалы на единицу услуги дороже или дешевле плана",
    "Лаборатории и внешние услуги": "Лаборатории и внешние услуги на единицу дороже или дешевле плана"}


# ------------------------------------------------------------------ оформление
def head(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def sign_format(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN))


# ------------------------------------------------------------------ ФОТ и валовая маржа
GRADES = ["junior", "middle", "senior", "lead"]


def prepare_fot():
    """Оплата врачей по грейдам, план врачей по ячейкам, общие затраты клиник (медперсонал, амортизация)."""
    pay = pd.read_csv(CORR / "payroll_1c.csv")
    doc = pd.read_csv(CORR / "doctors.csv")
    docp = pay[pay.role == "doctor"].merge(doc[["doctor_id", "grade"]], left_on="employee_id", right_on="doctor_id")
    rate = round(docp.contributions.sum() / docp.gross.sum(), 4)
    plan_pct = doc.groupby(["clinic_id", "direction_code"]).pay_pct.mean().rename("plan_pct").reset_index()

    grades = docp.groupby(CELL + ["grade"], as_index=False).agg(
        base_fact=("base_revenue", "sum"), fact_pay=("gross", "sum"), fact_contrib=("contributions", "sum"))
    grades = grades.merge(plan_pct, on=["clinic_id", "direction_code"])
    grades["grade"] = pd.Categorical(grades.grade, GRADES, ordered=True)
    grades = grades.sort_values(CELL + ["grade"]).reset_index(drop=True)
    grades["grade"] = grades.grade.astype(str)

    cells = grades.groupby(CELL, as_index=False).plan_pct.first()

    data, _ = bpf.prepare()
    common = data[data.article.isin(["Оплата труда медперсонала", "Амортизация"])]
    nurse_c = pay[pay.role == "nurse"].groupby(["period", "clinic_id"], as_index=False).contributions.sum().rename(
        columns={"contributions": "nurse_fact_contrib"})
    g = common.pivot_table(index=["period", "clinic_id"], columns="article", values=["plan", "fact"], aggfunc="sum").fillna(0.0)
    g.columns = [f"{a}|{b}" for a, b in g.columns]
    g = g.reset_index().rename(columns={
        "plan|Оплата труда медперсонала": "nurse_plan", "fact|Оплата труда медперсонала": "nurse_fact",
        "plan|Амортизация": "depr_plan", "fact|Амортизация": "depr_fact"})
    g = g.merge(nurse_c, on=["period", "clinic_id"], how="left").fillna(0.0)
    g.insert(2, "direction_code", "ОБЩ")
    g = g[["period", "clinic_id", "direction_code", "nurse_plan", "nurse_fact", "nurse_fact_contrib", "depr_plan", "depr_fact"]]
    return dict(rate=rate, grades=grades, cells=cells, common=g)


def build_gross(wb, d, fot):
    """Листы: ФОТ врачи / ФОТ ячейки / ФОТ общие (данные с формулами), Валовая маржа, Врачи по грейдам."""
    nd = len(d) + 1
    Dd = lambda c: f"'Данные'!${c}$2:${c}${nd}"  # noqa: E731
    gr, ce, co, rate = fot["grades"], fot["cells"], fot["common"], fot["rate"]

    def fill(ws, hdr, rows, blue_from, blue_to, widths, nfmt_from=5):
        for i, h in enumerate(hdr, 1):
            ws.cell(1, i, h)
        head(ws, 1, 1, len(hdr))
        ws.row_dimensions[1].height = 42
        for r, row in enumerate(rows, 2):
            for c, v in enumerate(row, 1):
                cell = ws.cell(r, c, v)
                cell.font = BLUE if blue_from <= c <= blue_to else BLACK
                if c >= nfmt_from:
                    cell.number_format = "#,##0.00"
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[col(i)].width = w
        ws.freeze_panes = "E2"

    # --- ФОТ врачи: период x клиника x направление x грейд
    wg = wb.create_sheet("ФОТ врачи")
    n1 = len(gr) + 1
    rows = gr[["period", "clinic_id", "direction_code", "grade", "base_fact", "fact_pay", "fact_contrib", "plan_pct"]].values.tolist()
    fill(wg, ["Период", "Клиника", "Направление", "Грейд", "База выплат (выручка врачей), факт", "Оплата врачей, факт",
              "Взносы с оплаты врачей, факт", "Плановая доля направления", "Гибкий план оплаты",
              "Эффект: ставка и грейды", "Эффект на взносы"], rows, 5, 8,
         [9, 8, 11, 9, 16, 14, 14, 12, 14, 14, 14])
    for r in range(2, n1 + 1):
        wg.cell(r, 9, f"=E{r}*H{r}")
        wg.cell(r, 10, f"=I{r}-F{r}")
        wg.cell(r, 11, f"=J{r}*$Q$1")
        for c in (9, 10, 11):
            wg.cell(r, c).font = BLACK
            wg.cell(r, c).number_format = "#,##0.00"
        wg.cell(r, 8).number_format = "0.0000"
    wg["P1"], wg["P1"].font = "Ставка взносов", BOLD
    wg["Q1"], wg["Q1"].font, wg["Q1"].fill, wg["Q1"].number_format = rate, BLUE, INPUT_FILL, "0.0000"
    wg.auto_filter.ref = f"A1:K{n1}"
    G = lambda c: f"'ФОТ врачи'!${c}$2:${c}${n1}"  # noqa: E731

    # --- ФОТ ячейки: период x клиника x направление (план врачей и гибкий план)
    wc_ = wb.create_sheet("ФОТ ячейки")
    n2 = len(ce) + 1
    fill(wc_, ["Период", "Клиника", "Направление", "Плановая доля направления", "План выручки", "План оплаты врачей",
               "База выплат, факт", "Гибкий план оплаты", "Эффект: оплата следует за выручкой", "Эффект на взносы",
               "План взносов с врачей"],
         ce[["period", "clinic_id", "direction_code", "plan_pct"]].values.tolist(), 4, 4,
         [9, 8, 11, 12, 15, 15, 15, 15, 16, 14, 14])
    for r in range(2, n2 + 1):
        wc_.cell(r, 5, f"=SUMIFS({Dd('Q')},{Dd('A')},A{r},{Dd('B')},B{r},{Dd('C')},C{r})")
        wc_.cell(r, 6, f"=E{r}*D{r}")
        wc_.cell(r, 7, f"=SUMIFS({G('E')},{G('A')},A{r},{G('B')},B{r},{G('C')},C{r})")
        wc_.cell(r, 8, f"=G{r}*D{r}")
        wc_.cell(r, 9, f"=F{r}-H{r}")
        wc_.cell(r, 10, f"=I{r}*'ФОТ врачи'!$Q$1")
        wc_.cell(r, 11, f"=F{r}*'ФОТ врачи'!$Q$1")
        for c in range(5, 12):
            wc_.cell(r, c).font = BLACK
            wc_.cell(r, c).number_format = "#,##0.00"
        wc_.cell(r, 4).number_format = "0.0000"
    wc_.auto_filter.ref = f"A1:K{n2}"
    Cc = lambda c: f"'ФОТ ячейки'!${c}$2:${c}${n2}"  # noqa: E731

    # --- ФОТ общие: период x клиника (медперсонал, амортизация)
    wo = wb.create_sheet("ФОТ общие")
    n3 = len(co) + 1
    fill(wo, ["Период", "Клиника", "Направление", "Медперсонал, план", "Медперсонал, факт", "Взносы с медперсонала, факт",
              "Амортизация, план", "Амортизация, факт", "Эффект: медперсонал", "Эффект: амортизация",
              "Эффект на взносы", "План взносов с медперсонала"],
         co.values.tolist(), 4, 8, [9, 8, 12, 15, 15, 15, 15, 15, 14, 14, 14, 15])
    for r in range(2, n3 + 1):
        wo.cell(r, 9, f"=D{r}-E{r}")
        wo.cell(r, 10, f"=G{r}-H{r}")
        wo.cell(r, 11, f"=I{r}*'ФОТ врачи'!$Q$1")
        wo.cell(r, 12, f"=D{r}*'ФОТ врачи'!$Q$1")
        for c in range(9, 13):
            wo.cell(r, c).font = BLACK
            wo.cell(r, c).number_format = "#,##0.00"
    wo.auto_filter.ref = f"A1:L{n3}"
    O = lambda c: f"'ФОТ общие'!${c}$2:${c}${n3}"  # noqa: E731

    # ================================================================ Валовая маржа
    ws = wb.create_sheet("Валовая маржа")
    ws["A1"], ws["A1"].font = "Факторный анализ валовой маржи, 2025 (тыс. ₽)", TITLE
    ws["A2"], ws["A2"].font = "Клиника (фильтр):", BOLD
    ws["A3"], ws["A3"].font = "Направление (фильтр):", BOLD
    for ref, wild, vals in (("B2", "C2", CLINICS), ("B3", "C3", list(DIRS))):
        ws[ref], ws[ref].fill, ws[ref].font = "Все", INPUT_FILL, BLUE
        ws[wild] = f'=IF($B${ref[1]}="Все","*",$B${ref[1]})'
        ws[wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ref)
    ws["D2"], ws["D3"] = "← Все, C01, C02 или C03", "← Все, DENT, FAM, COSM, PLAS или LASER"
    ws["D2"].font = ws["D3"].font = NOTE
    ws["A4"] = ("Плюс = выгодно для маржи. Валовая маржа = выручка − материалы − лаборатории − врачи − медперсонал − взносы − амортизация. "
                "Медперсонал и амортизация заданы по клинике: при выборе направления они не распределяются (строки равны нулю).")
    ws["A4"].font = NOTE
    for i, h in enumerate(["Показатель", "Сумма", "% плановой маржи", "Что означает"], 1):
        ws.cell(6, i, h)
    head(ws, 6, 1, 8)
    ws["D6"].alignment = Alignment(horizontal="left", vertical="center")
    cr = "$C$2", "$C$3"
    sd = lambda rng_fn, c: f"SUMIFS({rng_fn(c)},{rng_fn('B')},{cr[0]},{rng_fn('C')},{cr[1]})"  # noqa: E731
    # строки
    ws.cell(7, 1, "Плановая прямая маржа").font = BOLD
    ws.cell(7, 2, "=" + sd(Dd, "V"))
    plan_rows = [("  − плановая оплата врачей", "=-" + sd(Cc, "F")),
                 ("  − плановая оплата медперсонала", "=-" + sd(O, "D")),
                 ("  − плановые взносы", f"=-({sd(Cc, 'K')}+{sd(O, 'L')})"),
                 ("  − плановая амортизация", "=-" + sd(O, "G"))]
    for i, (nm, f_) in enumerate(plan_rows, 8):
        ws.cell(i, 1, nm).font = BLACK
        ws.cell(i, 2, f_).font = GREEN_LINK
    ws.cell(12, 1, "Плановая валовая маржа").font = BOLD
    ws.cell(12, 2, "=SUM(B7:B11)").font = BOLD
    for c in range(1, 5):
        ws.cell(12, c).fill = SEC_FILL
    names = [(nm, "=" + sd(Dd, c)) for nm, c in FACTORS] + [
        ("Врачи: оплата следует за выручкой", "=" + sd(Cc, "I")),
        ("Врачи: ставка и грейды", "=" + sd(G, "J")),
        ("Медперсонал", "=" + sd(O, "I")),
        ("Взносы с ФОТ", f"={sd(Cc, 'J')}+{sd(G, 'K')}+{sd(O, 'K')}"),
        ("Амортизация", "=" + sd(O, "J"))]
    meaning2 = dict({k: v for k, v in {
        "Врачи: оплата следует за выручкой": "Оплата врачей зависит от выручки: недобор выручки снижает её (гибкий план); плюс = экономия",
        "Врачи: ставка и грейды": "Факт оплаты против гибкого плана: другая структура выручки по грейдам, чем в плане",
        "Медперсонал": "Фиксированные затраты клиники: план против факта",
        "Взносы с ФОТ": "Страховые взносы следуют за эффектами врачей и медперсонала",
        "Амортизация": "Фиксированные затраты клиники: план против факта"}.items()})
    r = 13
    for nm, f_ in names:
        ws.cell(r, 1, nm).font = BLACK
        ws.cell(r, 2, f_).font = GREEN_LINK
        ws.cell(r, 3, f"=IF($B$12=0,0,B{r}/$B$12)")
        ws.cell(r, 4, MEANING.get(nm, meaning2.get(nm, ""))).font = NOTE
        r += 1
    fr = r  # 24
    ws.cell(fr, 1, "Фактическая валовая маржа").font = BOLD
    ws.cell(fr, 2, f"=B12+SUM(B13:B{fr - 1})").font = BOLD
    ws.cell(fr, 3, f"=IF($B$12=0,0,B{fr}/$B$12-1)")
    ws.cell(fr + 1, 1, "Контроль: факт маржи по данным минус факт по мосту").font = NOTE
    fact_direct = sd(Dd, "AC")
    fact_costs = f"{sd(G, 'F')}+{sd(O, 'E')}+{sd(G, 'G')}+{sd(O, 'F')}+{sd(O, 'H')}"
    ws.cell(fr + 1, 2, f"=({fact_direct}-({fact_costs}))-B{fr}").font = NOTE
    ws.cell(fr + 1, 4, "≈ 0 (остаток в несколько рублей: взносы в 1С округлены по сотрудникам)").font = NOTE
    for rr in range(7, fr + 2):
        ws.cell(rr, 2).number_format = NUM if rr != fr + 1 else "#,##0.00"
        ws.cell(rr, 3).number_format = PCT
    for c in range(1, 5):
        ws.cell(fr, c).border = TOP
    sign_format(ws, f"B13:C{fr - 1}")
    ws.cell(fr, 3).font = BOLD

    q0 = fr + 3
    for i, h in enumerate(["", "План", "Факт", "Откл."], 1):
        ws.cell(q0, i, h)
    head(ws, q0, 1, 4)
    ws.cell(q0 + 1, 1, "Выручка")
    ws.cell(q0 + 1, 2, "=" + sd(Dd, "Q"))
    ws.cell(q0 + 1, 3, "=" + sd(Dd, "J"))
    ws.cell(q0 + 2, 1, "Валовая маржа")
    ws.cell(q0 + 2, 2, "=B12")
    ws.cell(q0 + 2, 3, f"=B{fr}")
    ws.cell(q0 + 3, 1, "Валовая маржа, % выручки")
    for c in (2, 3):
        L = col(c)
        ws.cell(q0 + 3, c, f"=IF({L}{q0 + 1}=0,0,{L}{q0 + 2}/{L}{q0 + 1})")
    ws.cell(q0 + 1, 4, f"=C{q0 + 1}-B{q0 + 1}")
    ws.cell(q0 + 2, 4, f"=C{q0 + 2}-B{q0 + 2}")
    ws.cell(q0 + 3, 4, f"=(C{q0 + 3}-B{q0 + 3})*100")
    for rr in (q0 + 1, q0 + 2):
        for c in (2, 3, 4):
            ws.cell(rr, c).number_format = NUM
    for c in (2, 3):
        ws.cell(q0 + 3, c).number_format = PCT
    ws.cell(q0 + 3, 4).number_format = PP
    for rr in range(q0 + 1, q0 + 4):
        for c in range(1, 5):
            ws.cell(rr, c).font = BLACK
    sign_format(ws, f"D{q0 + 1}:D{q0 + 3}")
    ws.column_dimensions["A"].width = 46
    for c_ in "BCDEFGH":
        ws.column_dimensions[c_].width = 14
    ch = BarChart()
    ch.type, ch.title = "bar", "Влияние факторов на валовую маржу, тыс. ₽"
    ch.add_data(Reference(ws, min_col=2, min_row=13, max_row=fr - 1), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=1, min_row=13, max_row=fr - 1))
    ch.legend = None
    ch.series[0].invertIfNegative = False
    ch.series[0].graphicalProperties.solidFill = "2E75B6"
    ch.y_axis.numFmt = "#,##0,"
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.tickLblPos = "low"
    ch.height, ch.width = 10.5, 22
    ws.add_chart(ch, f"A{q0 + 6}")

    # ================================================================ Врачи по грейдам
    wv = wb.create_sheet("Врачи по грейдам")
    wv["A1"], wv["A1"].font = "Оплата врачей: эффект ставки и грейдов (тыс. ₽)", TITLE
    wv["A2"], wv["A2"].font = "Клиника (фильтр):", BOLD
    wv["A3"], wv["A3"].font = "Направление (фильтр):", BOLD
    for ref, wild, vals in (("B2", "C2", CLINICS), ("B3", "C3", list(DIRS))):
        wv[ref], wv[ref].fill, wv[ref].font = "Все", INPUT_FILL, BLUE
        wv[wild] = f'=IF($B${ref[1]}="Все","*",$B${ref[1]})'
        wv[wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        wv.add_data_validation(dv)
        dv.add(ref)
    wv["A4"] = ("Плюс = выгодно. План строится по простому среднему pay_pct врачей направления, то есть предполагает равный раздел выручки; "
                "фактически сильные врачи берут больше выручки и стоят дороже, поэтому эффект грейдов отчасти отражает упрощение плана.")
    wv["A4"].font = NOTE
    wv["A4"].alignment = Alignment(wrap_text=True, vertical="top")
    wv.merge_cells("A4:H4")
    wv.row_dimensions[4].height = 38
    hh = ["Грейд", "База выплат, факт", "Доля базы", "Оплата врачей, факт", "Доля врача, факт", "Доля врача, гибкий план",
          "Эффект: ставка и грейды", "Эффект на взносы"]
    for i, h in enumerate(hh, 1):
        wv.cell(6, i, h)
    head(wv, 6, 1, 8)
    wv.row_dimensions[6].height = 32
    crit_g = lambda c, key: f"SUMIFS({G(c)},{G('D')},{key},{G('B')},$C$2,{G('C')},$C$3)"  # noqa: E731
    for k, g_ in enumerate(GRADES + ["*"]):
        rr = 7 + k
        tot = g_ == "*"
        wv.cell(rr, 1, "Итого" if tot else g_).font = BOLD if tot else BLACK
        wv.cell(rr, 10, g_).font = Font(name=F, size=8, color="808080")
        wv.cell(rr, 2, "=" + crit_g("E", f"$J{rr}"))
        wv.cell(rr, 3, f"=IF($B$11=0,0,B{rr}/$B$11)")
        wv.cell(rr, 4, "=" + crit_g("F", f"$J{rr}"))
        wv.cell(rr, 5, f"=IF(B{rr}=0,0,D{rr}/B{rr})")
        wv.cell(rr, 6, f"=IF(B{rr}=0,0,{crit_g('I', f'$J{rr}')}/B{rr})")
        wv.cell(rr, 7, "=" + crit_g("J", f"$J{rr}"))
        wv.cell(rr, 8, "=" + crit_g("K", f"$J{rr}"))
        for c, fmt in ((2, NUM), (3, PCT), (4, NUM), (5, "0.0%"), (6, "0.0%"), (7, NUM), (8, NUM)):
            wv.cell(rr, c).number_format = fmt
            wv.cell(rr, c).font = BOLD if tot else BLACK
        if tot:
            for c in range(1, 9):
                wv.cell(rr, c).border = TOP
    sign_format(wv, "G7:H11")

    wv.cell(14, 1, "Эффект ставки и грейдов по направлениям и грейдам (тыс. ₽, фильтр клиники сверху)").font = BOLD
    wv.cell(15, 1, "Направление")
    for i, g_ in enumerate(GRADES + ["Итого"]):
        wv.cell(15, 2 + i, g_)
    head(wv, 15, 1, 6)
    for k, (dc, nm) in enumerate(list(DIRS.items()) + [("*", "Итого")]):
        rr = 16 + k
        wv.cell(rr, 1, nm).font = BOLD if dc == "*" else BLACK
        wv.cell(rr, 10, dc).font = Font(name=F, size=8, color="808080")
        for i, g_ in enumerate(GRADES + ["*"]):
            wv.cell(rr, 2 + i, f"=SUMIFS({G('J')},{G('C')},$J{rr},{G('D')},{'\"' + g_ + '\"'},{G('B')},$C$2)")
            wv.cell(rr, 2 + i).number_format = NUM
            wv.cell(rr, 2 + i).font = BOLD if (dc == "*" or g_ == "*") else BLACK
        if dc == "*":
            for c in range(1, 7):
                wv.cell(rr, c).border = TOP
    sign_format(wv, "B16:F21")
    wv.column_dimensions["A"].width = 28
    for c_ in "BCDEFGH":
        wv.column_dimensions[c_].width = 15
    wv.column_dimensions["J"].width = 6
    return wb


def build(d, fot):
    e = effects(d)
    wb = Workbook()
    n = len(d)
    nd = n + 1

    # ---------------------------------------------------------------- Данные
    wd = wb.active
    wd.title = "Данные"
    hdr = ["Период", "Клиника", "Направление", "Услуга", "Название", "Кол-во, план", "Кол-во, факт",
           "Цена, план", "Выручка факт по прайсу", "Выручка, факт", "Материалы ед., план", "Лаборатории ед., план",
           "Материалы, факт", "Лаборатории, факт",
           "Ключ ячейки", "Факт кол-во × план цена", "План выручка", "Маржа ед., план", "Маржа, план",
           "Факт кол-во × маржа ед. план", "Индекс объёма ячейки",
           "Маржа, план (ф)", "Объём", "Структура услуг", "Цена: прайс", "Цена: скидки", "Материалы", "Лаборатории",
           "Маржа, факт", "Проверка"]
    # буквы: A..N вход, O ключ, P..T помощники, U индекс, V план маржа, W..AB факторы, AC факт, AD проверка
    for i, h in enumerate(hdr, 1):
        wd.cell(1, i, h)
    head(wd, 1, 1, len(hdr))
    wd.row_dimensions[1].height = 42
    inp = d.values.tolist()
    for r, row in enumerate(inp, 2):
        for c, v in enumerate(row, 1):
            cell = wd.cell(r, c, v)
            cell.font = BLUE if c >= 6 else BLACK
        wd.cell(r, 15, f'=A{r}&"|"&B{r}&"|"&C{r}')
        wd.cell(r, 16, f"=G{r}*H{r}")
        wd.cell(r, 17, f"=F{r}*H{r}")
        wd.cell(r, 18, f"=H{r}-K{r}-L{r}")
        wd.cell(r, 19, f"=F{r}*R{r}")
        wd.cell(r, 20, f"=G{r}*R{r}")
        wd.cell(r, 21, f"=SUMIFS($P$2:$P${nd},$O$2:$O${nd},O{r})/SUMIFS($Q$2:$Q${nd},$O$2:$O${nd},O{r})")
        wd.cell(r, 22, f"=S{r}")
        wd.cell(r, 23, f"=S{r}*(U{r}-1)")
        wd.cell(r, 24, f"=T{r}-U{r}*S{r}")
        wd.cell(r, 25, f"=I{r}-P{r}")
        wd.cell(r, 26, f"=J{r}-I{r}")
        wd.cell(r, 27, f"=-(M{r}-G{r}*K{r})")
        wd.cell(r, 28, f"=-(N{r}-G{r}*L{r})")
        wd.cell(r, 29, f"=J{r}-M{r}-N{r}")
        wd.cell(r, 30, f"=AC{r}-(V{r}+W{r}+X{r}+Y{r}+Z{r}+AA{r}+AB{r})")
        for c in range(15, 31):
            wd.cell(r, c).font = BLACK
        for c in list(range(6, 15)) + list(range(16, 21)) + list(range(22, 31)):
            wd.cell(r, c).number_format = "#,##0.00"
        wd.cell(r, 21).number_format = "0.0000"
    widths = [9, 8, 11, 8, 36, 10, 10, 10, 14, 14, 11, 11, 13, 13, 22, 14, 14, 11, 14, 14, 11,
              14, 13, 13, 13, 13, 13, 13, 14, 10]
    for i, w in enumerate(widths, 1):
        wd.column_dimensions[col(i)].width = w
    wd.freeze_panes = "F2"
    wd.auto_filter.ref = f"A1:AD{nd}"
    D = lambda c: f"'Данные'!${c}$2:${c}${nd}"  # noqa: E731

    periods = [f"{YEAR}-{m:02d}" for m in range(1, 13)]

    def crit(clinic="$C$2", direction="$C$3"):
        parts = []
        if clinic:
            parts.append(f"{D('B')},{clinic}")
        if direction:
            parts.append(f"{D('C')},{direction}")
        return ",".join(parts)

    # ---------------------------------------------------------------- Мост
    ws = wb.create_sheet("Мост", 0)
    ws["A1"], ws["A1"].font = "Факторный анализ прямой маржи, 2025 (тыс. ₽)", TITLE
    ws["A2"], ws["A2"].font = "Клиника (фильтр):", BOLD
    ws["A3"], ws["A3"].font = "Направление (фильтр):", BOLD
    for ref, wild, vals in (("B2", "C2", CLINICS), ("B3", "C3", list(DIRS))):
        ws[ref], ws[ref].fill, ws[ref].font = "Все", INPUT_FILL, BLUE
        ws[wild] = f'=IF($B${ref[1]}="Все","*",$B${ref[1]})'
        ws[wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ref)
    ws["D2"], ws["D3"] = "← Все, C01, C02 или C03", "← Все, DENT, FAM, COSM, PLAS или LASER"
    ws["D2"].font = ws["D3"].font = NOTE
    ws["A4"] = "Плюс = выгодно для маржи, минус = невыгодно. Прямая маржа = выручка − материалы − лаборатории."
    ws["A4"].font = NOTE
    for i, h in enumerate(["Показатель", "Сумма", "% плановой маржи", "Что означает"], 1):
        ws.cell(6, i, h)
    head(ws, 6, 1, 8)
    ws["D6"].alignment = Alignment(horizontal="left", vertical="center")
    S = lambda c: f"=SUMIFS({D(c)},{crit()})"  # noqa: E731
    meaning = MEANING
    ws.cell(7, 1, "Плановая прямая маржа").font = BOLD
    ws.cell(7, 2, S("V")).font = BOLD
    r = 8
    for name, c in FACTORS:
        ws.cell(r, 1, name).font = BLACK
        ws.cell(r, 2, S(c)).font = GREEN_LINK
        ws.cell(r, 3, f"=IF($B$7=0,0,B{r}/$B$7)")
        ws.cell(r, 4, meaning[name]).font = NOTE
        r += 1
    fr = r  # строка факта
    ws.cell(fr, 1, "Фактическая прямая маржа").font = BOLD
    ws.cell(fr, 2, f"=B7+SUM(B8:B{fr - 1})").font = BOLD
    ws.cell(fr, 3, f"=IF($B$7=0,0,B{fr}/$B$7-1)")
    ws.cell(fr + 1, 1, "Контроль: факт маржи по данным минус факт по мосту").font = NOTE
    ws.cell(fr + 1, 2, f"=SUMIFS({D('AC')},{crit()})-B{fr}").font = NOTE
    ws.cell(fr + 1, 4, "должно быть 0").font = NOTE
    for rr in range(7, fr + 2):
        ws.cell(rr, 2).number_format = NUM if rr != fr + 1 else "#,##0.00"
        ws.cell(rr, 3).number_format = PCT
    for c in range(1, 5):
        ws.cell(fr, c).border = TOP
        ws.cell(7, c).fill = SEC_FILL
    sign_format(ws, f"B8:C{fr - 1}")
    ws.cell(fr, 3).font = BOLD

    q0 = fr + 3
    for i, h in enumerate(["", "План", "Факт", "Откл."], 1):
        ws.cell(q0, i, h)
    head(ws, q0, 1, 4)
    ws.cell(q0 + 1, 1, "Выручка")
    ws.cell(q0 + 1, 2, f"=SUMIFS({D('Q')},{crit()})")
    ws.cell(q0 + 1, 3, f"=SUMIFS({D('J')},{crit()})")
    ws.cell(q0 + 2, 1, "Прямая маржа")
    ws.cell(q0 + 2, 2, "=B7")
    ws.cell(q0 + 2, 3, f"=B{fr}")
    ws.cell(q0 + 3, 1, "Прямая маржа, % выручки")
    for c in (2, 3):
        L = col(c)
        ws.cell(q0 + 3, c, f"=IF({L}{q0 + 1}=0,0,{L}{q0 + 2}/{L}{q0 + 1})")
    ws.cell(q0 + 1, 4, f"=C{q0 + 1}-B{q0 + 1}")
    ws.cell(q0 + 2, 4, f"=C{q0 + 2}-B{q0 + 2}")
    ws.cell(q0 + 3, 4, f"=(C{q0 + 3}-B{q0 + 3})*100")
    for rr in (q0 + 1, q0 + 2):
        for c in (2, 3, 4):
            ws.cell(rr, c).number_format = NUM
    for c in (2, 3):
        ws.cell(q0 + 3, c).number_format = PCT
    ws.cell(q0 + 3, 4).number_format = PP
    for rr in range(q0 + 1, q0 + 4):
        for c in range(1, 5):
            ws.cell(rr, c).font = BLACK
    sign_format(ws, f"D{q0 + 1}:D{q0 + 3}")
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    for c_ in "EFGH":
        ws.column_dimensions[c_].width = 14

    ch = BarChart()
    ch.type, ch.title = "bar", "Влияние факторов на прямую маржу, тыс. ₽"
    ch.add_data(Reference(ws, min_col=2, min_row=8, max_row=fr - 1), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=1, min_row=8, max_row=fr - 1))
    ch.legend = None
    ch.series[0].invertIfNegative = False
    ch.series[0].graphicalProperties.solidFill = "2E75B6"
    ch.y_axis.numFmt = "#,##0,"
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.tickLblPos = "low"
    ch.height, ch.width = 8.5, 22
    ws.add_chart(ch, f"A{q0 + 6}")

    # ---------------------------------------------------------------- По месяцам
    wm = wb.create_sheet("По месяцам", 1)
    wm["A1"], wm["A1"].font = "Факторы по месяцам (тыс. ₽)", TITLE
    wm["A2"], wm["A2"].font = "Клиника (фильтр):", BOLD
    wm["A3"], wm["A3"].font = "Направление (фильтр):", BOLD
    for ref, wild, vals in (("B2", "C2", CLINICS), ("B3", "C3", list(DIRS))):
        wm[ref], wm[ref].fill, wm[ref].font = "Все", INPUT_FILL, BLUE
        wm[wild] = f'=IF($B${ref[1]}="Все","*",$B${ref[1]})'
        wm[wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        wm.add_data_validation(dv)
        dv.add(ref)
    wm["A4"] = "Объём в мае по стоматологии и в июле-сентябре по лазеру клиники C03: выбери направление и клинику."
    wm["A4"].font = NOTE
    wm.cell(6, 1, "Фактор")
    for i, p in enumerate(periods):
        wm.cell(6, 2 + i, p)
    wm.cell(6, 14, f"Итого {YEAR}")
    head(wm, 6, 1, 14)
    rr = 7
    wm.cell(rr, 1, "Плановая прямая маржа").font = BOLD
    for i in range(12):
        L = col(2 + i)
        wm.cell(rr, 2 + i, f"=SUMIFS({D('V')},{D('A')},{L}$6,{crit()})").font = BOLD
    wm.cell(rr, 14, "=SUM(B7:M7)").font = BOLD
    rr += 1
    for name, c in FACTORS:
        wm.cell(rr, 1, name).font = BLACK
        for i in range(12):
            L = col(2 + i)
            wm.cell(rr, 2 + i, f"=SUMIFS({D(c)},{D('A')},{L}$6,{crit()})")
        wm.cell(rr, 14, f"=SUM(B{rr}:M{rr})")
        rr += 1
    wm.cell(rr, 1, "Фактическая прямая маржа").font = BOLD
    for c in range(2, 15):
        L = col(c)
        wm.cell(rr, c, f"={L}7+SUM({L}8:{L}{rr - 1})").font = BOLD
        wm.cell(rr, c).border = TOP
    wm.cell(rr, 1).border = TOP
    for r2 in range(7, rr + 1):
        for c in range(2, 15):
            wm.cell(r2, c).number_format = NUM
            if wm.cell(r2, c).font.bold is not True:
                wm.cell(r2, c).font = BLACK
    sign_format(wm, f"B8:N{rr - 1}")
    wm.cell(rr + 1, 1, "Объём к плану, % (по выручке в плановых ценах)").font = BOLD
    for i in range(12):
        L = col(2 + i)
        wm.cell(rr + 1, 2 + i, f"=IF(SUMIFS({D('Q')},{D('A')},{L}$6,{crit()})=0,0,"
                               f"SUMIFS({D('P')},{D('A')},{L}$6,{crit()})/SUMIFS({D('Q')},{D('A')},{L}$6,{crit()})-1)")
        wm.cell(rr + 1, 2 + i).number_format = PCT
    wm.cell(rr + 1, 14, f"=IF(SUMIFS({D('Q')},{crit()})=0,0,SUMIFS({D('P')},{crit()})/SUMIFS({D('Q')},{crit()})-1)")
    wm.cell(rr + 1, 14).number_format = PCT
    from openpyxl.formatting.rule import ColorScaleRule
    wm.conditional_formatting.add(f"B{rr + 1}:N{rr + 1}", ColorScaleRule(
        start_type="num", start_value=-0.15, start_color="F8696B", mid_type="num", mid_value=0,
        mid_color="FFFFFF", end_type="num", end_value=0.05, end_color="63BE7B"))
    wm.column_dimensions["A"].width = 44
    for c in range(2, 15):
        wm.column_dimensions[col(c)].width = 11.5
    wm.freeze_panes = "B7"

    # ---------------------------------------------------------------- Срезы
    wc = wb.create_sheet("Срезы", 2)
    wc["A1"], wc["A1"].font = "Факторы по направлениям и клиникам, 2025 (тыс. ₽)", TITLE
    cols_h = ["План маржи", "Объём", "Структура услуг", "Цена: прайс", "Цена: скидки", "Материалы",
              "Лаборатории", "Факт маржи", "Откл. маржи", "Маржа план, %", "Маржа факт, %", "Откл., п.п."]
    srcs = ["V", "W", "X", "Y", "Z", "AA", "AB", "AC"]

    def slice_table(top, title, key_col, keys, names, extra_crit, filter_label, filter_ref, filter_wild, vals, filter_col):
        wc.cell(top, 1, title).font = BOLD
        wc.cell(top + 1, 1, filter_label).font = BOLD
        wc[filter_ref], wc[filter_ref].fill, wc[filter_ref].font = "Все", INPUT_FILL, BLUE
        wc[filter_wild] = f'=IF(${filter_ref[0]}${filter_ref[1:]}="Все","*",${filter_ref[0]}${filter_ref[1:]})'
        wc[filter_wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        wc.add_data_validation(dv)
        dv.add(filter_ref)
        h = top + 3
        wc.cell(h, 1, "Срез")
        for i, t in enumerate(cols_h):
            wc.cell(h, 2 + i, t)
        head(wc, h, 1, 13)
        wc.row_dimensions[h].height = 30
        r0 = h + 1
        for k, nm in zip(keys + ["*"], names + ["Итого"]):
            rr = h + 1 + (keys + ["*"]).index(k)
            bold = k == "*"
            wc.cell(rr, 1, nm).font = BOLD if bold else BLACK
            wc.cell(rr, 15, k).font = Font(name=F, size=8, color="808080")
            for i, s in enumerate(srcs):
                wc.cell(rr, 2 + i, f"=SUMIFS({D(s)},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]})")
                wc.cell(rr, 2 + i).number_format = NUM
                wc.cell(rr, 2 + i).font = BOLD if bold else BLACK
            wc.cell(rr, 10, f"=I{rr}-B{rr}")
            wc.cell(rr, 11, f"=IF(SUMIFS({D('Q')},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]})=0,0,B{rr}/SUMIFS({D('Q')},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]}))")
            wc.cell(rr, 12, f"=IF(SUMIFS({D('J')},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]})=0,0,I{rr}/SUMIFS({D('J')},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]}))")
            wc.cell(rr, 13, f"=(L{rr}-K{rr})*100")
            wc.cell(rr, 10).number_format = NUM
            wc.cell(rr, 11).number_format = wc.cell(rr, 12).number_format = PCT
            wc.cell(rr, 13).number_format = PP
            for c in (10, 11, 12, 13):
                wc.cell(rr, c).font = BOLD if bold else BLACK
            if bold:
                for c in range(1, 14):
                    wc.cell(rr, c).border = TOP
        last = h + len(keys) + 1
        sign_format(wc, f"C{h + 1}:H{last}")
        sign_format(wc, f"J{h + 1}:J{last}")
        sign_format(wc, f"M{h + 1}:M{last}")
        return last

    last = slice_table(2, "По направлениям", "C", list(DIRS), list(DIRS.values()), None,
                       "Клиника (фильтр):", "B3", "C3", CLINICS, "B")
    slice_table(last + 3, "По клиникам", "B", CLINICS, ["Невский", "Север", "Юг"], None,
                "Направление (фильтр):", f"B{last + 4}", f"C{last + 4}", list(DIRS), "C")
    wc.column_dimensions["A"].width = 30
    for c in range(2, 14):
        wc.column_dimensions[col(c)].width = 13
    wc.column_dimensions["O"].width = 6
    wc.freeze_panes = "B5"

    # ---------------------------------------------------------------- По услугам
    wsv = wb.create_sheet("По услугам", 3)
    wsv["A1"], wsv["A1"].font = "Факторы по услугам, 2025 (тыс. ₽)", TITLE
    wsv["A2"], wsv["A2"].font = "Клиника (фильтр):", BOLD
    wsv["B2"], wsv["B2"].fill, wsv["B2"].font = "Все", INPUT_FILL, BLUE
    wsv["C2"] = '=IF($B$2="Все","*",$B$2)'
    wsv["C2"].font = Font(name=F, size=8, color="808080")
    dv = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    wsv.add_data_validation(dv)
    dv.add("B2")
    svc = d[["service_id", "service_name", "direction_code"]].drop_duplicates().sort_values("service_id")
    for i, h in enumerate(["Услуга", "Название", "Напр."] + cols_h[:8], 1):
        wsv.cell(4, i, h)
    head(wsv, 4, 1, 11)
    wsv.row_dimensions[4].height = 30
    for j, s in enumerate(svc.itertuples(), 5):
        wsv.cell(j, 1, s.service_id).font = BLACK
        wsv.cell(j, 2, s.service_name).font = BLACK
        wsv.cell(j, 3, s.direction_code).font = BLACK
        for i, src in enumerate(srcs):
            wsv.cell(j, 4 + i, f"=SUMIFS({D(src)},{D('D')},$A{j},{D('B')},$C$2)")
            wsv.cell(j, 4 + i).number_format = NUM
            wsv.cell(j, 4 + i).font = BLACK
    sign_format(wsv, f"E5:J{4 + len(svc)}")
    wsv.column_dimensions["A"].width = 9
    wsv.column_dimensions["B"].width = 42
    wsv.column_dimensions["C"].width = 8
    for c in range(4, 12):
        wsv.column_dimensions[col(c)].width = 13
    wsv.freeze_panes = "D5"

    # ---------------------------------------------------------------- Топ факторов (значения)
    wt = wb.create_sheet("Топ факторов", 4)
    wt["A1"], wt["A1"].font = "Где сильнее всего сработали факторы (тыс. ₽)", TITLE
    wt["A2"] = "Значения рассчитаны скриптом по всей сети (не формулы). Фильтры живых листов: «Мост», «По месяцам», «Срезы»."
    wt["A2"].font = NOTE
    cell_e = e.groupby(CELL, as_index=False).agg(
        plan_margin=("plan_margin", "sum"), volume=("volume", "sum"), mix=("mix", "sum"),
        price_list=("price_list", "sum"), discount=("discount", "sum"), materials=("materials", "sum"),
        labs=("labs", "sum"), q_f_pp=("q_f_pp", "sum"), q_p_pp=("q_p_pp", "sum"))
    cell_e["vol_pct"] = cell_e.q_f_pp / cell_e.q_p_pp - 1

    def top_table(top, title, frame, extra=None):
        wt.cell(top, 1, title).font = BOLD
        hh = ["Период", "Клиника", "Направление", "Эффект", "Объём к плану, %"]
        for i, h in enumerate(hh, 1):
            wt.cell(top + 1, i, h)
        head(wt, top + 1, 1, len(hh))
        rr = top + 2
        for x in frame.itertuples():
            vals = [x.period, x.clinic_id, DIRS[x.direction_code], x.eff, x.vol_pct]
            for i, v in enumerate(vals, 1):
                cc = wt.cell(rr, i, v)
                cc.font = BLACK
                cc.number_format = NUM if i == 4 else PCT if i == 5 else "General"
            rr += 1
        return rr + 1

    nxt = top_table(4, "Объём: 10 самых слабых ячеек (месяц × клиника × направление)",
                    cell_e.assign(eff=cell_e.volume).sort_values("eff").head(10))
    nxt = top_table(nxt, "Цена, прайс: 10 самых невыгодных ячеек",
                    cell_e.assign(eff=cell_e.price_list).sort_values("eff").head(10))
    nxt = top_table(nxt, "Скидки: 10 самых больших ячеек",
                    cell_e.assign(eff=cell_e.discount).sort_values("eff").head(10))
    nxt = top_table(nxt, "Материалы: 10 самых невыгодных ячеек",
                    cell_e.assign(eff=cell_e.materials).sort_values("eff").head(10))
    for i, w in enumerate([10, 10, 30, 14, 16], 1):
        wt.column_dimensions[col(i)].width = w

    build_gross(wb, d, fot)

    # ---------------------------------------------------------------- Описание
    wi = wb.create_sheet("Описание", 0)
    wi.column_dimensions["A"].width = 125
    text = [
        ("Факторный анализ прямой маржи, 2025", TITLE),
        ("Что раскладываем", BOLD),
        ("Лист «Мост»: отклонение факта от плана по прямой марже (выручка − материалы − лаборатории). Именно эти статьи заданы "
         "в плане на уровне услуг: количество, цена и себестоимость единицы. Лист «Валовая маржа» достраивает мост статьями, "
         "которые в плане заданы по клиникам и направлениям: врачи, медперсонал, взносы, амортизация.", BLACK),
        ("Факторы (разложение точное, остатка нет; контроль на листе «Мост»)", BOLD),
        ("Объём: плановая маржа ячейки × (индекс объёма − 1). Индекс объёма считается в плановых ценах по ячейке "
         "«месяц × клиника × направление», поэтому дешёвые и дорогие услуги не складываются штуками.", BLACK),
        ("Структура услуг: сдвиг состава услуг внутри направления к более или менее маржинальным.", BLACK),
        ("Цена: прайс: разница между действовавшим прайсом и плановой ценой (план мог заложить другую индексацию).", BLACK),
        ("Цена: скидки: разница между фактической выручкой и выручкой по действовавшему прайсу.", BLACK),
        ("Материалы и лаборатории: фактическая стоимость единицы против плановой, умноженная на фактическое количество.", BLACK),
        ("Достройка до валовой маржи", BOLD),
        ("Врачи: оплата следует за выручкой: оплата врачей это процент от выручки, поэтому при недоборе выручки она снижается сама. "
         "Это экономия за счёт объёма, а не эффективность затрат: плановая оплата минус гибкий план (фактическая база выплат × плановая доля направления).", BLACK),
        ("Врачи: ставка и грейды: факт оплаты против гибкого плана. Раскладывается по грейдам на листе «Врачи по грейдам».", BLACK),
        ("Медперсонал и амортизация фиксированы по клинике, по направлениям не распределяются. Взносы: ставка × эффекты врачей и медперсонала.", BLACK),
        ("Источники", BOLD),
        ("Факт по услугам: visit_items и visits после исправлений сверки. План: budget_revenue. "
         "Затраты факта: materials_1c (1С хранит их по направлениям); на услуги разнесены пропорционально нормативу "
         "(количество × норма единицы), поэтому суммы совпадают с 1С.", BLACK),
        ("Ограничения", BOLD),
        ("1. Структура направлений и клиник отдельным фактором не выделена: она входит в объём, потому что объём "
         "считается по каждой ячейке «месяц × клиника × направление» и суммируется.", BLACK),
        ("2. Фактор «Материалы/Лаборатории» показывает цену единицы. Расход на единицу (норматив) в данных не меняется, "
         "поэтому перерасход нормы от роста цен отделить нельзя.", BLACK),
        ("3. Эффект объёма смешивает фоновый недобор плана и разовые события. Разовые события видны на листах "
         "«По месяцам» и «Топ факторов».", BLACK),
        ("4. Плановая доля врачей направления это простое среднее pay_pct врачей клиники и направления (так построен бюджет): "
         "план предполагает равный раздел выручки между врачами. Сильные врачи берут больше выручки и стоят дороже, поэтому эффект "
         "грейдов отчасти отражает упрощение плана, а не операционную проблему.", BLACK),
        ("5. Налог на прибыль и НДС не моделируются. Эквайринг и маркетинг лежат ниже валовой маржи и здесь не раскладываются.", BLACK),
        ("Цвета: синий шрифт: входные числа, зелёный: ссылка на лист «Данные», чёрный: расчёты, жёлтая заливка: фильтры.", NOTE),
    ]
    for i, (t, f) in enumerate(text, 1):
        c = wi.cell(i, 1, t)
        c.font, c.alignment = f, Alignment(wrap_text=True, vertical="top")
    order = ["Описание", "Мост", "Валовая маржа", "По месяцам", "Срезы", "По услугам", "Врачи по грейдам",
             "Топ факторов", "Данные", "ФОТ врачи", "ФОТ ячейки", "ФОТ общие"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    return wb, e


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    d, mat1c = prepare()
    fot = prepare_fot()
    wb, e = build(d, fot)
    path = OUT / "FactorAnalysis_2025.xlsx"
    wb.save(path)
    recalculate(path)

    # контроль в pandas
    t = e[["plan_margin", "volume", "mix", "price_list", "discount", "materials", "labs", "fact_margin"]].sum()
    print(f"Сохранено: {path}  (строк: {len(d)})")
    print(f"Макс. остаток разложения по строкам: {e.check.abs().max():.6f}")
    print(f"Затраты факта: материалы {d.mat_fact.sum():,.0f} (1С {mat1c.material_fact.sum():,.0f}); "
          f"лаборатории {d.ext_fact.sum():,.0f} (1С {mat1c.external_fact.sum():,.0f})")
    print(f"Плановая прямая маржа {t.plan_margin / 1e6:,.2f} млн ₽ -> факт {t.fact_margin / 1e6:,.2f} млн ₽ "
          f"(откл. {(t.fact_margin - t.plan_margin) / 1e6:,.2f})")
    for k, nm in (("volume", "Объём"), ("mix", "Структура услуг"), ("price_list", "Цена: прайс"),
                  ("discount", "Цена: скидки"), ("materials", "Материалы"), ("labs", "Лаборатории")):
        print(f"  {nm:<18}{t[k] / 1e6:>8.2f} млн ₽")
    print(f"Ставка взносов: {fot['rate']:.4f}; строк ФОТ: врачи {len(fot['grades'])}, ячейки {len(fot['cells'])}, общие {len(fot['common'])}")


if __name__ == "__main__":
    main()
