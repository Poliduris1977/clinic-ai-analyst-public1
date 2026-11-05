"""
Шаг 13: LTV пациентов по когортам первого визита, только наблюдённая часть (без прогноза), в Excel.

Вход:  data/corrected/ - patients, visits, visit_items
Выход: data/reports/PatientLTV_2025.xlsx

Определения:
  новый пациент  первый визит (first_visit_date) в 2025 году и есть завершённые визиты;
                 пациенты с историей до 2025 года исключены: их визитов до 2025 года в данных нет
  когорта        месяц первого завершённого визита (1 = январь ... 12 = декабрь)
  возраст k      число календарных месяцев после месяца первого визита (0 = месяц первого визита)
  LTV до k       накопленная выручка на пациента за возрасты 0..k. LTV12 это возраст 11 (12 календарных месяцев)
  наблюдаемость  данные заканчиваются 31.12.2025: когорта c наблюдается до возраста 12 - c
                 (январская до 11, декабрьская только 0); ненаблюдаемые значения не показываются и не прогнозируются

Листы:
  Описание  - определения, правила, ограничения, контроль
  Когорты   - треугольник накопленного LTV по когортам (фильтры: направление, клиника, канал)
  Срезы     - LTV до возраста k по направлению, клинике и каналу (k задаёт пользователь), только когорты, где k наблюдается
  Данные    - когорта x направление x клиника x канал, накопленные суммы по возрасту

Запуск:  python src/reports/build_ltv.py
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.recalc import recalculate  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
YEAR = 2025
AGES = 12
DIRS = {"DENT": "Стоматология", "FAM": "Семейная медицина", "COSM": "Косметология",
        "PLAS": "Пластическая хирургия", "LASER": "Лазерная коррекция зрения"}
CLINICS = ["C01", "C02", "C03"]
MONTHS = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
MIN_N = 100          # меньше пациентов в срезе: пометка «мало пациентов»

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
TOP = Border(top=Side(style="thin"))
NUM1 = "#,##0.0,;(#,##0.0,);-"       # ₽ -> тыс. ₽ с одним знаком
INT = "#,##0;(#,##0);-"


# ------------------------------------------------------------------ данные
def prepare():
    pat = pd.read_csv(CORR / "patients.csv", parse_dates=["first_visit_date"])
    visits = pd.read_csv(CORR / "visits.csv")
    items = pd.read_csv(CORR / "visit_items.csv")
    vc = visits[visits.status == "completed"].merge(
        items.groupby("visit_id", as_index=False).revenue.sum(), on="visit_id")
    vc["d"] = pd.to_datetime(vc.visit_date)
    vc = vc.merge(pat[["patient_id", "first_visit_date", "source"]], on="patient_id")
    new = vc[vc.first_visit_date >= f"{YEAR}-01-01"].copy()
    new = new.sort_values(["patient_id", "d", "visit_id"])
    first = new.groupby("patient_id").first()[["d", "direction_code", "clinic_id", "source"]]
    first["cohort"] = first.d.dt.month
    new = new.merge(first[["cohort", "d"]].rename(columns={"d": "d0"}), left_on="patient_id", right_index=True)
    new["age"] = (new.d.dt.year - new.d0.dt.year) * 12 + new.d.dt.month - new.d0.dt.month
    assert new.age.between(0, AGES - 1).all()
    rev = new.pivot_table(index="patient_id", columns="age", values="revenue", aggfunc="sum", fill_value=0) \
             .astype(float).reindex(columns=range(AGES), fill_value=0.0)
    vis = new.pivot_table(index="patient_id", columns="age", values="visit_id", aggfunc="count", fill_value=0) \
             .reindex(columns=range(AGES), fill_value=0)
    cum_rev = rev.cumsum(axis=1)
    cum_vis = vis.cumsum(axis=1)
    ret = (cum_vis >= 2).astype(int)
    pt = first.join(cum_rev.add_prefix("r")).join(cum_vis.add_prefix("v")).join(ret.add_prefix("t"))
    pt["rev_total"] = cum_rev.iloc[:, -1]
    keys = ["cohort", "direction_code", "clinic_id", "source"]
    agg = {"d": "size"}
    for p in ("r", "v", "t"):
        for k in range(AGES):
            agg[f"{p}{k}"] = "sum"
    agg["rev_total"] = "sum"
    g = pt.groupby(keys).agg(agg).rename(columns={"d": "patients"}).reset_index()
    # ненаблюдаемые возрасты (k > 12 - cohort) обнуляем: на листах они исключаются отбором по когорте
    for k in range(AGES):
        mask = g.cohort > AGES - k
        for p in ("r", "v", "t"):
            g.loc[mask, f"{p}{k}"] = 0
    g = g.sort_values(keys).reset_index(drop=True)
    ctrl = {
        "patients": float(len(pt)),
        "rev_new": float(pt.rev_total.sum()),
        "ltv12_jan": float(pt[pt.cohort == 1].r11.mean()),
        "n_jan": int((pt.cohort == 1).sum()),
        "rev_total": float(items.revenue.sum()),
        "rev_old": float(items.revenue.sum() - pt.rev_total.sum()),
    }
    return g, pt, ctrl


# ------------------------------------------------------------------ оформление
def head(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put_filter(ws, row, label, value, options, hint, crit=None):
    ws.cell(row, 1, label).font = BOLD
    c = ws.cell(row, 2, value)
    c.fill, c.font = INPUT_FILL, BLUE
    if crit:
        ws.cell(row, 3, crit).font = GREY
    ws.cell(row, 4, hint).font = HINT
    if options:
        dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(c)


def build(g, ctrl, sources):
    wb = Workbook()

    # ---------- Данные ----------
    wd = wb.active
    wd.title = "Данные"
    heads = ["Когорта", "Месяц когорты", "Направление", "Клиника", "Канал", "Пациентов"]
    heads += [f"Накоплено выручки, возраст {k}" for k in range(AGES)]
    heads += [f"Накоплено визитов, возраст {k}" for k in range(AGES)]
    heads += [f"Вернулись (2+ визита), возраст {k}" for k in range(AGES)]
    heads += ["Выручка за весь наблюдаемый период"]
    for i, h in enumerate(heads, 1):
        wd.cell(1, i, h)
    head(wd, 1, 1, len(heads))
    wd.row_dimensions[1].height = 54
    for r, row in enumerate(g.itertuples(index=False), 2):
        vals = [row.cohort, MONTHS[row.cohort - 1], row.direction_code, row.clinic_id, row.source, row.patients]
        vals += [getattr(row, f"r{k}") for k in range(AGES)]
        vals += [getattr(row, f"v{k}") for k in range(AGES)]
        vals += [getattr(row, f"t{k}") for k in range(AGES)]
        vals += [row.rev_total]
        for c, v in enumerate(vals, 1):
            cell = wd.cell(r, c, float(v) if isinstance(v, (np.floating, float)) else (int(v) if isinstance(v, (np.integer, int)) else v))
            cell.font = BLUE if c >= 6 else BLACK
    nd = len(g) + 1
    for i in range(1, len(heads) + 1):
        wd.column_dimensions[col(i)].width = 12 if i > 5 else 14
    wd.freeze_panes = "G2"
    C_REV, C_VIS, C_RET = 7, 7 + AGES, 7 + 2 * AGES          # первые колонки блоков по возрасту
    C_TOT = 7 + 3 * AGES
    D = lambda c: f"'Данные'!${col(c)}$2:${col(c)}${nd}"                          # noqa: E731
    DB = lambda c1, c2: f"'Данные'!${col(c1)}$2:${col(c2)}${nd}"                  # noqa: E731
    COH, DIR, CLI, SRC, PAT = 1, 3, 4, 5, 6

    # ======================= Когорты =======================
    wk = wb.create_sheet("Когорты", 0)
    wk["A1"], wk["A1"].font = "LTV пациентов по когортам первого визита: накопленная выручка на пациента (тыс. ₽)", TITLE
    put_filter(wk, 2, "Направление первого визита:", "Все", ["Все"] + list(DIRS), "← «Все» или код: " + ", ".join(DIRS),
               '=IF($B$2="Все","*",$B$2)')
    put_filter(wk, 3, "Клиника первого визита:", "Все", ["Все"] + CLINICS, "← «Все», C01, C02 или C03",
               '=IF($B$3="Все","*",$B$3)')
    put_filter(wk, 4, "Канал привлечения:", "Все", ["Все"] + sources, "← «Все» или канал",
               '=IF($B$4="Все","*",$B$4)')
    wk["A5"] = ("Показана только наблюдённая часть, прогноза нет: данные заканчиваются 31.12.2025, январская когорта "
                "наблюдается до 12-го календарного месяца, декабрьская только в месяц первого визита.")
    wk["A5"].font = Font(name=F, size=9, italic=True)
    wk.column_dimensions["A"].width = 38
    hdr = ["Когорта (месяц первого визита)", "Пациентов", "Месяцев наблюдения"] + [f"К {k + 1}-му мес." for k in range(AGES)]
    HR = 7
    for i, h in enumerate(hdr, 1):
        wk.cell(HR, i, h)
    head(wk, HR, 1, len(hdr))
    wk.row_dimensions[HR].height = 32
    crit = f'{D(DIR)},$C$2,{D(CLI)},$C$3,{D(SRC)},$C$4'
    for c in range(1, AGES + 1):
        r = HR + c
        wk.cell(r, 1, f"{c}. {MONTHS[c - 1]} {YEAR}")
        wk.cell(r, 2, f"=SUMIFS({D(PAT)},{D(COH)},{c},{crit})")
        wk.cell(r, 3, AGES + 1 - c)
        for k in range(AGES):
            cell = wk.cell(r, 4 + k)
            if k <= AGES - c:
                colref = col(C_REV + k)
                cell.value = f"=IF($B{r}=0,\"—\",SUMIFS('Данные'!${colref}$2:${colref}${nd},{D(COH)},{c},{crit})/$B{r})"
            else:
                cell.value = "—"
            cell.number_format = NUM1
            cell.alignment = Alignment(horizontal="right")
        for cc in range(1, len(hdr) + 1):
            wk.cell(r, cc).font = BLACK
        wk.cell(r, 2).number_format = INT
    rp = HR + AGES + 1
    wk.cell(rp, 1, "Пул: когорты, у которых возраст наблюдается").font = BOLD
    wk.cell(rp, 2, f"=SUMIFS({D(PAT)},{crit})").number_format = INT
    for k in range(AGES):
        colref = col(C_REV + k)
        lim = AGES - k
        wk.cell(rp, 4 + k, (f'=IF(SUMIFS({D(PAT)},{D(COH)},"<={lim}",{crit})=0,"—",'
                            f"SUMIFS('Данные'!${colref}$2:${colref}${nd},{D(COH)},\"<={lim}\",{crit})/"
                            f'SUMIFS({D(PAT)},{D(COH)},"<={lim}",{crit}))'))
        wk.cell(rp, 4 + k).number_format = NUM1
        wk.cell(rp, 4 + k).alignment = Alignment(horizontal="right")
    for cc in range(1, len(hdr) + 1):
        wk.cell(rp, cc).font = BOLD
        wk.cell(rp, cc).border = TOP
    wk.cell(rp + 1, 1, "Пул считается по когортам, где данный возраст наблюдается, поэтому с ростом возраста в нём остаётся "
                       "меньше когорт, а состав по сезонам меняется. Для сравнения срезов используйте лист «Срезы».").font = HINT
    wk.conditional_formatting.add(f"D{HR + 1}:O{HR + AGES}", ColorScaleRule(
        start_type="min", start_color="FFFFFF", end_type="max", end_color="A9D08E"))
    for c in range(2, len(hdr) + 1):
        wk.column_dimensions[col(c)].width = 11
    wk.freeze_panes = "B8"

    # ======================= Срезы =======================
    ws = wb.create_sheet("Срезы", 1)
    ws["A1"], ws["A1"].font = "Срезы LTV: направление, клиника и канал на одном возрасте когорты (тыс. ₽)", TITLE
    put_filter(ws, 2, "Возраст k, месяцев после месяца первого визита:", 5, [str(i) for i in range(AGES)],
               "← 0-11; 5 значит выручка за 6 календарных месяцев, 11 за 12 (LTV12)")
    ws["C2"] = "=12-$B$2"
    ws["C2"].font = GREY
    ws["D2"] = "← берутся когорты не позднее этого номера месяца"
    ws["D2"].font = HINT
    ws["A3"] = ("В срез входят только когорты, у которых возраст k наблюдается. Чем больше k, тем меньше когорт. "
                "Меньше 100 пациентов: пометка «мало пациентов».")
    ws["A3"].font = Font(name=F, size=9, italic=True)
    ws.column_dimensions["A"].width = 34
    sh = ["Срез", "Пациентов", "LTV до k, тыс. ₽", "Визитов на пациента", "Вернулись (2+ визита)", "Выручка, тыс. ₽",
          "Примечание"]

    def block(top, title, dimcol, values, labels):
        ws.cell(top, 1, title).font = BOLD
        for c in range(1, len(sh) + 1):
            ws.cell(top, c).fill = SEC_FILL
        for i, h in enumerate(sh, 1):
            ws.cell(top + 1, i, h)
        head(ws, top + 1, 1, len(sh))
        ws.row_dimensions[top + 1].height = 30
        r0 = top + 2
        rows = [(v, labels.get(v, v)) for v in values] + [("*", "Все")]
        for i, (val, lab) in enumerate(rows):
            r = r0 + i
            ws.cell(r, 1, lab)
            ws.cell(r, 8, val).font = GREY
            cr = f'{D(COH)},"<="&$C$2'
            dimc = f"{D(dimcol)},$H{r}"
            pats = f"SUMIFS({D(PAT)},{cr},{dimc})"
            ws.cell(r, 2, "=" + pats)
            rev_k = f"SUMIFS(INDEX({DB(C_REV, C_REV + AGES - 1)},0,$B$2+1),{cr},{dimc})"
            vis_k = f"SUMIFS(INDEX({DB(C_VIS, C_VIS + AGES - 1)},0,$B$2+1),{cr},{dimc})"
            ret_k = f"SUMIFS(INDEX({DB(C_RET, C_RET + AGES - 1)},0,$B$2+1),{cr},{dimc})"
            ws.cell(r, 3, f"=IF(B{r}=0,0,{rev_k}/B{r})")
            ws.cell(r, 4, f"=IF(B{r}=0,0,{vis_k}/B{r})")
            ws.cell(r, 5, f"=IF(B{r}=0,0,{ret_k}/B{r})")
            ws.cell(r, 6, "=" + rev_k)
            ws.cell(r, 7, f'=IF(B{r}=0,"нет пациентов",IF(B{r}<{MIN_N},"мало пациентов","—"))')
            for c in range(1, 8):
                cell = ws.cell(r, c)
                cell.font = BOLD if val == "*" else BLACK
                if val == "*":
                    cell.border = TOP
            ws.cell(r, 2).number_format = INT
            ws.cell(r, 3).number_format = NUM1
            ws.cell(r, 4).number_format = "0.00"
            ws.cell(r, 5).number_format = "0.0%"
            ws.cell(r, 6).number_format = "#,##0,;(#,##0,);-"
            ws.cell(r, 7).font = Font(name=F, size=9, italic=True, color="C00000")
        return r0 + len(rows)

    nxt = block(5, "По направлению первого визита", DIR, list(DIRS), DIRS) + 1
    nxt = block(nxt, "По клинике первого визита", CLI, CLINICS, {}) + 1
    block(nxt, "По каналу привлечения", SRC, sources, {})
    ws.column_dimensions["H"].width = 6
    for c in range(2, 8):
        ws.column_dimensions[col(c)].width = 16
    ws.column_dimensions["G"].width = 24
    ws.freeze_panes = "B5"

    # ======================= Описание =======================
    wi = wb.create_sheet("Описание", 0)
    wi["A1"], wi["A1"].font = "LTV пациентов 2025: описание и контроль", TITLE
    lines = [
        ("Определения", None),
        ("Новый пациент", "Первый визит (first_visit_date) в 2025 году и есть завершённые визиты. Пациенты с историей до 2025 года "
                          "исключены: их визитов до 2025 года в данных нет, LTV с первого визита по ним посчитать нельзя."),
        ("Когорта и возраст", "Когорта это месяц первого завершённого визита (у 4,4% новых пациентов он позже даты регистрации, потому что "
                              "первая запись была отменена или пациент не пришёл). Возраст k это число календарных месяцев после месяца "
                              "первого визита (0 = месяц первого визита)."),
        ("LTV до k", "Накопленная выручка на пациента (по датам визитов) за возрасты 0..k. LTV12 это возраст 11, то есть месяц первого "
                     "визита и 11 следующих."),
        ("Только факт", None),
        ("Без прогноза", "Данные заканчиваются 31.12.2025, поэтому когорта c наблюдается до возраста 12 − c: январская полностью, "
                         "декабрьская только в месяц первого визита. Ненаблюдаемые значения не показываются и не прогнозируются. "
                         "Прогнозы по врачам и направлениям строятся отдельным блоком."),
        ("Сравнение срезов", "Срезы сравниваются на одном возрасте k и только по когортам, где он наблюдается. LTV12 полностью наблюдается "
                             f"только у январской когорты ({str(format(ctrl['n_jan'], ',')).replace(',', ' ')} пациентов), поэтому по срезам при k = 11 "
                             "выводы статистически слабые: см. примечание «мало пациентов»."),
        ("Ограничения", None),
        ("Сезонность", "Когорты разных месяцев различаются по сезонности спроса и каналам, поэтому «пул» на листе «Когорты» при "
                       "росте возраста собирается из меньшего числа когорт и не является единой кривой жизни пациента."),
        ("Канал «Повторное обращение»", "У части новых пациентов канал «Повторное обращение»: для нового пациента это странная метка, "
                                        "оставлена как в источнике, срезы по ней осторожно."),
        ("Данные синтетические", "Модель клиники: цифры демонстрируют метод, а не реальную ценность пациентов."),
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
    wi.cell(r, 1, "Контроль (при фильтрах «Все»)").font = BOLD
    for c in (1, 2, 3, 4):
        wi.cell(r, c).fill = SEC_FILL
    r += 1
    for i, h in enumerate(["Показатель", "По книге", "Источник", "Разница"], 1):
        wi.cell(r, i, h)
    head(wi, r, 1, 4)
    rows_ = [("Новых пациентов", f"=SUM({D(PAT)})", ctrl["patients"], "#,##0"),
             ("Выручка новых пациентов за наблюдаемый период, тыс. ₽", f"=SUM({D(C_TOT)})", ctrl["rev_new"], "#,##0.00,"),
             ("LTV12 январской когорты, тыс. ₽ на пациента", f"='Когорты'!O{HR + 1}", ctrl["ltv12_jan"], "#,##0.00,")]
    for name, f, src, fmt in rows_:
        r += 1
        wi.cell(r, 1, name).font = BLACK
        wi.cell(r, 2, f).number_format = fmt
        wi.cell(r, 3, round(float(src), 2)).font = BLUE
        wi.cell(r, 3).number_format = fmt
        wi.cell(r, 4, f"=B{r}-C{r}").number_format = "#,##0.00;(#,##0.00);-"
    r += 1
    wi.cell(r, 1, "Справка: выручка прочих пациентов (история до 2025 года), тыс. ₽").font = BLACK
    wi.cell(r, 2, round(ctrl["rev_old"], 2)).number_format = "#,##0.00,"
    wi.cell(r, 2).font = BLUE
    wi.cell(r, 3, "Выручка ERP за год, тыс. ₽").font = HINT
    wi.cell(r, 4, round(ctrl["rev_total"], 2)).number_format = "#,##0.00,"
    wi.cell(r, 4).font = BLUE
    wi.column_dimensions["A"].width = 56
    wi.column_dimensions["B"].width = 110
    wi.column_dimensions["C"].width = 16
    wi.column_dimensions["D"].width = 16

    wb._sheets = [wb["Описание"], wb["Когорты"], wb["Срезы"], wb["Данные"]]
    wb.active = 1
    return wb


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    g, pt, ctrl = prepare()
    sources = sorted(pt.source.unique())
    wb = build(g, ctrl, sources)
    path = OUT / "PatientLTV_2025.xlsx"
    wb.save(path)
    recalculate(path)
    print(f"Сохранено: {path}  (групп: {len(g)}, новых пациентов: {int(ctrl['patients']):,})")
    print("Выручка новых пациентов", f"{ctrl['rev_new']:,.0f}", "| прочих", f"{ctrl['rev_old']:,.0f}",
          "| всего", f"{ctrl['rev_total']:,.0f}", "| LTV12 января", f"{ctrl['ltv12_jan']:,.0f}")


if __name__ == "__main__":
    main()
