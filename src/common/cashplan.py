"""
Плановый денежный бюджет (БДДС): плановые начисления x график платежей.

Плановые начисления берутся из budget_revenue и budget_opex (плановый P&L), график платежей это учётная политика и
договоры (восстановлен по фактической дате платежей в cashflow_1c):
  поступления от пациентов   в месяце оказания услуги, доля пациентов = 1 - доля ДМС по направлению
  поступления от страховых   ДМС-выручка месяца приходит в следующем месяце (30-45 дней)
  входящая дебиторка ДМС     сальдо 62.02 на 01.01, поступает поровну в январе и феврале
  зарплата                   начислено x (1 - НДФЛ) выплачивается 10-го числа следующего месяца
  НДФЛ и взносы              уплачиваются 28-го числа следующего месяца
  материалы и лаборатории    (план количество x плановые нормы) оплачиваются 20-го следующего месяца
  маркетинг                  оплачивается 15-го следующего месяца
  аренда, коммунальные,
  эквайринг                  в месяце начисления
  входящая кредиторка        сальдо 70, 68, 69 на 01.01 уплачивается в январе; сальдо 60 в январе делится между
                             материалами и маркетингом пропорционально их плановым начислениям января
  инвестпрограмма            утверждённые закупки оборудования

Допущения демо (описаны в ASSUMPTIONS.md): доли ДМС в плане равны фактическим долям 2025; инвестпрограмма равна факту;
налог на прибыль, НДС и дивиденды не планируются (как и в факте).
Знак: поступления плюс, выплаты минус.
"""
import pandas as pd

YEAR = 2025
NDFL = 0.13
ARTICLES = [  # (вид деятельности, статья ДДС) в порядке отчёта
    ("operating", "Поступления от пациентов"),
    ("operating", "Поступления от страховых компаний"),
    ("operating", "Выплата заработной платы"),
    ("operating", "НДФЛ"),
    ("operating", "Страховые взносы"),
    ("operating", "Материалы и лаборатории"),
    ("operating", "Аренда помещений"),
    ("operating", "Коммунальные услуги"),
    ("operating", "Маркетинг и реклама"),
    ("operating", "Комиссия эквайринга"),
    ("investing", "Приобретение оборудования"),
]
ACTIVITY = {a: act for act, a in ARTICLES}
CAPEX_PLAN = {("C01", 3): 12_000_000.0, ("C03", 6): 7_500_000.0, ("C02", 9): 4_200_000.0}
SAME_MONTH = {"Аренда": "Аренда помещений", "Коммунальные услуги": "Коммунальные услуги",
              "Комиссия эквайринга": "Комиссия эквайринга"}
PAYROLL = ["Оплата труда врачей", "Оплата труда медперсонала", "Оплата труда административного персонала"]
CONTRIB = ["Страховые взносы (себестоимость)", "Страховые взносы (администрация)"]


def _next(period):
    y, m = map(int, period.split("-"))
    return f"{y}-{m + 1:02d}" if m < 12 else None


def opening_balances(osv):
    """Сальдо на 01.01 по клиникам: {счёт: Series(клиника -> сумма)}; кредитовые остатки положительным числом."""
    o = osv[osv.period == f"{YEAR}-01"].copy()
    o["account"] = o.account.astype(str)
    g = o.groupby(["account", "clinic_id"])[["opening_debit", "opening_credit"]].sum()
    out = {}
    for acc in ("51", "60", "62.02", "68", "69", "70"):
        s = g.loc[acc] if acc in g.index.get_level_values(0) else pd.DataFrame(columns=g.columns)
        out[acc] = (s.opening_debit if acc in ("51", "62.02") else s.opening_credit).to_dict()
    return out


def dms_share(payments, visits):
    """Доля ДМС в оплатах по направлению (допущение: плановая доля равна фактической 2025)."""
    v = payments.merge(visits[["visit_id", "direction_code"]], on="visit_id")
    t = v.pivot_table(index="direction_code", columns="payer_type", values="amount", aggfunc="sum").fillna(0)
    return (t["dms"] / t.sum(axis=1)).to_dict()


def plan_cash(budget_rev, budget_opex, share, opening):
    rows = []     # (period, clinic, article, signed amount)
    rev = budget_rev.copy()
    rev["share"] = rev.direction_code.map(share).fillna(0.0)
    pat = rev.assign(a=rev.plan_revenue * (1 - rev.share)).groupby(["period", "clinic_id"]).a.sum()
    for (p, c), v in pat.items():
        rows.append((p, c, "Поступления от пациентов", v))
    dms = rev.assign(a=rev.plan_revenue * rev.share).groupby(["period", "clinic_id"]).a.sum()
    for (p, c), v in dms.items():
        q = _next(p)
        if q:
            rows.append((q, c, "Поступления от страховых компаний", v))
    for c, v in opening["62.02"].items():
        rows.append((f"{YEAR}-01", c, "Поступления от страховых компаний", v / 2))
        rows.append((f"{YEAR}-02", c, "Поступления от страховых компаний", v / 2))

    mat = rev.assign(a=rev.plan_qty * (rev.plan_unit_material + rev.plan_unit_external)) \
             .groupby(["period", "clinic_id"]).a.sum()
    for (p, c), v in mat.items():
        q = _next(p)
        if q:
            rows.append((q, c, "Материалы и лаборатории", -v))

    ox = budget_opex
    for src, art in SAME_MONTH.items():
        for r in ox[ox.article == src].itertuples():
            rows.append((r.period, r.clinic_id, art, -r.plan_amount))
    mk = ox[ox.article == "Маркетинг"]
    for r in mk.itertuples():
        q = _next(r.period)
        if q:
            rows.append((q, r.clinic_id, "Маркетинг и реклама", -r.plan_amount))
    gross = ox[ox.article.isin(PAYROLL)].groupby(["period", "clinic_id"]).plan_amount.sum()
    for (p, c), v in gross.items():
        q = _next(p)
        if q:
            rows.append((q, c, "Выплата заработной платы", -v * (1 - NDFL)))
            rows.append((q, c, "НДФЛ", -v * NDFL))
    contr = ox[ox.article.isin(CONTRIB)].groupby(["period", "clinic_id"]).plan_amount.sum()
    for (p, c), v in contr.items():
        q = _next(p)
        if q:
            rows.append((q, c, "Страховые взносы", -v))

    jan = f"{YEAR}-01"
    for c, v in opening["70"].items():
        rows.append((jan, c, "Выплата заработной платы", -v))
    for c, v in opening["68"].items():
        rows.append((jan, c, "НДФЛ", -v))
    for c, v in opening["69"].items():
        rows.append((jan, c, "Страховые взносы", -v))
    for c, v in opening["60"].items():
        wm = float(mat.get((jan, c), 0.0))
        wk = float(mk[(mk.period == jan) & (mk.clinic_id == c)].plan_amount.sum())
        tot = wm + wk
        wm, wk = (wm / tot, wk / tot) if tot else (1.0, 0.0)
        rows.append((jan, c, "Материалы и лаборатории", -v * wm))
        rows.append((jan, c, "Маркетинг и реклама", -v * wk))
    for (c, m), v in CAPEX_PLAN.items():
        rows.append((f"{YEAR}-{m:02d}", c, "Приобретение оборудования", -v))

    d = pd.DataFrame(rows, columns=["period", "clinic_id", "article", "plan"])
    d = d.groupby(["period", "clinic_id", "article"], as_index=False).plan.sum()
    d["activity"] = d.article.map(ACTIVITY)
    return d
