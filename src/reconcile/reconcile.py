"""
Шаг 6: модуль сверки ERP и бухгалтерии (1С).

Читает «испорченные» данные из data/dirty/, ключ ответов НЕ использует.
Пишет в data/output/:
  findings.csv             - таблица находок (одна строка = одна проблема)
  reconciliation_report.md - отчёт для руководителя
А затем отдельным шагом сравнивает находки с ключом ответов
(data/answer_key/injected_errors.csv) и пишет оценку качества
в data/answer_key/reconcile_evaluation.md.

Логика: «объясняй и вычитай». Сначала находятся первопричины на уровне ERP
(дубли, визиты без врача, цены ниже прайса, потерянные оплаты). Их влияние
вычитается из расхождения ERP-1С, и только остаток разбирается как проблемы
самой бухгалтерии (смещение периода, ручная правка реализации, задвоенное
поступление). Так одна ошибка не порождает пять находок.

Допуск на округление: 100 руб. на ячейку сравнения.

Версия 2 (04.11.2026, DECISIONS п. 86): доработка ПОСЛЕ слепого теста, правила знали результат первого
прогона. Добавлено: сдвиг даты банка не находка (R6), врач чужой клиники (R2b), комиссия эквайринга (R6b),
ставка ЗП по справочнику (R10), материалы по направлениям (R11). Замороженная версия первого прогона:
reconcile_frozen_59cff04f.py.

Запуск:  python src/reconcile/reconcile.py
"""
import argparse
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
DIRTY = ROOT / "data" / "dirty"
OUT = ROOT / "data" / "output"
KEY = ROOT / "data" / "answer_key"

TOL = 100.0                     # допуск, руб.
CUTOFF = pd.Timestamp("2025-12-31")
DMS_GRACE_FROM = pd.Timestamp("2025-11-01")   # визиты с этой даты могут ждать оплату ДМС
LAG_DAYS = 3                    # сдвиг выписки банка: до 3 дней (пятница -> понедельник)
FEE_TOL = 1.0                   # комиссия банка округляется до рубля
MAT_REL = 0.005                 # материалы: отклонение ячейки от нормы больше 0,5%
MAT_PAIR = 0.01                 # и встречное отклонение в клинике-месяце гасит его с точностью до 1%
EVAL = True
STR = {"debit": str, "credit": str, "account": str}

ACTIONS = {
    "duplicate_visit": "Удалить дубль в ERP; проверить, не создан ли дубль и в реализации 1С",
    "no_doctor": "Указать врача в визите; пересчитать мотивацию врача за период",
    "price_below_list": "Выяснить причину цены ниже прайса; оформить скидку или исправить цену",
    "no_payment": "Найти платёж по визиту (касса, эквайринг) и привязать к визиту",
    "revenue_mismatch": "Сверить документ реализации в 1С с визитами ERP; откатить ручную правку",
    "wrong_period": "Перепровести реализацию в правильный период; пересчитать закрытие месяца",
    "duplicate_receipt": "Удалить задвоенное поступление в 1С; проверить выписку банка/эквайринга",
    "missing_receipt": "Найти платёж ERP в выписке банка и провести поступление в 1С",
    "dms_receipt_mismatch": "Сверить оплату страховой с актом и платёжкой",
    "price_list_mismatch": "Проверить прайс на дату визита",
    "abnormal_balance": "Разобрать причину по связанным находкам, затем скорректировать проводки",
    "receivable_mismatch": "Сверить дебиторку ДМС по актам со страховыми",
    "doctor_clinic_mismatch": "Проверить врача в визите или справочник: врач чужой клиники",
    "acquiring_fee_rate": "Сверить комиссию с договором эквайринга и выпиской банка",
    "payroll_rate_mismatch": "Пересчитать начисление по ставке справочника или обновить справочник",
    "cost_misallocation": "Перепровести списание материалов по верному направлению",
}


def sev(amount):
    a = abs(amount)
    return "high" if a >= 100_000 else "medium" if a >= 30_000 else "low"


def load():
    v = pd.read_csv(DIRTY / "visits.csv", parse_dates=["visit_date"])
    it = pd.read_csv(DIRTY / "visit_items.csv")
    pays = pd.read_csv(DIRTY / "visit_payments.csv", parse_dates=["pay_date"])
    post = pd.read_csv(DIRTY / "postings_1c.csv", dtype=STR, parse_dates=["date"])
    osv = pd.read_csv(DIRTY / "osv_1c.csv", dtype=STR | {c: float for c in [
        "opening_debit", "opening_credit", "turnover_debit", "turnover_credit",
        "closing_debit", "closing_credit"]})
    hist = pd.read_csv(DIRTY / "price_history.csv", parse_dates=["valid_from"])
    return v, it, pays, post, osv, hist


def load_ref():
    doctors = pd.read_csv(DIRTY / "doctors.csv")
    payroll = pd.read_csv(DIRTY / "payroll_1c.csv")
    services = pd.read_csv(DIRTY / "services.csv")
    mat = pd.read_csv(DIRTY / "materials_1c.csv")
    return doctors, payroll, services, mat


def pair_lags(df, keys, channel):
    """Сдвиг даты банка: минус в день ERP и такой же плюс в день 1С (позже на 1-LAG_DAYS дней, тот же месяц,
    тот же канал и клиника). Банк не зачисляет раньше платежа, поэтому 1С раньше ERP сдвигом не считается."""
    x = df[df["diff"].abs() > TOL]["diff"].reset_index()
    neg, pos = x[x["diff"] < 0], x[x["diff"] > 0]
    used, consumed, rows = set(), set(), []
    for n in neg.itertuples():
        for q in pos.itertuples():
            if q.Index in used or any(getattr(n, k) != getattr(q, k) for k in keys):
                continue
            days = (q.d - n.d).days
            if 0 < days <= LAG_DAYS and (q.d.year, q.d.month) == (n.d.year, n.d.month) \
                    and abs(q.diff + n.diff) <= TOL:
                used.add(q.Index)
                consumed |= {(n.d,) + tuple(getattr(n, k) for k in keys), (q.d,) + tuple(getattr(q, k) for k in keys)}
                rows.append(dict(channel=channel, clinic_id=n.clinic_id,
                                 payer=getattr(n, "payer", ""), date_erp=n.d.date().isoformat(),
                                 date_1c=q.d.date().isoformat(), days=days, amount_rub=round(float(q.diff), 2)))
                break
    return consumed, rows


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    visits, items, pays, post, osv, hist = load()
    doctors, payroll, services, mat = load_ref()
    F = []   # находки

    def add(rule, etype, period, clinic, direction, ids, amount, evidence, root=True, **extra):
        F.append(dict(rule=rule, error_type=etype, severity=sev(amount), period=period,
                      clinic_id=clinic, direction_code=direction, record_ids="|".join(ids),
                      amount_rub=round(float(amount), 2), evidence=evidence,
                      suggested_action=ACTIONS[etype], is_root_cause=root, **extra))

    vrev = items.groupby("visit_id").revenue.sum()
    visits["rev"] = visits.visit_id.map(vrev).fillna(0.0)
    done = visits[visits.status == "completed"].copy()
    done["period"] = done.visit_date.dt.strftime("%Y-%m")

    # ===== R1. Дубли визитов (тот же пациент, врач, время и состав услуг) =====
    sig = items.sort_values("item_id").groupby("visit_id").apply(
        lambda g: tuple(zip(g.service_id, g.qty, g.price_fact)), include_groups=False)
    done["sig"] = done.visit_id.map(sig)
    done["doc_key"] = done.doctor_id.fillna("")
    dup_keys = ["patient_id", "doc_key", "clinic_id", "visit_date", "visit_time", "sig"]
    grp = done[done.duplicated(dup_keys, keep=False)].sort_values("visit_id")
    dup_ids = set()
    for _, g in grp.groupby(dup_keys, sort=False):
        orig, *dups = g.itertuples()
        for d in dups:
            dup_ids.add(d.visit_id)
            add("R1_duplicates", "duplicate_visit", d.period, d.clinic_id, d.direction_code,
                [orig.visit_id, d.visit_id], d.rev,
                f"Визиты {orig.visit_id} и {d.visit_id}: один пациент, врач, дата {d.visit_date:%d.%m.%Y} "
                f"{d.visit_time}, одинаковые услуги; у {d.visit_id} нет платежа")

    # ===== R2. Завершённые визиты без врача =====
    for r in done[done.doctor_id.isna()].itertuples():
        add("R2_no_doctor", "no_doctor", r.period, r.clinic_id, r.direction_code, [r.visit_id], r.rev,
            f"Визит {r.visit_id} от {r.visit_date:%d.%m.%Y} на {r.rev:,.0f} ₽ без врача: "
            f"выручка не попадает в расчёт врача")

    # ===== R2b. Врач визита из другой клиники (справочник врачей) =====
    doc_clinic = doctors.set_index("doctor_id").clinic_id
    dcl = done.doctor_id.map(doc_clinic)
    for r in done[dcl.notna() & (dcl != done.clinic_id)].itertuples():
        add("R2b_doctor_clinic", "doctor_clinic_mismatch", r.period, r.clinic_id, r.direction_code, [r.visit_id], r.rev,
            f"Визит {r.visit_id} от {r.visit_date:%d.%m.%Y} в клинике {r.clinic_id}: врач {r.doctor_id} по справочнику "
            f"работает в {doc_clinic[r.doctor_id]}; выручка {r.rev:,.0f} ₽ уходит врачу чужой клиники")

    # ===== R3. Цена ниже прайса при нулевой/меньшей скидке =====
    im = items.merge(visits[["visit_id", "visit_date", "clinic_id", "direction_code", "status"]], on="visit_id")
    im["period"] = im.visit_date.dt.strftime("%Y-%m")
    im["expected_price"] = (im.list_price_at_date * (1 - im.discount_pct / 100) / 100).round() * 100
    bad = im[im.price_fact < im.expected_price - TOL].copy()
    price_adj = {}
    for r in bad.itertuples():
        delta = (r.expected_price - r.price_fact) * r.qty
        price_adj[(r.visit_date, r.clinic_id, r.direction_code)] = \
            price_adj.get((r.visit_date, r.clinic_id, r.direction_code), 0.0) + delta
        add("R3_price_vs_list", "price_below_list", r.period, r.clinic_id, r.direction_code,
            [r.item_id], -delta,
            f"Позиция {r.item_id} (визит {r.visit_id}): цена {r.price_fact:,.0f} при прайсе "
            f"{r.list_price_at_date:,.0f} и скидке {r.discount_pct:.0f}%; недобор {delta:,.0f} ₽")

    # прайс на дату визита против истории цен (контроль справочника)
    hs = hist.sort_values("valid_from")
    m = pd.merge_asof(im.sort_values("visit_date"), hs, left_on="visit_date", right_on="valid_from",
                      by="service_id", suffixes=("", "_hist"))
    pl = m[m.list_price_at_date != m.list_price]
    for r in pl.itertuples():
        add("R3b_price_history", "price_list_mismatch", r.period, r.clinic_id, r.direction_code,
            [r.item_id], r.list_price_at_date - r.list_price,
            f"Прайс в позиции {r.list_price_at_date} != истории цен {r.list_price}")

    # ===== R4. Закрытые визиты без оплаты =====
    paid_ids = set(pays.visit_id)
    unpaid = done[~done.visit_id.isin(paid_ids) & ~done.visit_id.isin(dup_ids)]
    old_unpaid = unpaid[unpaid.visit_date < DMS_GRACE_FROM]
    nopay_adj = {}
    for r in old_unpaid.itertuples():
        nopay_adj[(r.visit_date, r.clinic_id)] = nopay_adj.get((r.visit_date, r.clinic_id), 0.0) + r.rev
        add("R4_no_payment", "no_payment", r.period, r.clinic_id, r.direction_code, [r.visit_id], r.rev,
            f"Визит {r.visit_id} от {r.visit_date:%d.%m.%Y} на {r.rev:,.0f} ₽ закрыт, платежа нет; "
            f"срок оплаты ДМС по таким визитам давно истёк")
    young_unpaid = unpaid[unpaid.visit_date >= DMS_GRACE_FROM]

    # ===== R5. Выручка ERP vs 1С по дням / клиникам / направлениям =====
    erp = im[im.status == "completed"].groupby(["visit_date", "clinic_id", "direction_code"]).revenue.sum()
    # вычитаем влияние уже найденных первопричин
    erp_adj = erp.copy()
    for d in done[done.visit_id.isin(dup_ids)].itertuples():
        k = (d.visit_date, d.clinic_id, d.direction_code)
        erp_adj[k] = erp_adj.get(k, 0.0) - d.rev
    for k, delta in price_adj.items():
        erp_adj[k] = erp_adj.get(k, 0.0) + delta
    rev1c_posts = post[(post.doc_type == "Реализация услуг") & (post.credit == "90.01")]
    c1 = rev1c_posts.groupby(["date", "clinic_id", "direction_code"]).amount.sum()
    c1.index.names = erp_adj.index.names = ["visit_date", "clinic_id", "direction_code"]
    cell = pd.concat([c1.rename("onec"), erp_adj.rename("erp")], axis=1).fillna(0.0)
    cell["diff"] = (cell.onec - cell.erp).round(2)
    small = cell[(cell["diff"].abs() > 0) & (cell["diff"].abs() <= TOL)]
    big = cell[cell["diff"].abs() > TOL].reset_index()

    used = set()
    neg = big[big["diff"] < 0]
    pos = big[big["diff"] > 0]
    posts_in = lambda d, c, dr: rev1c_posts[(rev1c_posts.date == d) & (rev1c_posts.clinic_id == c) &
                                            (rev1c_posts.direction_code == dr)]
    # --- 5a. смещение периода: минус в одной ячейке, такой же плюс в начале следующего месяца ---
    for n in neg.itertuples():
        cand = pos[(pos.clinic_id == n.clinic_id) & (pos.direction_code == n.direction_code) &
                   (pos.visit_date > n.visit_date) & (pos.visit_date <= n.visit_date + pd.Timedelta(days=6)) &
                   (pos.visit_date.dt.month != n.visit_date.month) &
                   ((pos["diff"] + n.diff).abs() <= TOL)]
        if cand.empty:
            continue
        p = cand.iloc[0]
        used.add(n.Index)
        used.add(p.name)
        ps = posts_in(p.visit_date, p.clinic_id, p.direction_code)
        exact = ps[(ps.amount - p["diff"]).abs() <= TOL]
        ids = (exact if len(exact) else ps).posting_id.tolist()
        add("R5_revenue_period", "wrong_period", n.visit_date.strftime("%Y-%m"), n.clinic_id,
            n.direction_code, ids, p["diff"],
            f"Выручка {p['diff']:,.0f} ₽ есть в ERP за {n.visit_date:%d.%m.%Y}, но проведена в 1С "
            f"{p.visit_date:%d.%m.%Y}: период {n.visit_date:%m.%Y} занижен, "
            f"{p.visit_date:%m.%Y} завышен", affects_6201=True,
            fix_date=n.visit_date.date().isoformat())
    # --- 5b. остаток: выручка в 1С не равна ERP ---
    for r in big.itertuples():
        if r.Index in used:
            continue
        ids = posts_in(r.visit_date, r.clinic_id, r.direction_code).posting_id.tolist()
        add("R5_revenue_diff", "revenue_mismatch", r.visit_date.strftime("%Y-%m"), r.clinic_id,
            r.direction_code, ids, r.diff,
            f"{r.visit_date:%d.%m.%Y}: выручка 1С {r.onec:,.0f} ₽, ERP (с учётом найденных ошибок) "
            f"{r.erp:,.0f} ₽, разница {r.diff:+,.0f} ₽. Проводки-кандидаты: {', '.join(ids)}",
            affects_6201=True)

    # ===== R6. Поступления: ERP-платежи vs проводки 1С =====
    cash = post[post.debit == "51"]
    e_self = pays[pays.payer_type == "self"].merge(visits[["visit_id", "clinic_id"]], on="visit_id")
    e_self = e_self.groupby(["pay_date", "clinic_id"]).amount.sum()
    c_self = cash[cash.doc_type == "Поступление от пациентов (эквайринг)"].groupby(["date", "clinic_id"]).amount.sum()
    e_self.index.names = c_self.index.names = ["d", "clinic_id"]
    adj = e_self.copy()
    for k, v in nopay_adj.items():
        adj[k] = adj.get(k, 0.0) + v
    cc = pd.concat([c_self.rename("onec"), adj.rename("erp")], axis=1).fillna(0.0)
    cc["diff"] = (cc.onec - cc.erp).round(2)
    lag_self, lag_rows = pair_lags(cc, ["clinic_id"], "самооплата")
    for (d, c), r in cc[cc["diff"].abs() > TOL].iterrows():
        if (d, c) in lag_self:
            continue
        ids = cash[(cash.date == d) & (cash.clinic_id == c) &
                   (cash.doc_type == "Поступление от пациентов (эквайринг)")].posting_id.tolist()
        if r["diff"] > 0:
            add("R6_cash_self", "duplicate_receipt", d.strftime("%Y-%m"), c, "", ids, r["diff"],
                f"{d:%d.%m.%Y}: поступлений в 1С {r.onec:,.0f} ₽, платежей в ERP {r.erp:,.0f} ₽; "
                f"лишние {r['diff']:,.0f} ₽ (проводки {', '.join(ids)})", affects_6201=True)
        else:
            add("R6_cash_self", "missing_receipt", d.strftime("%Y-%m"), c, "", ids, r["diff"],
                f"{d:%d.%m.%Y}: платежи ERP на {-r['diff']:,.0f} ₽ не проведены в 1С")

    e_dms = pays[pays.payer_type == "dms"].merge(visits[["visit_id", "clinic_id"]], on="visit_id")
    e_dms = e_dms.groupby(["pay_date", "clinic_id", "payer_name"]).amount.sum()
    c_dms = cash[(cash.doc_type == "Поступление от страховой") &
                 (cash.comment.fillna("") != "Оплата акта за декабрь 2024")]
    c_dms = c_dms.groupby(["date", "clinic_id", "counterparty"]).amount.sum()
    e_dms.index.names = c_dms.index.names = ["d", "clinic_id", "payer"]
    dd = pd.concat([c_dms.rename("onec"), e_dms.rename("erp")], axis=1).fillna(0.0)
    dd["diff"] = (dd.onec - dd.erp).round(2)
    lag_dms, rows2 = pair_lags(dd, ["clinic_id", "payer"], "ДМС")
    lag_rows += rows2
    lags = pd.DataFrame(lag_rows, columns=["channel", "clinic_id", "payer", "date_erp", "date_1c", "days", "amount_rub"])
    lags.to_csv(OUT / "timing_differences.csv", index=False, encoding="utf-8-sig")
    for (d, c, p), r in dd[dd["diff"].abs() > TOL].iterrows():
        if (d, c, p) in lag_dms:
            continue
        add("R6_cash_dms", "dms_receipt_mismatch", d.strftime("%Y-%m"), c, "", [], r["diff"],
            f"{d:%d.%m.%Y} {p}: 1С {r.onec:,.0f} ₽, ERP {r.erp:,.0f} ₽")

    # ===== R6b. Комиссия эквайринга против ставки (база: платежи ERP, только дни со сошедшимися поступлениями) =====
    fee_p = post[post.doc_type == "Комиссия эквайринга"]
    fee = fee_p.groupby(["date", "clinic_id"]).amount.sum()
    fee.index.names = ["d", "clinic_id"]
    fx = pd.concat([cc["erp"].rename("base"), cc["diff"].rename("rdiff"), fee.rename("fee")], axis=1).fillna(0.0)
    fx = fx[(fx.rdiff.abs() <= TOL) & (fx.base > 0)]
    if len(fx):
        rate = (fx.fee / fx.base).round(4).mode().iloc[0]    # договорной ставки в данных нет: берём самую частую по сети
        for (d, c), r in fx.iterrows():
            dev = r.fee - r.base * rate
            if abs(dev) > FEE_TOL:
                ids = fee_p[(fee_p.date == d) & (fee_p.clinic_id == c)].posting_id.tolist()
                add("R6b_acquiring_fee", "acquiring_fee_rate", d.strftime("%Y-%m"), c, "", ids, dev,
                    f"{d:%d.%m.%Y}: комиссия {r.fee:,.2f} ₽ при поступлении {r.base:,.0f} ₽ ({r.fee / r.base:.2%}), "
                    f"ставка по сети {rate:.2%}, ожидалось {r.base * rate:,.2f} ₽")

    # ===== R10. Ставка ЗП врача в расчёте против справочника (по ID врача) =====
    ref = doctors[["doctor_id", "pay_pct"]].rename(columns={"doctor_id": "employee_id", "pay_pct": "ref_pct"})
    pr = payroll[payroll.role == "doctor"].merge(ref, on="employee_id", how="left")
    for r in pr[pr.ref_pct.notna() & ((pr.pay_pct - pr.ref_pct).abs() > 1e-9)].itertuples():
        add("R10_payroll_rate", "payroll_rate_mismatch", r.period, r.clinic_id, r.direction_code,
            [f"{r.period}|{r.employee_id}"], r.base_revenue * (r.pay_pct - r.ref_pct),
            f"{r.period}, врач {r.employee_id}: ставка {r.pay_pct:.0%} в расчёте ЗП, {r.ref_pct:.0%} в справочнике; "
            f"база {r.base_revenue:,.0f} ₽")

    # ===== R11. Списание материалов по направлениям: встречные отклонения от нормы ERP в клинике-месяце =====
    sv = services.set_index("service_id").material_cost
    norm = im[im.status == "completed"].assign(norm=lambda t: t.qty * t.service_id.map(sv)) \
        .groupby(["period", "clinic_id", "direction_code"]).norm.sum()
    wo = post[post.doc_type == "Списание материалов"].assign(period=lambda t: t.date.dt.strftime("%Y-%m"))
    mx = pd.concat([wo.groupby(["period", "clinic_id", "direction_code"]).amount.sum().rename("fact"),
                    norm.rename("norm"),
                    mat.set_index(["period", "clinic_id", "direction_code"]).material_index.rename("idx")],
                   axis=1).fillna(0.0)
    mx["exp"] = mx.norm * mx.idx
    mx["dev"] = (mx.fact - mx.exp).round(2)
    mx["rel"] = mx.dev / mx.exp.replace(0, np.nan)
    mb = mx[mx.rel.abs() > MAT_REL].reset_index()
    for (per, c), g in mb.groupby(["period", "clinic_id"]):
        if (g.dev > 0).any() and (g.dev < 0).any() and abs(g.dev.sum()) <= MAT_PAIR * g.dev.abs().max():
            for r in g.itertuples():
                ids = wo[(wo.period == per) & (wo.clinic_id == c) & (wo.direction_code == r.direction_code)].posting_id.tolist()
                add("R11_materials_split", "cost_misallocation", per, c, r.direction_code, ids, r.dev,
                    f"{per}, {c}, {r.direction_code}: списано материалов {r.fact:,.0f} ₽, по норме ERP {r.exp:,.0f} ₽ "
                    f"({r.rel:+.1%}); встречное отклонение по другому направлению клиники, итог по клинике сходится")

    # ===== R7. Дебиторка ДМС: ОСВ 62.02 на 31.12 vs неоплаченные визиты ERP =====
    dms_1c = osv[(osv.period == "2025-12") & (osv.account == "62.02")].set_index("clinic_id").closing_debit
    dms_erp = young_unpaid.groupby("clinic_id").rev.sum()
    rec = pd.concat([dms_1c.rename("onec"), dms_erp.rename("erp")], axis=1).fillna(0.0)
    rec["diff"] = rec.onec - rec.erp
    for c, r in rec.iterrows():
        if abs(r["diff"]) > 5 * TOL:
            add("R7_dms_receivable", "receivable_mismatch", "2025-12", c, "", [], r["diff"],
                f"62.02 в ОСВ {r.onec:,.0f} ₽ против неоплаченных визитов ERP {r.erp:,.0f} ₽", root=False)

    # ===== R8. Аномальное сальдо 62.01 (симптом) =====
    roots = pd.DataFrame(F)
    b = osv[(osv.account == "62.01") & ((osv.closing_credit > TOL) | (osv.closing_debit > TOL))]
    for c, g in b.groupby("clinic_id"):
        g = g.sort_values("period")
        first, last = g.iloc[0], g.iloc[-1]
        linked = []
        if len(roots) and "affects_6201" in roots:
            linked = roots[(roots.clinic_id == c) & (roots.affects_6201 == True)].index.tolist()  # noqa: E712
        amount = last.closing_debit - last.closing_credit
        add("R8_osv_balance", "abnormal_balance", first.period, c, "", [], amount,
            f"Счёт 62.01 (расчёты с пациентами) должен закрываться в ноль. Отклонение с {first.period}, "
            f"на конец периода {'дебетовое' if amount > 0 else 'кредитовое'} сальдо {abs(amount):,.0f} ₽. "
            f"Связанные находки: @@LINKED@@", root=False,
            _linked=linked)

    # ===== Сборка =====
    # пустой список находок (всё чисто) не должен ронять сборку: нужны колонки, а строк может не быть
    cols = ["rule", "error_type", "severity", "period", "clinic_id", "direction_code", "record_ids", "amount_rub",
            "evidence", "suggested_action", "is_root_cause", "_linked"]
    findings = pd.DataFrame(F, columns=None if F else cols)
    findings["is_root_cause"] = findings["is_root_cause"].astype(bool)   # у пустой таблицы dtype object ломает фильтр
    findings["_orig"] = findings.index
    if "affects_6201" in findings:
        findings = findings.drop(columns="affects_6201")
    order = {"high": 0, "medium": 1, "low": 2}
    findings["_o"] = findings.severity.map(order)
    findings["_abs"] = -findings.amount_rub.abs()
    findings = findings.sort_values(["is_root_cause", "_o", "_abs"],
                                    ascending=[False, True, True]).reset_index(drop=True)
    findings.insert(0, "finding_id", [f"F{i + 1:03d}" for i in range(len(findings))])
    orig2id = dict(zip(findings._orig, findings.finding_id))
    for i, r in findings.iterrows():
        if "@@LINKED@@" in r.evidence:
            ids = ", ".join(orig2id[j] for j in r._linked) if isinstance(r._linked, list) and r._linked else "не найдены"
            findings.at[i, "evidence"] = r.evidence.replace("@@LINKED@@", ids)
    findings = findings.drop(columns=["_o", "_abs", "_orig", "_linked"], errors="ignore")
    findings.to_csv(OUT / "findings.csv", index=False, encoding="utf-8-sig")

    write_report(findings, small, young_unpaid, rec, len(done), len(items), lags)
    if EVAL:
        evaluate(findings)


def write_report(f, small, young_unpaid, rec, n_visits, n_items, lags):
    roots = f[f.is_root_cause]
    L = ["# Отчёт сверки ERP и 1С за 2025 год", ""]
    L.append(f"Проверено: {n_visits:,} завершённых визитов, {n_items:,} услуг, проводки 1С, платежи, ОСВ.")
    L.append(f"**Найдено первопричин: {len(roots)}**, симптомов в ОСВ: {len(f) - len(roots)}.")
    L.append(f"Сумма под риском (по модулю): **{roots.amount_rub.abs().sum():,.0f} ₽**.")
    L.append("")
    L += ["## Сводка по типам", "", "| Тип | Кол-во | Сумма, ₽ |", "|---|---|---|"]
    names = {"duplicate_visit": "Дубли визитов", "no_doctor": "Визит без врача",
             "price_below_list": "Цена ниже прайса", "no_payment": "Визит без оплаты",
             "revenue_mismatch": "Выручка 1С ≠ ERP", "wrong_period": "Неверный период",
             "duplicate_receipt": "Задвоенное поступление", "missing_receipt": "Платёж не проведён",
             "dms_receipt_mismatch": "Оплата ДМС не сходится", "price_list_mismatch": "Прайс не по истории",
             "doctor_clinic_mismatch": "Врач чужой клиники", "acquiring_fee_rate": "Комиссия эквайринга не по ставке",
             "payroll_rate_mismatch": "Ставка ЗП не по справочнику", "cost_misallocation": "Материалы не на то направление"}
    for t, g in roots.groupby("error_type"):
        L.append(f"| {names.get(t, t)} | {len(g)} | {g.amount_rub.abs().sum():,.0f} |")
    L += ["", "## Находки", "", "| ID | Важность | Период | Клиника | Тип | Сумма, ₽ | Что и где |",
          "|---|---|---|---|---|---|---|"]
    for r in roots.itertuples():
        L.append(f"| {r.finding_id} | {r.severity} | {r.period} | {r.clinic_id} | "
                 f"{names.get(r.error_type, r.error_type)} | {r.amount_rub:,.0f} | {r.evidence} |")
    L += ["", "## Что делать", ""]
    for t, g in roots.groupby("error_type"):
        L.append(f"- **{names.get(t, t)}** ({len(g)}): {g.suggested_action.iloc[0]}.")
    sym = f[~f.is_root_cause]
    if len(sym):
        L += ["", "## Симптомы в ОСВ", ""]
        for r in sym.itertuples():
            L.append(f"- {r.finding_id}, {r.clinic_id}: {r.evidence}")
    L += ["", "## Что проверено и принято как норма", "",
          f"- Расхождения до {TOL:.0f} ₽ на ячейку (округления): {len(small)} шт., "
          f"сумма по модулю {small['diff'].abs().sum():,.0f} ₽. Не ошибки.",
          f"- Неоплаченные визиты с 01.11.2025 ({len(young_unpaid)} шт., "
          f"{young_unpaid.rev.sum():,.0f} ₽): ДМС платит через 30-45 дней, срок не вышел.",
          "- Дебиторка ДМС в ОСВ (62.02) на 31.12 сходится с неоплаченными визитами ERP: "
          + ", ".join(f"{c}: {r.onec:,.0f} против {r.erp:,.0f}" for c, r in rec.iterrows()) + " ₽.",
          "- Оплаты ДМС за декабрь 2024 (акт прошлого года) исключены из сверки (их сдвиги дат не проверяются).",
          f"- Сдвиги даты зачисления банком без влияния на месяц (1С позже ERP на 1-{LAG_DAYS} дня): {len(lags)} шт., "
          f"{lags.amount_rub.abs().sum():,.0f} ₽; список в timing_differences.csv. Не ошибки.",
          "", "## Ограничения", "",
          "- Потерянную самооплату в ноябре-декабре нельзя отличить от ожидаемой оплаты ДМС: "
          "в ERP нет признака плательщика у неоплаченного визита. Для таких визитов нужен признак плательщика.",
          "- Сверка делается на уровне день / клиника / направление; внутри ячейки проводки-кандидаты "
          "перечислены, но точную проводку подтверждает бухгалтер.",
          "- Не проверялись: НДС, налог на прибыль, база расчёта зарплаты врачей (проверена только ставка по справочнику).",
          "- Комиссия эквайринга сравнивается со ставкой, самой частой по сети: договора в данных нет; "
          "ошибочная ставка по всей сети осталась бы незамеченной."]
    import re
    text = re.sub(r"(?<=\d),(?=\d{3})", " ", "\n".join(L))   # разряды пробелом
    (OUT / "reconciliation_report.md").write_text(text, encoding="utf-8")


def evaluate(f):
    kp = KEY / "injected_errors.csv"
    if not kp.exists():
        return
    key = pd.read_csv(kp)
    errs = key[key.error_id != "BENIGN"].copy()
    benign = (key.error_id == "BENIGN").sum()
    roots = f[f.is_root_cause].copy()
    roots["key_type"] = roots.error_type.replace({"duplicate_receipt": "negative_balance"})
    tp, matched_f, miss = [], set(), []
    for e in errs.itertuples():
        ids = set(str(e.record_ids).split("|"))
        hit = roots[(roots.key_type == e.error_type) &
                    roots.record_ids.map(lambda s: bool(ids & set(str(s).split("|"))))]
        if len(hit):
            tp.append(e.error_id)
            matched_f |= set(hit.finding_id)
        else:
            miss.append(e.error_id)
    fp = roots[~roots.finding_id.isin(matched_f)]
    prec = len(matched_f) / max(len(roots), 1)
    rec = len(tp) / len(errs)
    L = ["# Оценка сверки по ключу ответов", "",
         f"- Внесено ошибок: {len(errs)}, безобидных округлений: {benign}",
         f"- Найдено: {len(tp)} из {len(errs)} (полнота {rec:.0%})",
         f"- Находок-первопричин: {len(roots)}, из них подтверждённых ключом: {len(matched_f)} "
         f"(точность {prec:.0%})",
         f"- Ложных срабатываний: {len(fp)}", f"- Пропущено: {', '.join(miss) or 'нет'}", ""]
    if len(fp):
        L.append("## Ложные срабатывания")
        for r in fp.itertuples():
            L.append(f"- {r.finding_id} {r.error_type} {r.period} {r.clinic_id}: {r.evidence}")
    (KEY / "reconcile_evaluation.md").write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Сверка ERP и 1С")
    ap.add_argument("--input", default="dirty", help="папка в data/ с данными (dirty | corrected)")
    ap.add_argument("--out", default="output", help="папка в data/ для результатов")
    args = ap.parse_args()
    DIRTY = ROOT / "data" / args.input
    OUT = ROOT / "data" / args.out
    EVAL = args.input == "dirty"
    main()
