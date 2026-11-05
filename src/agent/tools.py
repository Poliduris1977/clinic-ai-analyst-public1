"""
Функции доступа к данным для ИИ-агента (п. 29 DECISIONS): агент сам ничего не считает,
все числа берутся отсюда. Логика переиспользует prepare()/effects() отчётных скриптов,
поэтому совпадает с книгами. Каждая функция возвращает dict с числами (рубли, доли в долях).
"""
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
SRC = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(SRC), str(SRC / "reports")]
import build_cashflow as bcf  # noqa: E402
import build_doctors as bd  # noqa: E402
import build_factor as bf  # noqa: E402
import build_forecast as bfc  # noqa: E402
import build_planfact as bpf  # noqa: E402
from common import forecast as fc  # noqa: E402

ROOT = SRC.parent
DIRS = bpf.DIRS
COST = [a for a in bpf.COST_ARTICLES + bpf.MGMT_ARTICLES + bpf.COMM_ARTICLES]
_C = {}


def _get(key, fn):
    if key not in _C:
        _C[key] = fn()
    return _C[key]


def _sel(df, month, clinic, direction):
    m = pd.Series(True, index=df.index)
    if month:
        m &= df.period == f"2025-{int(month):02d}"
    if clinic:
        m &= df.clinic_id == clinic
    if direction:
        m &= df.direction_code == direction
    return df[m]


def plan_gap(month=None, clinic=None, direction=None):
    """Вопрос 1: отклонение прибыли от плана и из чего оно складывается."""
    pf = _sel(_get("pf", lambda: bpf.prepare()[0]), month, clinic, direction)
    s = pf.groupby("article")[["plan", "fact"]].sum()
    if "Выручка" not in s.index:
        return None
    cost_arts = [a for a in s.index if a != "Выручка"]
    prof_p = s.loc["Выручка", "plan"] - s.loc[cost_arts, "plan"].sum()
    prof_f = s.loc["Выручка", "fact"] - s.loc[cost_arts, "fact"].sum()
    e = _sel(_get("eff", lambda: bf.effects(bf.prepare()[0])), month, clinic, direction)
    fac = {k: float(e[c].sum()) for k, c in (("volume", "volume"), ("mix", "mix"), ("price_list", "price_list"),
                                              ("discount", "discount"), ("materials", "materials"), ("labs", "labs"))}
    dm_dev = float(e.fact_margin.sum() - e.plan_margin.sum())
    other = (s.loc[cost_arts, "plan"] - s.loc[cost_arts, "fact"]).drop(["Материалы", "Лаборатории и внешние услуги"],
                                                                      errors="ignore")
    other = other.sort_values(ascending=False)
    # худшие срезы
    by_dir = pf.assign(sign=np.where(pf.article == "Выручка", 1.0, -1.0)).assign(
        dev=lambda d: (d.fact - d.plan) * d.sign).groupby("direction_code").dev.sum().drop("ОБЩ", errors="ignore").sort_values()
    by_mon = pf.assign(sign=np.where(pf.article == "Выручка", 1.0, -1.0)).assign(
        dev=lambda d: (d.fact - d.plan) * d.sign).groupby("month").dev.sum().sort_values()
    return {"profit_plan": float(prof_p), "profit_fact": float(prof_f), "profit_dev": float(prof_f - prof_p),
            "rev_plan": float(s.loc["Выручка", "plan"]), "rev_fact": float(s.loc["Выручка", "fact"]),
            "dm_dev": dm_dev, "factors": fac, "other_dev": float(prof_f - prof_p - dm_dev),
            "other_top": {k: float(v) for k, v in other.head(3).items()},
            "worst_dir": (by_dir.index[0], float(by_dir.iloc[0])), "worst_month": (int(by_mon.index[0]), float(by_mon.iloc[0])),
            "check": float(prof_f - prof_p - dm_dev - (other.sum()))}


def discounts():
    """Вопрос 3: скидки, объём и проверка связи скидок с потоком пациентов (по ячейкам месяц-клиника-направление)."""
    d = _get("fd", lambda: bf.prepare()[0])
    e = _get("eff", lambda: bf.effects(d))
    g = e.groupby(["period", "clinic_id", "direction_code"]).agg(
        disc=("discount", "sum"), lst=("list_total", "sum"), fq=("q_f_pp", "sum"), pq=("q_p_pp", "sum")).reset_index()
    g["rate"] = -g.disc / g.lst
    g["vol"] = g.fq / g.pq - 1
    corr = float(g.rate.corr(g.vol))
    bd_ = e.groupby("direction_code").agg(disc=("discount", "sum"), lst=("list_total", "sum"), vol=("volume", "sum"))
    return {"discount": float(e.discount.sum()), "volume": float(e.volume.sum()), "list_total": float(e.list_total.sum()),
            "rate": float(-e.discount.sum() / e.list_total.sum()), "corr": corr, "cells": int(len(g)),
            "by_dir": {k: {"disc": float(v.disc), "rate": float(-v.disc / v.lst), "vol": float(v.vol)} for k, v in bd_.iterrows()},
            "margin_dev": float(e.fact_margin.sum() - e.plan_margin.sum())}


def liquidity(delay=30, clinic=None, reserve=10):
    """Вопрос 2: остаток по дням в факте и при задержке страховых на delay дней (как лист «Ликвидность»)."""
    m, daily, open51, ctrl = _get("cf", bcf.prepare)
    cl = [clinic] if clinic else list(open51)
    dd = daily[daily.clinic_id.isin(cl)].groupby("d")[["inflow", "outflow", "ins"]].sum().sort_index()
    bal = sum(open51[c] for c in cl) + (dd.inflow + dd.outflow).cumsum()
    delay = int(delay)
    win = dd.ins.rolling(delay, min_periods=1).sum() if delay > 0 else dd.ins * 0
    stress = bal - win
    mon = dd.index.month
    rows = []
    for k in range(1, 13):
        mk = mon == k
        avg_out = -dd.outflow[mk].sum() / mk.sum()
        mn, sm = float(bal[mk].min()), float(stress[mk].min())
        rows.append({"month": k, "min": mn, "stress_min": sm, "avg_out": float(avg_out),
                     "days_min": mn / avg_out, "days_stress": sm / avg_out})
    st = lambda v, d: "разрыв" if v < 0 else ("ниже резерва" if d < reserve else "ок")  # noqa: E731
    bad_f = [r["month"] for r in rows if st(r["min"], r["days_min"]) != "ок"]
    bad_s = [r["month"] for r in rows if st(r["stress_min"], r["days_stress"]) != "ок"]
    gap_s = [r["month"] for r in rows if r["stress_min"] < 0]
    i_s = stress.idxmin()
    return {"delay": delay, "reserve": reserve, "opening": float(sum(open51[c] for c in cl)),
            "min_fact": float(bal.min()), "min_fact_date": bal.idxmin().strftime("%d.%m"),
            "min_stress": float(stress.min()), "min_stress_date": i_s.strftime("%d.%m"),
            "days_stress_min": min(r["days_stress"] for r in rows), "bad_fact": bad_f, "bad_stress": bad_s,
            "gap_stress": gap_s, "end_fact": float(bal.iloc[-1]), "ins_year": float(dd.ins.sum())}


def doctors(doctor=None, min_months=9):
    """Вопрос 4: вклад врача (прямая маржа - оплата - взносы) на врач-месяц, как лист «Врачи»."""
    dd, dl, _, _ = _get("dd", bd.prepare)
    d = dd[dd.doctor_id != bd.NO_DOC].copy()
    d["direct"] = d.revenue - d.materials - d.labs
    d["contrib"] = d.direct - d.pay_gross - d.pay_contrib
    g = d.groupby("doctor_id").agg(rev=("revenue", "sum"), direct=("direct", "sum"), contrib=("contrib", "sum"),
                                  months=("worked", "sum")).join(dl.set_index("doctor_id"))
    g["per_month"] = g.contrib / g.months
    dm = g.groupby("direction_code").apply(lambda x: x.contrib.sum() / x.months.sum())
    g["vs_dir"] = g.per_month / g.direction_code.map(dm) - 1
    log = pd.read_csv(ROOT / "data" / "output" / "corrections_log.csv")
    fnd = pd.read_csv(ROOT / "data" / "output" / "findings.csv")
    linked = log[log.error_type == "no_doctor"].set_index("after")[["target", "basis"]]
    linked["amount"] = [float(fnd[fnd.record_ids.astype(str).str.contains(v, regex=False)].amount_rub.iloc[0])
                        for v in linked.target]
    g["linked_visit"] = g.index.isin(linked.index)
    low = g.sort_values("per_month").head(5)
    pack = lambda r, i: {"id": i, "name": r.doctor_name, "dir": r.direction_code, "grade": r.grade,  # noqa: E731
                         "months": int(r.months), "per_month": float(r.per_month), "contrib": float(r.contrib),
                         "vs_dir": float(r.vs_dir), "partial": bool(r.months < min_months),
                         "linked_visit": bool(r.linked_visit), "pay_pct": float(r.pay_pct)}
    out = {"low": [pack(r, i) for i, r in low.iterrows()], "n_doctors": int(len(g)),
           "net_per_month": float(g.contrib.sum() / g.months.sum()), "min_months": min_months,
           "linked": {k: {"visit": v.target, "amount": float(v.amount)} for k, v in linked.iterrows()}}
    if doctor:
        key = doctor if doctor in g.index else (g.index[g.doctor_name.str.contains(doctor, regex=False)][0]
                                                if g.doctor_name.str.contains(doctor, regex=False).any() else None)
        out["doctor"] = pack(g.loc[key], key) if key else None
    return out


def forecast_q1():
    """Вопрос 5: прогноз I кв. 2026 (Хольт), коридор и надёжность по направлениям."""
    S, rev, wk, dl, bt, methods, by_series, ctrl = _get("fc", bfc.prepare)
    res = {}
    for c in bfc.CODES:
        f = float(fc.holt(S[c].values).sum())
        err = by_series[c]["mape"]
        rel = "высокая" if err < 0.15 else ("средняя" if err <= 0.30 else "низкая")
        res[c] = {"fc": f, "err": err, "lo": f * (1 - err), "hi": f * (1 + err), "rel": rel,
                  "q4": float(S[c].values[-3:].sum())}
    tot = sum(v["fc"] for v in res.values())
    err = by_series["Итого"]["mape"]
    res["Итого"] = {"fc": tot, "err": err, "lo": tot * (1 - err), "hi": tot * (1 + err),
                    "rel": "высокая" if err < 0.15 else ("средняя" if err <= 0.30 else "низкая"),
                    "q4": float(S.values[-3:].sum())}
    return {"dirs": res, "bias": by_series["Итого"]["bias"], "rev_2025": ctrl["rev_total"],
            "n_backtest": int(by_series["Итого"]["n"])}


def reconciliation():
    """Вопрос 6: итоги сверки ERP и 1С по журналам (файлы с ответами не читаются)."""
    out = ROOT / "data" / "output"
    f = pd.read_csv(out / "findings.csv")
    c = pd.read_csv(out / "corrections_log.csv")
    f2 = pd.read_csv(ROOT / "data" / "output_corrected" / "findings.csv")
    it = pd.read_csv(ROOT / "data" / "corrected" / "visit_items.csv")
    vs = pd.read_csv(ROOT / "data" / "corrected" / "visits.csv")
    ok = it.merge(vs[["visit_id", "status"]], on="visit_id")
    erp = float(ok[ok.status == "completed"].revenue.sum())
    pnl = float(_get("pf", lambda: bpf.prepare()[0]).query("article=='Выручка'").fact.sum())
    byt = c.groupby("error_type").size().to_dict()
    rv = c[c.status == "applied_review"].finding_id
    review_abs = float(f[f.finding_id.isin(rv)].amount_rub.abs().sum())
    return {"findings": int(len(f)), "root": int(f.is_root_cause.sum()), "derived": int((~f.is_root_cause).sum()),
            "applied": int((c.status == "applied").sum()), "review": int((c.status == "applied_review").sum()),
            "after": int(len(f2)), "erp": erp, "books": pnl, "diff": erp - pnl, "by_type": byt, "review_abs": review_abs}


def report_variants():
    """П. 34: значения показателя в разных версиях, определениях и охватах для сопоставления с цифрой отчёта.
    Ключ «месяц|направление|клиника» (0 и пустые строки = все). Значения: выручка [1С, ERP, ERP до исправлений, план],
    прибыль [факт, план], маржа [факт, план]. «До исправлений» = data/dirty (п. 81)."""
    pf = _get("pf", lambda: bpf.prepare()[0]).copy()
    sg = np.where(pf.article == "Выручка", 1.0, -1.0)
    pf["f"], pf["p"] = pf.fact * sg, pf.plan * sg
    rv = pf[pf.article == "Выручка"]
    e = _get("eff", lambda: bf.effects(bf.prepare()[0])).copy()
    e["month"] = e.period.str[5:7].astype(int)

    def erp(folder):
        v = pd.read_csv(ROOT / "data" / folder / "visits.csv")
        i = pd.read_csv(ROOT / "data" / folder / "visit_items.csv")
        m = i.merge(v, on="visit_id")
        m = m[m.status == "completed"].copy()
        m["month"] = m.visit_date.str[5:7].astype(int)
        return m
    ec, ed = erp("corrected"), erp("dirty")

    def cut(df, mo, di, cl):
        m = pd.Series(True, index=df.index)
        if mo:
            m &= df.month == mo
        if di:
            m &= df.direction_code == di
        if cl:
            m &= df.clinic_id == cl
        return df[m]
    out = {}
    for mo in range(13):
        for di in [""] + list(DIRS):
            for cl in ["", "C01", "C02", "C03"]:
                r = cut(rv, mo, di, cl)
                if r.empty:
                    continue
                a = cut(pf, mo, di, cl)
                x = cut(e, mo, di, cl)
                out[f"{mo}|{di}|{cl}"] = [round(float(v), 2) for v in (
                    r.fact.sum(), cut(ec, mo, di, cl).revenue.sum(), cut(ed, mo, di, cl).revenue.sum(), r.plan.sum(),
                    a.f.sum(), a.p.sum(), x.fact_margin.sum(), x.plan_margin.sum())]
                if di:  # прибыль по направлению не определена: общие затраты не распределены (код «ОБЩ»)
                    out[f"{mo}|{di}|{cl}"][4] = out[f"{mo}|{di}|{cl}"][5] = None
    return out
