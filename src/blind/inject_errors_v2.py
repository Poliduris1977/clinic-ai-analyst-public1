"""
Слепой набор ошибок, версия 2: внесение расхождений и ошибок между ERP и 1С.

Берёт чистые файлы из data/raw/, пишет «испорченные» копии в data/dirty_v2/
(те же имена файлов и колонки) и ключ ответов в
data/answer_key_v2/injected_errors_v2.csv.

Сид 33. 14 ошибок (E01-E14) и 12 «безобидных» расхождений (B01-B12), которые
находкой быть не должны.

Типы ошибок (в скобках - где лежит ошибка):
  E01  wrong_direction          часть выручки 1С проведена по чужому направлению (1С)
  E02  wrong_clinic             часть выручки 1С проведена по чужой клинике (1С)
  E03  insurer_mismatch         поступление от страховой отнесено на другую СК (1С)
  E04  payer_type_flip          платёж СК записан как оплата пациента (ERP)
  E05  status_conflict          оплаченный визит с позициями переведён в «отменён» (ERP)
  E06  acquiring_fee_rate       комиссия эквайринга по ставке не 1,8% (1С)      [мелкая]
  E07  missing_receipt_1c       поступление эквайринга не попало в 1С (1С)
  E08  payment_partial          платёж меньше стоимости визита, но не ноль (ERP)  [мелкая]
  E09  doctor_clinic_mismatch   у визита врач другой клиники (ERP)               [редкий срез]
  E10  payroll_rate_mismatch    ставка врача в расчёте ЗП не совпадает со справочником (1С)
  E11  cost_misallocation       списание материалов отнесено на другое направление (1С)
  E12  stale_list_price         позиция пробита по старому прайсу, всё сходится (ERP+1С)  [мелкая]
  E13  revenue_mismatch (старый тип, мелкая сумма): реализация 1С отличается на ~145 ₽
  E14  payment_date_shift       дата платежа в ERP уехала в следующий месяц (ERP)

Безобидные расхождения (НЕ ошибки):
  B01-B02  округление реализации 62.01 на 1-35 ₽
  B03-B04  округление реализации 62.02 на 1-40 ₽
  B05-B06  банк округлил комиссию эквайринга до рубля (< 1 ₽)
  B07-B08  поступление от СК отражено в 1С на день позже (в пределах месяца)
  B09      поступление эквайринга за субботу отражено в понедельник (в пределах месяца)
  B10-B11  отменённая запись и новая запись того же пациента к тому же врачу в тот же день
  B12      запись ФИО в расчёте ЗП отличается от справочника только пробелом

Принципы:
  * правки в 1С делаются на проводках, затем пересчитываются закрытие месяца,
    ОСВ и ДДС (common.ledger), поэтому файлы остаются внутренне согласованными;
  * ERP-правки выполняются в visits / visit_items / visit_payments;
  * расчёт ЗП (payroll_1c) и материалы (materials_1c) пересчитываются вместе
    с проводками там, где ошибка их касается;
  * ID записей в ключе соответствуют ID в испорченных файлах. Если запись
    удалена (E07), ключ - составной: клиника|дата|вид документа.

Запуск:  python src/blind/inject_errors_v2.py
"""
import random
import shutil
import sys
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.ledger import rebuild_closing, rebuild_cash, rebuild_osv  # noqa: E402

SEED = 33
YEAR = 2025
ROOT = Path(__file__).resolve().parents[2]
RAW = ROOT / "data" / "raw"
DIRTY = ROOT / "data" / "dirty_v2"
KEY = ROOT / "data" / "answer_key_v2"
rng = random.Random(SEED)

STR = {"debit": str, "credit": str, "account": str}
ENC = dict(index=False, encoding="utf-8-sig")
FEE_RATE = 0.018
RECEIPT_PAT = "Поступление от пациентов (эквайринг)"
RECEIPT_INS = "Поступление от страховой"
FEE = "Комиссия эквайринга"
REAL = "Реализация услуг"
CLINIC_DIRS = {"C01": ["DENT", "COSM", "PLAS", "FAM", "LASER"],
               "C02": ["DENT", "COSM", "FAM"],
               "C03": ["DENT", "FAM", "LASER"]}


def pick(idx):
    """Детерминированный выбор из индекса/списка."""
    return rng.choice(sorted(idx))


def main():
    DIRTY.mkdir(parents=True, exist_ok=True)
    KEY.mkdir(parents=True, exist_ok=True)

    visits = pd.read_csv(RAW / "visits.csv")
    items = pd.read_csv(RAW / "visit_items.csv")
    pays = pd.read_csv(RAW / "visit_payments.csv")
    doctors = pd.read_csv(RAW / "doctors.csv")
    price_hist = pd.read_csv(RAW / "price_history.csv")
    payroll = pd.read_csv(RAW / "payroll_1c.csv")
    mats = pd.read_csv(RAW / "materials_1c.csv")
    post = pd.read_csv(RAW / "postings_1c.csv", dtype=STR)
    osv_clean = pd.read_csv(RAW / "osv_1c.csv", dtype=STR | {c: float for c in [
        "opening_debit", "opening_credit", "turnover_debit", "turnover_credit",
        "closing_debit", "closing_credit"]})

    post["_tag"] = ""
    oper = post[post.doc_type != "Закрытие месяца"].copy()
    oper["month"] = pd.to_datetime(oper.date).dt.month

    vm = visits.copy()
    vm["month"] = pd.to_datetime(vm.visit_date).dt.month
    vrev = items.groupby("visit_id").revenue.sum()
    vm["rev"] = vm.visit_id.map(vrev).fillna(0)
    self_pay = pays[pays.payer_type == "self"].merge(
        vm[["visit_id", "clinic_id", "direction_code", "visit_date", "month", "rev", "status"]], on="visit_id")

    key_rows = []
    used_posts, used_visits, used_pays = set(), set(), set()

    def note(**kw):
        key_rows.append(kw)

    def receipt_idx(cid, d):
        s = oper[(oper.clinic_id == cid) & (oper.date == d) & (oper.doc_type == RECEIPT_PAT)]
        return s.index[0]

    def fee_idx(cid, d):
        s = oper[(oper.clinic_id == cid) & (oper.date == d) & (oper.doc_type == FEE)]
        return s.index[0]

    # =====================================================================
    # ОШИБКИ
    # =====================================================================

    # ---- E01: часть выручки 1С проведена по чужому направлению -------------
    # (в 1С каждый день есть реализация по каждому направлению клиники, поэтому
    #  ошибка - перенос части суммы между двумя проводками одного дня)
    cand = oper[(oper.clinic_id == "C01") & (oper.month == 5) & (oper.doc_type == REAL) &
                (oper.debit == "62.01") & (oper.amount >= 150_000)]
    i = pick(cand.index)
    d0 = oper.loc[i, "date"]
    tgt = oper[(oper.clinic_id == "C01") & (oper.date == d0) & (oper.doc_type == REAL) & (oper.debit == "62.01") &
               (oper.index != i)]
    j = pick(tgt.index)
    shift = float(round(rng.uniform(0.35, 0.5) * oper.loc[i, "amount"], -2))
    old_dir, new_dir = oper.loc[i, "direction_code"], oper.loc[j, "direction_code"]
    oper.loc[i, "amount"] = round(oper.loc[i, "amount"] - shift, 2)
    oper.loc[j, "amount"] = round(oper.loc[j, "amount"] + shift, 2)
    oper.loc[[i, j], "_tag"] = "E01"
    used_posts.update([i, j])
    note(error_id="E01", is_finding=1, subtle=0, error_type="wrong_direction", file="postings_1c.csv",
         period=f"{YEAR}-05", clinic_id="C01", direction_code=f"{old_dir}->{new_dir}", amount_effect_rub=shift,
         description=f"Часть реализации {d0} ({shift:,.0f} ₽) проведена в 1С по направлению {new_dir} вместо {old_dir}; "
                     f"итог по клинике и по дню сходится с ERP, по двум направлениям расходится в обе стороны")

    # ---- E02: часть выручки 1С проведена по чужой клинике ------------------
    cand = oper[(oper.clinic_id == "C02") & (oper.month == 8) & (oper.doc_type == REAL) &
                (oper.debit == "62.01") & (oper.direction_code.isin(["DENT", "FAM"])) &
                (oper.amount >= 100_000) & (~oper.index.isin(used_posts))]
    i = pick(cand.index)
    d0, dcode = oper.loc[i, "date"], oper.loc[i, "direction_code"]
    tgt = oper[(oper.clinic_id == "C03") & (oper.date == d0) & (oper.direction_code == dcode) &
               (oper.doc_type == REAL) & (oper.debit == "62.01")]
    j = tgt.index[0]
    shift = float(round(rng.uniform(0.3, 0.45) * oper.loc[i, "amount"], -2))
    oper.loc[i, "amount"] = round(oper.loc[i, "amount"] - shift, 2)
    oper.loc[j, "amount"] = round(oper.loc[j, "amount"] + shift, 2)
    oper.loc[[i, j], "_tag"] = "E02"
    used_posts.update([i, j])
    note(error_id="E02", is_finding=1, subtle=0, error_type="wrong_clinic", file="postings_1c.csv",
         period=f"{YEAR}-08", clinic_id="C02->C03", direction_code=dcode, amount_effect_rub=shift,
         description=f"Часть реализации {d0} по направлению {dcode} ({shift:,.0f} ₽) относится к визитам клиники C02, "
                     f"но проведена в 1С по клинике C03; по сети итог сходится, по двум клиникам расходится в обе стороны; в ОСВ у C02 кредитовое сальдо на 62.01, у C03 - дебетовое")

    # ---- E03: поступление от СК отнесено на другую СК ---------------------
    cand = oper[(oper.month == 7) & (oper.doc_type == RECEIPT_INS) & (oper.comment.isna()) &
                (oper.amount.between(300_000, 900_000))]
    insurers = sorted(oper.counterparty.dropna().unique())
    for _ in range(200):
        i = pick(cand.index)
        cid, d, cp = oper.loc[i, "clinic_id"], oper.loc[i, "date"], oper.loc[i, "counterparty"]
        free = [x for x in insurers if x != cp and not ((oper.clinic_id == cid) & (oper.date == d) &
                (oper.doc_type == RECEIPT_INS) & (oper.counterparty == x)).any()]
        if free:
            break
    else:
        raise RuntimeError("E03: нет подходящей проводки")
    new_cp = rng.choice(free)
    oper.loc[i, "counterparty"] = new_cp
    oper.loc[i, "_tag"] = "E03"
    used_posts.add(i)
    note(error_id="E03", is_finding=1, subtle=0, error_type="insurer_mismatch",
         file="postings_1c.csv/cashflow_1c.csv", period=f"{YEAR}-07", clinic_id=cid, direction_code="",
         amount_effect_rub=float(oper.loc[i, "amount"]),
         description=f"Поступление {d} на {oper.loc[i, 'amount']:,.0f} ₽ фактически от {cp}, "
                     f"в 1С (и в ДДС) отнесено на {new_cp}; общая сумма поступлений сходится, "
                     f"в разрезе страховых - нет")

    # ---- E04: платёж СК записан как оплата пациента (ERP) -----------------
    dp = pays[(pays.payer_type == "dms") & (pays.payer_name == "Согласие") &
              (pays.amount.between(2_000, 9_000)) &
              (pays.pay_date.between(f"{YEAR}-06-01", f"{YEAR}-08-31"))]
    r = pays.loc[pick(dp.index)]
    pays.loc[pays.payment_id == r.payment_id, ["payer_type", "payer_name"]] = ["self", "Пациент"]
    used_pays.add(r.payment_id)
    used_visits.add(r.visit_id)
    vv = vm[vm.visit_id == r.visit_id].iloc[0]
    note(error_id="E04", is_finding=1, subtle=1, error_type="payer_type_flip", file="visit_payments.csv",
         record_ids=r.payment_id, period=r.pay_date[:7], clinic_id=vv.clinic_id, direction_code=vv.direction_code,
         amount_effect_rub=float(r.amount),
         description=f"Платёж {r.payment_id} по визиту {r.visit_id} (оплата СК «Согласие» {r.pay_date}) записан в ERP как "
                     f"оплата пациента; в 1С он пришёл как поступление от страховой")

    # ---- E05: оплаченный визит с позициями переведён в «отменён» -----------
    cand = self_pay[(self_pay.clinic_id == "C03") & (self_pay.month == 9) & (self_pay.status == "completed") &
                    (self_pay.rev.between(15_000, 45_000)) & (~self_pay.visit_id.isin(used_visits))]
    r = cand.loc[pick(cand.index)]
    visits.loc[visits.visit_id == r.visit_id, "status"] = "cancelled"
    used_visits.add(r.visit_id)
    note(error_id="E05", is_finding=1, subtle=0, error_type="status_conflict", file="visits.csv",
         record_ids=r.visit_id, period=f"{YEAR}-09", clinic_id="C03", direction_code=r.direction_code,
         amount_effect_rub=float(r.rev),
         description=f"Визит {r.visit_id} от {r.visit_date} имеет статус «cancelled», но по нему есть позиции "
                     f"на {r.rev:,.0f} ₽ и оплата; в 1С выручка и поступление проведены")

    # ---- E06: комиссия эквайринга по ставке не 1,8% ------------------------
    cand = oper[(oper.clinic_id == "C02") & (oper.month.between(1, 6)) & (oper.doc_type == RECEIPT_PAT) &
                (oper.amount.between(150_000, 300_000))]
    ri = pick(cand.index)
    cid, d = oper.loc[ri, "clinic_id"], oper.loc[ri, "date"]
    fi = fee_idx(cid, d)
    old_fee = float(oper.loc[fi, "amount"])
    new_fee = round(float(oper.loc[ri, "amount"]) * 0.019, 2)
    oper.loc[fi, "amount"] = new_fee
    oper.loc[fi, "_tag"] = "E06"
    used_posts.update([ri, fi])
    note(error_id="E06", is_finding=1, subtle=1, error_type="acquiring_fee_rate",
         file="postings_1c.csv/cashflow_1c.csv", period=d[:7], clinic_id=cid, direction_code="",
         amount_effect_rub=round(new_fee - old_fee, 2),
         description=f"Комиссия эквайринга за {d} списана {new_fee:,.2f} ₽ при поступлении {oper.loc[ri, 'amount']:,.0f} ₽: "
                     f"ставка 1,9% вместо договорных 1,8% (должно быть {old_fee:,.2f} ₽)")

    # ---- E07: поступление эквайринга не попало в 1С ------------------------
    cand = oper[(oper.clinic_id == "C03") & (oper.month.between(9, 11)) & (oper.doc_type == RECEIPT_PAT) &
                (oper.amount.between(150_000, 350_000)) & (~oper.index.isin(used_posts))]
    ri = pick(cand.index)
    cid, d = oper.loc[ri, "clinic_id"], oper.loc[ri, "date"]
    fi = fee_idx(cid, d)
    amt, fee_amt = float(oper.loc[ri, "amount"]), float(oper.loc[fi, "amount"])
    raw_ids = f"{post.loc[ri, 'posting_id']}|{post.loc[fi, 'posting_id']}"
    oper = oper.drop(index=[ri, fi])
    used_posts.update([ri, fi])
    note(error_id="E07", is_finding=1, subtle=0, error_type="missing_receipt_1c",
         file="postings_1c.csv/cashflow_1c.csv", record_ids=f"{cid}|{d}|{RECEIPT_PAT}", period=d[:7],
         clinic_id=cid, direction_code="", amount_effect_rub=-amt,
         description=f"Поступление эквайринга {d} на {amt:,.0f} ₽ (и комиссия {fee_amt:,.2f} ₽) отсутствует в 1С: "
                     f"реестр не загружен; платежи в ERP есть, на 62.01 остался непогашенный долг "
                     f"(в чистых данных проводки {raw_ids})")

    # ---- E08: платёж меньше стоимости визита, но не ноль (ERP) -------------
    cand = self_pay[(self_pay.clinic_id == "C01") & (self_pay.month == 4) & (self_pay.status == "completed") &
                    (self_pay.amount.between(15_000, 45_000)) & (self_pay.amount == self_pay.rev) &
                    (~self_pay.visit_id.isin(used_visits))]
    r = cand.loc[pick(cand.index)]
    delta = int(rng.choice([1000, 1500, 2000, 2500]))
    pays.loc[pays.payment_id == r.payment_id, "amount"] = int(r.amount) - delta
    used_visits.add(r.visit_id)
    used_pays.add(r.payment_id)
    note(error_id="E08", is_finding=1, subtle=1, error_type="payment_partial", file="visit_payments.csv",
         record_ids=r.payment_id, period=f"{YEAR}-04", clinic_id="C01", direction_code=r.direction_code,
         amount_effect_rub=-float(delta),
         description=f"Платёж {r.payment_id} по визиту {r.visit_id}: {int(r.amount) - delta:,} ₽ при стоимости "
                     f"{int(r.rev):,} ₽ (недоплата {delta:,} ₽, но не ноль); в 1С поступление на полную сумму")

    # ---- E09: у визита врач другой клиники (ERP) ---------------------------
    c3_fam = doctors[(doctors.clinic_id == "C03") & (doctors.direction_code == "FAM")].doctor_id.tolist()
    cand = vm[(vm.clinic_id == "C01") & (vm.direction_code == "FAM") & (vm.month == 10) &
              (vm.status == "completed") & (vm.rev.between(4_000, 15_000)) & (~vm.visit_id.isin(used_visits))]
    r = cand.loc[pick(cand.index)]
    old_doc = r.doctor_id
    new_doc = rng.choice(sorted(c3_fam))
    visits.loc[visits.visit_id == r.visit_id, "doctor_id"] = new_doc
    used_visits.add(r.visit_id)
    note(error_id="E09", is_finding=1, subtle=1, error_type="doctor_clinic_mismatch", file="visits.csv",
         record_ids=r.visit_id, period=f"{YEAR}-10", clinic_id="C01", direction_code="FAM",
         amount_effect_rub=float(r.rev),
         description=f"Визит {r.visit_id} в клинике C01: врач {new_doc} работает в клинике C03 (был {old_doc}); "
                     f"выручка {r.rev:,.0f} ₽ уходит врачу чужой клиники, расчёт мотивации по врачам расходится с ERP")

    # ---- E10: ставка врача в расчёте ЗП не совпадает со справочником -------
    nxt = {0.28: 0.32, 0.32: 0.36, 0.36: 0.40}
    pr_dec = payroll[(payroll.period == f"{YEAR}-12") & (payroll.role == "doctor")]
    pr_dec = pr_dec.merge(doctors[["doctor_id", "pay_pct"]].rename(columns={"pay_pct": "ref_pct"}),
                          left_on="employee_id", right_on="doctor_id")
    pr_dec = pr_dec[(pr_dec.ref_pct == 0.32) & (pr_dec.base_revenue >= 1_000_000)]
    r = pr_dec.loc[pick(pr_dec.index)]
    old_pct, new_pct = float(r.pay_pct), nxt[float(r.pay_pct)]
    new_gross = int(np.round(r.base_revenue * new_pct))
    new_ndfl = int(np.round(new_gross * 0.13))
    new_net = new_gross - new_ndfl
    new_contr = int(np.round(new_gross * 0.22))
    dg, dn, dc = new_gross - int(r.gross), new_ndfl - int(r.ndfl), new_contr - int(r.contributions)
    m = (payroll.period == r.period) & (payroll.employee_id == r.employee_id)
    payroll.loc[m, ["pay_pct", "gross", "ndfl", "net_pay", "contributions"]] = [new_pct, new_gross, new_ndfl, new_net, new_contr]
    cid, dcode = r.clinic_id, r.direction_code
    # проводки: ЗП врачей (клиника+направление), НДФЛ (клиника), взносы (клиника+направление, врачебная часть)
    dir_doc_contr = int(payroll[(payroll.period == r.period) & (payroll.clinic_id == cid) &
                                (payroll.direction_code == dcode) & (payroll.role == "doctor")].contributions.sum()) - new_contr + int(r.contributions)
    p_sal = oper[(oper.date.str[:7] == r.period) & (oper.clinic_id == cid) & (oper.doc_type == "Начисление ЗП врачей") &
                 (oper.direction_code == dcode)]
    p_ndfl = oper[(oper.date.str[:7] == r.period) & (oper.clinic_id == cid) & (oper.doc_type == "Удержание НДФЛ")]
    p_con = oper[(oper.date.str[:7] == r.period) & (oper.clinic_id == cid) & (oper.debit == "20") & (oper.credit == "69") &
                 (oper.direction_code == dcode) & (oper.amount == float(dir_doc_contr))]
    assert len(p_sal) == len(p_ndfl) == len(p_con) == 1, (len(p_sal), len(p_ndfl), len(p_con))
    oper.loc[p_sal.index[0], "amount"] = round(oper.loc[p_sal.index[0], "amount"] + dg, 2)
    oper.loc[p_ndfl.index[0], "amount"] = round(oper.loc[p_ndfl.index[0], "amount"] + dn, 2)
    oper.loc[p_con.index[0], "amount"] = round(oper.loc[p_con.index[0], "amount"] + dc, 2)
    for ix in (p_sal.index[0], p_ndfl.index[0], p_con.index[0]):
        oper.loc[ix, "_tag"] = "E10"
    note(error_id="E10", is_finding=1, subtle=0, error_type="payroll_rate_mismatch",
         file="payroll_1c.csv/postings_1c.csv", record_ids=f"{r.period}|{r.employee_id}", period=r.period,
         clinic_id=cid, direction_code=dcode, amount_effect_rub=float(dg),
         description=f"В расчёте ЗП за {r.period} врачу {r.employee_id} применена ставка {new_pct:.0%} вместо {old_pct:.0%} "
                     f"из справочника: начислено на {dg:,} ₽ больше (НДФЛ +{dn:,} ₽, взносы +{dc:,} ₽); "
                     f"проводки и расчёт ЗП согласованы между собой, расхождение со справочником врачей")

    # ---- E11: списание материалов отнесено на другое направление -----------
    mm = mats[(mats.period == f"{YEAR}-07") & (mats.clinic_id == "C01")].set_index("direction_code")
    src = mm.material_fact.idxmax()
    dst = rng.choice(sorted(d for d in CLINIC_DIRS["C01"] if d not in (src, "LASER")))
    shift = float(round(rng.uniform(0.15, 0.25) * mm.loc[dst, "material_fact"], -2))
    sel_s = oper[(oper.date.str[:7] == f"{YEAR}-07") & (oper.clinic_id == "C01") & (oper.doc_type == "Списание материалов") &
                 (oper.direction_code == src)]
    sel_d = oper[(oper.date.str[:7] == f"{YEAR}-07") & (oper.clinic_id == "C01") & (oper.doc_type == "Списание материалов") &
                 (oper.direction_code == dst)]
    assert len(sel_s) == len(sel_d) == 1
    oper.loc[sel_s.index[0], "amount"] = round(oper.loc[sel_s.index[0], "amount"] - shift, 2)
    oper.loc[sel_d.index[0], "amount"] = round(oper.loc[sel_d.index[0], "amount"] + shift, 2)
    oper.loc[sel_s.index[0], "_tag"] = "E11"
    oper.loc[sel_d.index[0], "_tag"] = "E11"
    mk = (mats.period == f"{YEAR}-07") & (mats.clinic_id == "C01")
    mats.loc[mk & (mats.direction_code == src), "material_fact"] = round(mm.loc[src, "material_fact"] - shift, 2)
    mats.loc[mk & (mats.direction_code == dst), "material_fact"] = round(mm.loc[dst, "material_fact"] + shift, 2)
    note(error_id="E11", is_finding=1, subtle=0, error_type="cost_misallocation",
         file="postings_1c.csv/materials_1c.csv", period=f"{YEAR}-07", clinic_id="C01", direction_code=f"{src}->{dst}",
         amount_effect_rub=shift,
         description=f"Списание материалов на {shift:,.0f} ₽ за июль проведено по направлению {dst} вместо {src}: "
                     f"итог по клинике сходится, факт по направлениям отклоняется от плана в обе стороны "
                     f"(по {dst} выше плана, по {src} ниже)")

    # ---- E12: позиция пробита по старому прайсу (всё остальное сходится) ----
    stale = {"S001": ("2025-04-01", 1500), "S012": ("2025-04-01", 2000),
             "S033": ("2025-09-01", 4500), "S036": ("2025-09-01", 2500)}
    spf = pays[pays.payer_type == "self"][["payment_id", "visit_id", "pay_date", "amount"]]
    im = (items.merge(vm[["visit_id", "clinic_id", "direction_code", "visit_date", "status"]], on="visit_id")
          .merge(spf, on="visit_id"))
    pool = []
    for sid, (vf, oldp) in stale.items():
        lim = (date.fromisoformat(vf) + timedelta(days=9)).isoformat()
        c = im[(im.service_id == sid) & (im.visit_date.between(vf, lim)) & (im.discount_pct == 0) &
               (im.status == "completed") & (im.pay_date == im.visit_date) & (~im.visit_id.isin(used_visits)) &
               (im.list_price_at_date > oldp)]
        pool.append(c)
    pool = pd.concat(pool)
    r = pool.loc[pick(pool.index)]
    oldp = stale[r.service_id][1]
    delta = (int(r.list_price_at_date) - oldp) * int(r.qty)
    items.loc[items.item_id == r.item_id, ["list_price_at_date", "price_fact", "revenue"]] = [oldp, oldp, oldp * int(r.qty)]
    pays.loc[pays.payment_id == r.payment_id, "amount"] = int(r.amount) - delta
    used_visits.add(r.visit_id)
    used_pays.add(r.payment_id)
    doc_id = vm.loc[vm.visit_id == r.visit_id, "doctor_id"].iloc[0]
    ix = oper[(oper.clinic_id == r.clinic_id) & (oper.date == r.visit_date) & (oper.direction_code == r.direction_code) &
              (oper.doc_type == REAL) & (oper.debit == "62.01")].index
    assert len(ix) == 1
    oper.loc[ix[0], "amount"] = round(oper.loc[ix[0], "amount"] - delta, 2)
    ri = receipt_idx(r.clinic_id, r.visit_date)
    fi = fee_idx(r.clinic_id, r.visit_date)
    oper.loc[ri, "amount"] = round(oper.loc[ri, "amount"] - delta, 2)
    oper.loc[fi, "amount"] = round(oper.loc[ri, "amount"] * FEE_RATE, 2)
    used_posts.update([ix[0], ri, fi])
    note(error_id="E12", is_finding=1, subtle=1, error_type="stale_list_price",
         file="visit_items.csv", record_ids=r.item_id, period=r.visit_date[:7], clinic_id=r.clinic_id,
         direction_code=r.direction_code, amount_effect_rub=-float(delta),
         description=f"Позиция {r.item_id} ({r.service_id}) от {r.visit_date} пробита по старому прайсу {oldp:,} ₽ вместо "
                     f"{int(r.list_price_at_date):,} ₽ из прайса на дату, скидки нет; выручка занижена на {delta:,} ₽, "
                     f"платёж, реализация, поступление и комиссия в 1С скорректированы вместе с ним, поэтому ERP и 1С "
                     f"сходятся; видно только относительно прайса (и как расхождение {delta:,} ₽ между базой ЗП врача "
                     f"{doc_id} в payroll_1c и выручкой ERP)")

    # ---- E13: ручная правка реализации на малую сумму (старый тип) ---------
    cand = oper[(oper.clinic_id == "C02") & (oper.month == 11) & (oper.doc_type == REAL) & (oper.debit == "62.01") &
                (oper.amount > 20_000) & (~oper.index.isin(used_posts))]
    i = pick(cand.index)
    oper.loc[i, "amount"] = round(oper.loc[i, "amount"] + 145, 2)
    oper.loc[i, "_tag"] = "E13"
    used_posts.add(i)
    note(error_id="E13", is_finding=1, subtle=1, error_type="revenue_mismatch", file="postings_1c.csv",
         period=f"{YEAR}-11", clinic_id="C02", direction_code=oper.loc[i, "direction_code"], amount_effect_rub=145.0,
         description=f"Реализация в 1С за {oper.loc[i, 'date']} изменена вручную на +145 ₽ относительно ERP "
                     f"(мелкая сумма, чуть выше уровня округлений)")

    # ---- E14: дата платежа в ERP уехала в следующий месяц ------------------
    month_last = {3: "2025-03-31", 5: "2025-05-31", 6: "2025-06-30", 7: "2025-07-31", 9: "2025-09-30", 10: "2025-10-31"}
    mo = rng.choice(sorted(month_last))
    cand = self_pay[(self_pay.visit_date == month_last[mo]) & (self_pay.pay_date == self_pay.visit_date) &
                    (self_pay.status == "completed") & (self_pay.amount == self_pay.rev) &
                    (self_pay.amount.between(10_000, 60_000)) & (~self_pay.visit_id.isin(used_visits))]
    r = cand.loc[pick(cand.index)]
    new_pd = (date.fromisoformat(r.pay_date) + timedelta(days=1)).isoformat()
    pays.loc[pays.payment_id == r.payment_id, "pay_date"] = new_pd
    used_visits.add(r.visit_id)
    used_pays.add(r.payment_id)
    note(error_id="E14", is_finding=1, subtle=0, error_type="payment_date_shift", file="visit_payments.csv",
         record_ids=r.payment_id, period=r.visit_date[:7], clinic_id=r.clinic_id, direction_code=r.direction_code,
         amount_effect_rub=float(r.amount),
         description=f"Платёж {r.payment_id} по визиту {r.visit_id} от {r.visit_date} в ERP датирован {new_pd}: "
                     f"самоплата не бывает позже дня визита, в 1С поступление в {r.visit_date[:7]}, "
                     f"в ERP попало в следующий месяц")

    # =====================================================================
    # БЕЗОБИДНЫЕ РАСХОЖДЕНИЯ (НЕ ошибки)
    # =====================================================================
    benign = []

    def bnote(bid, btype, file, rid, period, cid, dcode, eff, desc):
        benign.append(dict(error_id=bid, is_finding=0, subtle=0, error_type=btype, file=file, record_ids=rid,
                           period=period, clinic_id=cid, direction_code=dcode, amount_effect_rub=eff,
                           description=desc))

    free_post = lambda: ~oper.index.isin(used_posts)  # noqa: E731

    # B01-B04: округления реализации
    for bid, debit, lim in [("B01", "62.01", 35), ("B02", "62.01", 35), ("B03", "62.02", 40), ("B04", "62.02", 40)]:
        pool = oper[(oper.doc_type == REAL) & (oper.debit == debit) & free_post() & (oper.amount > 3_000) &
                    (oper._tag == "")]
        i = pick(pool.index)
        d = rng.choice([-1, 1]) * rng.randint(1, lim)
        oper.loc[i, "amount"] = round(oper.loc[i, "amount"] + d, 2)
        oper.loc[i, "_tag"] = bid
        used_posts.add(i)
        bnote(bid, "rounding_noise", "postings_1c.csv", None, oper.loc[i, "date"][:7], oper.loc[i, "clinic_id"],
              oper.loc[i, "direction_code"], float(d),
              f"Реализация {oper.loc[i, 'date']} в 1С отличается от ERP на {d:+d} ₽ (округление): НЕ ошибка")

    # B05-B06: комиссия эквайринга округлена до рубля банком
    for bid in ("B05", "B06"):
        pool = oper[(oper.doc_type == FEE) & free_post() & (oper.amount % 1 > 0.1) & (oper.amount > 500)]
        i = pick(pool.index)
        old = float(oper.loc[i, "amount"])
        new = float(round(old))
        oper.loc[i, "amount"] = new
        oper.loc[i, "_tag"] = bid
        used_posts.add(i)
        bnote(bid, "fee_rounding", "postings_1c.csv/cashflow_1c.csv", None, oper.loc[i, "date"][:7],
              oper.loc[i, "clinic_id"], "", round(new - old, 2),
              f"Комиссия эквайринга за {oper.loc[i, 'date']} округлена банком до рубля ({old:,.2f} -> {new:,.0f} ₽): НЕ ошибка")

    # B07-B08: поступление от СК отражено в 1С на день позже (в пределах месяца)
    for bid in ("B07", "B08"):
        pool = oper[(oper.doc_type == RECEIPT_INS) & free_post()]
        pool = pool[pool.date.map(lambda s: date.fromisoformat(s).weekday() in (0, 1, 2, 3, 4) and
                                  (date.fromisoformat(s) + timedelta(days=1)).month == date.fromisoformat(s).month)]
        pool = pool[pool.apply(lambda x: not ((oper.clinic_id == x.clinic_id) & (oper.counterparty == x.counterparty) &
                                              (oper.doc_type == RECEIPT_INS) &
                                              (oper.date == (date.fromisoformat(x.date) + timedelta(days=1)).isoformat())).any(),
                               axis=1)]
        i = pick(pool.index)
        od = oper.loc[i, "date"]
        nd = (date.fromisoformat(od) + timedelta(days=1)).isoformat()
        oper.loc[i, "date"] = nd
        oper.loc[i, "_tag"] = bid
        used_posts.add(i)
        bnote(bid, "posting_lag_1d", "postings_1c.csv/cashflow_1c.csv", None, nd[:7], oper.loc[i, "clinic_id"], "",
              0.0, f"Поступление от СК {oper.loc[i, 'counterparty']} за {od} отражено в 1С {nd} (выписка банка на день "
                   f"позже, тот же месяц): НЕ ошибка")

    # B09: поступление эквайринга за субботу отражено в понедельник (тот же месяц)
    pool = oper[(oper.doc_type == RECEIPT_PAT) & free_post()]
    pool = pool[pool.date.map(lambda s: date.fromisoformat(s).weekday() == 5 and
                              (date.fromisoformat(s) + timedelta(days=2)).month == date.fromisoformat(s).month)]
    i = pick(pool.index)
    od, cid = oper.loc[i, "date"], oper.loc[i, "clinic_id"]
    nd = (date.fromisoformat(od) + timedelta(days=2)).isoformat()
    fi = fee_idx(cid, od)
    oper.loc[[i, fi], "date"] = nd
    oper.loc[[i, fi], "_tag"] = "B09"
    used_posts.update([i, fi])
    bnote("B09", "posting_lag_weekend", "postings_1c.csv/cashflow_1c.csv", None, nd[:7], cid, "", 0.0,
          f"Эквайринг за субботу {od} зачислен банком в понедельник {nd} (поступление и комиссия): НЕ ошибка")

    # B10-B11: отменённая запись + новая запись того же пациента к тому же врачу в тот же день
    next_v = int(vm.visit_id.str[1:].astype(int).max()) + 1
    new_visits = []
    for bid, same_time in [("B10", True), ("B11", False)]:
        pool = vm[(vm.status == "completed") & (vm.rev > 0) & (~vm.visit_id.isin(used_visits)) &
                  (vm.visit_time >= "10:00")]
        r = pool.loc[pick(pool.index)]
        used_visits.add(r.visit_id)
        row = visits[visits.visit_id == r.visit_id].iloc[0].to_dict()
        row["visit_id"] = f"V{next_v:07d}"
        next_v += 1
        row["status"] = "cancelled"
        if not same_time:
            h, mi = map(int, r.visit_time.split(":"))
            tot = h * 60 + mi - rng.choice([45, 60, 90])
            row["visit_time"] = f"{tot // 60:02d}:{tot % 60:02d}"
        new_visits.append(row)
        bnote(bid, "cancel_and_rebook", "visits.csv", row["visit_id"], r.visit_date[:7], r.clinic_id, r.direction_code,
              0.0, f"Отменённая запись {row['visit_id']} и состоявшийся визит {r.visit_id}: тот же пациент, врач и дата "
                   f"({'то же время' if same_time else 'другое время'}); перезапись, не дубль: НЕ ошибка")

    # B12: ФИО в расчёте ЗП отличается от справочника пробелом
    pool = payroll[(payroll.role == "doctor") & (payroll.employee_name.str.count(r"\.") == 2)]
    i = pick(pool.index)
    old_name = payroll.loc[i, "employee_name"]
    new_name = old_name.replace(".", ". ", 1)
    payroll.loc[i, "employee_name"] = new_name
    bnote("B12", "name_format", "payroll_1c.csv", f"{payroll.loc[i, 'period']}|{payroll.loc[i, 'employee_id']}",
          payroll.loc[i, "period"], payroll.loc[i, "clinic_id"], payroll.loc[i, "direction_code"], 0.0,
          f"В расчёте ЗП ФИО записано «{new_name}», в справочнике «{old_name}» (пробел в инициалах, ID тот же): НЕ ошибка")

    # =====================================================================
    # СБОРКА И ЗАПИСЬ
    # =====================================================================
    oper = oper.drop(columns="month")
    closing = rebuild_closing(oper)
    closing["_tag"] = ""
    dirty_post = pd.concat([oper, closing], ignore_index=True)
    dirty_post = dirty_post.sort_values(["date", "clinic_id", "doc_type"], kind="stable").reset_index(drop=True)
    dirty_post["posting_id"] = [f"PST{i:07d}" for i in range(1, len(dirty_post) + 1)]

    dirty_visits = pd.concat([visits, pd.DataFrame(new_visits)], ignore_index=True)

    osv_dirty = rebuild_osv(dirty_post.drop(columns="_tag"), osv_clean)
    chk = osv_dirty.groupby(["period", "clinic_id"])[["closing_debit", "closing_credit"]].sum()
    assert (chk.closing_debit - chk.closing_credit).abs().max() < 1, "ОСВ не сходится"
    cash_dirty = rebuild_cash(dirty_post)

    # id проводок в ключ
    for r in key_rows + benign:
        if r.get("record_ids") is None:
            ids = dirty_post.loc[dirty_post._tag == r["error_id"], "posting_id"].tolist()
            r["record_ids"] = "|".join(ids)
        elif r["error_id"] in ("E01", "E02", "E03", "E06", "E10", "E11", "E13"):
            ids = dirty_post.loc[dirty_post._tag == r["error_id"], "posting_id"].tolist()
            r["record_ids"] = "|".join(ids)

    for name in ["clinics", "doctors", "services", "patients", "price_history",
                 "budget_revenue", "budget_opex"]:
        shutil.copy(RAW / f"{name}.csv", DIRTY / f"{name}.csv")
    dirty_visits.to_csv(DIRTY / "visits.csv", **ENC)
    items.to_csv(DIRTY / "visit_items.csv", **ENC)
    pays.to_csv(DIRTY / "visit_payments.csv", **ENC)
    payroll.to_csv(DIRTY / "payroll_1c.csv", **ENC)
    mats.to_csv(DIRTY / "materials_1c.csv", **ENC)
    dirty_post.drop(columns="_tag")[post.columns.drop("_tag")].to_csv(DIRTY / "postings_1c.csv", **ENC)
    osv_dirty.to_csv(DIRTY / "osv_1c.csv", **ENC)
    cash_dirty.to_csv(DIRTY / "cashflow_1c.csv", **ENC)

    key = pd.DataFrame(key_rows + benign)
    cols = ["error_id", "is_finding", "subtle", "error_type", "file", "record_ids", "period", "clinic_id",
            "direction_code", "amount_effect_rub", "description"]
    key[cols].to_csv(KEY / "injected_errors_v2.csv", **ENC)

    errs = key[key.is_finding == 1]
    print(f"Внесено ошибок: {len(errs)} (мелких/сложных: {int(errs.subtle.sum())}), "
          f"безобидных расхождений: {int((key.is_finding == 0).sum())}")
    print(errs.error_type.value_counts().to_string())


if __name__ == "__main__":
    main()
