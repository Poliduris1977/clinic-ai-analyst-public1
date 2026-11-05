"""
Схема оплаты врачей «что если»: чистая логика расчёта без Excel.

Используется как эталон для проверки формул листа «Калькулятор» (build_doctors.py) и как функция для агента.

Вход data: строки «врач x месяц» с колонками grade, worked, revenue, materials, labs, pay_gross, pay_contrib
(лист «Данные» книги Doctors_2025.xlsx). Параметры params (деньги в рублях):
  rates         ставка оплаты по грейдам, доля выручки
  threshold     порог месячной выручки врача, ₽
  bonus         надбавка к ставке на выручку сверх порога, доля
  fixed         фиксированная часть на врача-месяц, ₽
  growth        сценарное изменение выручки у врач-месяцев выше порога, доля (допущение, не прогноз)
  contrib_rate  ставка страховых взносов

Правила:
  оплата = ROUND(ставка x выручка + надбавка x max(0, выручка - порог) + фикс), только для врач-месяцев с записью в payroll;
  выручка врач-месяца выше порога меняется на growth целиком, материалы и лаборатории пропорционально;
  взносы = ROUND(оплата x ставка взносов); вклад = выручка - материалы - лаборатории - оплата - взносы.
При исходных параметрах расчёт равен факту payroll до рубля.
"""
import numpy as np
import pandas as pd

GRADES = ["junior", "middle", "senior", "lead"]


def default_params(rates):
    return {"rates": dict(rates), "threshold": 3_000_000.0, "bonus": 0.0, "fixed": 0.0, "growth": 0.0,
            "contrib_rate": 0.22}


def _round(x):
    # как ROUND в Excel: половина от нуля (np.round округляет к чётному)
    x = np.asarray(x, dtype=float)
    return np.sign(x) * np.floor(np.abs(x) + 0.5)


def apply(data, params):
    d = data.copy()
    worked = d.worked == 1
    over = worked & (d.revenue > params["threshold"])
    d["rev_s"] = d.revenue * (1 + np.where(over, params["growth"], 0.0))
    cost = d.materials + d.labs
    d["cost_s"] = np.where(d.revenue > 0, cost * d.rev_s / d.revenue.replace(0, 1), 0.0)
    rate = d.grade.map(params["rates"]).fillna(0.0)
    raw = rate * d.rev_s + params["bonus"] * np.maximum(0.0, d.rev_s - params["threshold"]) + params["fixed"]
    d["pay_s"] = np.where(worked, _round(raw), 0.0)
    d["contr_s"] = _round(d.pay_s * params["contrib_rate"])
    d["contrib_s"] = d.rev_s - d.cost_s - d.pay_s - d.contr_s              # вклад по сценарию
    d["contrib_f"] = d.revenue - cost - d.pay_gross - d.pay_contrib         # вклад по факту
    return d


def by_grade(d):
    g = d[d.grade.isin(GRADES)].groupby("grade").agg(
        months=("worked", "sum"), rev_f=("revenue", "sum"), rev_s=("rev_s", "sum"),
        pay_f=("pay_gross", "sum"), pay_s=("pay_s", "sum"), contrib_f=("contrib_f", "sum"),
        contrib_s=("contrib_s", "sum"))
    return g.reindex(GRADES)
