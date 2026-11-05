"""
Прогноз выручки: тренд с затуханием (метод Хольта) по направлениям, откат и распределение по врачам.

Позиция (согласована 30.10.2026): рост 2025 года считается ростом бизнеса, сезонность по одному году оценить нельзя,
поэтому базовый прогноз без сезонности; сезонная поправка задаётся пользователем вручную (по умолчанию 1,00).

Метод Хольта с затуханием (месяцы 1..n):
  L1 = y1,  B1 = y2 - y1
  Lt = a*yt + (1-a)*(L(t-1) + phi*B(t-1)),  Bt = b*(Lt - L(t-1)) + (1-b)*phi*B(t-1),  t = 2..n
  прогноз на h шагов = Ln + (phi + phi^2 + ... + phi^h) * Bn
Параметры по умолчанию a=0,5, b=0,3, phi=0,9 не подбирались по данным.

Откат: обучение по первым o месяцам, прогноз следующих трёх, o = 6..9 (прогнозируемые месяцы не позже декабря);
погрешность = факт / прогноз - 1. Итог по сети считается как сумма прогнозов направлений (снизу вверх).

Врачи: прогноз направления делится пропорционально среднемесячной выручке врача за последние три месяца
(выручка за эти месяцы / число месяцев с записью в payroll), поэтому сумма врачей равна направлению.
"""
import numpy as np
import pandas as pd

ALPHA, BETA, PHI = 0.5, 0.3, 0.9
ORIGINS = range(6, 10)
HORIZON = 3


def holt(y, h=HORIZON, alpha=ALPHA, beta=BETA, phi=PHI):
    y = np.asarray(y, dtype=float)
    level, trend = y[0], y[1] - y[0]
    for t in range(1, len(y)):
        prev = level
        level = alpha * y[t] + (1 - alpha) * (prev + phi * trend)
        trend = beta * (level - prev) + (1 - beta) * phi * trend
    return np.array([level + sum(phi ** i for i in range(1, k + 1)) * trend for k in range(1, h + 1)])


def naive(y, h=HORIZON):
    return np.repeat(float(y[-1]), h)


def mean3(y, h=HORIZON):
    return np.repeat(float(np.mean(y[-3:])), h)


def linear(y, h=HORIZON, k=6):
    b, a = np.polyfit(np.arange(k), np.asarray(y[-k:], dtype=float), 1)
    return a + b * (k - 1 + np.arange(1, h + 1))


METHODS = {"Последний месяц": naive, "Среднее 3 месяцев": mean3, "Линейный тренд 6 месяцев": linear,
           "Хольт с затуханием": holt}


def backtest(series, method=holt):
    """series: DataFrame (месяцы 1..12 x направления). Возвращает построчный откат; итог = сумма прогнозов направлений."""
    rows = []
    cols = list(series.columns)
    for o in ORIGINS:
        fc = {c: method(series[c].values[:o], HORIZON) for c in cols}
        fc["Итого"] = sum(fc[c] for c in cols)
        act = {c: series[c].values[o:o + HORIZON] for c in cols}
        act["Итого"] = sum(act[c] for c in cols)
        for s in cols + ["Итого"]:
            for h in range(HORIZON):
                if o + h + 1 > 12:
                    continue
                rows.append((s, o, h + 1, fc[s][h], act[s][h]))
    d = pd.DataFrame(rows, columns=["series", "origin", "horizon", "forecast", "actual"])
    d["err"] = d.actual / d.forecast - 1
    return d


def summary(bt):
    return {"mape": float(bt.err.abs().mean()), "bias": float(bt.err.mean()),
            "q10": float(bt.err.quantile(0.10)), "q50": float(bt.err.quantile(0.50)),
            "q90": float(bt.err.quantile(0.90)), "n": int(len(bt))}


def doctor_shares(rev_q, months_q, direction):
    """rev_q, months_q, direction: Series по врачам. Доля в направлении по среднемесячной выручке."""
    rr = (rev_q / months_q.replace(0, np.nan)).fillna(0.0)
    tot = rr.groupby(direction).transform("sum")
    return (rr / tot.replace(0, np.nan)).fillna(0.0), rr
