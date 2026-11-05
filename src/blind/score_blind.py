"""Слепой тест сверки (п. 84): сравнение findings.csv с ключом второго набора.
Находка считается попаданием в ошибку, если совпали период и клиника (любая из
клиник ключа) и либо пересеклись record_ids, либо сумма совпала с точностью до 1 руб.
Исключение: для ошибок сдвига даты допускается соседний месяц (сдвиг уходит в другой период).
Находки-симптомы (is_root_cause=False) не считаются ни попаданиями, ни ложными.
"""
import sys, re, pandas as pd
from pathlib import Path
D = Path(__file__).resolve().parents[2] / "data"
key = pd.read_csv(D/"answer_key_v2/injected_errors_v2.csv")
f = pd.read_csv(D/"output_v2/findings.csv")
ids = lambda s: set(str(s).split("|")) if pd.notna(s) else set()
def clinics(r):  # E02: "C02->C03" лежит в колонке direction_code при сдвиге столбцов
    return set(re.findall(r"C\d\d", f"{r.clinic_id} {r.direction_code}"))
def amt(r):
    a = r.amount_effect_rub
    return abs(float(a)) if pd.notna(a) else None
def month_gap(a, b):
    return abs((int(a[:4]) - int(b[:4])) * 12 + int(a[5:]) - int(b[5:]))
def match(e, x):
    # ошибки со сдвигом даты (payment_date_shift) по замыслу видны в соседнем месяце
    gap = 1 if "date_shift" in e.error_type else 0
    if month_gap(x.period, e.period) > gap or x.clinic_id not in clinics(e): return False
    ov = bool(ids(e.record_ids) & ids(x.record_ids))
    am = amt(e) is not None and abs(abs(x.amount_rub) - amt(e)) <= 1
    return ov or am
rc = f[f.is_root_cause]
err = key[key.is_finding == 1]; ben = key[key.is_finding == 0]
rows = []; used = set()
for e in err.itertuples():
    m = [x.finding_id for x in rc.itertuples() if match(e, x)]
    used |= set(m)
    rows.append(dict(id=e.error_id, type=e.error_type, subtle=e.subtle, period=e.period, found=bool(m), findings="|".join(m)))
res = pd.DataFrame(rows)
fp = rc[~rc.finding_id.isin(used)]
def benign_of(x):
    for b in ben.itertuples():
        if x.period == b.period and x.clinic_id == b.clinic_id: return b.error_id
    return ""
fpt = fp.assign(benign=[benign_of(x) for x in fp.itertuples()])
res.to_csv(D/"output_v2/blind_score.csv", index=False)
fpt[["finding_id","rule","period","clinic_id","amount_rub","benign"]].to_csv(D/"output_v2/blind_false_positives.csv", index=False)
n=len(res); h=int(res.found.sum()); s=res[res.subtle==1]; ns=res[res.subtle==0]
bn=len(ben); bfail=len({b for b in fpt.benign if b})
print(f"ошибок {n}; найдено {h} ({h/n:.0%}); тонких {int(s.found.sum())}/{len(s)}; обычных {int(ns.found.sum())}/{len(ns)}")
print(f"находок-первопричин {len(rc)}; верных {len(used)}; ложных {len(fp)}; точность {len(used)/len(rc):.0%}")
print(f"безобидных {bn}; прошли без находки {bn-bfail}; сработали {bfail}: {sorted(b for b in set(fpt.benign) if b)}")
print("пропущено:", ", ".join(res[~res.found].id + " " + res[~res.found].type))
