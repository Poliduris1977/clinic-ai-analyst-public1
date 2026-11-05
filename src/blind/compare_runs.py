"""Слепой тест сверки (пп. 84-87): первый прогон (замороженные правила) и повторный (после доработки, не слепой) рядом.
Оба прогона идут по одному и тому же data/dirty_v2 и оцениваются одним и тем же score_blind.py (его логика не менялась).
Итоговое состояние data/output_v2 = повторный прогон. Результаты каждого прогона копируются в *_first.csv / *_second.csv.
Запуск (после чистой сборки цепочки и src/blind/inject_errors_v2.py): python3 src/blind/compare_runs.py"""
import hashlib, shutil, subprocess, sys
from pathlib import Path
import pandas as pd
ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "output_v2"
RUNS = [("first", "reconcile_frozen_59cff04f.py"), ("second", "reconcile.py")]
sh = lambda *a: subprocess.run([sys.executable, *map(str, a)], capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=ROOT)
for tag, script in RUNS:
    path = ROOT / "src" / "reconcile" / script
    r = sh(path, "--input", "dirty_v2", "--out", "output_v2"); assert r.returncode == 0, r.stderr[-500:]
    r = sh(ROOT / "src" / "blind" / "score_blind.py"); assert r.returncode == 0, r.stderr[-500:]
    print(f"[{tag}] {script} sha256 {hashlib.sha256(path.read_bytes()).hexdigest()[:16]}...\n{r.stdout}")
    for n in ["findings", "blind_score", "blind_false_positives"]:
        shutil.copy(OUT / f"{n}.csv", OUT / f"{n}_{tag}.csv")
a, b = pd.read_csv(OUT / "blind_score_first.csv"), pd.read_csv(OUT / "blind_score_second.csv")
m = a.merge(b, on=["id", "type", "subtle", "period"], suffixes=("_1", "_2"))
m["first"] = m.found_1.map({True: "да", False: "НЕТ"}); m["second"] = m.found_2.map({True: "да", False: "НЕТ"})
m[["id", "type", "subtle", "first", "second"]].to_csv(OUT / "blind_compare.csv", index=False)
print(m[["id", "type", "subtle", "first", "second"]].to_string(index=False))
