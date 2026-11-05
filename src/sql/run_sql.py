"""
Запускает SQL-запросы из папки sql/ на базах corrected и dirty и сверяет результаты с контрольными цифрами проекта.

  python src/sql/run_sql.py

На исправленных данных (corrected) значения должны совпасть с книгами и таблицей контролей в docs/START.md.
На «грязных» данных (dirty) проверяются значения, которые показывают, что запросы находят ошибки: сверка ERP и 1С
находит существенные расхождения, часть визитов без врача и т.д. Код выхода 1, если есть FAIL.
"""
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_db import ROOT, build  # noqa: E402

SQL = ROOT / "sql"
# файл: {колонка: (ожидаемое, допуск)}; текст сравнивается точно (допуск None)
CORRECTED = {
    "01_revenue_erp.sql": {"revenue_erp": (825_211_700, 0.5), "items": (81_532, 0)},
    "02_revenue_1c.sql": {"revenue_1c": (825_211_796, 0.5)},
    "03_reconcile_erp_1c.sql": {"cells": (36, 0), "cells_over_1000": (0, 0), "total_diff": (-96, 0.5)},
    "04_profit_before_tax.sql": {"revenue": (825_211_796, 0.5), "cost_of_sales": (557_108_306.27, 1),
                                 "profit_before_tax": (158_046_155.89, 1)},
    "05_cash_end_of_year.sql": {"opening_51": (31_000_000, 0.5), "net_flow_dds": (185_038_096.87, 1),
                                "closing_osv": (216_038_096.87, 1), "net_flow_postings": (185_038_096.87, 1)},
    "06_min_daily_balance.sql": {"min_balance": (16_579_356, 1), "min_date": ("2025-01-10", None),
                                 "end_balance": (216_038_096.87, 1), "days": (365, 0)},
    "07_stress_insurer_delay.sql": {"min_stress": (16_579_356, 1), "end_stress": (204_579_396.87, 1),
                                    "min_fact": (16_579_356, 1)},
    "08_new_patients_ltv.sql": {"new_patients": (27_924, 0), "rev_new": (494_762_800, 0.5), "jan_patients": (1_287, 0),
                                "ltv12_jan": (34_101.48, 0.01)},
    "09_doctors_coverage.sql": {"doctors": (43, 0), "doctor_months": (479, 0), "visits_no_doctor": (0, 0),
                                "revenue_by_doctors": (825_211_700, 0.5)},
}
DIRTY = {  # запросы должны ловить дефекты исходных данных
    "03_reconcile_erp_1c.sql": {"cells_over_1000": (10, 0), "total_diff": (-67_696, 0.5)},
    "09_doctors_coverage.sql": {"visits_no_doctor": (2, 0)},
}


def run(db: Path, expected: dict, label: str) -> int:
    con = sqlite3.connect(db)
    fails = 0
    for name, cols in expected.items():
        cur = con.execute((SQL / name).read_text(encoding="utf-8"))
        row = dict(zip([c[0] for c in cur.description], cur.fetchone()))
        for col, (want, tol) in cols.items():
            got = row[col]
            ok = got == want if tol is None else got is not None and abs(float(got) - float(want)) <= tol
            fails += not ok
            print(f"{'PASS' if ok else 'FAIL'}  [{label}] {name} / {col}: {got} (ожидалось {want})")
    con.close()
    return fails


def main():
    build("corrected")
    build("dirty")
    fails = run(ROOT / "data" / "sql" / "corrected.db", CORRECTED, "corrected")
    fails += run(ROOT / "data" / "sql" / "dirty.db", DIRTY, "dirty")
    total = sum(len(v) for v in CORRECTED.values()) + sum(len(v) for v in DIRTY.values())
    print(f"\nИтого: {total - fails} из {total} SQL-контролей пройдено")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
