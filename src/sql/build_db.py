"""
Собирает SQLite-базу из CSV одного набора данных проекта.

  python src/sql/build_db.py corrected   ->  data/sql/corrected.db
  python src/sql/build_db.py dirty       ->  data/sql/dirty.db

Таблицы названы по файлам (visits, visit_items, postings_1c, osv_1c, cashflow_1c и т.д.).
Номера счетов (debit, credit, account) хранятся как текст: иначе 90.01 превратится в число 90.01 и потеряет смысл.
Деньги хранятся числами, даты текстом ISO (YYYY-MM-DD), как в источнике.
"""
import sqlite3
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
TEXT_COLS = {"debit", "credit", "account", "article", "doctor_id", "patient_id", "visit_id", "clinic_id"}
INDEXES = ["CREATE INDEX IF NOT EXISTS ix_visits_date ON visits(visit_date)",
           "CREATE INDEX IF NOT EXISTS ix_visits_patient ON visits(patient_id)",
           "CREATE INDEX IF NOT EXISTS ix_items_visit ON visit_items(visit_id)",
           "CREATE INDEX IF NOT EXISTS ix_post_date ON postings_1c(date)",
           "CREATE INDEX IF NOT EXISTS ix_cash_date ON cashflow_1c(date)"]


def build(dataset: str = "corrected") -> Path:
    src = ROOT / "data" / dataset
    out = ROOT / "data" / "sql" / f"{dataset}.db"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.unlink(missing_ok=True)
    con = sqlite3.connect(out)
    for f in sorted(src.glob("*.csv")):
        head = pd.read_csv(f, nrows=0).columns
        df = pd.read_csv(f, dtype={c: str for c in head if c in TEXT_COLS})
        df.to_sql(f.stem, con, index=False)
    for q in INDEXES:
        con.execute(q)
    con.commit()
    con.close()
    return out


if __name__ == "__main__":
    ds = sys.argv[1] if len(sys.argv) > 1 else "corrected"
    print(build(ds))
