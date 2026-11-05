"""
Автоматическая проверка контрольных цифр проекта (таблица в docs/START.md).

Запуск из корня проекта после прогона всей цепочки:  python3 src/check_controls.py
Печатает PASS/FAIL по каждому контролю, в конце итог; код выхода 1, если есть FAIL.
Ничего не пересчитывает и не меняет: только читает data/ и data/reports/.
При осознанном изменении метода обновить ожидаемые значения здесь и в START.md.
"""
import sys
from pathlib import Path

import openpyxl
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
R = ROOT / "data" / "reports"
BOOKS = ["PnL_2025", "PlanFact_2025", "FactorAnalysis_2025", "Doctors_2025", "CashFlow_2025", "PatientLTV_2025",
         "Forecast_2026Q1"]
results = []
_cache = {}


def book(name):
    if name not in _cache:
        _cache[name] = openpyxl.load_workbook(R / f"{name}.xlsx", data_only=True)
    return _cache[name]


def val(name, sheet, cell):
    return book(name)[sheet][cell].value


def check(label, got, want, tol=0.0):
    ok = got is not None and abs(float(got) - float(want)) <= tol
    results.append((ok, label, got, want))


def main():
    try:
        corr = ROOT / "data" / "corrected"
        items = pd.read_csv(corr / "visit_items.csv")
        visits = pd.read_csv(corr / "visits.csv")
        check("Выручка ERP после исправлений, ₽", items.revenue.sum(), 825_211_700, 0.5)
        check("Визитов без врача после исправлений", int(visits.doctor_id.isna().sum()), 0)

        log = pd.read_csv(ROOT / "data" / "output" / "corrections_log.csv")
        vc = log.status.value_counts().to_dict()
        check("Исправлений применено", vc.get("applied", 0) + vc.get("applied_review", 0), 14)
        check("Из них с пометкой «подтвердить»", vc.get("applied_review", 0), 5)
        check("Ручной разбор", vc.get("manual_review", 0), 0)
        f1 = pd.read_csv(ROOT / "data" / "output" / "findings.csv")
        check("Первопричин в первой сверке", int(f1.is_root_cause.sum()), 14)
        f2 = pd.read_csv(ROOT / "data" / "output_corrected" / "findings.csv")
        check("Находок в повторной сверке", len(f2), 0)

        check("P&L: выручка (1С), ₽", val("PnL_2025", "По клиникам", "E5"), 825_211_796, 1)
        check("P&L: прибыль до налога, ₽", val("PnL_2025", "По клиникам", "E24"), 158_046_155.89, 1)
        pf = "План-факт"
        check("План-факт: план прибыли, ₽", val("PlanFact_2025", pf, "B25"), 179_563_350.58, 1)
        check("План-факт: факт прибыли, ₽", val("PlanFact_2025", pf, "C25"), 158_046_155.89, 1)
        check("План-факт: отклонение, ₽", val("PlanFact_2025", pf, "D25"), -21_517_194.69, 1)
        check("Мост: прямая маржа план, ₽", val("FactorAnalysis_2025", "Мост", "B19"), 727_463_146.11, 1)
        check("Мост: прямая маржа факт, ₽", val("FactorAnalysis_2025", "Мост", "C19"), 690_504_466.67, 1)
        check("Валовая маржа план, ₽", val("FactorAnalysis_2025", "Валовая маржа", "B29"), 291_428_430.92, 1)
        check("Валовая маржа факт, ₽", val("FactorAnalysis_2025", "Валовая маржа", "C29"), 268_103_392.11, 1)

        dw = book("Doctors_2025")["Врачи"]
        tot = next(r for r in range(6, 70) if dw.cell(r, 1).value == "Врачи итого")
        check("Врачи: вклад врачей, ₽", dw.cell(tot, 16).value, 361_131_371.65, 1)
        for r, want in zip(range(19, 23), (613.2, 764.9, 826.5, 852.6)):
            check(f"Калькулятор: вклад на врача-месяц, {val('Doctors_2025', 'Калькулятор', f'A{r}')}, тыс. ₽",
                  val("Doctors_2025", "Калькулятор", f"L{r}") / 1e3, want, 0.1)
        check("ДДС: остаток на конец года, ₽", val("CashFlow_2025", "ДДС", "N24"), 216_038_096.87, 1)
        check("Ликвидность: минимум дневного остатка в январе, ₽", val("CashFlow_2025", "Ликвидность", "F7"), 16_579_356, 1)
        check("Ликвидность: месяцев не «ок» (факт)", val("CashFlow_2025", "Ликвидность", "B20"), 0)
        check("Ликвидность: месяцев не «ок» (стресс)", val("CashFlow_2025", "Ликвидность", "C20"), 0)
        check("LTV12 январской когорты, ₽", val("PatientLTV_2025", "Когорты", "O8"), 34_101.48, 0.5)
        check("Пациентов в январской когорте", val("PatientLTV_2025", "Когорты", "B8"), 1287)
        check("Прогноз I кв. 2026 по сети, ₽", val("Forecast_2026Q1", "Прогноз", "F19"), 322_616_315.92, 1)
        check("Откат: ошибка по итогу", val("Forecast_2026Q1", "Прогноз", "H19"), 0.2023, 0.0005)

        # контрольные строки на листах «Описание»: разница не больше допуска округления
        for name in ("Doctors_2025", "CashFlow_2025", "PatientLTV_2025", "Forecast_2026Q1"):
            ws = book(name)["Описание"]
            num = (int, float)
            diffs = [abs(ws.cell(r, 4).value) for r in range(1, ws.max_row + 1)      # строки контроля: B, C, D числа
                     if ws.cell(r, 1).value and all(isinstance(ws.cell(r, c).value, num) for c in (2, 3, 4))]
            check(f"{name}: контрольных строк / максимальная разница", max(diffs) if diffs else None, 0, 0.2)

        # формулы: ошибки и пустые значения во всех книгах
        for name in BOOKS:
            wf = openpyxl.load_workbook(R / f"{name}.xlsx")
            wv = book(name)
            bad = n = 0
            for ws in wf.worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and c.value.startswith("="):
                            n += 1
                            v = wv[ws.title][c.coordinate].value
                            if v is None or (isinstance(v, str) and v.startswith("#")):
                                bad += 1
            check(f"{name}: ошибок формул и формул без значения (всего формул {n})", bad, 0)
    except Exception as e:   # любая ошибка чтения считается провалом проверки
        results.append((False, f"Проверка прервана: {type(e).__name__}: {e}", None, None))

    fails = [r for r in results if not r[0]]
    for ok, label, got, want in results:
        g = f"{got:,.4f}" if isinstance(got, float) else got
        print(("PASS " if ok else "FAIL ") + f"{label}: получено {g}, ожидалось {want}")
    print(f"\nИтого: {len(results) - len(fails)} из {len(results)} контролей пройдено")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
