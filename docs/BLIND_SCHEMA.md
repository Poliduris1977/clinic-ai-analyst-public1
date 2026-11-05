# Схема данных для автора слепого набора ошибок

Только структура и примеры чистых файлов `data/raw/`. Правил сверки здесь нет и быть не должно. Ошибки вносятся в копию этих файлов (`data/dirty_v2/`). Данные синтетические (воображаемая сеть клиник, 2025 год).


## budget_opex.csv  (360 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| period | str | 12 | 2025-01, 2025-01 |
| clinic_id | str | 3 | C01, C02, C03 |
| article | str | 10 | Оплата труда врачей, Оплата труда медперсонала |
| pnl_block | str | 3 | cost_of_sales, management, commercial |
| plan_amount | float64 | 158 | 6526759.2, 2400000.0 |

## budget_revenue.csv  (883 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| period | str | 12 | 2025-01, 2025-01 |
| clinic_id | str | 3 | C01, C02, C03 |
| direction_code | str | 5 | DENT, COSM, PLAS, FAM, LASER |
| service_id | str | 36 | S001, S002 |
| plan_qty | int64 | 281 | 107, 61 |
| plan_price | int64 | 60 | 1500, 7500 |
| plan_revenue | int64 | 708 | 160500, 457500 |
| plan_unit_material | float64 | 359 | 13.0, 422.0 |
| plan_unit_external | float64 | 167 | 0.0, 0.0 |

## cashflow_1c.csv  (2300 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| date | str | 341 | 2025-01-01, 2025-01-01 |
| clinic_id | str | 3 | C01, C02, C03 |
| activity | str | 2 | operating, investing |
| article | str | 11 | Комиссия эквайринга, Поступления от пациентов |
| flow | str | 2 | out, in |
| amount | float64 | 2200 | 2622.6, 145700.0 |
| signed_amount | float64 | 2200 | -2622.6, 145700.0 |
| counterparty | str | 5 | Согласие, Ингосстрах, АльфаСтрахование, СОГАЗ-Мед, РЕСО-Гарантия |

## clinics.csv  (3 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| clinic_id | str | 3 | C01, C02, C03 |
| clinic_name | str | 3 | Клиника Невский, Клиника Север, Клиника Юг |
| district | str | 3 | Центральный, Выборгский, Московский |
| directions | str | 3 | DENT|COSM|PLAS|FAM|LASER, DENT|COSM|FAM, DENT|FAM|LASER |
| rent_month | int64 | 3 | 1450000, 780000, 820000 |

## doctors.csv  (43 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| doctor_id | str | 43 | D001, D002 |
| doctor_name | str | 43 | Ширяев Е.А., Нестеров Е.Д. |
| direction_code | str | 5 | DENT, COSM, PLAS, FAM, LASER |
| direction | str | 5 | Стоматология, Косметология, Пластическая хирургия, Семейная медицина, Лазерная коррекция з |
| specialty | str | 12 | Стоматолог-ортопед, Ортодонт |
| clinic_id | str | 3 | C01, C02, C03 |
| grade | str | 4 | senior, junior, middle, lead |
| pay_pct | float64 | 4 | 0.36, 0.28, 0.32, 0.4 |
| popularity_weight | float64 | 4 | 1.5, 0.6, 1.0, 2.0 |
| hire_date | str | 38 | 2018-03-01, 2015-01-01 |

## materials_1c.csv  (132 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| period | str | 12 | 2025-01, 2025-01 |
| clinic_id | str | 3 | C01, C02, C03 |
| direction_code | str | 5 | COSM, DENT, FAM, LASER, PLAS |
| material_plan | int64 | 132 | 810940, 243350 |
| material_fact | float64 | 132 | 810959.95, 243355.99 |
| external_plan | int64 | 102 | 0, 1007500 |
| external_fact | float64 | 105 | 0.0, 1012014.79 |
| material_index | float64 | 36 | 1.0, 1.0 |
| external_index | float64 | 36 | 1.0045, 1.0045 |

## osv_1c.csv  (720 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| period | str | 12 | 2025-01, 2025-01 |
| clinic_id | str | 3 | C01, C02, C03 |
| account | float64 | 20 | 1.0, 2.0 |
| account_name | str | 20 | Основные средства, Амортизация основных средств |
| opening_debit | float64 | 115 | 70000000.0, 0.0 |
| opening_credit | float64 | 216 | 0.0, 24500000.0 |
| turnover_debit | float64 | 437 | 0.0, 0.0 |
| turnover_credit | float64 | 471 | 0.0, 1100000.0 |
| closing_debit | float64 | 116 | 70000000.0, 0.0 |
| closing_credit | float64 | 218 | 0.0, 25600000.0 |

## patients.csv  (36609 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| patient_id | str | 36609 | P000001, P000002 |
| first_visit_date | str | 1758 | 2024-08-02, 2022-07-18 |
| home_clinic_id | str | 3 | C01, C02, C03 |
| birth_year | int64 | 63 | 1957, 1965 |
| source | str | 7 | Яндекс, Яндекс |

## payroll_1c.csv  (551 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| period | str | 12 | 2025-01, 2025-01 |
| clinic_id | str | 3 | C01, C02, C03 |
| employee_id | str | 49 | D001, D002 |
| employee_name | str | 49 | Ширяев Е.А., Нестеров Е.Д. |
| role | str | 3 | doctor, nurse, admin |
| direction_code | str | 5 | DENT, COSM, PLAS, FAM, LASER |
| base_revenue | float64 | 475 | 2449000.0, 971300.0 |
| pay_pct | float64 | 4 | 0.36, 0.28, 0.32, 0.4 |
| gross | int64 | 546 | 881640, 271964 |
| ndfl | int64 | 545 | 114613, 35355 |
| net_pay | int64 | 546 | 767027, 236609 |
| contributions | int64 | 546 | 193961, 59832 |

## postings_1c.csv  (13145 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| posting_id | str | 13145 | PST0000001, PST0000002 |
| date | str | 342 | 2025-01-01, 2025-01-01 |
| clinic_id | str | 3 | C01, C02, C03 |
| debit | float64 | 18 | 26.0, 26.0 |
| credit | float64 | 18 | 60.0, 51.0 |
| amount | float64 | 5895 | 1450000.0, 2622.6 |
| direction_code | str | 5 | COSM, DENT, FAM, LASER, PLAS |
| counterparty | str | 5 | СОГАЗ-Мед, АльфаСтрахование, РЕСО-Гарантия, Согласие, Ингосстрах |
| doc_type | str | 26 | Аренда помещений, Комиссия эквайринга |
| comment | str | 1 | Оплата акта за декабрь 2024 |

## price_history.csv  (66 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| service_id | str | 36 | S001, S001 |
| valid_from | str | 4 | 2025-01-01, 2025-04-01, 2025-02-01, 2025-09-01 |
| list_price | int64 | 59 | 1500, 1600 |

## services.csv  (36 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| service_id | str | 36 | S001, S002 |
| direction_code | str | 5 | DENT, COSM, PLAS, FAM, LASER |
| direction | str | 5 | Стоматология, Косметология, Пластическая хирургия, Семейная медицина, Лазерная коррекция з |
| service_name | str | 36 | Консультация стоматолога, Профессиональная гигиена |
| list_price | int64 | 32 | 1500, 7500 |
| material_cost | int64 | 31 | 13, 422 |
| external_cost | int64 | 15 | 0, 0 |
| plan_cost | int64 | 33 | 13, 422 |
| plan_margin_pct | float64 | 34 | 0.9913, 0.9437 |
| duration_min | int64 | 11 | 30, 60 |

## visit_items.csv  (81532 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| item_id | str | 81532 | I0000001, I0000002 |
| visit_id | str | 66411 | V0000001, V0000001 |
| service_id | str | 36 | S001, S007 |
| qty | int64 | 3 | 1, 3, 2 |
| list_price_at_date | int64 | 59 | 1500, 24000 |
| discount_pct | int64 | 4 | 0, 5, 10, 15 |
| price_fact | int64 | 159 | 1500, 24000 |
| revenue | int64 | 168 | 1500, 24000 |

## visit_payments.csv  (63335 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| payment_id | str | 63335 | PAY0000001, PAY0000002 |
| visit_id | str | 63335 | V0000001, V0000002 |
| pay_date | str | 332 | 2025-01-01, 2025-01-01 |
| amount | int64 | 645 | 25500, 1500 |
| payer_type | str | 2 | self, dms |
| payer_name | str | 6 | Пациент, СОГАЗ-Мед, Ингосстрах, АльфаСтрахование, Согласие, РЕСО-Гарантия |

## visits.csv  (73890 строк)

| колонка | тип | уникальных | примеры |
|---|---|---|---|
| visit_id | str | 73890 | V0000001, V0000002 |
| patient_id | str | 36609 | P009461, P009462 |
| doctor_id | str | 43 | D001, D001 |
| clinic_id | str | 3 | C01, C02, C03 |
| direction_code | str | 5 | DENT, COSM, FAM, LASER, PLAS |
| visit_date | str | 313 | 2025-01-01, 2025-01-01 |
| visit_time | str | 44 | 13:00, 17:45 |
| status | str | 3 | completed, cancelled, no_show |
| is_first_in_direction | bool | 2 | True, False |
