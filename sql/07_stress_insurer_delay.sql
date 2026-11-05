-- Контроль 7. Стресс: задержка поступлений от страховых на 30 дней (вся сеть).
-- Остаток за день при задержке = факт минус страховые поступления за последние 30 дней.
-- Другая задержка N: заменить 29 на N-1 в рамке окна (для 0 дней вычитание убрать).
-- Ожидается остаток на конец года при задержке 204 579 397 ₽ и минимум года 16 579 356 ₽ (лист «Ликвидность» книги CashFlow_2025).
WITH RECURSIVE cal(d) AS (
  SELECT '2025-01-01'
  UNION ALL SELECT date(d, '+1 day') FROM cal WHERE d < '2025-12-31'),
flow AS (
  SELECT d, COALESCE(SUM(c.signed_amount), 0) AS net,
            COALESCE(SUM(CASE WHEN c.article = 'Поступления от страховых компаний' THEN c.signed_amount END), 0) AS ins
  FROM cal LEFT JOIN cashflow_1c c ON c.date = cal.d
  GROUP BY d),
bal AS (
  SELECT d,
         (SELECT SUM(opening_debit - opening_credit) FROM osv_1c WHERE account = '51' AND period = '2025-01')
           + SUM(net) OVER (ORDER BY d) AS balance,
         SUM(ins) OVER (ORDER BY d ROWS BETWEEN 29 PRECEDING AND CURRENT ROW) AS ins_window
  FROM flow),
st AS (SELECT d, balance, balance - ins_window AS stress FROM bal)
SELECT ROUND(MIN(stress), 2)                                    AS min_stress,
       (SELECT ROUND(stress, 2) FROM st WHERE d = '2025-12-31') AS end_stress,
       ROUND(MIN(balance), 2)                                   AS min_fact
FROM st;
