-- Контроль 6. Остаток по дням (оконная функция) и его минимум за год.
-- Календарь строится рекурсивно, чтобы в днях без движений остаток не пропадал.
-- Ожидается минимум 16 579 356 ₽ 10 января (день выплат начала года).
WITH RECURSIVE cal(d) AS (
  SELECT '2025-01-01'
  UNION ALL SELECT date(d, '+1 day') FROM cal WHERE d < '2025-12-31'),
flow AS (
  SELECT d, COALESCE(SUM(c.signed_amount), 0) AS net
  FROM cal LEFT JOIN cashflow_1c c ON c.date = cal.d
  GROUP BY d),
bal AS (
  SELECT d, (SELECT SUM(opening_debit - opening_credit) FROM osv_1c WHERE account = '51' AND period = '2025-01')
            + SUM(net) OVER (ORDER BY d) AS balance
  FROM flow)
SELECT ROUND(MIN(balance), 2)                                     AS min_balance,
       (SELECT d FROM bal ORDER BY balance LIMIT 1)               AS min_date,
       (SELECT ROUND(balance, 2) FROM bal WHERE d = '2025-12-31') AS end_balance,
       COUNT(*)                                                   AS days
FROM bal;
