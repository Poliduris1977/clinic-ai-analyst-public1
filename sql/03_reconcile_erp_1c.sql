-- Контроль 3. Сверка выручки ERP и 1С по клиникам и месяцам (36 ячеек: 3 клиники x 12 месяцев).
-- Порог существенности 1 000 ₽: меньшие расхождения это округление цен в ERP.
-- На «грязных» данных запрос находит 10 существенных расхождений, на исправленных ни одного.
WITH erp AS (
  SELECT v.clinic_id, substr(v.visit_date, 1, 7) AS period, SUM(i.revenue) AS erp
  FROM visit_items i JOIN visits v USING (visit_id)
  GROUP BY 1, 2),
c1 AS (
  SELECT clinic_id, substr(date, 1, 7) AS period, SUM(amount) AS c1
  FROM postings_1c
  WHERE credit = '90.01' AND debit IN ('62.01', '62.02')
  GROUP BY 1, 2),
d AS (
  SELECT erp.clinic_id, erp.period, erp.erp - COALESCE(c1.c1, 0) AS diff
  FROM erp LEFT JOIN c1 USING (clinic_id, period))
SELECT COUNT(*)                                          AS cells,
       SUM(ABS(diff) > 1000)                             AS cells_over_1000,
       ROUND(SUM(diff), 2)                               AS total_diff
FROM d;
