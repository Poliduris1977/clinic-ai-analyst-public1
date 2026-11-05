-- Контроль 8. Новые пациенты 2025 и LTV за 12 месяцев январской когорты.
-- Новый пациент: первый визит в 2025 году (patients.first_visit_date). Когорта: месяц первого завершённого визита.
-- Внимание: суммы выручки в SQLite целые, поэтому делим на 1.0 * сумма; иначе целочисленное деление даёт 34 101 вместо 34 101,48.
-- Ожидается 27 924 пациента, выручка новых 494 762 800 ₽, январская когорта 1 287 человек, LTV12 34 101 ₽ на пациента.
WITH vrev AS (
  SELECT v.patient_id, v.visit_date, SUM(i.revenue) AS revenue
  FROM visits v JOIN visit_items i USING (visit_id)
  WHERE v.status = 'completed'
  GROUP BY v.visit_id),
newv AS (
  SELECT r.* FROM vrev r JOIN patients p USING (patient_id)
  WHERE p.first_visit_date >= '2025-01-01'),
cohort AS (
  SELECT patient_id, MIN(substr(visit_date, 1, 7)) AS cohort_month, SUM(revenue) AS rev_total
  FROM newv GROUP BY patient_id)
SELECT COUNT(*)                                                                AS new_patients,
       ROUND(SUM(rev_total), 2)                                                AS rev_new,
       SUM(cohort_month = '2025-01')                                           AS jan_patients,
       ROUND(1.0 * SUM(CASE WHEN cohort_month = '2025-01' THEN rev_total END)
             / SUM(cohort_month = '2025-01'), 2)                               AS ltv12_jan
FROM cohort;
