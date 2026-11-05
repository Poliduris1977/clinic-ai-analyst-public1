-- Контроль 9. Врачи: все ли визиты привязаны к врачу и сколько врач-месяцев в расчёте вклада.
-- Ожидается 43 врача, 479 врач-месяцев, 0 визитов без врача; выручка по врачам равна выручке ERP 825 211 700 ₽.
SELECT (SELECT COUNT(DISTINCT doctor_id) FROM visits WHERE doctor_id IS NOT NULL)            AS doctors,
       (SELECT COUNT(*) FROM (SELECT DISTINCT doctor_id, substr(visit_date, 1, 7)
                              FROM visits WHERE doctor_id IS NOT NULL))                       AS doctor_months,
       (SELECT COUNT(*) FROM visits WHERE doctor_id IS NULL)                                  AS visits_no_doctor,
       (SELECT ROUND(SUM(i.revenue), 2) FROM visit_items i JOIN visits v USING (visit_id)
        WHERE v.doctor_id IS NOT NULL)                                                        AS revenue_by_doctors;
