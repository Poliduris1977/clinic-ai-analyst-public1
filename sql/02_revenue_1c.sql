-- Контроль 2. Выручка по данным 1С: проводки Дт 62.01/62.02 Кт 90.01 (реализация услуг).
-- Ожидается 825 211 796 ₽. Разница с ERP 96 ₽ это округление цен ERP до сотен рублей, не ошибка.
SELECT ROUND(SUM(amount), 2)  AS revenue_1c,
       COUNT(*)               AS postings
FROM postings_1c
WHERE credit = '90.01' AND debit IN ('62.01', '62.02');
