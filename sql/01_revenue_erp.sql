-- Контроль 1. Выручка по данным ERP (визиты и позиции услуг).
-- Ожидается 825 211 700 ₽ на исправленных данных (PnL: таблица контролей в docs/START.md).
SELECT ROUND(SUM(revenue), 2)  AS revenue_erp,
       COUNT(*)                AS items
FROM visit_items;
