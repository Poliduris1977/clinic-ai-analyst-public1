-- Контроль 5. Остаток денег на конец года тремя путями: ОСВ, ДДС, проводки по счёту 51.
-- Ожидается 216 038 097 ₽ (начало 31 000 000 плюс чистое изменение 185 038 097).
SELECT (SELECT ROUND(SUM(opening_debit - opening_credit), 2)
        FROM osv_1c WHERE account = '51' AND period = '2025-01')                       AS opening_51,
       (SELECT ROUND(SUM(signed_amount), 2) FROM cashflow_1c)                          AS net_flow_dds,
       (SELECT ROUND(SUM(closing_debit - closing_credit), 2)
        FROM osv_1c WHERE account = '51' AND period = '2025-12')                       AS closing_osv,
       (SELECT ROUND(SUM(CASE WHEN debit = '51' THEN amount ELSE 0 END)
                   - SUM(CASE WHEN credit = '51' THEN amount ELSE 0 END), 2)
        FROM postings_1c)                                                              AS net_flow_postings;
