-- Контроль 4. Прибыль до налога из проводок 1С: выручка минус закрытые расходы.
-- Ожидается 158 046 156 ₽ (P&L: 158 046 155,89; разница в копейках от округления проводок).
SELECT ROUND(SUM(CASE WHEN credit = '90.01' AND debit IN ('62.01', '62.02') THEN amount END), 2) AS revenue,
       ROUND(SUM(CASE WHEN debit = '90.02' THEN amount END), 2)                                  AS cost_of_sales,
       ROUND(SUM(CASE WHEN debit = '90.07' THEN amount END), 2)                                  AS selling_exp,
       ROUND(SUM(CASE WHEN debit = '90.08' THEN amount END), 2)                                  AS admin_exp,
       ROUND(SUM(CASE WHEN credit = '90.01' AND debit IN ('62.01', '62.02') THEN amount END)
           - SUM(CASE WHEN debit IN ('90.02', '90.07', '90.08') THEN amount END), 2)             AS profit_before_tax
FROM postings_1c;
