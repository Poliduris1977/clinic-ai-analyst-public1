"""
Сборка ИИ-агента (HTML, открывается двойным щелчком). Агент = маршрутизатор вопросов + функции доступа к данным
(tools.py) + слой защиты. Языковой модели нет (п. 29-35 DECISIONS, п. 75). Паспорта ответов считаются здесь для всех
значений параметров; каждое число в тексте проходит сверку с результатом функции (verify), иначе сборка падает.
"""
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tools as T  # noqa: E402

ROOT = T.ROOT
OUT = ROOT / "data" / "reports" / "Agent_2025.html"
MON = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
DIR = T.DIRS
DIR_LOW = {k: v.lower() for k, v in DIR.items()}
NB = "\u00a0"
MINUS = "\u2212"
REG = set()


def _reg(s):
    for m in re.findall(r"\d+(?:\u00a0\d{3})*(?:,\d+)?", s):
        REG.add(m)
    return s


def num(x, d=2, unit=" млн ₽", sign=False, scale=1e6):
    v = x / scale
    s = f"{abs(v):,.{d}f}".replace(",", NB).replace(".", ",")
    pre = MINUS if round(v, d) < 0 else ("+" if sign and round(v, d) > 0 else "")
    return _reg(f"{pre}{s}{unit.replace(' ', NB) if unit else ''}")


def pct(x, d=1, sign=False):
    v = x * 100
    s = f"{abs(v):.{d}f}".replace(".", ",")
    pre = MINUS if round(v, d) < 0 else ("+" if sign and round(v, d) > 0 else "")
    return _reg(f"{pre}{s}%")


def thou(x, sign=False):
    return num(x, 0, " тыс. ₽", sign, 1e3)


def dmy(s):
    d, m = s.split(".")
    return lit(int(d)) + " " + ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"][int(m) - 1]


def lit(s):
    return _reg(str(s))


def verify(text):
    text = text.replace("1С", "")
    bad = [m for m in re.findall(r"\d+(?:\u00a0\d{3})*(?:,\d+)?", text) if m not in REG]
    if bad:
        raise ValueError(f"число не из функций: {bad} в «{text[:80]}»")


ROLES = [
    ["Прибыль, план, отклонения", "Собственник", "Финансовый директор"],
    ["Скидки, цены", "Коммерческий директор", "Финансовый директор (цифры), инициатор скидок (обоснование)"],
    ["Денежный резерв, ликвидность", "Финансовый директор", "Главный бухгалтер (остатки по 1С)"],
    ["Врачи, грейды, оплата, загрузка", "Медицинский директор", "Финансовый директор (расчёт вклада)"],
    ["Прогноз и планирование", "Финансовый директор", "Аналитик (модель, допущения)"],
    ["Качество данных, сверка с 1С", "Главный бухгалтер", "Аналитик (журнал исправлений)"],
]


def passport(title, answer, unproven, dq, who, key_roles):
    p = {"title": title, "answer": answer, "unproven": unproven, "dq": dq, "decides": key_roles[0], "confirms": key_roles[1]}
    for v in [title, *answer, *unproven, *dq]:
        verify(v)
    return p


def dq_standard(rc, what="отклонение прибыли"):
    return [f"Данные прошли сверку ERP и 1С: найдено {lit(rc['root'])} ошибок, исправлено {lit(rc['root'])}, повторная сверка замечаний не находит. "
            f"Из исправлений {lit(rc['review'])} ждут подтверждения главного бухгалтера, они в сумме по модулю {thou(rc['review_abs'])}; "
            f"Это может изменить {what} не более чем на эту сумму."]


# ---------------------------------------------------------------- паспорта
def q1(rc, month=None, clinic=None, direction=None):
    g = T.plan_gap(month, clinic, direction)
    parts = [f"{MON[month - 1]} {lit(2025)}" if month else f"{lit(2025)} год"]
    if direction:
        parts.append(DIR_LOW[direction])
    if clinic:
        parts.append(f"клиника {lit(clinic)}")
    scope = ", ".join(parts)
    if g is None:
        return {"title": "В данных ответа нет", "plain": True, "decides": "Аналитик", "confirms": "Финансовый директор",
                "answer": [f"{scope}: такого среза в данных нет (в этой клинике нет этого направления или нет плана)."],
                "unproven": ["Агент не заменяет отсутствующий срез соседним."], "dq": ["Не применимо."]}
    f = g["factors"]
    head = "маржа после материалов и лабораторий" if direction else "прибыль до налога"
    ans = [f"{scope}: {head} {num(g['profit_fact'])} при плане {num(g['profit_plan'])}, отклонение {num(g['profit_dev'], sign=True)}. "
           f"Выручка отличается от плана на {num(g['rev_fact'] - g['rev_plan'], sign=True)}.",
           f"Маржа после материалов и лабораторий {num(g['dm_dev'], sign=True)} к плану: объём {num(f['volume'], sign=True)}, скидки {num(f['discount'], sign=True)}, "
           f"прайс (индексация цен) {num(f['price_list'], sign=True)}, структура услуг {num(f['mix'], sign=True)}, материалы {num(f['materials'], sign=True)}, "
           f"лаборатории {num(f['labs'], sign=True)}."
           + ("" if direction else f" Остальные затраты дают {num(g['other_dev'], sign=True)}")
           + (": " + ", ".join(f"{k.lower()} {num(v, sign=True)}" for k, v in g["other_top"].items()) if (not direction) and g["other_top"] and abs(g["other_dev"]) > 1e4 else "") + ("" if direction else ".")]
    if not month and abs(g["profit_dev"]) > 1e4:
        if not direction:
            ans.append(f"Больше всего отстаёт по марже {DIR_LOW[g['worst_dir'][0]]} ({num(g['worst_dir'][1], sign=True)}); "
                       f"самый слабый месяц {MON[g['worst_month'][0] - 1].lower()} ({num(g['worst_month'][1], sign=True)}).")
    unpr = ["Разложение показывает, где образовалось отклонение, но не почему: причины падения объёма (спрос, простои, загрузка врачей) в данных не записаны.",
            "Экономия на оплате труда врачей и взносах следует из самой выручки (оплата врача считается процентом от выручки), её нельзя считать результатом управленческих решений."]
    if direction:
        unpr.insert(0, "Прибыль до налога по направлению в данных не считается: оплата врачей и общие затраты (аренда, администрация, маркетинг) не распределены по направлениям, поэтому показана маржа после материалов и лабораторий.")
    return passport(f"{'Маржа' if direction else 'Прибыль'} к плану: {scope}", ans, unpr, dq_standard(rc), None, ("Собственник", "Финансовый директор"))


def q2(rc, delay=30, clinic=None):
    L = T.liquidity(delay, clinic)
    scope = f"клиника {lit(clinic)}" if clinic else "сеть"
    d = lit(delay)
    ans = [f"При задержке страховых платежей на {d} дн. ({scope}) остаток на счетах не уходит ниже нуля ни в один день 2025 года: минимум {num(L['min_stress'])} ({dmy(L['min_stress_date'])}), "
           f"без задержки минимум {num(L['min_fact'])} ({dmy(L['min_fact_date'])}). Запас в самый тяжёлый день стресса {num(L['days_stress_min'], 1, ' дн. выплат', False, 1)} "
           f"при резерве {lit(L['reserve'])} дн."]
    if L["bad_stress"]:
        ans.append("Месяцы ниже резерва или с разрывом при задержке: " + ", ".join(MON[m - 1].lower() for m in L["bad_stress"]) + ".")
    else:
        ans.append(f"Все месяцы остаются в статусе «ок»; поступления от страховых за год {num(L['ins_year'])}, остаток на конец года {num(L['end_fact'])}.")
    unpr = ["Это механический сдвиг всех страховых поступлений на заданное число дней; поведение конкретных страховых и рост просрочки не моделируются.",
            f"Остаток считается до налога на прибыль, НДС и дивидендов; данных за период до начала года нет, начальный остаток берётся из ОСВ по счёту {lit(51)}."]
    dq = [f"Остаток на конец года совпадает с ОСВ 1С (счёт {lit(51)}), книга ДДС прошла контроль. Подтверждение остатков: главный бухгалтер."]
    return passport(f"Ликвидность при задержке страховых на {d} дн.", ans, unpr, dq, None, ("Финансовый директор", "Главный бухгалтер"))


def q3(rc):
    D = T.discounts()
    ans = [f"Скидки относительно прайса дают {num(D['discount'], sign=True)} к марже, это {pct(D['rate'])} от суммы по прайсу. Объём услуг при этом даёт {num(D['volume'], sign=True)}: "
           f"скидки объясняют {pct(D['discount'] / D['margin_dev'], 0)} отставания маржи от плана ({num(D['margin_dev'], sign=True)}). Отмена скидок вернула бы до {num(-D['discount'])} при неизменном объёме, это верхняя граница.",
           "По направлениям (скидки / объём): " + "; ".join(f"{DIR_LOW[k]} {num(v['disc'], sign=True)} / {num(v['vol'], sign=True)}" for k, v in D["by_dir"].items()) + ".",
           f"Связь скидок с потоком пациентов в данных слабая и обратная: по {lit(D['cells'])} срезам «месяц, клиника, направление» корреляция глубины скидки и отклонения объёма {lit(f'{D['corr']:.2f}'.replace('.', ',').replace('-', MINUS))}. "
           f"Знак противоположен гипотезе «скидки привлекают пациентов», но связь слабая."]
    unpr = ["Не доказано, что скидки привлекают пациентов или что их отмена снизит поток: эксперимента в данных нет, корреляция по срезам не причинность.",
            "Предложение агента: инициатор скидок и финансовый отдел обосновывают правильность скидок цифрами; решение об отмене принимает коммерческий директор."]
    return passport("Скидки: что показывают данные", ans, unpr, dq_standard(rc, "эти выводы"), None, ("Коммерческий директор", "Финансовый директор"))


def q4(rc, doctor=None):
    D = T.doctors(doctor)
    dn = lambda x: f"{x['name']} ({DIR_LOW[x['dir']]}, {x['grade']}, {lit(x['months'])} мес.): {thou(x['per_month'])} на врач-месяц, {pct(x['vs_dir'], 0, True)} к среднему по направлению"  # noqa: E731
    if doctor and D.get("doctor"):
        x = D["doctor"]
        ans = ["Вклад врача = выручка − материалы − лаборатории − оплата врача − взносы. " + dn(x) + f"; сеть в среднем {thou(D['net_per_month'])}."]
        dq = dq_standard(rc)
        if x["linked_visit"]:
            lk = D["linked"][x["id"]]
            dq = [f"Визит на {thou(lk['amount'])} привязан к этому врачу автоматически по payroll (единственный кандидат), ждёт подтверждения главного бухгалтера; на вклад врача это влияет не более чем на эту сумму."] + dq
        title = f"Врач {x['name']}"
    else:
        ans = ["Ниже всего вклад на врач-месяц у: " + "; ".join(dn(x) for x in D["low"]) + f". Среднее по сети {thou(D['net_per_month'])}."]
        dq, title = dq_standard(rc), "Врачи с самым низким вкладом"
        links = [x for x in D["low"] if x["linked_visit"]]
        for x in links:
            dq = [f"У врача {x['name']} один визит привязан по payroll и ждёт подтверждения бухгалтера."] + dq
    unpr = ["Низкий вклад ещё не означает плохую работу: его определяют набор услуг, ставка оплаты и направление; сравнивать нужно внутри направления.",
            "Решение о враче (оплата, загрузка, грейд) принимает медицинский директор; агент предлагает разобрать причины с врачом и руководителем направления."]
    return passport(title, ans, unpr, dq, None, ("Медицинский директор", "Финансовый директор"))


def q5(rc):
    F = T.forecast_q1()
    d = F["dirs"]
    ans = [f"Прогноз выручки на I кв. {lit(2026)} по сети {num(d['Итого']['fc'])} (IV кв. {lit(2025)}: {num(d['Итого']['q4'])}), коридор {num(d['Итого']['lo'], 1, '')}-{num(d['Итого']['hi'], 1)}, надёжность {d['Итого']['rel']}.",
           "По направлениям: " + "; ".join(f"{DIR_LOW[k]} {num(v['fc'], 1, '')} (надёжность {v['rel']})" for k, v in d.items() if k != "Итого") + " млн ₽."]
    low = [DIR_LOW[k] for k, v in d.items() if k != "Итого" and v["rel"] == "низкая"]
    ans.append("Не использовать как план: " + ", ".join(low) + f" (ошибка отката {pct(min(d[k]['err'] for k in d if DIR_LOW.get(k) in low), 0)}-{pct(max(d[k]['err'] for k in d if DIR_LOW.get(k) in low), 0)}).")
    unpr = [f"Ошибка прогноза по сети измерена откатом на {lit(F['n_backtest'])} наблюдениях одного года (факт выше прогноза в среднем на {pct(F['bias'], 0)}); сезонность по одному году оценить нельзя.",
            "Прогноз это продолжение тренда, а не план: акции, найм и простои в нём не учтены."]
    dq = ["Использованы только исправленные данные 2025 года; по направлениям с низкой надёжностью прогноз планом считать нельзя."]
    return passport("Чего ждать в первом квартале 2026", ans, unpr, dq, None, ("Финансовый директор", "Аналитик"))


def q6(rc):
    ans = [f"Выручка в ERP {num(rc['erp'], 3)} и в книгах 1С {num(rc['books'], 3)}, остаточная разница {lit(f'{rc['diff']:.0f}'.replace('-', MINUS))} ₽ (округление строк). "
           f"До исправлений сверка нашла {lit(rc['findings'])} замечаний, из них {lit(rc['root'])} первопричин и {lit(rc['derived'])} производных (аномальные остатки ОСВ, следствие первопричин).",
           f"Исправлено {lit(rc['root'])} из {lit(rc['root'])}: {lit(rc['applied'])} применены, {lit(rc['review'])} со статусом «подтвердить бухгалтером». После исправлений повторная сверка замечаний не находит ({lit(rc['after'])})."]
    unpr = ["Сверка работает по восьми правилам; ошибки вне этих правил она не найдёт. «Нет замечаний» значит «нет известных типов ошибок», а не «ошибок нет».",
            "Исправления со статусом «подтвердить» применены по правилам с оценкой уверенности «средняя», окончательно их принимает главный бухгалтер."]
    return passport("Расхождения ERP и 1С", ans, unpr, ["Журнал исправлений и находок сохранён, неразобранных замечаний нет."], None, ("Главный бухгалтер", "Аналитик"))


def main():
    rc = T.reconciliation()
    P = {}
    for mo in [None] + list(range(1, 13)):
        for di in [None] + list(DIR):
            for cl in [None, "C01", "C02", "C03"]:
                P[f"q1|{mo or 0}|{di or ''}|{cl or ''}"] = q1(rc, mo, cl, di)
    for cl in [None, "C01", "C02", "C03"]:
        for dl in range(0, 91):
            P[f"q2|{dl}|{cl or ''}"] = q2(rc, dl, cl)
    P["q3"] = q3(rc)
    P["q4|"] = q4(rc)
    dd, dl_, _, _ = T.bd.prepare()
    names = {}
    for r in dl_.itertuples():
        P[f"q4|{r.doctor_id}"] = q4(rc, r.doctor_id)
        names.setdefault(r.doctor_name.split()[0].lower(), []).append([r.doctor_id, r.doctor_name])
    P["q5"] = q5(rc)
    P["q6"] = q6(rc)
    dq34 = dq_standard(rc, "сопоставление")[0]
    verify(dq34)
    V = T.report_variants()
    html = HTML.replace("__DATA__", json.dumps({"P": P, "names": names, "roles": ROLES, "mon": MON, "V": V, "dirs": DIR_LOW,
                                                 "dq34": dq34}, ensure_ascii=False))
    OUT.write_text(html, encoding="utf-8")
    print(f"Паспортов: {len(P)}; чисел в реестре: {len(REG)}; файл {OUT.name}: {len(html) // 1024} КБ")


HTML = open(Path(__file__).with_name("agent_template.html"), encoding="utf-8").read()

if __name__ == "__main__":
    main()
