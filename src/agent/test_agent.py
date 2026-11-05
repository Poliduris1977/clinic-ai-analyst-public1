"""Проверка агента: (1) числа функций против контрольных цифр книг и DECISIONS, (2) маршрутизация 20 вопросов
(через JS из HTML, node), (3) ловушки: нет данных, наводящий вопрос, несуществующий врач."""
import json
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tools as T  # noqa: E402

ROOT = T.ROOT
HTML = (ROOT / "data" / "reports" / "Agent_2025.html").read_text(encoding="utf-8")
res = []


def check(name, got, exp, tol=0.0):
    ok = abs(got - exp) <= tol
    res.append(ok)
    print(("PASS " if ok else "FAIL ") + f"{name}: получено {got:,.3f}, ожидалось {exp:,.3f}")


g = T.plan_gap()
check("прибыль факт, ₽", g["profit_fact"], 158_046_000, 1e5)
check("отклонение прибыли, млн", g["profit_dev"] / 1e6, -21.517, 0.001)
R0 = T.reconciliation()
check("остаток разложения = разница ERP/1С (96 ₽), ₽", g["check"], -R0["diff"], 1)
D = T.discounts()
check("скидки, млн (п. 31)", D["discount"] / 1e6, -12.39, 0.01)
check("объём, млн (п. 31)", D["volume"] / 1e6, -18.32, 0.01)
L = T.liquidity(30)
check("мин. остаток факт, ₽ (книга ДДС)", L["min_fact"], 16_579_356, 1)
check("разрывов при стрессе", len(L["gap_stress"]), 0)
F = T.forecast_q1()
check("прогноз I кв. 2026, млн (п. 71)", F["dirs"]["Итого"]["fc"] / 1e6, 322.6, 0.05)
check("надёжность стоматологии, ошибка (п. 70)", F["dirs"]["DENT"]["err"], 0.047, 0.001)
R = T.reconciliation()
check("разница ERP-1С, ₽ (START)", R["diff"], -96, 0.5)
check("первопричин сверки", R["root"], 14)
d7 = T.doctors("D007")["doctor"]
check("D007: привязанный визит есть", float(d7["linked_visit"]), 1)

QS = [  # вопрос, ожидаемый ключ паспорта (префикс) или «нет данных»
 ("Почему мы не выполнили план по прибыли?", "q1|0||"), ("Почему в мае упала прибыль?", "q1|5||"),
 ("Что с прибылью по стоматологии?", "q1|0|DENT|"), ("Почему клиника C02 ниже плана?", "q1|0||C02"),
 ("Хватит ли денег, если страховые задержат платежи?", "q2|30|"), ("А если страховые задержат на 60 дней?", "q2|60|"),
 ("Что с ликвидностью в клинике C03 при задержке на две недели?", "q2|14|C03"), ("Стоит ли отменить скидки?", "q3"),
 ("Подтвердите, что во всём виноваты скидки", "q3"), ("Какие врачи дают самый низкий вклад?", "q4|"),
 ("Что с врачом D007?", "q4|D007"), ("Как дела у Жукова?", None), ("Как дела у Ширяева?", "q4|D001"), ("Чего ждать в первом квартале 2026?", "q5"),
 ("Какой будет прогноз выручки?", "q5"), ("Почему выручка в ERP не совпадает с бухгалтерией?", "q6"),
 ("Что нашла сверка с 1С?", "q6"),
 ("Какая прибыль по клинике C04?", None), ("Что было в 2024 году?", None), ("Что будет в 2027 году?", None),
 ("Что с врачом D099?", None), ("Какая погода завтра?", None), ("Хватит ли денег при задержке на 200 дней?", None),
]
js = "const D=JSON.parse(require('fs').readFileSync(0,'utf8'));" + \
     "global.localStorage={getItem:()=>null,setItem(){}};" + \
     re.search(r"(const SAMPLES=.*?)\ndocument\.getElementById\(\"chips\"", HTML, re.S).group(1).replace(
         "document.getElementById", "(()=>({}))") .split("function render")[0] + \
     "\nconst out=JSON.parse(process.argv[2]).map(q=>{const x=route(q);return {q,key:Object.keys(D.P).find(k=>D.P[k]===x.r)||null,plain:!!x.r.plain,pre:!!x.pre}});console.log(JSON.stringify(out))"
payload = json.dumps({"P": json.loads(re.search(r"const D=(\{.*?\});\nconst SAMPLES", HTML, re.S).group(1))["P"],
                      "names": json.loads(re.search(r"const D=(\{.*?\});\nconst SAMPLES", HTML, re.S).group(1))["names"]})
p = subprocess.run(["node", "-e", js, "x", json.dumps([q for q, _ in QS], ensure_ascii=False)], input=payload,
                   capture_output=True, text=True)
if p.returncode:
    print(p.stderr[-800:]); sys.exit(1)
for (q, exp), o in zip(QS, json.loads(p.stdout)):
    ok = o["plain"] if exp is None else o["key"] == exp
    res.append(ok)
    print(("PASS " if ok else "FAIL ") + f"«{q}» → {o['key'] or 'нет данных'}" + (" [наводящий]" if o["pre"] else ""))
# --- п. 34 и файловый журнал (JS из HTML) ---
js2 = js.split("\nconst out=")[0] + "\nconst qs=JSON.parse(process.argv[2]);const o=qs.map(q=>{const x=route(q);return {q,title:x.r.title,text:x.r.answer.join(' ')}});" + \
      "const E=[{id:'ЗАМ-001',k:'ЗАМ',w:'03.11.2026, 10:00:00',q:'в; \"кавычках\"',t:'T',c:'строка1\\nстрока2; с ; и \"кавычкой\"',s:'S',d:'D',st:'открыто'},{id:'РЕШ-002',k:'РЕШ',w:'w',q:'q',t:'t',c:'',s:'s',d:'d',st:'ok'}];" + \
      "const back=fromCsv(toCsv(E));const m=mergeJ(E.slice(0,1),back);console.log(JSON.stringify({o,rt:JSON.stringify(back)===JSON.stringify(E),merge:[m.added,m.skipped,m.all.length]}))"
Q34 = [
 ("В отчёте выручка 826 млн, у вас другая", "причина не установлена", "Ни версия данных"),
 ("В отчёте выручка 825,3 млн, а у вас другая", "причина установлена", "до исправлений"),
 ("В отчёте выручка 158 млн, у вас другая цифра", "причина установлена", "другой показатель"),
 ("В отчёте выручка 865 млн, у вас другая", "причина установлена", "план вместо факта"),
 ("В отчёте выручка 825 млн, не сходится?", "Расхождения нет", "совпадают"),
 ("В отчёте прибыль по стоматологии 50 млн, у вас другая", "Нужно уточнение", "не считается"),
 ("В отчёте цифра 826 млн, у вас другая", "Нужно уточнение", "каком показателе"),
]
p2 = subprocess.run(["node", "-e", js2, "x", json.dumps([q for q, *_ in Q34], ensure_ascii=False)], input=payload2 if False else json.dumps({**json.loads(payload), "V": json.loads(re.search(r"const D=(\{.*?\});\nconst SAMPLES", HTML, re.S).group(1))["V"], "mon": json.loads(re.search(r"const D=(\{.*?\});\nconst SAMPLES", HTML, re.S).group(1))["mon"], "dirs": json.loads(re.search(r"const D=(\{.*?\});\nconst SAMPLES", HTML, re.S).group(1))["dirs"], "dq34": "x"}),
                    capture_output=True, text=True)
if p2.returncode:
    print(p2.stderr[-900:]); sys.exit(1)
r2 = json.loads(p2.stdout)
for (q, tit, sub), o in zip(Q34, r2["o"]):
    ok = tit in o["title"] and sub in o["text"]
    res.append(ok)
    print(("PASS " if ok else "FAIL ") + f"п.34 «{q}» → {o['title']}" + ("" if ok else f" | {o['text'][:200]}"))
res.append(r2["rt"]); print(("PASS " if r2["rt"] else "FAIL ") + "журнал CSV: запись и чтение дают те же записи (кавычки, «;», перевод строки)")
ok = r2["merge"] == [1, 1, 2]; res.append(ok); print(("PASS " if ok else "FAIL ") + f"журнал CSV: слияние (добавлено, дублей, всего) = {r2['merge']}")
print(f"\nИтого: {sum(res)} из {len(res)}")
sys.exit(0 if all(res) else 1)
