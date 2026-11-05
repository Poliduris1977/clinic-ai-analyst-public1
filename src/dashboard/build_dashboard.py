#!/usr/bin/env python3
"""Сборка Dashboard_2025.html (каркас + вкладка «Качество данных»).
Числа вкладки берутся из таблицы ошибок ниже; итоги пересчитываются и сверяются
с заявленными в docs/BLIND_RESULT.md (п. 29): расхождение останавливает сборку."""
import csv, json, sys, pathlib
import openpyxl

# (id, ошибка, тонкая, найдена в 1-м прогоне, причина названа в 1-м прогоне)
E = [
 ("E01","проводка не по тому направлению",0,1,"да"),
 ("E02","проводка не по той клинике",0,1,"да"),
 ("E03","поступление не от той страховой",0,1,"да"),
 ("E04","тип плательщика ДМС/пациент",1,1,"симптом"),
 ("E05","отменённый визит с выручкой",0,1,"симптом"),
 ("E06","ставка комиссии эквайринга",1,0,"—"),
 ("E07","поступление не загружено в 1С",0,1,"да"),
 ("E08","платёж меньше стоимости визита",1,1,"симптом"),
 ("E09","врач чужой клиники",1,0,"—"),
 ("E10","ставка ЗП не по справочнику",0,0,"—"),
 ("E11","затраты не на то направление",0,0,"—"),
 ("E12","цена по старому прайсу",1,1,"да"),
 ("E13","ручная правка выручки +145 ₽",1,1,"да"),
 ("E14","платёж сдвинут на следующий месяц",0,1,"симптом"),
]
RUN1 = dict(findings=19, correct=15, false=4, harmless=10)   # BLIND_RESULT, п. 85
RUN2 = dict(findings=20, correct=20, false=0, harmless=12)   # п. 87 (E11 даёт две строки)
HASH1, HASH2 = "59cff04f…", "53126a39…"

ARGS=[a for a in sys.argv[1:]]
OUT = pathlib.Path(ARGS[0] if ARGS else "Dashboard_2025.html")
DATA = pathlib.Path(ARGS[1]) if len(ARGS) > 1 else pathlib.Path(__file__).resolve().parents[2] / "data"
with open(DATA/"output_v2"/"blind_compare.csv", encoding="utf-8-sig") as f:
    CMP = {r["id"]: r for r in csv.DictReader(f)}
mism = [e[0] for e in E if (CMP[e[0]]["first"] == "да") != bool(e[3]) or CMP[e[0]]["second"] != "да"
        or int(CMP[e[0]]["subtle"]) != e[2]]
if mism: sys.exit(f"blind_compare.csv расходится с таблицей в скрипте: {mism}")

# --- Обзор: числа из книг (в рублях)
def rows(ws): return [list(r) for r in ws.iter_rows(values_only=True)]
pf = rows(openpyxl.load_workbook(DATA/"reports"/"PlanFact_2025.xlsx", data_only=True)["План-факт"])
bd = rows(openpyxl.load_workbook(DATA/"reports"/"CashFlow_2025.xlsx", data_only=True)["БДДС"])
def line(tbl, label, nth=0):
    return [r for r in tbl if r[0] == label][nth]
rev, prof = line(pf, "Выручка"), line(pf, "Прибыль до налога")
rev_p, rev_f = line(pf, "Выручка, план")[1:13], line(pf, "Выручка, факт")[1:13]
prof_imp = line(pf, "Прибыль до налога", 1)[1:13]           # влияние на прибыль по месяцам, тыс. ₽ в рублях
cash_end = line(bd, "Остаток на конец периода")
mi = next(i for i, r in enumerate(bd) if r[0] == "Месяц")
cash_m = [r for r in bd[mi+1:mi+12] if isinstance(r[1], (int, float))]
cash_p = [r[1] for r in cash_m] + [cash_end[1]]; cash_f = [r[2] for r in cash_m] + [cash_end[2]]
skip = {"Выручка", "Итого себестоимость", "Валовая прибыль", "Итого управленческие", "Прибыль до налога", "EBITDA"}
items = [(r[0], r[4]) for r in pf[5:27] if r[0] and r[0] not in skip and isinstance(r[4], (int, float)) and r[4]]
top = sorted(items, key=lambda x: -abs(x[1]))[:4]
EXTRA = ({
 "выручка факт 825 211 796": (round(rev[2]), 825_211_796), "прибыль факт 158 046 156": (round(prof[2]), 158_046_156),
 "прибыль план 179 563 351": (round(prof[1]), 179_563_351), "остаток на конец года факт 216 038 097": (round(cash_end[2]), 216_038_097),
 "12 месяцев выручки": (len(rev_p)+len(rev_f), 24), "12 точек остатка": (len(cash_p)+len(cash_f), 24),
 "сумма месяцев выручки = год": (round(sum(rev_f)), 825_211_796),
 "сумма месяцев влияния = отклонение прибыли": (round(sum(prof_imp)), round(prof[3])),
})
# --- P&L: мост валовой маржи (FactorAnalysis) и таблица план-факт (PlanFact)
fa = rows(openpyxl.load_workbook(DATA/"reports"/"FactorAnalysis_2025.xlsx", data_only=True)["Валовая маржа"])
lab = lambda r: (r[0] or "").strip()
i0 = next(i for i, r in enumerate(fa) if lab(r) == "Плановая валовая маржа")
i1 = next(i for i, r in enumerate(fa) if lab(r) == "Фактическая валовая маржа")
gm_plan, gm_fact = fa[i0][1], fa[i1][1]
effects = [(lab(r), r[1]) for r in fa[i0+1:i1]]
bridge = [dict(n="Плановая валовая маржа", v=gm_plan, k="t")] + [dict(n=n, v=v, k="d") for n, v in effects] + [dict(n="Фактическая валовая маржа", v=gm_fact, k="t")]
pl = []
for r in pf[5:27]:
    if r[0] and r[0] != "Статья":
        pl.append(dict(n=r[0], p=r[1], f=r[2], imp=r[4], tot=r[0].startswith(("Итого", "Валовая", "Прибыль", "EBITDA", "Выручка")), head=r[1] is None))
gap = rev and line(pf, "Валовая прибыль")[2] - gm_fact
EXTRA.update({
 "валовая маржа план 291 428 431": (round(gm_plan), 291_428_431), "валовая маржа факт 268 103 392": (round(gm_fact), 268_103_392),
 "мост сходится (план + эффекты = факт, допуск 2 ₽)": (round(abs(gm_plan + sum(v for _, v in effects) - gm_fact) <= 2), 1),
 "9+ эффектов моста": (len(effects) >= 9, True),
})
# --- Деньги: ДДС, БДДС, дневной остаток, ликвидность (CashFlow_2025), без стресс-колонок
cfw = openpyxl.load_workbook(DATA/"reports"/"CashFlow_2025.xlsx", data_only=True)
dds_t = rows(cfw["ДДС"]); liq_t = rows(cfw["Ликвидность"]); day_t = rows(cfw["Ликвидность дн"])[1:]
dl = lambda lbl: next(r for r in dds_t if r[0] == lbl)
dds_rows = [dict(n=r[0], m=r[1:13], tot=r[13], head=r[1] is None,
                 tot_row=r[0].startswith(("Итого", "Чистый", "Чистое", "Остаток")))
            for r in dds_t[4:24] if r[0]]
bi = next(i for i, r in enumerate(bd) if r[0] == "Статья")
bdds_rows = [dict(n=r[0], p=r[1], f=r[2], d=r[3], head=r[1] is None,
                  tot_row=r[0].startswith(("Итого", "Чистый", "Чистое", "Остаток")))
             for r in bd[bi+1:bi+21] if r[0]]
liq = [dict(m=r[0], close=r[4], minday=r[5], mindate=r[6].strftime("%d.%m"), days_min=r[9], status=r[10])
       for r in liq_t[6:18]]
daily = [r[5] for r in day_t]; dmin = min(range(len(daily)), key=lambda i: daily[i])
bnet = line(bd, "Чистое изменение денежных средств"); bclose = line(bd, "Остаток на конец периода")
bsum = lambda lbl: line(bd, lbl)
EXTRA_CASH = {
 "остаток на конец года план 251 776 637": (round(bclose[1]), 251_776_637),
 "чистое изменение денег факт 185 038 097": (round(bnet[2]), 185_038_097),
 "чистое изменение денег план 220 776 637": (round(bnet[1]), 220_776_637),
 "минимум дневного остатка 16 579 356": (round(daily[dmin]), 16_579_356),
 "минимум дня 10 января": (day_t[dmin][0].strftime("%d.%m"), "10.01"),
 "365 дней в дневном ряду": (len(daily), 365),
 "последний дневной остаток = остаток конца года": (round(daily[-1]), round(cash_end[2])),
 "сумма 12 месяцев ДДС = БДДС факт": (round(sum(dl("Чистое изменение денежных средств")[1:13])), round(bnet[2])),
 "ДДС итого за год = сумма месяцев": (round(dl("Итого поступления")[13]), round(sum(dl("Итого поступления")[1:13]))),
 "БДДС: поступления = пациенты + страховые (факт)": (round(bsum("Поступления от пациентов")[2] + bsum("Поступления от страховых компаний")[2]), round(bsum("Итого поступления")[2])),
 "12 строк ликвидности": (len(liq), 12),
 "месяцев не «ок» (факт) 0": (sum(l["status"] != "ок" for l in liq), 0),
 "остаток конца декабря в ликвидности = ДДС": (round(liq[-1]["close"]), round(cash_end[2])),
}
# --- Врачи: Doctors_2025 (листы «Врачи», «Грейды», «Калькулятор»: порог минимума)
dwb = openpyxl.load_workbook(DATA/"reports"/"Doctors_2025.xlsx", data_only=True)
dv = rows(dwb["Врачи"]); gr = rows(dwb["Грейды"]); kc = rows(dwb["Калькулятор"])
end_d = next(i for i, r in enumerate(dv) if r[0] == "Без врача")
docs = [dict(n=r[0], id=r[1], cl=r[2], dr=r[3], g=r[4], mo=r[6], rev=r[8], con=r[15], pm=r[19], rate=r[5]) for r in dv[5:end_d]]
dirname = {r[16]: r[0] for r in gr if len(r) > 16 and r[16] and r[1] == "Все грейды"}
net_g = [dict(g=r[1], n=r[2], mo=r[3], rev=r[4], con=r[7], pm=r[10], pc=r[13]) for r in gr[7:12]]
dir_g = [dict(d=r[0], g=r[1], n=r[2], mo=r[3], pm=r[10], pc=r[13], rel=r[14], note=r[15], all=r[1] == "Все грейды")
         for r in gr[15:40] if r[0] and r[0] != "Направление" and not (r[1] != "Все грейды" and r[2] == 0)]
dmin_thr = next(r[1] for r in kc if r[0] and str(r[0]).startswith("Минимальный вклад")) * 1000
below = [d for d in docs if d["pm"] < dmin_thr]
tot_d = next(r for r in dv if r[0] == "Врачи итого"); calc_below = next(r for r in kc if r[0] and str(r[0]).startswith("Врачей ниже"))[2]
EXTRA_DOC = {
 "врачей 43": (len(docs), 43), "врач-месяцев 479": (sum(d["mo"] for d in docs), 479),
 "выручка врачей 825 211 700": (round(sum(d["rev"] for d in docs)), 825_211_700),
 "вклад врачей = итог книги": (round(sum(d["con"] for d in docs)), round(tot_d[15])),
 "вклад врачей 361 131 372": (round(tot_d[15]), 361_131_372),
 "вклад на врача-месяц 753 928": (round(tot_d[19]), 753_928),
 "грейды: 43 врача": (sum(g["n"] for g in net_g[:4]), 43), "грейды: 479 врач-месяцев": (sum(g["mo"] for g in net_g[:4]), 479),
 "грейды: вклад = итог": (round(sum(g["con"] for g in net_g[:4])), round(tot_d[15])),
 "направления: 43 врача": (sum(r["n"] for r in dir_g if r["all"]), 43),
 "врачей ниже минимума = калькулятор (3)": (len(below), calc_below),
 "порог минимума 500 тыс.": (round(dmin_thr), 500_000),
 "врачей с 4-8 месяцами 6": (sum(d["mo"] < 12 for d in docs), 6),
}
# --- Пациенты и прогноз: PatientLTV_2025 (всё в рублях, заголовки книги «тыс. ₽» это формат отображения) и Forecast_2026Q1
lw = openpyxl.load_workbook(DATA/"reports"/"PatientLTV_2025.xlsx", data_only=True)
fw = openpyxl.load_workbook(DATA/"reports"/"Forecast_2026Q1.xlsx", data_only=True)
lt_d = rows(lw["Описание"]); lt_c = rows(lw["Когорты"]); lt_s = rows(lw["Срезы"])
num = lambda v: v if isinstance(v, (int, float)) else None
cohorts = [dict(n=r[0].split(". ")[1].split()[0], p=r[1], obs=r[2], ltv=[num(x) for x in r[3:15]]) for r in lt_c[7:19]]
pool = dict(p=lt_c[19][1], ltv=[num(x) for x in lt_c[19][3:15]])
def lt_ctl(prefix): return next(r for r in lt_d if r[0] and str(r[0]).startswith(prefix))[1]
new_pat, new_rev, others_rev = lt_ctl("Новых пациентов"), lt_ctl("Выручка новых"), lt_ctl("Справка")
def slices(title):
    i = next(k for k, r in enumerate(lt_s) if r[0] == title) + 2; out = []
    while True:
        r = lt_s[i]; out.append(dict(n=r[0], p=r[1], ltv=r[2], v=r[3], ret=r[4], rev=r[5], note=r[6], all=r[0] == "Все"))
        if r[0] == "Все": return out
        i += 1
sl = dict(dir=slices("По направлению первого визита"), cl=slices("По клинике первого визита"), ch=slices("По каналу привлечения"))
sl_k, sl_c = lt_s[1][1], lt_s[1][2]
ft = rows(fw["Прогноз"]); fb = rows(fw["Откат"]); fm_ = rows(fw["Данные направлений"])
fc = [dict(n=r[0], iv=r[1], m=r[2:5], q=r[5], g=r[6], err=r[7], lo=r[8], hi=r[9], rel=r[10], tot=r[0].startswith("Итого")) for r in ft[13:19]]
meth = [dict(n=r[0], mae=r[1], bias=r[2]) for r in fb[4:8]]
net_m = [sum(r[1:6]) for r in fm_[1:13]]
EXTRA_PAT = {
 "новых пациентов 27 924": (new_pat, 27_924), "сумма когорт = новые пациенты": (sum(c["p"] for c in cohorts), 27_924),
 "выручка новых 494 762 800": (round(new_rev), 494_762_800), "12 когорт": (len(cohorts), 12),
 "LTV12 январской когорты 34 101": (round(cohorts[0]["ltv"][11]), 34_101), "LTV12 январской = пул": (round(cohorts[0]["ltv"][11]), round(pool["ltv"][11])),
 "когорта k наблюдается до 12-k (треугольник)": (sum(sum(x is not None for x in c["ltv"]) for c in cohorts), sum(c["obs"] for c in cohorts)),
 "срезы: пациентов по направлениям = по клиникам = по каналам": (sum(r["p"] for r in sl["dir"] if not r["all"]) == sl["dir"][-1]["p"] == sum(r["p"] for r in sl["cl"] if not r["all"]) == sum(r["p"] for r in sl["ch"] if not r["all"]), True),
 "выручка 2025 по направлениям = 825 211 700": (round(sum(net_m)), 825_211_700),
 "IV квартал факт 287 233 000": (round(sum(net_m[9:12])), 287_233_000), "IV кв. по направлениям = сеть": (round(sum(r["iv"] for r in fc if not r["tot"])), round(fc[-1]["iv"])),
 "прогноз I кв. 322 616 316": (round(fc[-1]["q"]), 322_616_316), "прогноз: сумма направлений = итог": (round(sum(r["q"] for r in fc if not r["tot"])), round(fc[-1]["q"])),
 "прогноз: три месяца = квартал": (round(sum(sum(r["m"]) - r["q"] for r in fc)), 0),
 "надёжность по порогам 15%/30%": ([("высокая" if r["err"] < .15 else "средняя" if r["err"] <= .30 else "низкая") == r["rel"] for r in fc].count(True), 6),
 "Хольт лучший метод отката": (min(meth, key=lambda m: m["mae"])["n"], "Хольт с затуханием"),
}
# --- Стресс: задержка страховых (live в браузере). Исходные ряды: «Данные дн» (день x клиника) и «Остатки» (ОСВ 51).
import pandas as pd
sd = pd.DataFrame(rows(cfw["Данные дн"])[1:], columns=["d", "c", "inf", "out", "ins"])
sop = {r[0]: r[1] for r in rows(cfw["Остатки"])[1:]}
def py_stress(cl, delay, res=10):
    x = sd[sd.c.isin(cl)].groupby("d")[["inf", "out", "ins"]].sum().sort_index()
    bal = sum(sop[c] for c in cl) + (x.inf + x.out).cumsum()
    st = bal - (x.ins.rolling(delay, min_periods=1).sum() if delay > 0 else x.ins * 0)
    mon = x.index.month; bad, gap, dmin = [], [], []
    for k in range(1, 13):
        mk = mon == k; ao = -x.out[mk].sum() / mk.sum(); sm = st[mk].min(); dmin.append(sm / ao)
        if sm < 0: gap.append(k)
        if sm < 0 or sm / ao < res: bad.append(k)
    return dict(bal=bal, st=st, min=float(st.min()), date=st.idxmin().strftime("%d.%m"), days=min(dmin), bad=bad, gap=gap)
STRESS_CL = sorted(sop)
cl_ser = {c: dict(net=[round(v, 2) for v in (g.inf + g.out)], out=[round(v, 2) for v in g.out], ins=[round(v, 2) for v in g.ins])
          for c, g in ((c, sd[sd.c == c].sort_values("d")) for c in STRESS_CL)}
stress_data = dict(open=sop, mon=[d.month - 1 for d in sorted(sd.d.unique())], cl=cl_ser)
ref30 = py_stress(STRESS_CL, 30)
book_L = [r[11] for r in liq_t[6:18]]                                 # «Стресс: остаток на конец месяца» (задержка 30)
mend = lambda ser: [ser[ser.index.month == k].iloc[-1] for k in range(1, 13)]
EXTRA_STRESS = {
 "стресс 30 дней: остаток на конец года = книга": (round(ref30["st"].iloc[-1]), round(book_L[-1])),
 "стресс 30 дней: 12 концов месяцев = книга (допуск 1 ₽)": (all(abs(a - b) < 1 for a, b in zip(mend(ref30["st"]), book_L)), True),
 "стресс 30 дней: минимум года = книга": (round(ref30["min"]), round(liq_t[20][2])),
 "стресс 0 дней = факт": (round(py_stress(STRESS_CL, 0)["min"]), round(daily[dmin])),
 "дневных рядов 3 клиники x 365": (sum(len(v["net"]) for v in cl_ser.values()), 3 * 365),
 "сумма клиник по дням = сеть (net)": (round(sum(sum(v["net"]) for v in cl_ser.values())), round(sum(r[2] + r[3] for r in day_t))),
}
# --- проверка п. 29: итоги из таблицы = итоги из документа
found = sum(e[3] for e in E)
subtle = [e for e in E if e[2]]; plain = [e for e in E if not e[2]]
chk = {
 "найдено 10 из 14": (found, 10), "тонких найдено 4 из 6": (sum(e[3] for e in subtle), 4),
 "тонких всего 6": (len(subtle), 6), "обычных найдено 6 из 8": (sum(e[3] for e in plain), 6),
 "обычных всего 8": (len(plain), 8), "причина названа у 6": (sum(e[4]=="да" for e in E), 6),
 "верных+ложных=находок (1)": (RUN1["correct"]+RUN1["false"], RUN1["findings"]),
 "верных+ложных=находок (2)": (RUN2["correct"]+RUN2["false"], RUN2["findings"]),
 "точность 1-го прогона 79%": (round(100*RUN1["correct"]/RUN1["findings"]), 79),
}
chk.update(EXTRA)
chk.update(EXTRA_CASH)
chk.update(EXTRA_DOC)
chk.update(EXTRA_PAT)
chk.update(EXTRA_STRESS)
bad = {k: v for k, v in chk.items() if v[0] != v[1]}
if bad: sys.exit(f"Сборка остановлена, числа не сошлись: {bad}")

data = dict(
  errors=[dict(id=a, name=b, subtle=bool(c), first=bool(d), cause=f) for a,b,c,d,f in E],
  run1=RUN1, run2=RUN2, hash1=HASH1, hash2=HASH2,
  found1=found, subtle_found=sum(e[3] for e in subtle), plain_found=sum(e[3] for e in plain),
  recon=dict(found="14 из 14", residual="−96 ₽"),
  pnl=dict(bridge=bridge, rows=pl, gap=gap),
  pat=dict(cohorts=cohorts, pool=pool, new=new_pat, rev=new_rev, others=others_rev, sl=sl, k=sl_k, kc=sl_c,
           fc=fc, meth=meth, net_m=net_m, tot_rev=sum(net_m)),
  doc=dict(docs=docs, net=net_g, dirs=dir_g, below=[d["id"] for d in below], thr=dmin_thr, names=dirname),
  cash=dict(dds=dds_rows, bdds=bdds_rows, liq=liq, daily=daily, dmin=dmin, dmin_date=day_t[dmin][0].strftime("%d.%m"),
            stress=stress_data, days_at_min=liq[0]["days_min"], start=dl("Остаток на начало периода")[1]),
  ov=dict(rev=[rev[1], rev[2]], prof=[prof[1], prof[2]], cash=[cash_end[1], cash_end[2]],
          rev_p=rev_p, rev_f=rev_f, prof_imp=prof_imp, cash_p=cash_p, cash_f=cash_f, top=top),
)

HTML = (pathlib.Path(__file__).with_name("dashboard_template.html").read_text(encoding="utf-8")
        if pathlib.Path(__file__).with_name("dashboard_template.html").exists() else None)
if HTML is None: sys.exit("нет dashboard_template.html рядом со скриптом")
out = OUT
out.write_text(HTML.replace("/*DATA*/null", json.dumps(data, ensure_ascii=False)), encoding="utf-8")
print("OK", out, "| проверок:", len(chk))
