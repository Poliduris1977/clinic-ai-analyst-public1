import subprocess,sys,time
steps=["src/generate/generate_reference.py","src/generate/generate_operations.py","src/generate/generate_1c.py","src/generate/generate_budget.py","src/generate/inject_errors.py","src/reconcile/reconcile.py","src/reconcile/apply_corrections.py","src/reconcile/reconcile.py --input corrected --out output_corrected","src/reports/build_pnl.py","src/reports/build_planfact.py","src/reports/build_factor.py","src/reports/build_doctors.py","src/reports/build_cashflow.py","src/reports/build_ltv.py","src/reports/build_forecast.py","src/blind/inject_errors_v2.py","src/blind/compare_runs.py","src/dashboard/build_dashboard.py data/reports/Dashboard_2025.html","src/agent/build_agent.py","src/agent/test_agent.py","src/check_controls.py","src/sql/run_sql.py"]
for s in steps:
    t=time.time()
    r=subprocess.run([sys.executable]+s.split(),capture_output=True,text=True,encoding="utf-8",errors="replace")
    tail=(r.stdout.strip().splitlines() or [""])[-1]
    print(f"rc={r.returncode} {time.time()-t:5.1f}s {s} | {tail[:100]}",flush=True)
    if r.returncode and "check_controls" not in s and "test_agent" not in s and "run_sql" not in s:
        print(r.stderr[-800:]); break
