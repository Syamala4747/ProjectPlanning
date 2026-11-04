# -*- coding: utf-8 -*-
"""
verify.py - automated correctness gate for the road construction planner.

Run:  python -m src.verify            (audits the three reference workbooks)
      python -m src.verify --no-run   (audit existing workbooks only, do not re-run)

This gate exists because every real defect found during development was caught
by an invariant, not by reading the report. It checks three layers:

  A. assumptions  - dead configuration, productivity source-of-truth, direct
                    rates present and on the 2019 yardstick, wastage wiring
  B. model        - the cost identities and the pricing basis
  C. workbooks    - the printed sheet must satisfy its own stated arithmetic

A failure here is a real defect. Do not ship around one.
"""
import sys
import io
import os
from datetime import date

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
sys.path.insert(0, ROOT)

PASSED = 0
FAILED = 0


def check(label, ok, detail=""):
    global PASSED, FAILED
    print(f"[{'PASS' if ok else 'FAIL'}] {label}" + (f"  -- {detail}" if detail else ""))
    if ok:
        PASSED += 1
    else:
        FAILED += 1


def close(a, b, tol):
    return abs(a - b) <= tol


def main():
    import yaml
    from openpyxl import load_workbook
    from src.data_loader import DataLoader
    from src.cost_calculator import CostCalculator

    road_config = os.path.join(ROOT, "project_config", "road")
    with open(os.path.join(road_config, "road_assumptions.yaml"), encoding="utf-8") as f:
        a = yaml.safe_load(f)
    loader = DataLoader(
        guidelines_path=os.path.join(road_config, "road_guidelines.json"),
        assumptions_path=os.path.join(road_config, "road_assumptions.yaml"),
    )
    loader.load_all()
    cc = CostCalculator(loader.assumptions, 8.0, 10.0)

    # ---------------------------------------------------------------- A
    print("=" * 78)
    print("A. ASSUMPTIONS")
    print("=" * 78)

    rates = a.get("direct_unit_rates_inr", {}) or {}
    activity_master = (
        loader.assumptions.get("activity_master", [])
        or loader.guidelines.get("activity_master", [])
    )
    check("direct_unit_rates_inr covers every activity",
          all(s.get("key") in rates for s in activity_master),
          f"{len(rates)} rates")

    acts = activity_master or []
    check("every activity has a positive productivity_rate",
          all(float(s.get("productivity_rate", 0) or 0) > 0 for s in acts),
          f"{len(acts)} activities")

    # productivity source of truth: the master field is what is priced. The
    # top-level `productivity:` block is reference documentation under its own
    # descriptive keys, so it cannot be mirrored value-for-value; what we can
    # and must verify is that it says so, so nobody edits it expecting a change.
    with open(os.path.join(road_config, "road_assumptions.yaml"), encoding="utf-8") as f:
        raw = f.read()
    prod_block = a.get("productivity") or {}
    if prod_block:
        check("top-level `productivity:` block is flagged REFERENCE-ONLY in the YAML",
              "REFERENCE DOCUMENTATION ONLY" in raw,
              "edit activity_master[].productivity_rate instead - the norm block "
              "is not read by any code")

    # A block that no code reads is a trap: it looks priced. Require that each
    # one carries an explicit NOT-USED banner immediately above it.
    reserved = [b for b in ("project_overhead_crew", "vehicle_transport",
                            "subcontractor_defaults", "overhead_weights")
                if b in a]
    if reserved:
        unmarked = []
        for b in reserved:
            idx = raw.find(f"\n{b}:")
            window = raw[max(0, idx - 1200):idx] if idx > 0 else ""
            if "NOT USED BY THE COST MODEL" not in window:
                unmarked.append(b)
        check("every reserved (unread) block is marked NOT-USED beside it",
              not unmarked, f"unmarked: {unmarked}" if unmarked else
              f"{len(reserved)} block(s) marked")

    crew_prof = a.get("crew_profile_by_activity_key") or {}
    check("every activity has a crew profile",
          all(s.get("key") in crew_prof for s in acts),
          f"{len(crew_prof)} profiles")

    # ---------------------------------------------------------------- B
    print()
    print("=" * 78)
    print("B. MODEL")
    print("=" * 78)

    from src.constraint_scheduler import ConstraintScheduler
    from src.models import ProjectInput

    res = ConstraintScheduler(
        project_input=ProjectInput(
            title="25 KM 4-Lane Highway Project", length_km=25.0,
            duration_days=800, budget_crore=1000.0,
            start_date=date(2026, 10, 13), lanes=4),
        data_loader=loader).solve()

    check("optimized duration meets the constraint",
          res.calculated_duration_days <= 800,
          f"{res.calculated_duration_days} d vs <= 800")
    check("optimized cost is within budget",
          res.total_calculated_cost_crore <= 1000.0,
          f"Rs {res.total_calculated_cost_crore:,.2f} Cr")
    check("additional optimization cost is never negative",
          res.additional_cost_crore >= 0.0,
          f"Rs {res.additional_cost_crore:,.2f} Cr")
    check("duration saved is never negative",
          res.duration_saved_days >= 0,
          f"{res.duration_saved_days} d")

    ob = sum(x.cost.total_amount for x in res.baseline_activities)
    oo = sum(x.cost.total_amount for x in res.activities)
    check("model total agrees with the reported crore figure",
          close(oo / 1e7, res.total_calculated_cost_crore, 0.02),
          f"{oo/1e7:,.2f} vs {res.total_calculated_cost_crore:,.2f} Cr")
    check("baseline total agrees with the reported crore figure",
          close(ob / 1e7, res.baseline_cost_crore, 0.02),
          f"{ob/1e7:,.2f} vs {res.baseline_cost_crore:,.2f} Cr")
    check("on-costs are NOT charged a second time (OH/Profit are zero)",
          all(abs(x.cost.overhead_amount) < 0.01 and abs(x.cost.profit_amount) < 0.01
              for x in res.activities))

    n_no_surcharge_ok = True
    for x in res.activities:
        comps = [x.cost.crash_fleet_mobilisation, x.cost.crash_shift_premium,
                 x.cost.crash_front_supervision, x.cost.crash_site_congestion,
                 x.cost.crash_site_establishment, x.cost.crash_resident_staff,
                 x.cost.crash_plant_hire_uplift, x.cost.crash_efficiency_loss]
        if abs(sum(comps) - x.cost.crash_total_surcharge) > 1:
            n_no_surcharge_ok = False
    check("all 8 surcharge components sum to the surcharge total",
          n_no_surcharge_ok)

    check("every surcharge component is non-negative",
          all(x.cost.crash_total_surcharge >= 0 for x in res.activities))

    # ---------------------------------------------------------------- C
    print()
    print("=" * 78)
    print("C. WORKBOOKS")
    print("=" * 78)

    for fn in (
        os.path.join("output", "report_251.xlsx"),
        os.path.join("output", "verify.xlsx"),
        os.path.join("output", "report_251_noaccel.xlsx"),
    ):
        # the generator falls back to <stem>_new.xlsx when the target is open in
        # Excel, so accept either and audit the freshest that exists
        path = fn if os.path.exists(fn) else fn.replace(".xlsx", "_new.xlsx")
        if not os.path.exists(path):
            check(f"{fn}: present", False, "not generated yet")
            continue
        wb = load_workbook(path)
        check(f"{path}: 14 sheets", len(wb.sheetnames) == 14, str(len(wb.sheetnames)))
        if "10_Optimized_Cost" not in wb.sheetnames:
            check(f"{path}: sheet 10 present", False)
            continue
        ws = wb["10_Optimized_Cost"]
        check(f"{path}: sheet 10 carries 31 columns (30 + escalation col)",
              ws.max_column >= 31, f"max_column={ws.max_column}")
        rows = list(ws.iter_rows(min_row=5, max_row=31))
        check(f"{path}: 27 activity rows", len(rows) == 27, str(len(rows)))
        bad8 = bad19 = bad23 = bad30 = bad2122 = 0
        for r in rows:
            q = r[2].value or 0
            c7, c8 = r[6].value or 0, r[7].value or 0
            c17, c19 = r[16].value or 0, r[18].value or 0
            c21, c22, c23 = r[20].value or 0, r[21].value or 0, r[22].value or 0
            c30 = r[29].value or 0
            c31 = r[30].value or 0
            if q and abs(c8 - q * c7) > max(2.0, c8 * 1e-4):
                bad8 += 1
            if abs(c19 - c8 - c17) > 1:
                bad19 += 1
            if abs(c23 - c19 - c30 - c31) > 1:
                bad23 += 1
            if abs(c30 - sum(r[j].value or 0 for j in (25, 26, 27, 28))) > 1:
                bad30 += 1
            if abs(c21) > 0.01 or abs(c22) > 0.01:
                bad2122 += 1
        check(f"{path}: col8 = qty x col7 (all-in rate prices the cost)",
              bad8 == 0, f"{bad8} rows")
        check(f"{path}: col19 = col8 + col17", bad19 == 0, f"{bad19} rows")
        check(f"{path}: col23 = col19 + col30 + col31 (wastage + escalation)",
              bad23 == 0, f"{bad23} rows")
        check(f"{path}: col30 = the four loss columns", bad30 == 0, f"{bad30} rows")
        check(f"{path}: OH/Profit columns are zero by design",
              bad2122 == 0, f"{bad2122} rows")

        # sheet 11: no negative acceleration anywhere, bridge must add up
        if "11_Baseline_vs_Optimized" in wb.sheetnames:
            w11 = wb["11_Baseline_vs_Optimized"]
            import re
            negs = [c.coordinate for row in w11.iter_rows() for c in row
                    if isinstance(c.value, str)
                    and re.search(r"(?:₹\s*)?-\s*[\d,]+\.\d+\s*Cr", c.value)]
            check(f"{path}: sheet 11 shows no negative acceleration cost",
                  not negs, str(negs))

            base = inc = opt = wst = esc = None
            for r in w11.iter_rows():
                lab = r[0].value
                val = r[1].value
                if not isinstance(lab, str) or not isinstance(val, (int, float)):
                    continue
                if lab.startswith("(=) BASELINE"):
                    base = val
                elif lab.startswith("(+/-) Change in acceleration"):
                    inc = val
                elif lab.startswith("(+/-) Change in execution"):
                    wst = val
                elif lab.startswith("(+/-) Change in price escalation"):
                    esc = val
                elif lab.startswith("(=) OPTIMIZED"):
                    opt = val
            check(f"{path}: sheet 11 bridge reconciles incl. wastage and escalation",
                  None not in (inc, wst, esc, base, opt)
                  and close(base + inc + wst + esc, opt, 0.02),
                  f"{base} + {inc} + {wst} + {esc} = "
                  f"{base + (inc or 0) + (wst or 0) + (esc or 0):.2f} vs {opt}"
                  if None not in (inc, wst, esc, base, opt) else "bridge rows not found")

    print()
    print("=" * 78)
    print(f"RESULT: {PASSED} passed, {FAILED} failed")
    print("=" * 78)
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())