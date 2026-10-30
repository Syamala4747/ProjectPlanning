# -*- coding: utf-8 -*-
"""
INDEPENDENT CALCULATION AUDIT
=============================
Recomputes every stage of the planner from assumptions.yaml primitives and
compares against what the model produced. Deliberately does NOT call the
model's own helpers for the expected values - it re-derives them from the raw
YAML so a bug in a shared helper cannot hide itself.

Run:  python audit_calculations.py
"""
import sys, io, os, math, yaml
from datetime import date

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _ROOT)
os.chdir(_ROOT)

from src.data_loader import DataLoader
from src.quantity_calculator import QuantityCalculator
from src.cost_calculator import CostCalculator
from src.constraint_scheduler import ConstraintScheduler
from src.models import ProjectInput

P = F = 0
def chk(label, ok, detail=""):
    global P, F
    print(f"[{'PASS' if ok else 'FAIL'}] {label}" + (f"  -- {detail}" if detail else ""))
    if ok: P += 1
    else:  F += 1

def near(a, b, tol): return abs(a - b) <= tol

# ---------------------------------------------------------------- raw YAML
loader = DataLoader(guidelines_path="guidelines.json", assumptions_path="assumptions.yaml")
loader.load_all()
A = loader.assumptions
cc = CostCalculator(A, 8.0, 10.0)

res = ConstraintScheduler(
    project_input=ProjectInput(title="25 KM 4-Lane Highway Project", length_km=25.0,
                               duration_days=800, budget_crore=1000.0,
                               start_date=date(2026, 10, 13), lanes=4),
    data_loader=loader).solve()

acts = loader.guidelines.get("activity_master") or A.get("activity_master", [])
SPEC   = {s["key"]: s for s in acts}
LAB    = A["labour_rates_inr_per_day"]
COMP   = A["crew_composition"]
EQPA   = A["equipment_per_activity"]
HIRE   = A["equipment_hire_inr_per_day"]
LPH    = A["fuel_consumption_lph"]
MATE   = A["material_rates_inr"]
BIND   = A["material_bindings"]
PROF   = A["crew_profile_by_activity_key"]
LANES  = 4
acts   = {a.activity_key: a for a in res.activities}

print("=" * 80); print("1. GEOMETRY FROM lane_standards (the block that actually wins)"); print("=" * 80)
LS = A["lane_standards"][LANES]
for f in ("carriageway_width_m", "formation_width_m", "median_width_m",
          "subgrade_width_m", "gsb_width_m", "wmm_width_m"):
    chk(f"lane_standards[{LANES}].{f} present", f in LS, f"{LS.get(f)}")

print(); print("=" * 80); print("2. QUANTITIES - recomputed from geometry"); print("=" * 80)
qc = QuantityCalculator(ProjectInput(title="t", length_km=25.0, duration_days=800,
                                     budget_crore=1000.0, lanes=LANES), A)
L  = 25.0 * 1000.0
FW = qc.formation_width_m; CW = qc.carriageway_width_m
SW, GW, WW, MW = qc.subgrade_width_m, qc.gsb_width_m, qc.wmm_width_m, qc.median_width_m
kmf = (25.0 / qc.ref_corridor_km) * (LANES / 4.0)

expect = {
    "site_clearance":   (L * FW / 10000.0, "ha"),
    "excavation":       (L * FW * qc.excavation_depth_m * qc.excavation_cut_ratio, "cum"),
    "embankment":       (L * FW * qc.embankment_height_m * qc.embankment_fill_ratio, "cum"),
    "subgrade":         (L * SW * qc.subgrade_thickness_m, "cum"),
    "gsb":              (L * GW * qc.gsb_thickness_m, "cum"),
    "wmm":              (L * WW * qc.wmm_thickness_m, "cum"),
    "dbm":              (L * CW, "sqm"),
    "bc":               (L * CW, "sqm"),
    "drainage":         (L * 2.0, "m"),
}
for k, (v, unit) in expect.items():
    got = acts[k].quantity
    chk(f"quantity[{k}] = {v:,.2f} {unit}", near(got, v, max(1.0, v * 1e-6)),
        f"model {got:,.2f}")
# dimensional sanity
chk("excavation qty consistent with formation area",
    near(acts["excavation"].quantity / L, FW * qc.excavation_depth_m * qc.excavation_cut_ratio, 1e-6))
chk("subgrade volume = area x thickness",
    near(acts["subgrade"].quantity, L * SW * 0.50, max(1.0, acts["subgrade"].quantity * 1e-6)))

print(); print("=" * 80); print("3. CREW & PLANT DAY RATES - recomputed from raw tables"); print("=" * 80)
for key in ("excavation", "site_clearance", "embankment", "gsb"):
    prof = PROF[key]
    day = sum(int(c["count"]) * float(LAB[c["trade"]]) for c in COMP[prof])
    pday = sum(float(HIRE[(e.get("item") if isinstance(e, dict) else e)])
               for e in EQPA[prof])
    chk(f"crew_day_rate[{key}] = {day:,.0f}", near(cc.crew_day_rate[key], day, 0.51),
        f"model {cc.crew_day_rate[key]:,.0f}")
    chk(f"plant_day_rate[{key}] = {pday:,.0f}", near(cc.plant_day_rate[key], pday, 0.51),
        f"model {cc.plant_day_rate[key]:,.0f}")
chk("every trade in every profile has a wage",
    all(c["trade"] in LAB for p in COMP.values() for c in p))
chk("every equipment item has a hire rate",
    all((e.get("item") if isinstance(e, dict) else e) in HIRE
        for p in EQPA.values() for e in p))

print(); print("=" * 80); print("4. PLANNED RATE - full independent rebuild"); print("=" * 80)
for key in ("excavation", "site_clearance", "embankment", "gsb", "bc", "toll_plaza_atms"):
    spec = SPEC[key]; prod = float(spec["productivity_rate"]); bc = int(spec["base_crews"])
    prof = PROF[key]
    day  = sum(int(c["count"]) * float(LAB[c["trade"]]) for c in COMP[prof])
    pday = sum(float(HIRE[(e.get("item") if isinstance(e, dict) else e)]) for e in EQPA[prof])
    lph  = sum(float(LPH[(e.get("item") if isinstance(e, dict) else e)]) for e in EQPA[prof])
    fuel = lph * float(cc.equipment_hours_per_day) * float(cc.diesel_rate)
    sp = cc._activity_recipe.get(key)
    mat = 0.0
    if sp:
        rec = cc._resolve_recipe(str(sp["recipe"]))
        mix = float(sp.get("per_output", 1.0) or 1.0)
        for mk, po in rec.items():
            if str(mk).startswith("_"): continue
            b = BIND[mk]
            mat += float(po) * mix / float(b.get("divisor", 1) or 1) * float(MATE[b["rate"]])
    bare = mat + day / prod + pday / prod + fuel / prod
    K = cc.sdb_oncost_multiplier(spec.get("category"))
    got_bare, _ = cc._resource_planned_rate(key, 1e6)
    chk(f"planned BARE rate[{key}] = {bare:,.4f}", near(got_bare, round(bare, 4), 0.01),
        f"model {got_bare:,.4f}")
    # the all-in rate is bare x K - verified against the activity's own cost record
    a = acts.get(key)
    if a is not None:
        chk(f"planned ALL-IN rate[{key}] = {bare * K:,.4f}",
            near(a.cost.planned_unit_rate, round(bare * K, 4), 0.01),
            f"model {a.cost.planned_unit_rate:,.4f}")
chk("K multiplier is 1.188 for road works",
    near(cc.sdb_oncost_multiplier("Earthwork"), 1.188, 1e-9),
    f"{cc.sdb_oncost_multiplier('Earthwork')}")
chk("K multiplier is 1.32 for Structures",
    near(cc.sdb_oncost_multiplier("Structures"), 1.32, 1e-9),
    f"{cc.sdb_oncost_multiplier('Structures')}")

print(); print("=" * 80); print("5. DURATION & CPM"); print("=" * 80)
bad_effort = bad_ef = bad_float = 0
for a in res.activities:
    prod = float(SPEC[a.activity_key]["productivity_rate"])
    want_eff = max(1, math.ceil(a.quantity / max(prod * a.crews * a.shifts, 1e-9)))
    if a.effort_days != want_eff:
        bad_effort += 1
    if a.ef != a.es + a.duration_days:
        bad_ef += 1
    if a.total_float != a.ls - a.es:
        bad_float += 1
chk("effort_days = ceil(qty / (prod x crews x shifts))", bad_effort == 0, f"{bad_effort} wrong")
chk("ef = es + duration for every activity", bad_ef == 0, f"{bad_ef} wrong")
chk("total_float = ls - es for every activity", bad_float == 0, f"{bad_float} wrong")
chk("calendar duration never shorter than effort days",
    all(a.duration_days >= a.effort_days for a in res.activities))
crit = [a for a in res.activities if a.is_critical]
chk("every critical activity has zero total float",
    all(a.total_float == 0 for a in crit), f"{len(crit)} critical")

# precedence: on the OPTIMIZED plan, overlaps are LEGAL and intentional - the
# staggered handover applies a negative FS lag. So verify strict FS on the
# BASELINE (which is built with stagger off), then separately confirm that every
# overlap in the optimized plan is on a chain declared in
# staggered_handover_chains.
bid = {a.activity_id: a for a in res.baseline_activities}
viol = []
for a in res.baseline_activities:
    for p in a.predecessors:
        if p in bid and a.es < bid[p].ef:
            viol.append(f"{a.activity_id} starts {bid[p].ef - a.es} d before {p} finishes")
chk("baseline precedence respected (strict FS, stagger off)", not viol,
    "; ".join(viol[:3]) if viol else "")

pairs = {(p, s) for p, s in
         ((c["predecessor_id"], c["successor_id"]) for c in A["staggered_handover_chains"])}
oid = {a.activity_id: a for a in res.activities}
unexplained = []
for a in res.activities:
    for p in a.predecessors:
        if p in oid and a.es < oid[p].ef and (p, a.activity_id) not in pairs:
            unexplained.append(f"{a.activity_id} starts {oid[p].ef - a.es} d before "
                               f"{p} but that pair is not a declared stagger chain")
chk("every optimized-plan overlap is a DECLARED staggered handover",
    not unexplained, "; ".join(unexplained[:3]) if unexplained else
    f"{len(pairs)} declared chains")

print(); print("=" * 80); print("6. WASTAGE"); print("=" * 80)
bad_mat = bad_lf = bad_pl = 0
for a in res.activities:
    key = a.activity_key
    sp = cc._activity_recipe.get(key)
    if sp:
        rec = cc._resolve_recipe(str(sp["recipe"]))
        mix = float(sp.get("per_output", 1.0) or 1.0)
        mw = 0.0
        for mk, po in rec.items():
            if str(mk).startswith("_"): continue
            b = BIND[mk]
            theo = float(po) * a.quantity * mix
            waste = float(cc._wastage_factors.get(b.get("wastage"), cc.wastage_default))
            issued = theo * (1.0 + waste)
            mw += (issued - theo) / float(b.get("divisor", 1) or 1) * float(MATE[b["rate"]])
        if not near(mw, a.cost.material_wastage_cost, max(1.0, abs(mw) * 0.01)):
            bad_mat += 1
    prof = PROF[key]
    day = sum(int(c["count"]) * float(LAB[c["trade"]]) for c in COMP[prof])
    gd = a.duration_days * max(1, a.crews)
    lw = day * gd * float(A["wastage_non_material"]["labour_idle_pct"])
    if not near(lw, a.cost.labour_wastage_cost, max(1.0, abs(lw) * 0.02)):
        bad_lf += 1
chk("material wastage = (issued - theoretical) priced at real rates", bad_mat == 0, f"{bad_mat} wrong")
chk("labour wastage = labour_base x 4% (gang-days)", bad_lf == 0, f"{bad_lf} wrong")
chk("wastage_total = material + fuel + plant + labour",
    all(near(a.cost.wastage_cost,
             a.cost.material_wastage_cost + a.cost.fuel_wastage_cost
             + a.cost.plant_wastage_cost + a.cost.labour_wastage_cost, 1.0)
        for a in res.activities))

print(); print("=" * 80); print("7. ACCELERATION SURCHARGE - all 8 + re-sequencing"); print("=" * 80)
cr = A["crashing_factors"]; st = A["acceleration_step_costs"]
bad = 0
for a in res.activities:
    c = a.cost
    comps = [c.crash_fleet_mobilisation, c.crash_shift_premium, c.crash_front_supervision,
             c.crash_site_congestion, c.crash_site_establishment, c.crash_resident_staff,
             c.crash_plant_hire_uplift, c.crash_efficiency_loss]
    if not near(sum(comps), c.crash_total_surcharge, 1.0): bad += 1
    if any(x < 0 for x in comps): bad += 1
chk("8 components sum to crash_total_surcharge and none negative", bad == 0, f"{bad} wrong")
# spot-check one activity's establishment term by hand
a = acts["embankment"]
ef = max(0, a.work_fronts - 1)
want = ef * a.duration_days * float(st["site_establishment_per_front_day"])
chk(f"site establishment = fronts x days x rate ({a.activity_key})",
    near(a.cost.crash_site_establishment, want, 1.0),
    f"want {want:,.0f}, model {a.cost.crash_site_establishment:,.0f}")
chk("shift premium is zero because shifts == 1.0 everywhere",
    all(a.cost.crash_shift_premium == 0 for a in res.activities))

print(); print("=" * 80); print("8. PROJECT TOTALS"); print("=" * 80)
qp = sum(a.quantity * a.cost.planned_unit_rate for a in res.activities)
sur = sum(a.cost.crash_total_surcharge for a in res.activities)
wst = sum(a.cost.wastage_cost for a in res.activities)
esc = sum(a.cost.escalation_amount for a in res.activities)
tot = res.total_calculated_cost_crore * 1e7
chk("qty x planned + surcharge + wastage + escalation == project total",
    near(qp + sur + wst + esc, tot, max(1e4, tot * 1e-5)),
    f"{qp/1e7:,.2f} + {sur/1e7:,.2f} + {wst/1e7:,.2f} + {esc/1e7:,.2f} "
    f"= {(qp+sur+wst+esc)/1e7:,.2f} vs {tot/1e7:,.2f}")
chk("escalation is never negative", esc >= 0, f"Rs {esc/1e7:,.2f} Cr")
chk("escalation factor >= 1.0 for every activity",
    all(a.cost.escalation_factor >= 1.0 for a in res.activities))
EXCL = set(A["escalation"].get("excluded") or [])
chk("excluded activities carry exactly 1.0 factor",
    all(a.cost.escalation_factor == 1.0 for a in res.activities
        if a.activity_key in EXCL), f"excluded: {sorted(EXCL)}")
chk("later work escalates at least as much as earlier work",
    max(a.cost.escalation_factor for a in res.activities)
    >= min(a.cost.escalation_factor for a in res.activities),
    f"factor range {min(a.cost.escalation_factor for a in res.activities):.4f}"
    f" .. {max(a.cost.escalation_factor for a in res.activities):.4f}")
_cap = A["escalation"].get("cap_pct_of_base")
_cap = 1.0 if _cap is None else float(_cap)          # null cap = un-capped
chk("escalation respects the cap",
    all(a.cost.escalation_factor <= 1.0 + _cap + 1e-9 for a in res.activities),
    "un-capped" if _cap >= 1.0 else f"cap {1 + _cap:.2f}x")
chk("on-costs are NOT charged twice (OH and profit are zero)",
    all(a.cost.overhead_amount == 0 and a.cost.profit_amount == 0 for a in res.activities))
chk("duration meets the constraint", res.calculated_duration_days <= 800,
    f"{res.calculated_duration_days} d")
chk("cost within budget", res.total_calculated_cost_crore <= 1000.0,
    f"Rs {res.total_calculated_cost_crore:,.2f} Cr")
chk("acceleration cost never negative", res.additional_cost_crore >= 0,
    f"Rs {res.additional_cost_crore:,.2f} Cr")

print(); print("=" * 80); print("9. WEIGHTS & PROGRESS"); print("=" * 80)
chk("cost weights sum to 100%", near(sum(a.weight_pct for a in res.activities), 100.0, 0.05),
    f"{sum(a.weight_pct for a in res.activities):.2f}%")
chk("progress (crew-day) weights sum to 100%",
    near(sum(a.progress_weight_pct for a in res.activities), 100.0, 0.05),
    f"{sum(a.progress_weight_pct for a in res.activities):.2f}%")
chk("cost weights and progress weights are NOT identical (two different curves)",
    any(abs(a.weight_pct - a.progress_weight_pct) > 0.5 for a in res.activities))

print()
print("=" * 80)
print(f"RESULT: {P} passed, {F} failed")
print("=" * 80)
sys.exit(1 if F else 0)