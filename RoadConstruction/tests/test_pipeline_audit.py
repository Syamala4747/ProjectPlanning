# -*- coding: utf-8 -*-
"""Regression tests for the 89-activity pipeline / water-supply catalogue.

These lock in the invariants that are easy to break when the guideline
catalogue or the assumptions file is edited:

  * every catalogue key gets a quantity, a positive all-in rate and an activity
  * quantity x rate, surcharge, wastage and total-cost identities reconcile
  * zero-quantity (out-of-scope) activities carry zero duration/cost/weight
  * CPM ES/EF/LS/LF/TF/FF are internally consistent and no zero-duration
    pass-through node is reported as a critical work activity
  * the workbook's activity, cost and bridge sheets agree with the model
  * declared trenchless crossings do not create open-cut road quantities
"""
import json
import os
import sys
from datetime import date

import openpyxl
import pytest
import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.models import ProjectInput
from src.data_loader import DataLoader
from src.constraint_scheduler import ConstraintScheduler
from src.excel_report import ExcelReportGenerator
from src.quantity_calculator import PipelineQuantityCalculator

GUIDELINES = os.path.join(ROOT, "project_config", "pipeline", "pipeline_guidelines.json")
ASSUMPTIONS = os.path.join(ROOT, "project_config", "pipeline", "pipeline_assumptions.yaml")

SCENARIOS = {
    # Transmission corridor only: every facility package is out of scope.
    "corridor": dict(
        title="Regression corridor", length_km=10.0, pipe_diameter_mm=600,
        pipe_material="DI", trench_depth_m=2.0, number_of_work_fronts=2,
        duration_days=540, budget_crore=15.0,
    ),
    # Multi-village scheme: village count derives OHT/FHTC/WTP/IBP.
    "scheme": dict(
        title="Regression 10-village scheme", length_km=18.0, pipe_diameter_mm=300,
        pipe_material="DI", trench_depth_m=2.0, number_of_work_fronts=3,
        duration_days=900, budget_crore=60.0, num_villages=10,
    ),
}


@pytest.fixture(scope="module")
def runs():
    """Run each scenario once and cache the scheduler + result."""
    out = {}
    for name, kw in SCENARIOS.items():
        pi = ProjectInput(project_type="pipeline", start_date=date(2026, 1, 1), **kw)
        loader = DataLoader(guidelines_path=GUIDELINES, assumptions_path=ASSUMPTIONS)
        loader.load_all(terrain=pi.terrain)
        sched = ConstraintScheduler(pi, loader, enable_acceleration_cost=True)
        out[name] = (sched, sched.solve())
    return out


def _guideline_keys():
    with open(GUIDELINES, "r", encoding="utf-8") as fh:
        return [a["key"] for a in json.load(fh)["activity_master"]]


def test_every_catalogue_key_is_quantified_and_priced(runs):
    keys = set(_guideline_keys())
    assert len(keys) == 89

    for name, (sched, res) in runs.items():
        q = sched.quantities
        acts = res.activities
        assert set(q) == keys, f"{name}: quantity keys differ from catalogue"
        assert {a.activity_key for a in acts} == keys, f"{name}: activity keys differ"
        assert len(acts) == len(keys)

        for a in acts:
            c = a.cost
            assert c.planned_unit_rate > 0, f"{name}: {a.activity_key} has no rate"
            # Planned rate is DERIVED from the published source rate plus
            # wastage/idle losses and statutory OH+CP. It must not be a copy.
            assert c.planned_unit_rate > c.morth_direct_unit_rate, (
                f"{name}: {a.activity_key} planned rate was not built up"
            )
            assert c.planning_variation_pct > 0, (
                f"{name}: {a.activity_key} planning variation not reported"
            )
            ratio = c.planned_unit_rate / c.morth_direct_unit_rate
            assert 1.05 <= ratio <= 1.60, (
                f"{name}: {a.activity_key} planned/source ratio {ratio:.3f} out of band"
            )
            assert c.total_amount == pytest.approx(
                c.direct_cost + c.overhead_amount + c.profit_amount + c.escalation_amount
            ), f"{name}: {a.activity_key} total identity"
            assert c.direct_cost == pytest.approx(
                a.quantity * c.planned_unit_rate
                + c.crash_total_surcharge
                + c.wastage_cost
            ), f"{name}: {a.activity_key} direct identity"
            # Direct all-in rates already contain wastage, OH and profit.
            assert c.wastage_in_base_rate
            for field in ("wastage_cost", "material_wastage_cost", "fuel_wastage_cost",
                          "plant_wastage_cost", "labour_wastage_cost",
                          "resource_wastage_cost"):
                assert getattr(c, field) == 0, f"{name}: {a.activity_key} {field} != 0"
            assert c.overhead_amount == 0 and c.profit_amount == 0
            if a.quantity > 0:
                # On-costs are embedded in the planned rate and reported as memo
                # amounts; they are not charged a second time.
                assert c.overhead_embedded > 0, (
                    f"{name}: {a.activity_key} embedded overhead is missing"
                )
                assert c.profit_embedded > 0, (
                    f"{name}: {a.activity_key} embedded profit is missing"
                )
                assert (
                    c.overhead_embedded + c.profit_embedded
                    < a.quantity * c.planned_unit_rate
                )

            if a.quantity <= 0:
                assert a.duration_days == 0, f"{name}: {a.activity_key} zero-qty duration"
                assert c.total_amount == 0, f"{name}: {a.activity_key} zero-qty cost"
                assert a.progress_weight_pct == 0, f"{name}: {a.activity_key} zero-qty weight"


def test_cpm_and_progress_invariants(runs):
    for name, (_, res) in runs.items():
        acts = res.activities
        by_id = {a.activity_id: a for a in acts}
        duration = res.calculated_duration_days

        assert max(a.ef for a in acts) == duration, f"{name}: max EF != duration"
        for a in acts:
            assert a.ef - a.es == a.duration_days
            assert a.lf - a.ls == a.duration_days
            assert a.total_float == pytest.approx(a.ls - a.es)
            assert a.total_float >= 0
            for pid in a.predecessors:
                lag = int((a.predecessor_lags or {}).get(pid, 0))
                assert a.es >= by_id[pid].ef + lag - 1e-9
            if a.duration_days == 0:
                assert not a.is_critical, f"{name}: zero-duration {a.activity_id} critical"
            assert by_id[a.activity_id].duration_days >= 0

        for aid in res.critical_path:
            assert by_id[aid].duration_days > 0, (
                f"{name}: reported critical path contains zero-duration {aid}"
            )

        last = res.monthly_progress[-1]
        assert last.cumulative_expenditure_crore == pytest.approx(
            res.total_calculated_cost_crore, abs=0.05
        )
        assert last.cumulative_planned_progress_pct == pytest.approx(100.0, abs=0.2)
        assert sum(m.period_expenditure_crore for m in res.monthly_progress) == pytest.approx(
            res.total_calculated_cost_crore, abs=0.05
        )

        # Sheet-11 bridge identity.
        base = res.baseline_activities
        bridge = (
            res.baseline_cost_crore
            + (sum(a.cost.crash_total_surcharge for a in acts)
               - sum(a.cost.crash_total_surcharge for a in base)) / 1e7
            + (sum(a.cost.wastage_cost for a in acts)
               - sum(a.cost.wastage_cost for a in base)) / 1e7
            + (sum(a.cost.escalation_amount for a in acts)
               - sum(a.cost.escalation_amount for a in base)) / 1e7
        )
        assert bridge == pytest.approx(res.total_calculated_cost_crore, abs=0.02)


def test_workbook_reconciles_to_the_model(runs, tmp_path):
    _, res = runs["scheme"]
    out = tmp_path / "pipeline_scheme.xlsx"
    ExcelReportGenerator(res, str(out)).generate()
    wb = openpyxl.load_workbook(out, data_only=False)

    # Sheet 02: scheme scope is shown for a facility run.
    ws = wb["02_Project_Inputs"]
    vals = {str(r[0]): r[1] for r in ws.iter_rows(min_row=5, values_only=True) if r[0]}
    assert vals["Villages Served"] == res.pipeline_scope["villages"] == 10
    assert vals["Overhead Service Reservoirs (OHT/ESR)"] == res.pipeline_scope["ohts"] == 4
    assert vals["Water Treatment Plant Capacity"] == 5
    assert vals["Functional Household Tap Connections"] == res.pipeline_scope["fhtc"] == 2000
    assert vals["Intake / Booster Pumping Stations"] == res.pipeline_scope["ibps"] == 1

    by_id = {a.activity_id: a for a in res.activities}
    base_by_id = {a.activity_id: a for a in res.baseline_activities}

    # Sheet 05: quantity, baseline cost and optimized cost per activity.
    ws = wb["05_Activities"]
    seen = 0
    for row in ws.iter_rows(min_row=5, values_only=True):
        if row[2] and str(row[2]).startswith("ACT-"):
            aid, qty, bcost, ocost = row[2], row[4], row[9], row[13]
            assert float(qty) == pytest.approx(by_id[aid].quantity, abs=0.01)
            assert float(bcost) == pytest.approx(base_by_id[aid].cost.total_amount, abs=1.0)
            assert float(ocost) == pytest.approx(by_id[aid].cost.total_amount, abs=1.0)
            seen += 1
    assert seen == len(res.activities)

    # Sheet 10: optimized total, planned direct, surcharge and effective rate.
    ws = wb["10_Optimized_Cost"]
    headers = [c.value for c in ws[4]]
    col = {h: i for i, h in enumerate(headers)}
    ti = col["Total Activity Cost (₹)"]
    pdi = col["Planned Direct Cost (₹)"]
    sdi = col["Total Crashing Surcharge (₹)"]
    odi = col["Optimized Direct Cost (₹)"]
    ohi = col["Overhead (₹) — embedded in rate"]
    cpi = col["Profit (₹) — embedded in rate"]
    seen = 0
    for row in ws.iter_rows(min_row=5, values_only=True):
        if row[0] and str(row[0]).startswith("ACT-"):
            a = by_id[row[0]]
            assert float(row[ti]) == pytest.approx(a.cost.total_amount, abs=1.0)
            assert float(row[pdi]) == pytest.approx(
                a.quantity * a.cost.planned_unit_rate, abs=1.0)
            assert float(row[sdi]) == pytest.approx(a.cost.crash_total_surcharge, abs=1.0)
            assert float(row[odi]) == pytest.approx(a.cost.direct_cost, abs=1.0)
            assert float(row[ohi]) == pytest.approx(a.cost.overhead_embedded, abs=1.0)
            assert float(row[cpi]) == pytest.approx(a.cost.profit_embedded, abs=1.0)
            seen += 1
    assert seen == len(res.activities)

    # Sheet 11: the visible cost bridge closes to the optimized total.
    ws = wb["11_Baseline_vs_Optimized"]
    bridge = {}
    for row in ws.iter_rows(min_row=118, max_row=124, values_only=True):
        if row[0]:
            bridge[str(row[0])] = row[1]
    base_label = next(k for k in bridge if k.startswith("(=) BASELINE"))
    opt_label = next(k for k in bridge if k.startswith("(=) OPTIMIZED"))
    assert bridge[base_label] == pytest.approx(res.baseline_cost_crore, abs=0.01)
    assert bridge[opt_label] == pytest.approx(res.total_calculated_cost_crore, abs=0.01)
    deltas = sum(v for k, v in bridge.items()
                 if k.startswith("(+/-)") and isinstance(v, (int, float)))
    assert bridge[base_label] + deltas == pytest.approx(bridge[opt_label], abs=0.01)


def test_declared_crossings_do_not_create_open_cut_road_work():
    with open(ASSUMPTIONS, "r", encoding="utf-8") as fh:
        assumptions = yaml.safe_load(fh)
    inputs = assumptions.setdefault("pipeline_project_inputs", {})
    inputs.update({
        "number_of_road_crossings": 4,
        "number_of_highway_crossings": 2,
        "number_of_rail_crossings": 1,
        "number_of_river_crossings": 1,
        "number_of_hdd_crossings": 2,
    })
    pi = ProjectInput(project_type="pipeline", start_date=date(2026, 1, 1),
                      title="crossing check", duration_days=365, budget_crore=10.0,
                      length_km=10.0, pipe_diameter_mm=300, pipe_material="DI",
                      trench_depth_m=2.0)
    q = PipelineQuantityCalculator(pi, assumptions).calculate_quantities()
    assert q["dismantling_roads"]["quantity"] == pytest.approx(4 * 20.0)
    assert q["restoration_roads"]["quantity"] == pytest.approx(4 * 20.0)
    assert q["pipe_encasing"]["quantity"] == pytest.approx(4.0)
    assert q["highway_rail_crossing"]["quantity"] == pytest.approx(3 * 25.0)
    assert q["river_bridge_crossing"]["quantity"] == pytest.approx(60.0)


def test_intake_and_wtp_each_get_an_ht_substation():
    with open(ASSUMPTIONS, "r", encoding="utf-8") as fh:
        assumptions = yaml.safe_load(fh)
    inputs = assumptions.setdefault("pipeline_project_inputs", {})
    inputs["num_ibps"] = 2
    inputs["wtp_capacity_mld"] = 5.0
    pi = ProjectInput(project_type="pipeline", start_date=date(2026, 1, 1),
                      title="substation check", duration_days=365, budget_crore=10.0,
                      length_km=5.0, pipe_diameter_mm=300, pipe_material="DI",
                      trench_depth_m=2.0)
    q = PipelineQuantityCalculator(pi, assumptions).calculate_quantities()
    assert q["ht_substation"]["quantity"] == pytest.approx(3.0)  # 2 IBP + 1 WTP train
