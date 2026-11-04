# -*- coding: utf-8 -*-
"""Regression tests for the 40-activity low-rise building catalogue.

These lock in the fixes that were missing from the building path:

  * all 40 semantic keys get a real quantity (no 100-unit fallback)
  * site area is BUA / FAR, not BUA / (FAR x floors)
  * CPWD DSR rates are treated as bare and the planned rate is built up with
    wastage, idle losses and statutory OH/profit
  * escalation reads annual_rates / base_date (and still understands the legacy
    annual_rate_pct key)
  * night-work keys are catalogue keys, not the obsolete bld_* template keys
  * the workbook reports embedded OH/CP instead of claiming they are in the
    source rate
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

from src.constraint_scheduler import ConstraintScheduler
from src.data_loader import DataLoader
from src.escalation import EscalationModel
from src.excel_report import ExcelReportGenerator
from src.models import ProjectInput
from src.cost_calculator import CostCalculator
from src.plan_generator import build_plan
from src.quantity_calculator import BuildingQuantityCalculator
from src.workbook_reader import WorkbookReader

GUIDELINES = os.path.join(ROOT, "project_config", "building", "building_guidelines.json")
ASSUMPTIONS = os.path.join(ROOT, "project_config", "building", "building_assumptions.yaml")


def _catalogue():
    with open(GUIDELINES, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data.get("activity_master") or (data.get("data") or {}).get("activity_master")


def _assumptions():
    with open(ASSUMPTIONS, "r", encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def _input(**overrides):
    kw = dict(
        title="Building regression", length_km=0.0, project_type="building",
        start_date=date(2026, 11, 3), duration_days=540, budget_crore=6.0,
        built_up_area_sqm=2000.0, number_of_floors=4,
        building_type="residential", construction_type="rcc",
        quality_finish="standard", number_of_blocks=1,
    )
    kw.update(overrides)
    return ProjectInput(**kw)


@pytest.fixture(scope="module")
def run():
    pi = _input()
    loader = DataLoader(guidelines_path=GUIDELINES, assumptions_path=ASSUMPTIONS)
    loader.load_all(terrain=pi.terrain)
    sched = ConstraintScheduler(pi, loader, enable_acceleration_cost=True)
    return sched, sched.solve(), loader


def test_every_catalogue_key_is_quantified_and_units_match(run):
    sched, res, _ = run
    catalogue = _catalogue()
    keys = {a["key"] for a in catalogue}
    units = {a["key"]: a["unit"] for a in catalogue}

    assert len(keys) == 40
    assert set(sched.quantities) == keys
    assert {a.activity_key for a in res.activities} == keys
    assert len(res.activities) == 40
    assert not sched._quantity_fallbacks, "building used a quantity fallback"

    for a in res.activities:
        q = sched.quantities[a.activity_key]
        assert q["unit"] == units[a.activity_key], (
            f"{a.activity_key}: quantity unit {q['unit']!r} != catalogue "
            f"{units[a.activity_key]!r}"
        )
        assert "fallback" not in str(q.get("formula", "")).lower(), (
            f"{a.activity_key}: quantity came from the legacy fallback"
        )

    # In an RCC residential building these two activities are genuinely out of
    # scope; every other row must carry work.
    zero_ok = {"roof_trusses", "curtain_wall"}
    for key in keys:
        if key in zero_ok:
            assert sched.quantities[key]["quantity"] == 0.0
        else:
            assert sched.quantities[key]["quantity"] > 0.0, f"{key} has no quantity"


def test_workbook_round_trip_preserves_building_inputs(tmp_path):
    """update-plan must retain blocks, quality and construction type."""
    pi = _input(number_of_blocks=3, quality_finish="premium",
                construction_type="steel", basement=True)
    loader = DataLoader(guidelines_path=GUIDELINES, assumptions_path=ASSUMPTIONS)
    loader.load_all(terrain=pi.terrain)
    res = ConstraintScheduler(pi, loader, enable_acceleration_cost=True).solve()

    out = tmp_path / "building_round_trip.xlsx"
    ExcelReportGenerator(res, str(out)).generate()

    rebuilt, _ = WorkbookReader(str(out)).read()
    assert rebuilt.project_type == "building"
    assert rebuilt.built_up_area_sqm == pytest.approx(2000.0)
    assert rebuilt.number_of_floors == 4
    assert rebuilt.number_of_blocks == 3
    assert rebuilt.construction_type == "steel"
    assert rebuilt.quality_finish == "premium"
    assert rebuilt.basement is True


def test_site_area_is_not_divided_by_floors_twice():
    assumptions = _assumptions()
    two = BuildingQuantityCalculator(_input(number_of_floors=2), assumptions)
    four = BuildingQuantityCalculator(_input(number_of_floors=4), assumptions)
    # Same BUA and FAR => the plot is the same size regardless of storeys.
    assert two.site_area_sqm == pytest.approx(four.site_area_sqm)
    assert four.site_area_sqm == pytest.approx(2000.0 / 1.5, abs=0.1)


def test_quantities_follow_geometry_and_totals(run):
    sched, _, _ = run
    q = {k: v["quantity"] for k, v in sched.quantities.items()}
    assumptions = _assumptions()
    p = assumptions["building_parameters"]

    assert q["site_setting_out"] == pytest.approx(1333.3, abs=0.2)
    assert q["site_clearance"] == pytest.approx(1333.3, abs=0.2)
    assert q["gf_slab"] == pytest.approx(500.0, abs=0.2)
    assert q["anti_termite"] == pytest.approx(500.0, abs=0.2)

    frame_keys = [
        "gf_columns", "ff_beams_slab", "gf_stairs", "ff_columns",
        "sf_beams_slab", "ff_stairs", "sf_columns", "roof_beams",
    ]
    frame_total = sum(q[k] for k in frame_keys)
    assert frame_total == pytest.approx(2000.0 * p["frame_rcc_per_sqm"], abs=0.5)

    masonry_total = q["gf_walls"] + q["ff_walls"] + q["sf_walls"]
    assert masonry_total == pytest.approx(2000.0 * p["masonry_per_sqm"], abs=0.5)

    assert q["piling"] > 0 and q["pile_testing"] >= 3
    pile_depth = q["piling"] / q["pile_head_chipping"]
    assert 8.0 <= pile_depth <= 18.0
    assert q["electrical"] == pytest.approx(1000.0, abs=0.1)
    assert q["plumbing"] == pytest.approx(2000.0, abs=0.1)
    assert q["doors"] == pytest.approx(80.0, abs=0.1)
    assert q["windows"] == pytest.approx(100.0, abs=0.1)


def test_planned_rates_and_wastage_use_shared_cost_method(run):
    _, res, _ = run
    variations = []
    for a in res.activities:
        c = a.cost
        assert c.planned_unit_rate > 0
        assert c.planning_variation_pct == pytest.approx(
            (c.planned_unit_rate - c.morth_direct_unit_rate)
            / c.morth_direct_unit_rate * 100,
            abs=0.0001,
        ), a.activity_key
        variations.append(c.planning_variation_pct)
        assert not c.wastage_in_base_rate, a.activity_key
        if a.quantity > 0:
            assert c.wastage_cost > 0, a.activity_key
        assert c.overhead_amount == 0 and c.profit_amount == 0
        if a.quantity > 0:
            assert c.overhead_embedded > 0, a.activity_key
            assert c.profit_embedded > 0, a.activity_key
            assert c.overhead_embedded + c.profit_embedded < a.quantity * c.planned_unit_rate
        assert c.total_amount == pytest.approx(
            c.direct_cost + c.overhead_amount + c.profit_amount + c.escalation_amount
        )
    assert len({round(value, 3) for value in variations}) > 1


def test_building_rate_uses_recipe_and_source_share_fallback():
    assumptions = _assumptions()
    loader = DataLoader(
        guidelines_path=GUIDELINES,
        assumptions_path=ASSUMPTIONS,
    )
    loader.load_all(terrain="plain")
    calculator = CostCalculator(loader.assumptions)

    recipe_key = "gf_columns"
    fallback_key = "site_clearance"
    recipe_rate, _ = calculator._resource_planned_rate(recipe_key, quantity=1.0)
    fallback_rate, _ = calculator._resource_planned_rate(fallback_key, quantity=1.0)
    recipe_spec = calculator._activity_recipe[recipe_key]
    recipe = calculator._resolve_recipe(recipe_spec["recipe"])
    recipe_material_cost = sum(
        float(amount) * float(recipe_spec.get("per_output", 1.0))
        / float(calculator._material_bindings[material].get("divisor", 1) or 1)
        * float(calculator._material_rates[
            calculator._material_bindings[material]["rate"]
        ])
        for material, amount in recipe.items()
        if not str(material).startswith("_")
    )
    source_rate = float(calculator.direct_rates[fallback_key])
    fallback_spec = next(
        item for item in loader.assumptions["activity_master"]
        if item["key"] == fallback_key
    )
    fallback_material = (
        source_rate * calculator._shares_for(fallback_spec["category"])["material"]
    )

    assert assumptions["rate_build_up_mode"] == "resource_build_up"
    assert recipe_rate > recipe_material_cost
    assert fallback_rate > fallback_material


def test_building_plan_shows_recipe_materials_per_crew_day(run):
    _, result, _ = run
    activity = next(
        row for row in build_plan(result)
        if row.activity_id == "ACT-14"
    )
    cement = next(
        item for item in activity.daily_material_assumptions
        if item["name"] == "cement_kg"
    )

    assert cement["quantity_per_crew_day"] == pytest.approx(2400.0)
    assert cement["unit"] == "kg"
    assert cement["output_unit"] == "cum"
    assert cement["verification_status"] == "configured_recipe"
    assert "excludes material wastage" in cement["cost_treatment"]
    assert activity.material_cost > 0


def test_building_run_cost_and_duration_are_plausible(run):
    _, res, _ = run
    # Resource-derived rates can exceed the old source-rate-only estimate.
    # The total must still reconcile exactly to the calculated activities.
    assert res.total_calculated_cost_crore > 0
    assert 250 <= res.calculated_duration_days <= 900
    assert res.total_calculated_cost_crore == pytest.approx(
        sum(a.cost.total_amount for a in res.activities) / 1e7, abs=0.01
    )
    assert any(a.cost.escalation_amount > 0 for a in res.activities)


def test_workbook_reports_embedded_oh_cp(run, tmp_path):
    _, res, _ = run
    out = tmp_path / "building.xlsx"
    ExcelReportGenerator(res, str(out)).generate()
    wb = openpyxl.load_workbook(out, data_only=True)

    ws = wb["03_Assumptions"]
    header_row = next(
        row for row in ws.iter_rows(values_only=True)
        if row and row[0] == "Activity ID"
    )
    headers = list(header_row)
    oh_col = headers.index("Overhead % (embedded)")
    cp_col = headers.index("Profit % (embedded)")
    direct_col = headers.index("Direct Unit Rate (₹)")
    planned_col = headers.index("Planned Unit Rate (₹)")
    seen = 0
    for row in ws.iter_rows(min_row=4, values_only=True):
        if row[0] and str(row[0]).startswith("ACT-"):
            assert float(row[planned_col]) == pytest.approx(
                float(row[direct_col])
                * (1 + float(row[headers.index("Planning Variation (%)")]) / 100),
                abs=0.01,
            )
            assert "%" in str(row[oh_col]) and row[oh_col] != "In source rate"
            assert "%" in str(row[cp_col]) and row[cp_col] != "In source rate"
            seen += 1
    assert seen == 40


def test_night_work_and_pricing_config_use_catalogue_keys():
    catalogue_keys = {a["key"] for a in _catalogue()}
    assumptions = _assumptions()
    assert assumptions["use_direct_rate_as_planned"] is False
    assert assumptions["rate_build_up_mode"] == "resource_build_up"
    assert assumptions["rate_oncost_basis"] == "statutory"
    assert set(assumptions["night_work_eligible"]) <= catalogue_keys
    assert not any(str(k).startswith("bld_") for k in assumptions["night_work_eligible"])
    assert "annual_rates" in assumptions["escalation"]
    assert "base_date" in assumptions["escalation"]


def test_four_blocks_accelerate_through_night_work_keys():
    pi = _input(number_of_blocks=4, duration_days=360, budget_crore=8.0)
    loader = DataLoader(guidelines_path=GUIDELINES, assumptions_path=ASSUMPTIONS)
    loader.load_all(terrain=pi.terrain)
    res = ConstraintScheduler(pi, loader, enable_acceleration_cost=True).solve()
    assert res.calculated_duration_days < res.baseline_duration_days
    assert res.optimized_duration_days < res.baseline_duration_days
    assert res.duration_saved_days > 0
    assert res.additional_cost_crore > 0


def test_escalation_reads_annual_rates_and_legacy_pct():
    cfg = _assumptions()
    cfg["escalation"]["enabled"] = True
    cfg["escalation"]["base_date"] = "2025-01-01"
    model = EscalationModel(cfg)
    assert model.enabled
    assert model.annual_rates["material"] == pytest.approx(0.04)
    assert model.annual_rates["labour"] == pytest.approx(0.05)
    factor = model.blended_factor(
        1.0, {"material": 0.55, "labour": 0.25, "plant": 0.12, "fuel": 0.08}
    )
    assert factor > 1.0

    legacy = EscalationModel({
        "escalation": {
            "enabled": True,
            "annual_rate_pct": 5.0,
            "base_date": "2025-01-01",
        }
    })
    assert legacy.annual_rates == {
        "material": 0.05, "labour": 0.05, "plant": 0.05, "fuel": 0.05,
    }
    assert legacy.default_rate == pytest.approx(0.05)
