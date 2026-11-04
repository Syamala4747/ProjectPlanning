"""Resource-plan rows derived from the solved schedule and cost engine output."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any, Dict, List


def _material_unit(material_key: str, binding: Dict[str, Any]) -> str | None:
    configured_unit = binding.get("unit")
    if configured_unit:
        return str(configured_unit)
    if material_key.endswith("_lump"):
        return None
    if material_key == "delineator_set_km":
        return "set"
    suffix_units = (
        ("_per_metre", "m"),
        ("_metre", "m"),
        ("_litre", "L"),
        ("_thousand", "per 1,000"),
        ("_cum", "m3"),
        ("_sqm", "m2"),
        ("_kg", "kg"),
        ("_no", "No."),
        ("_set", "set"),
        ("_per_tree", "per tree"),
        ("_per_cum", "per m3"),
    )
    return next(
        (unit for suffix, unit in suffix_units if material_key.endswith(suffix)),
        None,
    )


def _resolve_material_recipe(
    recipes: Dict[str, Any], recipe_name: str
) -> Dict[str, Any]:
    seen = set()
    current = recipe_name
    while True:
        if current in seen:
            raise ValueError(
                f"Material recipe '{recipe_name}' contains a circular _use reference."
            )
        seen.add(current)
        recipe = recipes.get(current)
        if recipe is None:
            raise ValueError(
                f"Activity references missing material recipe '{current}'."
            )
        target = recipe.get("_use") if isinstance(recipe, dict) else None
        if not target:
            return recipe
        current = str(target)


@dataclass
class PlanActivity:
    activity_id: str
    activity_name: str
    phase: str
    milestone: str
    planned_start: date | None
    planned_finish: date | None
    duration_days: int
    quantity: float
    unit: str
    productivity: float
    crew_multiplier: int
    status: str
    source_unit_rate: float
    planning_variation_pct: float
    planned_unit_rate: float
    labour_cost: float
    equipment_cost: float
    fuel_cost: float
    material_cost: float
    other_cost: float
    total_execution_cost: float
    fuel_consumption: float
    fuel_rate: float
    labour: List[Dict[str, Any]] = field(default_factory=list)
    equipment: List[Dict[str, Any]] = field(default_factory=list)
    materials: List[Dict[str, Any]] = field(default_factory=list)
    daily_material_assumptions: List[Dict[str, Any]] = field(default_factory=list)
    unitemized_rate_material_cost: float = 0.0


def build_plan(result: Any) -> List[PlanActivity]:
    """Build report rows without re-pricing or re-quantifying activities."""
    assumptions = getattr(result, "assumptions_snapshot", {}) or {}
    calendar = assumptions.get("calendar", {}) or {}
    hours_per_day = float(calendar.get("equipment_hours_per_day", 8.0) or 8.0)
    fuel_rates = assumptions.get("material_rates_inr", {}) or {}
    fuel_rate = float(
        assumptions.get("diesel_rate_inr_per_litre")
        or fuel_rates.get("diesel_litre", 0.0)
        or 0.0
    )
    rows: List[PlanActivity] = []
    crew_composition = assumptions.get("crew_composition", {}) or {}
    profile_by_activity = assumptions.get("crew_profile_by_activity_key", {}) or {}
    equipment_by_profile = assumptions.get("equipment_per_activity", {}) or {}
    equipment_rates = assumptions.get("equipment_hire_inr_per_day", {}) or {}
    equipment_fuel_rates = assumptions.get("fuel_consumption_lph", {}) or {}
    labour_rates = assumptions.get("labour_rates_inr_per_day", {}) or {}
    material_bindings = assumptions.get("material_bindings", {}) or {}
    daily_material_config = assumptions.get("daily_material_assumptions", {}) or {}
    daily_material_by_activity = daily_material_config.get("activities", {}) or {}
    recipes = {}
    recipes.update(assumptions.get("material_per_unit_pipeline", {}) or {})
    recipes.update(assumptions.get("material_per_unit", {}) or {})
    recipes.update(assumptions.get("material_per_unit_road", {}) or {})
    activity_recipe_by_key = assumptions.get("activity_material_recipe", {}) or {}

    for activity in list(getattr(result, "activities", []) or []):
        cost = activity.cost
        quantity = float(activity.quantity or 0.0)
        duration = max(0, int(activity.duration_days or 0))
        crews = max(1, int(activity.crews or activity.base_crews or 1))
        active = quantity > 0 and duration > 0
        labour_cost = float(cost.direct_labour or 0.0)
        equipment_cost = float(cost.direct_equipment or 0.0)
        fuel_cost = float(cost.direct_fuel or 0.0)
        material_cost = float(cost.direct_material or cost.bottom_up_material_cost or 0.0)
        material_lines = list(getattr(cost, "material_lines", []) or [])
        unitemized_rate_material_cost = 0.0
        if (
            assumptions.get("rate_build_up_mode")
            in ("source_plus_oncosts", "resource_build_up")
            and not material_lines
            and material_cost > 0
        ):
            unitemized_rate_material_cost = material_cost
            material_cost = 0.0
        component_total = labour_cost + equipment_cost + fuel_cost + material_cost
        other_cost = max(0.0, float(cost.direct_cost or 0.0) - component_total)

        profile = profile_by_activity.get(activity.activity_key, "")
        configured_labour = crew_composition.get(profile, []) or []
        raw_labour = [
            float(item.get("count", 0) or 0) * crews * duration
            * float(labour_rates.get(item.get("trade"), 0.0) or 0.0)
            for item in configured_labour
        ]
        raw_labour_total = sum(raw_labour)
        labour = []
        for item, raw in zip(configured_labour, raw_labour):
            share = raw / raw_labour_total if raw_labour_total else 0.0
            trade = str(item.get("trade", "configured labour"))
            count = int(item.get("count", 0) or 0) * crews if active else 0
            labour.append({
                "category": trade,
                "workers": count,
                "working_days": duration if active else 0,
                "daily_rate": float(labour_rates.get(trade, 0.0) or 0.0),
                "total_cost": labour_cost * share,
            })
        if not labour:
            labour = [{
                "category": "Unavailable in project assumptions",
                "workers": 0,
                "working_days": 0,
                "daily_rate": 0.0,
                "total_cost": 0.0,
            }]

        configured_equipment = equipment_by_profile.get(profile, []) or []
        raw_equipment = []
        for name in configured_equipment:
            rate = float(equipment_rates.get(name, 0.0) or 0.0)
            units = crews if active else 0
            raw_equipment.append(rate * duration * units)
        raw_equipment_total = sum(raw_equipment)
        equipment = []
        fuel_quantity = 0.0
        raw_fuel_quantities = [
            duration * hours_per_day * (crews if active else 0)
            * float(equipment_fuel_rates.get(name, 0.0) or 0.0)
            for name in configured_equipment
        ]
        raw_fuel_total = sum(raw_fuel_quantities)
        for name, raw in zip(configured_equipment, raw_equipment):
            rate = float(equipment_rates.get(name, 0.0) or 0.0)
            lph = float(equipment_fuel_rates.get(name, 0.0) or 0.0)
            units = crews if active else 0
            total_hours = duration * hours_per_day * units
            fuel_qty = total_hours * lph
            share = raw / raw_equipment_total if raw_equipment_total else 0.0
            line_equipment_cost = equipment_cost * share
            fuel_share = fuel_qty / raw_fuel_total if raw_fuel_total else 0.0
            line_fuel_cost = fuel_cost * fuel_share
            equipment.append({
                "name": name,
                "units": units,
                "operating_hours": hours_per_day if active else 0.0,
                "total_hours": total_hours,
                "fuel_rate_lph": lph,
                "fuel_consumption": fuel_qty,
                "fuel_cost": line_fuel_cost,
                "hire_rate": rate,
                "total_cost": line_equipment_cost,
            })
            fuel_quantity += fuel_qty
        if fuel_cost > 0 and not any(
            float(item.get("fuel_cost", 0.0) or 0.0) > 0 for item in equipment
        ):
            # The cost engine may carry a fuel allowance even when the
            # configured plant has no fuel-consumption rate.
            fuel_quantity = fuel_cost / fuel_rate if fuel_rate else fuel_quantity
            equipment.append({
                "name": "Fuel allowance (original calculation)",
                "units": 0,
                "operating_hours": 0.0,
                "total_hours": 0.0,
                "fuel_rate_lph": 0.0,
                "fuel_consumption": fuel_quantity,
                "fuel_cost": fuel_cost,
                "hire_rate": 0.0,
                "total_cost": equipment_cost if not equipment else 0.0,
            })
        if not equipment:
            if equipment_cost > 0 or fuel_cost > 0:
                # Some catalogues calculate plant/fuel allowances without
                # retaining an equipment profile. Keep both original amounts
                # auditable instead of reporting false zero detail.
                fuel_quantity = fuel_cost / fuel_rate if fuel_rate else 0.0
                equipment = [{
                    "name": "Plant/fuel allowance (original calculation)",
                    "units": 0,
                    "operating_hours": 0.0,
                    "total_hours": 0.0,
                    "fuel_rate_lph": 0.0,
                    "fuel_consumption": fuel_quantity,
                    "fuel_cost": fuel_cost,
                    "hire_rate": 0.0,
                    "total_cost": equipment_cost,
                }]
            else:
                equipment = [{
                    "name": "Unavailable in project assumptions",
                    "units": 0,
                    "operating_hours": 0.0,
                    "total_hours": 0.0,
                    "fuel_rate_lph": 0.0,
                    "fuel_consumption": 0.0,
                    "fuel_cost": 0.0,
                    "hire_rate": 0.0,
                    "total_cost": 0.0,
                }]
        if fuel_quantity and fuel_cost:
            fuel_quantity = fuel_cost / fuel_rate if fuel_rate else fuel_quantity

        materials = []
        for line in material_lines:
            materials.append({
                "name": line.get("material", "Configured material"),
                "quantity": float(line.get("issued_qty", 0.0) or 0.0),
                "unit": line.get("unit") or material_bindings.get(
                    line.get("material"), {}
                ).get("unit", "configured unit"),
                "unit_rate": float(line.get("unit_rate", 0.0) or 0.0),
                "total_cost": float(line.get("cost", 0.0) or 0.0),
            })
        if not materials and material_cost:
            materials.append({
                "name": "Material detail unavailable in assumptions",
                "quantity": 0.0,
                "unit": "not verifiable",
                "unit_rate": 0.0,
                "total_cost": material_cost,
            })

        daily_material_assumptions = []
        daily_material_specs = daily_material_by_activity.get(
            activity.activity_key, []
        ) or []
        if daily_material_specs:
            daily_output = float(activity.productivity_rate or 0.0)
            if daily_output <= 0:
                raise ValueError(
                    f"daily_material_assumptions for '{activity.activity_key}' "
                    "requires a positive activity productivity_rate."
                )
            for item in daily_material_specs:
                per_output_unit = float(item["quantity_per_output_unit"])
                if per_output_unit < 0:
                    raise ValueError(
                        f"daily_material_assumptions for '{activity.activity_key}' "
                        "cannot contain a negative quantity_per_output_unit."
                    )
                daily_material_assumptions.append({
                    "name": item["material"],
                    "quantity_per_crew_day": per_output_unit * daily_output,
                    "unit": item["material_unit"],
                    "output_unit": item["output_unit"],
                    "cost_treatment": daily_material_config.get(
                        "cost_treatment", "Memo quantity only."
                    ),
                    "verification_status": item.get(
                        "verification_status", "provisional"
                    ),
                })
        else:
            activity_recipe = activity_recipe_by_key.get(activity.activity_key)
            if activity_recipe:
                recipe_name = str(activity_recipe.get("recipe") or "")
                if not recipe_name:
                    raise ValueError(
                        f"activity_material_recipe for '{activity.activity_key}' "
                        "does not specify a recipe."
                    )
                recipe = _resolve_material_recipe(recipes, recipe_name)
                per_output = float(activity_recipe.get("per_output", 1.0) or 1.0)
                daily_output = float(activity.productivity_rate or 0.0)
                if daily_output <= 0:
                    raise ValueError(
                        f"Activity '{activity.activity_key}' has a material recipe "
                        "but no positive productivity_rate for its daily takeoff."
                    )
                for material_key, recipe_quantity in recipe.items():
                    if str(material_key).startswith("_"):
                        continue
                    binding = material_bindings.get(material_key, {}) or {}
                    unit = _material_unit(str(material_key), binding)
                    if unit is None:
                        continue
                    daily_material_assumptions.append({
                        "name": str(material_key),
                        "quantity_per_crew_day": (
                            float(recipe_quantity) * per_output * daily_output
                        ),
                        "unit": unit,
                        "output_unit": activity.unit,
                        "cost_treatment": (
                            "Informational only; this recipe is already priced in "
                            "the planned rate. Daily quantity excludes material wastage."
                        ),
                        "verification_status": "configured_recipe",
                    })

        rows.append(PlanActivity(
            activity_id=activity.activity_id,
            activity_name=activity.name,
            phase=activity.work_package or activity.category,
            milestone="Critical activity" if activity.is_critical else "",
            planned_start=activity.planned_start,
            planned_finish=activity.planned_finish,
            duration_days=duration,
            quantity=quantity,
            unit=activity.unit,
            productivity=float(activity.productivity_rate or 0.0),
            crew_multiplier=crews,
            status="Active" if active else "Out of scope",
            source_unit_rate=float(cost.morth_direct_unit_rate or 0.0),
            planning_variation_pct=float(cost.planning_variation_pct or 0.0),
            planned_unit_rate=float(cost.planned_unit_rate or 0.0),
            labour_cost=labour_cost,
            equipment_cost=equipment_cost,
            fuel_cost=fuel_cost,
            material_cost=material_cost,
            other_cost=other_cost,
            total_execution_cost=float(cost.direct_cost or 0.0),
            fuel_consumption=fuel_quantity,
            fuel_rate=fuel_rate,
            labour=labour,
            equipment=equipment,
            materials=materials,
            daily_material_assumptions=daily_material_assumptions,
            unitemized_rate_material_cost=unitemized_rate_material_cost,
        ))
    return rows
