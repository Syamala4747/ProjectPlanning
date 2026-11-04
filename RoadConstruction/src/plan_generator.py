"""Resource-plan rows derived from the solved schedule and cost engine output."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any, Dict, List


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
        for line in list(getattr(cost, "material_lines", []) or []):
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
        ))
    return rows
