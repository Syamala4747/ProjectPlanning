"""Calendar-aware planned execution rows for the Engineer_DailyLog sheet."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any, List

from .project_calendar import ProjectCalendar
from .plan_generator import PlanActivity, build_plan


@dataclass
class DailyLogRow:
    day_number: int
    activity_working_day: int
    date: date
    weekday: str
    activity_id: str
    activity_name: str
    phase: str
    quantity: float
    unit: str
    crew: str
    equipment: str
    materials: str
    labour_cost: float
    equipment_cost: float
    fuel_consumption: float
    fuel_cost: float
    material_cost: float
    other_cost: float
    total_cost: float
    cumulative_cost: float


def build_daily_log(result: Any, plan: List[PlanActivity] | None = None) -> List[DailyLogRow]:
    """Distribute each active activity by calendar-weighted working days."""
    assumptions = getattr(result, "assumptions_snapshot", {}) or {}
    project = result.project_input
    calendar = ProjectCalendar(
        assumptions.get("calendar", {}) or {},
        project.start_date,
        terrain=getattr(project, "terrain", "plain"),
    )
    plan = plan if plan is not None else build_plan(result)
    records = []

    for row in plan:
        if row.status != "Active" or not row.planned_start or not row.planned_finish:
            continue
        dates = []
        current = row.planned_start
        while current <= row.planned_finish:
            if calendar.is_working(current):
                dates.append(current)
            current += timedelta(days=1)
        if not dates:
            dates = [row.planned_start]
        weights = [max(0.0, calendar.productivity_factor(d)) for d in dates]
        total_weight = sum(weights) or float(len(dates))
        for activity_day, (d, weight) in enumerate(zip(dates, weights), start=1):
            share = weight / total_weight
            equipment = row.equipment[0]
            available_materials = [
                m for m in row.materials
                if m["quantity"] > 0 and m["unit"] not in ("not verifiable", "—", "rate apportionment")
            ]
            if available_materials:
                materials = "; ".join(
                    f"{m['name']}: {m['quantity'] * share:.3f} {m['unit']}"
                    for m in available_materials
                )
            else:
                # Use the descriptive label from plan_generator
                fallback = row.materials[0]["name"] if row.materials else "No material takeoff configured"
                materials = fallback
            records.append({
                "date": d,
                "activity_id": row.activity_id,
                "row": DailyLogRow(
                    day_number=0,
                    activity_working_day=activity_day,
                    date=d,
                    weekday=d.strftime("%A"),
                    activity_id=row.activity_id,
                    activity_name=row.activity_name,
                    phase=row.phase,
                    quantity=row.quantity * share,
                    unit=row.unit,
                    crew=row.labour[0]["category"],
                    equipment=equipment["name"],
                    materials=materials,
                    labour_cost=row.labour_cost * share,
                    equipment_cost=row.equipment_cost * share,
                    fuel_consumption=equipment["fuel_consumption"] * share,
                    fuel_cost=row.fuel_cost * share,
                    material_cost=row.material_cost * share,
                    other_cost=row.other_cost * share,
                    total_cost=(
                        row.labour_cost + row.equipment_cost + row.fuel_cost
                        + row.material_cost + row.other_cost
                    ) * share,
                    cumulative_cost=0.0,
                ),
            })

    records.sort(key=lambda item: (item["date"], item["activity_id"]))
    cumulative = 0.0
    project_day = {}
    working_day = 0
    last_date = None
    for item in records:
        if item["date"] != last_date:
            working_day += 1
            project_day[item["date"]] = working_day
            last_date = item["date"]
    output = []
    for item in records:
        row = item["row"]
        cumulative += row.total_cost
        row.day_number = project_day[item["date"]]
        row.cumulative_cost = cumulative
        output.append(row)
    return output
