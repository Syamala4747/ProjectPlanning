"""Independent validation and sanity checks for generated project plans.

The checks are deliberately result-oriented: they validate the schedule and
cost objects after calculation and do not replace the calculation engines.
Errors indicate an internally invalid plan; warnings identify assumptions or
engineering decisions that require review.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
import math
from typing import Any, Iterable, List, Optional


@dataclass
class ValidationIssue:
    severity: str
    code: str
    message: str
    activity_id: Optional[str] = None


@dataclass
class ValidationReport:
    issues: List[ValidationIssue] = field(default_factory=list)

    @property
    def errors(self) -> List[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == "ERROR"]

    @property
    def warnings(self) -> List[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == "WARNING"]

    @property
    def passed(self) -> bool:
        return not self.errors

    def add(
        self,
        severity: str,
        code: str,
        message: str,
        activity_id: Optional[str] = None,
    ) -> None:
        self.issues.append(ValidationIssue(severity, code, message, activity_id))

    @property
    def status(self) -> str:
        return "PASS" if self.passed else "FAIL"


def _near(left: float, right: float, absolute: float = 0.05) -> bool:
    return abs(float(left) - float(right)) <= absolute


def validate_schedule_result(
    result: Any,
    assumptions: Optional[dict] = None,
    *,
    tolerance_inr: float = 0.05,
    calendar: Optional[Any] = None,
) -> ValidationReport:
    """Validate a ``ProjectScheduleResult`` and return errors/warnings.

    The function is safe for road, pipeline, and building results. All project
    types use the same direct-cost identity: quantity × planned rate +
    wastage + acceleration surcharge.
    """
    report = ValidationReport()
    activities = list(getattr(result, "activities", []) or [])
    by_id = {a.activity_id: a for a in activities}

    if not activities:
        report.add("ERROR", "NO_ACTIVITIES", "The plan contains no activities.")
        return report

    project_type = str(getattr(result.project_input, "project_type", "road") or "road").lower()
    budget_crore = float(getattr(result.project_input, "budget_crore", 0.0) or 0.0)
    overhead_pct = _validation_overhead_pct(assumptions or {}, budget_crore)
    if project_type == "pipeline" and assumptions:
        _validate_pipeline_inputs(result, assumptions, report, tolerance_inr)
        _validate_engineering_sequence(result, report)
    _validate_activity_catalogue(result, assumptions, report, project_type)
    _validate_resource_allocation(result, assumptions, report)
    _validate_escalation(result, assumptions, report)
    _validate_cycles(result, report)
    _validate_plan_calculations(result, report, tolerance_inr)
    _validate_resource_plan_and_daily_log(result, report)

    # Quantity, productivity, duration, and CPM invariants.
    for activity in activities:
        aid = activity.activity_id
        quantity = float(activity.quantity or 0.0)
        productivity = float(activity.productivity_rate or 0.0)
        duration = int(activity.duration_days or 0)
        cost = activity.cost

        zero_quantity = quantity <= 0
        if quantity < 0:
            report.add("ERROR", "NEGATIVE_QUANTITY", "Quantity cannot be negative.", aid)
        if productivity <= 0:
            report.add("ERROR", "INVALID_PRODUCTIVITY", "Productivity must be positive.", aid)
        if duration < 1 and not zero_quantity:
            report.add("ERROR", "INVALID_DURATION", "Active activities must last at least one day.", aid)
        effective_shifts = float(activity.shifts or 1)
        if calendar is not None:
            effective_shifts = calendar.effective_shifts(effective_shifts)
        elif effective_shifts > 1:
            # Keep validation independent of the scheduler while matching the
            # documented night-shift productivity rule.
            night_factor = float(
                ((assumptions or {}).get("calendar") or {})
                .get("night_shift_productivity_factor", 0.8)
                or 0.8
            )
            effective_shifts = 1.0 + (effective_shifts - 1.0) * night_factor
        expected_effort = 0 if zero_quantity else max(
            1,
            math.ceil(quantity / max(productivity * float(activity.crews or 1) *
                                    effective_shifts, 1e-9)),
        )
        if int(getattr(activity, "effort_days", expected_effort) or 0) != expected_effort:
            report.add(
                "ERROR",
                "PRODUCTIVITY_DURATION_MISMATCH",
                f"Effort days {activity.effort_days} does not match independent "
                f"ceil(quantity / productivity / crews / shifts) = {expected_effort}.",
                aid,
            )
        if activity.ef != activity.es + duration:
            report.add("ERROR", "CPM_EF_IDENTITY", "EF must equal ES plus duration.", aid)
        if activity.lf != activity.ls + duration:
            report.add("ERROR", "CPM_LF_IDENTITY", "LF must equal LS plus duration.", aid)
        if activity.total_float != activity.ls - activity.es:
            report.add("ERROR", "CPM_FLOAT_IDENTITY", "Total float must equal LS minus ES.", aid)
        if activity.total_float < 0 or activity.free_float < 0:
            report.add("ERROR", "NEGATIVE_FLOAT", "CPM float cannot be negative.", aid)

        for predecessor in activity.predecessors:
            if predecessor not in by_id:
                report.add(
                    "ERROR",
                    "MISSING_PREDECESSOR",
                    f"Predecessor {predecessor} is not present in the active schedule.",
                    aid,
                )
            else:
                lag = int((activity.predecessor_lags or {}).get(predecessor, 0))
                if activity.es < by_id[predecessor].ef + lag:
                    report.add(
                        "ERROR",
                        "PRECEDENCE_VIOLATION",
                        f"ES {activity.es} is before predecessor EF plus lag "
                        f"{by_id[predecessor].ef + lag}.",
                        aid,
                    )

        if float(cost.planned_unit_rate or 0.0) <= 0:
            report.add("ERROR", "INVALID_PLANNED_RATE", "Planned unit rate must be positive.", aid)
        if float(cost.direct_cost or 0.0) < 0 or float(cost.total_amount or 0.0) < 0:
            report.add("ERROR", "NEGATIVE_COST", "Cost values cannot be negative.", aid)

        expected_direct = quantity * float(cost.planned_unit_rate or 0.0)
        if not bool(getattr(cost, "wastage_in_base_rate", False)):
            expected_direct += float(cost.wastage_cost or 0.0)
        expected_direct += float(cost.crash_total_surcharge or 0.0)
        if not _near(cost.direct_cost, expected_direct, tolerance_inr):
            report.add(
                "ERROR",
                "DIRECT_COST_RECONCILIATION",
                f"Direct cost {cost.direct_cost:.2f} does not equal "
                f"quantity × planned rate + surcharge {expected_direct:.2f}.",
                aid,
            )

        expected_total = (
            float(cost.direct_cost or 0.0)
            + float(cost.overhead_amount or 0.0)
            + float(cost.profit_amount or 0.0)
            + float(cost.escalation_amount or 0.0)
        )
        if not _near(cost.total_amount, expected_total, tolerance_inr):
            report.add(
                "ERROR",
                "TOTAL_COST_RECONCILIATION",
                f"Total cost {cost.total_amount:.2f} does not equal "
                f"direct + overhead + profit + escalation {expected_total:.2f}.",
                aid,
            )

        crash_parts = (
            "crash_fleet_mobilisation",
            "crash_shift_premium",
            "crash_front_supervision",
            "crash_site_congestion",
            "crash_site_establishment",
            "crash_resident_staff",
            "crash_plant_hire_uplift",
            "crash_efficiency_loss",
        )
        expected_surcharge = sum(float(getattr(cost, name, 0.0) or 0.0) for name in crash_parts)
        if not _near(cost.crash_total_surcharge, expected_surcharge, tolerance_inr):
            report.add(
                "ERROR",
                "SURCHARGE_RECONCILIATION",
                f"Acceleration surcharge {cost.crash_total_surcharge:.2f} does not "
                f"equal its component sum {expected_surcharge:.2f}.",
                aid,
            )

        wastage_fields = (
            "material_wastage_cost",
            "fuel_wastage_cost",
            "plant_wastage_cost",
            "labour_wastage_cost",
        )
        wastage_sum = sum(float(getattr(cost, name, 0.0) or 0.0) for name in wastage_fields)
        if float(cost.wastage_cost or 0.0) < 0 or wastage_sum < 0:
            report.add("ERROR", "NEGATIVE_WASTAGE", "Wastage values cannot be negative.", aid)
        if not _near(cost.wastage_cost, wastage_sum, tolerance_inr):
            report.add(
                "ERROR",
                "WASTAGE_RECONCILIATION",
                f"Reported wastage {cost.wastage_cost:.2f} does not equal "
                f"its component sum {wastage_sum:.2f}.",
                aid,
            )
        if assumptions and assumptions.get("rate_build_up_mode") == "source_plus_oncosts":
            _validate_source_plus_oncost_rate(
                activity, assumptions, report, tolerance_inr, overhead_pct
            )

    # Project-level checks.
    max_finish = max(a.ef for a in activities)
    if max_finish != int(getattr(result, "calculated_duration_days", max_finish)):
        report.add(
            "ERROR",
            "PROJECT_DURATION_RECONCILIATION",
            f"Project duration {result.calculated_duration_days} does not equal "
            f"the maximum activity finish {max_finish}.",
        )

    expected_cost_crore = sum(float(a.cost.total_amount or 0.0) for a in activities) / 10_000_000.0
    if not _near(result.total_calculated_cost_crore, expected_cost_crore, 0.01):
        report.add(
            "ERROR",
            "PROJECT_COST_RECONCILIATION",
            f"Project total {result.total_calculated_cost_crore:.4f} Cr does not equal "
            f"activity total {expected_cost_crore:.4f} Cr.",
        )

    target_duration = int(getattr(result.project_input, "duration_days", 0) or 0)
    budget = float(getattr(result.project_input, "budget_crore", 0.0) or 0.0)
    if target_duration > 0 and result.calculated_duration_days > target_duration:
        report.add(
            "WARNING",
            "DURATION_TARGET_EXCEEDED",
            f"Calculated duration {result.calculated_duration_days} exceeds target "
            f"{target_duration} days.",
        )
    if budget > 0 and result.total_calculated_cost_crore > budget:
        report.add(
            "WARNING",
            "BUDGET_EXCEEDED",
            f"Calculated cost {result.total_calculated_cost_crore:.2f} Cr exceeds "
            f"budget {budget:.2f} Cr.",
        )

    # Monthly reconciliation checks.
    monthly = list(getattr(result, "monthly_progress", []) or [])
    if not monthly:
        report.add("ERROR", "NO_MONTHLY_PROGRESS", "No monthly progress was generated.")
    else:
        expenditure = sum(float(item.period_expenditure_crore or 0.0) for item in monthly)
        if not _near(expenditure, result.total_calculated_cost_crore, 0.05):
            report.add(
                "ERROR",
                "MONTHLY_COST_RECONCILIATION",
                f"Monthly expenditure {expenditure:.4f} Cr does not equal "
                f"project total {result.total_calculated_cost_crore:.4f} Cr.",
            )
        final_progress = float(monthly[-1].cumulative_planned_progress_pct or 0.0)
        if abs(final_progress - 100.0) > 0.2:
            report.add(
                "ERROR",
                "PROGRESS_NOT_COMPLETE",
                f"Final cumulative planned progress is {final_progress:.2f}%, not 100%.",
            )
        for activity in activities:
            allocated = sum(
                (float(getattr(item, "period_planned_progress_pct", 0.0) or 0.0)
                 for item in monthly)
            )
            if allocated < 0:
                report.add("ERROR", "NEGATIVE_PROGRESS", "Progress cannot be negative.", activity.activity_id)

    _validate_calendar(result, calendar, report)
    _validate_optimization(result, report)

    # Explicitly surface provisional pipeline inputs rather than silently
    # presenting them as approved design or contract facts.
    if assumptions:
        sources = assumptions.get("sources") or {}
        escalation = assumptions.get("escalation") or {}
        if "provisional" in str(escalation).lower():
            report.add(
                "WARNING",
                "PROVISIONAL_ESCALATION",
                "Escalation source/clause is provisional and requires contract verification.",
            )
        if "confirm" in str(sources.get("pipe_rates_note", "")).lower():
            report.add(
                "WARNING",
                "UNVERIFIED_PIPE_RATE",
                "Pipeline pipe supply/rate inclusions require BOQ or source-document confirmation.",
            )

    return report


def _validate_resource_plan_and_daily_log(result: Any, report: ValidationReport) -> None:
    """Validate the report-facing Plan and Engineer_DailyLog data."""
    from .engineer_daily_log import build_daily_log
    from .plan_generator import build_plan

    plan = build_plan(result)
    activities = {
        activity.activity_id: activity
        for activity in list(getattr(result, "activities", []) or [])
    }
    if {row.activity_id for row in plan} != set(activities):
        report.add("ERROR", "PLAN_ACTIVITY_COVERAGE",
                   "Plan must contain exactly one activity header for every scheduled activity.")

    log = build_daily_log(result, plan)
    seen = set()
    by_activity = {}
    activity_days = {}
    project_days = {}
    for row in log:
        key = (row.activity_id, row.date)
        if key in seen:
            report.add("ERROR", "DAILY_LOG_DUPLICATE",
                       f"Duplicate activity/date record: {row.activity_id} {row.date}.",
                       row.activity_id)
        seen.add(key)
        if row.date not in project_days:
            project_days[row.date] = row.day_number
        if project_days[row.date] != row.day_number:
            report.add("ERROR", "DAILY_LOG_PROJECT_DAY",
                       "All activities on a date must share the same project day.",
                       row.activity_id)
        activity_days.setdefault(row.activity_id, []).append(
            (row.activity_working_day, row.date)
        )
        if row.activity_id not in activities:
            report.add("ERROR", "DAILY_LOG_UNKNOWN_ACTIVITY",
                       f"Daily log references unknown activity {row.activity_id}.",
                       row.activity_id)
        if row.quantity < -0.05 or row.total_cost < -0.05:
            report.add("ERROR", "DAILY_LOG_NEGATIVE_VALUE",
                       "Daily quantity and cost must not be negative.", row.activity_id)
        by_activity.setdefault(row.activity_id, {"quantity": 0.0, "cost": 0.0})
        by_activity[row.activity_id]["quantity"] += row.quantity
        by_activity[row.activity_id]["cost"] += row.total_cost

    for activity_id, entries in activity_days.items():
        ordered = sorted(entries, key=lambda item: item[1])
        expected_days = list(range(1, len(ordered) + 1))
        if [item[0] for item in ordered] != expected_days:
            report.add("ERROR", "DAILY_LOG_ACTIVITY_DAY",
                       "Activity working-day numbers must start at 1 and be sequential.",
                       activity_id)

    for row in plan:
        actual = by_activity.get(row.activity_id, {"quantity": 0.0, "cost": 0.0})
        if row.status == "Active":
            if not _near(actual["quantity"], row.quantity, 0.1):
                report.add("ERROR", "DAILY_LOG_QUANTITY_RECONCILIATION",
                           f"Daily quantity {actual['quantity']:.3f} does not match "
                           f"Plan quantity {row.quantity:.3f}.", row.activity_id)
            if not _near(actual["cost"], row.total_execution_cost, 1.0):
                report.add("ERROR", "DAILY_LOG_COST_RECONCILIATION",
                           f"Daily cost {actual['cost']:.2f} does not match "
                           f"Plan execution cost {row.total_execution_cost:.2f}.",
                           row.activity_id)
        elif actual["quantity"] > 0.1:
            report.add("ERROR", "ZERO_SCOPE_DAILY_CONSUMPTION",
                       "Out-of-scope activities must not generate daily consumption.",
                       row.activity_id)

    previous = 0.0
    for row in sorted(log, key=lambda item: (item.date, item.activity_id)):
        if row.cumulative_cost + 1.0 < previous:
            report.add("ERROR", "DAILY_LOG_CUMULATIVE_COST",
                       "Cumulative cost must not decrease.", row.activity_id)
        previous = max(previous, row.cumulative_cost)
        activity = activities.get(row.activity_id)
        if activity and activity.planned_start and activity.planned_finish:
            if not (activity.planned_start <= row.date <= activity.planned_finish):
                report.add("ERROR", "DAILY_LOG_OUTSIDE_ACTIVITY_DATES",
                           "Daily record falls outside its activity dates.",
                           row.activity_id)


def _validate_plan_calculations(
    result: Any,
    report: ValidationReport,
    tolerance_inr: float,
) -> None:
    """Reconcile report-facing Plan values with the original cost calculation."""
    from .plan_generator import build_plan

    source = {
        activity.activity_id: activity
        for activity in list(getattr(result, "activities", []) or [])
    }
    for row in build_plan(result):
        activity = source.get(row.activity_id)
        if activity is None:
            report.add("ERROR", "PLAN_UNKNOWN_ACTIVITY",
                       "Plan contains an activity absent from the original calculation.",
                       row.activity_id)
            continue

        cost = activity.cost
        checks = (
            ("quantity", row.quantity, float(activity.quantity or 0.0), 0.1),
            ("duration", row.duration_days, int(activity.duration_days or 0), 0),
            ("crew", row.crew_multiplier, max(1, int(activity.crews or activity.base_crews or 1)), 0),
            ("labour cost", row.labour_cost, float(cost.direct_labour or 0.0), tolerance_inr),
            ("equipment cost", row.equipment_cost, float(cost.direct_equipment or 0.0), tolerance_inr),
            ("fuel cost", row.fuel_cost, float(cost.direct_fuel or 0.0), tolerance_inr),
            (
                "material cost",
                row.material_cost,
                float(cost.direct_material or cost.bottom_up_material_cost or 0.0),
                tolerance_inr,
            ),
            ("execution cost", row.total_execution_cost, float(cost.direct_cost or 0.0), tolerance_inr),
        )
        for label, actual, expected, tolerance in checks:
            if not _near(actual, expected, max(tolerance, 0.05)):
                report.add(
                    "ERROR",
                    "PLAN_ORIGINAL_CALCULATION_MISMATCH",
                    f"Plan {label} {float(actual):.2f} does not match "
                    f"original calculation {float(expected):.2f}.",
                    row.activity_id,
                )

        component_total = (
            row.labour_cost + row.equipment_cost
            + row.fuel_cost + row.material_cost + row.other_cost
        )
        if not _near(component_total, row.total_execution_cost, tolerance_inr):
            report.add(
                "ERROR",
                "PLAN_COMPONENT_TOTAL_MISMATCH",
                f"Plan resource components total {component_total:.2f}, but "
                f"execution cost is {row.total_execution_cost:.2f}.",
                row.activity_id,
            )

        detail_labour = sum(float(item.get("total_cost", 0.0) or 0.0) for item in row.labour)
        detail_equipment = sum(float(item.get("total_cost", 0.0) or 0.0) for item in row.equipment)
        detail_fuel = sum(float(item.get("fuel_cost", 0.0) or 0.0) for item in row.equipment)
        detail_material = sum(float(item.get("total_cost", 0.0) or 0.0) for item in row.materials)
        detail_checks = [("labour", detail_labour, row.labour_cost)]
        if any(item.get("name") != "Unavailable in project assumptions" for item in row.equipment):
            detail_checks.extend((
                ("equipment", detail_equipment, row.equipment_cost),
                ("fuel", detail_fuel, row.fuel_cost),
            ))
        material_lines_total = sum(
            float(line.get("cost", 0.0) or 0.0)
            for line in list(getattr(cost, "material_lines", []) or [])
        )
        if row.materials and material_lines_total:
            # Material detail is issued quantity including wastage; the
            # original direct-material component excludes that separate loss.
            detail_checks.append(("material", detail_material, material_lines_total))
        for label, detail, summary in detail_checks:
            if not _near(detail, summary, tolerance_inr):
                report.add(
                    "ERROR",
                    "PLAN_DETAIL_TOTAL_MISMATCH",
                    f"Plan {label} detail total {detail:.2f} does not match "
                    f"the activity summary {summary:.2f}.",
                    row.activity_id,
                )


def _validate_pipeline_inputs(result: Any, assumptions: dict, report: ValidationReport, tol: float) -> None:
    """Recheck pipeline geometry, strata partition, and selected pipe branch."""
    inp = result.project_input
    params = assumptions.get("pipeline_parameters") or {}
    length_m = float(inp.length_km or 0.0) * 1000.0
    width = float(getattr(inp, "trench_width_m", 0.0) or 0.0)
    if width <= 0:
        width = round(float(inp.pipe_diameter_mm or 300) / 1000.0 + 0.60, 2)
    depth = float(inp.trench_depth_m or 0.0)
    shares = params.get("excavation_strata_shares") or {}
    share_total = sum(float(shares.get(key, 0.0) or 0.0) for key in
                      ("ordinary", "moorum", "soft_rock", "hard_rock"))
    if abs(share_total - 1.0) > 1e-9:
        report.add("ERROR", "SOIL_SHARES_NOT_100", f"Soil strata shares total {share_total:.6f}, not 1.0.")

    expected = {
        "excavation_ordinary_0_1_5": length_m * width * min(depth, 1.5) * float(shares.get("ordinary", 0.0) or 0.0),
        "excavation_ordinary_1_5_3": length_m * width * max(depth - 1.5, 0.0) * float(shares.get("ordinary", 0.0) or 0.0),
        "excavation_moorum": length_m * width * depth * float(shares.get("moorum", 0.0) or 0.0),
        "excavation_soft_rock": length_m * width * depth * float(shares.get("soft_rock", 0.0) or 0.0),
        "excavation_hard_rock": length_m * width * depth * float(shares.get("hard_rock", 0.0) or 0.0),
    }
    actual = {a.activity_key: float(a.quantity or 0.0) for a in result.activities}
    for key, value in expected.items():
        if key in actual and abs(actual[key] - round(value, 2)) > max(tol, abs(value) * 1e-5):
            report.add("ERROR", "QUANTITY_FORMULA_MISMATCH",
                       f"{key} is {actual[key]:.2f}; independent formula gives {value:.2f}.")
    total_exc = sum(actual.get(key, 0.0) for key in expected)
    if abs(total_exc - length_m * width * depth) > max(tol, length_m * width * depth * 1e-5):
        report.add("ERROR", "EXCAVATION_PARTITION_MISMATCH",
                   "Soil activity quantities do not partition the total trench volume.")


def _validate_source_plus_oncost_rate(
    activity: Any,
    assumptions: dict,
    report: ValidationReport,
    tolerance_inr: float,
    overhead_pct: float,
) -> None:
    rates = assumptions.get("direct_unit_rates_inr") or {}
    source = rates.get(activity.activity_key)
    if isinstance(source, dict):
        source = source.get("value")
    if source is None:
        report.add("ERROR", "RATE_SOURCE_MISSING", "Source rate is missing.", activity.activity_id)
        return
    category = next(
        (item.get("category") for item in assumptions.get("activity_master", []) or []
         if item.get("key") == activity.activity_key),
        None,
    )
    source = float(source)
    profit = float((assumptions.get("statutory_and_overheads") or {}).get("profit_pct", 10.0))
    # Wastage is a separate direct-cost component. The planned rate contains
    # only the source rate plus the configured OH/CP on-costs.
    expected = source * (1.0 + overhead_pct / 100.0) * (1.0 + profit / 100.0)
    if abs(float(activity.cost.planned_unit_rate) - round(expected, 4)) > max(tolerance_inr, abs(expected) * 1e-5):
        report.add(
            "ERROR",
            "UNIT_RATE_BUILDUP_MISMATCH",
            f"Planned rate {activity.cost.planned_unit_rate:.4f} does not match "
            f"independent source-plus-oncosts planned rate {expected:.4f}.",
            activity.activity_id,
        )


def _validation_overhead_pct(assumptions: dict, budget_crore: float) -> float:
    """Resolve the same overhead slab used by DataLoader/CostCalculator."""
    explicit = (assumptions.get("cost_parameters") or {}).get("overhead_percentage")
    if explicit is not None:
        return float(explicit)
    statutory = assumptions.get("statutory_and_overheads") or {}
    for slab in statutory.get("overhead_slabs") or []:
        if budget_crore <= float(slab.get("max_budget_cr", float("inf"))):
            return float(slab.get("overhead_pct", 8.0))
    return float(statutory.get("overhead_pct", 8.0))


def _validate_cycles(result: Any, report: ValidationReport) -> None:
    graph = {activity.activity_id: list(activity.predecessors) for activity in result.activities}
    visiting = set()
    visited = set()

    def visit(node: str, path: list[str]) -> None:
        if node in visiting:
            report.add("ERROR", "CIRCULAR_DEPENDENCY", f"Cycle detected: {' -> '.join(path + [node])}.")
            return
        if node in visited:
            return
        visiting.add(node)
        for predecessor in graph.get(node, []):
            if predecessor in graph:
                visit(predecessor, path + [node])
        visiting.remove(node)
        visited.add(node)

    for node in graph:
        visit(node, [])


def _validate_engineering_sequence(result: Any, report: ValidationReport) -> None:
    """Surface pipeline handover gaps without imposing a new scheduling policy."""
    by_id = {activity.activity_id: activity for activity in result.activities}
    laying = by_id.get("ACT-18")
    if laying is not None and not laying.predecessors:
        report.add(
            "WARNING",
            "ENGINEERING_HANDOVER_REVIEW",
            "ACT-18 pipe laying has no explicit excavation/bedding handover predecessor; "
            "confirm whether section-wise handovers are represented elsewhere.",
            laying.activity_id,
        )
    backfill = by_id.get("ACT-19")
    commissioning = by_id.get("ACT-20")
    if backfill is not None and laying is not None and "ACT-18" not in backfill.predecessors:
        report.add("ERROR", "BACKFILL_BEFORE_LAYING", "ACT-19 must follow ACT-18.", backfill.activity_id)
    if commissioning is not None and backfill is not None and "ACT-19" not in commissioning.predecessors:
        report.add("ERROR", "COMMISSIONING_BEFORE_TEST", "ACT-20 must follow ACT-19.", commissioning.activity_id)


def _validate_activity_catalogue(
    result: Any,
    assumptions: Optional[dict],
    report: ValidationReport,
    project_type: str = "road",
) -> None:
    if not assumptions:
        return
    catalogue = assumptions.get("activity_master") or []
    if not catalogue:
        report.add("WARNING", "NO_ACTIVITY_MASTER", "Activity master was not available for catalogue validation.")
        return
    keys = {str(item.get("key")) for item in catalogue if item.get("key")}
    active = {str(a.activity_key) for a in result.activities}
    if not active <= keys:
        report.add("ERROR", "UNKNOWN_ACTIVITY", f"Scheduled activity keys not in activity master: {sorted(active - keys)}.")
    if any(float(a.quantity or 0.0) <= 0 for a in result.activities):
        severity = "ERROR" if project_type == "pipeline" else "WARNING"
        report.add(
            severity,
            "ZERO_ACTIVE_QUANTITY",
            "An activity row has zero quantity; it is retained for catalogue/audit visibility."
            if project_type != "pipeline"
            else "An active pipeline activity has zero quantity.",
        )


def _validate_resource_allocation(result: Any, assumptions: Optional[dict], report: ValidationReport) -> None:
    for activity in result.activities:
        if int(activity.crews or 0) < 1:
            report.add("ERROR", "INVALID_CREW_COUNT", "Active activity must have at least one crew.", activity.activity_id)
        if float(activity.shifts or 0.0) < 1:
            report.add("ERROR", "INVALID_SHIFT_COUNT", "Active activity must have at least one shift.", activity.activity_id)
    cap = float((assumptions or {}).get("crew_multiplier_cap", 0) or 0)
    if cap > 0:
        for activity in result.activities:
            if activity.crews > max(activity.base_crews, int(activity.base_crews * cap)):
                report.add("ERROR", "CREW_CAP_EXCEEDED", "Crew count exceeds configured cap.", activity.activity_id)


def _validate_escalation(result: Any, assumptions: Optional[dict], report: ValidationReport) -> None:
    cfg = (assumptions or {}).get("escalation") or {}
    if cfg.get("enabled"):
        try:
            base = date.fromisoformat(str(cfg.get("base_date"))[:10])
            if base.year < 2000:
                raise ValueError
        except (TypeError, ValueError):
            report.add("ERROR", "INVALID_ESCALATION_BASE_DATE", "Escalation base_date is missing or invalid.")
        rates = cfg.get("annual_rates") or {}
        for name in ("material", "labour", "plant", "fuel"):
            if float(rates.get(name, 0.0) or 0.0) < 0:
                report.add("ERROR", "NEGATIVE_ESCALATION_RATE", f"{name} escalation rate cannot be negative.")
    for activity in result.activities:
        cost = activity.cost
        if (
            float(cost.direct_cost or 0.0) > 0
            and (cost.escalation_years < 0 or cost.escalation_factor < 1.0 or cost.escalation_amount < 0)
        ):
            report.add("ERROR", "ESCALATION_OUTPUT_INVALID", "Escalation output is invalid.", activity.activity_id)


def _validate_calendar(result: Any, calendar: Optional[Any], report: ValidationReport) -> None:
    if calendar is None:
        report.add("WARNING", "CALENDAR_NOT_RECHECKED", "Calendar rules were not independently rechecked.")
        return
    start = result.project_input.start_date
    for activity in result.activities:
        if activity.duration_days < activity.effort_days:
            report.add("ERROR", "CALENDAR_SHORTER_THAN_EFFORT",
                       "Calendar duration is shorter than effort duration.", activity.activity_id)
        expected_duration = calendar.calendar_days_needed(
            start + timedelta(days=int(activity.es)),
            activity.effort_days,
        )
        if int(activity.duration_days) != int(expected_duration):
            report.add(
                "ERROR",
                "CALENDAR_DURATION_MISMATCH",
                f"Duration {activity.duration_days} does not match independent "
                f"calendar calculation {expected_duration}.",
                activity.activity_id,
            )
        if activity.duration_days:
            finish = start + timedelta(days=activity.ef)
            if finish < start + timedelta(days=activity.es):
                report.add("ERROR", "INVALID_ACTIVITY_DATES", "Activity finish precedes its start.", activity.activity_id)
    if calendar.working_days_per_week < 1 or calendar.working_days_per_week > 7:
        report.add("ERROR", "INVALID_WORKING_WEEK", "Working days per week must be between 1 and 7.")
    if not 0 < float(calendar.monsoon_factor) <= 1:
        report.add("ERROR", "INVALID_MONSOON_FACTOR", "Monsoon productivity factor must be in (0, 1].")


def _validate_optimization(result: Any, report: ValidationReport) -> None:
    if result.calculated_duration_days != max(a.ef for a in result.activities):
        report.add("ERROR", "OPTIMIZED_FINISH_MISMATCH", "Optimized finish does not match activity network.")
    target = int(result.project_input.duration_days or 0)
    budget = float(result.project_input.budget_crore or 0.0)
    if result.is_feasible and target > 0 and result.calculated_duration_days > target:
        report.add("ERROR", "FEASIBILITY_DURATION_MISMATCH", "Plan is marked feasible but exceeds duration target.")
    if result.is_feasible and budget > 0 and result.total_calculated_cost_crore > budget:
        report.add("ERROR", "FEASIBILITY_BUDGET_MISMATCH", "Plan is marked feasible but exceeds budget.")


def validate_exported_workbook(path: str, result: Any) -> ValidationReport:
    """Re-open an exported workbook and verify its structural control totals."""
    report = ValidationReport()
    try:
        import openpyxl
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except (OSError, ValueError, ImportError) as exc:
        report.add("ERROR", "WORKBOOK_OPEN_FAILED", f"Could not reopen workbook: {exc}")
        return report

    required = {
        "01_Project_Summary", "02_Project_Inputs", "05_Activities",
        "09_Optimized_CPM", "10_Optimized_Cost", "12_Monthly_Schedule",
        "13_S_Curve", "15_Validation_Checks", "Plan", "Engineer_DailyLog",
    }
    missing = sorted(required - set(workbook.sheetnames))
    if missing:
        report.add("ERROR", "WORKBOOK_SHEETS_MISSING", f"Missing required sheets: {missing}.")
    else:
        expected_plan_rows = len(getattr(result, "activities", []) or [])
        if workbook["Plan"].max_row < expected_plan_rows + 2:
            report.add("ERROR", "WORKBOOK_PLAN_EMPTY",
                       "Plan sheet does not contain all scheduled activity rows.")
        if workbook["Engineer_DailyLog"].max_row < 2:
            report.add("ERROR", "WORKBOOK_DAILY_LOG_EMPTY",
                       "Engineer_DailyLog does not contain planned execution rows.")
        _validate_exported_plan_calculations(workbook["Plan"], result, report)

    validation = workbook["15_Validation_Checks"] if "15_Validation_Checks" in workbook.sheetnames else None
    if validation is not None:
        status = validation["B4"].value
        errors = validation["B5"].value
        warnings = validation["B6"].value
        expected = getattr(getattr(result, "validation_report", None), "status", "NOT RUN")
        if status != expected:
            report.add("ERROR", "WORKBOOK_VALIDATION_STATUS_MISMATCH",
                       f"Workbook status {status!r} does not match in-memory status {expected!r}.")
        if errors != len(getattr(getattr(result, "validation_report", None), "errors", []) or []):
            report.add("ERROR", "WORKBOOK_ERROR_COUNT_MISMATCH", "Workbook error count differs from validation report.")
        if warnings != len(getattr(getattr(result, "validation_report", None), "warnings", []) or []):
            report.add("ERROR", "WORKBOOK_WARNING_COUNT_MISMATCH", "Workbook warning count differs from validation report.")

    activity_sheet = workbook["05_Activities"] if "05_Activities" in workbook.sheetnames else None
    if activity_sheet is not None:
        exported_ids = {
            row[2] for row in activity_sheet.iter_rows(min_row=2, values_only=True)
            if len(row) > 2 and isinstance(row[2], str) and row[2].startswith("ACT-")
        }
        expected_ids = {a.activity_id for a in result.activities}
        if exported_ids != expected_ids:
            report.add("ERROR", "WORKBOOK_ACTIVITY_SET_MISMATCH",
                       "Exported activity IDs differ from the in-memory optimized schedule.")
    curve = workbook["13_S_Curve"] if "13_S_Curve" in workbook.sheetnames else None
    if curve is not None:
        rows = [row for row in curve.iter_rows(min_row=5, values_only=True)
                if row and row[0] is not None]
        progress_total = sum(float(row[2] or 0.0) for row in rows)
        final_cumulative = float(rows[-1][3] or 0.0) if rows else 0.0
        cost_total = sum(float(row[4] or 0.0) for row in rows)
        if abs(progress_total - 1.0) > 0.002 or abs(final_cumulative - 1.0) > 0.002:
            report.add("ERROR", "WORKBOOK_SCURVE_PROGRESS_MISMATCH",
                       "S-curve monthly progress does not reconcile to 100%.")
        if abs(cost_total - float(result.total_calculated_cost_crore)) > 0.01:
            report.add("ERROR", "WORKBOOK_SCURVE_COST_MISMATCH",
                       "S-curve monthly costs do not reconcile to project total.")
    return report


def _validate_exported_plan_calculations(
    sheet: Any,
    result: Any,
    report: ValidationReport,
) -> None:
    """Check the persisted Plan activity rows and totals against source costs."""
    from .plan_generator import build_plan

    expected = {row.activity_id: row for row in build_plan(result)}
    seen = set()
    project_execution = 0.0
    project_billed = 0.0
    for values in sheet.iter_rows(min_row=3, values_only=True):
        activity_id = values[0] if values else None
        row_type = values[12] if len(values) > 12 else None
        if row_type == "Activity" and isinstance(activity_id, str):
            row = expected.get(activity_id)
            if row is None:
                report.add("ERROR", "WORKBOOK_PLAN_UNKNOWN_ACTIVITY",
                           f"Plan contains unknown activity {activity_id}.", activity_id)
                continue
            seen.add(activity_id)
            checks = (
                ("quantity", values[7], row.quantity, 0.1),
                ("duration", values[6], row.duration_days, 0),
                ("crew", values[10], row.crew_multiplier, 0),
            )
            for label, actual, source, tolerance in checks:
                if actual is None or not _near(actual, source, max(tolerance, 0.05)):
                    report.add(
                        "ERROR",
                        "WORKBOOK_PLAN_INPUT_MISMATCH",
                        f"Plan {label} does not match the generated Plan data.",
                        activity_id,
                    )
        elif row_type == "Subtotal" and isinstance(activity_id, str):
            row = expected.get(activity_id)
            if row is None:
                continue
            execution = float(values[29] or 0.0)
            billed = float(values[30] or 0.0)
            project_execution += execution
            project_billed += billed
            if not _near(execution, row.total_execution_cost, 1.0):
                report.add("ERROR", "WORKBOOK_PLAN_EXECUTION_MISMATCH",
                           "Plan subtotal does not match the original direct cost.",
                           activity_id)
            source_activity = next(
                (item for item in result.activities if item.activity_id == activity_id),
                None,
            )
            if source_activity is not None and not _near(
                billed, float(source_activity.cost.total_amount or 0.0), 1.0
            ):
                report.add("ERROR", "WORKBOOK_PLAN_BILLED_MISMATCH",
                           "Plan billed subtotal does not match the original total cost.",
                           activity_id)

    if seen != set(expected):
        report.add("ERROR", "WORKBOOK_PLAN_ACTIVITY_SET_MISMATCH",
                   "Persisted Plan activity rows differ from the generated Plan.")

    for values in sheet.iter_rows(min_row=3, values_only=True):
        if values and values[0] == "PROJECT TOTALS":
            execution = float(values[29] or 0.0)
            billed = float(values[30] or 0.0)
            if not _near(execution, project_execution, 1.0):
                report.add("ERROR", "WORKBOOK_PLAN_PROJECT_EXECUTION_MISMATCH",
                           "Plan project execution total does not equal activity subtotals.")
            if not _near(billed, project_billed, 1.0):
                report.add("ERROR", "WORKBOOK_PLAN_PROJECT_BILLED_MISMATCH",
                           "Plan project billed total does not equal activity subtotals.")
            break


def validate_update_result(result: Any, assumptions: Optional[dict] = None) -> ValidationReport:
        """Validate an update-plan result and its remaining-work cost model."""
        from types import SimpleNamespace

        report = ValidationReport()
        revised = list(getattr(result, "revised_activities", []) or [])
        statuses = list(getattr(result, "activity_status", []) or [])
        baseline = getattr(result, "baseline_result", None)
        if not revised and statuses:
            report.add("ERROR", "UPDATE_NO_REVISED_ACTIVITIES", "Update contains status rows but no revised activities.")

        baseline_ids = {a.activity_id for a in getattr(baseline, "activities", []) or []}
        status_ids = {s.activity_id for s in statuses}
        if baseline_ids and status_ids != baseline_ids:
            report.add("ERROR", "UPDATE_ACTIVITY_STATUS_SET_MISMATCH",
                       "Activity status rows do not match the baseline activity set.")

        revised_ids = {a.activity_id for a in revised}
        if not revised_ids <= status_ids:
            report.add("ERROR", "UPDATE_REVISED_ACTIVITY_SET_MISMATCH",
                       "Revised activities contain IDs without corresponding actual-status rows.")

        for status in statuses:
            pct = float(status.actual_completion_pct or 0.0)
            if not 0.0 <= pct <= 100.0:
                report.add("ERROR", "UPDATE_PROGRESS_OUT_OF_RANGE",
                           f"Actual completion {pct:.2f}% is outside 0-100%.", status.activity_id)
            if abs(
                float(status.completed_quantity or 0.0)
                + float(status.remaining_quantity or 0.0)
                - float(status.original_quantity or 0.0)
            ) > 0.05:
                report.add("ERROR", "UPDATE_QUANTITY_RECONCILIATION",
                           "Completed quantity plus remaining quantity does not equal original quantity.",
                           status.activity_id)

        by_id = {a.activity_id: a for a in revised}
        for activity in revised:
            aid = activity.activity_id
            if activity.ef != activity.es + activity.duration_days:
                report.add("ERROR", "UPDATE_CPM_EF_IDENTITY", "Revised EF must equal ES plus duration.", aid)
            if activity.lf != activity.ls + activity.duration_days:
                report.add("ERROR", "UPDATE_CPM_LF_IDENTITY", "Revised LF must equal LS plus duration.", aid)
            if activity.total_float != activity.ls - activity.es:
                report.add("ERROR", "UPDATE_CPM_FLOAT_IDENTITY", "Revised float must equal LS minus ES.", aid)
            for predecessor in activity.predecessors:
                if predecessor not in by_id:
                    report.add("ERROR", "UPDATE_MISSING_PREDECESSOR",
                               f"Revised predecessor {predecessor} is not present.", aid)
                elif activity.es < by_id[predecessor].ef + int(
                    (activity.predecessor_lags or {}).get(predecessor, 0)
                ):
                    report.add("ERROR", "UPDATE_PRECEDENCE_VIOLATION",
                               "Revised activity starts before its predecessor is complete.", aid)
            cost = activity.cost
            expected_direct = (
                float(activity.quantity or 0.0) * float(cost.planned_unit_rate or 0.0)
                + float(cost.wastage_cost or 0.0)
                + float(cost.crash_total_surcharge or 0.0)
            )
            if not _near(cost.direct_cost, expected_direct, 0.05):
                report.add("ERROR", "UPDATE_DIRECT_COST_RECONCILIATION",
                           f"Direct cost {cost.direct_cost:.2f} does not match planned rate, "
                           f"wastage and surcharge {expected_direct:.2f}.", aid)
            expected_total = (
                float(cost.direct_cost or 0.0)
                + float(cost.overhead_amount or 0.0)
                + float(cost.profit_amount or 0.0)
                + float(cost.escalation_amount or 0.0)
            )
            if not _near(cost.total_amount, expected_total, 0.05):
                report.add("ERROR", "UPDATE_TOTAL_COST_RECONCILIATION",
                           "Activity total does not equal direct cost plus escalation.", aid)
            wastage_sum = sum(float(getattr(cost, field, 0.0) or 0.0) for field in (
                "material_wastage_cost", "fuel_wastage_cost",
                "plant_wastage_cost", "labour_wastage_cost",
            ))
            if not _near(cost.wastage_cost, wastage_sum, 0.05):
                report.add("ERROR", "UPDATE_WASTAGE_RECONCILIATION",
                           "Wastage total does not equal its component losses.", aid)
            if cost.escalation_factor < 1.0 or cost.escalation_amount < 0:
                report.add("ERROR", "UPDATE_ESCALATION_OUTPUT_INVALID",
                           "Escalation factor or amount is invalid.", aid)

        max_finish = max((a.ef for a in revised), default=0)
        if max_finish != int(getattr(result, "revised_duration_days", max_finish)):
            report.add("ERROR", "UPDATE_DURATION_RECONCILIATION",
                       "Revised duration does not equal the latest revised activity finish.")
        remaining = sum(float(a.cost.total_amount or 0.0) for a in revised) / 10000000.0
        if not _near(remaining, float(getattr(result, "remaining_cost_crore", 0.0)), 0.01):
            report.add("ERROR", "UPDATE_REMAINING_COST_RECONCILIATION",
                       "Remaining cost does not equal revised activity totals.")
        total = float(getattr(result, "spent_cost_crore", 0.0)) + float(
            getattr(result, "remaining_cost_crore", 0.0)
        )
        if not _near(total, float(getattr(result, "total_project_cost_crore", 0.0)), 0.01):
            report.add("ERROR", "UPDATE_PROJECT_COST_RECONCILIATION",
                       "Spent plus remaining cost does not equal revised project cost.")
        monthly_total = sum(
            float(item.period_expenditure_crore or 0.0)
            for item in getattr(result, "revised_monthly_progress", []) or []
        )
        if not _near(monthly_total, remaining, 0.01):
            report.add("ERROR", "UPDATE_SCURVE_COST_RECONCILIATION",
                       "Revised S-curve expenditure does not equal remaining cost.")
        if getattr(result, "revised_completion_date", None) is None:
            report.add("ERROR", "UPDATE_COMPLETION_DATE_MISSING",
                       "Revised completion date was not calculated.")
        if int(getattr(result, "projected_delay_days", 0) or 0) < 0:
            report.add("ERROR", "UPDATE_NEGATIVE_DELAY", "Projected delay cannot be negative.")
        for warning in getattr(result, "warnings", []) or []:
            report.add("WARNING", "UPDATE_PLANNING_WARNING", str(warning))
        revised_view = SimpleNamespace(
            project_input=getattr(result, "project_input", None),
            activities=revised,
            assumptions_snapshot=assumptions or getattr(
                baseline, "assumptions_snapshot", {}
            ) or {},
        )
        _validate_resource_plan_and_daily_log(revised_view, report)
        return report


def validate_exported_update_workbook(path: str, result: Any) -> ValidationReport:
        """Validate the update workbook structure and its validation summary."""
        import openpyxl

        report = ValidationReport()
        workbook = openpyxl.load_workbook(path, data_only=True)
        required = {
            "01_Update_Summary", "02_Baseline_Schedule", "03_Actual_Progress",
            "04_Activity_Status", "05_Revised_CPM", "06_Revised_Cost",
            "07_S_Curve_Comparison", "08_Target_Feasibility", "Plan",
            "Engineer_DailyLog", "09_Validation_Checks",
        }
        missing = sorted(required - set(workbook.sheetnames))
        if missing:
            report.add("ERROR", "UPDATE_WORKBOOK_SHEETS_MISSING", f"Missing update sheets: {missing}.")
        else:
            if workbook["Plan"].max_row < 2:
                report.add("ERROR", "UPDATE_WORKBOOK_PLAN_EMPTY",
                           "Updated Plan sheet does not contain revised work.")
            if workbook["Engineer_DailyLog"].max_row < 2:
                report.add("ERROR", "UPDATE_WORKBOOK_DAILY_LOG_EMPTY",
                           "Updated Engineer_DailyLog does not contain revised work.")
        sheet = workbook["09_Validation_Checks"] if "09_Validation_Checks" in workbook.sheetnames else None
        expected = getattr(getattr(result, "validation_report", None), "status", "NOT RUN")
        if sheet is not None:
            if sheet["B4"].value != expected:
                report.add("ERROR", "UPDATE_WORKBOOK_STATUS_MISMATCH",
                           "Workbook validation status differs from the in-memory report.")
            if sheet["B5"].value != len(getattr(getattr(result, "validation_report", None), "errors", []) or []):
                report.add("ERROR", "UPDATE_WORKBOOK_ERROR_COUNT_MISMATCH",
                           "Workbook error count differs from the in-memory report.")
        revised_sheet = workbook["05_Revised_CPM"] if "05_Revised_CPM" in workbook.sheetnames else None
        if revised_sheet is not None:
            exported = {
                row[0] for row in revised_sheet.iter_rows(min_row=5, values_only=True)
                if row and isinstance(row[0], str) and row[0].startswith("ACT-")
            }
            if exported != {a.activity_id for a in getattr(result, "revised_activities", []) or []}:
                report.add("ERROR", "UPDATE_WORKBOOK_ACTIVITY_SET_MISMATCH",
                           "Exported revised activity IDs differ from the update result.")
        return report

def format_validation_report(report: ValidationReport) -> str:
    """Return a concise human-readable report for CLI output and logs."""
    lines = [
        f"Validation: {'PASS' if report.passed else 'FAIL'}",
        f"Errors: {len(report.errors)}",
        f"Warnings: {len(report.warnings)}",
    ]
    for issue in report.issues:
        suffix = f" [{issue.activity_id}]" if issue.activity_id else ""
        lines.append(f"{issue.severity} {issue.code}{suffix}: {issue.message}")
    return "\n".join(lines)
