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
from collections import Counter
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
    def incomplete(self) -> List[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == "INCOMPLETE"]

    @property
    def passed(self) -> bool:
        return not self.errors and not self.incomplete

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
        if self.errors:
            return "FAIL"
        if self.incomplete:
            return "INCOMPLETE"
        return "PASS"


def _near(left: float, right: float, absolute: float = 0.05) -> bool:
    return abs(float(left) - float(right)) <= absolute


def _validate_activity_structure(
    result: Any,
    activities: List[Any],
    report: ValidationReport,
    *,
    prefix: str = "",
) -> bool:
    """Validate the minimum shape required by the activity checks."""
    if result is None:
        report.add("ERROR", f"{prefix}RESULT_MISSING", "Validation result is missing.")
        return False
    if not hasattr(result, "project_input") or result.project_input is None:
        report.add("ERROR", f"{prefix}PROJECT_INPUT_MISSING", "Project input is missing.")
    if not activities:
        report.add("ERROR", f"{prefix}NO_ACTIVITIES", "The plan contains no activities.")
        return False

    ids = []
    structurally_valid = True
    for index, activity in enumerate(activities):
        aid = getattr(activity, "activity_id", None)
        if not isinstance(aid, str) or not aid.strip():
            report.add(
                "ERROR",
                f"{prefix}ACTIVITY_ID_MISSING",
                f"Activity at index {index} has no valid activity ID.",
            )
            structurally_valid = False
        else:
            ids.append(aid)
        for field_name in ("activity_key", "quantity", "productivity_rate", "cost"):
            if not hasattr(activity, field_name) or getattr(activity, field_name) is None:
                report.add(
                    "ERROR",
                    f"{prefix}ACTIVITY_FIELD_MISSING",
                    f"Activity {aid or index} is missing required field '{field_name}'.",
                    aid if isinstance(aid, str) else None,
                )
                structurally_valid = False
        if hasattr(activity, "cost") and getattr(activity, "cost", None) is not None:
            for field_name in ("planned_unit_rate", "direct_cost", "total_amount"):
                if not hasattr(activity.cost, field_name):
                    report.add(
                        "ERROR",
                        f"{prefix}COST_FIELD_MISSING",
                        f"Activity {aid or index} cost is missing '{field_name}'.",
                        aid if isinstance(aid, str) else None,
                    )
                    structurally_valid = False

    for aid, count in Counter(ids).items():
        if count > 1:
            report.add(
                "ERROR",
                f"{prefix}DUPLICATE_ACTIVITY_ID",
                f"Activity ID {aid} occurs {count} times.",
                aid,
            )
            structurally_valid = False
    return structurally_valid


def _run_check(report: ValidationReport, code: str, check: Any) -> None:
    """Report unexpected validator failures instead of allowing a false pass."""
    try:
        check()
    except Exception as exc:
        report.add("ERROR", code, f"Validation check could not complete: {exc}")


def _validate_schedule_result_impl(
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
    if not _validate_activity_structure(result, activities, report):
        return report
    by_id = {a.activity_id: a for a in activities}

    project_type = str(getattr(result.project_input, "project_type", "road") or "road").lower()
    budget_crore = float(getattr(result.project_input, "budget_crore", 0.0) or 0.0)
    overhead_pct = _validation_overhead_pct(assumptions or {}, budget_crore)
    _run_check(
        report, "ASSUMPTION_VALIDATION_EXCEPTION",
        lambda: _validate_assumption_sources(result, assumptions or {}, report),
    )
    if project_type == "pipeline" and assumptions:
        _run_check(
            report, "PIPELINE_INPUT_VALIDATION_EXCEPTION",
            lambda: _validate_pipeline_inputs(result, assumptions, report, tolerance_inr),
        )
        _run_check(
            report, "ENGINEERING_SEQUENCE_VALIDATION_EXCEPTION",
            lambda: _validate_engineering_sequence(result, report),
        )
    _run_check(
        report, "ACTIVITY_CATALOGUE_VALIDATION_EXCEPTION",
        lambda: _validate_activity_catalogue(result, assumptions, report, project_type),
    )
    _run_check(
        report, "RESOURCE_ALLOCATION_VALIDATION_EXCEPTION",
        lambda: _validate_resource_allocation(result, assumptions, report),
    )
    _run_check(
        report, "ESCALATION_VALIDATION_EXCEPTION",
        lambda: _validate_escalation(result, assumptions, report),
    )
    _run_check(
        report, "CYCLE_VALIDATION_EXCEPTION",
        lambda: _validate_cycles(result, report),
    )
    _run_check(
        report, "PLAN_VALIDATION_EXCEPTION",
        lambda: _validate_plan_calculations(result, report, tolerance_inr),
    )
    _run_check(
        report, "RESOURCE_PLAN_VALIDATION_EXCEPTION",
        lambda: _validate_resource_plan_and_daily_log(result, report),
    )

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
        reference = getattr(activity, "reference", None)
        if reference and not getattr(reference, "rate_source", ""):
            report.add("WARNING", "RATE_METADATA_MISSING",
                       "Rate source metadata is missing.", aid)
        if reference and getattr(reference, "rate_verification_status", "unverified") == "unverified":
            report.add("WARNING", "RATE_NOT_VERIFIED",
                       "Rate metadata is provisional or unverified.", aid)

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
        elif assumptions and assumptions.get("rate_build_up_mode") == "resource_build_up":
            _validate_resource_build_up_rate(
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
        previous_cumulative = 0.0
        for item in monthly:
            period_progress = float(item.period_planned_progress_pct or 0.0)
            cumulative_progress = float(item.cumulative_planned_progress_pct or 0.0)
            if period_progress < 0 or cumulative_progress < 0:
                report.add("ERROR", "NEGATIVE_PROGRESS", "Progress cannot be negative.")
            if cumulative_progress + 0.2 < previous_cumulative:
                report.add(
                    "ERROR",
                    "PROGRESS_NOT_MONOTONIC",
                    "Cumulative planned progress cannot decrease.",
                )
            previous_cumulative = cumulative_progress

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


def validate_schedule_result(
    result: Any,
    assumptions: Optional[dict] = None,
    *,
    tolerance_inr: float = 0.05,
    calendar: Optional[Any] = None,
) -> ValidationReport:
    """Run schedule validation and report unexpected validator failures."""
    try:
        return _validate_schedule_result_impl(
            result,
            assumptions,
            tolerance_inr=tolerance_inr,
            calendar=calendar,
        )
    except Exception as exc:
        report = ValidationReport()
        report.add("ERROR", "VALIDATION_EXCEPTION",
                   f"Schedule validation could not complete: {exc}")
        return report


def _validate_assumption_sources(
    result: Any,
    assumptions: dict,
    report: ValidationReport,
) -> None:
    """Verify that scheduled engineering inputs resolve to assumptions."""
    catalog = {
        str(item.get("key")): item
        for item in assumptions.get("activity_master", []) or []
        if item.get("key")
    }
    rates = assumptions.get("direct_unit_rates_inr", {}) or {}
    if not catalog:
        report.add(
            "ERROR",
            "ASSUMPTION_SOURCE_MISSING",
            "No activity_master assumptions were loaded for the schedule.",
        )
        return

    for activity in list(getattr(result, "activities", []) or []):
        aid = activity.activity_id
        key = str(getattr(activity, "activity_key", "") or "")
        spec = catalog.get(key)
        if spec is None:
            report.add(
                "ERROR",
                "ASSUMPTION_SOURCE_MISSING",
                f"Activity key '{key}' has no activity_master assumption.",
                aid,
            )
            continue
        missing = [
            field
            for field in ("productivity_rate", "productivity_unit", "base_crews")
            if spec.get(field) in (None, "")
        ]
        if missing:
            report.add(
                "ERROR",
                "ASSUMPTION_SOURCE_MISSING",
                f"activity_master[{key}] is missing: {', '.join(missing)}.",
                aid,
            )
        if key not in rates or rates.get(key) in (None, ""):
            report.add(
                "ERROR",
                "ASSUMPTION_RATE_MISSING",
                f"direct_unit_rates_inr[{key}] is missing.",
                aid,
            )
        else:
            try:
                if float(rates[key]) <= 0:
                    report.add(
                        "ERROR",
                        "ASSUMPTION_RATE_INVALID",
                        f"direct_unit_rates_inr[{key}] must be positive.",
                        aid,
                    )
            except (TypeError, ValueError):
                report.add(
                    "ERROR",
                    "ASSUMPTION_RATE_INVALID",
                    f"direct_unit_rates_inr[{key}] must be numeric.",
                    aid,
                )


def _validate_resource_plan_and_daily_log(result: Any, report: ValidationReport) -> None:
    """Validate the report-facing Plan and Engineer_DailyLog data."""
    from .engineer_daily_log import build_daily_log
    from .plan_generator import build_plan

    plan = build_plan(result)
    activities = {
        activity.activity_id: activity
        for activity in list(getattr(result, "activities", []) or [])
    }
    plan_ids = [row.activity_id for row in plan]
    duplicate_plan_ids = [aid for aid, count in Counter(plan_ids).items() if count > 1]
    if duplicate_plan_ids:
        report.add(
            "ERROR",
            "PLAN_DUPLICATE_ACTIVITY",
            f"Plan contains duplicate activity rows: {sorted(duplicate_plan_ids)}.",
        )
    if set(plan_ids) != set(activities):
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

    active_plan_ids = {
        row.activity_id for row in plan if row.status == "Active"
    }
    if active_plan_ids and not active_plan_ids <= set(activity_days):
        report.add(
            "ERROR",
            "DAILY_LOG_ACTIVITY_COVERAGE",
            "Daily log is missing records for one or more active plan activities.",
        )

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
        source_material_cost = float(
            cost.direct_material or cost.bottom_up_material_cost or 0.0
        )
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
                source_material_cost - row.unitemized_rate_material_cost,
                tolerance_inr,
            ),
            (
                "unitemized source-rate material allowance",
                row.unitemized_rate_material_cost,
                source_material_cost
                if (
                    (getattr(result, "assumptions_snapshot", {}) or {}).get(
                        "rate_build_up_mode"
                    ) in ("source_plus_oncosts", "resource_build_up")
                    and not cost.material_lines
                )
                else 0.0,
                tolerance_inr,
            ),
            ("source unit rate", row.source_unit_rate,
             float(cost.morth_direct_unit_rate or 0.0), tolerance_inr),
            ("planned unit rate", row.planned_unit_rate,
             float(cost.planned_unit_rate or 0.0), tolerance_inr),
            ("planning variation", row.planning_variation_pct,
             float(cost.planning_variation_pct or 0.0), 0.0001),
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


def _validate_resource_build_up_rate(
    activity: Any,
    assumptions: dict,
    report: ValidationReport,
    tolerance_inr: float,
    overhead_pct: float,
) -> None:
    """Independently reconcile a resource-derived planned unit rate."""
    activity_key = activity.activity_key
    rates = assumptions.get("direct_unit_rates_inr") or {}
    source = rates.get(activity_key)
    if isinstance(source, dict):
        source = source.get("value")
    if source is None:
        report.add("ERROR", "RATE_SOURCE_MISSING", "Source rate is missing.", activity.activity_id)
        return
    source = float(source)
    if assumptions.get("use_direct_rate_as_planned", False):
        if not _near(activity.cost.planned_unit_rate, source, tolerance_inr):
            report.add(
                "ERROR", "UNIT_RATE_BUILDUP_MISMATCH",
                f"Direct-rate override produced {activity.cost.planned_unit_rate:.4f}; "
                f"source rate is {source:.4f}.", activity.activity_id,
            )
        if abs(float(activity.cost.planning_variation_pct or 0.0)) > 0.0001:
            report.add(
                "ERROR", "PLANNING_VARIATION_MISMATCH",
                "Direct-rate override must have zero planning variation.",
                activity.activity_id,
            )
        return

    spec = next(
        (item for item in assumptions.get("activity_master", []) or []
         if item.get("key") == activity_key),
        None,
    )
    if not spec:
        report.add("ERROR", "RATE_ACTIVITY_MASTER_MISSING",
                   "Activity master data is missing for the planned rate.", activity.activity_id)
        return
    productivity = float(spec.get("productivity_rate") or 0.0)
    base_crews = int(spec.get("base_crews") or 0)
    if productivity <= 0 or base_crews <= 0:
        report.add("ERROR", "RATE_PRODUCTIVITY_MISSING",
                   "Positive productivity and base crews are required for the planned rate.",
                   activity.activity_id)
        return

    category = spec.get("category")
    shares = (assumptions.get("cost_breakdown_shares_by_category") or {}).get(
        category, assumptions.get("cost_breakdown_shares") or {}
    )
    material_share = float(shares.get("material_share", 0.0) or 0.0)
    material_per_unit = 0.0
    recipe_spec = (assumptions.get("activity_material_recipe") or {}).get(activity_key)
    if recipe_spec:
        recipes = {}
        for key in ("material_per_unit_pipeline", "material_per_unit",
                    "material_per_unit_road"):
            recipes.update(assumptions.get(key, {}) or {})
        recipe_name = str(recipe_spec.get("recipe") or "")
        seen = set()
        while recipe_name:
            if recipe_name in seen or recipe_name not in recipes:
                report.add("ERROR", "RATE_RECIPE_INVALID",
                           f"Material recipe '{recipe_name}' is missing or circular.",
                           activity.activity_id)
                return
            seen.add(recipe_name)
            recipe = recipes[recipe_name]
            alias = recipe.get("_use") if isinstance(recipe, dict) else None
            if not alias:
                break
            recipe_name = str(alias)
        bindings = assumptions.get("material_bindings") or {}
        material_rates = assumptions.get("material_rates_inr") or {}
        mix = float(recipe_spec.get("per_output", 1.0) or 1.0)
        for material, per_output in recipe.items():
            if str(material).startswith("_"):
                continue
            binding = bindings.get(material)
            if not binding or binding.get("rate") not in material_rates:
                report.add("ERROR", "RATE_RECIPE_BINDING_MISSING",
                           f"Material '{material}' has no usable rate binding.",
                           activity.activity_id)
                return
            divisor = float(binding.get("divisor", 1) or 1)
            if divisor <= 0:
                report.add("ERROR", "RATE_RECIPE_DIVISOR_INVALID",
                           f"Material '{material}' has an invalid rate divisor.",
                           activity.activity_id)
                return
            material_per_unit += (
                float(per_output) * mix / divisor
                * float(material_rates[binding["rate"]])
            )
    else:
        material_per_unit = source * material_share

    profile = (assumptions.get("crew_profile_by_activity_key") or {}).get(
        activity_key, ""
    )
    crew_items = (assumptions.get("crew_composition") or {}).get(profile, []) or []
    labour_rates = assumptions.get("labour_rates_inr_per_day") or {}
    labour_per_day = sum(
        float(item.get("count", 0) or 0)
        * float(labour_rates.get(item.get("trade"), 0.0) or 0.0)
        for item in crew_items
    )

    equipment = (assumptions.get("equipment_per_activity") or {}).get(profile, []) or []
    equipment_rates = assumptions.get("equipment_hire_inr_per_day") or {}
    major_plant = set(assumptions.get("major_plant") or [])
    major_rate = sum(
        float(equipment_rates.get(name, 0.0) or 0.0)
        for name in equipment if name in major_plant
    )
    mobile_rate = sum(
        float(equipment_rates.get(name, 0.0) or 0.0)
        for name in equipment if name not in major_plant
    )
    major_sets = min(base_crews, max(
        1, int(assumptions.get("major_plant_max_units", 0) or 0)
    ))
    plant_per_unit = (
        mobile_rate + major_rate * major_sets / base_crews
    ) / productivity

    fuel_rates = assumptions.get("fuel_consumption_lph") or {}
    fuel_lph = sum(float(fuel_rates.get(name, 0.0) or 0.0) for name in equipment)
    calendar = assumptions.get("calendar") or {}
    diesel = float((assumptions.get("material_rates_inr") or {}).get(
        "diesel_litre", 0.0
    ) or 0.0)
    fuel_per_unit = (
        fuel_lph * float(calendar.get("equipment_hours_per_day", 8.0) or 8.0)
        * diesel / productivity
    )
    bare_rate = material_per_unit + labour_per_day / productivity + plant_per_unit + fuel_per_unit

    oncost_basis = str(assumptions.get("rate_oncost_basis", "sdb") or "sdb")
    if oncost_basis == "statutory":
        profit = float((assumptions.get("statutory_and_overheads") or {}).get(
            "profit_pct", 10.0
        ))
        oncost = (1 + overhead_pct / 100) * (1 + profit / 100)
    elif oncost_basis in ("none", "off"):
        oncost = 1.0
    else:
        sdb = assumptions.get("sdb_rate_oncost") or {}
        group = (
            "structures"
            if category in set(sdb.get("structure_categories", []) or ["Structures"])
            else "road"
        )
        pct = sdb.get(group) or {}
        oh = float(pct.get("overhead_pct", 8.0))
        profit = float(pct.get("contractor_profit_pct", 10.0))
        oncost = (1 + oh / 100) * (1 + profit / 100)

    expected = round(bare_rate * oncost, 4)
    variation = (bare_rate * oncost - source) / source * 100 if source else 0.0
    if abs(float(activity.cost.planned_unit_rate) - expected) > max(
        tolerance_inr, abs(expected) * 1e-5
    ):
        report.add(
            "ERROR", "UNIT_RATE_BUILDUP_MISMATCH",
            f"Planned rate {activity.cost.planned_unit_rate:.4f} does not match "
            f"resource-derived rate {expected:.4f}.", activity.activity_id,
        )
    if abs(float(activity.cost.planning_variation_pct) - variation) > max(
        0.0001, abs(variation) * 1e-5
    ):
        report.add(
            "ERROR", "PLANNING_VARIATION_MISMATCH",
            f"Planning variation {activity.cost.planning_variation_pct:.4f}% does not "
            f"match the source-rate comparison {variation:.4f}%.",
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
    except Exception as exc:
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
        _run_check(
            report,
            "WORKBOOK_PLAN_VALIDATION_EXCEPTION",
            lambda: _validate_exported_plan_calculations(workbook["Plan"], result, report),
        )

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
        exported_id_values = [
            row[2] for row in activity_sheet.iter_rows(min_row=2, values_only=True)
            if len(row) > 2 and isinstance(row[2], str) and row[2].startswith("ACT-")
        ]
        exported_ids = set(exported_id_values)
        duplicate_ids = [
            activity_id for activity_id, count in Counter(exported_id_values).items()
            if count > 1
        ]
        if duplicate_ids:
            report.add(
                "ERROR",
                "WORKBOOK_DUPLICATE_ACTIVITY_ID",
                f"Exported activity IDs are duplicated: {sorted(duplicate_ids)}.",
            )
        expected_ids = {a.activity_id for a in result.activities}
        if exported_ids != expected_ids:
            report.add("ERROR", "WORKBOOK_ACTIVITY_SET_MISMATCH",
                       "Exported activity IDs differ from the in-memory optimized schedule.")
    curve = workbook["13_S_Curve"] if "13_S_Curve" in workbook.sheetnames else None
    if curve is not None:
        rows = [row for row in curve.iter_rows(min_row=5, values_only=True)
                if row and row[0] is not None]
        if not rows:
            report.add("ERROR", "WORKBOOK_SCURVE_EMPTY", "S-curve contains no monthly rows.")
            return report
        try:
            progress_total = sum(float(row[2] or 0.0) for row in rows)
            final_cumulative = float(rows[-1][3] or 0.0)
            cost_total = sum(float(row[4] or 0.0) for row in rows)
        except (TypeError, ValueError, IndexError) as exc:
            report.add("ERROR", "WORKBOOK_SCURVE_ROW_INVALID",
                       f"S-curve contains an invalid monthly value: {exc}.")
            return report
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
    activity_row_ids = []
    subtotal_row_ids = []
    project_execution = 0.0
    project_billed = 0.0
    project_totals_rows = []
    for values in sheet.iter_rows(min_row=3, values_only=True):
        activity_id = values[0] if values else None
        row_type = values[12] if len(values) > 12 else None
        if row_type == "Activity" and isinstance(activity_id, str):
            activity_row_ids.append(activity_id)
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
            subtotal_row_ids.append(activity_id)
            row = expected.get(activity_id)
            if row is None:
                continue
            if len(values) <= 30:
                report.add("ERROR", "WORKBOOK_PLAN_ROW_INCOMPLETE",
                           "Plan subtotal row is missing required total columns.",
                           activity_id)
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

    duplicate_activity_rows = [
        aid for aid, count in Counter(activity_row_ids).items() if count > 1
    ]
    duplicate_subtotal_rows = [
        aid for aid, count in Counter(subtotal_row_ids).items() if count > 1
    ]
    if duplicate_activity_rows:
        report.add("ERROR", "WORKBOOK_PLAN_DUPLICATE_ACTIVITY_ROW",
                   f"Plan contains duplicate activity rows: {sorted(duplicate_activity_rows)}.")
    if duplicate_subtotal_rows:
        report.add("ERROR", "WORKBOOK_PLAN_DUPLICATE_SUBTOTAL",
                   f"Plan contains duplicate subtotal rows: {sorted(duplicate_subtotal_rows)}.")
    if set(activity_row_ids) != set(expected) or set(subtotal_row_ids) != set(expected):
        report.add("ERROR", "WORKBOOK_PLAN_ACTIVITY_SET_MISMATCH",
                   "Persisted Plan activity and subtotal rows differ from the generated Plan.")

    for values in sheet.iter_rows(min_row=3, values_only=True):
        if values and values[0] == "PROJECT TOTALS":
            project_totals_rows.append(values)
    if len(project_totals_rows) != 1:
        report.add(
            "ERROR",
            "WORKBOOK_PLAN_PROJECT_TOTALS_MISSING" if not project_totals_rows
            else "WORKBOOK_PLAN_PROJECT_TOTALS_DUPLICATE",
            "Plan must contain exactly one PROJECT TOTALS row.",
        )
    else:
        values = project_totals_rows[0]
        if len(values) <= 30:
            report.add("ERROR", "WORKBOOK_PLAN_TOTAL_ROW_INCOMPLETE",
                       "PROJECT TOTALS row is missing required total columns.")
        else:
            execution = float(values[29] or 0.0)
            billed = float(values[30] or 0.0)
            if not _near(execution, project_execution, 1.0):
                report.add("ERROR", "WORKBOOK_PLAN_PROJECT_EXECUTION_MISMATCH",
                           "Plan project execution total does not equal activity subtotals.")
            if not _near(billed, project_billed, 1.0):
                report.add("ERROR", "WORKBOOK_PLAN_PROJECT_BILLED_MISMATCH",
                           "Plan project billed total does not equal activity subtotals.")


def _validate_update_result_impl(result: Any, assumptions: Optional[dict] = None) -> ValidationReport:
        """Validate an update-plan result and its remaining-work cost model."""
        from types import SimpleNamespace

        report = ValidationReport()
        if result is None:
            report.add("ERROR", "UPDATE_RESULT_MISSING", "Update result is missing.")
            return report
        revised = list(getattr(result, "revised_activities", []) or [])
        statuses = list(getattr(result, "activity_status", []) or [])
        baseline = getattr(result, "baseline_result", None)
        if baseline is None:
            report.add("ERROR", "UPDATE_BASELINE_MISSING", "Update baseline result is missing.")
            return report

        baseline_activities = list(getattr(baseline, "activities", []) or [])
        baseline_id_values = [getattr(a, "activity_id", None) for a in baseline_activities]
        baseline_ids = {aid for aid in baseline_id_values if isinstance(aid, str) and aid}
        if not baseline_ids:
            report.add("ERROR", "UPDATE_BASELINE_ACTIVITIES_MISSING",
                       "Update baseline contains no valid activities.")
            return report
        for label, values, code in (
            ("baseline", baseline_id_values, "UPDATE_DUPLICATE_BASELINE_ID"),
            ("status", [getattr(s, "activity_id", None) for s in statuses],
             "UPDATE_DUPLICATE_STATUS_ID"),
            ("revised", [getattr(a, "activity_id", None) for a in revised],
             "UPDATE_DUPLICATE_REVISED_ID"),
        ):
            duplicates = [
                aid for aid, count in Counter(values).items()
                if isinstance(aid, str) and aid and count > 1
            ]
            if duplicates:
                report.add("ERROR", code,
                           f"Duplicate {label} activity IDs: {sorted(duplicates)}.")

        if not statuses:
            report.add("ERROR", "UPDATE_STATUSES_MISSING",
                       "Update contains no activity status rows.")
        if not revised and statuses and any(
            str(getattr(status, "status", "")).lower() != "completed"
            for status in statuses
        ):
            report.add("ERROR", "UPDATE_NO_REVISED_ACTIVITIES", "Update contains status rows but no revised activities.")

        status_ids = {s.activity_id for s in statuses}
        if status_ids != baseline_ids:
            report.add("ERROR", "UPDATE_ACTIVITY_STATUS_SET_MISMATCH",
                       "Activity status rows do not match the baseline activity set.")

        for status in statuses:
            if not isinstance(getattr(status, "activity_id", None), str) or not status.activity_id:
                report.add("ERROR", "UPDATE_STATUS_ID_MISSING",
                           "Every update status row must have an activity ID.")
            for field_name in (
                "actual_completion_pct", "completed_quantity",
                "remaining_quantity", "original_quantity",
            ):
                if not hasattr(status, field_name) or getattr(status, field_name) is None:
                    report.add("ERROR", "UPDATE_STATUS_FIELD_MISSING",
                               f"Status row is missing '{field_name}'.",
                               getattr(status, "activity_id", None))

        revised_structure_invalid = False
        for activity in revised:
            aid = getattr(activity, "activity_id", None)
            for field_name in ("quantity", "duration_days", "es", "ef", "ls", "lf", "cost"):
                if not hasattr(activity, field_name) or getattr(activity, field_name) is None:
                    report.add("ERROR", "UPDATE_ACTIVITY_FIELD_MISSING",
                               f"Revised activity is missing '{field_name}'.", aid)
                    revised_structure_invalid = True
            if getattr(activity, "cost", None) is not None:
                for field_name in ("planned_unit_rate", "direct_cost", "total_amount"):
                    if not hasattr(activity.cost, field_name):
                        report.add("ERROR", "UPDATE_COST_FIELD_MISSING",
                                   f"Revised activity cost is missing '{field_name}'.", aid)
                        revised_structure_invalid = True
        if revised_structure_invalid:
            return report

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
                + (
                    0.0
                    if bool(getattr(cost, "wastage_in_base_rate", False))
                    else float(cost.wastage_cost or 0.0)
                )
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
        if revised:
            _run_check(
                report,
                "UPDATE_RESOURCE_PLAN_VALIDATION_EXCEPTION",
                lambda: _validate_resource_plan_and_daily_log(revised_view, report),
            )
        return report


def validate_update_result(result: Any, assumptions: Optional[dict] = None) -> ValidationReport:
        """Run update validation and report unexpected validator failures."""
        try:
            return _validate_update_result_impl(result, assumptions)
        except Exception as exc:
            report = ValidationReport()
            report.add("ERROR", "UPDATE_VALIDATION_EXCEPTION",
                       f"Update validation could not complete: {exc}")
            return report


def validate_exported_update_workbook(path: str, result: Any) -> ValidationReport:
        """Validate the update workbook structure and its validation summary."""
        report = ValidationReport()
        try:
            import openpyxl
            workbook = openpyxl.load_workbook(path, data_only=True)
        except Exception as exc:
            report.add("ERROR", "UPDATE_WORKBOOK_OPEN_FAILED",
                       f"Could not reopen update workbook: {exc}")
            return report
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
            exported_values = [
                row[0] for row in revised_sheet.iter_rows(min_row=5, values_only=True)
                if row and isinstance(row[0], str) and row[0].startswith("ACT-")
            ]
            exported = set(exported_values)
            duplicates = [
                aid for aid, count in Counter(exported_values).items() if count > 1
            ]
            if duplicates:
                report.add("ERROR", "UPDATE_WORKBOOK_DUPLICATE_ACTIVITY_ID",
                           f"Exported revised activity IDs are duplicated: {sorted(duplicates)}.")
            if exported != {a.activity_id for a in getattr(result, "revised_activities", []) or []}:
                report.add("ERROR", "UPDATE_WORKBOOK_ACTIVITY_SET_MISMATCH",
                           "Exported revised activity IDs differ from the update result.")
        return report

def format_validation_report(report: ValidationReport) -> str:
    """Return a concise human-readable report for CLI output and logs."""
    lines = [
        f"Validation: {report.status}",
        f"Errors: {len(report.errors)}",
        f"Warnings: {len(report.warnings)}",
        f"Incomplete: {len(report.incomplete)}",
    ]
    for issue in report.issues:
        suffix = f" [{issue.activity_id}]" if issue.activity_id else ""
        lines.append(f"{issue.severity} {issue.code}{suffix}: {issue.message}")
    return "\n".join(lines)
