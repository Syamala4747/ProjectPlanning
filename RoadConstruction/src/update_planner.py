"""
Update Plan Engine — revises an existing project plan against actual progress.

Given a reconstructed baseline (from the existing workbook) and the actual
cumulative progress at a status date, this module:

  1. Computes the baseline cumulative progress at the status date and the
     progress gap (baseline - actual).
  2. Estimates per-activity actual completion (from activity-level input when
     available, or by scaling baseline completions proportionally when only
     overall progress is given).
  3. Builds a *remaining-work* activity list: completed activities get zero
     duration; in-progress activities get a duration proportional to their
     remaining quantity; not-started activities keep their original duration.
  4. Re-runs CPM (same CPMEngine + ProjectCalendar) starting from the status
     date to produce the revised schedule.
  5. If the revised completion exceeds the target date, applies the same
     least-cost crashing logic as ConstraintScheduler (add crews / shifts on
     critical activities) to compress the remaining schedule.
  6. Checks target-date and target-progress feasibility and reports any
     projected delay.

All calculation logic is reused from the existing modules — no separate
scheduling engine is built.
"""
import copy
import math
from datetime import date, timedelta
from typing import List, Dict, Optional, Tuple, Any

from .models import (
    ProjectInput, ProjectScheduleResult, RoadActivity, ActivityCostBreakdown,
    ActivityReference, MonthlyProgressItem,
    UpdatePlanInput, UpdatePlanResult, ActivityActualStatus,
)
from .data_loader import DataLoader
from .cpm_engine import CPMEngine
from .cost_calculator import CostCalculator
from .duration_calculator import DurationCalculator
from .project_calendar import ProjectCalendar
from .progress_calculator import ProgressCalculator
from .dependency_manager import DependencyManager

# Shift intensities offered to the crashing search (mirrors ConstraintScheduler).
SHIFT_LEVELS = (1.0, 1.3, 1.6, 1.8)


class UpdatePlanner:
    """Compares actual progress with the baseline and replans remaining work."""

    def __init__(
        self,
        baseline_result: ProjectScheduleResult,
        update_input: UpdatePlanInput,
        data_loader: DataLoader,
    ):
        self.baseline_result = baseline_result
        self.baseline_acts = baseline_result.activities  # optimized plan
        self.base_baseline_acts = baseline_result.baseline_activities or baseline_result.activities
        self.update_input = update_input
        self.loader = data_loader
        self.assumptions = data_loader.assumptions

        self.p_in = baseline_result.project_input
        self.status_date = update_input.status_date
        if update_input.target_date is not None:
            self.target_date = update_input.target_date
        else:
            self.target_date = self.p_in.start_date + timedelta(days=baseline_result.calculated_duration_days)
        self.actual_progress = float(update_input.actual_progress_pct)

        # Calendar anchored at the status date for the remaining-work schedule.
        self.calendar = ProjectCalendar(
            self.assumptions.get("calendar"),
            self.status_date,
            terrain=self.p_in.terrain,
        )

        overhead_pct = data_loader.get_overhead_percentage(self.p_in.budget_crore)
        profit_pct = data_loader.get_profit_percentage()
        self.cost_calc = CostCalculator(
            self.assumptions, overhead_pct, profit_pct,
            enable_acceleration_cost=update_input.enable_acceleration_cost,
        )
        self.dur_calc = DurationCalculator(
            self.assumptions, self.p_in.length_km, self.calendar)

        self.activity_specs = self.assumptions.get("activity_master", [])
        self.spec_by_id = {s["id"]: s for s in self.activity_specs}
        self.spec_by_key = {s["key"]: s for s in self.activity_specs}

        # Crashing infrastructure — max work fronts for acceleration during replanning.
        # Uses the same project-type-aware logic as ConstraintScheduler.
        self.night_ok = set(self.assumptions.get("night_work_eligible", []) or [])
        ptype = str(getattr(self.p_in, "project_type", "road") or "road").lower()
        if ptype == "pipeline":
            supplied_fronts = max(1, int(getattr(self.p_in, "number_of_work_fronts", 1) or 1))
            self.corridor_fronts = min(supplied_fronts, 4)
        elif ptype == "building":
            supplied_blocks = max(1, int(getattr(self.p_in, "number_of_blocks", 1) or 1))
            self.corridor_fronts = min(supplied_blocks, 6)
        else:
            length = float(self.p_in.length_km or 0.0)
            self.corridor_fronts = max(2, min(8, int(math.ceil(length / 3.5))))
        cap = float(self.assumptions.get("crew_multiplier_cap", 6) or 6)
        self.max_multiplier = max(1, int(min(cap, self.corridor_fronts)))

        self.warnings: List[str] = []
        self.limitations: List[str] = []

    # ============================================================== public
    def plan(self) -> UpdatePlanResult:
        """Execute the full update planning pipeline."""
        status_day = (self.status_date - self.p_in.start_date).days
        available_days = max(0, (self.target_date - self.status_date).days)

        # 1. Baseline progress at the status date
        baseline_progress = self._baseline_progress_at_day(status_day)
        progress_gap = round(baseline_progress - self.actual_progress, 2)

        # 2. Target progress
        target_progress = self._determine_target_progress(status_day)
        required_progress = round(target_progress - self.actual_progress, 2)

        # 3. Per-activity actual completion
        activity_status = self._estimate_activity_completion(status_day)

        # 4. Build remaining-work activities
        remaining_acts, spent_cost_inr = self._build_remaining_activities(
            activity_status)

        # 5. Initial CPM (at original resource levels)
        revised_acts, revised_dur, revised_crit = self._run_cpm(remaining_acts)
        original_completion = self.status_date + timedelta(days=revised_dur)

        # 6. Compute initial remaining cost
        for act in revised_acts:
            if act.quantity <= 0:
                act.cost = ActivityCostBreakdown()
                continue
            esc_years = self.cost_calc.escalation.years_to(
                act.planned_start, duration_days=act.duration_days)
            act.cost = self.cost_calc.calculate_activity_cost(
                activity_key=act.activity_key,
                quantity=act.quantity,
                crews=act.crews,
                base_crews=act.base_crews,
                shifts=act.shifts,
                work_fronts=act.work_fronts,
                duration_days=act.duration_days,
                escalation_years=esc_years,
            )
        remaining_base_cost_inr = sum(a.cost.total_amount for a in revised_acts)
        remaining_base_cost_cr = round(remaining_base_cost_inr / 10000000.0, 2)
        spent_cost_cr = round(spent_cost_inr / 10000000.0, 2)
        baseline_total_cr = self.baseline_result.total_calculated_cost_crore

        # 7. Apply weights (for progress curve)
        self._apply_weights(revised_acts)

        # 8. Crashing if needed
        resources_summary: List[Dict[str, Any]] = []
        additional_accel_cost_cr = 0.0
        accel_cost_inr = 0.0
        work_fronts = 1
        for a in revised_acts:
            work_fronts = max(work_fronts, a.work_fronts)

        target_feasible = original_completion <= self.target_date
        feasibility_msg = ""

        if not target_feasible and revised_dur > available_days:
            revised_acts, revised_dur, revised_crit, accel_cost_inr, fronts, summary = \
                self._apply_crashing(revised_acts, available_days)
            additional_accel_cost_cr = round(accel_cost_inr / 10000000.0, 2)
            resources_summary = summary
            work_fronts = fronts
            target_feasible = revised_dur <= available_days

        revised_completion = self.status_date + timedelta(days=revised_dur)
        projected_delay = max(0, (revised_completion - self.target_date).days)
        target_feasible = (projected_delay == 0)

        if target_feasible:
            feasibility_msg = (
                f"Revised schedule completes on {revised_completion.isoformat()} "
                f"({revised_dur} days from status date), meeting the target date "
                f"{self.target_date.isoformat()}.")
        else:
            feasibility_msg = (
                f"Revised schedule completes on {revised_completion.isoformat()} "
                f"({revised_dur} days from status date), exceeding the target "
                f"date {self.target_date.isoformat()} by {projected_delay} days. "
                f"Further compression is not possible within available resources.")

        # 9. Revised progress at target date (projected)
        progress_at_target = self._projected_progress_at_day(
            revised_acts, revised_dur, available_days)

        # 10. Revised S-curve (monthly progress from status date)
        revised_monthly = self._revised_monthly_progress(revised_acts, revised_dur)

        # 11. Cost Reconciliation
        remaining_cost_inr = sum(a.cost.total_amount for a in revised_acts)
        remaining_cost_cr = round(remaining_cost_inr / 10000000.0, 2)
        total_cost_inr = spent_cost_inr + remaining_cost_inr
        total_cost_cr = round(spent_cost_cr + remaining_cost_cr, 2)

        # 12. Validate revised network
        errors = DependencyManager.validate_network(revised_acts)
        if errors:
            for err in errors:
                self.warnings.append(f"Network error: {err}")

        return UpdatePlanResult(
            project_input=self.p_in,
            baseline_result=self.baseline_result,
            actual_progress_pct=self.actual_progress,
            baseline_progress_at_status_date=baseline_progress,
            progress_gap_pct=progress_gap,
            status_date=self.status_date,
            target_date=self.target_date,
            target_progress_pct=target_progress,
            required_progress_pct=required_progress,
            required_calendar_days=available_days,
            activity_status=activity_status,
            revised_activities=revised_acts,
            revised_duration_days=revised_dur,
            revised_cost_crore=total_cost_cr,
            revised_critical_path=revised_crit,
            revised_monthly_progress=revised_monthly,
            original_revised_completion_date=original_completion,
            revised_completion_date=revised_completion,
            target_feasible=target_feasible,
            feasibility_message=feasibility_msg,
            projected_delay_days=projected_delay,
            projected_progress_at_target_date=progress_at_target,
            additional_acceleration_cost_crore=additional_accel_cost_cr,
            work_fronts=work_fronts,
            resources_added_summary=resources_summary,
            spent_cost_crore=spent_cost_cr,
            remaining_base_cost_crore=remaining_base_cost_cr,
            remaining_cost_crore=remaining_cost_cr,
            total_project_cost_crore=total_cost_cr,
            baseline_total_cost_crore=baseline_total_cr,
            spent_cost_inr=round(spent_cost_inr, 2),
            remaining_cost_inr=round(remaining_cost_inr, 2),
            total_project_cost_inr=round(total_cost_inr, 2),
            warnings=list(self.warnings),
            limitations=list(self.limitations),
            is_feasible=target_feasible,
        )

    # ============================================================== progress
    def _baseline_progress_at_day(self, day: int) -> float:
        """Cumulative baseline (optimized) progress at a given day ordinal."""
        total = 0.0
        for act in self.baseline_acts:
            comp = self._activity_completion_at_day(act, day)
            total += comp * act.progress_weight_pct
        return round(min(100.0, total), 2)

    @staticmethod
    def _activity_completion_at_day(act: RoadActivity, day: int) -> float:
        """Fraction (0.0-1.0) of an activity completed at the given day ordinal."""
        act_dur = max(act.duration_days, 1)
        if day <= act.es:
            return 0.0
        if day >= act.ef:
            return 1.0
        return (day - act.es) / act_dur

    def _determine_target_progress(self, status_day: int) -> float:
        """
        Determine the target cumulative progress at the target date.
        If the user provided --target-progress, use it; otherwise interpolate
        from the baseline monthly progress at the target date.
        """
        if self.update_input.target_progress_pct is not None:
            return float(self.update_input.target_progress_pct)
        target_day = (self.target_date - self.p_in.start_date).days
        return self._baseline_progress_at_day(target_day)

    # ============================================================== activity completion
    def _estimate_activity_completion(
        self, status_day: int
    ) -> List[ActivityActualStatus]:
        """
        Estimate per-activity actual completion from the overall actual progress.
        If activity-level progress was provided, use it directly.
        Otherwise scale each activity's baseline completion by the actual /
        baseline ratio and report the limitation.
        """
        activity_progress_input = self.update_input.activity_progress
        baseline_progress = self._baseline_progress_at_day(status_day)

        statuses: List[ActivityActualStatus] = []

        if activity_progress_input:
            # Direct activity-level progress provided
            for act in self.baseline_acts:
                pct = float(activity_progress_input.get(act.activity_id, 0.0))
                pct = max(0.0, min(100.0, pct))
                completed = act.quantity * pct / 100.0
                remaining = act.quantity - completed
                base_comp = self._activity_completion_at_day(act, status_day) * 100.0
                status = self._classify_status(pct)
                statuses.append(ActivityActualStatus(
                    activity_id=act.activity_id,
                    name=act.name,
                    actual_completion_pct=pct,
                    completed_quantity=completed,
                    remaining_quantity=remaining,
                    original_quantity=act.quantity,
                    status=status,
                    baseline_completion_pct=round(base_comp, 1),
                    is_delayed=pct < base_comp - 0.5,
                ))
        else:
            # When activity-level progress is not provided, allocate the overall
            # actual progress chronologically using the Earned Schedule (ES) principle:
            # find the baseline day where cumulative progress equals actual progress.
            # Completed predecessor works (e.g. site clearance, earthworks) are retired
            # as 100% completed milestones, avoiding artificial delays on finished tasks.
            max_day = float(max((act.ef for act in self.baseline_acts), default=500))
            low, high = 0.0, max_day
            for _ in range(60):
                mid = (low + high) / 2.0
                val = sum(self._activity_completion_at_day(a, mid) * a.progress_weight_pct for a in self.baseline_acts)
                if val < self.actual_progress:
                    low = mid
                else:
                    high = mid
            d_eff = (low + high) / 2.0

            self.limitations.append(
                f"Overall actual progress ({self.actual_progress:.1f}%) was allocated chronologically "
                f"using Earned Schedule (baseline day {d_eff:.1f}). Completed predecessor activities "
                "are recognized as completed milestones, and remaining work begins from the status date.")

            for act in self.baseline_acts:
                base_comp = self._activity_completion_at_day(act, status_day) * 100.0
                frac = self._activity_completion_at_day(act, d_eff)
                pct = round(frac * 100.0, 1)
                if pct >= 99.5:
                    pct = 100.0
                    completed = act.quantity
                    remaining = 0.0
                elif pct <= 0.5:
                    pct = 0.0
                    completed = 0.0
                    remaining = act.quantity
                else:
                    completed = round(act.quantity * pct / 100.0, 2)
                    remaining = round(max(0.0, act.quantity - completed), 2)

                status = self._classify_status(pct)
                statuses.append(ActivityActualStatus(
                    activity_id=act.activity_id,
                    name=act.name,
                    actual_completion_pct=pct,
                    completed_quantity=completed,
                    remaining_quantity=remaining,
                    original_quantity=act.quantity,
                    status=status,
                    baseline_completion_pct=round(base_comp, 1),
                    is_delayed=pct < base_comp - 0.5,
                ))

        return statuses

    @staticmethod
    def _classify_status(pct: float) -> str:
        if pct >= 99.5:
            return "completed"
        if pct > 0.5:
            return "in_progress"
        return "not_started"

    # ============================================================== remaining activities
    def _build_remaining_activities(
        self,
        activity_status: List[ActivityActualStatus],
    ) -> Tuple[List[RoadActivity], float]:
        """
        Build a fresh activity list for the remaining-work CPM. Completed
        activities get zero effort; in-progress activities get effort proportional
        to their remaining quantity; not-started activities keep original effort.

        Returns (activities, spent_cost_inr).
        """
        status_map = {s.activity_id: s for s in activity_status}
        spent_cost_inr = 0.0
        activities: List[RoadActivity] = []

        for orig in self.baseline_acts:
            st = status_map.get(orig.activity_id)
            if st is None:
                st = ActivityActualStatus(
                    activity_id=orig.activity_id, name=orig.name,
                    original_quantity=orig.quantity)

            remaining_qty = max(0.0, st.remaining_quantity)

            # Compute remaining effort days
            if remaining_qty <= 0:
                remaining_effort = 0
            else:
                remaining_effort, _, _, _ = self.dur_calc.calculate_duration(
                    activity_key=orig.activity_key,
                    quantity=remaining_qty,
                    crews=orig.crews,
                    shifts=orig.shifts,
                )

            # Spent cost (completed portion at original unit rate)
            if orig.quantity > 0:
                orig_unit_cost = orig.cost.total_amount / orig.quantity
            else:
                orig_unit_cost = 0.0
            spent_cost_inr += st.completed_quantity * orig_unit_cost

            # Clone the activity with remaining data
            act = copy.deepcopy(orig)
            act.quantity = round(remaining_qty, 2)
            act.effort_days = remaining_effort
            act.duration_days = remaining_effort  # CPM will overwrite with calendar days
            # Reset CPM fields — will be recomputed
            act.es = 0
            act.ef = 0
            act.ls = 0
            act.lf = 0
            act.total_float = 0
            act.free_float = 0
            act.is_critical = False
            act.planned_start = None
            act.planned_finish = None
            act.predecessor_lags = dict(orig.predecessor_lags)
            activities.append(act)

        return activities, spent_cost_inr

    # ============================================================== CPM
    def _run_cpm(
        self, activities: List[RoadActivity]
    ) -> Tuple[List[RoadActivity], int, List[str]]:
        """
        Run CPM on the remaining-work activities starting from the status date.
        Returns (activities, duration_days, critical_path).
        """
        # Create a ProjectInput with start_date = status_date
        rev_input = ProjectInput(
            title=self.p_in.title,
            length_km=self.p_in.length_km,
            duration_days=self.p_in.duration_days,
            budget_crore=self.p_in.budget_crore,
            start_date=self.status_date,
            lanes=self.p_in.lanes,
            terrain=self.p_in.terrain,
        )
        cpm = CPMEngine(
            activities=activities,
            project_input=rev_input,
            calendar=self.calendar,
        )
        dur, crit = cpm.run_cpm()

        # Completed activities (quantity <= 0 or duration_days == 0) are completed milestones
        for a in activities:
            if a.quantity <= 0 or a.duration_days == 0:
                a.is_critical = False
                a.duration_days = 0
                a.planned_start = self.status_date
                a.planned_finish = self.status_date

        # The remaining critical path should only trace activities with duration > 0
        crit_remaining = [
            aid for aid in crit
            if next((a.duration_days for a in activities if a.activity_id == aid), 0) > 0
        ]
        return activities, dur, crit_remaining

    # ============================================================== weights
    def _apply_weights(self, activities: List[RoadActivity]) -> None:
        """Apply progress weights based on remaining work content."""
        work = []
        for a in activities:
            if a.quantity > 0 and a.productivity_rate > 0:
                w = a.quantity / a.productivity_rate
            else:
                w = max(1.0, a.effort_days) if a.effort_days > 0 else 0.0
            work.append(w if w > 0 else 0.0)

        tot_work = sum(work)
        if tot_work > 0:
            for a, w in zip(activities, work):
                a.progress_weight_pct = round((w / tot_work) * 100.0, 2)
        else:
            for a in activities:
                a.progress_weight_pct = 0.0

        # Fix rounding to sum to 100
        if activities and tot_work > 0:
            diff = round(100.0 - sum(a.progress_weight_pct for a in activities), 2)
            activities[-1].progress_weight_pct = round(
                activities[-1].progress_weight_pct + diff, 2)

    # ============================================================== crashing
    def _build_crash_levels(self, act: RoadActivity) -> List[Tuple[int, float]]:
        """
        Build the ordered list of (crews, shifts) levels above the activity's
        current resource level, mirroring ConstraintScheduler._build_crash_levels
        but starting from the current (possibly already crashed) organisation.
        """
        spec = self.spec_by_key.get(act.activity_key)
        base_crews = int(spec.get("base_crews", 1)) if spec else 1
        key = act.activity_key
        eligible = key in self.night_ok if self.night_ok else True

        levels: List[Tuple[int, float]] = []
        for mult in range(1, self.max_multiplier + 1):
            crews = base_crews * mult
            levels.append((crews, 1.0))
            if eligible and mult > 1:
                for s in SHIFT_LEVELS[1:]:
                    levels.append((crews, s))

        # Sort by capacity and deduplicate
        levels.sort(key=lambda c: c[0] * max(0.5, c[1]))
        unique: List[Tuple[int, float]] = []
        for c in levels:
            cap = c[0] * max(0.5, c[1])
            if not unique or cap > unique[-1][0] * max(0.5, unique[-1][1]) + 1e-9:
                unique.append(c)

        # Find current level and return only levels above it
        current_cap = act.crews * max(0.5, act.shifts)
        above = [c for c in unique
                 if c[0] * max(0.5, c[1]) > current_cap + 1e-9]
        return above

    def _apply_crashing(
        self,
        activities: List[RoadActivity],
        available_days: int,
    ) -> Tuple[List[RoadActivity], int, List[str], float, int, List[Dict[str, Any]]]:
        """
        Least-cost crashing loop on the revised schedule, mirroring
        ConstraintSolver's approach: try advancing each critical activity by one
        resource level, pick the cheapest cost-per-day-saved, repeat.

        Returns (activities, duration, critical_path, accel_cost_inr,
        work_fronts, resources_summary).
        """
        max_budget = self.p_in.budget_crore
        spent_cr = sum(
            (a.quantity if a.quantity > 0 else 0) * 0  # spent already accounted
            for a in activities)  # placeholder; actual spent tracked outside
        remaining_budget = max_budget  # simplified: use full budget as ceiling

        current_acts = activities
        current_dur = 0
        current_crit: List[str] = []
        # Initial evaluation
        current_acts, current_dur, current_crit = self._run_cpm(current_acts)
        self._recost(current_acts)
        current_cost_inr = sum(a.cost.total_amount for a in current_acts)
        initial_cost_inr = current_cost_inr

        resources_summary: List[Dict[str, Any]] = []
        visited = set()
        max_iters = 200

        for _ in range(max_iters):
            if current_dur <= available_days:
                break

            critical_ids = [a.activity_id for a in current_acts if a.is_critical]
            candidates: List[Dict[str, Any]] = []

            for aid in critical_ids:
                act = next((a for a in current_acts if a.activity_id == aid), None)
                if act is None or act.quantity <= 0:
                    continue
                above = self._build_crash_levels(act)
                if not above:
                    continue
                next_level = above[0]
                # Trial: build a copy with the new resource level
                trial_acts = copy.deepcopy(current_acts)
                t_act = next(a for a in trial_acts if a.activity_id == aid)
                t_act.crews = next_level[0]
                t_act.shifts = next_level[1]
                t_act.work_fronts = max(1, t_act.crews // t_act.base_crews)
                # Recompute effort days for remaining quantity
                eff, _, _, _ = self.dur_calc.calculate_duration(
                    activity_key=t_act.activity_key,
                    quantity=t_act.quantity,
                    crews=t_act.crews,
                    shifts=t_act.shifts,
                )
                t_act.effort_days = eff
                t_act.duration_days = eff

                sig = (aid, next_level)
                if sig in visited:
                    continue
                visited.add(sig)

                trial_acts, t_dur, t_crit = self._run_cpm(trial_acts)
                self._recost(trial_acts)
                t_cost = sum(a.cost.total_amount for a in trial_acts)

                saved = current_dur - t_dur
                if saved <= 0:
                    continue
                if t_cost / 10000000.0 > remaining_budget:
                    continue

                ratio = (t_cost - current_cost_inr) / saved
                candidates.append({
                    "act_id": aid,
                    "act_name": act.name,
                    "level": next_level,
                    "trial_acts": trial_acts,
                    "t_dur": t_dur,
                    "t_cost": t_cost,
                    "t_crit": t_crit,
                    "ratio": ratio,
                    "saved": saved,
                })

            if not candidates:
                break

            best = min(candidates, key=lambda c: c["ratio"])
            old_act = next(a for a in current_acts
                           if a.activity_id == best["act_id"])
            extra_g = best["level"][0] - old_act.crews
            resources_summary.append({
                "activity_id": best["act_id"],
                "activity_name": best["act_name"],
                "baseline_crews": old_act.crews,
                "optimized_crews": best["level"][0],
                "work_fronts": max(1, best["level"][0] // old_act.base_crews),
                "shifts": best["level"][1],
                "days_saved": best["saved"],
                "cost_increase_lakhs": round(
                    (best["t_cost"] - current_cost_inr) / 100000.0, 2),
            })
            current_acts = best["trial_acts"]
            current_dur = best["t_dur"]
            current_crit = best["t_crit"]
            current_cost_inr = best["t_cost"]

        accel_cost_inr = max(0.0, current_cost_inr - initial_cost_inr)
        fronts = 1
        for a in current_acts:
            fronts = max(fronts, a.work_fronts)

        return (current_acts, current_dur, current_crit,
                accel_cost_inr, fronts, resources_summary)

    def _recost(self, activities: List[RoadActivity]) -> None:
        """Recalculate cost for each activity at its current resource level."""
        for act in activities:
            if act.quantity <= 0:
                act.cost = ActivityCostBreakdown()
                continue
            esc_years = self.cost_calc.escalation.years_to(
                act.planned_start, duration_days=act.duration_days)
            act.cost = self.cost_calc.calculate_activity_cost(
                activity_key=act.activity_key,
                quantity=act.quantity,
                crews=act.crews,
                base_crews=act.base_crews,
                shifts=act.shifts,
                work_fronts=act.work_fronts,
                duration_days=act.duration_days,
                escalation_years=esc_years,
            )

    # ============================================================== S-curve
    def _revised_monthly_progress(
        self, activities: List[RoadActivity], duration_days: int
    ) -> List[MonthlyProgressItem]:
        """Generate the revised forecast monthly progress from the status date."""
        rev_input = ProjectInput(
            title=self.p_in.title,
            length_km=self.p_in.length_km,
            duration_days=self.p_in.duration_days,
            budget_crore=self.p_in.budget_crore,
            start_date=self.status_date,
            lanes=self.p_in.lanes,
            terrain=self.p_in.terrain,
        )
        calc = ProgressCalculator(
            activities=activities,
            project_input=rev_input,
            calculated_duration_days=duration_days,
        )
        monthly, _ = calc.generate_monthly_schedule()
        return monthly

    def _projected_progress_at_day(
        self, activities: List[RoadActivity], revised_dur: int, day: int
    ) -> float:
        """Projected cumulative progress (remaining work) at the given day from
        the status date."""
        total = 0.0
        for act in activities:
            comp = self._activity_completion_at_day(act, day)
            total += comp * act.progress_weight_pct
        return round(min(100.0, total), 2)
