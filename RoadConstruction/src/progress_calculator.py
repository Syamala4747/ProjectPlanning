"""
Monthly and Periodic Progress Calculation Engine for Road Construction Projects.
Maps CPM early start/finish dates into calendar periods.

Two different curves are produced:

* PHYSICAL PROGRESS  - weighted by `progress_weight_pct` (crew-days of real work
  content), so it measures how much of the road has actually been built.
* EXPENDITURE        - weighted by each activity's own cost, so it measures money
  leaving the bank.

They are weighted independently on purpose. Progress built on cost weights is
identical to the expenditure curve by construction, which makes the S-curve
useless as a status tool: a plan can be 100% "on schedule" purely because it
spent 100% of the money.
"""
from datetime import date, timedelta
from typing import List, Dict, Any, Tuple
import math
from .models import RoadActivity, MonthlyProgressItem, ProjectInput


class ProgressCalculator:
    def __init__(self, activities: List[RoadActivity], project_input: ProjectInput,
                 calculated_duration_days: int):
        self.activities = activities
        self.project_input = project_input
        self.duration_days = max(calculated_duration_days, 1)

    # ------------------------------------------------------------------ periods
    def _periods(self) -> List[Tuple[int, int, date]]:
        """
        Real calendar periods within the plan, as [start_ordinal, end_ordinal)
        pairs plus the period's start date. The first period runs from the
        project start to the end of that month, so a mid-month start produces a
        short opening month rather than pretending every month is 30 days.
        """
        start = self.project_input.start_date
        periods: List[Tuple[int, int, date]] = []
        cur = start
        while True:
            o0 = (cur - start).days
            if o0 >= self.duration_days:
                break
            nxt = date(cur.year + (cur.month == 12), (cur.month % 12) + 1, 1)
            o1 = min((nxt - start).days, self.duration_days)
            periods.append((o0, o1, cur))
            cur = nxt
        return periods or [(0, self.duration_days, start)]

    def generate_monthly_schedule(self) -> Tuple[List[MonthlyProgressItem], Dict[str, List[float]]]:
        """
        Generates:
        1. List of MonthlyProgressItem with periodic and cumulative progress %.
        2. Matrix of activity_id -> list of progress % in each period.
        """
        periods = self._periods()
        num_months = len(periods)

        activity_monthly_matrix: Dict[str, List[float]] = {
            a.activity_id: [0.0] * num_months for a in self.activities}
        monthly_items: List[MonthlyProgressItem] = []

        # Per-activity percentage of the plan falling in each period.
        for act in self.activities:
            act_dur = max(act.duration_days, 1)
            for m, (p_start, p_end, _) in enumerate(periods):
                overlap_days = max(0, min(act.ef, p_end) - max(act.es, p_start))
                month_act_pct = (overlap_days / act_dur) * 100.0
                activity_monthly_matrix[act.activity_id][m] = round(month_act_pct, 1)
            if (
                str(getattr(self.project_input, "project_type", "road") or "road").lower()
                == "pipeline"
            ):
                # Preserve the displayed one-decimal precision while making
                # each pipeline activity reconcile exactly to 100%.
                rounded_total = sum(activity_monthly_matrix[act.activity_id])
                if activity_monthly_matrix[act.activity_id]:
                    last_index = max(
                        index for index, value in enumerate(
                            activity_monthly_matrix[act.activity_id]
                        ) if value > 0
                    )
                    last = activity_monthly_matrix[act.activity_id][last_index]
                    activity_monthly_matrix[act.activity_id][last_index] = round(
                        last + (100.0 - rounded_total), 1
                    )

        cum_progress = 0.0
        cum_expenditure = 0.0
        total_project_cost = sum(a.cost.total_amount for a in self.activities)

        for m, (p_start, p_end, p_date) in enumerate(periods):
            period_progress = 0.0
            period_cost_inr = 0.0

            for act in self.activities:
                act_month_pct = activity_monthly_matrix[act.activity_id][m]
                frac = act_month_pct / 100.0
                period_progress += frac * act.progress_weight_pct
                period_cost_inr += frac * act.cost.total_amount

            cum_progress += period_progress
            period_cost_cr = period_cost_inr / 10000000.0
            cum_expenditure += period_cost_cr

            monthly_items.append(MonthlyProgressItem(
                month_index=m + 1,
                month_name=f"Month {m + 1} ({p_date.strftime('%b %Y')})",
                period_start_day=p_start + 1,
                period_end_day=p_end,
                period_planned_progress_pct=round(period_progress, 2),
                cumulative_planned_progress_pct=round(min(100.0, cum_progress), 2),
                period_expenditure_crore=round(period_cost_cr, 2),
                cumulative_expenditure_crore=round(cum_expenditure, 2)
            ))

        if monthly_items:
            monthly_items[-1].cumulative_planned_progress_pct = 100.0
            # The workbook displays expenditure to two decimals. Reconcile
            # that displayed series to the activity total so every project
            # type has an exact S-curve cost identity after export.
            displayed_total = sum(
                item.period_expenditure_crore for item in monthly_items
            )
            monthly_items[-1].period_expenditure_crore = round(
                monthly_items[-1].period_expenditure_crore
                + (total_project_cost / 10000000.0 - displayed_total),
                2,
            )
            monthly_items[-1].cumulative_expenditure_crore = round(
                total_project_cost / 10000000.0, 2
            )

        return monthly_items, activity_monthly_matrix
