"""
Duration Calculation Engine for Road Construction Activities.

Enforces:
    effort_days = ceil(Quantity / (Daily_Productivity_per_crew * Crews * Effective_Shifts))

`effort_days` is PRODUCTION days, not calendar days. The project calendar
(ProjectCalendar) later stretches them into calendar days by skipping the
weekly rest day, public holidays and the monsoon slowdown, because a productive
day and a calendar day are not the same thing.
"""
import math
from typing import Dict, Any, Tuple, Optional

from .project_calendar import ProjectCalendar


class DurationCalculator:
    """
    Computes activity effort using the standard productivity formula.
    """

    def __init__(self, assumptions: Dict[str, Any], project_length_km: float,
                 calendar: Optional[ProjectCalendar] = None):
        self.assumptions = assumptions
        self.project_length_km = project_length_km
        self.calendar = calendar

        # Build lookup table from activity_master in assumptions.yaml
        self.prod_map: Dict[str, Tuple[float, str, int]] = {}
        for item in assumptions.get("activity_master", []):
            key = item.get("key")
            rate = float(item.get("productivity_rate", 100.0))
            unit = str(item.get("productivity_unit", "units/day"))
            base_crews = int(item.get("base_crews", 1))
            self.prod_map[key] = (rate, unit, base_crews)

    def effective_shifts(self, shifts: float) -> float:
        """
        Night work does not deliver full output. `shifts` > 1 means part of the
        day is a night shift, which counts for night_shift_productivity_factor
        rather than a full day (assumptions.yaml `calendar` block).
        """
        if self.calendar is None:
            return float(shifts or 1.0)
        return self.calendar.effective_shifts(shifts)

    def calculate_duration(
        self,
        activity_key: str,
        quantity: float,
        crews: Optional[int] = None,
        shifts: float = 1.0
    ) -> Tuple[int, float, int, str]:
        """
        Returns (effort_days, prod_rate, crews, unit_str).

        effort_days = ceil(quantity / (prod_rate * crews * effective_shifts))
        """
        prod_rate, unit_str, default_base_crews = self.prod_map.get(
            activity_key,
            (100.0, "units/day", 2 if self.project_length_km >= 15 else 1)
        )

        if crews is None or crews <= 0:
            crews = default_base_crews

        # A physical activity with no measured quantity is not work. Do not
        # allocate a minimum day to it: that would create artificial calendar
        # time and labour/plant/fuel idle losses for a zero-scope activity.
        if quantity <= 0:
            return 0, prod_rate, crews, unit_str

        eff_shifts = max(0.5, self.effective_shifts(shifts))
        daily_capacity = prod_rate * crews * eff_shifts
        effort_days = max(1, math.ceil(quantity / max(daily_capacity, 0.001)))

        return effort_days, prod_rate, crews, unit_str
