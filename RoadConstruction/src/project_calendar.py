"""
Project calendar: converts productivity EFFORT days into real CALENDAR days.

The planner used to treat "1 day of production" and "1 day of the calendar" as
the same thing, so a 6-day working week, public holidays and the monsoon season
had no effect on any date. This module owns that conversion and is driven
entirely by the `calendar` block of assumptions.yaml:

    working_days_per_week      6   -> Mon-Sat working, Sunday rest
    monsoon_months             [6,7,8,9]
    monsoon_productivity_factor 0.65 -> 65% output while it rains
    holidays_per_year          14
    night_shift_productivity_factor 0.8
    night_shift_cost_multiplier    1.5

`holidays` may be given explicitly in the config as a list of "YYYY-MM-DD" or
"MM-DD" strings; those are added on top of a fixed set of Indian national
holidays. If `holidays_per_year` asks for more than either provides, the
remainder are spread evenly through the year as stand-ins for the movable
festivals (Holi/Diwali/Eid) that cannot be fixed to a date in advance.
"""
from datetime import date, timedelta
from typing import Any, Dict, Iterable, List, Optional, Set

# Fixed-date public holidays observed every year.
_FIXED_HOLIDAYS = (
    (1, 1),    # New Year
    (1, 26),   # Republic Day
    (4, 14),   # Ambedkar Jayanti
    (5, 1),    # Labour Day
    (8, 15),   # Independence Day
    (10, 2),   # Gandhi Jayanti
    (12, 25),  # Christmas
)

# Safety valve so a mis-configured factor can never spin forever.
_MAX_DAYS = 100_000


class ProjectCalendar:
    def __init__(self, calendar_cfg: Optional[Dict[str, Any]] = None,
                 project_start: Optional[date] = None,
                 terrain: str = "plain"):
        cfg = dict(calendar_cfg or {})

        wpw = int(cfg.get("working_days_per_week", 7) or 7)
        self.working_days_per_week = max(1, min(7, wpw))
        # Rest days are the tail of the ISO week, so 6 -> Sunday, 5 -> Sat+Sun.
        self.rest_weekdays: Set[int] = set(range(self.working_days_per_week, 7))

        self.monsoon_months: Set[int] = {int(m) for m in (cfg.get("monsoon_months") or [])}
        self.monsoon_factor = float(cfg.get("monsoon_productivity_factor", 1.0) or 1.0)
        self.terrain = str(terrain or "plain").strip().lower()
        self.terrain_factor = (
            float(cfg.get("hilly_productivity_factor", 1.0) or 1.0)
            if self.terrain == "hilly" else 1.0
        )
        self.holidays_per_year = int(cfg.get("holidays_per_year", 0) or 0)
        self.night_shift_productivity_factor = float(
            cfg.get("night_shift_productivity_factor", 1.0) or 1.0)
        self.night_shift_cost_multiplier = float(
            cfg.get("night_shift_cost_multiplier", 1.0) or 1.0)

        explicit, monthdays = self._parse_explicit(cfg.get("holidays"))
        anchor = project_start or date.today()
        self.holidays = self._build_holidays(explicit, monthdays,
                                             anchor.year - 1, anchor.year + 2)

    # ------------------------------------------------------------------ setup
    @staticmethod
    def _parse_explicit(raw: Any):
        """Accepts "YYYY-MM-DD" (year-specific) and "MM-DD" (recurring)."""
        explicit: Dict[int, Set[int]] = {}
        monthdays: Set[int] = set()
        for item in raw or []:
            s = str(item).strip()
            try:
                if len(s) == 10:  # YYYY-MM-DD
                    d = date(int(s[0:4]), int(s[5:7]), int(s[8:10]))
                    explicit.setdefault(d.year, set()).add(d.toordinal())
                else:             # MM-DD
                    m, dd = s.split("-")
                    monthdays.add(int(m) * 100 + int(dd))
            except (ValueError, TypeError):
                continue
        return explicit, monthdays

    def _build_holidays(self, explicit: Dict[int, Set[int]],
                        monthdays: Set[int], y0: int, y1: int) -> Dict[int, Set[int]]:
        """
        Fixed national holidays are always observed, explicit `holidays` entries
        are added on top of them, and `holidays_per_year` is then topped up with
        evenly spaced stand-ins so the configured count is always reached.
        """
        holidays: Dict[int, Set[int]] = {}
        for y in range(y0, y1 + 1):
            days: Set[date] = {date(y, m, d) for m, d in _FIXED_HOLIDAYS}
            for ordv in explicit.get(y, ()):
                days.add(date.fromordinal(ordv))
            for md in monthdays:
                try:
                    days.add(date(y, md // 100, md % 100))
                except ValueError:
                    pass
            need = self.holidays_per_year - len(days)
            if need > 0:
                year_days = [date(y, 1, 1) + timedelta(days=i) for i in range(365)]
                open_days = [d for d in year_days if d not in days]
                stride = max(1, len(open_days) // need)
                taken = 0
                for i in range(0, len(open_days), stride):
                    if taken >= need:
                        break
                    days.add(open_days[i])
                    taken += 1
            holidays[y] = {d.toordinal() for d in days}
        return holidays

    # ----------------------------------------------------------- day queries
    def is_working(self, d: date) -> bool:
        if d.weekday() in self.rest_weekdays:
            return False
        return d.toordinal() not in self.holidays.get(d.year, ())

    def productivity_factor(self, d: date) -> float:
        """Fraction of nominal output this calendar day delivers."""
        if d.month in self.monsoon_months:
            return self.monsoon_factor * self.terrain_factor
        return self.terrain_factor

    def effective_shifts(self, shifts: float) -> float:
        """
        Apply night_shift_productivity_factor to any portion of the day that is
        night work. A full day shift is unaffected; an extra 0.8 of a night
        shift delivers only 0.8 x 0.8 of nominal output.
        """
        s = float(shifts or 1.0)
        if s <= 1.0:
            return s
        f = self.night_shift_productivity_factor
        return 1.0 + (s - 1.0) * f

    # -------------------------------------------------------- duration bridge
    def calendar_days_needed(self, start: date, effort_days: float) -> int:
        """
        Calendar days from `start` required to deliver `effort_days` of
        production. Rest days and holidays deliver nothing; monsoon days
        deliver `monsoon_factor` of a normal day.
        """
        effort = float(effort_days)
        if effort <= 0:
            return 0

        consumed = 0.0
        elapsed = 0
        d = start
        while consumed < effort and elapsed < _MAX_DAYS:
            elapsed += 1
            if self.is_working(d):
                consumed += self.productivity_factor(d)
            d += timedelta(days=1)
        return max(1, elapsed)

    # ---------------------------------------------------------------- report
    def describe(self) -> List[str]:
        """Human-readable summary of which calendar rules are in force."""
        wk = f"{self.working_days_per_week}-day week (rest: " + (
            ", ".join(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][i]
                      for i in sorted(self.rest_weekdays)) or "none") + ")"
        lines = [f"Working pattern : {wk}"]
        if self.terrain == "hilly":
            lines.append(f"Terrain penalty : {self.terrain_factor:.0%} productivity")
        if self.monsoon_months:
            lines.append(
                f"Monsoon penalty : months {sorted(self.monsoon_months)} at "
                f"{self.monsoon_factor:.0%} productivity")
        lines.append(
            f"Holidays        : {self.holidays_per_year} per year "
            f"({len(self.holidays.get(date.today().year, ()))} loaded for {date.today().year})")
        if self.night_shift_productivity_factor != 1.0:
            lines.append(
                f"Night work      : {self.night_shift_productivity_factor:.0%} productivity, "
                f"x{self.night_shift_cost_multiplier:g} cost")
        return lines
