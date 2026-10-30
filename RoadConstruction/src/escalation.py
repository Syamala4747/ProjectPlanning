"""
Price escalation.

Every rate in `assumptions.yaml` is priced at a single base date. A build that
runs for years must not be priced flat, or the estimate systematically
understates the outturn.

The method here is deliberately simple and auditable:

    years       = days from the base date to the MID-POINT of this activity's
                  own spend, in years
    factor_c    = (1 + annual_rate_c) ** years        for each cost component c
    blended     = SUM( share_c x factor_c )            weighted by cost shares
    escalation  = base_cost x (blended - 1)

Escalating to the MID-POINT of each activity's own spend - rather than to the
project end date - matters: an activity finished in month 3 should not carry
two years of inflation.

The cap mirrors how works contracts are actually written: escalation exposure
is bounded, so an uncapped index over a long monsoon-delayed programme would
overstate the liability.
"""
from __future__ import annotations

from datetime import date
from typing import Any, Dict, Optional

DAYS_PER_YEAR = 365.25


class EscalationModel:
    """Reads the `escalation` block and prices an activity's spend forward."""

    def __init__(self, assumptions: Dict[str, Any]):
        cfg = (assumptions or {}).get("escalation") or {}

        self.enabled = bool(cfg.get("enabled", False))
        self.base_date = self._as_date(cfg.get("base_date")) or date.today()

        rates = cfg.get("annual_rates") or {}
        self.annual_rates = {
            "material": float(rates.get("material", 0.0) or 0.0),
            "labour":   float(rates.get("labour",   0.0) or 0.0),
            "plant":    float(rates.get("plant",    0.0) or 0.0),
            "fuel":     float(rates.get("fuel",     0.0) or 0.0),
        }

        # A component with no rate of its own (e.g. an overhead recovery line)
        # escalates at the material rate rather than at zero, so nothing silently
        # becomes free of inflation.
        self.default_rate = float(cfg.get("default_annual_rate", 0.0) or 0.0)
        if not any(self.annual_rates.values()):
            self.default_rate = self.annual_rate("material")

        overrides = cfg.get("index_overrides") or {}
        self.overrides: Dict[str, float] = {}
        for name, ov in overrides.items():
            if isinstance(ov, dict) and ov.get("annual") is not None:
                self.overrides[name] = float(ov["annual"])
            elif isinstance(ov, (int, float)):
                self.overrides[name] = float(ov)

        self.cap_pct = float(cfg.get("cap_pct_of_base", 1.0) or 1.0)
        self.cap_factor = 1.0 + self.cap_pct

        method = str(cfg.get("method", "spend_curve_midpoint"))
        if method not in ("spend_curve_midpoint", "linear_by_calendar"):
            raise ValueError(
                f"escalation.method must be 'spend_curve_midpoint' or "
                f"'linear_by_calendar' (got {method!r}). Fix it in assumptions.yaml."
            )
        self.method = method

        self.excluded = set(cfg.get("excluded") or [])

    # ------------------------------------------------------------------ utils
    @staticmethod
    def _as_date(v: Any) -> Optional[date]:
        if isinstance(v, date):
            return v
        if isinstance(v, str):
            try:
                return date.fromisoformat(v.strip()[:10])
            except ValueError:
                return None
        return None

    def annual_rate(self, component: str) -> float:
        return float(self.annual_rates.get(component, self.default_rate))

    def material_rate(self, material_key: str) -> float:
        """
        Escalation rate for ONE recipe material.

        `index_overrides` are keyed by index name ("bitumen", "steel") and are
        matched against the recipe key ("bitumen_kg", "tmt_steel_fe500d_mt"),
        so an activity carrying bitumen escalates that bitumen at its own index
        while its aggregates still escalate at the headline material rate.
        Without this the override list would be parsed and then ignored.
        """
        key = str(material_key or "").lower()
        for name, rate in self.overrides.items():
            if name.lower() in key:
                return float(rate)
        return self.annual_rate("material")

    # ------------------------------------------------------------------ maths
    def years_to(self, spend_date: Optional[date], duration_days: float = 0.0) -> float:
        """
        Years from the base date to the point the activity's spend is priced at.

        `spend_curve_midpoint` uses the MID-POINT of the activity (start + half
        its duration), which is the fair representation of when the money moves.
        `linear_by_calendar` uses the FINISH date, which is the conservative
        reading and matches a contract that escalates on completion.

        Nothing spent before the base date is discounted - the year count is
        floored at zero.
        """
        if not self.enabled or spend_date is None:
            return 0.0
        if self.method == "linear_by_calendar":
            offset = float(duration_days or 0.0)
        else:
            offset = float(duration_days or 0.0) / 2.0
        days = (spend_date - self.base_date).days + offset
        return max(0.0, days / DAYS_PER_YEAR)

    def blended_factor(self, years: float, shares: Dict[str, float],
                       activity_key: str = "",
                       material_lines: Optional[list] = None) -> float:
        """Cost-share weighted escalation factor for one activity."""
        if not self.enabled or years <= 0:
            return 1.0
        if activity_key and activity_key in self.excluded:
            return 1.0
        total_share = sum(float(v) for v in (shares or {}).values())
        if total_share <= 0:
            # Nothing to weight - escalate nothing rather than return a
            # meaningless factor of zero.
            return 1.0
        factor = 0.0
        for comp, share in (shares or {}).items():
            w = float(share) / total_share
            if comp == "material" and material_lines and self.overrides:
                # Escalate the ACTUAL material mix of this activity, so an
                # overridden material uses its own index and the rest use the
                # headline material rate.
                total_cost = sum(float(ln.get("cost", 0.0) or 0.0)
                                 for ln in material_lines) or 1.0
                sub = 0.0
                for ln in material_lines:
                    mc = float(ln.get("cost", 0.0) or 0.0) / total_cost
                    sub += mc * ((1.0 + self.material_rate(ln.get("material", ""))) ** years)
                factor += w * sub
            else:
                factor += w * ((1.0 + self.annual_rate(comp)) ** years)
        return min(factor, self.cap_factor)

    def apply(self, base_cost: float, years: float, shares: Dict[str, float],
              activity_key: str = "", material_lines: Optional[list] = None) -> Dict[str, float]:
        """Return {factor, years, amount} for `base_cost`."""
        factor = self.blended_factor(years, shares, activity_key, material_lines)
        amount = max(0.0, float(base_cost) * (factor - 1.0))
        return {"factor": factor, "years": years, "amount": amount}

    def describe(self) -> str:
        if not self.enabled:
            return "Price escalation: DISABLED"
        r = self.annual_rates
        return (f"Escalation      : base {self.base_date.isoformat()}, "
                f"material {r['material']:.1%} / labour {r['labour']:.1%} / "
                f"plant {r['plant']:.1%} / fuel {r['fuel']:.1%} per year, "
                f"{self.method}, cap {self.cap_pct:.0%}")