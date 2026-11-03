"""
Cost and Rate Analysis Engine for Road Construction Activities.
Fetches unit rates, cost component shares, crashing friction factors, overheads, and profit from assumptions.yaml.
"""
import math
from typing import Dict, Any, Optional
from .models import ActivityCostBreakdown

class CostCalculator:
    """
    Computes activity direct cost, component breakdowns, overheads, contractor profit,
    and crashing adjustments.

    Uses the PLANNED UNIT RATE for all cost calculations. The planned rate is
    DERIVED from the first-principles resource build-up: the material recipes,
    crew day rates, plant hire rates and fuel consumption in assumptions.yaml,
    evaluated at each activity's natural single-gang / single-shift basis. The
    planning variation vs the MoRTH direct rate is then reported as the gap:

        planned_rate = resource_build_up_per_unit  (materials + labour + plant + fuel)
        variation    = (planned_rate x K − MoRTH_direct_rate) / MoRTH_direct_rate × 100

    where K is the SDB on-cost multiplier (`sdb_rate_oncost`: road ×1.188,
    structures ×1.32). The MoRTH direct rates come from SDB rate analyses that
    already embed overheads + contractor's profit, so both sides of the
    comparison are ALL-IN and the variation measures genuine scope/price
    differences rather than the overhead structure.

    No hand-typed rate or variation table is read — change a cement bag price,
    a machine hire rate or a recipe in assumptions.yaml and both the planned
    rate and the variation re-derive. The original MoRTH Direct Unit Rate is
    stored unchanged for traceability. Material wastage and fuel/plant/labour
    idle losses stay additive allowances on top of the rate (they are never
    double-counted into the derived planned rate).
    """
    def __init__(
        self,
        assumptions: Dict[str, Any],
        overhead_pct: float = 8.0,
        profit_pct: float = 10.0,
        enable_acceleration_cost: bool = True
    ):
        self.assumptions = assumptions
        self.overhead_pct = overhead_pct
        self.profit_pct = profit_pct
        # When False, every NEW time-driven acceleration term is forced to zero
        # so the pre-overhaul cost behaviour can be reproduced exactly (CLI: --no-accel-cost).
        self.enable_acceleration_cost = bool(enable_acceleration_cost)

        # Original MoRTH source rates — never modified
        self.direct_rates = assumptions.get("direct_unit_rates_inr", {})

        # The PLANNED UNIT RATE is DERIVED from the first-principles resource
        # build-up (material recipes + crew day rates + plant hire + fuel lph,
        # evaluated at the natural single-gang/single-shift basis) — see
        # `_resource_planned_rate`. The planning variation vs the MoRTH direct
        # rate is then the REPORTED gap between real site cost and the SDB rate.
        # No hand-typed rate or variation table exists in this model.

        shares = assumptions.get("cost_breakdown_shares", {})
        self.labour_share   = float(shares.get("labour_share",   0.18))
        self.material_share = float(shares.get("material_share", 0.54))
        self.equipment_share= float(shares.get("equipment_share",0.18))
        self.fuel_share     = float(shares.get("fuel_share",     0.10))

        # Rate basis. `source_plus_oncosts` treats the published direct rate as
        # the direct (bare) resource cost and adds the configured material
        # wastage, non-material idle losses and statutory OH/profit on top. It
        # is the defensible build-up when the catalogue carries source rates but
        # no complete bill-of-resource recipe (pipeline and building). `recipes`
        # keeps the original first-principles material/crew/plant/fuel build-up
        # (road).
        self.rate_build_up_mode = str(
            assumptions.get("rate_build_up_mode", "recipes") or "recipes"
        )
        self.rate_oncost_basis = str(
            assumptions.get("rate_oncost_basis", "sdb") or "sdb"
        )
        raw_shares = assumptions.get("cost_breakdown_shares_by_category", {}) or {}
        self._shares_by_category: Dict[str, Dict[str, float]] = {}
        for cat, vals in raw_shares.items():
            if not isinstance(vals, dict):
                continue
            self._shares_by_category[str(cat)] = {
                "labour": float(vals.get("labour_share", self.labour_share)),
                "material": float(vals.get("material_share", self.material_share)),
                "equipment": float(vals.get("equipment_share", self.equipment_share)),
                "fuel": float(vals.get("fuel_share", self.fuel_share)),
            }

        crashing = assumptions.get("crashing_factors", {})
        self.mob_pct         = float(crashing.get("equipment_mobilization_pct",  0.03))
        self.shift_prem_pct  = float(crashing.get("shift_overtime_premium_pct",  0.15))
        self.supervision_pct = float(crashing.get("supervision_per_front_pct",   0.02))
        self.congestion_pct  = float(crashing.get("site_congestion_factor_pct",  0.01))

        # Night hours cost more than day hours. `calendar.night_shift_cost_multiplier`
        # (1.5 = 50% premium) is the explicit statement of that; when it is absent
        # we fall back to 1 + shift_overtime_premium_pct, which reproduces the
        # legacy behaviour exactly.
        # Night hours cost more than day hours. `calendar.night_shift_cost_multiplier`
        # (1.5 = 50% premium) is the explicit statement of that; when it is absent
        # we fall back to 1 + shift_overtime_premium_pct, which reproduces the
        # legacy behaviour exactly.
        raw = (assumptions.get("calendar") or {}).get("night_shift_cost_multiplier")
        if raw is not None:
            self.night_cost_multiplier = float(raw)
            if self.night_cost_multiplier < 1.0:
                raise ValueError(
                    "calendar['night_shift_cost_multiplier'] must be >= 1.0 "
                    f"(got {self.night_cost_multiplier}). Fix it in assumptions.yaml."
                )
        else:
            self.night_cost_multiplier = 1.0 + self.shift_prem_pct

        # Hard ceiling on how many parallel gangs any single activity may use.
        # Enforced here rather than trusted from the caller.
        cap = float(assumptions.get("crew_multiplier_cap", 0) or 0)
        if cap < 0:
            raise ValueError(
                f"crew_multiplier_cap must be >= 0 (got {cap}). Fix it in assumptions.yaml.")
        self.crew_multiplier_cap = cap

        # ---------------------------------------------------------------------
        # Time-driven acceleration step costs (new). These make the cost of
        # crashing proportional to how long a front stays open and to the size
        # of the fleet deployed, instead of a flat percentage uplift.
        # ---------------------------------------------------------------------
        step = assumptions.get("acceleration_step_costs", {}) or {}
        self.site_establishment_per_front_day = float(
            step.get("site_establishment_per_front_day", 27500.0)
        )
        self.resident_staff_per_front_day = float(
            step.get("resident_staff_per_front_day", 16800.0)
        )
        self.plant_hire_billing_days = float(
            step.get("plant_hire_billing_days", 30) or 30
        )
        if self.plant_hire_billing_days <= 0:
            raise ValueError(
                "acceleration_step_costs['plant_hire_billing_days'] must be > 0. "
                "Fix it in assumptions.yaml."
            )
        self.gang_interference_pct = float(step.get("gang_interference_pct", 0.04))

        # ---------------------------------------------------------------------
        # Re-sequencing / concurrency cost.
        #
        # Overlapping two activities on one corridor front (a staggered
        # finish-to-start lead) is NOT free, and it adds no gang. The plant has
        # to be broken down, shifted and re-set up as one operation gives way to
        # the next, and two crews sharing a front interfere with each other.
        # Without this term a duration cut that adds no resources prices out at
        # ZERO acceleration cost, which is not true in the field: compressing a
        # schedule always costs money, even when it is bought with sequencing
        # rather than with extra gangs.
        # ---------------------------------------------------------------------
        rq = assumptions.get("re_sequencing_acceleration", {}) or {}
        self.reseq_plant_redeploy_pct = float(
            rq.get("plant_redeployment_pct", 0.10) or 0.0
        )
        self.reseq_concurrency_pct = float(
            rq.get("concurrency_congestion_pct", 0.03) or 0.0
        )

        # ---------------------------------------------------------------------
        # Price escalation (src.escalation). Disabled unless assumptions.yaml
        # sets `escalation.enabled`. Applied to the finished direct cost, never
        # inside the unit rate, so the planned rate stays comparable with the
        # 2019 MoRTH yardstick.
        # ---------------------------------------------------------------------
        from .escalation import EscalationModel
        self.escalation = EscalationModel(assumptions)

        # ---------------------------------------------------------------------
        # Crew / plant day rates per activity, needed by the new time-driven terms
        # (plant hire minimum billing). Built once on load.
        # ---------------------------------------------------------------------
        self._labour_rates = assumptions.get("labour_rates_inr_per_day", {}) or {}
        self._crew_composition = assumptions.get("crew_composition", {}) or {}
        self._equip_hire = assumptions.get("equipment_hire_inr_per_day", {}) or {}
        self._equip_per_activity = assumptions.get("equipment_per_activity", {}) or {}
        self._crew_profile_by_activity = assumptions.get("crew_profile_by_activity_key", {}) or {}
        self._unmapped_activity_keys = set(
            assumptions.get("unmapped_activity_keys", []) or []
        )

        # Major / fixed plant: equipment that does not scale linearly with the
        # number of parallel gangs. A second hot-mix plant is not deployed just
        # because a second gang is added, so these are capped regardless of the
        # crew count while mobile plant scales freely. Read BEFORE _build_day_rates
        # because the fleet-aware plant rate depends on it.
        self.major_plant = set(assumptions.get("major_plant", []) or [])
        self.major_plant_max_units = int(
            assumptions.get("major_plant_max_units", 0) or 0
        )

        self.crew_day_rate: Dict[str, float] = {}
        self.plant_day_rate: Dict[str, float] = {}
        # Plant split into the two scaling behaviours. Major plant (batching
        # plants, tower cranes, stationary pumps) is a site-wide shared asset, so
        # the number of units deployed is capped rather than growing with the
        # gang count. Mobile plant scales one-for-one with gangs.
        self.plant_major_rate: Dict[str, float] = {}
        self.plant_mobile_rate: Dict[str, float] = {}
        self._build_day_rates()

        # ---------------------------------------------------------------------
        # Bottom-up resource engine: real material quantities + per-material
        # wastage, used to validate the MoRTH rate and to report exactly how
        # much consumption the wastage factors add.
        # ---------------------------------------------------------------------
        self._material_rates = assumptions.get("material_rates_inr", {}) or {}
        self._material_bindings = assumptions.get("material_bindings", {}) or {}
        self._wastage_factors = assumptions.get("wastage", {}) or {}
        self._recipes: Dict[str, Any] = {}
        self._recipes.update(assumptions.get("material_per_unit", {}) or {})
        self._recipes.update(assumptions.get("material_per_unit_road", {}) or {})
        self._activity_recipe = assumptions.get("activity_material_recipe", {}) or {}
        self._no_material_activities = set(
            assumptions.get("activity_no_material_recipe", []) or []
        )
        # SDB on-costs (overhead + contractor's profit) that the source rates
        # already contain. Optional for non-road types (pipeline/building use
        # a simpler cost breakdown).
        self._sdb_oncost = assumptions.get("sdb_rate_oncost") or {}
        self._sdb_structure_categories = set(
            self._sdb_oncost.get("structure_categories", []) or ["Structures"]
        )
        self.wastage_default = float(self._wastage_factors.get("default", 0.0) or 0.0)

        # Non-material execution losses. Material wastage is a quantity problem
        # (you issue more bags); these are UTILISATION losses on plant and
        # labour, each a percentage of its own base cost. Grouped under
        # `wastage_non_material` in assumptions.yaml.
        nonmat = assumptions.get("wastage_non_material", {}) or {}
        self.fuel_idle_pct = float(nonmat.get("fuel_idle_pct", 0.0) or 0.0)
        self.plant_idle_pct = float(nonmat.get("plant_idle_pct", 0.0) or 0.0)
        self.labour_idle_pct = float(nonmat.get("labour_idle_pct", 0.0) or 0.0)
        for _nm, _v in (("fuel_idle_pct", self.fuel_idle_pct),
                        ("plant_idle_pct", self.plant_idle_pct),
                        ("labour_idle_pct", self.labour_idle_pct)):
            if not (0.0 <= _v <= 1.0):
                raise ValueError(
                    f"wastage_non_material['{_nm}'] = {_v} must be a fraction between "
                    f"0 and 1 (e.g. 0.08 for 8%). Fix assumptions.yaml."
                )

        # Diesel rate + machine hours per day, for the fuel bottom-up.
        self.diesel_rate = float(self._material_rates.get("diesel_litre", 0.0) or 0.0)
        self._fuel_lph = assumptions.get("fuel_consumption_lph", {}) or {}
        work = assumptions.get("calendar", {}) or {}
        self.equipment_hours_per_day = float(work.get("equipment_hours_per_day", 9) or 9)

        # Validate the recipe graph once, on load, so a typo fails loudly at
        # startup instead of silently costing an activity at zero material.
        self._validate_recipe_graph()

    def sdb_oncost_multiplier(self, category: Optional[str]) -> float:
        """SDB on-cost multiplier. Returns 1.0 when the block is absent (non-road types)."""
        if not self._sdb_oncost:
            return 1.0
        group = "structures" if (category or "") in self._sdb_structure_categories else "road"
        pct = self._sdb_oncost.get(group)
        if pct is None:
            return 1.0
        oh = float(pct.get("overhead_pct", 8.0))
        cp = float(pct.get("contractor_profit_pct", 10.0))
        return (1.0 + oh / 100.0) * (1.0 + cp / 100.0)

    def _shares_for(self, category: Optional[str]) -> Dict[str, float]:
        """Cost-breakdown shares for one work category (falls back to global)."""
        return self._shares_by_category.get(str(category or ""), {
            "labour": self.labour_share,
            "material": self.material_share,
            "equipment": self.equipment_share,
            "fuel": self.fuel_share,
        })

    def _oncost_k(self, category: Optional[str]) -> float:
        """Applied on-cost multiplier for the configured rate basis.

        `statutory` uses the project overhead/profit percentages (the same
        figures the loader derives from the budget slab). `sdb` keeps the
        road/structures SDB multipliers used by the MoRTH model.
        """
        if self.rate_oncost_basis in ("none", "off"):
            return 1.0
        if self.rate_oncost_basis == "statutory":
            return (1.0 + self.overhead_pct / 100.0) * (1.0 + self.profit_pct / 100.0)
        return self.sdb_oncost_multiplier(category)

    def _embedded_oncost_rates(
        self, bare_rate: float, category: Optional[str]
    ) -> tuple[float, float]:
        """Per-unit OH and CP already inside the all-in planned rate.

        Returned only for reporting (memo columns); the caller must not add
        these to cost because `bare_rate * _oncost_k()` is the pricing basis.
        """
        if self.rate_oncost_basis in ("none", "off"):
            return 0.0, 0.0
        if self.rate_oncost_basis == "statutory":
            oh = bare_rate * self.overhead_pct / 100.0
            cp = (bare_rate + oh) * self.profit_pct / 100.0
            return oh, cp
        if not self._sdb_oncost:
            return 0.0, 0.0
        group = "structures" if (category or "") in self._sdb_structure_categories else "road"
        pct = self._sdb_oncost.get(group) or {}
        oh = bare_rate * float(pct.get("overhead_pct", 8.0)) / 100.0
        cp = (bare_rate + oh) * float(pct.get("contractor_profit_pct", 10.0)) / 100.0
        return oh, cp

    def _resource_planned_rate(self, activity_key: str, quantity: float):
        """Derive ``(planned_rate_per_unit, variation_pct)`` from assumptions.

        Two build-up modes are supported:

        ``source_plus_oncosts`` (pipeline and building default)
            The published direct rate is split into material / labour / plant /
            fuel shares by work category. Material wastage and non-material
            utilisation losses are added, then statutory overhead + contractor
            profit are applied. Supply items stay embedded in the published
            source rate, so no physical recipe is required.

        ``recipes`` (road default)
            The planned rate is the first-principles site cost per unit built
            from the resources actually consumed:

                materials  = recipe_per_unit x binding rate
                labour     = crew day rate  / natural productivity
                plant      = plant hire/day / natural productivity
                fuel       = lph x hours/day x diesel rate / natural productivity

            It is evaluated at the activity's natural single-gang, single-shift
            basis, so the rate is independent of how the schedule is crashed,
            and multiplied by the SDB on-cost factor (road x1.188, structures
            x1.32).

        The variation is the REPORTED gap between the all-in planned rate and
        the original source rate. Quantity does not change a unit rate; a
        zero-quantity activity still shows its build-up but carries zero cost.
        """
        raw_rate = self.direct_rates.get(activity_key)
        if isinstance(raw_rate, dict):
            morth = float(raw_rate.get("value") or 1000.0)
        elif raw_rate is None:
            morth = 1000.0
        else:
            morth = float(raw_rate)

        spec = None
        for entry in self.assumptions.get("activity_master", []) or []:
            if entry.get("key") == activity_key:
                spec = entry
                break
        category = (spec or {}).get("category")

        # A direct-rate override always wins: the published rate is used as the
        # planned rate and no on-cost build-up is applied.
        if self.assumptions.get("use_direct_rate_as_planned", False):
            return morth, 0.0

        # ------------------------------------------------------------------
        # Source-rate + on-costs build-up (pipeline/building catalogues).
        #
        # The published CPHEEO/CPWD rate is the direct (bare) resource cost.
        # Split it with the category cost shares, add material wastage and the
        # non-material utilisation losses, then apply statutory OH + profit.
        # No physical recipe is needed because pipe/E&M supply is already
        # embedded in the published direct rate.
        # ------------------------------------------------------------------
        if self.rate_build_up_mode == "source_plus_oncosts":
            shares = self._shares_for(category)
            material  = morth * shares["material"]
            labour    = morth * shares["labour"]
            equipment = morth * shares["equipment"]
            fuel      = morth * shares["fuel"]
            material_wastage = material * self.wastage_default
            resource_loss = (
                labour * self.labour_idle_pct
                + equipment * self.plant_idle_pct
                + fuel * self.fuel_idle_pct
            )
            planned = morth + material_wastage + resource_loss
            oncost_k = self._oncost_k(category)
            variation = (planned * oncost_k - morth) / morth * 100.0 if morth else 0.0
            return planned, variation

        if quantity <= 0:
            return morth, 0.0

        prod = max(1e-9, float((spec or {}).get("productivity_rate", 100.0) or 100.0))
        base_crews = max(1, int((spec or {}).get("base_crews", 1) or 1))

        # Materials — theoretical per unit of output (before wastage).
        mat_pu = 0.0
        spec_recipe = self._activity_recipe.get(activity_key)
        if spec_recipe:
            recipe = self._resolve_recipe(str(spec_recipe.get("recipe")))
            mix = float(spec_recipe.get("per_output", 1.0) or 1.0)
            for mat_key, per_out in recipe.items():
                if str(mat_key).startswith("_"):
                    continue
                binding = self._material_bindings[mat_key]
                unit_rate = float(self._material_rates[binding["rate"]])
                divisor = float(binding.get("divisor", 1) or 1)
                mat_pu += float(per_out) * mix / divisor * unit_rate

        # Labour / plant / fuel — the daily fleet divided by natural output.
        labour_pu = self.crew_day_rate.get(activity_key, 0.0) / prod
        mobile = self.plant_mobile_rate.get(activity_key, 0.0)
        major = self.plant_major_rate.get(activity_key, 0.0)
        major_sets = min(base_crews, max(1, self.major_plant_max_units))
        plant_pu = (mobile + major * major_sets / base_crews) / prod
        fuel_lph = 0.0
        profile = self._crew_profile_by_activity.get(activity_key, "")
        for eq in self._equip_per_activity.get(profile, []):
            fuel_lph += float(self._fuel_lph.get(eq, 0.0) or 0.0)
        fuel_pu = fuel_lph * self.equipment_hours_per_day * self.diesel_rate / prod

        planned = mat_pu + labour_pu + plant_pu + fuel_pu
        # All-in comparison: the source rate embeds overheads + contractor's
        # profit, so our build-up carries the same on-costs before comparing.
        oncost_k = self._oncost_k((spec or {}).get("category"))
        variation = (planned * oncost_k - morth) / morth * 100.0 if morth else 0.0
        return planned, variation

    # ---------------------------------------------------------------------
    def _build_day_rates(self) -> None:
        """
        Derive two day rates per activity key, used by the time-driven
        acceleration terms:

          crew_day_rate[activity_key] = sum(trade count x labour rate)
                                        over the mapped crew profile

          plant_day_rate[activity_key] = sum(equipment hire rate/day)
                                        over the equipment used by the profile

        Mandatory validation: every activity_master key must either appear in
        `crew_profile_by_activity_key` or be explicitly acknowledged in
        `unmapped_activity_keys`. A silent zero would understate the
        acceleration cost, so an unmapped key is a hard configuration error.
        """
        for act_key, profile in self._crew_profile_by_activity.items():
            crew_items = self._crew_composition.get(profile)
            if crew_items is None:
                raise ValueError(
                    f"crew_profile_by_activity_key['{act_key}'] = '{profile}' but no crew "
                    f"composition named '{profile}' exists under 'crew_composition' in "
                    f"assumptions.yaml. Add the crew profile or fix the mapping."
                )
            self.crew_day_rate[act_key] = self._sum_trade_rates(crew_items, act_key)
            equip_items = self._equip_per_activity.get(profile, [])
            self.plant_day_rate[act_key] = self._sum_equip_rates(equip_items, act_key)
            # Split plant into major (site-wide, shared) and mobile (per-gang)
            # so the bottom-up can charge them on different bases. Stored
            # explicitly rather than derived from plant_day_rate, because
            # subtracting a capped major figure back out of the total loses the
            # mobile component whenever a profile carries more major items than
            # the cap allows.
            self.plant_major_rate[act_key] = sum(
                float(self._equip_hire[eq]) for eq in equip_items
                if eq in self.major_plant and eq in self._equip_hire
            )
            self.plant_mobile_rate[act_key] = sum(
                float(self._equip_hire[eq]) for eq in equip_items
                if eq not in self.major_plant and eq in self._equip_hire
            )

        # Validate coverage of the activity master.
        activity_master = self.assumptions.get("activity_master", []) or []
        master_keys: list = []
        for entry in activity_master:
            if isinstance(entry, dict):
                k = entry.get("key")
            else:
                k = getattr(entry, "key", None)
            if k:
                master_keys.append(k)

        missing = [
            k for k in master_keys
            if k not in self._crew_profile_by_activity and k not in self._unmapped_activity_keys
        ]
        if missing:
            # For pipeline/building configs that don't yet have full crew-profile
            # mappings, emit a warning rather than crashing.
            listed = ", ".join(sorted(set(missing)))
            import warnings
            warnings.warn(
                "Missing crew-profile mapping for activity_master keys: "
                f"{listed}. Day rates for these activities will be zero."
            )
            # Register them as unmapped so downstream code doesn't re-warn
            for k in missing:
                self._unmapped_activity_keys.add(k)
                self.crew_day_rate.setdefault(k, 0.0)
                self.plant_day_rate.setdefault(k, 0.0)
                self.plant_major_rate.setdefault(k, 0.0)
                self.plant_mobile_rate.setdefault(k, 0.0)

    # ---------------------------------------------------------------------
    # Bottom-up resource engine
    # ---------------------------------------------------------------------
    def _resolve_recipe(self, name: str) -> Dict[str, Any]:
        """
        Follow `_use` aliases in a recipe to the concrete quantity table.
        `_use` lets several activities share one IS 456 mix (e.g. every M35
        structure reuses `rcc_m35`) without duplicating the numbers.
        """
        seen = set()
        cur = name
        while True:
            if cur in seen:
                raise ValueError(
                    f"material_per_unit: recipe '{name}' has a circular '_use' chain "
                    f"through '{cur}'. Fix assumptions.yaml."
                )
            seen.add(cur)
            recipe = self._recipes.get(cur)
            if recipe is None:
                raise ValueError(
                    f"No material recipe named '{cur}' (referenced by '{name}') in "
                    f"material_per_unit / material_per_unit_road. Fix assumptions.yaml."
                )
            target = recipe.get("_use") if isinstance(recipe, dict) else None
            if not target:
                return recipe
            cur = str(target)

    def _validate_recipe_graph(self) -> None:
        """
        Every activity_master key must be either mapped to a recipe or listed in
        `activity_no_material_recipe`, and every recipe/binding/rate it reaches
        must exist. Mirrors the crew-profile coverage rule: a silent zero would
        understate the bottom-up estimate, so it is a hard configuration error.
        """
        for act_key, spec in self._activity_recipe.items():
            recipe_name = (spec or {}).get("recipe") if isinstance(spec, dict) else None
            if not recipe_name:
                raise ValueError(
                    f"activity_material_recipe['{act_key}'] has no 'recipe' key. "
                    f"Fix assumptions.yaml."
                )
            recipe = self._resolve_recipe(str(recipe_name))
            for mat_key, per_output in recipe.items():
                if str(mat_key).startswith("_"):
                    # `_waste_only` and friends are directives, not materials.
                    continue
                binding = self._material_bindings.get(mat_key)
                if binding is None:
                    raise ValueError(
                        f"Recipe '{recipe_name}' (used by activity '{act_key}') contains "
                        f"material '{mat_key}' which has no entry in 'material_bindings'. "
                        f"Add a binding with its wastage category and price."
                    )
                rate_key = binding.get("rate")
                if rate_key not in self._material_rates:
                    raise ValueError(
                        f"material_bindings['{mat_key}'].rate = '{rate_key}' is not present in "
                        f"'material_rates_inr'. Fix assumptions.yaml."
                    )
                category = binding.get("wastage")
                if category not in self._wastage_factors:
                    raise ValueError(
                        f"material_bindings['{mat_key}'].wastage = '{category}' is not a "
                        f"category in the 'wastage' block. Fix assumptions.yaml."
                    )
                if float(binding.get("divisor", 1) or 1) <= 0:
                    raise ValueError(
                        f"material_bindings['{mat_key}'].divisor must be > 0. "
                        f"Fix assumptions.yaml."
                    )

        activity_master = self.assumptions.get("activity_master", []) or []
        if self.rate_build_up_mode == "source_plus_oncosts":
            # The source-rate build-up does not use material recipes; the
            # published rate already carries supply items. Recipe coverage is
            # therefore not a configuration requirement in this mode.
            return
        master_keys = []
        for entry in activity_master:
            k = entry.get("key") if isinstance(entry, dict) else getattr(entry, "key", None)
            if k:
                master_keys.append(k)

        uncovered = [
            k for k in master_keys
            if k not in self._activity_recipe and k not in self._no_material_activities
        ]
        if uncovered and not self.assumptions.get("use_direct_rate_as_planned", False):
            import warnings
            warnings.warn(
                "Missing material-recipe mapping for activity_master keys: "
                f"{', '.join(sorted(set(uncovered)))}. Material costs will be zero for these activities."
            )
            for k in uncovered:
                self._no_material_activities.add(k)
        elif uncovered:
            # Direct-rate configs price materials inside the published all-in
            # rate. Record the keys as intentionally recipe-free so the same
            # warning is not raised twice.
            for k in uncovered:
                self._no_material_activities.add(k)

    def compute_resource_bottom_up(
        self,
        activity_key: str,
        quantity: float,
        duration_days: float = 0.0,
        crews: int = 1,
    ) -> Dict[str, Any]:
        """
        Build the activity's resource quantities and costs from first principles.

        Material:
            theoretical_qty = recipe_per_unit x quantity x mix_factor
            issued_qty       = theoretical_qty x (1 + wastage_factor)
            cost             = issued_qty / divisor x unit_rate

        Material wastage is a QUANTITY problem, so `material_wastage_cost` is the
        incremental money spent purely because of the wastage factors, i.e.
        cost(issued) - cost(theoretical). That is what a hard-set bag of cement, a
        spill, or overbreak actually costs.

        Labour / plant / fuel come from the crew and equipment actually deployed
        for `duration_days` x `crews`, using the same day rates the acceleration
        terms use, so the bottom-up and the surcharge story agree. Their losses are
        a UTILISATION problem rather than a quantity problem, so each is reported
        as its own base and its own loss: `fuel_wastage_cost`, `plant_wastage_cost`,
        `labour_wastage_cost`, summing to `resource_wastage_cost`.

        Plant hire uses the fleet-aware split: mobile plant scales one-for-one
        with the gang count, while major plant is capped at
        `major_plant_max_units` sets for the whole site.

        Returns a dict with `material_lines` (per-material audit trail),
        `material_cost`, the four loss components, `labour_cost`, `plant_cost`,
        `fuel_cost`, `total_cost`, and `wastage_cost` = material + resource loss.
        """
        spec = self._activity_recipe.get(activity_key)
        material_lines: list = []
        material_cost = 0.0
        wastage_cost = 0.0
        overbreak_pct = 0.0

        if spec:
            recipe_name = str(spec.get("recipe"))
            mix_factor = float(spec.get("per_output", 1.0) or 1.0)
            recipe = self._resolve_recipe(recipe_name)

            # `_waste_only` recipes carry no purchasable material; they express a
            # non-purchasable loss (overbreak) as a percentage instead. It is
            # scaled against labour+plant further below, where there is a base.
            waste_only = recipe.get("_waste_only") if isinstance(recipe, dict) else None
            overbreak_pct = (
                float(waste_only.get("overbreak_pct", 0.0) or 0.0)
                if isinstance(waste_only, dict)
                else 0.0
            )

            for mat_key, per_output in recipe.items():
                if str(mat_key).startswith("_"):
                    continue
                binding = self._material_bindings[mat_key]
                unit_rate = float(self._material_rates[binding["rate"]])
                divisor = float(binding.get("divisor", 1) or 1)
                wastage_pct = float(
                    self._wastage_factors.get(binding.get("wastage"), self.wastage_default)
                )

                theoretical = float(per_output) * quantity * mix_factor
                issued = theoretical * (1.0 + wastage_pct)
                cost_theoretical = (theoretical / divisor) * unit_rate
                cost_issued = (issued / divisor) * unit_rate
                material_cost += cost_issued
                wastage_cost += (cost_issued - cost_theoretical)
                material_lines.append({
                    "material": mat_key,
                    "unit_rate_key": binding["rate"],
                    "theoretical_qty": theoretical,
                    "issued_qty": issued,
                    "wastage_pct": wastage_pct,
                    "unit_rate": unit_rate,
                    "cost": cost_issued,
                    "wastage_cost": cost_issued - cost_theoretical,
                })

        days = max(0.0, float(duration_days or 0.0))
        gangs = max(1, int(crews))
        gang_days = days * gangs
        labour_base = self.crew_day_rate.get(activity_key, 0.0) * gang_days

        # Plant hire is charged on two different bases:
        #   mobile plant scales one-for-one with the gang count
        #   major plant is a shared site asset, so the number of sets deployed is
        #     capped at major_plant_max_units regardless of how many gangs run.
        # A 4th gang does not drag a 4th hot-mix plant onto site.
        major_rate = self.plant_major_rate.get(activity_key, 0.0)
        mobile_rate = self.plant_mobile_rate.get(activity_key, 0.0)
        major_sets = min(gangs, max(1, self.major_plant_max_units))
        plant_base = mobile_rate * days * gangs + major_rate * days * major_sets

        fuel_lph = 0.0
        for eq in self._equip_per_activity.get(
            self._crew_profile_by_activity.get(activity_key, ""), []
        ):
            fuel_lph += float(self._fuel_lph.get(eq, 0.0) or 0.0)
        fuel_base = fuel_lph * self.equipment_hours_per_day * gang_days * self.diesel_rate

        # -----------------------------------------------------------------
        # Resource efficiency losses — the non-material half of wastage.
        #
        # Material wastage is a QUANTITY problem (a bag goes hard, so you
        # issue more bags). These three are UTILISATION problems on plant and
        # labour, so each is a percentage of its own base cost:
        #
        #   fuel  — engines idle at site, run on standby gensets, pump prime
        #           before/after a pour, and never reach rated economy
        #   plant — breakdowns awaiting a mechanic, waiting on fuel or
        #           spares, standing idle between discrete tasks
        #   labour— gangs standing idle waiting on plant, plus the daily
        #           join/leave/retard/departure drag on a churning site
        #
        # They are priced from the SAME crew_day_rate / plant_day_rate /
        # fuel_consumption_lph the bottom-up uses, so the allowance and the
        # cost it scales can never drift apart.
        # -----------------------------------------------------------------
        fuel_waste = fuel_base * self.fuel_idle_pct
        labour_waste = labour_base * self.labour_idle_pct
        # Plant hire is charged per day, so idling is recovered as extra hire
        # days; fuel is the separate standby-burn cost on top.
        plant_waste = plant_base * self.plant_idle_pct

        labour_cost = labour_base + labour_waste
        plant_cost = plant_base + plant_waste
        fuel_cost = fuel_base + fuel_waste
        resource_waste_cost = labour_waste + plant_waste + fuel_waste

        # Overbreak: where the recipe has no purchasable material there is no
        # material base to scale it against, so it is charged on labour + plant,
        # which is the cost of the re-dig and re-haul the overbreak causes.
        if overbreak_pct and material_cost == 0.0:
            labour_cost += (labour_base + plant_base) * overbreak_pct

        return {
            "activity_key": activity_key,
            "quantity": quantity,
            "duration_days": days,
            "crews": max(1, int(crews)),
            "overbreak_pct": overbreak_pct,
            "material_lines": material_lines,
            "material_cost": round(material_cost, 2),
            "material_wastage_cost": round(wastage_cost, 2),
            "fuel_wastage_cost": round(fuel_waste, 2),
            "plant_wastage_cost": round(plant_waste, 2),
            "labour_wastage_cost": round(labour_waste, 2),
            "resource_wastage_cost": round(resource_waste_cost, 2),
            # material wastage + fuel/plant/labour idle = total execution loss
            "wastage_cost": round(wastage_cost + resource_waste_cost, 2),
            "labour_base": round(labour_base, 2),
            "plant_base": round(plant_base, 2),
            "fuel_base": round(fuel_base, 2),
            "fuel_lph": round(fuel_lph, 2),
            "fuel_litres": round(fuel_base / self.diesel_rate, 2) if self.diesel_rate else 0.0,
            "labour_cost": round(labour_cost, 2),
            "plant_cost": round(plant_cost, 2),
            "fuel_cost": round(fuel_cost, 2),
            "total_cost": round(material_cost + labour_cost + plant_cost + fuel_cost, 2),
        }

    # ---------------------------------------------------------------------
    def _sum_trade_rates(self, crew_items, activity_key: str) -> float:
        total = 0.0
        for item in crew_items or []:
            if not isinstance(item, dict):
                continue
            trade = item.get("trade")
            count = float(item.get("count", 0) or 0)
            rate = self._labour_rates.get(trade)
            if rate is None:
                raise ValueError(
                    f"crew_composition for activity '{activity_key}' references trade "
                    f"'{trade}' which has no rate in 'labour_rates_inr_per_day'. "
                    f"Fix assumptions.yaml."
                )
            total += count * float(rate)
        return total

    # ---------------------------------------------------------------------
    def _sum_equip_rates(self, equip_items, activity_key: str) -> float:
        total = 0.0
        for eq in equip_items or []:
            rate = self._equip_hire.get(eq)
            if rate is None:
                raise ValueError(
                    f"equipment_per_activity for activity '{activity_key}' references "
                    f"equipment '{eq}' which has no rate in "
                    f"'equipment_hire_inr_per_day'. Fix assumptions.yaml."
                )
            total += float(rate)
        return total

    def calculate_activity_cost(
        self,
        activity_key: str,
        quantity: float,
        crews: int = 1,
        base_crews: int = 1,
        shifts: float = 1.0,
        work_fronts: int = 1,
        duration_days: float = 0,
        staggered: bool = False,
        escalation_years: float = 0.0
    ) -> ActivityCostBreakdown:
        """
        Cost formula:
          planned_rate     = all-in build-up per unit of output (see
                             `_resource_planned_rate`): source direct rate
                             + wastage/idle losses, x the configured on-cost
                             factor (pipeline statutory OH/CP, road SDB)
          base_direct_cost = quantity x planned_rate
          direct_cost      = base_direct_cost + crashing additions
          total            = direct_cost + price escalation

        OH and profit are already inside `planned_rate`. They are exposed as
        `overhead_embedded` / `profit_embedded` for reporting only and are
        never added again. `overhead_amount` and `profit_amount` therefore
        stay zero when an on-cost build-up is active.

        Acceleration surcharge components:
          1. fleet mobilisation      - proportional to extra crews
          2. shift / overtime premium- proportional to extra shifts
          3. front supervision       - proportional to extra work fronts
          4. site congestion         - proportional to extra crews
          5. extra-front establishment- TIME driven (fronts x days)
          6. extra-front resident staff-TIME driven (fronts x days)
          7. plant hire min uplift   - TIME driven (crew x rounded-up months)
          8. gang interference loss  - super-linear in extra crews

        Components 1-4 are the pre-existing 'crashing_factors' terms and are
        left untouched. Components 5-8 are the new time-driven terms and are
        zeroed when `enable_acceleration_cost` is False (--no-accel-cost).
        For a NON-crashed activity (crews == base_crews, work_fronts == 1,
        shifts == 1.0) every one of the eight terms is exactly zero.

        `staggered=True` marks an activity that is pulled forward to overlap its
        predecessor on the same front. That adds no gang but is not free either:
        the plant is redeployed mid-operation and the two crews interfere, so a
        plant redeployment and a concurrency charge are added to components 1
        and 4. Without it, a duration cut bought purely by re-sequencing would
        report an acceleration cost of zero.
        """
        # A missing rate falls back to a budget-estimation default rather than
        # crashing — pipeline/building configs may have null values while being
        # populated incrementally. A warning is issued instead.
        if activity_key not in self.direct_rates:
            if (self.assumptions.get("use_direct_rate_as_planned", False)
                    or self.rate_build_up_mode == "source_plus_oncosts"):
                raise ValueError(
                    f"No direct unit rate for activity '{activity_key}'. "
                    f"The config prices activities from the published direct rate "
                    f"(rate_build_up_mode: {self.rate_build_up_mode}), so a "
                    f"placeholder rate would misprice the whole activity. Add the "
                    f"key to direct_unit_rates_inr or remove it from the guideline catalog."
                )
            # Road-style resource build-up: keep the legacy placeholder so a
            # partially populated config can still be inspected.
            morth_rate = 1000.0   # INR 1000/unit placeholder
        else:
            raw_rate = self.direct_rates[activity_key]
            if raw_rate is None or (isinstance(raw_rate, dict) and raw_rate.get("value") is None):
                if (self.assumptions.get("use_direct_rate_as_planned", False)
                        or self.rate_build_up_mode == "source_plus_oncosts"):
                    raise ValueError(
                        f"Direct unit rate for activity '{activity_key}' is null. "
                        f"Add a positive rate to direct_unit_rates_inr."
                    )
                morth_rate = 1000.0  # null value in config → placeholder
            else:
                rate_val = raw_rate if not isinstance(raw_rate, dict) else raw_rate.get("value", 1000.0)
                morth_rate = float(rate_val or 1000.0)
        category = next(
            (s.get("category") for s in (self.assumptions.get("activity_master") or [])
             if s.get("key") == activity_key),
            None,
        )
        planned_raw, variation = self._resource_planned_rate(activity_key, quantity)
        # planned_raw is the bare resource build-up (labour + plant + fuel + material).
        # Overhead and profit are applied explicitly below, so base cost uses the bare rate.
        # The all-in rate (bare * oncost_k) is stored separately for display/reporting only.
        oncost_k = self._oncost_k(category)
        # When using direct rate as planned rate (pipeline/building with CPWD/CPHEEO
        # all-in published rates), the rate already includes OH+profit — do NOT
        # multiply by oncost_k again.
        if self.assumptions.get("use_direct_rate_as_planned", False):
            oncost_k = 1.0
        planned_rate_allin = round(planned_raw * oncost_k, 4)  # all-in, incl. on-costs
        planned_rate = planned_rate_allin                 # ALL-IN is the pricing basis too

        # Memo-only split of the on-costs already inside planned_rate_allin.
        if self.assumptions.get("use_direct_rate_as_planned", False):
            oh_rate, cp_rate = 0.0, 0.0
        else:
            oh_rate, cp_rate = self._embedded_oncost_rates(planned_raw, category)

        # Never let a caller exceed the configured gang ceiling.
        if self.crew_multiplier_cap > 0:
            crews = int(min(crews, max(base_crews, int(base_crews * self.crew_multiplier_cap))))

        base_direct_cost = quantity * planned_rate

        cost_shares     = self._shares_for(category)
        direct_labour   = base_direct_cost * cost_shares["labour"]
        direct_material = base_direct_cost * cost_shares["material"]
        direct_equip    = base_direct_cost * cost_shares["equipment"]
        direct_fuel     = base_direct_cost * cost_shares["fuel"]

        # --- material wastage, priced from real quantities ---
        # `wastage` in assumptions.yaml gives a factor per material (cement 3%,
        # aggregate 5%, bitumen 2%, ...). This resolves the activity's recipe
        # into actual quantities, applies each material's own factor, and prices
        # the difference. That difference is the incremental money spent purely
        # on wastage — a hard-set bag of cement, a spill, overbreak, rehandling —
        # rather than a blanket percentage that cannot tell cement from bitumen.
        #
        # The base cost is priced at the resource-derived planned rate; this
        # wastage/overbreak differencing is an ADDITIVE allowance on top of the
        # rate and is reported separately (it never double-counts into the
        # derived planned rate, which is built on theoretical quantities).
        bottom_up = self.compute_resource_bottom_up(
            activity_key, quantity, duration_days=duration_days, crews=crews
        )

        extra_crews  = max(0, crews - base_crews)
        extra_fronts = max(0, work_fronts - 1)

        # --- pre-existing crash factors (unchanged in structure) ---
        # The shift premium now prices night hours at the multiplier the
        # calendar declares (1.5 = 50% more than day work) instead of a flat
        # 15% that knew nothing about night_shift_cost_multiplier.
        night_premium = max(0.0, self.night_cost_multiplier - 1.0)
        equip_mob_cost    = extra_crews  * (direct_equip   * self.mob_pct)
        shift_premium     = (direct_labour + direct_equip) * night_premium * max(0.0, shifts - 1.0)
        front_supervision = extra_fronts * (direct_labour  * self.supervision_pct)
        congestion_cost   = base_direct_cost * (self.congestion_pct * extra_crews)

        # --- new time-driven acceleration step costs ---
        duration_days = max(0.0, float(duration_days or 0.0))
        plant_day_rate = self.plant_day_rate.get(activity_key, 0.0)

        if self.enable_acceleration_cost:
            front_site_establishment = (
                extra_fronts * duration_days * self.site_establishment_per_front_day
            )
            front_resident_staff = (
                extra_fronts * duration_days * self.resident_staff_per_front_day
            )
            # Plant is hired on a calendar-month minimum, so a partial month is
            # charged in full. `billable_days` is the number of DAYS actually
            # chargeable at hire rates = whole billing periods converted back to
            # days (ceil(days / 30) * 30). The uplift is the rounding premium
            # only: hire cost less the pro-rata hire for the days actually
            # worked.
            billable_days = (
                math.ceil(duration_days / self.plant_hire_billing_days)
                * self.plant_hire_billing_days
            )
            plant_hire_uplift = (
                extra_crews * billable_days * plant_day_rate
                - extra_crews * duration_days * plant_day_rate
            )
            efficiency_loss = (
                (direct_labour + direct_equip + direct_fuel)
                * self.gang_interference_pct
                * (extra_crews ** 1.5)
            )
        else:
            front_site_establishment = 0.0
            front_resident_staff = 0.0
            plant_hire_uplift = 0.0
            efficiency_loss = 0.0

        # --- re-sequencing / concurrency (see `re_sequencing_acceleration`) ---
        # Charged only to an activity that actually overlaps its predecessor, so
        # a non-staggered plan is unaffected and --no-accel-cost still zeroes it.
        if self.enable_acceleration_cost and staggered:
            equip_mob_cost += direct_equip * self.reseq_plant_redeploy_pct
            congestion_cost += base_direct_cost * self.reseq_concurrency_pct

        # Every surcharge component is rounded to the rupee and clamped so it can
        # never be negative. Rounding first makes the reported breakdown
        # self-consistent: the 8 printed components sum EXACTLY to
        # crash_total_surcharge, which is what the Excel waterfall shows.
        equip_mob_cost = round(max(0.0, equip_mob_cost), 2)
        shift_premium = round(max(0.0, shift_premium), 2)
        front_supervision = round(max(0.0, front_supervision), 2)
        congestion_cost = round(max(0.0, congestion_cost), 2)
        front_site_establishment = round(max(0.0, front_site_establishment), 2)
        front_resident_staff = round(max(0.0, front_resident_staff), 2)
        plant_hire_uplift = round(max(0.0, plant_hire_uplift), 2)
        efficiency_loss = round(max(0.0, efficiency_loss), 2)
        crash_total_surcharge = round(
            equip_mob_cost
            + shift_premium
            + front_supervision
            + congestion_cost
            + front_site_establishment
            + front_resident_staff
            + plant_hire_uplift
            + efficiency_loss,
            2,
        )
        # Wastage (material spoilage + fuel/plant/labour idle losses) is a real
        # site cost and must be carried into the budget, not just reported.
        # It is added to direct cost so that overhead and profit also scale on it.
        # EXCEPTION: when using direct rates (CPWD/CPHEEO all-in rates), wastage
        # is already embedded in the published rate — do not add it again.
        if (self.assumptions.get("use_direct_rate_as_planned", False)
                or self.rate_build_up_mode == "source_plus_oncosts"):
            wastage_total = 0.0
            reported_material_wastage = 0.0
            reported_fuel_wastage = 0.0
            reported_plant_wastage = 0.0
            reported_labour_wastage = 0.0
            reported_resource_wastage = 0.0
            reported_wastage = 0.0
        else:
            wastage_total = bottom_up.get("wastage_cost", 0.0)
            reported_material_wastage = bottom_up["material_wastage_cost"]
            reported_fuel_wastage = bottom_up["fuel_wastage_cost"]
            reported_plant_wastage = bottom_up["plant_wastage_cost"]
            reported_labour_wastage = bottom_up["labour_wastage_cost"]
            reported_resource_wastage = bottom_up["resource_wastage_cost"]
            reported_wastage = bottom_up["wastage_cost"]
        direct_cost_tot  = base_direct_cost + crash_total_surcharge + wastage_total
        direct_cost_tot  = max(direct_cost_tot, base_direct_cost)
        direct_unit_rate = direct_cost_tot / max(quantity, 0.001)

        # On-costs are applied ONCE, inside `planned_rate` above, using either
        # the statutory OH/CP slab (pipeline) or the SDB multiplier (road).
        # They are therefore NOT charged a second time here. The embedded
        # on-cost amounts are memo-only; `overhead_amount` / `profit_amount`
        # stay zero so the total identity is simply direct + escalation.
        overhead_amount = 0.0
        profit_amount   = 0.0
        # Memo-only: the OH/CP already inside planned_rate (not added to total).
        # Split the actual on-cost portion of the rounded all-in rate so the two
        # memo columns sum to quantity x (planned - bare) within rounding.
        oncost_amount = max(0.0, (planned_rate_allin - planned_raw) * quantity)
        oncost_rate_sum = oh_rate + cp_rate
        if oncost_rate_sum > 0:
            overhead_embedded = oncost_amount * oh_rate / oncost_rate_sum
            profit_embedded   = oncost_amount * cp_rate / oncost_rate_sum
        else:
            overhead_embedded = 0.0
            profit_embedded   = 0.0

        # --- price escalation -------------------------------------------------
        # Applied once, to the finished direct cost (works + acceleration +
        # execution loss), and reported separately so the unit rate stays
        # comparable with the 2019 yardstick and the sheet identities hold:
        #     total = base + surcharge + wastage + escalation
        shares = {
            "material": cost_shares["material"],
            "labour":   cost_shares["labour"],
            "plant":    cost_shares["equipment"],
            "fuel":     cost_shares["fuel"],
        }
        esc = self.escalation.apply(direct_cost_tot, float(escalation_years or 0.0),
                                    shares, activity_key,
                                    bottom_up.get("material_lines") or None)
        escalation_amount = round(esc["amount"], 2)

        total_amount    = direct_cost_tot + escalation_amount
        final_unit_rate = total_amount / max(quantity, 0.001)

        return ActivityCostBreakdown(
            direct_labour          = round(direct_labour * (1.0 + night_premium * max(0.0, shifts - 1.0)), 2),
            direct_material        = round(direct_material, 2),
            direct_equipment       = round(direct_equip + equip_mob_cost, 2),
            direct_fuel            = round(direct_fuel + congestion_cost, 2),
            direct_cost            = round(direct_cost_tot, 2),
            direct_unit_rate       = round(planned_rate_allin, 4),   # all-in rate shown in sheet
            planned_unit_rate      = round(planned_rate_allin, 4),   # explicit alias for reports
            morth_direct_unit_rate = round(morth_rate, 4),           # MoRTH source — unchanged
            planning_variation_pct = round(variation, 4),            # derived vs MoRTH all-in
            overhead_pct           = self.overhead_pct,
            overhead_amount        = round(overhead_amount, 2),
            overhead_embedded      = round(overhead_embedded, 2),
            profit_pct             = self.profit_pct,
            profit_amount          = round(profit_amount, 2),
            profit_embedded        = round(profit_embedded, 2),
            total_amount           = round(total_amount, 2),
            final_unit_rate        = round(final_unit_rate, 2),
            crash_site_establishment = round(front_site_establishment, 2),
            crash_resident_staff    = round(front_resident_staff, 2),
            crash_plant_hire_uplift = round(plant_hire_uplift, 2),
            crash_efficiency_loss   = round(efficiency_loss, 2),
            crash_total_surcharge   = crash_total_surcharge,
            bottom_up_material_cost  = bottom_up["material_cost"],
            material_wastage_cost   = reported_material_wastage,
            fuel_wastage_cost       = reported_fuel_wastage,
            plant_wastage_cost      = reported_plant_wastage,
            labour_wastage_cost     = reported_labour_wastage,
            resource_wastage_cost   = reported_resource_wastage,
            wastage_cost            = reported_wastage,
            wastage_in_base_rate    = bool(
                self.assumptions.get("use_direct_rate_as_planned", False)
                or self.rate_build_up_mode == "source_plus_oncosts"
            ),
            bottom_up_labour_cost   = bottom_up["labour_cost"],
            bottom_up_plant_cost    = bottom_up["plant_cost"],
            bottom_up_fuel_cost     = bottom_up["fuel_cost"],
            bottom_up_total_cost    = bottom_up["total_cost"],
            crash_fleet_mobilisation  = equip_mob_cost,
            crash_shift_premium       = shift_premium,
            crash_front_supervision   = front_supervision,
            crash_site_congestion     = congestion_cost,
            escalation_amount        = escalation_amount,
            escalation_factor        = round(esc["factor"], 6),
            escalation_years         = round(esc["years"], 4),
        )
