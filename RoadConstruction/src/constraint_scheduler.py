"""
Constraint-Based Scheduling and Resource Optimization Engine for Road
Construction Projects.

Solves for a schedule satisfying both hard constraints:

    1. Project duration <= required duration
    2. Total project cost <= maximum budget

...while preserving the uncrashed baseline for side-by-side auditing.

How the search works
--------------------
Every activity has an ordered list of resource levels, from (base crews, single
day shift) up to (base crews x max multiplier, extended shifts). A solver state
is the chosen level per activity plus an optional staggered-handover flag.

The solver starts at the baseline and repeatedly makes the LEAST COST move: for
each activity on the current critical path it prices advancing that activity by
one level (and separately prices enabling staggered handover), computes rupees
per day saved, and applies the cheapest move that still fits the budget. It
stops the moment the duration target is met, so the answer is the cheapest
schedule that meets the deadline rather than the most over-cooked one.

Because durations come from the project calendar (see ProjectCalendar), every
candidate is priced by actually re-running CPM - the critical path can and does
move between moves.
"""
import copy
import math
import os
from typing import List, Dict, Tuple, Optional, Any
from datetime import date, timedelta

from .models import (ProjectInput, RoadActivity, ActivityReference,
                     ActivityCostBreakdown, ProjectScheduleResult)
from .data_loader import DataLoader
from .quantity_calculator import QuantityCalculator
from .cost_calculator import CostCalculator
from .duration_calculator import DurationCalculator
from .cpm_engine import CPMEngine
from .progress_calculator import ProgressCalculator
from .project_calendar import ProjectCalendar

# Shift intensities offered to the crashing search. Anything above 1.0 is part
# night work, so it delivers less output and costs more (ProjectCalendar).
SHIFT_LEVELS = (1.0, 1.3, 1.6, 1.8)


class ConstraintScheduler:
    """
    Solves for a feasible construction schedule satisfying:
    1. Project Duration <= required_duration
    2. Total Project Cost <= maximum_budget
    while preserving baseline results for side-by-side auditing.
    """

    def __init__(
        self,
        project_input: ProjectInput,
        data_loader: DataLoader,
        enable_acceleration_cost: bool = True
    ):
        self.input = project_input
        self.loader = data_loader
        self.assumptions = data_loader.assumptions
        self.guidelines = getattr(data_loader, "guidelines", {}) or {}
        self.enable_acceleration_cost = bool(enable_acceleration_cost)

        # The project calendar converts productivity EFFORT days into real
        # CALENDAR days (weekly rest, public holidays, monsoon slowdown).
        self.calendar = ProjectCalendar(
            self.assumptions.get("calendar"), project_input.start_date,
            terrain=project_input.terrain)

        self.qty_calc = QuantityCalculator(project_input, self.assumptions)
        self.quantities = self.qty_calc.calculate_quantities()

        overhead_pct = self.loader.get_overhead_percentage(project_input.budget_crore)
        profit_pct = self.loader.get_profit_percentage()
        self.cost_calc = CostCalculator(
            self.assumptions, overhead_pct, profit_pct,
            enable_acceleration_cost=self.enable_acceleration_cost,
        )
        self.dur_calc = DurationCalculator(
            self.assumptions, project_input.length_km, self.calendar)

        all_activity_specs = (
            self.guidelines.get("activity_master")
            or (self.guidelines.get("data") or {}).get("activity_master")
            or self.assumptions.get("activity_master", [])
        )
        if str(getattr(self.input, "project_type", "road") or "road").lower() == "pipeline":
            # Keep the catalogue and quantity ledger complete for auditability,
            # but do not schedule alternative branches or facility packages
            # whose derived quantity is zero for this project scope.
            self.activity_specs = [
                spec for spec in all_activity_specs
                if float(self.quantities.get(spec["key"], {}).get("quantity", 0.0) or 0.0) > 0
            ]
        else:
            self.activity_specs = all_activity_specs
        self.spec_by_id = {s["id"]: s for s in self.activity_specs}
        self.spec_by_key = {s["key"]: s for s in self.activity_specs}
        self.stagger_chains = [
            (item["predecessor_id"], item["successor_id"])
            for item in (self.assumptions.get("staggered_handover_chains", []) or [])
            if item["predecessor_id"] in self.spec_by_id
            and item["successor_id"] in self.spec_by_id
        ]

        # Labels for ActivityReference — read from config document block
        doc = (self.guidelines.get("data") or {}).get("document") or self.guidelines.get("document") or {}
        self._guidelines_filename = (
            os.path.basename(data_loader.guidelines_path)
            if hasattr(data_loader, "guidelines_path") else "guidelines.json"
        )
        self._assumptions_filename = (
            os.path.basename(data_loader.assumptions_path)
            if hasattr(data_loader, "assumptions_path") else "assumptions.yaml"
        )
        self._guideline_ref_label = doc.get("title", "Project Guidelines")
        self.stagger_chains = [
            (item["predecessor_id"], item["successor_id"])
            for item in (self.assumptions.get("staggered_handover_chains", []) or [])
            if item["predecessor_id"] in self.spec_by_id
            and item["successor_id"] in self.spec_by_id
        ]

        self.max_fronts = 1
        self.corridor_fronts = 1
        self.baseline_fronts = 1
        self.max_multiplier = 1
        # Keys whose project-type calculator does not cover the current catalog.
        # Pipeline raises instead of falling back; road/building record a warning.
        self._quantity_fallbacks = set()
        self.levels: Dict[str, List[Tuple[int, float]]] = {}

    # ------------------------------------------------------------------ setup
    def _build_crash_levels(self) -> None:
        """
        Pre-compute, for every activity, the ordered list of (crews, shifts)
        levels the search may step through. Index 0 is always the smallest
        organisation (one gang, single day shift); every later index is
        strictly more resource than the one before it.

        Three numbers bound the search:

        * `corridor_fronts`  - physical ceiling, one front per 3.5 km of road.
          You cannot deploy more gangs than there are stretches to work on.
        * `baseline_fronts`  - the NORMAL organisation, one gang per
          `work_front_span_km` km. A 25 km corridor is built by several gangs
          in parallel from day one; only a road <= that span starts with a
          single gang. This is where the solver starts.
        * `max_multiplier`   - how far crews may escalate. Bounded by
          `crew_multiplier_cap`, but never below `baseline_fronts`: the cap
          limits escalation, it cannot force a corridor back to one gang.
        """
        ptype = str(getattr(self.input, "project_type", "road") or "road").lower()
        cap = float(self.assumptions.get("crew_multiplier_cap", 6) or 6)
        if cap < 1:
            raise ValueError(
                "assumptions.yaml: crew_multiplier_cap must be >= 1 "
                f"(got {cap}). Fix it in assumptions.yaml."
            )

        if ptype == "pipeline":
            # Pipeline: user supplies number_of_work_fronts directly (from CLI --fronts).
            # Physical ceiling = min(fronts, 4) to avoid unrealistic over-splitting.
            # Baseline uses 1 front (single crew set, correct cost); optimizer can
            # scale up to the user-supplied fronts to shorten the schedule.
            supplied = max(1, int(getattr(self.input, "number_of_work_fronts", 1) or 1))
            corridor_fronts = min(supplied, 4)
            baseline_fronts = 1  # Single front at baseline — cost priced at 1 crew set
            # length check
            length = float(self.input.length_km or 0.0)
            if length <= 0:
                raise ValueError("pipeline length must be greater than 0 km.")

        elif ptype == "building":
            # Building: parallelism = number of blocks/towers.
            # Each block is an independent work front, but the BASELINE cost is
            # priced at 1 front (one crew set covers the whole building).
            # Duration at baseline = total work / 1 front; adding blocks cuts
            # duration but doesn't multiply cost — the same crew just moves
            # block-to-block. The optimizer can then add fronts up to the
            # number of blocks to shorten the schedule.
            length = float(getattr(self.input, "built_up_area_sqm", 1000) or 1000)
            supplied_blocks = max(1, int(getattr(self.input, "number_of_blocks", 1) or 1))
            corridor_fronts = min(supplied_blocks, 6)
            baseline_fronts = 1  # Single front at baseline — cost doesn't multiply by blocks

        else:
            # Road: original corridor-length formula
            length = float(self.input.length_km or 0.0)
            if length <= 0:
                raise ValueError("project length must be greater than 0 km.")

            span = float(self.assumptions.get("work_front_span_km", 12) or 0)
            if span <= 0:
                raise ValueError(
                    "assumptions.yaml: work_front_span_km must be greater than 0 "
                    "(got %s). Fix it in assumptions.yaml." % span)

            corridor_fronts = max(2, min(8, int(math.ceil(length / 3.5))))
            baseline_fronts = max(1, min(corridor_fronts, int(math.ceil(length / span))))

        self.max_fronts = corridor_fronts
        self.corridor_fronts = corridor_fronts
        self.baseline_fronts = baseline_fronts
        self.max_multiplier = max(baseline_fronts, int(min(cap, corridor_fronts)))

        # Activities that can physically run at night (earthwork, compaction,
        # structures). Surfaces that must be laid in daylight / dry conditions
        # are excluded because night shifts would cost more and deliver less.
        raw_night = self.assumptions.get("night_work_eligible", None)
        night_ok = None if raw_night is None else set(raw_night)

        self.levels = {}
        for spec in self.activity_specs:
            key = spec["key"]
            base = int(spec.get("base_crews", 1))
            eligible = night_ok is None or key in night_ok

            combos: List[Tuple[int, float]] = []
            for mult in range(1, self.max_multiplier + 1):
                crews = base * mult
                combos.append((crews, 1.0))
                if eligible and mult > 1:
                    for s in SHIFT_LEVELS[1:]:
                        combos.append((crews, s))

            combos = [c for c in combos if c != (base, 1.0)]
            combos.sort(key=self._capacity)

            levels = [(base, 1.0)]
            for c in combos:
                if self._capacity(c) > self._capacity(levels[-1]) + 1e-9:
                    levels.append(c)
            self.levels[key] = levels

    @staticmethod
    def _capacity(level: Tuple[int, float]) -> float:
        """Nominal daily output of a resource level, in crew-shifts."""
        crews, shifts = level
        return crews * max(0.5, shifts)

    def _baseline_index_for(self, key: str) -> int:
        """Level index corresponding to the NORMAL organisation for `key`."""
        base = int(self.spec_by_key[key].get("base_crews", 1))
        target = base * self.baseline_fronts
        for i, (crews, shifts) in enumerate(self.levels.get(key) or [(base, 1.0)]):
            if crews == target and shifts == 1.0:
                return i
        return 0

    def _initial_state(self) -> Dict[str, Any]:
        """Start from the normal organisation, not from a single gang."""
        return {"levels": {s["key"]: self._baseline_index_for(s["key"])
                           for s in self.activity_specs},
                "stagger": False}

    def _state_from(self, base_state: Dict[str, Any],
                    deltas: List[Tuple[str, str, str]]) -> Dict[str, Any]:
        """Rebuild a solver state by replaying accepted moves over the baseline."""
        levels = dict(base_state["levels"])
        stagger = base_state["stagger"]
        for kind, key, _label in deltas:
            if kind == "level":
                levels[key] += 1
            else:
                stagger = True
        return {"levels": levels, "stagger": stagger}

    @staticmethod
    def _state_sig(state: Dict[str, Any]) -> Tuple:
        return (tuple(sorted(state["levels"].items())), bool(state["stagger"]))

    def _fronts_in(self, state: Dict[str, Any]) -> int:
        fronts = 1
        for spec in self.activity_specs:
            opts = self.levels[spec["key"]]
            idx = min(state["levels"].get(spec["key"], 0), len(opts) - 1)
            fronts = max(fronts, opts[idx][0] // int(spec.get("base_crews", 1)))
        return fronts

    def _lags_for(self, state: Dict[str, Any],
                  effort_days: Dict[str, int]) -> Dict[str, Dict[str, int]]:
        """
        Staggered handover: successors start a lead-time before their
        predecessor finishes (negative FS lag). The lead is a fraction of the
        predecessor's duration, capped so at least 2 days of the predecessor
        still remain when the successor takes over.
        """
        if not state["stagger"] or not self.stagger_chains:
            return {}
        fronts = self._fronts_in(state)
        if fronts <= 1:
            return {}
        lead_ratio = max(0.20, min(0.65, (fronts - 1) / fronts))

        lags: Dict[str, Dict[str, int]] = {}
        for pred_id, succ_id in self.stagger_chains:
            pred_dur = effort_days.get(pred_id, 10)
            lead = int(math.floor(pred_dur * lead_ratio))
            safe = max(0, min(lead, pred_dur - 2))
            if safe > 0:
                lags.setdefault(succ_id, {})[pred_id] = -safe
        return lags

    # ------------------------------------------------------------ construction
    def _build_activities(self, state: Dict[str, Any]) -> List[RoadActivity]:
        """
        Materialise the activity list for a solver state.

        `duration_days` is seeded with EFFORT days here; CPMEngine overwrites it
        with the CALENDAR duration for the early start date it assigns, which is
        why costing happens after the forward pass.
        """
        from .models import ActivityReference  # local to avoid cycle at import
        chosen = state["levels"]
        activities: List[RoadActivity] = []

        for spec in self.activity_specs:
            key = spec["key"]
            act_id = spec["id"]
            base_crews = int(spec.get("base_crews", 1))

            opts = self.levels.get(key) or [(base_crews, 1.0)]
            idx = min(max(chosen.get(key, 0), 0), len(opts) - 1)
            crews, shifts = opts[idx]
            work_fronts = max(1, crews // base_crews)

            if key not in self.quantities:
                ptype = str(getattr(self.input, "project_type", "road") or "road").lower()
                if ptype == "pipeline":
                    raise ValueError(
                        f"Quantity calculator produced no quantity for activity key "
                        f"'{key}'. A silent fallback would misprice the activity. "
                        f"Add an explicit formula in the quantity calculator or remove "
                        f"the activity from the guideline catalog."
                    )
                # Road/building catalogs are outside the pipeline audit scope.
                # Keep their legacy behaviour but record a visible warning
                # instead of silently pricing the activity as 100 units.
                self._quantity_fallbacks.add(f"{act_id}/{key}")
                qty_info = {
                    "quantity": 100.0,
                    "unit": spec.get("unit", "units"),
                    "formula": "Legacy 100-unit fallback (missing quantity formula)",
                }
            else:
                qty_info = self.quantities[key]
            qty = qty_info["quantity"]
            unit = qty_info.get("unit", spec.get("unit", "units"))

            effort_days, prod_rate, _, prod_unit = self.dur_calc.calculate_duration(
                activity_key=key, quantity=qty, crews=crews, shifts=shifts)

            ref = ActivityReference(
                morth_spec=spec.get("morth_spec", spec.get("spec_ref", "")),
                data_book_item=spec.get("morth_item", spec.get("data_book_item", "")),
                page_number=spec.get("page", 0),
                source_file=self._guidelines_filename,
                guideline_ref=self._guideline_ref_label,
                assumption_ref=self._assumptions_filename,
            )

            if crews > base_crews or shifts > 1.0 or work_fronts > 1:
                crew_desc = (f"{crews} Active Gang"
                             f"{'s' if crews > 1 else ''} "
                             f"({work_fronts} Front"
                             f"{'s' if work_fronts > 1 else ''})")
                if shifts > 1.0:
                    crew_desc += f", {shifts:.1f} Shifts/day"
                equip_desc = f"{key} equipment train ({crews} sets)"
            else:
                crew_desc = f"{base_crews} Standard Gang{'s' if base_crews > 1 else ''} (1 Front)"
                equip_desc = f"{key} standard equipment ({base_crews} sets)"

            activities.append(RoadActivity(
                activity_id=act_id,
                wbs_code=spec["wbs_code"],
                parent_wbs=spec["parent_wbs"],
                wbs_level=spec.get("wbs_level", 3),
                work_package=spec.get("work_package", ""),
                name=spec["name"],
                description=f"{spec['name']} as per {spec.get('morth_spec', '')}",
                category=spec.get("category", ""),
                unit=unit,
                quantity=qty,
                productivity_rate=prod_rate,
                productivity_unit=prod_unit,
                crew_description=crew_desc,
                equipment_description=equip_desc,
                duration_days=effort_days,
                effort_days=effort_days,
                activity_key=key,
                predecessors=[
                    predecessor for predecessor in spec.get("predecessors", [])
                    if predecessor in self.spec_by_id
                ],
                cost=ActivityCostBreakdown(),
                reference=ref,
                crews=crews,
                work_fronts=work_fronts,
                shifts=shifts,
                base_crews=base_crews,
                predecessor_lags={},
            ))

        if state["stagger"]:
            lags = self._lags_for(
                state, {a.activity_id: a.duration_days for a in activities})
            if lags:
                by_id = {a.activity_id: a for a in activities}
                for succ_id, pair in lags.items():
                    succ = by_id.get(succ_id)
                    if succ is None:
                        continue
                    succ.predecessor_lags = {**succ.predecessor_lags, **pair}
                    for pid in pair:
                        if pid not in succ.predecessors:
                            succ.predecessors.append(pid)

        return activities

    def _apply_weights(self, activities: List[RoadActivity]) -> None:
        """
        Two independent weights:

        * `weight_pct`            - share of MONEY, drives the expenditure curve.
        * `progress_weight_pct`   - share of PHYSICAL work content, measured in
          crew-days (quantity / productivity rate) and therefore unaffected by
          rates, crashing or front counts. Drives the progress S-curve.

        Keeping them separate is the whole point: a plan where "percent complete"
        and "percent of budget spent" are the same number cannot tell you whether
        you are ahead of schedule or merely over-spending.
        """
        tot_cost = sum(a.cost.total_amount for a in activities)
        for act in activities:
            act.weight_pct = round(
                (act.cost.total_amount / tot_cost) * 100.0, 2) if tot_cost > 0 else 0.0
        if activities and tot_cost > 0:
            # Put the rounding residual on the LARGEST money activity, never on
            # the last row: that row can be an out-of-scope zero-cost activity,
            # which would then report a phantom weight.
            diff = round(100.0 - sum(a.weight_pct for a in activities), 2)
            if diff:
                target = max(activities, key=lambda a: a.cost.total_amount)
                target.weight_pct = round(target.weight_pct + diff, 2)

        # `productivity_rate` is the single-crew rate, so quantity / rate is a
        # crew-day count that does not move when gangs are multiplied up.
        # A zero-quantity activity is genuinely out of scope (it has zero
        # duration), so it must carry ZERO physical weight, not a nominal day.
        work = [max(a.quantity, 0.0) / (a.productivity_rate or 1.0)
                for a in activities]
        tot_work = sum(work)
        for act, w in zip(activities, work):
            act.progress_weight_pct = round((w / tot_work) * 100.0, 2) if tot_work > 0 else 0.0
        if activities and tot_work > 0:
            diff = round(100.0 - sum(a.progress_weight_pct for a in activities), 2)
            if diff:
                target = max(zip(activities, work), key=lambda pair: pair[1])[0]
                target.progress_weight_pct = round(
                    target.progress_weight_pct + diff, 2)

    def _evaluate(self, state: Dict[str, Any]) -> Tuple[List[RoadActivity], int,
                                                        float, List[str], List[str]]:
        """Build -> CPM (calendar durations) -> cost -> weights.

        Returns (activities, duration_days, cost_inr, critical_path,
        all_zero_float_ids). The last item is the full critical SET, which is
        what the crashing search must consider - the reported path is only one
        chain out of any parallel ones.
        """
        acts = self._build_activities(state)
        cpm = CPMEngine(activities=acts, project_input=self.input,
                        calendar=self.calendar)
        dur, crit = cpm.run_cpm()

        # Activities that the stagger pulls forward to overlap their predecessor.
        # They add no gang, so without this they would carry no acceleration cost
        # at all and a duration cut bought purely by re-sequencing would price at
        # zero. `_lags_for` keys the negative FS lags by SUCCESSOR, so those are
        # exactly the overlapped activities.
        staggered_ids: set = set()
        if state["stagger"] and self.stagger_chains and self._fronts_in(state) > 1:
            staggered_ids = {succ for _pred, succ in self.stagger_chains}

        # Costs are priced against the CALENDAR duration: a front stays open,
        # staffed and hired for as many days as it appears on the calendar.
        for act in acts:
            # Escalation is measured to the MID-POINT of this activity's own
            # spend, so work finished early in the programme does not carry
            # years of inflation it was never exposed to.
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
                staggered=act.activity_id in staggered_ids,
                escalation_years=esc_years,
            )
        self._apply_weights(acts)

        total = sum(a.cost.total_amount for a in acts)
        zero_float = [a.activity_id for a in acts if a.is_critical]
        return acts, dur, total, crit, zero_float

    # ------------------------------------------------------------------- solve
    @staticmethod
    def _log_entry(iteration: int, description: str, fronts: int,
                   duration_days: int, cost_crore: float,
                   critical_path: List[str], req: int,
                   budget: float) -> Dict[str, Any]:
        dur_ok = duration_days <= req
        cost_ok = cost_crore <= budget
        feasible = dur_ok and cost_ok
        if feasible:
            status = "PASS (FEASIBLE)" if iteration == 0 else "PASS (OPTIMAL)"
        elif not dur_ok:
            status = "FAIL (DURATION)"
        else:
            status = "FAIL (BUDGET)"
        return {
            "iteration": iteration,
            "description": description,
            "work_fronts": fronts,
            "duration_days": duration_days,
            "cost_crore": cost_crore,
            "duration_ok": "PASS" if dur_ok else "FAIL",
            "cost_ok": "PASS" if cost_ok else "FAIL",
            "is_feasible": feasible,
            "status": status,
            "critical_path": critical_path,
        }

    @staticmethod
    def _move_label(act_id: str, name: str, old: Tuple[int, float],
                    new: Tuple[int, float]) -> str:
        parts = []
        if new[0] != old[0]:
            parts.append(f"{old[0]}->{new[0]} gangs")
        if new[1] != old[1]:
            parts.append(f"{old[1]:.1f}->{new[1]:.1f} shifts")
        return f"{act_id} {name}: " + ", ".join(parts)

    def solve(self) -> ProjectScheduleResult:
        """
        Runs the baseline, then least-cost crashes onto the critical path until
        the duration target is met or no affordable move remains.
        """
        req_dur = self.input.duration_days
        max_budget = self.input.budget_crore

        self._build_crash_levels()

        # ---------------------------------------------------------- baseline
        base_state = self._initial_state()
        base_acts, base_dur, base_cost_inr, base_crit, _ = self._evaluate(base_state)
        base_cost_cr = round(base_cost_inr / 10000000.0, 2)

        for a in base_acts:
            a.base_duration_days = a.duration_days
            a.base_cost = a.cost.total_amount
            a.base_crews = a.crews
        base_activities = copy.deepcopy(base_acts)

        fronts_word = "Front" if self.baseline_fronts == 1 else "Fronts"
        baseline_label = (f"Baseline (Normal Resources, {self.baseline_fronts} "
                          f"{fronts_word}, Sequential FS)")
        log: List[Dict[str, Any]] = [self._log_entry(
            0, baseline_label, self.baseline_fronts,
            base_dur, base_cost_cr, base_crit, req_dur, max_budget)]

        if base_dur <= req_dur and base_cost_cr <= max_budget:
            return self._assemble(
                base_acts, base_dur, base_cost_cr, base_crit, log,
                base_activities, base_dur, base_cost_cr, base_crit,
                is_feasible=True,
                msg=("Baseline schedule satisfies duration and budget "
                     "constraints without crashing."),
                violations=[], gates=self._gates(base_dur, base_cost_cr),
                note="")

        # ------------------------------------------------- least-cost crashing
        state = self._initial_state()
        cur_acts, cur_dur, cur_cost, cur_crit, cur_critical = self._evaluate(state)
        cur_cost_cr = round(cur_cost / 10000000.0, 2)

        # If the baseline already busts the budget, no candidate could ever be
        # rejected for cost. Switch the filter off rather than silently refusing
        # to search - otherwise the report quotes a "shortest achievable
        # duration" it never actually looked for.
        force = cur_cost_cr > max_budget

        # Accepted moves, so the result can replay or back them out again.
        deltas: List[Tuple[str, str, str]] = []
        visited = {self._state_sig(state)}
        stop_reason = ""
        for move_no in range(1, 301):
            if cur_dur <= req_dur:
                break

            candidates: List[Dict[str, Any]] = []
            if not state["stagger"] and self.stagger_chains:
                candidates.append({
                    "label": "Staggered Handover (FS lead on critical chain)",
                    "kind": "stagger", "key": "",
                    "trial": {"levels": dict(state["levels"]), "stagger": True}})

            for aid in cur_critical:
                spec = self.spec_by_id.get(aid)
                if spec is None:
                    continue
                key = spec["key"]
                if state["levels"][key] >= len(self.levels[key]) - 1:
                    continue
                trial = {"levels": dict(state["levels"]),
                         "stagger": state["stagger"]}
                trial["levels"][key] = state["levels"][key] + 1
                candidates.append({
                    "label": self._move_label(
                        aid, spec["name"],
                        self.levels[key][state["levels"][key]],
                        self.levels[key][trial["levels"][key]]),
                    "kind": "level", "key": key, "trial": trial})

            if not candidates:
                stop_reason = ("the resource ceiling has been reached - no critical "
                               "activity has a higher resource level left, so "
                               "further crashing is impossible")
                break

            best = None
            n_saving = 0
            n_over_budget = 0
            for cand in candidates:
                trial = cand["trial"]
                sig = self._state_sig(trial)
                if sig in visited:
                    continue
                visited.add(sig)
                t_acts, t_dur, t_cost, t_crit, t_critical = self._evaluate(trial)
                saved = cur_dur - t_dur
                if saved <= 0:
                    continue
                n_saving += 1
                if (not force) and round(t_cost / 10000000.0, 2) > max_budget:
                    n_over_budget += 1
                    continue
                ratio = (t_cost - cur_cost) / saved
                if best is None or ratio < best[0]:
                    best = (ratio, cand, t_acts, t_dur, t_cost, t_crit, t_critical)

            if best is None:
                if n_saving == 0:
                    stop_reason = ("the resource ceiling has been reached - no "
                                   "remaining move shortens the schedule")
                elif n_over_budget > 0:
                    stop_reason = (
                        f"every remaining move that shortens the schedule "
                        f"({n_over_budget} candidate(s)) would push cost above "
                        f"the \u20b9{max_budget:,.2f} Cr budget")
                else:
                    stop_reason = "no candidate move shortens the schedule"
                break

            _, cand, t_acts, t_dur, t_cost, t_crit, t_critical = best
            state = cand["trial"]
            deltas.append((cand["kind"], cand["key"], cand["label"]))
            cur_acts, cur_dur, cur_cost, cur_crit = t_acts, t_dur, t_cost, t_crit
            cur_critical = t_critical
            cur_cost_cr = round(cur_cost / 10000000.0, 2)
            log.append(self._log_entry(
                move_no, cand["label"], self._fronts_in(state), cur_dur,
                cur_cost_cr, cur_crit, req_dur, max_budget))

        # ------------------------------------------ back off unneeded moves
        # The loop stops at the FIRST state that meets the target, so it can
        # leave gangs working the job longer than the deadline needs. Back off
        # any move the target can do without - but only while cost does not
        # RISE: a longer front can be dearer (it sits across the monsoon, and
        # time-driven site/plant cost runs on), in which case the "extra"
        # crashing is already the cheaper plan and nothing should be undone.
        # Among the removals that qualify, take the one landing closest to the
        # deadline, since the deadline is a maximum rather than a score.
        backed_off = 0
        if cur_dur <= req_dur and cur_cost_cr <= max_budget and deltas:
            accepted_cost = cur_cost
            passes = 0
            while passes <= len(deltas) + 1:
                passes += 1
                pick = None
                for i in range(len(deltas)):
                    trial_deltas = deltas[:i] + deltas[i + 1:]
                    _, t_dur, t_cost, _, _ = self._evaluate(
                        self._state_from(base_state, trial_deltas))
                    if t_dur > req_dur:
                        continue
                    if round(t_cost / 10000000.0, 2) > max_budget:
                        continue
                    if t_cost > accepted_cost + 0.01:
                        continue
                    if pick is None or t_dur > pick[0]:
                        pick = (t_dur, trial_deltas, t_cost)
                if pick is None:
                    break
                _dur, trial_deltas, t_cost = pick
                deltas = trial_deltas
                accepted_cost = t_cost
                backed_off += 1

            if backed_off:
                # Replay the surviving moves to rebuild state, activities and log.
                log = [log[0]]
                st = base_state
                for n, (kind, key, label) in enumerate(deltas, start=1):
                    st = self._state_from(st, [(kind, key, label)])
                    cur_acts, cur_dur, cur_cost, cur_crit, cur_critical = \
                        self._evaluate(st)
                    log.append(self._log_entry(
                        n, label, self._fronts_in(st), cur_dur,
                        round(cur_cost / 10000000.0, 2), cur_crit,
                        req_dur, max_budget))
                state = st
                cur_cost_cr = round(cur_cost / 10000000.0, 2)

        # -------------------------------------------------------- verdict
        is_feasible = (cur_dur <= req_dur and cur_cost_cr <= max_budget)
        gates = self._gates(cur_dur, cur_cost_cr)
        violations: List[str] = []

        note = stop_reason
        if backed_off:
            note = (f"{backed_off} crashing step(s) backed off after the target "
                    f"was met - the plan does not need them")

        if is_feasible:
            msg = (f"Optimal feasible schedule found: {cur_dur} days "
                   f"(target: {req_dur} days) at \u20b9{cur_cost_cr:,.2f} Cr "
                   f"(budget: \u20b9{max_budget:,.2f} Cr).")
            if backed_off:
                msg += f" {note[0].upper() + note[1:]}."
            headroom = req_dur - cur_dur
            if headroom > 0:
                msg += (f" Finishes {headroom} day(s) before the deadline: the "
                        f"deadline is a ceiling, and every move that would use "
                        f"more of it up also costs more than it buys.")
        elif cur_dur <= req_dur:
            msg = "Duration target met, but no schedule fits the budget."
            violations.append(
                f"Duration target met ({cur_dur} <= {req_dur} days) but cost "
                f"\u20b9{cur_cost_cr:,.2f} Cr exceeds budget "
                f"\u20b9{max_budget:,.2f} Cr.")
        else:
            msg = "No feasible schedule found within the given budget and duration."
            violations.append(
                f"Shortest schedule found: {cur_dur} days "
                f"(required: {req_dur} days).")
            violations.append(
                f"Estimated cost \u20b9{cur_cost_cr:,.2f} Cr "
                f"(budget \u20b9{max_budget:,.2f} Cr).")
        if stop_reason and not is_feasible:
            violations.append(f"Search stopped: {stop_reason}.")

        return self._assemble(
            cur_acts, cur_dur, cur_cost_cr, cur_crit, log,
            base_activities, base_dur, base_cost_cr, base_crit,
            is_feasible=is_feasible, msg=msg,
            violations=violations, gates=gates, note=note)

    # ------------------------------------------------------------------ output
    def _gates(self, duration_days: int, cost_crore: float) -> List[str]:
        """
        Reality checks from assumptions.yaml. These are deliberately warnings
        rather than hard failures: a plan that disagrees with an industry
        benchmark by a wide margin is not necessarily wrong, but it is never
        something to report silently.
        """
        warnings: List[str] = []

        # 1. Cost per km against the published per-unit benchmark.
        benchmarks = self.assumptions.get("per_unit_benchmarks_crore", {}) or {}
        lanes = 0
        try:
            from .models import parse_lanes_from_title
            lanes = parse_lanes_from_title(self.input.title, default_lanes=0)
        except Exception:
            lanes = 0
        bench_key = ("highway_4lane_greenfield_plains_per_km" if lanes >= 4
                     else "highway_2lane_per_km" if lanes == 2
                     else "rural_road_per_km" if lanes in (1,) else "")
        bench = benchmarks.get(bench_key)
        length = float(self.input.length_km or 0.0)
        if bench and length > 0:
            expected = float(bench) * length
            actual = cost_crore
            tol = float((self.assumptions.get("benchmark_tolerances", {}) or {})
                        .get("estimate_vs_benchmark_pct", 25.0))
            dev = (actual - expected) / expected * 100.0
            if abs(dev) > tol:
                warnings.append(
                    f"Estimate \u20b9{actual:,.2f} Cr is {dev:+.0f}% vs the "
                    f"\u20b9{expected:,.2f} Cr benchmark "
                    f"({bench_key} = \u20b9{float(bench):g} Cr/km x "
                    f"{length:g} km), tolerance +/-{tol:.0f}%.")

        # 2. Bromilow time-cost curve: T = K x C^B.
        brom = self.assumptions.get("bromilow", {}) or {}
        kb = brom.get("highway") or brom.get("generic") or {}
        k = float(kb.get("K", 0) or 0)
        b = float(kb.get("B", 0) or 0)
        if k and b and cost_crore > 0:
            predicted = k * (cost_crore ** b)
            tol = float(brom.get("tolerance_pct", 25.0))
            low, high = predicted * (1 - tol / 100.0), predicted * (1 + tol / 100.0)
            if not (low <= duration_days <= high):
                side = "below" if duration_days < low else "above"
                warnings.append(
                    f"Planned {duration_days} days is {side} the Bromilow band "
                    f"{low:.0f}-{high:.0f} days (K={k:g}, B={b:g}, "
                    f"C=\u20b9{cost_crore:,.2f} Cr, +/-{tol:.0f}%).")

        return warnings

    def _monsoon_exposure(self, activities: List[RoadActivity]) -> Optional[str]:
        """How much of the resulting plan actually lands in the wet season."""
        cal = self.assumptions.get("calendar", {}) or {}
        monsoon = set(cal.get("monsoon_months") or [])
        if not monsoon:
            return None
        start = self.input.start_date
        total = 0
        wet = 0
        for act in activities:
            for d in range(act.es, act.ef):
                total += 1
                if (start + timedelta(days=d)).month in monsoon:
                    wet += 1
        if total == 0:
            return None
        factor = float(cal.get("monsoon_productivity_factor", 1.0) or 1.0)
        return (f"{wet / total * 100.0:.0f}% of scheduled activity-days fall in "
                f"the monsoon months {sorted(monsoon)} at {factor:.0%} "
                f"productivity; the calendar already prices this in.")

    @staticmethod
    def _parallel_critical_note(activities: List[RoadActivity],
                                critical_path: List[str]) -> Optional[str]:
        """
        The report shows ONE critical chain, but zero float can belong to a
        parallel chain that also determines the finish date. Say so explicitly -
        otherwise "the critical path" looks like the complete answer when it is
        only one of several binding sequences.
        """
        on_path = set(critical_path)
        others = [a.activity_id for a in activities
                  if a.is_critical and a.activity_id not in on_path]
        if not others:
            return None
        return (f"{len(others)} further zero-float activit"
                f"{'ies' if len(others) != 1 else 'y'} lie"
                f"{'' if len(others) != 1 else 's'} on PARALLEL critical "
                f"chain(s) in addition to the reported path: {', '.join(others)}. "
                f"All of them must be compressed to shorten the finish date.")

    def _assemble(self, acts, final_dur, final_cost_cr, final_crit, log,
                  base_activities, base_dur, base_cost_cr, base_crit,
                  is_feasible, msg, violations, gates, note=""):
        base_by_id = {a.activity_id: a for a in base_activities}
        resources_summary = []
        fronts = 1
        for act in acts:
            fronts = max(fronts, act.work_fronts)
            base_act = base_by_id.get(act.activity_id)
            if base_act is None:
                act.resources_added_desc = "-"
                continue
            act.base_duration_days = base_act.duration_days
            act.base_cost = base_act.cost.total_amount
            act.base_crews = base_act.crews
            # Only rows where resources were actually added. An activity whose
            # duration merely moved because the calendar re-placed it (e.g. it
            # no longer sits across the monsoon) did not have gangs added and
            # must not appear in this table.
            extra_g = max(0, act.crews - base_act.crews)
            escalated = act.work_fronts > base_act.work_fronts or act.shifts > base_act.shifts
            if extra_g > 0 or escalated:
                shifts_desc = f", {act.shifts:.1f} shifts" if act.shifts > 1.0 else ""
                act.resources_added_desc = (
                    f"{act.crews} gangs across {act.work_fronts} front"
                    f"{'s' if act.work_fronts > 1 else ''} "
                    f"(+{extra_g} extra gang{'s' if extra_g != 1 else ''}{shifts_desc})")
                resources_summary.append({
                    "activity_id": act.activity_id,
                    "activity_name": act.name,
                    "baseline_crews": base_act.crews,
                    "optimized_crews": act.crews,
                    "work_fronts": act.work_fronts,
                    "shifts": act.shifts,
                    "days_saved": base_act.duration_days - act.duration_days,
                    "cost_increase_lakhs": round(
                        (act.cost.total_amount - base_act.cost.total_amount) / 100000.0, 2),
                })
            else:
                act.resources_added_desc = "-"

        prog_calc = ProgressCalculator(activities=acts,
                                       project_input=self.input,
                                       calculated_duration_days=final_dur)
        monthly_items, _ = prog_calc.generate_monthly_schedule()

        max_budget = self.input.budget_crore
        cost_var_cr = round(final_cost_cr - max_budget, 2)
        cost_var_pct = round((cost_var_cr / max_budget) * 100.0, 2) if max_budget > 0 else 0.0

        warnings = list(violations)
        gates = list(gates)
        if self._quantity_fallbacks:
            keys = sorted(self._quantity_fallbacks)
            shown = ", ".join(keys[:5])
            more = f" (+{len(keys) - 5} more)" if len(keys) > 5 else ""
            gates.append(
                f"{len(keys)} activity key(s) have no quantity formula for the "
                f"current catalog: {shown}{more}. The legacy 100-unit fallback "
                f"was used, so this project type's quantities are placeholders."
            )
        mono = self._monsoon_exposure(acts)
        if mono:
            gates.append(mono)
        parallel = self._parallel_critical_note(acts, final_crit)
        if parallel:
            gates.append(parallel)

        return ProjectScheduleResult(
            project_input=self.input,
            activities=acts,
            total_calculated_cost_crore=final_cost_cr,
            cost_variance_crore=cost_var_cr,
            cost_variance_pct=cost_var_pct,
            calculated_duration_days=final_dur,
            schedule_variance_days=final_dur - self.input.duration_days,
            critical_path=final_crit,
            monthly_progress=monthly_items,
            warnings=warnings,
            validation_warnings=gates,
            pipeline_scope=self.qty_calc.scheme_scope(),
            accel_site_establishment_rate=self.cost_calc.site_establishment_per_front_day,
            accel_resident_staff_rate=self.cost_calc.resident_staff_per_front_day,
            is_feasible=is_feasible,
            feasibility_message=msg,
            baseline_duration_days=base_dur,
            baseline_cost_crore=base_cost_cr,
            baseline_critical_path=base_crit,
            baseline_activities=copy.deepcopy(base_activities),
            optimized_duration_days=final_dur,
            optimized_cost_crore=final_cost_cr,
            duration_saved_days=max(0, base_dur - final_dur),
            # Acceleration cost measures what the SCHEDULE DECISION cost, so the
            # price-level effect is excluded: a shorter programme spends its money
            # earlier and therefore carries LESS escalation, and netting that
            # against the surcharge would understate the real cost of compression
            # (and could even make it vanish). Escalation is reported on its own
            # line on the cost sheet and in the sheet-11 bridge.
            additional_cost_crore=max(0.0, round(
                (sum(a.cost.total_amount - a.cost.escalation_amount
                     for a in acts) / 10000000.0)
                - (sum(a.cost.total_amount - a.cost.escalation_amount
                       for a in base_activities) / 10000000.0),
                2)),
            # Acceleration can only ever COST money - it can never make a project
            # cheaper, so this figure is floored at zero. A negative delta is not
            # a saving: it is a calendar artefact. When the optimizer beats the
            # duration target WITHOUT adding resources (e.g. a staggered FS lead
            # that re-times the chain), the resource-based planned cost is
            # unchanged (it is quantity x rate, with no duration term) while the
            # TIME-driven terms - extra-front site establishment, resident staff,
            # plant-hire billing uplift and the gang-day fuel/plant/labour wastage
            # - shrink with the shortened calendar spans. The honest acceleration
            # spend in that case is zero, not negative, so it is reported as 0.
            work_fronts=fronts,
            baseline_work_fronts=self.baseline_fronts,
            resources_added_summary=resources_summary,
            optimization_iterations=log,
            optimization_note=note or "",
            assumptions_snapshot=self.assumptions,
        )
