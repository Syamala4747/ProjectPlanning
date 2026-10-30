"""
Data loader module for reading road_guidelines.json and road_assumptions.yaml.
"""
import os
import json
import copy
import yaml
from typing import Dict, Any, List

class DataLoader:
    def __init__(self, guidelines_path: str = "road_guidelines.json", assumptions_path: str = "road_assumptions.yaml"):
        self.guidelines_path = guidelines_path
        self.assumptions_path = assumptions_path
        self.guidelines: Dict[str, Any] = {}
        self.assumptions: Dict[str, Any] = {}
        # Data-integrity problems found at load time. Surfaced by main.py so a
        # broken assumption is never silently accepted.
        self.warnings: List[str] = []

    def load_all(self, terrain: str = "plain") -> bool:
        """Loads guidelines and assumptions files, then validates cross-references."""
        self.load_guidelines()
        self.load_assumptions()
        self.select_terrain(terrain)
        self.warnings = self.validate()
        return True

    def select_terrain(self, terrain: str = "plain") -> str:
        """Select the activity catalog used by all downstream calculators."""
        selected = str(terrain or "plain").strip().lower()
        if selected not in {"plain", "hilly"}:
            raise ValueError("terrain must be either 'plain' or 'hilly'.")

        activity_master = (
            self.guidelines.get("activity_master")
            or self.guidelines.get("data", {}).get("activity_master")
            or self.assumptions.get("activity_master")
            or []
        )

        if selected == "hilly":
            hilly = self.assumptions.get("hilly_terrain_standards") or {}
            catalog = hilly.get("activity_master") or []
            if not catalog:
                raise ValueError(
                    "Hilly terrain was requested, but road_assumptions.yaml has no "
                    "hilly_terrain_standards.activity_master catalog."
                )
            plain_by_key = {
                item.get("key"): item
                for item in activity_master
            }
            normalized_catalog = []
            incomplete = []
            for item in catalog:
                key = item.get("key", item.get("id", "?"))
                normalized = copy.deepcopy(plain_by_key.get(key, {}))
                normalized.update(copy.deepcopy(item))
                if "hilly_productivity_rate" in item:
                    normalized["productivity_rate"] = item.get("hilly_productivity_rate")
                if normalized.get("productivity_rate") is None:
                    incomplete.append(str(key))
                normalized_catalog.append(normalized)
            if incomplete:
                raise ValueError(
                    "Hilly terrain cannot be scheduled yet: productivity_rate is "
                    "missing for " + ", ".join(incomplete) + ". Add the exact "
                    "Volume-II-derived productivity for these activities; plain "
                    "terrain values are never copied into hilly terrain."
                )
            self.assumptions["activity_master"] = normalized_catalog
        else:
            self.assumptions["activity_master"] = copy.deepcopy(activity_master)

        self.assumptions["selected_terrain"] = selected
        return selected

    def validate(self) -> List[str]:
        """
        Cross-checks road_assumptions.yaml against road_guidelines.json.

        Every activity must carry a resolvable MoRTH/SDB item reference and a
        published unit rate: `guideline_ref` is printed on the report, and a
        missing rate would otherwise be priced at an arbitrary default.
        """
        problems: List[str] = []

        acts = (((self.guidelines or {}).get("data") or {})
                .get("activities") or [])
        known_ids = {str(x.get("activity_id", "")).strip() for x in acts}
        known_items = {str(x.get("item_number", "")).strip() for x in acts}

        rates = self.assumptions.get("direct_unit_rates_inr", {}) or {}
        unresolvable = []
        unpriced = []
        for spec in self.assumptions.get("activity_master", []) or []:
            key = spec.get("key", "?")
            item = str(spec.get("morth_item", "")).strip()
            if item and item not in known_ids and item not in known_items:
                unresolvable.append(f"{spec.get('id', key)} -> '{item}'")
            if key not in rates:
                unpriced.append(key)

        if unresolvable:
            problems.append(
                f"{len(unresolvable)} of {len(self.assumptions.get('activity_master', []))} "
                f"activities cite a MoRTH/SDB item that does not exist in "
                f"road_guidelines.json (the printed guideline_ref will be a dangling "
                f"citation): " + "; ".join(unresolvable))
        if unpriced:
            problems.append(
                f"direct_unit_rates_inr has no rate for: {', '.join(unpriced)} "
                f"- costing will fail until these are supplied.")

        # --- dead-configuration guard -------------------------------------
        # `productivity:` (the top-level block) is NOT read by anything: the
        # cost engine (cost_calculator._resource_planned_rate) and the duration
        # engine (duration_calculator.DurationCalculator) BOTH read
        # `activity_master[].productivity_rate`. Editing the top-level block
        # therefore changes nothing while looking authoritative — which is
        # exactly the trap that left excavation priced on a stale 800 cum/day
        # after the norm block had been corrected. Cross-check the two and say
        # so loudly when they disagree.
        prod_block = self.assumptions.get("productivity") or {}
        if prod_block:
            master_rates = {
                s.get("key"): float(s.get("productivity_rate", 0.0) or 0.0)
                for s in self.assumptions.get("activity_master", []) or []
            }
            blocked = {float(v) for v in prod_block.values()
                       if isinstance(v, (int, float))}
            unshared = [k for k, v in master_rates.items() if v and v not in blocked]
            problems.append(
                "the top-level `productivity:` block is REFERENCE ONLY - no code "
                "reads it. Cost AND duration both use "
                "activity_master[].productivity_rate. Do not edit the norm block "
                "to change a rate; edit the activity_master entry."
                + (f" {len(unshared)} activity_master rate(s) are not mirrored "
                   f"there, so the block is already out of step." if unshared else ""))

        # --- unused assumption blocks -------------------------------------
        # Blocks that are defined but never read look live to anyone reviewing
        # the YAML. Name them explicitly instead of letting them mislead.
        reserved = ["project_overhead_crew", "vehicle_transport",
                    "subcontractor_defaults", "overhead_weights"]
        present = [b for b in reserved if b in self.assumptions]
        if present:
            problems.append(
                f"assumption block(s) {', '.join(present)} are defined but NEVER "
                f"read by the cost model. Site overhead is already carried by "
                f"sdb_rate_oncost (OH+CP inside the all-in rate), so adding "
                f"project_overhead_crew on top would double count. Treat them "
                f"as documentation until they are wired in deliberately.")

        # --- source-citation guard --------------------------------------------
        # An estimate that cannot name its source cannot defend its numbers.
        # Escalation is the easiest place to invent a number that merely looks
        # plausible, so it is held to the strictest rule: it may only run when a
        # source AND the contract's own clause are recorded.
        esc = self.assumptions.get("escalation") or {}
        if esc.get("enabled") and not (esc.get("source") or "").strip():
            problems.append(
                "escalation.enabled is true but escalation.source is empty - the "
                "annual rates are PLACEHOLDERS, not published figures. Record the "
                "contract's escalation clause (MoRTH/CPWD circular, index and "
                "weights) and its source, or set enabled: false and price flat.")
        if esc.get("enabled") and not (esc.get("contract_escalation_clause") or "").strip():
            problems.append(
                "escalation.contract_escalation_clause is empty - escalation is "
                "governed by the contract clause, so without it the model is "
                "assuming a rate rather than applying one.")

        register = self.assumptions.get("sources") or {}
        if not register:
            problems.append(
                "no `sources:` register in road_assumptions.yaml - no rate in this "
                "file can currently be traced to an official publication.")
        for block, label in (("material_rates_inr", "CPWD DSR / GeM / state PWD"),
                             ("equipment_hire_inr_per_day", "CPWD DSR equipment schedule")):
            if block in self.assumptions and not (esc.get("source") or "").strip():
                pass
            if block in self.assumptions and f"{block}:" not in str(register):
                problems.append(
                    f"`{block}` is not listed in the sources register (expected "
                    f"basis: {label}). Add the edition and base date so each rate "
                    f"is citable.")

        # Planning Variation is now DERIVED at run time from the first-principles
        # resource build-up (recipe + crew + plant + fuel at the natural basis).
        # The old hand tables are no longer read by the cost engine; if they are
        # still present, flag them so a user migrating to the new model is nudged
        # to delete them.
        planned_rates = self.assumptions.get("planned_unit_rates_inr", {}) or {}
        legacy_variations = self.assumptions.get("planning_rate_variations", {}) or {}
        if planned_rates or legacy_variations:
            leftovers = []
            if planned_rates:
                leftovers.append("planned_unit_rates_inr")
            if legacy_variations:
                leftovers.append("planning_rate_variations")
            problems.append(
                f"{' and '.join(leftovers)} is/are present in road_assumptions.yaml but "
                f"the planner no longer reads them: the planned rate and its "
                f"variation are derived from the resource build-up (material "
                f"recipes, crew day rates, plant hire, fuel lph). Remove the "
                f"stale tables to keep the file honest.")

        return problems

    def load_assumptions(self) -> Dict[str, Any]:
        """Loads road_assumptions.yaml."""
        if not os.path.exists(self.assumptions_path):
            raise FileNotFoundError(f"Assumptions file not found at: {self.assumptions_path}")
        
        with open(self.assumptions_path, "r", encoding="utf-8") as f:
            self.assumptions = yaml.safe_load(f)

        acts = (
            self.guidelines.get("activity_master")
            or (self.guidelines.get("data") or {}).get("activity_master")
        )
        if acts and "activity_master" not in self.assumptions:
            self.assumptions["activity_master"] = copy.deepcopy(acts)

        return self.assumptions

    def load_guidelines(self) -> Dict[str, Any]:
        """Loads road_guidelines.json."""
        if not os.path.exists(self.guidelines_path):
            raise FileNotFoundError(f"Guidelines file not found at: {self.guidelines_path}")
        
        with open(self.guidelines_path, "r", encoding="utf-8") as f:
            self.guidelines = json.load(f)

        acts = (
            self.guidelines.get("activity_master")
            or (self.guidelines.get("data") or {}).get("activity_master")
        )
        if acts:
            self.guidelines["activity_master"] = acts
            if "activity_master" not in self.assumptions:
                self.assumptions["activity_master"] = copy.deepcopy(acts)

        return self.guidelines

    def get_overhead_percentage(self, project_budget_crore: float) -> float:
        """
        Determines overhead percentage based on project size from road_assumptions.yaml statutory_and_overheads.
        """
        statutory = self.assumptions.get("statutory_and_overheads", {})
        slabs = statutory.get("overhead_slabs", [])
        for slab in slabs:
            if project_budget_crore <= slab.get("max_budget_cr", float("inf")):
                return float(slab.get("overhead_pct", 8.0))
        return float(statutory.get("overhead_pct", 8.0))

    def get_profit_percentage(self) -> float:
        """Returns contractor profit percentage from road_assumptions.yaml."""
        statutory = self.assumptions.get("statutory_and_overheads", {})
        return float(statutory.get("profit_pct", 10.0))
