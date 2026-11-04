"""
Data models and type definitions for the Construction Project Scheduler.
Supports road, pipeline, and building project types.
"""
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from datetime import date

import re

# Valid project types
PROJECT_TYPES = {"road", "pipeline", "building"}

def parse_lanes_from_title(title: str, default_lanes: int = 4) -> int:
    """
    Extracts the number of highway lanes from the project title.
    Supports numbers (e.g. '2-lane', '6 lane', '7lane', '4L', '10km_2lane.xlsx')
    and English words (e.g. 'two-lane', 'six lane', 'single lane', 'dual 3-lane').
    """
    if not title:
        return default_lanes

    t = str(title).strip().lower()

    word_to_num = {
        "single": 1, "one": 1, "two": 2, "three": 3, "four": 4,
        "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
        "ten": 10, "twelve": 12, "fourteen": 14, "sixteen": 16
    }

    # 1. Dual X-lane patterns (e.g. 'dual 3-lane' -> 6 lanes, 'dual 2-lane' -> 4 lanes)
    dual_match = re.search(r"dual\s+(\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s*[-_ ]?\s*lanes?\b", t)
    if dual_match:
        val_str = dual_match.group(1)
        per_carriageway = int(val_str) if val_str.isdigit() else word_to_num.get(val_str, 2)
        return max(1, per_carriageway * 2)

    # 2. Number + lane/lanes (e.g. '2-lane', '2 lane', '2lane', '6-lanes', '7_lane', '10km_2lane')
    num_match = re.search(r"(\d+)\s*[-_ ]?\s*lanes?\b", t)
    if num_match:
        return max(1, int(num_match.group(1)))

    # 3. Compact file/slug format like '10km_2lane' or '2lane'
    compact_match = re.search(r"(\d+)lane", t)
    if compact_match:
        return max(1, int(compact_match.group(1)))

    # 4. Short 'NL' notation like '4L', '6L', '2L' (e.g. '25km 4L highway' or '6l expressway')
    short_match = re.search(r"\b(\d+)\s*l\b", t)
    if short_match:
        val = int(short_match.group(1))
        if 1 <= val <= 20:
            return val

    # 5. English word numbers + lane (e.g. 'two lane', 'six-lane', 'single lane')
    words_pattern = r"\b(single|one|two|three|four|five|six|seven|eight|nine|ten|twelve|fourteen|sixteen)\s*[-_ ]?\s*lanes?\b"
    word_match = re.search(words_pattern, t)
    if word_match:
        w = word_match.group(1)
        return max(1, word_to_num.get(w, default_lanes))

    return default_lanes


def parse_length_from_title(title: str, default_length: float = 0.0) -> float:
    """Extracts project length in km from title if present (e.g. '10 KM 2-Lane' -> 10.0, '10km_2lane' -> 10.0)."""
    if not title:
        return default_length
    match = re.search(r"(\d+(\.\d+)?)\s*[-_ ]?\s*km\b", str(title), re.IGNORECASE)
    if match:
        return float(match.group(1))
    return default_length


def parse_terrain_from_title(title: str, default_terrain: str = "plain") -> str:
    """Detect an explicit hilly terrain marker; plain is the safe default."""
    if re.search(r"\b(hilly|hill|mountain|mountainous)\b", str(title or ""), re.IGNORECASE):
        return "hilly"
    return default_terrain


@dataclass
class ProjectInput:
    title: str
    length_km: float
    duration_days: int
    budget_crore: float
    start_date: date = field(default_factory=lambda: date.today())
    lanes: int = 4
    terrain: str = ""
    carriageway_width_m: float = 0.0
    formation_width_m: float = 0.0
    gst_pct: float = 18.0
    # Project type: "road", "pipeline", or "building"
    project_type: str = "road"

    # --- Pipeline-specific fields ---
    pipe_diameter_mm: float = 0.0       # e.g. 600 for DN600
    pipe_material: str = "DI"           # DI / HDPE / MS / RCC
    trench_depth_m: float = 2.0
    trench_width_m: float = 0.0         # 0 = auto-derive from diameter
    installation_method: str = "open_trench"  # open_trench / trenchless
    number_of_work_fronts: int = 1      # user-supplied parallelism
    pipeline_type: str = "water"        # water / sewerage / gas

    # --- Scheme-level pipeline fields (multi-village water supply) ---
    # These switch on the facility packages in the 89-activity catalogue:
    # intake, WTP, CWR, OHT/ESR, FHTC and allied E&M.  Zero means "not in
    # scope", which keeps a plain transmission-main run as a corridor-only
    # estimate.  `--villages` derives sensible values for the rest.
    num_villages: int = 0               # villages served by the scheme
    num_oht: int = 0                    # overhead service reservoirs
    wtp_capacity_mld: float = 0.0       # water treatment plant capacity, MLD
    num_ibps: int = 0                   # intake / booster pumping stations
    num_fhtc: int = 0                   # functional household tap connections

    # Declared crossings. These override the configured defaults in
    # pipeline_assumptions.yaml, so a scheme variant does not need a second
    # assumptions file just to declare its crossings.
    number_of_road_crossings: int = 0
    number_of_highway_crossings: int = 0
    number_of_rail_crossings: int = 0
    number_of_hdd_crossings: int = 0
    number_of_river_crossings: int = 0

    # --- Building-specific fields ---
    built_up_area_sqm: float = 0.0      # total built-up area across all floors (sqft input converted before passing)
    number_of_floors: int = 0           # G+N => total floors = N+1; e.g. G+3 => 4
    building_type: str = "residential"  # residential / commercial
    construction_type: str = "rcc"      # rcc / load_bearing / steel
    quality_finish: str = "standard"    # basic / standard / premium
    basement: bool = False
    number_of_blocks: int = 1           # parallel execution units

    def __post_init__(self):
        self.project_type = str(self.project_type or "road").strip().lower()
        if self.project_type not in PROJECT_TYPES:
            raise ValueError(f"project_type must be one of {PROJECT_TYPES}.")

        if self.project_type == "road":
            self._init_road()
        elif self.project_type == "pipeline":
            self._init_pipeline()
        elif self.project_type == "building":
            self._init_building()

    def _init_road(self):
        raw_terrain = self.terrain or parse_terrain_from_title(self.title)
        self.terrain = str(raw_terrain or "plain").strip().lower()
        if self.terrain in {"flat", "plains", "plain terrain"}:
            self.terrain = "plain"
        elif self.terrain in {"hill", "mountain", "mountainous", "hilly terrain"}:
            self.terrain = "hilly"
        if self.terrain not in {"plain", "hilly"}:
            raise ValueError("terrain must be either 'plain' or 'hilly'.")

        if self.title:
            extracted_lanes = parse_lanes_from_title(self.title, default_lanes=self.lanes)
            self.lanes = extracted_lanes
            if self.length_km <= 0:
                extracted_len = parse_length_from_title(self.title, default_length=self.length_km)
                if extracted_len > 0:
                    self.length_km = extracted_len

    def _init_pipeline(self):
        raw_terrain = self.terrain or parse_terrain_from_title(self.title)
        self.terrain = str(raw_terrain or "plain").strip().lower()
        if self.terrain in {"flat", "plains"}:
            self.terrain = "plain"
        elif self.terrain in {"hill", "mountain", "mountainous"}:
            self.terrain = "hilly"
        if self.terrain not in {"plain", "hilly"}:
            self.terrain = "plain"
        # Auto-derive trench width from pipe diameter if not supplied
        if self.trench_width_m <= 0 and self.pipe_diameter_mm > 0:
            od_m = self.pipe_diameter_mm / 1000.0
            self.trench_width_m = round(od_m + 0.6, 2)  # pipe OD + 300mm each side clearance
        if self.number_of_work_fronts < 1:
            self.number_of_work_fronts = 1
        # Extract length from title if not given
        if self.length_km <= 0 and self.title:
            extracted_len = parse_length_from_title(self.title, default_length=0.0)
            if extracted_len > 0:
                self.length_km = extracted_len

    def _init_building(self):
        self.terrain = "plain"  # buildings always on flat/prepared ground
        if self.built_up_area_sqm <= 0 and self.length_km > 0:
            # Allow length_km to carry sqm value for internal routing
            self.built_up_area_sqm = self.length_km
        # number_of_floors: accept "G+3" style as 4
        if self.number_of_floors < 1:
            self.number_of_floors = 1
        # Derive number_of_blocks from area: 1 block per 3000 sqm (typical residential staircase block)
        if self.number_of_blocks < 1:
            self.number_of_blocks = max(1, round(self.built_up_area_sqm / 3000.0))



@dataclass
class ActivityReference:
    morth_spec: str
    data_book_item: str
    page_number: Optional[int] = None
    source_file: str = "guidelines.json"
    guideline_ref: str = "Project Guidelines"
    assumption_ref: str = "assumptions.yaml"

@dataclass
class ActivityCostBreakdown:
    direct_labour: float = 0.0
    direct_material: float = 0.0
    direct_equipment: float = 0.0
    direct_fuel: float = 0.0
    direct_cost: float = 0.0
    direct_unit_rate: float = 0.0       # planned (resource-derived) rate used for cost
    morth_direct_unit_rate: float = 0.0 # original MoRTH source rate, never modified
    planning_variation_pct: float = 0.0 # derived = (planned all-in rate - source rate) / source rate x 100 (reporting only; both sides all-in, so the gap shows the build-up uplift)
    overhead_pct: float = 10.0
    overhead_amount: float = 0.0
    # Memo-only amounts when OH/CP are embedded in the planned rate. These are
    # already inside `planned_unit_rate`, so they are NEVER added to the total.
    overhead_embedded: float = 0.0
    profit_pct: float = 10.0
    profit_amount: float = 0.0
    profit_embedded: float = 0.0
    total_amount: float = 0.0
    final_unit_rate: float = 0.0
    crash_site_establishment: float = 0.0
    crash_resident_staff: float = 0.0
    crash_plant_hire_uplift: float = 0.0
    crash_efficiency_loss: float = 0.0
    # Execution losses/wastage on resources actually consumed (material spoilage,
    # overbreak, fuel wastage, plant idle, minor rework). NOT an acceleration
    # cost, so it is kept out of crash_total_surcharge and reported separately.
    # Bottom-up resource check, priced from material_per_unit + wastage + real
    # material_rates_inr. Reported alongside the MoRTH rate for validation; it
    # does not re-price the estimate.
    bottom_up_material_cost: float = 0.0
    # Execution losses, priced from first principles and reported separately.
    # Material wastage is a QUANTITY problem (you issue more bags); the other
    # three are UTILISATION losses, each a fraction of its own base cost.
    # `resource_wastage_cost` = fuel + plant + labour;
    # `wastage_cost` = material + resource, i.e. the total execution loss.
    material_wastage_cost: float = 0.0
    fuel_wastage_cost: float = 0.0
    plant_wastage_cost: float = 0.0
    labour_wastage_cost: float = 0.0
    resource_wastage_cost: float = 0.0
    wastage_cost: float = 0.0
    # Compatibility flag for rate sources that embed wastage. The active road,
    # pipeline and building configurations keep this false so wastage is a
    # separate, auditable direct-cost component.
    wastage_in_base_rate: bool = False
    bottom_up_labour_cost: float = 0.0
    bottom_up_plant_cost: float = 0.0
    bottom_up_fuel_cost: float = 0.0
    bottom_up_total_cost: float = 0.0
    crash_total_surcharge: float = 0.0
    planned_unit_rate: float = 0.0
    # Legacy `crashing_factors` components, stored individually so the
    # acceleration waterfall is auditable end-to-end (Excel sheet 10 shows
    # all 8 components and they must sum exactly to crash_total_surcharge).
    crash_fleet_mobilisation: float = 0.0
    crash_shift_premium: float = 0.0
    crash_front_supervision: float = 0.0
    crash_site_congestion: float = 0.0
    # Price escalation: the same work, later in the programme, costs more. Kept
    # OUT of planned_unit_rate and OUT of crash_total_surcharge so the unit rate
    # stays comparable with the 2019 MoRTH yardstick and the acceleration
    # waterfall still reconciles. Reported as its own line on sheet 10.
    escalation_amount: float = 0.0   # rupees added by escalation to this activity
    escalation_factor: float = 1.0   # blended multiplier applied (1.0 = none)
    escalation_years: float = 0.0    # years from base date to this activity's mid-spend


@dataclass
class RoadActivity:
    activity_id: str
    wbs_code: str
    parent_wbs: str
    wbs_level: int
    work_package: str
    name: str
    description: str
    category: str
    unit: str
    quantity: float
    productivity_rate: float
    productivity_unit: str
    crew_description: str
    equipment_description: str
    duration_days: int
    predecessors: List[str] = field(default_factory=list)
    cost: ActivityCostBreakdown = field(default_factory=ActivityCostBreakdown)
    reference: ActivityReference = field(default_factory=lambda: ActivityReference("", ""))
    weight_pct: float = 0.0
    # Physical work content (crew-days = quantity / productivity rate), used to
    # weight PHYSICAL progress on the S-curve. Deliberately independent of cost
    # so the progress curve and the expenditure curve are different things.
    progress_weight_pct: float = 0.0
    # Effort days of production (productivity formula) before the project
    # calendar stretches them into calendar days. `duration_days` is the
    # calendar duration and is what CPM and costing use.
    effort_days: int = 0
    activity_key: str = ""
    
    # CPM Fields
    es: int = 0
    ef: int = 0
    ls: int = 0
    lf: int = 0
    total_float: int = 0
    free_float: int = 0
    is_critical: bool = False
    
    # Dates
    planned_start: Optional[date] = None
    planned_finish: Optional[date] = None

    # Optimization & Baseline Fields
    crews: int = 1
    work_fronts: int = 1
    shifts: float = 1.0
    base_duration_days: int = 0
    base_cost: float = 0.0
    base_crews: int = 1
    resources_added_desc: str = ""
    predecessor_lags: Dict[str, int] = field(default_factory=dict)

@dataclass
class MonthlyProgressItem:
    month_index: int  # 1, 2, ...
    month_name: str
    period_start_day: int
    period_end_day: int
    period_planned_progress_pct: float = 0.0
    cumulative_planned_progress_pct: float = 0.0
    period_expenditure_crore: float = 0.0
    cumulative_expenditure_crore: float = 0.0

@dataclass
class ProjectScheduleResult:
    project_input: ProjectInput
    activities: List[RoadActivity]
    total_calculated_cost_crore: float
    cost_variance_crore: float
    cost_variance_pct: float
    calculated_duration_days: int
    schedule_variance_days: int
    critical_path: List[str]
    monthly_progress: List[MonthlyProgressItem]
    warnings: List[str] = field(default_factory=list)
    
    # Constraint-Based Optimization Results
    is_feasible: bool = True
    feasibility_message: str = ""
    baseline_duration_days: int = 0
    baseline_cost_crore: float = 0.0
    baseline_critical_path: List[str] = field(default_factory=list)
    baseline_activities: List[RoadActivity] = field(default_factory=list)
    optimized_duration_days: int = 0
    optimized_cost_crore: float = 0.0
    duration_saved_days: int = 0
    additional_cost_crore: float = 0.0
    work_fronts: int = 1
    # Fronts in the NORMAL (uncrashed) organisation - >1 on long corridors.
    baseline_work_fronts: int = 1
    resources_added_summary: List[Dict[str, Any]] = field(default_factory=list)
    optimization_iterations: List[Dict[str, Any]] = field(default_factory=list)
    # Why the search stopped (empty when the target was met on schedule).
    optimization_note: str = ""
    # Post-solve sanity gates (benchmarks from assumptions.yaml).
    validation_warnings: List[str] = field(default_factory=list)
    # Effective pipeline scheme scope AFTER village-derived defaults have been
    # applied by the quantity take-off (empty for road/building projects).
    pipeline_scope: Dict[str, Any] = field(default_factory=dict)
    # Time-driven acceleration step rates actually used by the cost engine, so
    # report labels cite the project-type config instead of hard-coded numbers.
    accel_site_establishment_rate: float = 0.0
    accel_resident_staff_rate: float = 0.0


# ======================================================================
# UPDATE-PLAN workflow models
# ======================================================================

@dataclass
class ActivityActualStatus:
    """Per-activity actual completion status used during replanning."""
    activity_id: str
    name: str = ""
    # 0.0 - 100.0  cumulative completion of this activity's quantity
    actual_completion_pct: float = 0.0
    completed_quantity: float = 0.0
    remaining_quantity: float = 0.0
    original_quantity: float = 0.0
    # "completed" | "in_progress" | "not_started"
    status: str = "not_started"
    # baseline cumulative completion at the status date (for comparison)
    baseline_completion_pct: float = 0.0
    is_delayed: bool = False


@dataclass
class UpdatePlanInput:
    """All inputs for the update-plan command."""
    input_path: str
    start_date: Optional[date] = None        # original project start date
    status_date: date = field(default_factory=lambda: date.today())
    actual_progress_pct: float = 0.0          # cumulative actual progress at status date
    target_date: Optional[date] = None        # Optional: None => use previous Excel optimized final date
    target_progress_pct: Optional[float] = None  # None => use baseline target at target_date
    output_path: str = "road_project_schedule_updated.xlsx"
    # Optional activity-level progress: {activity_id: completion_pct}
    activity_progress: Optional[Dict[str, float]] = None
    guidelines_path: str = "road_guidelines.json"
    assumptions_path: str = "road_assumptions.yaml"
    enable_acceleration_cost: bool = True


@dataclass
class UpdatePlanResult:
    """Complete output of the update planning engine."""
    project_input: ProjectInput
    baseline_result: ProjectScheduleResult

    # Progress comparison
    actual_progress_pct: float = 0.0
    baseline_progress_at_status_date: float = 0.0
    progress_gap_pct: float = 0.0          # baseline - actual (positive = behind)
    status_date: date = field(default_factory=lambda: date.today())
    target_date: date = field(default_factory=lambda: date.today())
    target_progress_pct: float = 100.0
    required_progress_pct: float = 0.0      # target - actual
    required_calendar_days: int = 0        # target_date - status_date

    # Per-activity status
    activity_status: List[ActivityActualStatus] = field(default_factory=list)

    # Revised plan (remaining work)
    revised_activities: List[RoadActivity] = field(default_factory=list)
    revised_duration_days: int = 0
    revised_cost_crore: float = 0.0
    revised_critical_path: List[str] = field(default_factory=list)
    revised_monthly_progress: List[MonthlyProgressItem] = field(default_factory=list)

    # Feasibility
    original_revised_completion_date: Optional[date] = None  # before acceleration
    revised_completion_date: Optional[date] = None           # after acceleration
    target_feasible: bool = False
    feasibility_message: str = ""
    projected_delay_days: int = 0          # > 0 means late vs target date
    projected_progress_at_target_date: float = 0.0

    # Acceleration
    additional_acceleration_cost_crore: float = 0.0
    work_fronts: int = 1
    resources_added_summary: List[Dict[str, Any]] = field(default_factory=list)

    # Cost summary
    spent_cost_crore: float = 0.0          # cost of completed work
    remaining_base_cost_crore: float = 0.0 # remaining cost at normal resource levels
    remaining_cost_crore: float = 0.0      # total remaining cost (normal + acceleration)
    total_project_cost_crore: float = 0.0  # spent + remaining
    baseline_total_cost_crore: float = 0.0
    spent_cost_inr: float = 0.0
    remaining_cost_inr: float = 0.0
    total_project_cost_inr: float = 0.0

    # Reporting
    warnings: List[str] = field(default_factory=list)
    limitations: List[str] = field(default_factory=list)
    is_feasible: bool = True
    validation_report: Any = None
