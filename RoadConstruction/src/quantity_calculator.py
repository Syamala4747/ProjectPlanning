"""
Quantity Calculation Engine for Road / Pipeline / Building Construction Projects.
Dispatches to the appropriate sub-calculator based on project_type.
All geometric standards and structural allowances are loaded from assumptions.yaml.
"""
import math
from typing import Dict, Any, Optional
from .models import ProjectInput


# ─────────────────────────────────────────────────────────────────────────────
# Public façade — single entry-point unchanged from the road-only era
# ─────────────────────────────────────────────────────────────────────────────

class QuantityCalculator:
    """
    Dispatches to RoadQuantityCalculator, PipelineQuantityCalculator, or
    BuildingQuantityCalculator based on project_input.project_type.

    Exposed interface is identical for all types:
        .calculate_quantities() → Dict[str, Dict[str, Any]]
    where each key is an activity `key` from the activity_master catalog.
    """

    def __init__(self, project_input: ProjectInput,
                 assumptions: Optional[Dict[str, Any]] = None):
        self._input = project_input
        self._assumptions = assumptions or {}
        ptype = str(getattr(project_input, "project_type", "road") or "road").lower()

        if ptype == "pipeline":
            self._calc = PipelineQuantityCalculator(project_input, self._assumptions)
        elif ptype == "building":
            self._calc = BuildingQuantityCalculator(project_input, self._assumptions)
        else:
            self._calc = RoadQuantityCalculator(project_input, self._assumptions)

    def calculate_quantities(self) -> Dict[str, Dict[str, Any]]:
        return self._calc.calculate_quantities()


# ─────────────────────────────────────────────────────────────────────────────
# Road quantity calculator (original logic, unchanged)
# ─────────────────────────────────────────────────────────────────────────────

class RoadQuantityCalculator:
    """
    Computes road engineering quantities derived purely from parameters in
    assumptions.yaml (highway_parameters, lane_standards, structure_allowances).
    """

    def __init__(self, project_input: ProjectInput,
                 assumptions: Optional[Dict[str, Any]] = None):
        self.input = project_input
        self.assumptions = assumptions or {}
        self.length_km = max(0.1, float(project_input.length_km))
        self.length_m = self.length_km * 1000.0

        lanes = max(1, int(getattr(project_input, 'lanes', 4)))
        self.lanes = lanes
        self.lane_factor = lanes / 4.0

        params = self.assumptions.get("highway_parameters", {})
        lane_standards = self.assumptions.get("lane_standards", {})

        if lanes in lane_standards:
            std = lane_standards[lanes]
            self.carriageway_width_m = float(std.get("carriageway_width_m", 18.0))
            self.formation_width_m = float(std.get("formation_width_m", 28.0))
            self.median_width_m = float(std.get("median_width_m", 4.5))
            self.subgrade_width_m = float(std.get("subgrade_width_m", 22.0))
            self.gsb_width_m = float(std.get("gsb_width_m", 20.5))
            self.wmm_width_m = float(std.get("wmm_width_m", 19.5))
            self.service_road_sides = float(std.get("service_road_sides", 2.0))
        else:
            base_formation = float(params.get("formation_width_m", 28.0))
            base_carriageway = float(params.get("carriageway_width_m", 18.0))
            base_subgrade = float(params.get("subgrade_width_m", 22.0))
            base_gsb = float(params.get("gsb_width_m", 20.5))
            base_wmm = float(params.get("wmm_width_m", 19.5))

            self.carriageway_width_m = round(base_carriageway * self.lane_factor, 2)
            self.median_width_m = float(params.get("median_width_m", 4.5))
            self.formation_width_m = round(base_formation * (0.35 + 0.65 * self.lane_factor), 2)
            self.subgrade_width_m = round(base_subgrade * self.lane_factor, 2)
            self.gsb_width_m = round(base_gsb * self.lane_factor, 2)
            self.wmm_width_m = round(base_wmm * self.lane_factor, 2)
            self.service_road_sides = 2.0 if self.lanes >= 4 else (1.0 if self.lanes >= 2 else 0.0)

        self.input.carriageway_width_m = self.carriageway_width_m
        self.input.formation_width_m = self.formation_width_m

        self.subgrade_thickness_m = float(params.get("subgrade_thickness_m", 0.50))
        self.gsb_thickness_m = float(params.get("gsb_thickness_m", 0.20))
        self.wmm_thickness_m = float(params.get("wmm_thickness_m", 0.25))
        self.dbm_thickness_m = float(params.get("dbm_thickness_m", 0.100))
        self.bc_thickness_m = float(params.get("bc_thickness_m", 0.050))

        self.embankment_height_m = float(params.get("embankment_avg_height_m", 1.80))
        self.embankment_fill_ratio = float(params.get("embankment_fill_ratio", 0.75))
        self.excavation_depth_m = float(params.get("excavation_avg_depth_m", 0.75))
        self.excavation_cut_ratio = float(params.get("excavation_cut_ratio", 0.40))

        self.trees_per_km = float(params.get("trees_per_km", 200))
        self.box_culverts_per_km = float(params.get("box_culverts_per_km", 1.6))
        self.box_culvert_concrete_cum = float(params.get("box_culvert_concrete_cum", 350.0))
        self.pipe_culverts_per_km = float(params.get("pipe_culverts_per_km", 1.2))
        self.pipe_culvert_length_m = float(params.get("pipe_culvert_length_m", 64.0))

        self.service_road_ratio = float(params.get("service_road_ratio", 0.60))
        self.service_road_width_m = float(params.get("service_road_width_m", 7.0))
        self.slope_coverage_ratio = float(params.get("slope_coverage_ratio", 0.75))

        struct = self.assumptions.get("structure_allowances", {})
        ref_km = float(struct.get("reference_corridor_km", 25.0))
        self.ref_corridor_km = ref_km
        self.piling_ref_m = float(struct.get("bridge_piling_m", 18500.0))
        self.bridge_sub_ref_cum = float(struct.get("bridge_substructure_cum", 18000.0))
        self.bridge_super_ref_cum = float(struct.get("bridge_superstructure_cum", 26000.0))
        self.flyover_ref_cum = float(struct.get("flyover_structures_cum", 16000.0))
        self.vup_ref_cum = float(struct.get("vup_structures_cum", 14000.0))
        self.re_walls_ref_sqm = float(struct.get("re_walls_sqm", 65000.0))

    def calculate_quantities(self) -> Dict[str, Dict[str, Any]]:
        quantities = {}
        km_factor = (self.length_km / self.ref_corridor_km) * self.lane_factor

        quantities["site_clearance"] = {
            "quantity": round((self.length_m * self.formation_width_m) / 10000.0, 2),
            "unit": "ha",
            "formula": "Length (m) × Formation Width (m) / 10,000"
        }
        quantities["utility_shifting"] = {
            "quantity": round(self.length_km, 2),
            "unit": "km",
            "formula": "Corridor Length (km)"
        }
        quantities["tree_felling"] = {
            "quantity": float(max(10, int(round(self.length_km * self.trees_per_km)))),
            "unit": "Nos",
            "formula": "Length (km) × Trees/km"
        }
        quantities["site_mobilization"] = {
            "quantity": round(max(0.04, km_factor), 3),
            "unit": "set",
            "formula": "Site Base Camp & Plants Setup (Scaled to Project Size)"
        }
        quantities["excavation"] = {
            "quantity": round(self.length_m * self.formation_width_m * self.excavation_depth_m * self.excavation_cut_ratio, 2),
            "unit": "cum",
            "formula": "Length × Formation Width × Cut Depth × Cut Ratio"
        }
        quantities["embankment"] = {
            "quantity": round(self.length_m * self.formation_width_m * self.embankment_height_m * self.embankment_fill_ratio, 2),
            "unit": "cum",
            "formula": "Length × Formation Width × Height × Fill Ratio"
        }
        quantities["subgrade"] = {
            "quantity": round(self.length_m * self.subgrade_width_m * self.subgrade_thickness_m, 2),
            "unit": "cum",
            "formula": "Length × Subgrade Width × Thickness"
        }
        quantities["median_dressing"] = {
            "quantity": round(self.length_m * self.median_width_m, 2),
            "unit": "sqm",
            "formula": "Length × Median Width"
        }
        quantities["gsb"] = {
            "quantity": round(self.length_m * self.gsb_width_m * self.gsb_thickness_m, 2),
            "unit": "cum",
            "formula": "Length × GSB Width × Thickness"
        }
        quantities["wmm"] = {
            "quantity": round(self.length_m * self.wmm_width_m * self.wmm_thickness_m, 2),
            "unit": "cum",
            "formula": "Length × WMM Width × Thickness"
        }
        quantities["prime_coat"] = {
            "quantity": round(self.length_m * self.wmm_width_m, 2),
            "unit": "sqm",
            "formula": "Length × WMM Width"
        }
        quantities["dbm"] = {
            "quantity": round(self.length_m * self.carriageway_width_m, 2),
            "unit": "sqm",
            "formula": "Length × Carriageway Width"
        }
        quantities["tack_coat"] = {
            "quantity": round(self.length_m * self.carriageway_width_m, 2),
            "unit": "sqm",
            "formula": "Length × Carriageway Width"
        }
        quantities["bc"] = {
            "quantity": round(self.length_m * self.carriageway_width_m, 2),
            "unit": "sqm",
            "formula": "Length × Carriageway Width"
        }
        box_count = max(0.2, self.length_km * self.box_culverts_per_km)
        pipe_count = max(0.2, self.length_km * self.pipe_culverts_per_km)
        culvert_width_factor = round(self.formation_width_m / 28.0, 2)
        quantities["box_culverts"] = {
            "quantity": round(box_count * self.box_culvert_concrete_cum * culvert_width_factor, 2),
            "unit": "cum",
        }
        quantities["pipe_culverts"] = {
            "quantity": round(pipe_count * self.pipe_culvert_length_m * culvert_width_factor, 2),
            "unit": "metre",
        }
        quantities["bridge_piling"] = {
            "quantity": round(km_factor * self.piling_ref_m, 2),
            "unit": "metre",
        }
        quantities["bridge_substructure"] = {
            "quantity": round(km_factor * self.bridge_sub_ref_cum, 2),
            "unit": "cum",
        }
        quantities["bridge_superstructure"] = {
            "quantity": round(km_factor * self.bridge_super_ref_cum, 2),
            "unit": "cum",
        }
        quantities["flyover_structures"] = {
            "quantity": round(km_factor * self.flyover_ref_cum, 2),
            "unit": "cum",
        }
        quantities["vup_structures"] = {
            "quantity": round(km_factor * self.vup_ref_cum, 2),
            "unit": "cum",
        }
        quantities["re_walls"] = {
            "quantity": round(km_factor * self.re_walls_ref_sqm, 2),
            "unit": "sqm",
        }
        quantities["drainage"] = {
            "quantity": round(self.length_m * 2.0, 2),
            "unit": "metre",
        }
        quantities["slope_protection"] = {
            "quantity": round(2.0 * self.length_m * (self.embankment_height_m * 2.24) * self.slope_coverage_ratio, 2),
            "unit": "sqm",
        }
        quantities["service_roads"] = {
            "quantity": round((self.service_road_ratio * self.length_m) * self.service_road_width_m * self.service_road_sides, 2),
            "unit": "sqm",
        }
        quantities["road_furniture"] = {
            "quantity": round(self.length_km * (0.5 + 0.5 * self.lane_factor), 2),
            "unit": "km",
        }
        quantities["toll_plaza_atms"] = {
            "quantity": round(max(0.04, km_factor), 3),
            "unit": "set",
        }
        return quantities


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline quantity calculator
# ─────────────────────────────────────────────────────────────────────────────

class PipelineQuantityCalculator:
    """
    Computes water-pipeline engineering quantities from project inputs and
    pipeline_assumptions.yaml (pipeline_parameters section).

    All formulas are keyed to the activity `key` field in pipeline_guidelines.json
    activity_master (pip_001 … pip_017).
    """

    def __init__(self, project_input: ProjectInput,
                 assumptions: Optional[Dict[str, Any]] = None):
        self.input = project_input
        self.assumptions = assumptions or {}
        self.length_km = max(0.1, float(project_input.length_km))
        self.length_m = self.length_km * 1000.0

        # Pipe geometry
        self.diameter_mm = max(100.0, float(getattr(project_input, "pipe_diameter_mm", 300) or 300))
        self.pipe_material = str(getattr(project_input, "pipe_material", "DI") or "DI").upper()
        self.trench_depth_m = max(1.0, float(getattr(project_input, "trench_depth_m", 2.0) or 2.0))
        od_m = self.diameter_mm / 1000.0
        # Trench width: pipe OD + 300 mm each side (IS:4745 / CPHEEO manual)
        supplied_width = float(getattr(project_input, "trench_width_m", 0.0) or 0.0)
        self.trench_width_m = supplied_width if supplied_width > 0 else round(od_m + 0.60, 2)

        # Pipeline parameters from assumptions
        params = self.assumptions.get("pipeline_parameters", {})
        self.bedding_thickness_m = float(params.get("bedding_thickness_m", 0.15))
        # Joint spacing: DI/MS/RCC => 5.5m standard, HDPE => 6m (butt fused)
        default_joint_spacing = 6.0 if self.pipe_material == "HDPE" else 5.5
        self.joint_spacing_m = float(params.get("joint_spacing_m", default_joint_spacing))
        # Valve chambers: every 500m for water, every 300m for sewerage (CPHEEO)
        pipeline_type = str(getattr(project_input, "pipeline_type", "water") or "water").lower()
        default_valve_spacing = 500.0 if pipeline_type == "water" else 300.0
        self.valve_chamber_spacing_m = float(params.get("valve_chamber_spacing_m", default_valve_spacing))
        # Hydro-test sections: 500m per section (CPHEEO Part III 6.4.3)
        self.hydrotest_section_length_m = float(params.get("hydrotest_section_length_m", 500.0))
        # Thrust block concrete: ~0.8 cum per bend/tee; one fitting per 250m average
        self.thrust_block_cum_per_km = float(params.get("thrust_block_cum_per_km", 3.2))
        # Road-crossing trenchless: typically 5% of length needs HDD/jacking
        self.trenchless_ratio = float(params.get("trenchless_ratio", 0.05))
        # Work corridor for clearance/access (m either side of trench centre)
        self.work_corridor_m = float(params.get("work_corridor_m", 3.0))
        # Reinstatement: 80% of trench surface area (some areas are open field)
        self.reinstatement_ratio = float(params.get("reinstatement_ratio", 0.80))

    def calculate_quantities(self) -> Dict[str, Dict[str, Any]]:
        q: Dict[str, Dict[str, Any]] = {}

        # Total trench volume
        trench_volume = round(self.length_m * self.trench_width_m * self.trench_depth_m, 1)
        # Bedding volume (below pipe)
        bedding_volume = round(self.length_m * self.trench_width_m * self.bedding_thickness_m, 1)
        # Number of joints
        joint_count = max(1, int(math.ceil(self.length_m / self.joint_spacing_m)))
        # Number of valve chambers
        chamber_count = max(2, int(math.ceil(self.length_m / self.valve_chamber_spacing_m)))
        # Thrust / anchor blocks
        thrust_block_cum = round(self.length_km * self.thrust_block_cum_per_km, 1)
        # Initial backfill (selected surround: 300mm above pipe crown)
        pipe_od_m = self.diameter_mm / 1000.0
        initial_backfill_depth = self.bedding_thickness_m + pipe_od_m + 0.30
        initial_backfill_vol = round(self.length_m * self.trench_width_m * initial_backfill_depth, 1)
        # Final backfill (remaining trench)
        final_backfill_depth = max(0.1, self.trench_depth_m - initial_backfill_depth)
        final_backfill_vol = round(self.length_m * self.trench_width_m * final_backfill_depth, 1)
        # Hydrotest sections
        test_sections = max(1, int(math.ceil(self.length_m / self.hydrotest_section_length_m)))
        # Work corridor area for site clearance
        corridor_area = round(self.length_m * (self.trench_width_m + 2 * self.work_corridor_m), 0)
        # Reinstatement area
        reinstatement_area = round(self.length_m * self.trench_width_m * self.reinstatement_ratio, 1)

        # PIP-001: Survey and route marking
        q["pip_001"] = {"quantity": round(self.length_km, 2), "unit": "km",
                        "formula": "Pipeline length (km)"}

        # PIP-002: Mobilization
        q["pip_002"] = {"quantity": 1.0, "unit": "LS",
                        "formula": "Lump sum"}

        # PIP-003: Site clearance
        q["pip_003"] = {"quantity": corridor_area, "unit": "sqm",
                        "formula": "Length × (trench width + 2×corridor)"}

        # PIP-004: Trench excavation
        q["pip_004"] = {"quantity": trench_volume, "unit": "cum",
                        "formula": "L × trench_width × trench_depth"}

        # PIP-005: Dewatering (duration-based; priced per day as proportion of trench days)
        dewater_days = max(5, int(math.ceil(trench_volume / 300.0)))  # ~300 cum/day excavation
        q["pip_005"] = {"quantity": float(dewater_days), "unit": "day",
                        "formula": "Proportional to trench excavation duration"}

        # PIP-006: Bedding and pipe foundation
        q["pip_006"] = {"quantity": bedding_volume, "unit": "cum",
                        "formula": "L × trench_width × bedding_thickness"}

        # PIP-007: Pipe supply, delivery, stacking
        q["pip_007"] = {"quantity": round(self.length_m * 1.02, 1), "unit": "m",
                        "formula": "Length + 2% wastage/overlap"}

        # PIP-008: Pipe lowering, laying, alignment
        q["pip_008"] = {"quantity": round(self.length_m, 1), "unit": "m",
                        "formula": "Pipeline length (m)"}

        # PIP-009: Jointing and fittings
        q["pip_009"] = {"quantity": float(joint_count), "unit": "joint",
                        "formula": f"Length / {self.joint_spacing_m}m joint spacing"}

        # PIP-010: Valves, chambers and appurtenances
        q["pip_010"] = {"quantity": float(chamber_count), "unit": "no",
                        "formula": f"Length / {self.valve_chamber_spacing_m}m spacing"}

        # PIP-011: Thrust/anchor blocks
        q["pip_011"] = {"quantity": thrust_block_cum, "unit": "cum",
                        "formula": f"{self.thrust_block_cum_per_km} cum/km × length"}

        # PIP-012: Initial backfilling (selected surround)
        q["pip_012"] = {"quantity": initial_backfill_vol, "unit": "cum",
                        "formula": "L × trench_width × (bed+OD+0.30m)"}

        # PIP-013: Final backfilling and compaction
        q["pip_013"] = {"quantity": final_backfill_vol, "unit": "cum",
                        "formula": "Remaining trench volume above initial backfill"}

        # PIP-014: Hydrostatic pressure testing
        q["pip_014"] = {"quantity": float(test_sections), "unit": "test_section",
                        "formula": f"Length / {self.hydrotest_section_length_m}m per section"}

        # PIP-015: Flushing, disinfection, water quality
        q["pip_015"] = {"quantity": round(self.length_km, 2), "unit": "km",
                        "formula": "Pipeline length (km)"}

        # PIP-016: Road/surface reinstatement
        q["pip_016"] = {"quantity": reinstatement_area, "unit": "sqm",
                        "formula": f"Length × trench_width × {self.reinstatement_ratio} reinstatement ratio"}

        # PIP-017: Commissioning, handover, demobilization
        q["pip_017"] = {"quantity": 1.0, "unit": "LS",
                        "formula": "Lump sum"}

        return q


# ─────────────────────────────────────────────────────────────────────────────
# Building quantity calculator
# ─────────────────────────────────────────────────────────────────────────────

class BuildingQuantityCalculator:
    """
    Computes residential/commercial building quantities from project inputs and
    building_assumptions.yaml (building_parameters section).

    Activity keys map to BLD-001 … BLD-020 in building_guidelines.json.
    Quantities are derived from built-up area, floors, and standard
    CPWD DSR 2021 / IS 456 ratios.
    """

    def __init__(self, project_input: ProjectInput,
                 assumptions: Optional[Dict[str, Any]] = None):
        self.input = project_input
        self.assumptions = assumptions or {}

        self.bua_sqm = max(100.0, float(getattr(project_input, "built_up_area_sqm", 0) or 0))
        self.floors = max(1, int(getattr(project_input, "number_of_floors", 1) or 1))
        self.basement = bool(getattr(project_input, "basement", False))
        self.building_type = str(getattr(project_input, "building_type", "residential") or "residential").lower()
        self.construction_type = str(getattr(project_input, "construction_type", "rcc") or "rcc").lower()
        self.quality = str(getattr(project_input, "quality_finish", "standard") or "standard").lower()
        self.blocks = max(1, int(getattr(project_input, "number_of_blocks", 1) or 1))

        # Per-floor area
        self.floor_area_sqm = self.bua_sqm / self.floors

        # Building parameters from assumptions
        params = self.assumptions.get("building_parameters", {})

        # Structural ratios (per sqm of BUA) — CPWD DSR 2021 / IS 456 based
        # Foundation: footings + raft; ~0.35 cum RCC per sqm for G+3 residential
        self.foundation_rcc_per_sqm = float(params.get("foundation_rcc_per_sqm", 0.35))
        # Frame RCC: columns + beams + slabs; ~0.40 cum/sqm for RCC frame
        self.frame_rcc_per_sqm = float(params.get("frame_rcc_per_sqm", 0.40))
        # Masonry: ~0.18 cum/sqm BUA (230mm brick wall, openings deducted)
        self.masonry_per_sqm = float(params.get("masonry_per_sqm", 0.18))
        # Plaster area: 3.5× BUA (walls both sides + ceilings; IS standard)
        self.plaster_area_factor = float(params.get("plaster_area_factor", 3.5))
        # Waterproofing: terrace + toilet wet areas; ~0.30× BUA
        self.waterproofing_area_factor = float(params.get("waterproofing_area_factor", 0.30))
        # Flooring = 100% BUA (slab area)
        self.flooring_area_factor = float(params.get("flooring_area_factor", 1.0))
        # Painting = external + internal surfaces; ~4.0× BUA
        self.painting_area_factor = float(params.get("painting_area_factor", 4.0))
        # Basement: 0 if no basement; area = floor_area × 1 level
        self.basement_excavation_depth_m = float(params.get("basement_excavation_depth_m", 3.5))
        # External works: landscaping + paving = plot coverage; ~0.5× BUA
        self.external_works_factor = float(params.get("external_works_factor", 0.5))
        # Site area estimate: BUA / FAR; default FAR = 1.5 (typical residential)
        self.far = float(params.get("floor_area_ratio", 1.5))
        self.site_area_sqm = round(self.bua_sqm / (self.far * self.floors), 1)

        # Quality multipliers on finishing quantities
        quality_map = {"basic": 0.8, "standard": 1.0, "premium": 1.3}
        self.quality_mult = quality_map.get(self.quality, 1.0)

    def calculate_quantities(self) -> Dict[str, Dict[str, Any]]:
        q: Dict[str, Dict[str, Any]] = {}

        # Derived quantities
        foundation_rcc = round(self.bua_sqm * self.foundation_rcc_per_sqm, 1)
        frame_rcc = round(self.bua_sqm * self.frame_rcc_per_sqm, 1)
        masonry_cum = round(self.bua_sqm * self.masonry_per_sqm, 1)
        plaster_sqm = round(self.bua_sqm * self.plaster_area_factor * self.quality_mult, 1)
        waterproofing_sqm = round(self.bua_sqm * self.waterproofing_area_factor, 1)
        flooring_sqm = round(self.bua_sqm * self.flooring_area_factor * self.quality_mult, 1)
        painting_sqm = round(self.bua_sqm * self.painting_area_factor * self.quality_mult, 1)
        # Electrical / plumbing / HVAC / fire — MEP in sqm of served area
        elec_sqm = round(self.bua_sqm * self.quality_mult, 1)
        plumbing_sqm = round(self.bua_sqm * self.quality_mult, 1)
        hvac_sqm = round(self.bua_sqm * (0.5 if self.building_type == "residential" else 1.0) * self.quality_mult, 1)
        fire_sqm = round(self.bua_sqm, 1)
        external_sqm = round(self.site_area_sqm * self.external_works_factor, 1)

        # Basement
        basement_excavation = 0.0
        if self.basement:
            basement_excavation = round(self.floor_area_sqm * self.basement_excavation_depth_m, 1)

        site_excavation = round(
            max(100.0, self.site_area_sqm * 0.5),  # strip + level; ~0.5m average
            1
        ) + basement_excavation

        # BLD-001: Survey, soil investigation
        q["bld_001"] = {"quantity": 1.0, "unit": "LS",
                        "formula": "Lump sum"}

        # BLD-002: Mobilization
        q["bld_002"] = {"quantity": 1.0, "unit": "LS",
                        "formula": "Lump sum"}

        # BLD-003: Site clearance and temporary works
        q["bld_003"] = {"quantity": round(self.site_area_sqm, 1), "unit": "sqm",
                        "formula": "Estimated site area = BUA / (FAR × floors)"}

        # BLD-004: Bulk excavation and earthwork
        q["bld_004"] = {"quantity": site_excavation, "unit": "cum",
                        "formula": "Strip + level + basement (if any)"}

        # BLD-005: Foundation (PCC + footings)
        q["bld_005"] = {"quantity": foundation_rcc, "unit": "cum",
                        "formula": f"BUA × {self.foundation_rcc_per_sqm} cum/sqm"}

        # BLD-006: RCC frame (columns + beams + slabs)
        q["bld_006"] = {"quantity": frame_rcc, "unit": "cum",
                        "formula": f"BUA × {self.frame_rcc_per_sqm} cum/sqm"}

        # BLD-007: Masonry / brickwork
        q["bld_007"] = {"quantity": masonry_cum, "unit": "cum",
                        "formula": f"BUA × {self.masonry_per_sqm} cum/sqm"}

        # BLD-008: Plastering and pointing
        q["bld_008"] = {"quantity": plaster_sqm, "unit": "sqm",
                        "formula": f"BUA × {self.plaster_area_factor} area factor × quality"}

        # BLD-009: Waterproofing
        q["bld_009"] = {"quantity": waterproofing_sqm, "unit": "sqm",
                        "formula": f"BUA × {self.waterproofing_area_factor} (terrace + wet areas)"}

        # BLD-010: Flooring and tiling
        q["bld_010"] = {"quantity": flooring_sqm, "unit": "sqm",
                        "formula": f"BUA × {self.flooring_area_factor} × quality"}

        # BLD-011: Doors, windows and glazing
        # ~1 door per 25 sqm, ~1 window per 20 sqm (typical residential)
        door_count = max(2, int(round(self.bua_sqm / 25.0)))
        window_count = max(4, int(round(self.bua_sqm / 20.0)))
        q["bld_011"] = {"quantity": float(door_count + window_count), "unit": "no",
                        "formula": "BUA / 25 doors + BUA / 20 windows"}

        # BLD-012: External finishes and facade
        external_finish_sqm = round(self.bua_sqm * 0.4, 1)  # external wall area ~ 0.4× BUA
        q["bld_012"] = {"quantity": external_finish_sqm, "unit": "sqm",
                        "formula": "BUA × 0.4 (external wall area)"}

        # BLD-013: Internal painting
        q["bld_013"] = {"quantity": painting_sqm, "unit": "sqm",
                        "formula": f"BUA × {self.painting_area_factor} area factor × quality"}

        # BLD-014: Electrical works
        q["bld_014"] = {"quantity": elec_sqm, "unit": "sqm",
                        "formula": "BUA × quality multiplier"}

        # BLD-015: Plumbing and sanitary
        q["bld_015"] = {"quantity": plumbing_sqm, "unit": "sqm",
                        "formula": "BUA × quality multiplier"}

        # BLD-016: HVAC works
        q["bld_016"] = {"quantity": hvac_sqm, "unit": "sqm",
                        "formula": "BUA × (0.5 residential / 1.0 commercial) × quality"}

        # BLD-017: Fire fighting
        q["bld_017"] = {"quantity": fire_sqm, "unit": "sqm",
                        "formula": "BUA (IS 1641 coverage)"}

        # BLD-018: Lifts and escalators
        # 1 lift per 2000 sqm BUA (typical residential)
        lift_count = max(0, int(math.ceil(self.bua_sqm / 2000.0)) if self.floors >= 4 else 0)
        q["bld_018"] = {"quantity": float(max(1, lift_count)) if self.floors >= 4 else 0.0,
                        "unit": "no", "formula": "BUA / 2000 sqm (for G+3 and above)"}

        # BLD-019: External works (landscaping, paving, compound wall)
        q["bld_019"] = {"quantity": external_sqm, "unit": "sqm",
                        "formula": f"Site area × {self.external_works_factor} external factor"}

        # BLD-020: Commissioning and handover
        q["bld_020"] = {"quantity": 1.0, "unit": "LS",
                        "formula": "Lump sum"}

        return q
