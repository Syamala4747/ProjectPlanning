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

    def scheme_scope(self) -> Dict[str, Any]:
        """Effective scheme scope after defaults (pipeline only; else empty).

        Only meaningful after ``calculate_quantities()`` has run, because the
        village-count default derivation happens there.
        """
        calc = self._calc
        if not hasattr(calc, "num_villages"):
            return {}
        wtp_mld = float(getattr(calc, "wtp_capacity_mld", 0.0) or 0.0)
        ref = float(getattr(calc, "reference_wtp_mld", 0.0) or 0.0)
        return {
            "villages": int(getattr(calc, "num_villages", 0) or 0),
            "ohts": int(getattr(calc, "num_oht", 0) or 0),
            "wtp_mld": wtp_mld,
            "reference_wtp_mld": ref,
            "wtp_trains": (wtp_mld / ref) if (wtp_mld > 0 and ref > 0) else 0.0,
            "cwr_units": (wtp_mld / ref) if (wtp_mld > 0 and ref > 0) else 0.0,
            "ibps": int(getattr(calc, "num_ibps", 0) or 0),
            "fhtc": int(getattr(calc, "num_fhtc", 0) or 0),
        }


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
            self.service_road_sides = float(
                getattr(project_input, "service_road_sides", 0)
            )
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
            self.service_road_sides = float(
                getattr(project_input, "service_road_sides", 0)
            )

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
        bridge_count = int(getattr(self.input, "bridges", 0) or 0)
        flyover_count = int(getattr(self.input, "flyovers", 0) or 0)
        underpass_count = int(getattr(self.input, "underpasses", 0) or 0)
        quantities["bridge_piling"] = {
            "quantity": round(bridge_count * self.piling_ref_m, 2),
            "unit": "metre",
        }
        quantities["bridge_substructure"] = {
            "quantity": round(bridge_count * self.bridge_sub_ref_cum, 2),
            "unit": "cum",
        }
        quantities["bridge_superstructure"] = {
            "quantity": round(bridge_count * self.bridge_super_ref_cum, 2),
            "unit": "cum",
        }
        quantities["flyover_structures"] = {
            "quantity": round(flyover_count * self.flyover_ref_cum, 2),
            "unit": "cum",
        }
        quantities["vup_structures"] = {
            "quantity": round(underpass_count * self.vup_ref_cum, 2),
            "unit": "cum",
        }
        quantities["re_walls"] = {
            "quantity": round(
                float(getattr(self.input, "retaining_walls_sqm", 0) or 0), 2
            ),
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
            "quantity": round(
                int(getattr(self.input, "toll_plazas", 0) or 0)
                + int(getattr(self.input, "atms", 0) or 0), 3
            ),
            "unit": "set",
        }
        return quantities


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline quantity calculator
# ─────────────────────────────────────────────────────────────────────────────

class PipelineQuantityCalculator:
    """
    Computes water-pipeline and water-supply-scheme quantities from project
    inputs and pipeline_assumptions.yaml (pipeline_parameters /
    pipeline_project_inputs sections).

    Every activity `key` in pipeline_guidelines.json activity_master
    (survey_setting_out … scada_automation) has an explicit formula here. A
    quantity of 0 means the activity is genuinely out of scope for the run
    (e.g. a DI main does not carry HDPE laying, and a corridor-only run has no
    WTP/OHT package) and therefore takes 0 days and 0 cost.
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
        # Valve chambers: every 500m for water, every 300m for sewerage (CPHEEO)
        pipeline_type = str(getattr(project_input, "pipeline_type", "water") or "water").lower()
        default_valve_spacing = 500.0 if pipeline_type == "water" else 300.0
        self.valve_chamber_spacing_m = float(params.get("valve_chamber_spacing_m", default_valve_spacing))
        # Thrust blocks are counted per bend/tee (guideline unit is Nos), so no
        # concrete-volume assumption is carried here.
        # Road-crossing trenchless: explicit crossing counts win; otherwise a
        # trenchless installation takes this share of the alignment.
        self.trenchless_ratio = float(params.get("trenchless_ratio", 0.05))

        # --- Scheme-level counts -------------------------------------------
        # ProjectInput wins; pipeline_project_inputs supplies the configured
        # defaults. A value of 0 leaves that package out of scope.
        inputs = self.assumptions.get("pipeline_project_inputs", {}) or {}

        def _count(attr: str, key: str, cast):
            raw = getattr(project_input, attr, 0) or 0
            if not raw:
                raw = inputs.get(key) or 0
            try:
                return cast(raw or 0)
            except (TypeError, ValueError):
                return cast(0)

        self.num_villages = _count("num_villages", "num_villages", int)
        self.num_oht = _count("num_oht", "num_oht", int)
        self.wtp_capacity_mld = _count("wtp_capacity_mld", "wtp_capacity_mld", float)
        self.num_ibps = _count("num_ibps", "num_ibps", int)
        self.num_fhtc = _count("num_fhtc", "num_fhtc", int)

        self.number_of_valves = int(inputs.get("number_of_valves", 0) or 0)
        self.number_of_air_valves = int(inputs.get("number_of_air_valves", 0) or 0)
        self.number_of_scour_valves = int(inputs.get("number_of_scour_valves", 0) or 0)
        self.number_of_thrust_blocks = int(inputs.get("number_of_thrust_blocks", 0) or 0)
        self.number_of_road_crossings = _count(
            "number_of_road_crossings", "number_of_road_crossings", int)
        self.number_of_highway_crossings = _count(
            "number_of_highway_crossings", "number_of_highway_crossings", int)
        self.number_of_rail_crossings = _count(
            "number_of_rail_crossings", "number_of_rail_crossings", int)
        self.number_of_hdd_crossings = _count(
            "number_of_hdd_crossings", "number_of_hdd_crossings", int)
        self.number_of_river_crossings = _count(
            "number_of_river_crossings", "number_of_river_crossings", int)
        self.number_of_service_connections = int(
            inputs.get("number_of_service_connections", 0) or 0)
        self.installation_method = str(
            getattr(project_input, "installation_method", "open_trench")
            or "open_trench").lower()

        # Crossings and facility scaling lengths
        self.reference_wtp_mld = max(1.0, float(params.get("reference_wtp_mld", 20.0) or 20.0))
        self.hdd_small_length_m = float(params.get("hdd_small_length_m", 40.0) or 40.0)
        self.hdd_large_length_m = float(params.get("hdd_large_length_m", 60.0) or 60.0)
        self.highway_rail_crossing_length_m = float(
            params.get("highway_rail_crossing_length_m", 25.0) or 25.0)
        self.river_crossing_length_m = float(
            params.get("river_crossing_length_m", 60.0) or 60.0)
        self.road_dismantle_length_m = float(
            params.get("road_dismantle_length_m", 20.0) or 20.0)
        self.boundary_wall_length_m = float(
            params.get("boundary_wall_length_m", 500.0) or 500.0)
        self.internal_road_sqm_per_mld = float(
            params.get("internal_road_sqm_per_mld", 100.0) or 100.0)
        self.storm_drain_length_m = float(
            params.get("storm_drain_length_m", 300.0) or 300.0)
        self.earth_fill_cum_per_mld = float(
            params.get("earth_fill_cum_per_mld", 150.0) or 150.0)
        self.ht_lt_cable_length_m = float(
            params.get("ht_lt_cable_length_m", 500.0) or 500.0)

        # Excavation strata split. The three rock/moorum shares are carved out
        # of the same trench volume, not added on top of it.
        shares = params.get("excavation_strata_shares", {}) or {}
        self.ordinary_share = float(shares.get("ordinary", 0.90) or 0.90)
        self.moorum_share = float(shares.get("moorum", 0.07) or 0.07)
        self.soft_rock_share = float(shares.get("soft_rock", 0.02) or 0.02)
        self.hard_rock_share = float(shares.get("hard_rock", 0.01) or 0.01)

    def calculate_quantities(self) -> Dict[str, Dict[str, Any]]:
        q: Dict[str, Dict[str, Any]] = {}
        Lkm = self.length_km
        Lm = self.length_m
        tw = self.trench_width_m
        td = self.trench_depth_m

        # ------------------------------------------------------------------
        # Scheme counts and derived facility units
        # ------------------------------------------------------------------
        villages = self.num_villages
        oht = self.num_oht
        wtp_mld = self.wtp_capacity_mld
        ibps = self.num_ibps
        fhtc = self.num_fhtc

        # A village count is the single switch that turns on a full scheme:
        # derive the facilities it implies unless they were supplied explicitly.
        if villages > 0:
            if oht <= 0:
                oht = max(1, int(math.ceil(villages / 3.0)))
            if fhtc <= 0:
                fhtc = villages * 200
            if wtp_mld <= 0:
                # A rural multi-village scheme designs around 0.5 MLD per
                # village (domestic + livestock + institutional demand with a
                # planning margin), not 2 MLD. The minimum keeps a single
                # compact treatment train rather than a fractional plant.
                wtp_mld = max(1.0, villages * 0.5)
        if wtp_mld > 0 and ibps <= 0:
            ibps = max(1, int(math.ceil(wtp_mld / self.reference_wtp_mld)))
        if fhtc <= 0 and self.number_of_service_connections > 0:
            fhtc = self.number_of_service_connections

        # Persist the effective scope so the report shows the facilities that
        # were actually priced, not just the raw user inputs.
        self.num_oht = oht
        self.wtp_capacity_mld = wtp_mld
        self.num_ibps = ibps
        self.num_fhtc = fhtc

        wtp_trains = wtp_mld / self.reference_wtp_mld if wtp_mld > 0 else 0.0
        cwr_units = wtp_trains
        has_wtp = wtp_mld > 0

        # ------------------------------------------------------------------
        # Linear corridor quantities
        # ------------------------------------------------------------------
        trench_volume = Lm * tw * td

        # Strata partition of the SAME trench volume.
        ordinary_0_1_5 = Lm * tw * min(td, 1.5) * self.ordinary_share
        ordinary_1_5_3 = Lm * tw * max(0.0, td - 1.5) * self.ordinary_share
        moorum_vol = trench_volume * self.moorum_share
        soft_rock_vol = trench_volume * self.soft_rock_share
        hard_rock_vol = trench_volume * self.hard_rock_share

        # Valves / chambers / crossings: configured counts win, otherwise
        # derive only the counts the alignment itself implies. A rail, river,
        # highway or HDD crossing is NEVER invented when the user has not
        # declared one: those are site-specific and drive large quantities.
        sluice_count = self.number_of_valves or max(2, int(math.ceil(Lm / self.valve_chamber_spacing_m)))
        air_count = self.number_of_air_valves or max(1, int(math.ceil(Lm / 1000.0)))
        scour_count = self.number_of_scour_valves or max(1, int(math.ceil(Lm / 1000.0)))
        valve_total = sluice_count + air_count + scour_count
        thrust_count = self.number_of_thrust_blocks or max(2, int(math.ceil(Lm / 250.0)))
        road_cross = self.number_of_road_crossings or max(1, int(math.ceil(Lkm / 2.0)))
        highway_cross = self.number_of_highway_crossings
        rail_cross = self.number_of_rail_crossings
        river_cross = self.number_of_river_crossings

        # Pipe string: only the selected material/size class carries length.
        mat = self.pipe_material
        dia = self.diameter_mm
        hdpe_small = Lm if mat == "HDPE" and dia <= 110 else 0.0
        hdpe_medium = Lm if mat == "HDPE" and 110 < dia <= 200 else 0.0
        hdpe_large = Lm if mat == "HDPE" and dia > 200 else 0.0
        hdpe_total = hdpe_small + hdpe_medium + hdpe_large
        dik7 = Lm if mat == "DI" and dia <= 300 else 0.0
        dik9_small_med = Lm if mat == "DI" and 300 < dia <= 350 else 0.0
        dik9_large = Lm if mat in ("DI", "RCC") and (dia > 350 or mat == "RCC") else 0.0
        dik9_total = dik9_small_med + dik9_large
        ms_line = Lm if mat == "MS" else 0.0

        # HDD: an explicit crossing count wins. Otherwise a trenchless
        # installation takes `trenchless_ratio` of the alignment; open-trench
        # work carries no HDD length at all.
        if self.number_of_hdd_crossings > 0:
            n = self.number_of_hdd_crossings
            hdd_small_qty = n * self.hdd_small_length_m if dia <= 150 else 0.0
            hdd_large_qty = n * self.hdd_large_length_m if dia > 150 else 0.0
            hdd_formula = f"{n} declared HDD crossing(s)"
        elif self.installation_method in ("trenchless", "hdd"):
            hdd_total = Lm * self.trenchless_ratio
            hdd_small_qty = hdd_total if dia <= 150 else 0.0
            hdd_large_qty = hdd_total if dia > 150 else 0.0
            hdd_formula = (f"{self.trenchless_ratio:.1%} of {Lkm:g} km alignment "
                           f"= {hdd_total:,.0f} m ({self.installation_method})")
        else:
            hdd_small_qty = hdd_large_qty = 0.0
            hdd_formula = "No HDD: open-trench installation and no declared crossings"

        # Scheme facility quantities (0 when the package is out of scope).
        highway_qty = (highway_cross + rail_cross) * self.highway_rail_crossing_length_m
        river_qty = river_cross * self.river_crossing_length_m
        # Only open-cut road crossings break and restore the road surface.
        # NH/SH and rail crossings are trenchless (ACT-26), so they carry no
        # dismantling, restoration or RCC encasing quantity.
        dismantle_qty = road_cross * self.road_dismantle_length_m

        boundary_wall_qty = self.boundary_wall_length_m if has_wtp else 0.0
        gates_qty = (2 * wtp_trains if has_wtp else 0.0) + oht
        internal_road_qty = wtp_mld * self.internal_road_sqm_per_mld if has_wtp else 0.0
        storm_drain_qty = self.storm_drain_length_m if has_wtp else 0.0
        earth_fill_qty = wtp_mld * self.earth_fill_cum_per_mld if has_wtp else 0.0
        allied_civil_qty = max(1.0, wtp_trains) if has_wtp else 0.0
        pump_sets_qty = (2 * ibps + 2 * wtp_trains) if (ibps > 0 or has_wtp) else 0.0
        substation_qty = (wtp_trains + ibps) if (ibps > 0 or has_wtp) else 0.0
        transformer_qty = (ibps + wtp_trains) if (ibps > 0 or has_wtp) else 0.0
        cabling_qty = (self.ht_lt_cable_length_m * (wtp_trains + ibps)
                       if (ibps > 0 or has_wtp) else 0.0)
        dg_qty = (ibps + wtp_trains) if (ibps > 0 or has_wtp) else 0.0
        scada_qty = (wtp_trains + ibps) if (ibps > 0 or has_wtp) else 0.0

        # ------------------------------------------------------------------
        # 1. Preliminary works & mobilization
        # ------------------------------------------------------------------
        q["survey_setting_out"] = {
            "quantity": round(Lkm, 2), "unit": "km",
            "formula": "Pipeline length (km)"}
        q["utility_mapping"] = {
            "quantity": round(Lkm, 2), "unit": "km",
            "formula": "Pipeline length (km) — utility corridor survey"}
        q["site_mobilization"] = {
            "quantity": 1.0, "unit": "set",
            "formula": f"One site establishment for {max(1, int(getattr(self.input, 'number_of_work_fronts', 1) or 1))} front(s)"}

        # ------------------------------------------------------------------
        # 2. Pipeline trenching & earthwork
        # ------------------------------------------------------------------
        q["excavation_ordinary_0_1_5"] = {
            "quantity": round(ordinary_0_1_5, 2), "unit": "cum",
            "formula": "L × trench width × min(depth,1.5) × ordinary share"}
        q["excavation_ordinary_1_5_3"] = {
            "quantity": round(ordinary_1_5_3, 2), "unit": "cum",
            "formula": "L × trench width × max(depth−1.5,0) × ordinary share"}
        q["excavation_moorum"] = {
            "quantity": round(moorum_vol, 2), "unit": "cum",
            "formula": f"Trench volume × {self.moorum_share:.0%} moorum share"}
        q["excavation_soft_rock"] = {
            "quantity": round(soft_rock_vol, 2), "unit": "cum",
            "formula": f"Trench volume × {self.soft_rock_share:.0%} soft-rock share"}
        q["excavation_hard_rock"] = {
            "quantity": round(hard_rock_vol, 2), "unit": "cum",
            "formula": f"Trench volume × {self.hard_rock_share:.0%} hard-rock share"}

        # ------------------------------------------------------------------
        # 3. Pipe laying, testing & commissioning by material
        # ------------------------------------------------------------------
        pipe_rows = [
            ("hdpe_laying_small", hdpe_small, "HDPE pipe length ≤ DN110 (m)"),
            ("hdpe_laying_medium", hdpe_medium, "HDPE pipe length DN125–DN200 (m)"),
            ("hdpe_laying_large", hdpe_large, "HDPE pipe length > DN200 (m)"),
            ("hdpe_backfill_test", hdpe_total, "Active HDPE length (m)"),
            ("hdpe_commissioning", hdpe_total, "Active HDPE length (m)"),
            ("dik7_laying", dik7, "DI K7 pipe length ≤ DN300 (m)"),
            ("dik7_backfill_test", dik7, "Active DI K7 length (m)"),
            ("dik7_commissioning", dik7, "Active DI K7 length (m)"),
            ("dik9_laying_small_med", dik9_small_med, "DI K9 pipe length DN300–DN350 (m)"),
            ("dik9_laying_large", dik9_large, "DI K9 pipe length > DN350 (m)"),
            ("dik9_backfill_test", dik9_total, "Active DI K9 length (m)"),
            ("dik9_commissioning", dik9_total, "Active DI K9 length (m)"),
            ("ms_laying", ms_line, "MS pipe length (m)"),
            ("ms_backfill_test", ms_line, "Active MS length (m)"),
            ("ms_commissioning", ms_line, "Active MS length (m)"),
        ]
        for key, qty, formula in pipe_rows:
            q[key] = {"quantity": round(qty, 2), "unit": "Mtr", "formula": formula}

        # ------------------------------------------------------------------
        # 4. Trenchless, road and river crossings
        # ------------------------------------------------------------------
        q["hdd_small"] = {"quantity": round(hdd_small_qty, 2), "unit": "Mtr",
                          "formula": f"{hdd_formula} — ≤ DN150"}
        q["hdd_large"] = {"quantity": round(hdd_large_qty, 2), "unit": "Mtr",
                          "formula": f"{hdd_formula} — > DN150"}
        q["highway_rail_crossing"] = {"quantity": round(highway_qty, 2), "unit": "Mtr",
                                       "formula": f"({highway_cross} NH/SH + {rail_cross} rail) declared crossings × {self.highway_rail_crossing_length_m:g} m"}
        q["river_bridge_crossing"] = {"quantity": round(river_qty, 2), "unit": "Mtr",
                                       "formula": f"{river_cross} declared river/nallah crossing(s) × {self.river_crossing_length_m:g} m"}
        q["dismantling_roads"] = {"quantity": round(dismantle_qty, 2), "unit": "Mtr",
                                   "formula": f"{road_cross} open-cut road crossing(s) × {self.road_dismantle_length_m:g} m (trenchless NH/SH/rail excluded)"}
        q["restoration_roads"] = {"quantity": round(dismantle_qty, 2), "unit": "Mtr",
                                   "formula": "Same length as dismantled open-cut road surface"}

        # ------------------------------------------------------------------
        # 5. House service connections
        # ------------------------------------------------------------------
        q["fhtc"] = {"quantity": float(fhtc), "unit": "Nos",
                     "formula": f"{villages} village(s) × 200 FHTC/village"
                                if (fhtc and villages and fhtc == villages * 200)
                                else "Configured household connections"}

        # ------------------------------------------------------------------
        # 6. Valves, chambers and thrust blocks
        # ------------------------------------------------------------------
        for key, desc in (
            ("valve_supply", "Supply of sluice/air/scour valves"),
            ("valve_install", "Installation of valves"),
            ("dismantling_joints", "Dismantling joints"),
            ("valve_test_comm", "Valve testing & commissioning"),
        ):
            q[key] = {"quantity": float(valve_total), "unit": "Each",
                      "formula": f"{sluice_count} sluice + {air_count} air + {scour_count} scour valves"}

        q["chamber_scour"] = {"quantity": float(scour_count), "unit": "Nos",
                               "formula": f"{scour_count} scour valve chambers"}
        q["chamber_sluice"] = {"quantity": float(sluice_count), "unit": "Nos",
                                "formula": f"{sluice_count} sluice valve chambers"}
        q["chamber_air"] = {"quantity": float(air_count), "unit": "Nos",
                             "formula": f"{air_count} air valve chambers"}
        q["thrust_blocks"] = {"quantity": float(thrust_count), "unit": "Nos",
                               "formula": f"One thrust block per 250 m ({thrust_count} total)"}
        q["pipe_encasing"] = {"quantity": float(road_cross),
                               "unit": "Nos",
                               "formula": f"{road_cross} open-cut road crossing(s); trenchless NH/SH/rail use their own casing"}

        # ------------------------------------------------------------------
        # 7. Overhead service reservoirs
        # ------------------------------------------------------------------
        oht_desc = {
            "oht_foundation": "ESR foundations",
            "oht_staging": "ESR staging columns & bracing",
            "oht_stairs": "ESR staircase",
            "oht_dome_ring": "ESR ring beam / domes / balcony",
            "oht_cylindrical_wall": "ESR cylindrical wall",
            "oht_roof_dome": "ESR roof dome",
            "oht_piping_finishing": "ESR piping, ladder & finishing",
            "oht_testing_comm": "ESR testing & commissioning",
        }
        for key, desc in oht_desc.items():
            q[key] = {"quantity": float(oht), "unit": "Nos",
                      "formula": f"{oht} OHT/ESR(s) — {desc}"}

        # ------------------------------------------------------------------
        # 8. Intake well cum pump house
        # ------------------------------------------------------------------
        intake_desc = {
            "intake_cofferdam": "coffer dam",
            "intake_foundation_staining": "foundation & well staining",
            "intake_casting_sinking": "well casting & sinking",
            "intake_discharge_floor": "discharge floor",
            "intake_bridge_foundation": "approach bridge foundation / pier",
            "intake_bridge_deck": "approach bridge deck",
            "intake_pumphouse": "pump house to roof slab",
            "intake_fixing": "inlet port, valves, specials & gantry",
        }
        for key, desc in intake_desc.items():
            q[key] = {"quantity": float(ibps), "unit": "Job",
                      "formula": f"{ibps} intake/pumping station(s) — {desc}"}

        # ------------------------------------------------------------------
        # 9. Clear water reservoir & pump house
        # ------------------------------------------------------------------
        cwr_desc = {
            "cwr_earthwork_pcc": "earthwork & PCC bed",
            "cwr_raft": "raft foundation",
            "cwr_wall": "CWR walls",
            "cwr_roof": "roof slab",
            "cwr_finishing": "painting & finishing",
            "cwr_pipe_fittings": "pipe & fittings",
            "cwr_test_comm": "testing & commissioning",
        }
        for key, desc in cwr_desc.items():
            q[key] = {"quantity": round(cwr_units, 3), "unit": "Job",
                      "formula": f"{wtp_mld:g} MLD / {self.reference_wtp_mld:g} MLD reference — {desc}"}

        # ------------------------------------------------------------------
        # 10. Water treatment plant — civil
        # ------------------------------------------------------------------
        wtp_civil_desc = {
            "wtp_cascade_aerator": "cascade aerator / inlet chamber",
            "wtp_flash_mixer": "flash mixer & inlet channel",
            "wtp_clariflocculator_civil": "pre-settling tank / clariflocculator civil",
            "wtp_filter_house": "filter house",
            "wtp_wash_water_tank": "wash water tank",
            "wtp_chlorination": "chlorination block & neutralisation tank",
            "wtp_admin_lab_store": "administrative block, lab & store",
            "wtp_recycling": "recycling arrangement",
            "wtp_road_site_dev": "internal roads & site development",
            "wtp_test_comm": "complete plant testing & commissioning",
        }
        for key, desc in wtp_civil_desc.items():
            q[key] = {"quantity": round(wtp_trains, 3), "unit": "Job",
                      "formula": f"{wtp_mld:g} MLD / {self.reference_wtp_mld:g} MLD reference — {desc}"}

        # ------------------------------------------------------------------
        # 11. Water treatment plant — electro-mechanical
        # ------------------------------------------------------------------
        wtp_em_desc = {
            "wtp_em_clariflocculator": "clariflocculator mechanism, agitator & gates",
            "wtp_em_pumps_blower": "wash-water / backwash pumps & air blower",
            "wtp_em_dosing": "dosing pumps & alum/PE/lime tanks",
            "wtp_em_electrical_scada": "electrical panel, lighting, instrumentation & PLC/SCADA",
        }
        for key, desc in wtp_em_desc.items():
            q[key] = {"quantity": round(wtp_trains, 3), "unit": "Job",
                      "formula": f"{wtp_mld:g} MLD / {self.reference_wtp_mld:g} MLD reference — {desc}"}

        # ------------------------------------------------------------------
        # 12. Boundary wall, gates & ancillary civil
        # ------------------------------------------------------------------
        q["boundary_wall"] = {"quantity": round(boundary_wall_qty, 2), "unit": "Mtr",
                               "formula": "Configured WTP campus perimeter"}
        q["ms_gates"] = {"quantity": round(gates_qty, 2), "unit": "Nos",
                          "formula": "2 gates per WTP train + 1 per OHT"}
        q["internal_road_campus"] = {"quantity": round(internal_road_qty, 2), "unit": "Sqm",
                                      "formula": f"{self.internal_road_sqm_per_mld:g} sqm per MLD"}
        q["storm_water_drain"] = {"quantity": round(storm_drain_qty, 2), "unit": "Mtr",
                                   "formula": "Configured campus drain length"}
        q["earth_sand_filling"] = {"quantity": round(earth_fill_qty, 2), "unit": "Cum",
                                    "formula": f"{self.earth_fill_cum_per_mld:g} cum per MLD"}
        q["allied_civil_structures"] = {"quantity": round(allied_civil_qty, 2), "unit": "Nos",
                                         "formula": "RTU/chlorination/guard/DG/meter rooms & staff quarters"}
        q["allied_electrical_finish"] = {"quantity": round(allied_civil_qty, 2), "unit": "Nos",
                                          "formula": "Electrical finishing for the allied buildings"}

        # ------------------------------------------------------------------
        # 13. Electrical & mechanical — pumping & power
        # ------------------------------------------------------------------
        q["pump_sets"] = {"quantity": round(pump_sets_qty, 2), "unit": "Nos",
                           "formula": "2 pump sets per intake/booster + 2 per WTP train"}
        q["ht_substation"] = {"quantity": round(substation_qty, 2), "unit": "Set",
                               "formula": "One HT substation per WTP train and per intake/booster station"}
        q["power_transformers"] = {"quantity": round(transformer_qty, 2), "unit": "Nos",
                                    "formula": "One transformer per intake/booster and per WTP train"}
        q["ht_lt_cabling"] = {"quantity": round(cabling_qty, 2), "unit": "Mtr",
                               "formula": f"{self.ht_lt_cable_length_m:g} m per WTP train and per intake/booster station"}
        q["dg_set"] = {"quantity": round(dg_qty, 2), "unit": "Nos",
                        "formula": "One standby DG per intake/booster and per WTP train"}
        q["scada_automation"] = {"quantity": round(scada_qty, 2), "unit": "Set",
                                  "formula": "One SCADA/PLC integration per WTP train and per intake/booster station"}

        return q



# ─────────────────────────────────────────────────────────────────────────────
# Building quantity calculator
# ─────────────────────────────────────────────────────────────────────────────

class BuildingQuantityCalculator:
    """
    Computes the 40-activity low-rise building quantities from project inputs
    and building_assumptions.yaml (building_parameters section).

    All 40 keys in building_guidelines.json are emitted. Activities that are
    genuinely out of scope for the selected building (for example roof trusses
    on an RCC roof, or upper-floor work in a single-storey building) are
    returned with a zero quantity. They stay in the catalogue so the workbook
    and CPM keep the same activity set for every run.

    Quantities are derived from built-up area, floors, and standard CPWD DSR
    2021 / IS 456 ratios, not from per-activity placeholders.
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
        self.foundation_type = str(
            getattr(project_input, "foundation_type", "pile") or "pile"
        ).strip().lower()

        # Total floor plate (all blocks combined).
        self.floor_area_sqm = self.bua_sqm / self.floors

        # Building parameters from assumptions
        params = self.assumptions.get("building_parameters", {})

        # Structural ratios (per sqm of BUA) — CPWD DSR 2021 / IS 456 based.
        self.foundation_rcc_per_sqm = float(params.get("foundation_rcc_per_sqm", 0.35))
        self.frame_rcc_per_sqm = float(params.get("frame_rcc_per_sqm", 0.40))
        self.masonry_per_sqm = float(params.get("masonry_per_sqm", 0.18))
        self.plaster_area_factor = float(params.get("plaster_area_factor", 3.5))
        self.waterproofing_area_factor = float(params.get("waterproofing_area_factor", 0.30))
        self.flooring_area_factor = float(params.get("flooring_area_factor", 1.0))
        self.painting_area_factor = float(params.get("painting_area_factor", 4.0))
        self.basement_excavation_depth_m = float(params.get("basement_excavation_depth_m", 3.5))
        self.external_works_factor = float(params.get("external_works_factor", 0.5))
        self.far = float(params.get("floor_area_ratio", 1.5))

        # Geometry and take-off factors. These are explicit so the estimate can
        # be tuned without changing code.
        self.floor_height_m = float(params.get("floor_height_m", 3.2))
        self.external_wall_opening_factor = float(params.get("external_wall_opening_factor", 0.20))
        self.pile_spacing_sqm = float(params.get("pile_spacing_sqm", 12.0))
        self.pile_depth_base_m = float(params.get("pile_depth_base_m", 10.0))
        self.pile_depth_per_floor_m = float(params.get("pile_depth_per_floor_m", 0.75))
        self.pile_cap_cum_per_pile = float(params.get("pile_cap_cum_per_pile", 1.0))
        self.column_stump_cum_per_pile = float(params.get("column_stump_cum_per_pile", 0.6))
        self.ground_beam_cum_per_m = float(params.get("ground_beam_cum_per_m", 0.18))
        self.plinth_fill_depth_m = float(params.get("plinth_fill_depth_m", 0.30))
        self.strip_excavation_depth_m = float(params.get("strip_excavation_depth_m", 0.30))
        self.door_per_sqm = float(params.get("door_per_sqm", 0.04))
        self.window_per_sqm = float(params.get("window_per_sqm", 0.05))
        self.electrical_points_per_sqm = float(params.get("electrical_points_per_sqm", 0.5))
        self.plumbing_m_per_sqm = float(params.get("plumbing_m_per_sqm", 1.0))
        self.sanitary_per_sqm = float(params.get("sanitary_per_sqm", 0.02))
        self.hardware_per_sqm = float(params.get("hardware_per_sqm", 0.01))
        self.railing_m_per_floor = float(params.get("railing_m_per_floor", 5.0))
        self.tiling_area_factor = float(params.get("tiling_area_factor", 1.15))
        self.roof_truss_mt_per_sqm = float(params.get("roof_truss_mt_per_sqm", 0.05))
        self.curtain_wall_factor_commercial = float(params.get("curtain_wall_factor_commercial", 0.35))
        self.curtain_wall_factor_premium = float(params.get("curtain_wall_factor_premium", 0.10))

        # Plot area: FAR relates total BUA to plot area. Do NOT divide by the
        # number of floors again — that error understated the site by `floors`
        # and made clearance/external works shrink as the building got taller.
        self.site_area_sqm = round(
            max(self.floor_area_sqm, self.bua_sqm / max(self.far, 1e-9)), 1
        )

        # Total perimeter of all blocks (square blocks plus an irregularity
        # allowance). Multi-block sites have more external wall than one block.
        self.perimeter_m = round(
            4.0 * math.sqrt(max(1.0, self.floor_area_sqm * self.blocks)) * 1.15, 1
        )
        self.external_wall_sqm = round(
            self.perimeter_m * self.floor_height_m * self.floors
            * (1.0 - self.external_wall_opening_factor), 1
        )

        # Construction-type adjustments. The catalogue is an RCC frame, so
        # load-bearing and steel buildings scale the applicable portions rather
        # than silently keeping an RCC quantity.
        self.frame_factor = {"rcc": 1.0, "steel": 0.60, "load_bearing": 0.45}.get(
            self.construction_type, 1.0)
        self.masonry_factor = {"rcc": 1.0, "steel": 0.60, "load_bearing": 1.35}.get(
            self.construction_type, 1.0)
        self.pile_factor = {"rcc": 1.0, "steel": 0.80, "load_bearing": 0.0}.get(
            self.construction_type, 1.0)

        # Quality multipliers on finishing quantities
        quality_map = {"basic": 0.8, "standard": 1.0, "premium": 1.3}
        self.quality_mult = quality_map.get(self.quality, 1.0)

    def calculate_quantities(self) -> Dict[str, Dict[str, Any]]:
        q: Dict[str, Dict[str, Any]] = {}

        def row(quantity: float, unit: str, formula: str) -> Dict[str, Any]:
            return {"quantity": round(max(0.0, float(quantity)), 1),
                    "unit": unit, "formula": formula}

        bua = self.bua_sqm
        floors = self.floors
        footprint = self.floor_area_sqm
        site = self.site_area_sqm
        perimeter = self.perimeter_m
        external_wall = self.external_wall_sqm
        quality = self.quality_mult

        # --- Frame RCC: allocate the total BUA-based volume across the
        # catalogue's floor-specific activities. Extra floors beyond the G+2
        # template are folded into the top-floor beams/columns so the total
        # concrete volume remains correct.
        frame_rcc = bua * self.frame_rcc_per_sqm * self.frame_factor
        frame_weights: Dict[str, float] = {
            "gf_columns": 0.0, "ff_beams_slab": 0.0, "gf_stairs": 0.0,
            "ff_columns": 0.0, "sf_beams_slab": 0.0, "ff_stairs": 0.0,
            "sf_columns": 0.0, "roof_beams": 0.0,
        }
        if floors >= 1:
            frame_weights["gf_columns"] = 1.0
            frame_weights["roof_beams"] = 1.5
        if floors >= 2:
            frame_weights["ff_beams_slab"] = 2.0
            frame_weights["gf_stairs"] = 0.4
            frame_weights["ff_columns"] = 1.0
        if floors >= 3:
            frame_weights["sf_beams_slab"] = 2.0
            frame_weights["ff_stairs"] = 0.4
            frame_weights["sf_columns"] = 1.0
        if floors > 3 and getattr(self.input, "floor_wise_structural_detail", False):
            for floor_no in range(4, floors + 1):
                frame_weights[f"floor_{floor_no}_beams_slab"] = 3.0
                frame_weights[f"floor_{floor_no}_columns"] = 1.0

        weight_total = sum(frame_weights.values()) or 1.0
        frame_q = {k: frame_rcc * w / weight_total for k, w in frame_weights.items()}

        # --- Masonry: split by storey, with any extra floors assigned to the
        # top-floor wall row.
        masonry = bua * self.masonry_per_sqm * self.masonry_factor
        if floors == 1:
            gf_walls, ff_walls, sf_walls = masonry, 0.0, 0.0
        elif floors == 2:
            gf_walls = ff_walls = masonry / 2.0
            sf_walls = 0.0
        else:
            gf_walls = ff_walls = masonry / floors
            sf_walls = masonry * (floors - 2) / floors

        # --- Piles and foundation. One pile per `pile_spacing_sqm` of floor
        # plate; depth grows with the number of storeys.
        pile_count = 0
        if self.pile_factor > 0 and self.foundation_type == "pile":
            pile_count = max(12, int(math.ceil(footprint / max(self.pile_spacing_sqm, 1e-9))))
        pile_depth = 0.0
        if pile_count:
            pile_depth = max(
                8.0,
                self.pile_depth_base_m + self.pile_depth_per_floor_m * max(0, floors - 1),
            )
        piling_m = pile_count * pile_depth
        pile_tests = max(3, int(math.ceil(pile_count * 0.02))) if pile_count else 0
        footing_count = max(4, int(math.ceil(footprint / 25.0)))
        foundation_volume = (
            footprint * self.foundation_rcc_per_sqm
            if self.foundation_type == "raft"
            else footing_count * 0.80
        )

        # --- Roof and envelope.
        roof_area = footprint * 1.05
        roof_trusses_mt = (
            roof_area * self.roof_truss_mt_per_sqm
            if self.construction_type == "steel" else 0.0
        )
        if self.building_type == "commercial":
            curtain_factor = self.curtain_wall_factor_commercial
        elif self.quality == "premium":
            curtain_factor = self.curtain_wall_factor_premium
        else:
            curtain_factor = 0.0

        # --- Services. Quantities are in the units used by the catalogue:
        # electrical is "points", plumbing is "metre", sanitary/hardware are
        # "Nos". The old calculator returned sqm for all of them.
        commercial = self.building_type == "commercial"
        electrical_points = bua * self.electrical_points_per_sqm * (1.6 if commercial else 1.0) * quality
        plumbing_m = bua * self.plumbing_m_per_sqm * (1.2 if commercial else 1.0) * quality
        sanitary_count = max(4, int(round(bua * self.sanitary_per_sqm))) * (1.2 if commercial else 1.0)
        hardware_count = max(1, int(round(bua * self.hardware_per_sqm)))
        door_count = max(2, int(round(bua * self.door_per_sqm)))
        window_count = max(4, int(round(bua * self.window_per_sqm)))

        # --- Finishes.
        external_render = external_wall * quality
        internal_plaster = max(0.0, bua * self.plaster_area_factor * quality - external_render)
        internal_painting = max(0.0, bua * self.painting_area_factor * quality - external_render)
        tiling = bua * self.tiling_area_factor * quality
        waterproofing = bua * self.waterproofing_area_factor

        # --- Earthwork and site.
        excavation = max(
            100.0,
            site * self.strip_excavation_depth_m
            + (footprint * self.basement_excavation_depth_m if self.basement else 0.0),
        )

        q["site_setting_out"] = row(site, "sqm", "Plot area = BUA / FAR")
        q["site_clearance"] = row(site, "sqm", "Plot area = BUA / FAR")
        q["site_mobilization"] = row(1.0, "set", "One mobilisation set")
        q["excavation"] = row(
            excavation, "cum",
            "Plot stripping + basement excavation (if any)")
        q["piling"] = row(
            piling_m, "metre",
            f"{pile_count} piles x {pile_depth:.2f} m depth")
        q["pile_testing"] = row(
            float(pile_tests), "Nos",
            "Initial/routine tests = max(3, 2% of piles)")
        q["pile_head_chipping"] = row(
            float(pile_count), "Nos", "One per pile")
        q["pile_caps"] = row(
            pile_count * self.pile_cap_cum_per_pile
            if self.foundation_type == "pile" else foundation_volume,
            "cum",
            (f"{pile_count} piles x {self.pile_cap_cum_per_pile} cum/pile"
             if self.foundation_type == "pile"
             else f"{self.foundation_type} foundation allowance"))
        q["column_stumps"] = row(
            pile_count * self.column_stump_cum_per_pile, "cum",
            f"{pile_count} piles x {self.column_stump_cum_per_pile} cum/pile")
        q["ground_beams"] = row(
            perimeter * self.ground_beam_cum_per_m, "cum",
            f"Perimeter {perimeter:.1f} m x {self.ground_beam_cum_per_m} cum/m")
        q["plinth_filling"] = row(
            footprint * self.plinth_fill_depth_m, "cum",
            f"Floor plate x {self.plinth_fill_depth_m} m fill")
        q["anti_termite"] = row(footprint, "sqm", "Ground-floor plate area")
        q["gf_slab"] = row(footprint, "sqm", "Ground-floor slab-on-grade area")

        q["gf_columns"] = row(frame_q["gf_columns"], "cum", "Frame RCC allocation")
        q["ff_beams_slab"] = row(frame_q["ff_beams_slab"], "cum", "Frame RCC allocation")
        q["gf_stairs"] = row(frame_q["gf_stairs"], "cum", "Frame RCC allocation")
        q["gf_walls"] = row(gf_walls, "cum", "Masonry split by storey")
        q["ff_columns"] = row(frame_q["ff_columns"], "cum", "Frame RCC allocation")
        q["sf_beams_slab"] = row(frame_q["sf_beams_slab"], "cum", "Frame RCC allocation")
        q["ff_stairs"] = row(frame_q["ff_stairs"], "cum", "Frame RCC allocation")
        q["ff_walls"] = row(ff_walls, "cum", "Masonry split by storey")
        q["sf_columns"] = row(frame_q["sf_columns"], "cum", "Frame RCC allocation")
        q["roof_beams"] = row(frame_q["roof_beams"], "cum", "Frame RCC allocation")
        if getattr(self.input, "floor_wise_structural_detail", False):
            for floor_no in range(4, floors + 1):
                q[f"floor_{floor_no}_beams_slab"] = row(
                    frame_q.get(f"floor_{floor_no}_beams_slab", 0.0),
                    "cum", f"Explicit floor {floor_no} frame RCC allocation")
                q[f"floor_{floor_no}_columns"] = row(
                    frame_q.get(f"floor_{floor_no}_columns", 0.0),
                    "cum", f"Explicit floor {floor_no} frame RCC allocation")
        q["roof_trusses"] = row(
            roof_trusses_mt, "MT",
            "Steel roof only; zero for an RCC roof")
        q["roof_covering"] = row(roof_area, "sqm", "Roof plate + 5% overhang")
        q["sf_walls"] = row(sf_walls, "cum", "Masonry split by storey")
        q["curtain_wall"] = row(
            external_wall * curtain_factor, "sqm",
            "Commercial/premium facade glazing")
        q["railings"] = row(
            floors * self.railing_m_per_floor + 0.25 * perimeter, "metre",
            "Staircase + balcony railing allowance")
        q["doors"] = row(float(door_count), "Nos", "BUA x doors/sqm")
        q["windows"] = row(float(window_count), "Nos", "BUA x windows/sqm")
        q["electrical"] = row(electrical_points, "points", "BUA x points/sqm x quality")
        q["plumbing"] = row(plumbing_m, "metre", "BUA x pipe metre/sqm x quality")
        q["sanitary_fittings"] = row(sanitary_count, "Nos", "BUA x fixtures/sqm")
        q["internal_plaster"] = row(internal_plaster, "sqm", "Plaster factor less external render")
        q["tiling"] = row(tiling, "sqm", "BUA x tiling factor x quality")
        q["hardware"] = row(float(hardware_count), "Nos", "BUA x hardware units/sqm")
        q["internal_painting"] = row(internal_painting, "sqm", "Paint factor less external render")
        q["external_render"] = row(external_render, "sqm", "External wall area x quality")
        q["external_works"] = row(site * self.external_works_factor * quality, "sqm", "Plot area x external-works factor")
        q["cleaning_handover"] = row(bua, "sqm", "Built-up area")

        return q
