"""
Geometric Quantity Calculation Engine for Road Construction Projects.
Computes smooth, continuously scalable engineering quantities for ANY project length and lane configuration.
All geometric standards, layer thicknesses, and structural allowances are loaded from assumptions.yaml.
"""
from typing import Dict, Any, Optional
from .models import ProjectInput

class QuantityCalculator:
    """
    Computes engineering quantities derived purely from parameters in assumptions.yaml.
    """
    def __init__(self, project_input: ProjectInput, assumptions: Optional[Dict[str, Any]] = None):
        self.input = project_input
        self.assumptions = assumptions or {}
        self.length_km = max(0.1, float(project_input.length_km))
        self.length_m = self.length_km * 1000.0
        
        lanes = max(1, int(getattr(project_input, 'lanes', 4)))
        self.lanes = lanes
        self.lane_factor = lanes / 4.0
        
        # Load geometric parameters and lane standards from assumptions.yaml
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

        # Synchronize project_input geometry
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

        # Structural allowances
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
        """Computes quantities using pure data-driven formulas."""
        quantities = {}
        km_factor = (self.length_km / self.ref_corridor_km) * self.lane_factor

        # 1. Pre-Construction & Utilities
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

        # 2. Earthwork & Subgrade
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

        # 3. Granular Pavement
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

        # 4. Bituminous Pavement
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

        # 5. Cross Drainage Culverts
        box_count = max(0.2, self.length_km * self.box_culverts_per_km)
        pipe_count = max(0.2, self.length_km * self.pipe_culverts_per_km)
        culvert_width_factor = round(self.formation_width_m / 28.0, 2)
        quantities["box_culverts"] = {
            "quantity": round(box_count * self.box_culvert_concrete_cum * culvert_width_factor, 2),
            "unit": "cum",
            "formula": f"Length (km) × {self.box_culverts_per_km} Box Culverts/km × {self.box_culvert_concrete_cum} cum × Formation Factor ({culvert_width_factor:.2f})"
        }
        quantities["pipe_culverts"] = {
            "quantity": round(pipe_count * self.pipe_culvert_length_m * culvert_width_factor, 2),
            "unit": "metre",
            "formula": f"Length (km) × {self.pipe_culverts_per_km} Pipe Culverts/km × {self.pipe_culvert_length_m}m Twin Pipes × Formation Factor ({culvert_width_factor:.2f})"
        }

        # 6. Major & Minor Bridges + ROB (Proportional structural allowance per km from assumptions)
        quantities["bridge_piling"] = {
            "quantity": round(km_factor * self.piling_ref_m, 2),
            "unit": "metre",
            "formula": "Bridges & ROB Piling Allowance (Scaled from assumptions.yaml)"
        }
        quantities["bridge_substructure"] = {
            "quantity": round(km_factor * self.bridge_sub_ref_cum, 2),
            "unit": "cum",
            "formula": "Bridges & ROB Substructure Allowance (Scaled from assumptions.yaml)"
        }
        quantities["bridge_superstructure"] = {
            "quantity": round(km_factor * self.bridge_super_ref_cum, 2),
            "unit": "cum",
            "formula": "Bridges & ROB Superstructure Allowance (Scaled from assumptions.yaml)"
        }

        # 7. Grade Separators, Flyovers & RE Walls
        quantities["flyover_structures"] = {
            "quantity": round(km_factor * self.flyover_ref_cum, 2),
            "unit": "cum",
            "formula": "Elevated Flyovers & Interchanges Allowance (Scaled from assumptions.yaml)"
        }
        quantities["vup_structures"] = {
            "quantity": round(km_factor * self.vup_ref_cum, 2),
            "unit": "cum",
            "formula": "Vehicular/Pedestrian Underpasses Allowance (Scaled from assumptions.yaml)"
        }
        quantities["re_walls"] = {
            "quantity": round(km_factor * self.re_walls_ref_sqm, 2),
            "unit": "sqm",
            "formula": "RE Retaining Walls Allowance (Scaled from assumptions.yaml)"
        }

        # 8. Drainage & Slope Protection
        quantities["drainage"] = {
            "quantity": round(self.length_m * 2.0, 2),
            "unit": "metre",
            "formula": "Length (m) × 2 sides"
        }
        quantities["slope_protection"] = {
            "quantity": round(2.0 * self.length_m * (self.embankment_height_m * 2.24) * self.slope_coverage_ratio, 2),
            "unit": "sqm",
            "formula": "2 sides × Length × Slope Hypotenuse × Coverage Ratio"
        }

        # 9. Service Roads
        quantities["service_roads"] = {
            "quantity": round((self.service_road_ratio * self.length_m) * self.service_road_width_m * self.service_road_sides, 2),
            "unit": "sqm",
            "formula": f"Built-up Length × Width × {int(self.service_road_sides)} sides"
        }

        # 10. Road Safety, Toll Plaza & ATMS
        quantities["road_furniture"] = {
            "quantity": round(self.length_km * (0.5 + 0.5 * self.lane_factor), 2),
            "unit": "km",
            "formula": "Project Length (km) × Lane Scale Factor"
        }
        quantities["toll_plaza_atms"] = {
            "quantity": round(max(0.04, km_factor), 3),
            "unit": "set",
            "formula": "Toll Plaza & Corridor ATMS Network (Scaled to Highway Size)"
        }

        return quantities
