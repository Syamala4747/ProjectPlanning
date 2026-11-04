"""
Workbook Reader for the update-plan workflow.

Reads an existing 14-sheet Excel report produced by ExcelReportGenerator and
reconstructs the full project state needed for replanning:

  * ProjectInput  - parsed from Sheet 02 (Project Inputs)
  * The full ProjectScheduleResult - re-derived by re-running the
    deterministic ConstraintScheduler with the same ProjectInput, then
    cross-checked against the duration/cost figures printed in Sheet 01.

The planning pipeline is deterministic (same inputs + same assumptions.yaml
=> same output), so re-deriving reproduces the original baseline exactly.
When assumptions.yaml has drifted since the workbook was generated, the
re-derived figures are cross-checked against Sheet 01 and a warning is raised
so the user knows the baseline may have shifted.

This approach reuses every existing calculation module and requires NO
modification to the original ExcelReportGenerator: the 14 visible sheets are
the data source.
"""
import os
import re
from datetime import date, datetime
from typing import Optional, Tuple, Dict, Any, List

import openpyxl

from .models import ProjectInput, ProjectScheduleResult
from .workbook_metadata import read_metadata


class WorkbookReader:
    """Reads an existing Excel workbook and reconstructs the project state."""

    def __init__(self, input_path: str):
        self.input_path = input_path
        self.warnings: List[str] = []

    # ------------------------------------------------------------------ public
    def read(self) -> Tuple[ProjectInput, Dict[str, Any]]:
        """
        Returns (ProjectInput, sheet_metadata) where sheet_metadata is a dict
        of key figures extracted from the visible sheets for cross-checking.
        Raises FileNotFoundError if the workbook does not exist, ValueError if
        the workbook does not look like a planner output.
        """
        if not os.path.exists(self.input_path):
            raise FileNotFoundError(
                f"Input workbook not found: {self.input_path}")

        wb = openpyxl.load_workbook(self.input_path, data_only=True)
        self._validate_workbook(wb)

        proj_input = self._extract_project_input(wb)
        metadata = self._extract_summary(wb)
        metadata["workbook_provenance"] = read_metadata(wb)

        opt_date = self._extract_optimized_finish_date(wb)
        if opt_date:
            metadata["optimized_finish_date"] = opt_date

        return proj_input, metadata

    # ------------------------------------------------------------------ validate
    @staticmethod
    def _validate_workbook(wb) -> None:
        """Ensure the workbook has the expected sheet structure."""
        required = {"02_Project_Inputs", "05_Activities"}
        missing = required - set(wb.sheetnames)
        if missing:
            raise ValueError(
                f"Workbook '{wb}' does not appear to be a planner output: "
                f"missing sheet(s) {sorted(missing)}. Expected sheets like "
                f"'02_Project_Inputs' and '05_Activities'.")

    # ------------------------------------------------------------------ inputs
    def _extract_project_input(self, wb) -> ProjectInput:
        """
        Parse Sheet 02 (Project Inputs) to reconstruct the ProjectInput.
        Column A holds the parameter name, column B holds the value.
        Auto-detects project type from Sheet 01 banner title or from the
        parameters present in Sheet 02.
        """
        ws = wb["02_Project_Inputs"]

        # Build a label -> value lookup from column A/B pairs (rows 5+).
        values: Dict[str, Any] = {}
        for row in ws.iter_rows(min_row=5, max_col=5, values_only=True):
            label = row[0]
            val = row[1]
            if label is not None:
                values[str(label).strip()] = val

        # --- Detect project type from Sheet 01 banner or from parameter keys ---
        project_type = self._detect_project_type(wb, values)

        # --- Common fields (all project types) ---
        duration_days = self._parse_int(values.get("Required Completion Duration"))
        if duration_days is None or duration_days <= 0:
            raise ValueError(
                "Could not read 'Required Completion Duration' from Sheet 02.")

        budget_crore = self._parse_float(values.get("Maximum Available Budget"))
        if budget_crore is None or budget_crore <= 0:
            raise ValueError(
                "Could not read 'Maximum Available Budget' from Sheet 02.")

        start_date_raw = values.get("Project Start Date", "")
        start_date = self._parse_date(start_date_raw) or date.today()

        if project_type == "pipeline":
            return self._extract_pipeline_input(values, duration_days, budget_crore, start_date)
        elif project_type == "building":
            return self._extract_building_input(values, duration_days, budget_crore, start_date)
        else:
            return self._extract_road_input(values, duration_days, budget_crore, start_date)

    def _detect_project_type(self, wb, values: Dict[str, Any]) -> str:
        """Detect project type from the Sheet 01 dashboard banner title."""
        if "01_Project_Summary" in wb.sheetnames:
            ws = wb["01_Project_Summary"]
            banner = str(ws.cell(1, 1).value or "").upper()
            if "PIPELINE" in banner or "WATER" in banner:
                return "pipeline"
            if "BUILDING" in banner:
                return "building"
        # Fall back to checking which parameter labels are present in Sheet 02
        if "Pipeline Type" in values or "Pipe Diameter" in values or "Total Pipeline Length" in values:
            return "pipeline"
        if "Building Type" in values or "Built-up Area" in values or "Number of Floors" in values:
            return "building"
        return "road"

    def _extract_road_input(self, values: Dict[str, Any],
                            duration_days: int, budget_crore: float,
                            start_date: date) -> ProjectInput:
        """Reconstruct a road ProjectInput from Sheet 02 values."""
        length_km = self._parse_float(values.get("Project Length"))
        if length_km is None or length_km <= 0:
            raise ValueError(
                "Could not read 'Project Length' from Sheet 02.")

        lanes = self._parse_int(values.get("Number of Lanes"))
        if lanes is None or lanes <= 0:
            lanes = 4

        terrain_raw = str(values.get("Terrain", "")).strip().lower()
        terrain = "hilly" if terrain_raw in ("hilly", "hill", "mountain") else "plain"
        def nonnegative_int(label):
            return max(0, int(self._parse_int(values.get(label)) or 0))
        retaining_walls = max(
            0.0, float(self._parse_float(values.get("Retaining Walls")) or 0.0)
        )

        title = f"{length_km:g} KM {lanes}-Lane Highway Project"
        return ProjectInput(
            title=title,
            length_km=length_km,
            duration_days=duration_days,
            budget_crore=budget_crore,
            start_date=start_date,
            lanes=lanes,
            terrain=terrain,
            project_type="road",
            service_road_sides=nonnegative_int("Service Road Sides"),
            toll_plazas=nonnegative_int("Toll Plazas"),
            atms=nonnegative_int("ATMS"),
            bridges=nonnegative_int("Bridges"),
            flyovers=nonnegative_int("Flyovers"),
            underpasses=nonnegative_int("Underpasses"),
            retaining_walls_sqm=retaining_walls,
        )

    def _extract_pipeline_input(self, values: Dict[str, Any],
                                duration_days: int, budget_crore: float,
                                start_date: date) -> ProjectInput:
        """Reconstruct a pipeline ProjectInput from Sheet 02 values."""
        # Total Pipeline Length
        length_km = self._parse_float(
            values.get("Total Pipeline Length") or values.get("Project Length"))
        if length_km is None or length_km <= 0:
            raise ValueError("Could not read pipeline length from Sheet 02.")

        # Pipe Diameter
        dia_raw = values.get("Pipe Diameter (Nominal)", 0)
        dia_mm = float(self._parse_float(dia_raw) or 300.0)

        # Pipe Material
        mat_raw = str(values.get("Pipe Material", "DI") or "DI").strip().upper()
        if mat_raw not in {"DI", "HDPE", "MS", "RCC"}:
            mat_raw = "DI"

        # Trench depth / width
        trench_depth = float(
            self._parse_float(values.get("Trench Depth (Average)")) or 2.0)
        trench_width = float(
            self._parse_float(values.get("Trench Width (Excavated)")) or 0.0)

        # Installation method
        method_raw = str(
            values.get("Installation Method", "Open Trench") or "Open Trench"
        ).strip().lower()
        installation_method = (
            "trenchless" if ("trenchless" in method_raw or "hdd" in method_raw)
            else "open_trench"
        )

        # Work fronts
        fronts = int(self._parse_int(values.get("Number of Work Fronts")) or 1)

        # Terrain
        terrain_raw = str(values.get("Terrain Type", "plain") or "plain").strip().lower()
        terrain = "hilly" if "hill" in terrain_raw else "plain"

        # Pipeline type
        pipe_type = str(values.get("Pipeline Type", "water") or "water").strip().lower()
        if pipe_type not in {"water", "sewerage", "gas"}:
            pipe_type = "water"

        # Scheme-level inputs. These labels are present only for a facility
        # run; a corridor-only workbook leaves them absent (zero).
        num_villages = int(self._parse_int(values.get("Villages Served")) or 0)
        num_oht = int(self._parse_int(
            values.get("Overhead Service Reservoirs (OHT/ESR)")) or 0)
        wtp_mld = float(
            self._parse_float(values.get("Water Treatment Plant Capacity")) or 0.0)
        num_ibps = int(self._parse_int(
            values.get("Intake / Booster Pumping Stations")) or 0)
        num_fhtc = int(self._parse_int(
            values.get("Functional Household Tap Connections")) or 0)

        # Declared site-specific crossings. Old workbooks do not carry these
        # rows, so they fall back to zero and the quantity calculator derives
        # only the road crossings implied by the alignment.
        number_of_road_crossings = int(self._parse_int(
            values.get("Road Crossings (Declared)")) or 0)
        number_of_highway_crossings = int(self._parse_int(
            values.get("Highway Crossings (Declared)")) or 0)
        number_of_rail_crossings = int(self._parse_int(
            values.get("Rail Crossings (Declared)")) or 0)
        number_of_hdd_crossings = int(self._parse_int(
            values.get("HDD / Trenchless Crossings (Declared)")) or 0)
        number_of_river_crossings = int(self._parse_int(
            values.get("River Crossings (Declared)")) or 0)

        title = f"{length_km:g} km DN{int(dia_mm)} {mat_raw} Pipeline"
        return ProjectInput(
            title=title,
            length_km=length_km,
            duration_days=duration_days,
            budget_crore=budget_crore,
            start_date=start_date,
            project_type="pipeline",
            pipe_diameter_mm=dia_mm,
            pipe_material=mat_raw,
            trench_depth_m=trench_depth,
            trench_width_m=trench_width,
            installation_method=installation_method,
            number_of_work_fronts=fronts,
            pipeline_type=pipe_type,
            terrain=terrain,
            num_villages=num_villages,
            num_oht=num_oht,
            wtp_capacity_mld=wtp_mld,
            num_ibps=num_ibps,
            num_fhtc=num_fhtc,
            number_of_road_crossings=number_of_road_crossings,
            number_of_highway_crossings=number_of_highway_crossings,
            number_of_rail_crossings=number_of_rail_crossings,
            number_of_hdd_crossings=number_of_hdd_crossings,
            number_of_river_crossings=number_of_river_crossings,
        )

    def _extract_building_input(self, values: Dict[str, Any],
                                duration_days: int, budget_crore: float,
                                start_date: date) -> ProjectInput:
        """Reconstruct a building ProjectInput from Sheet 02 values."""
        # Built-up area
        bua = self._parse_float(values.get("Built-up Area (Total)") or
                                values.get("Built-up Area"))
        if bua is None or bua <= 0:
            raise ValueError("Could not read built-up area from Sheet 02.")

        # Floors
        floors_raw = values.get("Number of Floors", 4)
        floors = int(self._parse_int(floors_raw) or 4)

        # Building type
        btype = str(values.get("Building Type", "residential") or "residential").strip().lower()
        if btype not in {"residential", "commercial"}:
            btype = "residential"

        # Construction type
        ctype_raw = str(values.get("Construction Type", "RCC") or "RCC").strip().lower()
        ctype = "rcc" if "rcc" in ctype_raw else ("steel" if "steel" in ctype_raw else "load_bearing")

        # Quality finish
        quality_raw = str(values.get("Quality / Finish Level", "standard") or "standard").strip().lower()
        quality = quality_raw if quality_raw in {"basic", "standard", "premium"} else "standard"

        # Basement
        basement_raw = str(values.get("Basement", "No") or "No").strip().lower()
        basement = basement_raw in {"yes", "true", "1"}

        # Blocks / towers determine the external envelope and the available
        # parallel fronts, so they must survive the workbook round-trip.
        blocks = int(self._parse_int(
            values.get("Number of Blocks / Towers")) or 1)
        foundation = str(values.get("Foundation Type", "pile") or "pile").strip().lower()
        if foundation not in {"pile", "raft", "isolated_footing"}:
            foundation = "pile"
        floor_detail = str(
            values.get("Floor-wise Structural Detail", "No") or "No"
        ).strip().lower() in {"yes", "true", "1"}

        title = f"{bua:,.0f} sqm G+{floors - 1} {btype.title()} Building"
        return ProjectInput(
            title=title,
            length_km=bua,  # carries bua for routing
            duration_days=duration_days,
            budget_crore=budget_crore,
            start_date=start_date,
            project_type="building",
            built_up_area_sqm=bua,
            number_of_floors=floors,
            building_type=btype,
            construction_type=ctype,
            quality_finish=quality,
            basement=basement,
            number_of_blocks=blocks,
            foundation_type=foundation,
            floor_wise_structural_detail=floor_detail,
        )

    # ------------------------------------------------------------------ summary
    def _extract_summary(self, wb) -> Dict[str, Any]:
        """
        Extract key figures from Sheet 01 (Project Summary) for cross-checking
        the re-derived baseline against the original workbook.
        """
        meta: Dict[str, Any] = {}
        if "01_Project_Summary" not in wb.sheetnames:
            return meta
        ws = wb["01_Project_Summary"]

        baseline_days = self._parse_int(
            self._find_dashboard_value(ws, "Baseline Duration"))
        if baseline_days is not None and baseline_days > 0:
            meta["baseline_duration_days"] = baseline_days

        baseline_cost = self._parse_float(
            self._find_dashboard_value(ws, "Baseline Cost"))
        if baseline_cost is not None and baseline_cost > 0:
            meta["baseline_cost_crore"] = baseline_cost

        # Also grab the S-curve monthly data for baseline progress lookup.
        scurve = self._extract_scurve(wb)
        if scurve:
            meta["scurve_months"] = scurve

        return meta

    @staticmethod
    def _find_dashboard_value(ws, label: str):
        """Find a KPI label and return the value in the cell below it.

        The Sheet 01 dashboard writes labels in one row and values directly
        beneath them (for example E5='Baseline Duration (Normal)', E6='537
        Days'). It falls back to the cell to the right for layouts that use
        label/value pairs on one row.
        """
        wanted = str(label).strip().lower()
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str):
                    continue
                if not value.strip().lower().startswith(wanted):
                    continue
                below = ws.cell(row=cell.row + 1, column=cell.column).value
                if below is not None and str(below).strip():
                    return below
                right = ws.cell(row=cell.row, column=cell.column + 1).value
                if right is not None and str(right).strip():
                    return right
        return None

    def _extract_scurve(self, wb) -> List[Dict[str, Any]]:
        """Parse Sheet 13 (S-Curve) for monthly cumulative progress data."""
        if "13_S_Curve" not in wb.sheetnames:
            return []
        ws = wb["13_S_Curve"]
        months: List[Dict[str, Any]] = []
        for row in ws.iter_rows(min_row=5, values_only=True):
            month_name = row[0]
            cum_pct = row[3]
            if month_name is None:
                continue
            cum_val = None
            if isinstance(cum_pct, (int, float)):
                cum_val = float(cum_pct) * 100.0  # stored as fraction
            months.append({
                "month_name": str(month_name),
                "cumulative_progress_pct": cum_val,
            })
        return months

    # ------------------------------------------------------------------ helpers
    @staticmethod
    def _parse_float(val) -> Optional[float]:
        if val is None:
            return None
        if isinstance(val, (int, float)):
            return float(val)
        # Dashboard cells are formatted like "₹ 43.33 Crore" or "537 Days".
        # Keep only the numeric token so the value is not lost to the unit.
        s = re.sub(r"[^\d.\-]", "", str(val).strip())
        try:
            return float(s)
        except ValueError:
            return None

    @staticmethod
    def _parse_int(val) -> Optional[int]:
        if val is None:
            return None
        if isinstance(val, (int, float)):
            return int(val)
        s = re.sub(r"[^\d.]", "", str(val).strip())
        try:
            return int(round(float(s)))
        except (ValueError, TypeError):
            return None

    def _extract_optimized_finish_date(self, wb) -> Optional[date]:
        """
        Extract the optimized final completion date from the workbook.
        Checks:
          1. '01_Project_Summary' for 'Optimized Finish Date' / 'Optimized Completion Date'
          2. '09_Optimized_CPM' for the maximum Planned Finish Date
          3. '01_Update_Summary' for 'Revised Completion Date' / 'Target Date'
          4. '05_Revised_CPM' for the maximum Planned Finish Date
          5. '06_Baseline_CPM' for the maximum Planned Finish Date
        """
        # 1. Sheet 01: Project Summary
        if "01_Project_Summary" in wb.sheetnames:
            ws = wb["01_Project_Summary"]
            for row in ws.iter_rows(values_only=True):
                if not any(row):
                    continue
                for col_idx, cell_val in enumerate(row):
                    if cell_val is None:
                        continue
                    s = str(cell_val).strip().lower()
                    if "optimized finish date" in s or "optimized completion date" in s:
                        for next_val in row[col_idx + 1:]:
                            if next_val is not None and str(next_val).strip():
                                parsed = self._parse_date(next_val)
                                if parsed:
                                    return parsed

        # 2. Sheet 09: Optimized CPM
        if "09_Optimized_CPM" in wb.sheetnames:
            ws = wb["09_Optimized_CPM"]
            d = self._extract_max_planned_finish(ws)
            if d:
                return d

        # 3. Sheet 01: Update Summary (if previous was an update workbook)
        if "01_Update_Summary" in wb.sheetnames:
            ws = wb["01_Update_Summary"]
            for row in ws.iter_rows(values_only=True):
                if not any(row):
                    continue
                for col_idx, cell_val in enumerate(row):
                    if cell_val is None:
                        continue
                    s = str(cell_val).strip().lower()
                    if "revised completion date" in s or "target date" in s:
                        for next_val in row[col_idx + 1:]:
                            if next_val is not None and str(next_val).strip():
                                parsed = self._parse_date(next_val)
                                if parsed:
                                    return parsed

        # 4. Sheet 05: Revised CPM
        if "05_Revised_CPM" in wb.sheetnames:
            ws = wb["05_Revised_CPM"]
            d = self._extract_max_planned_finish(ws)
            if d:
                return d

        # 5. Sheet 06: Baseline CPM
        if "06_Baseline_CPM" in wb.sheetnames:
            ws = wb["06_Baseline_CPM"]
            d = self._extract_max_planned_finish(ws)
            if d:
                return d

        return None

    def _extract_max_planned_finish(self, ws) -> Optional[date]:
        """Find the maximum Planned Finish Date from a CPM schedule sheet."""
        finish_col_idx = None
        dates: List[date] = []
        for row in ws.iter_rows(values_only=True):
            if not any(row):
                continue
            if finish_col_idx is None:
                for idx, val in enumerate(row):
                    if val and "planned finish" in str(val).lower():
                        finish_col_idx = idx
                        break
            else:
                if finish_col_idx < len(row):
                    val = row[finish_col_idx]
                    if val is not None and str(val).strip():
                        parsed = self._parse_date(val)
                        if parsed:
                            dates.append(parsed)
        return max(dates) if dates else None

    @classmethod
    def _parse_date(cls, s: Any) -> Optional[date]:
        if s is None:
            return None
        if isinstance(s, datetime):
            return s.date()
        if isinstance(s, date):
            return s
        s_str = str(s).strip()
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(s_str, fmt).date()
            except ValueError:
                continue
        # If it's a serial date number from Excel
        try:
            val_num = float(s_str)
            if 30000 <= val_num <= 80000:
                return date.fromordinal(int(val_num) + 693594)
        except (ValueError, TypeError):
            pass
        return None
