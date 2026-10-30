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
        """
        ws = wb["02_Project_Inputs"]

        # Build a label -> value lookup from column A/B pairs (rows 5+).
        values: Dict[str, Any] = {}
        for row in ws.iter_rows(min_row=5, max_col=5, values_only=True):
            label = row[0]
            val = row[1]
            if label is not None:
                values[str(label).strip()] = val

        # --- Length ---
        length_km = self._parse_float(values.get("Project Length"))
        if length_km is None or length_km <= 0:
            raise ValueError(
                "Could not read 'Project Length' from Sheet 02. "
                "Ensure the workbook is a valid planner output.")

        # --- Lanes ---
        lanes = self._parse_int(values.get("Number of Lanes"))
        if lanes is None or lanes <= 0:
            raise ValueError(
                "Could not read 'Number of Lanes' from Sheet 02.")

        # --- Terrain ---
        terrain_raw = str(values.get("Terrain", "")).strip().lower()
        terrain = "hilly" if terrain_raw in ("hilly", "hill", "mountain") else "plain"

        # --- Duration ---
        duration_days = self._parse_int(values.get("Required Completion Duration"))
        if duration_days is None or duration_days <= 0:
            raise ValueError(
                "Could not read 'Required Completion Duration' from Sheet 02.")

        # --- Budget ---
        budget_crore = self._parse_float(values.get("Maximum Available Budget"))
        if budget_crore is None or budget_crore <= 0:
            raise ValueError(
                "Could not read 'Maximum Available Budget' from Sheet 02.")

        # --- Start date ---
        start_date_raw = values.get("Project Start Date", "")
        start_date = self._parse_date(start_date_raw) or date.today()

        # Reconstruct a title (not stored directly; auto-detect lanes/length).
        title = f"{length_km:g} KM {lanes}-Lane Highway Project"

        return ProjectInput(
            title=title,
            length_km=length_km,
            duration_days=duration_days,
            budget_crore=budget_crore,
            start_date=start_date,
            lanes=lanes,
            terrain=terrain,
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
        for row in ws.iter_rows(values_only=True):
            for val in row:
                if val and isinstance(val, str):
                    meta["_raw_" + str(val)[:40]] = True

        # Also grab the S-curve monthly data for baseline progress lookup.
        scurve = self._extract_scurve(wb)
        if scurve:
            meta["scurve_months"] = scurve

        return meta

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
        s = str(val).strip().replace(",", "").replace("₹", "").replace("Cr", "")
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
