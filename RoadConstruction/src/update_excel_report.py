"""
Excel Report Generator for the update-plan workflow.

Produces an 8-sheet executive workbook that clearly distinguishes:
  * Original baseline values
  * Actual values (as of the status date)
  * Revised forecast values (remaining work)

Sheets:
  01_Update_Summary        — executive dashboard of the update
  02_Baseline_Schedule     — original optimized CPM schedule
  03_Actual_Progress       — actual vs baseline progress at status date
  04_Activity_Status      — completed / remaining quantities per activity
  05_Revised_CPM          — revised schedule for remaining work
  06_Revised_Cost         — remaining-work cost and acceleration surcharge
  07_S_Curve_Comparison   — baseline / actual / revised forecast curves
  08_Target_Feasibility   — feasibility verdict and resource changes
"""
import os
from types import SimpleNamespace
from datetime import date, timedelta
from typing import List, Dict, Any, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import AxDataSource, StrRef

from .models import (
    UpdatePlanResult, RoadActivity, ActivityActualStatus,
)
from .plan_generator import build_plan
from .engineer_daily_log import build_daily_log
from .excel_report import ExcelReportGenerator
from .workbook_metadata import attach_metadata, build_metadata


class UpdateExcelReportGenerator:
    """Generates the updated-plan Excel workbook."""

    def __init__(self, result: UpdatePlanResult, output_path: str):
        self.r = result
        self.p_in = result.project_input
        self.output_path = output_path

        # Color palette
        self.c_navy = "1B365D"
        self.c_slate = "4A777A"
        self.c_accent = "2E75B6"
        self.c_light_bg = "F0F4F8"
        self.c_crit_bg = "FFEBEE"
        self.c_crit_fg = "C62828"
        self.c_pass_bg = "E8F5E9"
        self.c_pass_fg = "2E7D32"
        self.c_border = "D0D7DE"
        self.c_card_bg = "F8FAFC"
        self.c_gold = "D97706"
        self.c_actual = "7B1FA2"    # purple for actual
        self.c_revised = "00695C"   # teal for revised

        # Styles
        self.font_title = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
        self.font_subtitle = Font(name="Calibri", size=11, italic=True, color="E0E0E0")
        self.font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        self.font_bold = Font(name="Calibri", size=11, bold=True, color="000000")
        self.font_regular = Font(name="Calibri", size=10, color="000000")
        self.font_small = Font(name="Calibri", size=9, color="555555")
        self.font_critical = Font(name="Calibri", size=10, bold=True, color=self.c_crit_fg)
        self.font_pass = Font(name="Calibri", size=10, bold=True, color=self.c_pass_fg)

        self.fill_primary = PatternFill("solid", start_color=self.c_navy, end_color=self.c_navy)
        self.fill_accent = PatternFill("solid", start_color=self.c_accent, end_color=self.c_accent)
        self.fill_zebra = PatternFill("solid", start_color=self.c_light_bg, end_color=self.c_light_bg)
        self.fill_crit = PatternFill("solid", start_color=self.c_crit_bg, end_color=self.c_crit_bg)
        self.fill_pass = PatternFill("solid", start_color=self.c_pass_bg, end_color=self.c_pass_bg)
        self.fill_card = PatternFill("solid", start_color=self.c_card_bg, end_color=self.c_card_bg)
        self.fill_actual = PatternFill("solid", start_color="F3E5F5", end_color="F3E5F5")
        self.fill_revised = PatternFill("solid", start_color="E0F2F1", end_color="E0F2F1")

        self.border_thin = Border(
            left=Side(style="thin", color=self.c_border),
            right=Side(style="thin", color=self.c_border),
            top=Side(style="thin", color=self.c_border),
            bottom=Side(style="thin", color=self.c_border),
        )
        self.border_double = Border(
            bottom=Side(style="double", color="000000"),
            top=Side(style="thin", color="000000"),
        )

    # ------------------------------------------------------------------ generate
    def generate(self) -> str:
        out_dir = os.path.dirname(self.output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        self._build_summary(wb)
        self._build_baseline_schedule(wb)
        self._build_assumptions(wb)
        self._build_actual_progress(wb)
        self._build_activity_status(wb)
        self._build_revised_cpm(wb)
        self._build_revised_cost(wb)
        self._build_scurve(wb)
        self._build_feasibility(wb)
        self._build_revised_resource_sheets(wb)
        self._build_validation(wb)
        baseline = getattr(self.r, "baseline_result", None)
        guidelines = getattr(baseline, "guidelines_path", "")
        assumptions = getattr(baseline, "assumptions_path", "")
        if guidelines and assumptions:
            attach_metadata(wb, build_metadata(self.p_in, guidelines, assumptions))

        try:
            wb.save(self.output_path)
            return self.output_path
        except PermissionError:
            base, ext = os.path.splitext(self.output_path)
            fallback = f"{base}_new{ext}"
            wb.save(fallback)
            print(f"\n  Note: '{self.output_path}' was open in Excel. "
                  f"Saved as '{fallback}' instead.")
            return fallback

    def _build_revised_resource_sheets(self, wb):
        """Export resource data for remaining work, never the old baseline."""
        view = SimpleNamespace(
            project_input=self.r.project_input,
            activities=self.r.revised_activities,
            baseline_activities=[],
            assumptions_snapshot=getattr(
                self.r.baseline_result, "assumptions_snapshot", {}
            ),
        )
        resource_reporter = ExcelReportGenerator(view, "")
        resource_reporter._build_sheet_plan(wb)
        resource_reporter._build_sheet_daily_log(wb)

    def _build_assumptions(self, wb):
        """Keep the update workbook's assumptions in the third sheet."""
        view = SimpleNamespace(
            project_input=self.r.project_input,
            activities=self.r.revised_activities,
            baseline_activities=[],
            assumptions_snapshot=getattr(
                self.r.baseline_result, "assumptions_snapshot", {}
            ),
        )
        ExcelReportGenerator(view, "")._build_sheet_03_assumptions(wb)

    def _format_resource_sheet(self, ws, headers):
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{max(2, ws.max_row)}"
        for row in ws.iter_rows(min_row=3):
            for cell in row:
                cell.border = self.border_thin
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if isinstance(cell.value, date):
                    cell.number_format = "yyyy-mm-dd"
                elif isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
        for index, header in enumerate(headers, start=1):
            ws.column_dimensions[get_column_letter(index)].width = min(
                32, max(12, len(header) + 2)
            )

    def _build_validation(self, wb):
        """Write the same clear PASS/WARNING/ERROR register as new plans."""
        ws = wb.create_sheet(title="09_Validation_Checks")
        report = getattr(self.r, "validation_report", None)
        status = getattr(report, "status", "NOT RUN")
        issues = list(getattr(report, "issues", []) or [])
        errors = list(getattr(report, "errors", []) or [])
        warnings = list(getattr(report, "warnings", []) or [])
        catalog = [
            ("UPDATE_ACTIVITY_STATUS_SET_MISMATCH", "Actual status rows match baseline activities"),
            ("UPDATE_PROGRESS_OUT_OF_RANGE", "Actual progress is between 0% and 100%"),
            ("UPDATE_QUANTITY_RECONCILIATION", "Completed plus remaining quantity equals original"),
            ("UPDATE_CPM_EF_IDENTITY", "Revised EF equals ES plus duration"),
            ("UPDATE_CPM_LF_IDENTITY", "Revised LF equals LS plus duration"),
            ("UPDATE_CPM_FLOAT_IDENTITY", "Revised float equals LS minus ES"),
            ("UPDATE_MISSING_PREDECESSOR", "All revised predecessors exist"),
            ("UPDATE_PRECEDENCE_VIOLATION", "Revised predecessor dates are respected"),
            ("UPDATE_DIRECT_COST_RECONCILIATION", "Revised direct cost reconciles"),
            ("UPDATE_TOTAL_COST_RECONCILIATION", "Revised total cost reconciles"),
            ("UPDATE_WASTAGE_RECONCILIATION", "Revised wastage components reconcile"),
            ("UPDATE_ESCALATION_OUTPUT_INVALID", "Revised escalation output is valid"),
            ("UPDATE_DURATION_RECONCILIATION", "Revised duration matches activity finish"),
            ("UPDATE_REMAINING_COST_RECONCILIATION", "Remaining cost matches revised activities"),
            ("UPDATE_PROJECT_COST_RECONCILIATION", "Spent plus remaining equals project cost"),
            ("UPDATE_SCURVE_COST_RECONCILIATION", "Revised S-curve cost reconciles"),
            ("UPDATE_COMPLETION_DATE_MISSING", "Revised completion date exists"),
            ("UPDATE_NEGATIVE_DELAY", "Projected delay is non-negative"),
            ("UPDATE_PLANNING_WARNING", "Update planning warnings are surfaced"),
        ]
        self._create_banner(
            ws,
            "UPDATE VALIDATION & SANITY CHECK RESULTS",
            "Complete update register: PASS checks, review warnings, and hard errors",
            7,
        )
        for row, label, value in (
            (4, "Overall status", status),
            (5, "Error count", len(errors)),
            (6, "Warning count", len(warnings)),
        ):
            ws.cell(row=row, column=1, value=label).font = self.font_bold
            ws.cell(row=row, column=2, value=value).border = self.border_thin
            ws.cell(row=row, column=1).border = self.border_thin
        ws.cell(row=7, column=1, value="Passed check count").font = self.font_bold
        headers = ["Severity", "Check code", "Activity ID", "Result", "Check / expected rule", "Finding", "Action"]
        self._apply_header(ws, 10, headers)
        issue_codes = {str(issue.code) for issue in issues}
        rows = [(
            issue.severity, issue.code, issue.activity_id or "-",
            "ERROR" if issue.severity == "ERROR" else "WARNING",
            next((text for code, text in catalog if code == issue.code), "Update validation finding"),
            issue.message,
            "Correct update calculation before issue" if issue.severity == "ERROR"
            else "Review/confirm update assumption",
        ) for issue in issues]
        rows.extend((
            "PASS", code, "-", "PASS", text,
            "Check completed successfully; no exception was reported.",
            "No action required",
        ) for code, text in catalog if code not in issue_codes)
        ws.cell(row=7, column=2, value=sum(1 for row in rows if row[0] == "PASS"))
        for row_index, values in enumerate(rows, start=11):
            severity = str(values[0])
            for col_index, value in enumerate(values, start=1):
                cell = ws.cell(row=row_index, column=col_index, value=value)
                cell.border = self.border_thin
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.fill = self.fill_crit if severity == "ERROR" else (
                    self.fill_pass if severity == "PASS" else self.fill_card
                )
        ws.freeze_panes = "A11"
        ws.auto_filter.ref = f"A10:G{max(11, 10 + len(rows))}"
        for col, width in {"A": 13, "B": 38, "C": 14, "D": 12, "E": 48, "F": 75, "G": 34}.items():
            ws.column_dimensions[col].width = width

    # ------------------------------------------------------------------ helpers
    def _create_banner(self, ws, title: str, subtitle: str, max_cols: int = 10):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
        ws.row_dimensions[1].height = 32
        c1 = ws.cell(row=1, column=1, value=title)
        c1.font = self.font_title
        c1.fill = self.fill_primary
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_cols)
        ws.row_dimensions[2].height = 20
        c2 = ws.cell(row=2, column=1, value=subtitle)
        c2.font = self.font_subtitle
        c2.fill = self.fill_accent
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    def _apply_header(self, ws, row: int, titles: List[str]):
        ws.row_dimensions[row].height = 26
        for col, title in enumerate(titles, 1):
            cell = ws.cell(row=row, column=col, value=title)
            cell.font = self.font_header
            cell.fill = self.fill_primary
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = self.border_thin

    def _auto_fit(self, ws, max_cap: int = 50):
        for col in ws.columns:
            max_len = 0
            letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    s = str(cell.value)
                    if len(s) < 80:
                        max_len = max(max_len, len(s))
            ws.column_dimensions[letter].width = max(max_len + 4, 13)

    @staticmethod
    def _status_color(status: str) -> Optional[PatternFill]:
        if status == "completed":
            return PatternFill("solid", start_color="E8F5E9", end_color="E8F5E9")
        if status == "in_progress":
            return PatternFill("solid", start_color="FFF3E0", end_color="FFF3E0")
        return None

    # ------------------------------------------------------------------ sheets
    # SHEET 1: UPDATE SUMMARY
    # ------------------------------------------------------------------
    def _build_summary(self, wb):
        ws = wb.create_sheet(title="01_Update_Summary")
        self._create_banner(
            ws,
            "PROJECT PLAN UPDATE — EXECUTIVE SUMMARY",
            f"Original Plan Start: {self.p_in.start_date} | Status Inspection Date: {self.r.status_date}",
            6,
        )
        r = self.r
        p = self.p_in

        # KPI cards
        ptype = str(getattr(p, "project_type", "road") or "road").lower()
        if ptype == "pipeline":
            mat = getattr(p, "pipe_material", "DI") or "DI"
            dia = int(getattr(p, "pipe_diameter_mm", 0) or 0)
            scope_val = f"{p.length_km:.1f} km, DN{dia} {mat} Pipeline"
        elif ptype == "building":
            bua = getattr(p, "built_up_area_sqm", 0) or 0
            floors = getattr(p, "number_of_floors", 1) or 1
            scope_val = f"{bua:,.0f} sqm, G+{floors-1} {getattr(p, 'building_type', 'Residential').title()} Building"
        else:
            scope_val = f"{p.length_km:.1f} km, {p.lanes}-Lane Highway"

        kpis = [
            ("Project Scope", scope_val),
            ("Original Start Date", str(p.start_date)),
            ("Current Status Date", str(r.status_date)),
            ("Actual Progress", f"{r.actual_progress_pct:.1f}%"),
            ("Baseline Progress at Status Date", f"{r.baseline_progress_at_status_date:.1f}%"),
            ("Progress Gap (Behind)", f"{r.progress_gap_pct:.1f} pp"),
            ("Target Date", str(r.target_date)),
            ("Target Progress", f"{r.target_progress_pct:.1f}%"),
            ("Required Progress (Remaining)", f"{r.required_progress_pct:.1f} pp"),
            ("Revised Duration (from status)", f"{r.revised_duration_days} Days"),
            ("Revised Completion Date", str(r.revised_completion_date or "")),
            ("Target Feasible", "YES" if r.target_feasible else "NO"),
            ("Projected Delay", f"{r.projected_delay_days} Days"),
            ("Original Baseline Cost", f"Rs {r.baseline_total_cost_crore:,.2f} Cr"),
            ("Spent Cost (Completed Work)", f"Rs {r.spent_cost_crore:,.2f} Cr"),
            ("Remaining Base Cost (Normal)", f"Rs {r.remaining_base_cost_crore:,.2f} Cr"),
            ("Additional Acceleration Cost", f"Rs {r.additional_acceleration_cost_crore:,.2f} Cr"),
            ("Total Remaining Cost", f"Rs {r.remaining_cost_crore:,.2f} Cr"),
            ("Total Revised Project Cost", f"Rs {r.total_project_cost_crore:,.2f} Cr"),
            ("Work Fronts (Revised)", f"{r.work_fronts} Fronts"),
        ]

        ws.cell(row=4, column=1,
                value="EXECUTIVE KEY PERFORMANCE INDICATORS").font = \
            Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells("A4:F4")

        for idx, (title, val) in enumerate(kpis):
            row = 5 + (idx // 3) * 3
            col = 1 + (idx % 3) * 2

            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
            ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)

            ct = ws.cell(row=row, column=col, value=title)
            ct.font = Font(size=9, bold=True, color="555555")
            ct.fill = self.fill_card
            ct.alignment = Alignment(horizontal="center", vertical="center")
            ct.border = self.border_thin

            cv = ws.cell(row=row + 1, column=col, value=val)
            cv.font = Font(size=13, bold=True, color=self.c_navy)
            cv.fill = self.fill_card
            cv.alignment = Alignment(horizontal="center", vertical="center")
            cv.border = self.border_thin

            if title == "Target Feasible":
                cv.fill = self.fill_pass if r.target_feasible else self.fill_crit
                cv.font = Font(size=13, bold=True,
                               color=self.c_pass_fg if r.target_feasible else self.c_crit_fg)
            elif title == "Projected Delay" and r.projected_delay_days > 0:
                cv.fill = self.fill_crit
                cv.font = Font(size=13, bold=True, color=self.c_crit_fg)

        curr = 5 + ((len(kpis) + 2) // 3) * 3 + 1
        ws.cell(row=curr, column=1, value="FEASIBILITY VERDICT").font = \
            Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        curr += 1

        vcell = ws.cell(row=curr, column=1, value=r.feasibility_message)
        vcell.font = self.font_bold
        vcell.alignment = Alignment(wrap_text=True)
        vcell.fill = self.fill_pass if r.target_feasible else self.fill_crit
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        ws.row_dimensions[curr].height = 36
        curr += 2

        if r.limitations:
            ws.cell(row=curr, column=1, value="LIMITATIONS & ASSUMPTIONS").font = \
                Font(size=11, bold=True, color=self.c_gold)
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
            curr += 1
            for lim in r.limitations:
                c = ws.cell(row=curr, column=1, value=f"  - {lim}")
                c.font = self.font_small
                c.alignment = Alignment(wrap_text=True)
                ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
                ws.row_dimensions[curr].height = 30
                curr += 1

        if r.warnings:
            curr += 1
            ws.cell(row=curr, column=1, value="WARNINGS").font = \
                Font(size=11, bold=True, color=self.c_crit_fg)
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
            curr += 1
            for w in r.warnings:
                c = ws.cell(row=curr, column=1, value=f"  - {w}")
                c.font = self.font_small
                c.alignment = Alignment(wrap_text=True)
                ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
                ws.row_dimensions[curr].height = 20
                curr += 1

        self._auto_fit(ws)

    # SHEET 2: BASELINE SCHEDULE
    # ------------------------------------------------------------------
    def _build_baseline_schedule(self, wb):
        ws = wb.create_sheet(title="02_Baseline_Schedule")
        self._create_banner(
            ws,
            "ORIGINAL BASELINE SCHEDULE (OPTIMIZED PLAN)",
            "The original optimized CPM schedule from the existing workbook",
            13,
        )
        headers = [
            "Activity ID", "Activity Name", "Predecessors", "Duration (d)",
            "ES", "EF", "LS", "LF", "Total Float", "Free Float",
            "Planned Start", "Planned Finish", "Critical?",
        ]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "A5"

        for i, act in enumerate(self.r.baseline_result.activities, 5):
            preds = ", ".join(act.predecessors) if act.predecessors else "-"
            vals = [
                act.activity_id, act.name, preds, act.duration_days,
                act.es, act.ef, act.ls, act.lf,
                act.total_float, act.free_float,
                str(act.planned_start or ""), str(act.planned_finish or ""),
                "CRITICAL" if act.is_critical else "",
            ]
            for col_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (4, 5, 6, 7, 8, 9, 10, 13):
                    cell.alignment = Alignment(horizontal="center")
                if act.is_critical:
                    cell.fill = self.fill_crit
                    if col_idx == 13:
                        cell.font = self.font_critical
                elif i % 2 == 0:
                    cell.fill = self.fill_zebra
            ws.row_dimensions[i].height = 20

        self._auto_fit(ws)

    # SHEET 3: ACTUAL PROGRESS
    # ------------------------------------------------------------------
    def _build_actual_progress(self, wb):
        ws = wb.create_sheet(title="03_Actual_Progress")
        self._create_banner(
            ws,
            "ACTUAL PROGRESS vs BASELINE (AS OF STATUS DATE)",
            f"Per-activity comparison of baseline planned vs actual completion as of Status Date: {self.r.status_date}",
            8,
        )
        headers = [
            "Activity ID", "Activity Name", "Baseline Completion %",
            "Actual Completion %", "Gap (pp)", "Status",
            "Delayed?", "Notes",
        ]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "A5"

        for i, st in enumerate(self.r.activity_status, 5):
            gap = round(st.baseline_completion_pct - st.actual_completion_pct, 1)
            vals = [
                st.activity_id, st.name,
                st.baseline_completion_pct / 100.0,
                st.actual_completion_pct / 100.0,
                gap / 100.0,
                st.status,
                "YES" if st.is_delayed else "",
                "Behind baseline" if st.is_delayed else
                ("On track" if abs(gap) < 1 else
                 "Ahead of baseline" if gap < 0 else ""),
            ]
            for col_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (3, 4, 5):
                    cell.number_format = "0.0%"
                elif col_idx in (1, 6, 7):
                    cell.alignment = Alignment(horizontal="center")

                fill = self._status_color(st.status)
                if fill:
                    cell.fill = fill
                elif i % 2 == 0:
                    cell.fill = self.fill_zebra

                if col_idx == 7 and st.is_delayed:
                    cell.font = self.font_critical
            ws.row_dimensions[i].height = 20

        # Totals row
        tot_r = 5 + len(self.r.activity_status)
        ws.cell(row=tot_r, column=2, value="PROJECT TOTAL").font = self.font_bold
        ws.cell(row=tot_r, column=3,
                value=self.r.baseline_progress_at_status_date / 100.0).number_format = "0.0%"
        ws.cell(row=tot_r, column=4,
                value=self.r.actual_progress_pct / 100.0).number_format = "0.0%"
        ws.cell(row=tot_r, column=5,
                value=self.r.progress_gap_pct / 100.0).number_format = "0.0%"
        for c in range(1, 9):
            cell = ws.cell(row=tot_r, column=c)
            cell.font = self.font_bold
            cell.border = self.border_double
            cell.fill = self.fill_zebra

        self._auto_fit(ws)

    # SHEET 4: ACTIVITY STATUS (COMPLETED / REMAINING)
    # ------------------------------------------------------------------
    def _build_activity_status(self, wb):
        ws = wb.create_sheet(title="04_Activity_Status")
        self._create_banner(
            ws,
            "ACTIVITY-WISE COMPLETED & REMAINING QUANTITIES",
            f"Original, completed, and remaining quantities for each activity as of Status Date: {self.r.status_date}",
            9,
        )
        headers = [
            "Activity ID", "Activity Name", "Unit", "Original Quantity",
            "Completed Quantity", "Remaining Quantity",
            "Completion %", "Status", "Delayed?",
        ]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "A5"

        orig_acts = {a.activity_id: a for a in self.r.baseline_result.activities}

        for i, st in enumerate(self.r.activity_status, 5):
            orig = orig_acts.get(st.activity_id)
            unit = orig.unit if orig else ""
            vals = [
                st.activity_id, st.name, unit,
                st.original_quantity, st.completed_quantity,
                st.remaining_quantity,
                st.actual_completion_pct / 100.0,
                st.status,
                "YES" if st.is_delayed else "",
            ]
            for col_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (4, 5, 6):
                    cell.number_format = "#,##0.00"
                elif col_idx == 7:
                    cell.number_format = "0.0%"
                elif col_idx in (1, 8, 9):
                    cell.alignment = Alignment(horizontal="center")

                fill = self._status_color(st.status)
                if fill:
                    cell.fill = fill
                elif i % 2 == 0:
                    cell.fill = self.fill_zebra
            ws.row_dimensions[i].height = 20

        self._auto_fit(ws)

    # SHEET 5: REVISED CPM
    # ------------------------------------------------------------------
    def _build_revised_cpm(self, wb):
        ws = wb.create_sheet(title="05_Revised_CPM")
        self._create_banner(
            ws,
            "REVISED CPM SCHEDULE (REMAINING WORK)",
            f"Revised schedule starting from Status Date: {self.r.status_date}, duration: {self.r.revised_duration_days} days to {self.r.revised_completion_date or ''}",
            15,
        )

        ws.cell(row=4, column=1, value="REVISED SCHEDULE SUMMARY").font = \
            Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells("A4:O4")

        summary = [
            ("Status Date", str(self.r.status_date)),
            ("Revised Duration (from status date)", f"{self.r.revised_duration_days} Days"),
            ("Revised Completion Date", str(self.r.revised_completion_date or "")),
            ("Target Date", str(self.r.target_date)),
            ("Revised Critical Path", " -> ".join(self.r.revised_critical_path) if self.r.revised_critical_path else "None"),
            ("Target Feasible", "YES" if self.r.target_feasible else "NO"),
            ("Projected Delay", f"{self.r.projected_delay_days} Days"),
            ("Work Fronts", f"{self.r.work_fronts}"),
        ]
        curr = 5
        for label, val in summary:
            c1 = ws.cell(row=curr, column=1, value=label)
            c1.font = self.font_bold
            c1.fill = self.fill_card
            c1.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=2)

            c2 = ws.cell(row=curr, column=3, value=val)
            c2.font = self.font_regular
            c2.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=15)
            if "Feasible" in label:
                c2.fill = self.fill_pass if self.r.target_feasible else self.fill_crit
                c2.font = self.font_pass if self.r.target_feasible else self.font_critical
            curr += 1

        curr += 1
        headers = [
            "Activity ID", "Activity Name", "Remaining Qty", "Unit",
            "Crews", "Shifts", "Duration (d)", "ES", "EF", "LS", "LF",
            "Total Float", "Planned Start", "Planned Finish", "Critical / Status",
        ]
        self._apply_header(ws, curr, headers)
        curr += 1

        for ri, act in enumerate(self.r.revised_activities):
            is_done = (act.quantity <= 0 or act.duration_days == 0)
            crit_text = "COMPLETED" if is_done else ("CRITICAL" if act.is_critical else "")
            vals = [
                act.activity_id, act.name,
                act.quantity if act.quantity > 0 else 0.0,
                act.unit,
                act.crews if not is_done else "-",
                f"{act.shifts:.1f}" if not is_done else "-",
                act.duration_days, act.es, act.ef, act.ls, act.lf,
                act.total_float if not is_done else 0,
                str(act.planned_start or self.r.status_date),
                str(act.planned_finish or self.r.status_date),
                crit_text,
            ]
            for col_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=curr, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (5, 6, 7, 8, 9, 10, 11, 12, 15):
                    cell.alignment = Alignment(horizontal="center")
                if col_idx == 3 and isinstance(val, (int, float)):
                    cell.number_format = "#,##0.00"

                if is_done:
                    cell.fill = self.fill_pass
                    if col_idx == 15:
                        cell.font = self.font_pass
                elif act.is_critical:
                    cell.fill = self.fill_crit
                    if col_idx == 15:
                        cell.font = self.font_critical
                elif ri % 2 == 0:
                    cell.fill = self.fill_zebra

            ws.row_dimensions[curr].height = 20
            curr += 1

        self._auto_fit(ws)

    # SHEET 6: REVISED COST
    # ------------------------------------------------------------------
    def _build_revised_cost(self, wb):
        ws = wb.create_sheet(title="06_Revised_Cost")
        self._create_banner(
            ws,
            "REVISED COST ESTIMATE (REMAINING WORK)",
            f"Remaining work cost, acceleration surcharge, and project reconciliation as of Status Date: {self.r.status_date}",
            11,
        )
        headers = [
            "Activity ID", "Activity Name", "Remaining Qty", "Unit",
            "Crews", "Shifts", "Duration (d)",
            "Base Direct Cost (Rs)", "Acceleration Surcharge (Rs)",
            "Price Escalation (Rs)", "Total Remaining Cost (Rs)",
        ]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "A5"

        start_r = 5
        curr = start_r
        sum_base = 0.0
        sum_accel = 0.0
        sum_esc = 0.0
        sum_total = 0.0

        for act in self.r.revised_activities:
            if act.quantity <= 0:
                continue
            c = act.cost
            base_dir = max(0.0, c.direct_cost - c.crash_total_surcharge)
            accel_sur = c.crash_total_surcharge
            esc_amt = c.escalation_amount
            tot_amt = c.total_amount

            sum_base += base_dir
            sum_accel += accel_sur
            sum_esc += esc_amt
            sum_total += tot_amt

            vals = [
                act.activity_id, act.name, act.quantity, act.unit,
                act.crews, f"{act.shifts:.1f}", act.duration_days,
                base_dir, accel_sur, esc_amt, tot_amt,
            ]
            for col_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=curr, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (5, 6, 7):
                    cell.alignment = Alignment(horizontal="center")
                elif col_idx == 3:
                    cell.number_format = "#,##0.00"
                elif col_idx in (8, 9, 10, 11):
                    cell.number_format = "\u20B9#,##0.00"
                if (curr - start_r) % 2 == 1:
                    cell.fill = self.fill_zebra
            ws.row_dimensions[curr].height = 20
            curr += 1

        # Table Totals Row
        tot_table_r = curr
        ws.cell(row=tot_table_r, column=2, value="TOTAL REMAINING WORK").font = self.font_bold
        ws.cell(row=tot_table_r, column=8, value=sum_base).number_format = "\u20B9#,##0.00"
        ws.cell(row=tot_table_r, column=9, value=sum_accel).number_format = "\u20B9#,##0.00"
        ws.cell(row=tot_table_r, column=10, value=sum_esc).number_format = "\u20B9#,##0.00"
        ws.cell(row=tot_table_r, column=11, value=sum_total).number_format = "\u20B9#,##0.00"
        for col_idx in range(1, 12):
            c_cell = ws.cell(row=tot_table_r, column=col_idx)
            c_cell.font = self.font_bold
            c_cell.border = self.border_double
            c_cell.fill = self.fill_zebra
        ws.row_dimensions[tot_table_r].height = 24

        tot_r = tot_table_r + 2

        ws.cell(row=tot_r, column=2, value="PROJECT COST SUMMARY (RECONCILED)").font = self.font_bold
        cost_rows = [
            ("Spent Cost (Completed Work)", self.r.spent_cost_inr),
            ("Remaining Base Direct Cost (Normal)", sum_base),
            ("Additional Acceleration Cost", sum_accel),
            ("Remaining Price Escalation", sum_esc),
            ("Total Remaining Cost", sum_total),
            ("Total Revised Project Cost", self.r.total_project_cost_inr),
            ("Original Baseline Total Cost", self.r.baseline_total_cost_crore * 10000000),
            ("Cost Variance vs Baseline", self.r.total_project_cost_inr - (self.r.baseline_total_cost_crore * 10000000)),
        ]
        for j, (label, val) in enumerate(cost_rows):
            r = tot_r + 1 + j
            c1 = ws.cell(row=r, column=2, value=label)
            c1.font = self.font_bold
            c1.border = self.border_thin
            c1.fill = self.fill_card
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
            c2 = ws.cell(row=r, column=8, value=val)
            c2.font = self.font_bold
            c2.border = self.border_thin
            c2.number_format = "\u20B9#,##0.00"
            if "Total Revised" in label:
                c2.fill = self.fill_pass
            ws.row_dimensions[r].height = 22

        self._auto_fit(ws)

    # SHEET 7: S-CURVE COMPARISON
    # ------------------------------------------------------------------
    # SHEET 7: S-CURVE COMPARISON
    # ------------------------------------------------------------------
    def _build_scurve(self, wb):
        ws = wb.create_sheet(title="07_S_Curve_Comparison")
        self._create_banner(
            ws,
            "S-CURVE COMPARISON: BASELINE vs ACTUAL vs REVISED FORECAST",
            f"Physical cumulative progress trajectory as of Status Date: {self.r.status_date} | "
            f"Target Date: {self.r.target_date} | Projected Delay: {self.r.projected_delay_days} Days",
            8,
        )

        # Build a unified monthly timeline from project start to revised completion
        start = self.p_in.start_date
        rev_finish = self.r.revised_completion_date or start
        baseline_finish = start + timedelta(days=self.r.baseline_result.calculated_duration_days)
        target_finish = self.r.target_date or baseline_finish
        timeline_end = max(rev_finish, baseline_finish, target_finish)

        # Executive KPI Cards (rows 4-5)
        kpis = [
            ("Status Date", f"{self.r.status_date} ({self.r.actual_progress_pct:.1f}% Actual)"),
            ("Target Date", f"{self.r.target_date}"),
            ("Revised Completion", f"{rev_finish}"),
            ("Projected Delay", f"+{self.r.projected_delay_days} Days Delay" if self.r.projected_delay_days > 0 else "0 Days (On Target)"),
        ]
        for idx, (title, val) in enumerate(kpis):
            c_start = 1 + idx * 2
            c_end = c_start + 1
            ws.merge_cells(start_row=4, start_column=c_start, end_row=4, end_column=c_end)
            ws.merge_cells(start_row=5, start_column=c_start, end_row=5, end_column=c_end)

            c_t = ws.cell(row=4, column=c_start, value=title)
            c_t.font = Font(size=9, bold=True, color="555555")
            c_t.fill = self.fill_card
            c_t.alignment = Alignment(horizontal="center", vertical="center")
            c_t.border = self.border_thin

            c_v = ws.cell(row=5, column=c_start, value=val)
            c_v.font = Font(size=12, bold=True, color=self.c_navy)
            c_v.fill = self.fill_card
            c_v.alignment = Alignment(horizontal="center", vertical="center")
            c_v.border = self.border_thin

            if title == "Projected Delay":
                if self.r.projected_delay_days > 0:
                    c_v.fill = self.fill_crit
                    c_v.font = Font(size=12, bold=True, color=self.c_crit_fg)
                else:
                    c_v.fill = self.fill_pass
                    c_v.font = Font(size=12, bold=True, color=self.c_pass_fg)

        ws.row_dimensions[4].height = 18
        ws.row_dimensions[5].height = 26

        # Collect month starts
        months: List[date] = []
        cur = date(start.year, start.month, 1)
        while cur <= timeline_end:
            months.append(cur)
            nxt = date(cur.year + (cur.month == 12), (cur.month % 12) + 1, 1)
            cur = nxt

        status_date = self.r.status_date
        status_month = date(status_date.year, status_date.month, 1)
        target_date = self.r.target_date
        target_month = date(target_date.year, target_date.month, 1)
        projected_delay = self.r.projected_delay_days
        has_delay = projected_delay > 0 and rev_finish > target_date

        # Baseline cumulative map by month date
        baseline_monthly = self.r.baseline_result.monthly_progress
        baseline_cum_by_date: Dict[date, float] = {}
        for bm in baseline_monthly:
            d_mid = start + timedelta(days=(bm.period_start_day + bm.period_end_day) // 2)
            m_key = date(d_mid.year, d_mid.month, 1)
            baseline_cum_by_date[m_key] = bm.cumulative_planned_progress_pct

        # Revised monthly cumulative list (of remaining work)
        revised_monthly = self.r.revised_monthly_progress
        revised_cum: List[float] = [m.cumulative_planned_progress_pct for m in revised_monthly]

        headers = [
            "Month",
            "Baseline Planned %",
            "Actual Progress %",
            "Revised Forecast (Target) %",
            "Delay Period Forecast %",
            "Total Cumulative Forecast %",
            "Schedule Delay Impact",
            "Notes & Milestone Markers",
        ]
        self._apply_header(ws, 7, headers)

        start_row = 8
        prev_revised = self.r.actual_progress_pct

        for i, month_start in enumerate(months):
            month_label = month_start.strftime("%b\n%Y")

            # 1. Baseline Planned
            if month_start in baseline_cum_by_date:
                base_val = baseline_cum_by_date[month_start]
            else:
                day = (month_start - start).days
                base_val = self._interp_baseline(day)
            base_val = min(100.0, max(0.0, base_val))

            # 2. Actual Progress (recorded up to status date only)
            actual_val = None
            if month_start < status_month:
                day = max(0, (month_start - start).days)
                sday = max(1, (status_date - start).days)
                actual_val = min(self.r.actual_progress_pct,
                                 self.r.actual_progress_pct * (day / sday))
            elif month_start == status_month:
                actual_val = self.r.actual_progress_pct

            # 3. Total Cumulative Forecast %
            total_forecast_val = None
            if month_start < status_month:
                total_forecast_val = None
            elif month_start == status_month:
                total_forecast_val = self.r.actual_progress_pct
                prev_revised = total_forecast_val
            else:
                k = (month_start.year - status_month.year) * 12 + (month_start.month - status_month.month)
                rem_cum = 0.0
                if 0 <= k - 1 < len(revised_cum):
                    rem_cum = revised_cum[k - 1]
                elif k - 1 >= len(revised_cum):
                    rem_cum = 100.0

                p_act = self.r.actual_progress_pct
                overall_pct = p_act + (rem_cum * (100.0 - p_act) / 100.0)
                if i == len(months) - 1 or month_start >= date(rev_finish.year, rev_finish.month, 1):
                    overall_pct = 100.0
                total_forecast_val = min(100.0, max(prev_revised, overall_pct))
                prev_revised = total_forecast_val

            # 4. Split into On-Target Forecast vs Delay Period Forecast
            revised_target_val = None
            delay_period_val = None
            is_delay_row = False

            if total_forecast_val is not None:
                if not has_delay:
                    revised_target_val = total_forecast_val
                    delay_period_val = None
                else:
                    if target_month < date(rev_finish.year, rev_finish.month, 1):
                        if month_start < target_month:
                            revised_target_val = total_forecast_val
                            delay_period_val = None
                        elif month_start == target_month:
                            revised_target_val = total_forecast_val
                            delay_period_val = total_forecast_val  # Connect line seamlessly!
                        else:
                            # Beyond target month = DELAY PERIOD
                            revised_target_val = None
                            delay_period_val = total_forecast_val
                            is_delay_row = True
                    else:
                        # target_month == rev_finish_month, but rev_finish > target_date
                        if month_start < target_month:
                            revised_target_val = total_forecast_val
                            if i + 1 < len(months) and months[i + 1] == target_month:
                                delay_period_val = total_forecast_val
                        else:
                            revised_target_val = None
                            delay_period_val = total_forecast_val
                            is_delay_row = True

            # Delay impact text & notes
            delay_text = "-"
            notes = ""
            if month_start == status_month:
                notes = f"<-- Current Status Date ({self.r.actual_progress_pct:.1f}% Actual)"
            elif month_start == target_month:
                notes = f"<-- Target Date ({target_date.isoformat()})"
                if not is_delay_row:
                    delay_text = "Target Milestone"

            if is_delay_row:
                delay_text = f"+{projected_delay} Days Delay"
                notes = f"⚠️ DELAY PERIOD: Exceeds target date ({target_date.isoformat()}) by {projected_delay} days"
            elif total_forecast_val is not None and not is_delay_row:
                delay_text = "On Schedule"

            vals = [
                month_label,
                base_val / 100.0 if base_val is not None else None,
                actual_val / 100.0 if actual_val is not None else None,
                revised_target_val / 100.0 if revised_target_val is not None else None,
                delay_period_val / 100.0 if delay_period_val is not None else None,
                total_forecast_val / 100.0 if total_forecast_val is not None else None,
                delay_text,
                notes,
            ]

            row_idx = start_row + i
            for col_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                cell.alignment = Alignment(
                    horizontal="center" if col_idx == 1 else "right",
                    vertical="center",
                    wrap_text=col_idx == 1,
                )
                if col_idx in (2, 3, 4, 5, 6):
                    cell.number_format = "0.0%"
                    cell.alignment = Alignment(horizontal="right")
                elif col_idx in (1, 7):
                    cell.alignment = Alignment(horizontal="center")

                # Highlighting delay days and milestones in distinct colors
                if is_delay_row:
                    # Clear identification of delay days in alert red
                    cell.fill = self.fill_crit
                    if col_idx in (1, 5, 6, 7):
                        cell.font = self.font_critical
                    elif col_idx == 8:
                        cell.font = Font(name="Calibri", size=9, bold=True, color=self.c_crit_fg)
                elif month_start == status_month:
                    if col_idx in (1, 3):
                        cell.fill = self.fill_actual
                        cell.font = Font(name="Calibri", size=10, bold=True, color=self.c_actual)
                    elif col_idx == 8:
                        cell.font = Font(name="Calibri", size=9, bold=True, color=self.c_gold)
                elif month_start == target_month and not is_delay_row:
                    if col_idx in (1, 4, 7):
                        cell.fill = PatternFill("solid", start_color="FFF8E1", end_color="FFF8E1")
                        cell.font = Font(name="Calibri", size=10, bold=True, color=self.c_gold)
                    elif col_idx == 8:
                        cell.font = Font(name="Calibri", size=9, bold=True, color=self.c_gold)
                elif row_idx % 2 == 0:
                    cell.fill = self.fill_zebra

            ws.row_dimensions[row_idx].height = 30

        end_row = start_row + len(months) - 1

        # Build LineChart
        chart = LineChart()
        chart.title = "Project S-Curve: Planned Baseline vs. Actual Progress vs. Revised Forecast"
        chart.style = 13
        chart.y_axis.title = "Cumulative Physical Progress (%)"
        chart.x_axis.title = "Project Timeline (Months)"
        chart.y_axis.scaling.min = 0.0
        chart.y_axis.scaling.max = 1.0
        chart.y_axis.majorUnit = 0.1
        chart.y_axis.number_format = "0%"
        chart.y_axis.majorGridlines = ChartLines()
        chart.y_axis.tickLblPos = "low"
        chart.y_axis.delete = False
        chart.x_axis.tickLblSkip = 1
        chart.x_axis.tickMarkSkip = 1
        chart.x_axis.tickLblPos = "low"
        chart.x_axis.delete = False
        chart.width = 25
        chart.height = 15
        chart.legend.legendPos = "b"

        planned_and_actual = Reference(
            ws, min_col=2, min_row=7, max_col=3, max_row=end_row
        )
        chart.add_data(planned_and_actual, titles_from_data=True)
        revised_forecast = Reference(
            ws, min_col=6, min_row=7, max_col=6, max_row=end_row
        )
        chart.add_data(revised_forecast, titles_from_data=True)
        chart.series[2].tx = openpyxl.chart.series.SeriesLabel(
            v="Revised Forecast %"
        )
        cats = Reference(ws, min_col=1, min_row=start_row, max_row=end_row)
        chart.set_categories(cats)
        category_ref = StrRef(
            f=f"'{ws.title}'!$A${start_row}:$A${end_row}"
        )
        for series in chart.series:
            series.cat = AxDataSource(strRef=category_ref)

        # Plot the total revised forecast as one series, including any delay.
        colors = [self.c_navy, self.c_actual, self.c_revised]
        for idx, series in enumerate(chart.series):
            if idx < len(colors):
                series.graphicalProperties.line.solidFill = colors[idx]
                series.graphicalProperties.line.width = 25000
                series.smooth = True

        ws.add_chart(chart, "J4")
        self._auto_fit(ws)
        ws.column_dimensions["A"].width = 14

    def _interp_baseline(self, day: int) -> float:
        """Linear interpolation of baseline cumulative progress at a day ordinal."""
        if day <= 0:
            return 0.0
        base_dur = max(self.r.baseline_result.calculated_duration_days, 1)
        if day >= base_dur:
            return 100.0
        total = 0.0
        for act in self.r.baseline_result.activities:
            act_dur = max(act.duration_days, 1)
            if day <= act.es:
                comp = 0.0
            elif day >= act.ef:
                comp = 1.0
            else:
                comp = (day - act.es) / act_dur
            total += comp * act.progress_weight_pct
        return min(100.0, total)

    # SHEET 8: TARGET FEASIBILITY
    # ------------------------------------------------------------------
    def _build_feasibility(self, wb):
        ws = wb.create_sheet(title="08_Target_Feasibility")
        self._create_banner(
            ws,
            "TARGET FEASIBILITY & RESOURCE CHANGES",
            f"Assessment of whether the target date and progress are achievable as of Status Date: {self.r.status_date}",
            7,
        )

        r = self.r
        ws.cell(row=4, column=1, value="FEASIBILITY ASSESSMENT").font = \
            Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells("A4:G4")

        rows = [
            ("Status Date", str(r.status_date)),
            ("Target Date", str(r.target_date)),
            ("Available Calendar Days", str(r.required_calendar_days)),
            ("Actual Progress at Status Date", f"{r.actual_progress_pct:.1f}%"),
            ("Baseline Progress at Status Date", f"{r.baseline_progress_at_status_date:.1f}%"),
            ("Progress Gap (Behind)", f"{r.progress_gap_pct:.1f} pp"),
            ("Target Progress", f"{r.target_progress_pct:.1f}%"),
            ("Required Progress (Remaining)", f"{r.required_progress_pct:.1f} pp"),
            ("Revised Duration (from status)", f"{r.revised_duration_days} Days"),
            ("Revised Completion Date", str(r.revised_completion_date or "")),
            ("Original Revised Completion (no accel.)", str(r.original_revised_completion_date or "")),
            ("Projected Delay vs Target", f"{r.projected_delay_days} Days"),
            ("Projected Progress at Target Date", f"{r.projected_progress_at_target_date:.1f}%"),
            ("Target Feasible", "YES" if r.target_feasible else "NO"),
            ("Additional Acceleration Cost", f"Rs {r.additional_acceleration_cost_crore:,.2f} Cr"),
            ("Work Fronts (Revised)", f"{r.work_fronts}"),
        ]

        curr = 5
        for label, val in rows:
            c1 = ws.cell(row=curr, column=1, value=label)
            c1.font = self.font_bold
            c1.fill = self.fill_card
            c1.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=3)

            c2 = ws.cell(row=curr, column=4, value=val)
            c2.font = self.font_regular
            c2.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=4, end_row=curr, end_column=7)

            if "Feasible" in label:
                if r.target_feasible:
                    c2.fill = self.fill_pass
                    c2.font = self.font_pass
                else:
                    c2.fill = self.fill_crit
                    c2.font = self.font_critical
            elif "Delay" in label and r.projected_delay_days > 0:
                c2.fill = self.fill_crit
                c2.font = self.font_critical

            ws.row_dimensions[curr].height = 20
            curr += 1

        # Resource changes table
        curr += 2
        ws.cell(row=curr, column=1,
                value="RESOURCE CHANGES APPLIED (ACCELERATION)").font = \
            Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=7)
        curr += 1

        if r.resources_added_summary:
            headers = [
                "Activity ID", "Activity Name", "From Crews", "To Crews",
                "Fronts", "Shifts", "Days Saved", "Cost Increase (Lakhs)",
            ]
            self._apply_header(ws, curr, headers)
            curr += 1
            for item in r.resources_added_summary:
                vals = [
                    item.get("activity_id", ""),
                    item.get("activity_name", ""),
                    item.get("baseline_crews", ""),
                    item.get("optimized_crews", ""),
                    item.get("work_fronts", ""),
                    f"{item.get('shifts', 1.0):.1f}",
                    item.get("days_saved", ""),
                    item.get("cost_increase_lakhs", ""),
                ]
                for col_idx, val in enumerate(vals, 1):
                    cell = ws.cell(row=curr, column=col_idx, value=val)
                    cell.border = self.border_thin
                    cell.font = self.font_regular
                    if col_idx in (1, 3, 4, 5, 6, 7):
                        cell.alignment = Alignment(horizontal="center")
                    elif col_idx == 8:
                        cell.number_format = "#,##0.00"
                curr += 1
        else:
            ws.cell(row=curr, column=1,
                    value="No resource changes were needed — the revised schedule "
                          "meets the target date at current resource levels.").font = \
                self.font_regular
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=7)

        self._auto_fit(ws)
