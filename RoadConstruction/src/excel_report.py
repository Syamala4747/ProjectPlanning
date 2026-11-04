"""
Excel Report Generation Module using openpyxl.
Produces an executive-grade 14-sheet workbook with strict separation between
BASELINE (normal resources) and OPTIMIZED (constraint-based schedule satisfying duration & budget),
featuring an active hierarchical Work Breakdown Structure (WBS) engine.
"""
import os
from typing import List, Dict, Any
from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

from .models import ProjectScheduleResult, RoadActivity, MonthlyProgressItem
from .wbs_generator import WBSGenerator, WBSNode

class ExcelReportGenerator:
    def __init__(self, result: ProjectScheduleResult, output_path: str = "road_project_schedule.xlsx"):
        self.res = result
        self.p_in = result.project_input
        self.opt_acts = result.activities
        self.base_acts = result.baseline_activities if result.baseline_activities else result.activities
        self.output_path = output_path
        
        # Color Palette - Modern Corporate Navy
        self.c_navy = "1B365D"       # Primary Headers
        self.c_slate = "4A777A"      # Subheaders
        self.c_accent = "2E75B6"     # Accent Blue
        self.c_light_bg = "F0F4F8"   # Alternate rows
        self.c_crit_bg = "FFEBEE"    # Critical soft red
        self.c_crit_fg = "C62828"    # Critical bold red
        self.c_pass_bg = "E8F5E9"    # Green
        self.c_pass_fg = "2E7D32"
        self.c_border = "D0D7DE"
        self.c_card_bg = "F8FAFC"
        self.c_gold = "D97706"

        # Styles
        self.font_title = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
        self.font_subtitle = Font(name="Calibri", size=11, italic=True, color="E0E0E0")
        self.font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        self.font_bold = Font(name="Calibri", size=11, bold=True, color="000000")
        self.font_regular = Font(name="Calibri", size=10, color="000000")
        self.font_small = Font(name="Calibri", size=9, color="555555")
        self.font_critical = Font(name="Calibri", size=10, bold=True, color=self.c_crit_fg)
        self.font_pass = Font(name="Calibri", size=10, bold=True, color=self.c_pass_fg)
        
        self.fill_primary = PatternFill(start_color=self.c_navy, end_color=self.c_navy, fill_type="solid")
        self.fill_accent = PatternFill(start_color=self.c_accent, end_color=self.c_accent, fill_type="solid")
        self.fill_zebra = PatternFill(start_color=self.c_light_bg, end_color=self.c_light_bg, fill_type="solid")
        self.fill_crit = PatternFill(start_color=self.c_crit_bg, end_color=self.c_crit_bg, fill_type="solid")
        self.fill_pass = PatternFill(start_color=self.c_pass_bg, end_color=self.c_pass_bg, fill_type="solid")
        self.fill_card = PatternFill(start_color=self.c_card_bg, end_color=self.c_card_bg, fill_type="solid")
        
        self.border_thin = Border(
            left=Side(style='thin', color=self.c_border),
            right=Side(style='thin', color=self.c_border),
            top=Side(style='thin', color=self.c_border),
            bottom=Side(style='thin', color=self.c_border)
        )
        self.border_double = Border(bottom=Side(style='double', color='000000'), top=Side(style='thin', color='000000'))

    def generate(self) -> str:
        """Builds all 14 sheets exactly as requested and saves workbook."""
        out_dir = os.path.dirname(self.output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
            
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove default sheet

        self._build_sheet_01_project_summary(wb)
        self._build_sheet_02_project_inputs(wb)
        self._build_sheet_03_assumptions(wb)
        self._build_sheet_04_wbs_hierarchy(wb)
        self._build_sheet_05_activities(wb)
        self._build_sheet_06_baseline_cpm(wb)
        self._build_sheet_07_baseline_cost(wb)
        self._build_sheet_08_constraint_optimization(wb)
        self._build_sheet_09_optimized_cpm(wb)
        self._build_sheet_10_optimized_cost(wb)
        self._build_sheet_11_baseline_vs_optimized(wb)
        self._build_sheet_12_monthly_schedule(wb)
        self._build_sheet_13_scurve(wb)
        self._build_sheet_14_critical_path(wb)
        self._build_sheet_15_validation(wb)

        try:
            wb.save(self.output_path)
            return self.output_path
        except PermissionError:
            base, ext = os.path.splitext(self.output_path)
            fallback_path = f"{base}_new{ext}"
            wb.save(fallback_path)
            print(f"\n⚠️  Note: '{self.output_path}' was open in Excel. Saved as '{fallback_path}' instead.")
            return fallback_path

    def _build_sheet_15_validation(self, wb):
        """Write the structured validation and sanity-check result."""
        ws = wb.create_sheet(title="15_Validation_Checks")
        report = getattr(self.res, "validation_report", None)
        status = getattr(report, "status", "NOT RUN")
        errors = list(getattr(report, "errors", []) or [])
        warnings = list(getattr(report, "warnings", []) or [])
        issues = list(getattr(report, "issues", []) or [])

        self._create_banner(
            ws,
            "VALIDATION & SANITY CHECK RESULTS",
            "Complete validation register: PASS checks, review warnings, and hard errors",
            7,
        )
        ws.cell(row=4, column=1, value="Overall status").font = self.font_bold
        ws.cell(row=4, column=2, value=status)
        ws.cell(row=5, column=1, value="Error count").font = self.font_bold
        ws.cell(row=5, column=2, value=len(errors))
        ws.cell(row=6, column=1, value="Warning count").font = self.font_bold
        ws.cell(row=6, column=2, value=len(warnings))
        ws.cell(row=7, column=1, value="Passed check count").font = self.font_bold
        check_catalog = [
            ("NO_ACTIVITIES", "Activity list is present"),
            ("QUANTITY_FORMULA", "Pipeline quantities and soil partition reconcile"),
            ("SOIL_SHARES_NOT_100", "Soil percentages total 100%"),
            ("ACTIVITY_MASTER", "Scheduled activities exist in the activity master"),
            ("ZERO_ACTIVE_QUANTITY", "Active activities have positive quantities"),
            ("CIRCULAR_DEPENDENCY", "Activity network is cycle-free"),
            ("MISSING_PREDECESSOR", "All predecessors exist"),
            ("PRECEDENCE_VIOLATION", "Predecessor dates are respected"),
            ("ENGINEERING_HANDOVER_REVIEW", "Engineering handovers are represented"),
            ("PRODUCTIVITY_DURATION_MISMATCH", "Effort days match quantity/productivity/crews"),
            ("CALENDAR_DURATION_MISMATCH", "Calendar duration matches working-day rules"),
            ("INVALID_ACTIVITY_DATES", "Activity dates are valid"),
            ("CPM_EF_IDENTITY", "CPM EF equals ES plus duration"),
            ("CPM_LF_IDENTITY", "CPM LF equals LS plus duration"),
            ("CPM_FLOAT_IDENTITY", "CPM float equals LS minus ES"),
            ("NEGATIVE_FLOAT", "CPM floats are non-negative"),
            ("INVALID_CREW_COUNT", "Crew allocations are valid"),
            ("INVALID_SHIFT_COUNT", "Shift allocations are valid"),
            ("CREW_CAP_EXCEEDED", "Crew caps are respected"),
            ("UNIT_RATE_BUILDUP_MISMATCH", "Planned unit-rate build-up reconciles"),
            ("INVALID_PLANNED_RATE", "Planned rates are positive"),
            ("DIRECT_COST_RECONCILIATION", "Direct cost reconciles"),
            ("TOTAL_COST_RECONCILIATION", "Total cost reconciles"),
            ("EMBEDDED_ONCOST_DUPLICATION", "Embedded overhead/profit are not double-counted"),
            ("WASTAGE_RECONCILIATION", "Separate wastage components reconcile"),
            ("ESCALATION_OUTPUT_INVALID", "Escalation outputs are valid"),
            ("PROVISIONAL_ESCALATION", "Escalation source and clause are approved"),
            ("PROJECT_DURATION_RECONCILIATION", "Project finish reconciles to activity network"),
            ("FEASIBILITY_DURATION_MISMATCH", "Duration constraint is satisfied"),
            ("FEASIBILITY_BUDGET_MISMATCH", "Budget constraint is satisfied"),
            ("MONTHLY_COST_RECONCILIATION", "Monthly costs reconcile to project total"),
            ("PROGRESS_NOT_COMPLETE", "Monthly progress reaches 100%"),
            ("WORKBOOK_ACTIVITY_SET_MISMATCH", "Workbook activity set matches schedule"),
            ("WORKBOOK_SCURVE_PROGRESS_MISMATCH", "Workbook S-curve progress reconciles"),
            ("WORKBOOK_SCURVE_COST_MISMATCH", "Workbook S-curve costs reconcile"),
            ("SCHEDULER_SANITY", "Scheduler planning-risk warnings"),
        ]
        for row in range(4, 8):
            for col in range(1, 3):
                cell = ws.cell(row=row, column=col)
                cell.border = self.border_thin
                if col == 2:
                    cell.fill = self.fill_pass if (
                        (row == 4 and status == "PASS") or
                        (row == 7) or
                        (row != 4 and cell.value == 0)
                    ) else self.fill_crit

        ws.cell(row=8, column=1, value="Interpretation").font = self.font_bold
        ws.merge_cells(start_row=8, start_column=2, end_row=8, end_column=7)
        ws.cell(
            row=8,
            column=2,
            value=(
                "ERROR rows are hard reconciliation failures and stop plan delivery. "
                "WARNING rows are sanity or source-data review items and do not alter totals."
            ),
        )
        ws.cell(row=8, column=2).alignment = Alignment(wrap_text=True)

        headers = ["Severity", "Check code", "Activity ID", "Result", "Check / expected rule", "Finding", "Action"]
        self._apply_header(ws, 10, headers)
        issue_codes = {str(issue.code) for issue in issues}
        rows = [(
            str(issue.severity),
            issue.code,
            issue.activity_id or "-",
            "ERROR" if issue.severity == "ERROR" else "WARNING",
            next((description for code, description in check_catalog if code == issue.code), "Exception reported by validation engine"),
            issue.message,
            "Correct calculation before issue" if issue.severity == "ERROR" else "Review/confirm assumption or source data",
        ) for issue in issues]
        rows.extend((
            "PASS",
            code,
            "-",
            "PASS",
            description,
            "Check completed successfully; no exception was reported.",
            "No action required",
        ) for code, description in check_catalog if code not in issue_codes)
        ws.cell(row=7, column=2, value=sum(1 for row in rows if row[0] == "PASS"))

        for row_index, values in enumerate(rows, start=11):
            severity = str(values[0])
            for col_index, value in enumerate(values, start=1):
                cell = ws.cell(row=row_index, column=col_index, value=value)
                cell.border = self.border_thin
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.font = self.font_critical if severity == "ERROR" else self.font_regular
                cell.fill = (
                    self.fill_crit if severity == "ERROR"
                    else self.fill_pass if severity == "PASS"
                    else self.fill_card
                )

        if not rows:
            ws.cell(row=11, column=1, value="PASS")
            ws.cell(row=11, column=2, value="NO_ISSUES")
            ws.cell(row=11, column=4, value="PASS")
            ws.cell(row=11, column=5, value="All validation and sanity checks passed.")
            for col in range(1, 8):
                ws.cell(row=11, column=col).fill = self.fill_pass
                ws.cell(row=11, column=col).border = self.border_thin

        ws.freeze_panes = "A11"
        ws.auto_filter.ref = f"A10:G{max(11, 10 + len(rows))}"
        ws.column_dimensions["A"].width = 13
        ws.column_dimensions["B"].width = 32
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 48
        ws.column_dimensions["F"].width = 75
        ws.column_dimensions["G"].width = 32

    def _apply_header(self, ws, row: int, titles: List[str]):
        ws.row_dimensions[row].height = 26
        for col, title in enumerate(titles, 1):
            cell = ws.cell(row=row, column=col, value=title)
            cell.font = self.font_header
            cell.fill = self.fill_primary
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = self.border_thin

    def _auto_fit_columns(self, ws, max_len_cap: int = 50):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = cell.value
                if val is not None:
                    val_str = str(val)
                    if len(val_str) < 80:
                        max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    def _create_banner(self, ws, title: str, subtitle: str, max_cols: int = 10):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
        ws.row_dimensions[1].height = 32
        cell1 = ws.cell(row=1, column=1, value=title)
        cell1.font = self.font_title
        cell1.fill = self.fill_primary
        cell1.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_cols)
        ws.row_dimensions[2].height = 20
        cell2 = ws.cell(row=2, column=1, value=subtitle)
        cell2.font = self.font_subtitle
        cell2.fill = self.fill_accent
        cell2.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    @staticmethod
    def _format_productivity_str(rate: float, unit: str) -> str:
        if rate is None:
            return f"- {unit}"
        r_float = float(rate)
        if r_float.is_integer():
            return f"{int(r_float)} {unit}"
        elif round(r_float, 2) == round(r_float, 1):
            return f"{r_float:.2f} {unit}"
        else:
            return f"{r_float:g} {unit}"

    # -------------------------------------------------------------
    # SHEET 1: PROJECT SUMMARY
    # -------------------------------------------------------------
    # ------------------------------------------------------------------
    # Helpers — project-type-aware display strings
    # ------------------------------------------------------------------
    def _dashboard_title(self) -> str:
        ptype = getattr(self.p_in, "project_type", "road") or "road"
        return {
            "pipeline": "WATER PIPELINE PROJECT SUMMARY DASHBOARD",
            "building": "BUILDING CONSTRUCTION PROJECT SUMMARY DASHBOARD",
        }.get(ptype, "ROAD CONSTRUCTION PROJECT SUMMARY DASHBOARD")

    def _scope_label(self) -> str:
        ptype = getattr(self.p_in, "project_type", "road") or "road"
        if ptype == "pipeline":
            mat = getattr(self.p_in, "pipe_material", "DI") or "DI"
            dia = int(getattr(self.p_in, "pipe_diameter_mm", 0) or 0)
            return f"{self.p_in.length_km:.1f} km, DN{dia} {mat} Pipeline"
        if ptype == "building":
            bua = getattr(self.p_in, "built_up_area_sqm", 0) or 0
            floors = getattr(self.p_in, "number_of_floors", 1) or 1
            btype = (getattr(self.p_in, "building_type", "Residential") or "Residential").title()
            return f"{bua:,.0f} sqm, G+{floors - 1} {btype} Building"
        return f"{self.p_in.length_km:.1f} km, {self.p_in.lanes}-Lane Highway"

    def _source_rate_label(self) -> str:
        ptype = getattr(self.p_in, "project_type", "road") or "road"
        return {
            "pipeline": "Source Unit Rate",
            "building": "CPWD Source Unit Rate",
        }.get(ptype, "MoRTH Direct Unit Rate")

    def _spec_ref_label(self) -> str:
        ptype = getattr(self.p_in, "project_type", "road") or "road"
        return {
            "pipeline": "CPHEEO / IS Specification Reference",
            "building": "CPWD DSR / IS Specification Reference",
        }.get(ptype, "MoRTH Specification Reference")

    def _rate_basis_subtitle(self) -> str:
        ptype = getattr(self.p_in, "project_type", "road") or "road"
        return {
            "pipeline": "Direct Unit Rate (CPHEEO/CPWD source) | Planning Variation (%) | Planned Unit Rate (source + statutory OH/CP) | Separate wastage",
            "building": "Direct Unit Rate (CPWD DSR source) | Planning Variation (%) | Planned Unit Rate (ALL-IN) | OH & CP (embedded in rate)",
        }.get(ptype, "Direct Unit Rate (MoRTH source) | Planning Variation (%) | Planned Unit Rate (ALL-IN incl. on-costs) | OH & CP (embedded in rate)")

    def _baseline_cost_subtitle(self) -> str:
        ptype = getattr(self.p_in, "project_type", "road") or "road"
        return {
            "pipeline": "Standard rate analysis: CPHEEO/CPWD source rate + statutory OH/profit; wastage is a separate direct-cost line (Not forced to equal budget)",
            "building": "Standard rate analysis based on CPWD DSR 2021 (Not forced to equal budget)",
        }.get(ptype, "Standard rate analysis based on MoRTH Standard Data Book (Not forced to equal budget)")

    def _activities_banner(self) -> str:
        ptype = getattr(self.p_in, "project_type", "road") or "road"
        return {
            "pipeline": "PIPELINE CONSTRUCTION MASTER ACTIVITIES DIRECTORY",
            "building": "BUILDING CONSTRUCTION MASTER ACTIVITIES DIRECTORY",
        }.get(ptype, "ROAD CONSTRUCTION MASTER ACTIVITIES DIRECTORY")

    def _project_type_inputs_list(self) -> list:
        """Return the project-type-specific engineering parameter rows for Sheet 02."""
        ptype = getattr(self.p_in, "project_type", "road") or "road"

        common_head = [
            ("Required Completion Duration", self.p_in.duration_days, "Days", "User Constraint", "Required Duration = CONSTRAINT"),
            ("Maximum Available Budget", self.p_in.budget_crore, "Crore (INR)", "User Constraint", "Maximum Budget = CONSTRAINT"),
            ("Project Start Date", str(self.p_in.start_date), "Date", "Scheduling", "Baseline Start Date"),
        ]
        common_tail = [
            ("Working Days per Week", 6, "Days", "Calendar", "assumptions.yaml: working_days_per_week"),
            ("Productive Hours per Day", 8, "Hours", "Calendar", "assumptions.yaml: productive_hours_per_day (equipment utilisation 9 h/day)"),
            ("Statutory GST", "18.0% (excluded from totals)", "Tax %", "Statutory Dues", "Rates are quoted ex-GST; GST is added at invoice, not in the estimate"),
        ]

        if ptype == "pipeline":
            dia = int(getattr(self.p_in, "pipe_diameter_mm", 0) or 0)
            mat = (getattr(self.p_in, "pipe_material", "DI") or "DI").upper()
            di_class = "K7" if dia <= 300 else "K9"
            pipe_std = {"DI": f"IS 8329 ({di_class} DI pipe)", "HDPE": "IS 4984 (PE 100)", "MS": "IS 3589 (ERW/SAW)", "RCC": "IS 458 (NP4)"}.get(mat, "IS standard")
            td = getattr(self.p_in, "trench_depth_m", 2.0) or 2.0
            tw = getattr(self.p_in, "trench_width_m", 0.0) or 0.0
            fronts = getattr(self.p_in, "number_of_work_fronts", 1) or 1
            method = (getattr(self.p_in, "installation_method", "open_trench") or "open_trench").replace("_", " ").title()
            pt = (getattr(self.p_in, "pipeline_type", "water") or "water").title()
            specific = [
                ("Pipeline Type", pt, "", "Scope", "Water supply / Sewerage / Gas"),
                ("Total Pipeline Length", self.p_in.length_km, "km", "Scope", "User Input"),
                ("Pipe Diameter (Nominal)", dia, "mm", "Pipe Specification", f"DN{dia} — {pipe_std}"),
                ("Pipe Material", mat, "", "Pipe Specification", pipe_std),
                ("Terrain Type", self.p_in.terrain.title(), "", "Scope", "Affects productivity factor"),
                ("Installation Method", method, "", "Construction Method", "Open Trench / Trenchless (HDD)"),
                ("Trench Depth (Average)", td, "m", "Engineering Input", f"Avg cover = {td:.1f}m (CPHEEO 6.2)"),
                ("Trench Width (Excavated)", round(tw, 2), "m", "Engineering Input", f"Pipe OD + 300mm clearance each side"),
                ("Number of Work Fronts", fronts, "Fronts", "Construction Method", "Parallel simultaneous trench-lay gangs"),
                ("Joint Spacing (approx)", 5.5 if mat != "HDPE" else 6.0, "m", "Pipe Specification", "Rubber-ring (DI/RCC) or butt-fused (HDPE)"),
                ("Bedding Type", "Granular (IS 12288)", "", "Engineering Input", "150mm compacted granular bedding"),
                ("Hydrotest Pressure", "1.5 × working pressure", "", "QA/Testing", "CPHEEO 6.7.1; IS 12288 Clause 8"),
                ("Road Crossings (Declared)", int(getattr(self.p_in, "number_of_road_crossings", 0) or 0), "Nos", "Crossings",
                 "0 = derive open-cut road crossings from length"),
                ("Highway Crossings (Declared)", int(getattr(self.p_in, "number_of_highway_crossings", 0) or 0), "Nos", "Crossings",
                 "NH/SH crossings are never invented"),
                ("Rail Crossings (Declared)", int(getattr(self.p_in, "number_of_rail_crossings", 0) or 0), "Nos", "Crossings",
                 "Rail crossings are never invented"),
                ("HDD / Trenchless Crossings (Declared)", int(getattr(self.p_in, "number_of_hdd_crossings", 0) or 0), "Nos", "Crossings",
                 "Trenchless crossings are never invented"),
                ("River Crossings (Declared)", int(getattr(self.p_in, "number_of_river_crossings", 0) or 0), "Nos", "Crossings",
                 "River crossings are never invented"),
            ]
            # Effective scheme-level scope supplied/derived for the water-supply
            # package. Zero facilities mean the package is out of scope and the
            # run is a pure transmission-corridor estimate.
            scope = getattr(self.res, "pipeline_scope", {}) or {}
            # `reference_wtp_mld` is always present, so test only the scope
            # switches. A pure corridor run (all zero) must not print empty
            # scheme-facility rows.
            if any(scope.get(k) for k in
                   ("villages", "ohts", "wtp_mld", "ibps", "fhtc")):
                wtp_mld = float(scope.get("wtp_mld", 0.0) or 0.0)
                trains = float(scope.get("wtp_trains", 0.0) or 0.0)
                ref_mld = float(scope.get("reference_wtp_mld", 0.0) or 0.0)
                specific += [
                    ("Villages Served", int(scope.get("villages", 0) or 0), "Nos", "Scheme Scope",
                     "Habitations covered by the water-supply scheme"),
                    ("Overhead Service Reservoirs (OHT/ESR)", int(scope.get("ohts", 0) or 0), "Nos", "Scheme Facilities",
                     "Village-derived at 1 OHT/ESR per 3 villages unless overridden"),
                    ("Water Treatment Plant Capacity", wtp_mld, "MLD", "Scheme Facilities",
                     f"{trains:g} × {ref_mld:g} MLD reference train(s) — CPHEEO treatment train"),
                    ("Treatment Trains / CWR Units", round(trains, 3), "Nos", "Scheme Facilities",
                     "Civil and E&M packages scale with MLD / reference MLD"),
                    ("Intake / Booster Pumping Stations", int(scope.get("ibps", 0) or 0), "Nos", "Scheme Facilities",
                     "Intake well-cum-pump house, boosting and allied E&M"),
                    ("Functional Household Tap Connections", int(scope.get("fhtc", 0) or 0), "Nos", "Scheme Scope",
                     "200 FHTC per village unless overridden"),
                ]
        elif ptype == "building":
            bua = getattr(self.p_in, "built_up_area_sqm", 0) or 0
            floors = getattr(self.p_in, "number_of_floors", 1) or 1
            btype = (getattr(self.p_in, "building_type", "residential") or "residential").title()
            ctype = (getattr(self.p_in, "construction_type", "rcc") or "rcc").upper()
            quality = (getattr(self.p_in, "quality_finish", "standard") or "standard").title()
            basement = "Yes" if getattr(self.p_in, "basement", False) else "No"
            blocks = getattr(self.p_in, "number_of_blocks", 1) or 1
            const_std = {"RCC": "IS 456 (RCC Frame)", "LOAD_BEARING": "IS 1905 (Masonry)", "STEEL": "IS 800 (Steel)"}.get(ctype, "IS standard")
            specific = [
                ("Building Type", btype, "", "Scope", "Residential / Commercial"),
                ("Built-up Area (Total)", bua, "sqm", "Scope", "Total built-up area across all floors"),
                ("Number of Floors", floors, "Floors", "Scope", f"G+{floors-1} ({floors} levels incl. ground)"),
                ("Construction Type", ctype, "", "Construction Method", const_std),
                ("Quality / Finish Level", quality, "", "Specification", "Basic / Standard / Premium — affects MEP and finishes rates"),
                ("Basement", basement, "", "Scope", "Adds basement excavation, waterproofing and structure"),
                ("Number of Blocks / Towers", blocks, "Blocks", "Construction Method", "Determines parallel work fronts"),
                ("Foundation Type", "RCC Isolated / Raft", "", "Substructure", "IS 456 Clause 34; soil report governs choice"),
                ("Structural System", f"RCC Frame ({ctype})", "", "Engineering Input", const_std),
                ("Cement Grade", "M25 (Foundation) / M30 (Frame)", "", "Specification", "IS 456 Table 1; IS 10262"),
                ("Steel Grade", "Fe500D TMT", "", "Specification", "IS 1786 — high ductility"),
                ("MEP Coverage", "Full MEP (Electrical + Plumbing + HVAC + Fire)", "", "Scope", "NBC 2016 Part 8; IS 732; IS 1641"),
            ]
        else:
            # Road — original rows
            lanes = self.p_in.lanes
            irc_spec = ("IRC:SP:73 (2-Lane)" if lanes <= 2 else
                       ("IRC:SP:84 (4-Lane Divided)" if lanes == 4 else
                       ("IRC:SP:87 (6-Lane Divided)" if lanes == 6 else f"IRC Standards ({lanes}-Lane)")))
            median_desc = "Single Carriageway (Undivided)" if lanes <= 2 else "IRC Raised / Depressed Median (4.5m)"
            subgrade_w = self.p_in.carriageway_width_m + (4.0 if lanes >= 4 else 2.0)
            gsb_w = self.p_in.carriageway_width_m + (2.5 if lanes >= 4 else 1.0)
            wmm_w = self.p_in.carriageway_width_m + (1.5 if lanes >= 4 else 0.5)
            specific = [
                ("Terrain", self.p_in.terrain.title(), "", "Scope", "Plain by default; hilly catalog only when selected"),
                ("Project Length", self.p_in.length_km, "km", "Scope", "User Input"),
                ("Number of Lanes", lanes, "Lanes", "Scope", f"{lanes}-Lane Highway Standard ({irc_spec})"),
                ("Formation Width", self.p_in.formation_width_m, "m", "Engineering Input", f"{irc_spec} Formation Width ({self.p_in.formation_width_m:.1f}m)"),
                ("Carriageway Width", self.p_in.carriageway_width_m, "m", "Engineering Input", f"{lanes}-Lane Carriageway + Paved Shoulders ({self.p_in.carriageway_width_m:.1f}m)"),
                ("Median Width", 0.0 if lanes <= 2 else 4.50, "m", "Engineering Input", median_desc),
                ("Subgrade Width", subgrade_w, "m", "Engineering Input", f"{subgrade_w:.1f}m Compacted Subgrade Bed"),
                ("GSB Width", gsb_w, "m", "Engineering Input", f"{gsb_w:.1f}m Granular Sub-base Width"),
                ("WMM Width", wmm_w, "m", "Engineering Input", f"{wmm_w:.1f}m Wet Mix Macadam Width"),
                ("Subgrade Depth", 0.50, "m", "Engineering Input", "500mm Compacted Depth (CBR >= 8%)"),
                ("GSB Thickness", 0.20, "m", "Layer Thickness", "200mm Granular Sub-base (MoRTH 401)"),
                ("WMM Thickness", 0.25, "m", "Layer Thickness", "250mm Wet Mix Macadam (MoRTH 406)"),
                ("DBM Thickness", 0.075, "m", "Layer Thickness", "75mm Dense Bituminous Macadam (MoRTH 505)"),
                ("BC Thickness", 0.040, "m", "Layer Thickness", "40mm Bituminous Concrete (MoRTH 507)"),
            ]
        return common_head + specific + common_tail

    # ------------------------------------------------------------------

    def _build_sheet_01_project_summary(self, wb):
        ws = wb.create_sheet(title="01_Project_Summary")
        self._create_banner(ws, self._dashboard_title(), "Baseline vs. Constraint-Optimized Executive Summary", 6)
        ws.freeze_panes = "A4"

        # KPI Metrics Cards
        kpis = [
            ("Project Scope", self._scope_label()),
            ("Target Duration (Constraint)", f"{self.p_in.duration_days} Days"),
            ("Baseline Duration (Normal)", f"{self.res.baseline_duration_days} Days"),
            ("Optimized Duration", f"{self.res.calculated_duration_days} Days"),
            ("Duration Saved", f"{self.res.duration_saved_days} Days"),
            ("Maximum Budget (Constraint)", f"₹ {self.p_in.budget_crore:,.2f} Crore"),
            ("Baseline Cost", f"₹ {self.res.baseline_cost_crore:,.2f} Crore"),
            ("Optimized Total Cost", f"₹ {self.res.total_calculated_cost_crore:,.2f} Crore"),
            ("Acceleration Cost (Excl. Escalation)", f"₹ {self.res.additional_cost_crore:,.2f} Crore"),
            ("Net Total-Cost Change", f"₹ {self.res.total_calculated_cost_crore - self.res.baseline_cost_crore:+,.2f} Crore"),
            ("Parallel Work Fronts", f"{self.res.work_fronts} Fronts"),
        ]

        ws.cell(row=4, column=1, value="EXECUTIVE KEY PERFORMANCE INDICATORS (KPIs)").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells("A4:F4")

        # Grid of KPI Cards (2 rows of 5)
        for idx, (title, val) in enumerate(kpis):
            r = 5 + (idx // 3) * 3
            c = 1 + (idx % 3) * 2

            ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+1)
            ws.merge_cells(start_row=r+1, start_column=c, end_row=r+1, end_column=c+1)

            cell_t = ws.cell(row=r, column=c, value=title)
            cell_t.font = Font(name="Calibri", size=9, bold=True, color="555555")
            cell_t.fill = self.fill_card
            cell_t.alignment = Alignment(horizontal="center", vertical="center")
            cell_t.border = self.border_thin

            cell_v = ws.cell(row=r+1, column=c, value=val)
            cell_v.font = Font(name="Calibri", size=13, bold=True, color=self.c_navy)
            cell_v.fill = self.fill_card
            cell_v.alignment = Alignment(horizontal="center", vertical="center")
            cell_v.border = self.border_thin

            ws.row_dimensions[r].height = 18
            ws.row_dimensions[r+1].height = 28

        curr = 19
        ws.cell(row=curr, column=1, value="EXECUTIVE SUMMARY COMPARISON: BASELINE vs. OPTIMIZED").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        curr += 1

        base_finish = self.p_in.start_date + timedelta(days=self.res.baseline_duration_days)
        opt_finish = self.p_in.start_date + timedelta(days=self.res.calculated_duration_days)
        dur_pass = self.res.calculated_duration_days <= self.p_in.duration_days
        cost_pass = self.res.total_calculated_cost_crore <= self.p_in.budget_crore
        final_pass = dur_pass and cost_pass

        # Section 1: BASELINE PLAN
        ws.cell(row=curr, column=1, value="SECTION 1: BASELINE PLAN (NORMAL RESOURCES & PRODUCTIVITY)").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        curr += 1

        base_rows = [
            ("Baseline Total Duration", f"{self.res.baseline_duration_days} Days"),
            ("Baseline Total Cost", f"₹ {self.res.baseline_cost_crore:,.2f} Crore"),
            ("Baseline Start Date", str(self.p_in.start_date)),
            ("Baseline Finish Date", str(base_finish)),
            ("Baseline Work Fronts", "1 Front (Sequential Gangs)"),
            ("Baseline Resource Utilization", "Standard normative equipment & gang sizes without crashing additions"),
        ]

        for label, val in base_rows:
            c1 = ws.cell(row=curr, column=1, value=label)
            c1.font = self.font_bold
            c1.fill = self.fill_card
            c1.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=2)

            c2 = ws.cell(row=curr, column=3, value=val)
            c2.font = self.font_regular
            c2.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=6)
            ws.row_dimensions[curr].height = 20
            curr += 1

        curr += 1
        # Section 2: OPTIMIZED / CONSTRAINT-BASED PLAN
        ws.cell(row=curr, column=1, value="SECTION 2: OPTIMIZED / CONSTRAINT-BASED PLAN").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        curr += 1

        opt_rows = [
            ("Required Duration", f"{self.p_in.duration_days} Days"),
            ("Optimized Duration", f"{self.res.calculated_duration_days} Days"),
            ("Maximum Budget", f"₹ {self.p_in.budget_crore:,.2f} Crore"),
            ("Optimized Cost", f"₹ {self.res.total_calculated_cost_crore:,.2f} Crore"),
            ("Optimized Finish Date", str(opt_finish)),
            ("Duration Constraint Status", "PASS (Met)" if dur_pass else "FAIL (Exceeded)"),
            ("Budget Constraint Status", "PASS (Met)" if cost_pass else "FAIL (Exceeded)"),
            ("Final Schedule Status", "FEASIBLE" if final_pass else "NOT FEASIBLE"),
        ]

        for label, val in opt_rows:
            c1 = ws.cell(row=curr, column=1, value=label)
            c1.font = self.font_bold
            c1.fill = self.fill_zebra
            c1.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=2)

            c2 = ws.cell(row=curr, column=3, value=val)
            c2.font = self.font_regular
            c2.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=6)

            if "PASS" in val or "FEASIBLE" == val:
                c2.fill = self.fill_pass
                c2.font = self.font_pass
            elif "FAIL" in val or "NOT FEASIBLE" == val:
                c2.fill = self.fill_crit
                c2.font = self.font_critical

            ws.row_dimensions[curr].height = 20
            curr += 1

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 2: PROJECT INPUTS
    # -------------------------------------------------------------
    def _build_sheet_02_project_inputs(self, wb):
        ws = wb.create_sheet(title="02_Project_Inputs")
        self._create_banner(ws, "PROJECT INPUTS & CONSTRAINT SPECIFICATIONS", "User constraints and standard engineering geometry inputs", 5)

        headers = ["Parameter", "Value", "Unit", "Category", "Constraint / Standard Status"]
        self._apply_header(ws, 4, headers)

        inputs_list = self._project_type_inputs_list()

        for i, (param, val, unit, cat, status) in enumerate(inputs_list, 5):
            ws.cell(row=i, column=1, value=param).border = self.border_thin
            ws.cell(row=i, column=2, value=val).border = self.border_thin
            ws.cell(row=i, column=3, value=unit).border = self.border_thin
            ws.cell(row=i, column=4, value=cat).border = self.border_thin
            
            c5 = ws.cell(row=i, column=5, value=status)
            c5.border = self.border_thin
            if "CONSTRAINT" in status:
                c5.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                c5.font = Font(name="Calibri", size=10, bold=True, color="856404")
            elif i % 2 == 0:
                for c in range(1, 6):
                    ws.cell(row=i, column=c).fill = self.fill_zebra

            ws.row_dimensions[i].height = 20

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 3: ASSUMPTIONS
    # -------------------------------------------------------------
    def _build_sheet_03_assumptions(self, wb):
        ws = wb.create_sheet(title="03_Assumptions")
        self._create_banner(
            ws,
            "NORMATIVE BASELINE ASSUMPTIONS & UNIT RATES",
            self._rate_basis_subtitle(),
            11
        )

        headers = [
            "Activity ID", "Activity Name", "Productivity Rate", "Baseline Crews",
            "Equipment & Resource Set",
            f"Direct Unit Rate (₹)",      # col 6 — source rate, never modified
            "Planning Variation (%)",     # col 7 — derived from planned vs direct rate
            "Planned Unit Rate (₹)",      # col 8
            "Overhead % (embedded)", "Profit % (embedded)",
            self._spec_ref_label()
        ]
        self._apply_header(ws, 4, headers)

        # Styling for the variation and planned rate columns
        amber_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        amber_font = Font(name="Calibri", size=10, bold=True, color="856404")
        blue_fill  = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
        blue_font  = Font(name="Calibri", size=10, bold=True, color=self.c_navy)

        base_map = {a.activity_id: a for a in self.base_acts}
        # A planned rate built above the source rate carries a known on-cost
        # split (statutory OH/CP for pipeline, SDB for road). When the planned
        # rate equals the source rate the source is used directly and its OH/CP
        # split is unknown, so the columns say "In source rate" instead of
        # implying a split that was never applied.

        for i, opt_act in enumerate(self.opt_acts, 5):
            b_act   = base_map.get(opt_act.activity_id, opt_act)
            c       = opt_act.cost
            morth_r = c.morth_direct_unit_rate
            var_pct = c.planning_variation_pct
            planned = c.planned_unit_rate

            row_vals = [
                opt_act.activity_id,
                opt_act.name,
                self._format_productivity_str(opt_act.productivity_rate, opt_act.productivity_unit),
                b_act.crews,
                opt_act.equipment_description,
                morth_r,    # 6 — Direct Unit Rate
                var_pct,    # 7 — Planning Variation %
                planned,    # 8 — Planned Unit Rate
                f"{c.overhead_pct:.1f}%" if planned > morth_r + 1e-9 else "In source rate",
                f"{c.profit_pct:.1f}%" if planned > morth_r + 1e-9 else "In source rate",
                opt_act.reference.morth_spec,
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font   = self.font_regular

                if col_idx == 6:
                    cell.number_format = "₹#,##0.00"
                elif col_idx == 7:
                    cell.number_format = '0.00"%"'
                    cell.fill = amber_fill
                    cell.font = amber_font
                elif col_idx == 8:
                    cell.number_format = "₹#,##0.00"
                    cell.fill = blue_fill
                    cell.font = blue_font
                elif col_idx in (1, 3, 4, 9, 10):
                    cell.alignment = Alignment(horizontal="center")

                if i % 2 == 0 and col_idx not in (7, 8):
                    cell.fill = self.fill_zebra

            ws.row_dimensions[i].height = 20

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 4: WBS HIERARCHY (ACTIVE WORK BREAKDOWN STRUCTURE ENGINE)
    # -------------------------------------------------------------
    def _build_sheet_04_wbs_hierarchy(self, wb):
        ws = wb.create_sheet(title="04_WBS_Hierarchy")
        self._create_banner(
            ws,
            "WORK BREAKDOWN STRUCTURE (WBS) HIERARCHY & ROLL-UP ANALYSIS",
            "3-Level Hierarchical EPC Scope of Works: Project -> Work Packages -> Activities with dynamic roll-ups",
            13
        )

        headers = [
            "WBS Code", "Level", "Work Package / Activity Name", "Activity ID",
            "Scope Description", "Quantity", "Unit",
            "Baseline Cost (₹ Cr)", "Optimized Cost (₹ Cr)", "Duration (Days)",
            "Start - Finish (Days)", "Weight %", "Critical?"
        ]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "A5"

        wbs_nodes = WBSGenerator.generate_wbs_tree(
            project_title=self.p_in.title,
            activities=self.opt_acts,
            baseline_activities=self.base_acts
        )

        pkg_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        pkg_font = Font(name="Calibri", size=10, bold=True, color="1B365D")

        for i, node in enumerate(wbs_nodes, 5):
            if node.level == 1:
                display_name = node.name.upper()
                level_str = "Level 1: Project Root"
            elif node.level == 2:
                display_name = f"  {node.name}"
                level_str = "Level 2: Work Package"
            else:
                display_name = f"    {node.name}"
                level_str = "Level 3: Activity"

            timeline_str = f"Day {node.start_day} - {node.finish_day}"
            crit_str = "CRITICAL" if node.is_critical else "Non-Critical"

            row_vals = [
                node.wbs_code,
                level_str,
                display_name,
                node.activity_id,
                node.description,
                node.quantity if node.quantity is not None else "",
                node.unit if node.unit else "",
                node.baseline_cost_crore,
                node.optimized_cost_crore,
                node.duration_days,
                timeline_str,
                node.weight_pct / 100.0,
                crit_str
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin

                if node.level == 1:
                    cell.fill = self.fill_primary
                    cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                    if col_idx in (8, 9):
                        cell.number_format = '"₹"#,##0.00" Cr"'
                    elif col_idx == 12:
                        cell.number_format = "0.00%"
                elif node.level == 2:
                    cell.fill = pkg_fill
                    cell.font = pkg_font
                    if col_idx in (8, 9):
                        cell.number_format = '"₹"#,##0.00" Cr"'
                    elif col_idx == 12:
                        cell.number_format = "0.00%"
                else:
                    cell.font = self.font_regular
                    if node.is_critical:
                        cell.fill = self.fill_crit
                        if col_idx == 13:
                            cell.font = self.font_critical
                    elif i % 2 == 0:
                        cell.fill = self.fill_zebra

                    if col_idx == 6:
                        cell.number_format = "#,##0.00"
                    elif col_idx in (8, 9):
                        cell.number_format = '"₹"#,##0.00" Cr"'
                    elif col_idx == 12:
                        cell.number_format = "0.00%"

                if col_idx in (1, 2, 4, 7, 10, 11, 13):
                    cell.alignment = Alignment(horizontal="center")

            ws.row_dimensions[i].height = 22 if node.level <= 2 else 20

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 5: ACTIVITIES (BASELINE & OPTIMIZED COLUMNS WITH WBS)
    # -------------------------------------------------------------
    def _build_sheet_05_activities(self, wb):
        ws = wb.create_sheet(title="05_Activities")
        self._create_banner(
            ws,
            self._activities_banner(),
            "WBS-aligned Baseline vs Optimized crews, durations, costs, days saved, and additional costs",
            16
        )
        
        headers = [
            "WBS Code", "Work Package", "Activity ID", "Activity Name", "Quantity", "Unit", "Productivity",
            "Baseline Crews", "Baseline Duration (d)", "Baseline Cost (₹)", "Predecessors",
            "Optimized Crews", "Optimized Duration (d)", "Optimized Cost (₹)", "Duration Saved (d)", "Additional Cost (₹)"
        ]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "A5"

        base_map = {a.activity_id: a for a in self.base_acts}

        start_r = 5
        for i, opt_act in enumerate(self.opt_acts, start_r):
            b_act = base_map.get(opt_act.activity_id, opt_act)
            dur_saved_formula = f"=I{i}-M{i}"
            add_cost_formula = f"=N{i}-J{i}"
            preds = ", ".join(opt_act.predecessors) if opt_act.predecessors else "-"

            row_vals = [
                opt_act.wbs_code,
                opt_act.work_package,
                opt_act.activity_id,
                opt_act.name,
                opt_act.quantity,
                opt_act.unit,
                self._format_productivity_str(opt_act.productivity_rate, opt_act.productivity_unit),
                b_act.crews,
                b_act.duration_days,
                b_act.cost.total_amount,
                preds,
                opt_act.crews,
                opt_act.duration_days,
                opt_act.cost.total_amount,
                dur_saved_formula,
                add_cost_formula
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                
                if col_idx in (5,):
                    cell.number_format = "#,##0.00"
                elif col_idx in (1, 3, 6, 8, 9, 12, 13, 15):
                    cell.alignment = Alignment(horizontal="center")
                elif col_idx in (10, 14, 16):
                    cell.number_format = "₹#,##0.00"

                if i % 2 == 0:
                    cell.fill = self.fill_zebra

            ws.row_dimensions[i].height = 20

        end_r = start_r + len(self.opt_acts) - 1
        tot_r = end_r + 1

        ws.cell(row=tot_r, column=4, value="TOTAL COST & CRASHING SUMMARY").font = self.font_bold
        ws.cell(row=tot_r, column=10, value=f"=SUM(J{start_r}:J{end_r})").number_format = "₹#,##0.00"
        ws.cell(row=tot_r, column=14, value=f"=SUM(N{start_r}:N{end_r})").number_format = "₹#,##0.00"
        ws.cell(row=tot_r, column=16, value=f"=SUM(P{start_r}:P{end_r})").number_format = "₹#,##0.00"

        for c_idx in range(1, 17):
            cell = ws.cell(row=tot_r, column=c_idx)
            cell.font = self.font_bold
            cell.border = self.border_double
            cell.fill = self.fill_zebra

        ws.row_dimensions[tot_r].height = 24
        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 6: BASELINE CPM (ONLY BASELINE SCHEDULE)
    # -------------------------------------------------------------
    def _build_sheet_06_baseline_cpm(self, wb):
        ws = wb.create_sheet(title="06_Baseline_CPM")
        self._create_banner(ws, "BASELINE CRITICAL PATH METHOD (CPM) SCHEDULE", "Normal execution using default resources and standard productivity (Unconstrained)", 12)
        
        base_finish = self.p_in.start_date + timedelta(days=self.res.baseline_duration_days)

        # Baseline Summary Box
        ws.cell(row=4, column=1, value="BASELINE SCHEDULE SUMMARY").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells("A4:L4")

        summary_rows = [
            ("Baseline Project Duration", f"{self.res.baseline_duration_days} Days"),
            ("Baseline Start Date", str(self.p_in.start_date)),
            ("Baseline Finish Date", str(base_finish)),
            ("Baseline Critical Path", " -> ".join(self.res.baseline_critical_path)),
            ("Baseline Resource Strategy", "1 Work Front, Standard Gangs, Sequential Finish-to-Start"),
        ]

        curr = 5
        for label, val in summary_rows:
            c1 = ws.cell(row=curr, column=1, value=label)
            c1.font = self.font_bold
            c1.fill = self.fill_card
            c1.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=2)

            c2 = ws.cell(row=curr, column=3, value=val)
            c2.font = self.font_regular
            c2.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=12)
            ws.row_dimensions[curr].height = 20
            curr += 1

        curr += 1
        headers = [
            "Activity ID", "Activity Description", "Predecessors", "Duration (d)",
            "Early Start (ES)", "Early Finish (EF)", "Late Start (LS)", "Late Finish (LF)",
            "Total Float (TF)", "Free Float (FF)", "Planned Start Date", "Planned Finish Date", "Critical?"
        ]
        self._apply_header(ws, curr, headers)
        ws.freeze_panes = f"A{curr+1}"
        curr += 1

        for i, act in enumerate(self.base_acts, curr):
            preds = ", ".join(act.predecessors) if act.predecessors else "-"
            crit_str = "YES" if act.is_critical else "NO"

            row_vals = [
                act.activity_id,
                act.name,
                preds,
                act.duration_days,
                act.es,
                act.ef,
                act.ls,
                act.lf,
                act.total_float,
                act.free_float,
                str(act.planned_start),
                str(act.planned_finish),
                crit_str
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (4, 5, 6, 7, 8, 9, 10, 11, 12, 13):
                    cell.alignment = Alignment(horizontal="center")
                if act.is_critical:
                    cell.fill = self.fill_crit
                    if col_idx == 13:
                        cell.font = self.font_critical
                elif i % 2 == 0:
                    cell.fill = self.fill_zebra

            ws.row_dimensions[i].height = 20

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 7: BASELINE COST (WITH WBS CODE & WORK PACKAGE)
    # -------------------------------------------------------------
    def _build_sheet_07_baseline_cost(self, wb):
        ws = wb.create_sheet(title="07_Baseline_Cost")
        self._create_banner(ws, "BASELINE QUANTITY & COST ESTIMATION SHEET", self._baseline_cost_subtitle(), 16)
        headers = [
            "WBS Code", "Work Package", "Activity ID", "Activity Description", "Quantity", "Unit",
            f"{self._source_rate_label()} (₹)",
            "Planned Unit Rate (₹)",
            "Base Direct Cost (₹)",
            "Execution Wastage Cost (₹)",
            "Planned Direct Cost (₹)",
            "Price Escalation (₹)",
            "OH Embedded in Rate (₹)",
            "CP Embedded in Rate (₹)",
            "Total Activity Cost (₹)",
            "Weight %"
        ]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "A5"

        blue_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
        blue_font = Font(name="Calibri", size=10, bold=True, color=self.c_navy)

        start_row = 5
        for i, act in enumerate(self.base_acts, start_row):
            c = act.cost
            base_direct = round(act.quantity * c.direct_unit_rate, 2)
            row_vals = [
                act.wbs_code,                 # 1
                act.work_package,             # 2
                act.activity_id,              # 3
                act.name,                     # 4
                act.quantity,                 # 5
                act.unit,                     # 6
                c.morth_direct_unit_rate,     # 7  MoRTH source rate
                c.direct_unit_rate,           # 8  planned (resource-derived) rate
                base_direct,                  # 9  Quantity x Planned Rate
                c.wastage_cost,               # 10 Execution Wastage Cost
                c.direct_cost,                # 11 Planned Direct Cost (Base + Wastage)
                c.escalation_amount,          # 12 Price Escalation
                c.overhead_embedded,          # 13 OH embedded in planned rate
                c.profit_embedded,            # 14 CP embedded in planned rate
                c.total_amount,               # 15 Total Activity Cost (Direct + Escalation)
                act.weight_pct / 100.0,       # 16 Weight %
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font   = self.font_regular

                if col_idx == 5:
                    cell.number_format = "#,##0.00"
                elif col_idx in (1, 3, 6):
                    cell.alignment = Alignment(horizontal="center")
                elif col_idx in (7, 8, 9, 10, 11, 12, 13, 14, 15):
                    cell.number_format = "₹#,##0.00"
                elif col_idx == 16:
                    cell.number_format = "0.00%"

                if col_idx == 8:
                    cell.fill = blue_fill
                    cell.font = blue_font
                elif i % 2 == 0:
                    cell.fill = self.fill_zebra

            ws.row_dimensions[i].height = 20

        end_row = start_row + len(self.base_acts) - 1
        tot_row  = end_row + 1

        ws.cell(row=tot_row, column=4, value="TOTAL PROJECT COST (FROM PLANNED RATES)").font = self.font_bold
        for col_letter, col_idx in [("I", 9), ("J", 10), ("K", 11), ("L", 12), ("M", 13), ("N", 14), ("O", 15)]:
            cell = ws.cell(row=tot_row, column=col_idx,
                           value=f"=SUM({col_letter}{start_row}:{col_letter}{end_row})")
            cell.number_format = "₹#,##0.00"
        ws.cell(row=tot_row, column=16,
                value=f"=SUM(P{start_row}:P{end_row})").number_format = "0.00%"

        for c_idx in range(1, 17):
            cell = ws.cell(row=tot_row, column=c_idx)
            cell.font   = self.font_bold
            cell.border = self.border_double
            cell.fill   = self.fill_zebra
        ws.row_dimensions[tot_row].height = 24

        curr = tot_row + 2
        ws.cell(row=curr, column=4, value="Total Planning Cost (Planned Rates):").font = self.font_bold
        ws.cell(row=curr, column=7, value=f"₹ {self.res.baseline_cost_crore:,.2f} Crore").font = self.font_bold
        curr += 1
        ws.cell(row=curr, column=4, value="Maximum Budget (Constraint):").font = self.font_bold
        ws.cell(row=curr, column=7,
                value=f"₹ {self.p_in.budget_crore:,.2f} Crore").font = self.font_bold
        curr += 1
        ws.cell(row=curr, column=4, value="Note:").font = self.font_small
        embedded_wastage = any(
            getattr(a.cost, "wastage_in_base_rate", False) for a in self.base_acts
        )
        wastage_note = (
            "Execution wastage is EMBEDDED in the planned rate for this "
            "project type, so col 10 is zero by design and is NOT added again."
            if embedded_wastage else
            "Execution Wastage Cost (col 10) captures site handling, compaction & "
            "idle losses. Planned Direct Cost (col 11) = Base Direct Cost + Wastage."
        )
        ws.cell(row=curr, column=5,
                value="All costs use the ALL-IN Planned Unit Rate (col 8). For "
                      "pipeline the planned rate is the published direct rate plus "
                      "material wastage, resource idle losses and statutory "
                      "overhead + contractor's profit (OH/CP are embedded and shown "
                      "as memo amounts in cols 13-14). For road it is the "
                      "first-principles build-up x the SDB on-cost multiplier "
                      "(road x1.188 / structures x1.32). "
                      "Base Direct Cost (col 9) = Quantity x Planned Unit Rate. "
                      + wastage_note + " "
                      "Price Escalation (col 12) reflects inflation up to the "
                      "mid-spend date. Total Activity Cost (col 15) = Planned "
                      "Direct Cost + Price Escalation.").font = self.font_small

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 8: CONSTRAINT OPTIMIZATION
    # -------------------------------------------------------------
    def _build_sheet_08_constraint_optimization(self, wb):
        ws = wb.create_sheet(title="08_Constraint_Optimization")
        self._create_banner(ws, "CONSTRAINT-BASED RESOURCE OPTIMIZATION LOG", "Iterative schedule crashing & parallel work front optimization to meet user constraints", 8)
        
        ws.freeze_panes = "A8"
        ws.views.sheetView[0].showGridLines = True

        # Target and Baseline summary cards
        ws.cell(row=4, column=1, value="TARGET CONSTRAINTS vs. BASELINE STARTING POINT").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells("A4:H4")

        summary_kpis = [
            ("Required Target Duration (Constraint)", f"{self.p_in.duration_days} Days", "Maximum Project Budget (Constraint)", f"₹ {self.p_in.budget_crore:,.2f} Crore"),
            ("Baseline Duration (Starting Point)", f"{self.res.baseline_duration_days} Days", "Baseline Total Cost (Starting Point)", f"₹ {self.res.baseline_cost_crore:,.2f} Crore"),
        ]

        curr = 5
        for l1, v1, l2, v2 in summary_kpis:
            c1 = ws.cell(row=curr, column=1, value=l1)
            c1.font = self.font_bold
            c1.fill = self.fill_card
            c1.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=2)

            c2 = ws.cell(row=curr, column=3, value=v1)
            c2.font = Font(bold=True, color=self.c_navy)
            c2.border = self.border_thin

            c3 = ws.cell(row=curr, column=4, value=l2)
            c3.font = self.font_bold
            c3.fill = self.fill_card
            c3.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=4, end_row=curr, end_column=5)

            c4 = ws.cell(row=curr, column=6, value=v2)
            c4.font = Font(bold=True, color=self.c_navy)
            c4.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=6, end_row=curr, end_column=8)

            ws.row_dimensions[curr].height = 20
            curr += 1

        curr += 1
        ws.cell(row=curr, column=1, value="OPTIMIZATION PROCESS ITERATION LOG").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=8)
        curr += 1

        headers = [
            "Iteration #", "Resource Change / Strategy Deployed", "Work Fronts",
            "Project Duration", "Project Cost (₹ Cr)", "Duration Constraint", "Budget Constraint", "Status"
        ]
        self._apply_header(ws, curr, headers)
        curr += 1

        opt_log = self.res.optimization_iterations
        for item in opt_log:
            dur_val = f"{item['duration_days']} days"
            cost_val = f"₹ {item['cost_crore']:,.2f} Cr"
            dur_status = item.get("duration_ok", "FAIL")
            cost_status = item.get("cost_ok", "PASS")
            status_text = item.get("status", "FAIL")

            row_vals = [
                f"Iteration {item['iteration']}",
                item["description"],
                f"{item['work_fronts']} Front{'s' if item['work_fronts'] > 1 else ''}",
                dur_val,
                cost_val,
                dur_status,
                cost_status,
                status_text
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=curr, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (1, 3, 4, 5, 6, 7, 8):
                    cell.alignment = Alignment(horizontal="center")

                if col_idx == 6:
                    if val == "PASS":
                        cell.fill = self.fill_pass
                        cell.font = self.font_pass
                    else:
                        cell.fill = self.fill_crit
                        cell.font = self.font_critical

                if col_idx == 7:
                    if val == "PASS":
                        cell.fill = self.fill_pass
                        cell.font = self.font_pass
                    else:
                        cell.fill = self.fill_crit
                        cell.font = self.font_critical

                if col_idx == 8:
                    if "PASS" in val or "OPTIMAL" in val:
                        cell.fill = self.fill_pass
                        cell.font = Font(bold=True, color=self.c_pass_fg)
                    else:
                        cell.fill = self.fill_crit
                        cell.font = self.font_critical

            ws.row_dimensions[curr].height = 22
            curr += 1

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 9: OPTIMIZED CPM (FINAL CONSTRAINT-BASED SCHEDULE)
    # -------------------------------------------------------------
    def _build_sheet_09_optimized_cpm(self, wb):
        ws = wb.create_sheet(title="09_Optimized_CPM")
        self._create_banner(
            ws,
            "FINAL OPTIMIZED CRITICAL PATH METHOD (CPM) SCHEDULE",
            f"Final scheduled plan satisfying Project Duration (<= {self.p_in.duration_days} Days) & Budget (<= ₹ {self.p_in.budget_crore:,.2f} Cr)",
            14
        )

        ws.freeze_panes = "A8"
        ws.views.sheetView[0].showGridLines = True

        opt_finish = self.p_in.start_date + timedelta(days=self.res.calculated_duration_days)

        # Check feasibility
        if not self.res.is_feasible:
            ws.cell(row=4, column=1, value="NO FEASIBLE SCHEDULE FOUND WITH CURRENT RESOURCES, PRODUCTIVITY AND BUDGET.").font = Font(size=14, bold=True, color=self.c_crit_fg)
            ws.merge_cells("A4:N4")
            for c in range(1, 15):
                ws.cell(row=4, column=c).fill = self.fill_crit
        else:
            ws.cell(row=4, column=1, value=f"FEASIBLE OPTIMAL PLAN ACHIEVED: {self.res.calculated_duration_days} Days (Target: <= {self.p_in.duration_days} d) | Total Cost: ₹ {self.res.total_calculated_cost_crore:,.2f} Cr (Budget: <= ₹ {self.p_in.budget_crore:,.2f} Cr)").font = Font(size=11, bold=True, color=self.c_pass_fg)
            ws.merge_cells("A4:N4")
            for c in range(1, 15):
                ws.cell(row=4, column=c).fill = self.fill_pass

        duration_ok = self.res.calculated_duration_days <= self.p_in.duration_days
        cost_ok = self.res.total_calculated_cost_crore <= self.p_in.budget_crore
        summary_rows = [
            ("Optimized Project Duration",
             f"{self.res.calculated_duration_days} Days "
             f"({'Met' if duration_ok else 'Not met'} <= {self.p_in.duration_days} d)"),
            ("Optimized Scheduled Cost",
             f"₹ {self.res.total_calculated_cost_crore:,.2f} Crore "
             f"({'Met' if cost_ok else 'Not met'} <= ₹ {self.p_in.budget_crore:,.2f} Cr)"),
            ("Scheduled Finish Date", str(opt_finish)),
            ("Optimized Critical Path", " -> ".join(self.res.critical_path)),
        ]

        curr = 5
        for label, val in summary_rows:
            c1 = ws.cell(row=curr, column=1, value=label)
            c1.font = self.font_bold
            c1.fill = self.fill_card
            c1.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=2)

            c2 = ws.cell(row=curr, column=3, value=val)
            c2.font = self.font_regular
            c2.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=14)
            ws.row_dimensions[curr].height = 20
            curr += 1

        curr += 1
        headers = [
            "Activity ID", "Activity Description", "Optimized Crews", "Predecessors",
            "Optimized Duration (d)", "Early Start (ES)", "Early Finish (EF)", "Late Start (LS)", "Late Finish (LF)",
            "Total Float (TF)", "Free Float (FF)", "Planned Start Date", "Planned Finish Date", "Critical?"
        ]
        self._apply_header(ws, curr, headers)
        curr += 1

        for i, act in enumerate(self.opt_acts, curr):
            preds = ", ".join(act.predecessors) if act.predecessors else "-"
            crit_str = "CRITICAL" if act.is_critical else "Non-Critical"

            row_vals = [
                act.activity_id,
                act.name,
                f"{act.crews} gangs ({self.res.work_fronts} fronts)",
                preds,
                act.duration_days,
                act.es,
                act.ef,
                act.ls,
                act.lf,
                act.total_float,
                act.free_float,
                str(act.planned_start),
                str(act.planned_finish),
                crit_str
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14):
                    cell.alignment = Alignment(horizontal="center")
                if act.is_critical:
                    cell.fill = self.fill_crit
                    if col_idx == 14:
                        cell.font = self.font_critical
                elif i % 2 == 0:
                    cell.fill = self.fill_zebra

            ws.row_dimensions[i].height = 20

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 10: OPTIMIZED COST (WITH WBS CODE & WORK PACKAGE)
    # -------------------------------------------------------------
    def _build_sheet_10_optimized_cost(self, wb):
        ws = wb.create_sheet(title="10_Optimized_Cost")
        self._create_banner(
            ws,
            "OPTIMIZED QUANTITY & COST ESTIMATION — ACCELERATION COST WATERFALL",
            "Time-driven, activity-dependent surcharge for extra gangs / extra work fronts. "
            "Quantity and the source rate are unchanged; only the DERIVED effective rate rises.",
            25
        )

        headers = [
            "Activity ID",                    # 1
            "Activity Description",           # 2
            "Quantity",                       # 3
            "Unit",                           # 4
            f"{self._source_rate_label()} (₹)",   # 5
            "Planning Variation (%)",         # 6
            "Planned Unit Rate (₹)",          # 7
            "Planned Direct Cost (₹)",        # 8
            "Fleet Mobilisation (₹)",         # 9
            "Shift / Overtime Premium (₹)",   # 10
            "Front Supervision (₹)",          # 11
            "Site Congestion (₹)",            # 12
            "Extra-Front Site Establishment (₹)",  # 13
            "Extra-Front Resident Staff (₹)",      # 14
            "Plant Hire Minimum Uplift (₹)",       # 15
            "Gang Interference Loss (₹)",          # 16
            "Total Crashing Surcharge (₹)",        # 17
            "Surcharge as % of Planned",           # 18
            "Optimized Direct Cost (₹)",           # 19
            "Effective Direct Unit Rate (₹)",      # 20
            "Overhead (₹) — embedded in rate",                # 21
            "Profit (₹) — embedded in rate",                  # 22
            "Total Activity Cost (₹)",                        # 23
            "All-In Unit Rate (₹)",                           # 24
            "Weight %",                            # 25
            "Material Wastage Cost (₹)",           # 26
            "Fuel Wastage Cost (₹)",               # 27
            "Plant Wastage Cost (₹)",              # 28
            "Labour Wastage Cost (₹)",             # 29
            "Total Wastage Cost (₹)",              # 30
            "Price Escalation (₹)",               # 31
        ]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "C5"

        green_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
        green_font = Font(name="Calibri", size=10, bold=True, color="065F46")
        amber_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        amber_font = Font(name="Calibri", size=10, bold=True, color="92400E")

        start_row = 5
        for i, act in enumerate(self.opt_acts, start_row):
            c = act.cost
            qty = max(act.quantity, 0.001)

            # Planned side: quantity x planned rate (unchanged by acceleration)
            planned_rate      = c.direct_unit_rate
            planned_cost      = round(act.quantity * planned_rate, 2)

            # The 8 acceleration components (cols 9-16)
            comp = [
                c.crash_fleet_mobilisation,
                c.crash_shift_premium,
                c.crash_front_supervision,
                c.crash_site_congestion,
                c.crash_site_establishment,
                c.crash_resident_staff,
                c.crash_plant_hire_uplift,
                c.crash_efficiency_loss,
            ]
            # Col 17 == sum(cols 9..16) exactly
            total_surcharge = round(sum(comp), 2)
            surcharge_pct = (total_surcharge / planned_cost * 100.0) if planned_cost > 0 else 0.0

            # Col 19 == col 8 + col 17
            optimized_direct_cost = round(
                planned_cost + total_surcharge + c.wastage_cost, 2
            )
            # Col 20 == col 19 / quantity. Left blank when the quantity is zero
            # (e.g. median dressing on a 2-lane road) because a rate over a
            # zero quantity is undefined - printing Rs 0.00 would imply a
            # real rate was derived.
            effective_rate = round(optimized_direct_cost / qty, 2) if act.quantity > 0 else None

            row_vals = [
                act.activity_id,                 # 1
                act.name,                        # 2
                act.quantity,                    # 3
                act.unit,                        # 4
                c.morth_direct_unit_rate,        # 5  MoRTH source — never modified
                c.planning_variation_pct,        # 6
                planned_rate,                    # 7  resource build-up (bare site cost)
                planned_cost,                    # 8  quantity x planned rate
                comp[0],                         # 9  Fleet mobilisation
                comp[1],                         # 10 Shift / overtime premium
                comp[2],                         # 11 Front supervision
                comp[3],                         # 12 Site congestion
                comp[4],                         # 13 Extra-front site establishment
                comp[5],                         # 14 Extra-front resident staff
                comp[6],                         # 15 Plant hire minimum uplift
                comp[7],                         # 16 Gang interference loss
                total_surcharge,                 # 17 Total crashing surcharge
                surcharge_pct,                   # 18 Surcharge as % of planned
                optimized_direct_cost,           # 19 Optimized direct cost
                effective_rate,                  # 20 Effective direct unit rate
                c.overhead_embedded,             # 21 OH embedded in planned rate
                c.profit_embedded,               # 22 CP embedded in planned rate
                c.total_amount,                  # 23
                c.final_unit_rate,               # 24 All-in unit rate
                act.weight_pct / 100.0,          # 25
                c.material_wastage_cost,         # 26 Bottom-up material wastage
                c.fuel_wastage_cost,             # 27 Fuel idle / standby burn
                c.plant_wastage_cost,            # 28 Plant idle / breakdown
                c.labour_wastage_cost,           # 29 Labour waiting on plant
                c.wastage_cost,                  # 30 Total execution loss
                c.escalation_amount,             # 31 Price escalation (base -> spend date)
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font   = self.font_regular

                if col_idx in (1, 4):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 3:
                    cell.number_format = "#,##0.00"
                elif col_idx in (5, 6):
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="center")
                elif col_idx == 7:
                    cell.number_format = "₹#,##0.00"
                elif col_idx == 8:
                    cell.number_format = "₹#,##0.00"
                elif col_idx in (9, 10, 11, 12, 13, 14, 15, 16):
                    cell.number_format = "₹#,##0.00"
                elif col_idx == 17:
                    cell.number_format = "₹#,##0.00"
                    cell.font = amber_font
                elif col_idx == 18:
                    cell.number_format = '0.00"%"'
                    cell.alignment = Alignment(horizontal="center")
                    if val > 0:
                        cell.font = amber_font
                elif col_idx == 19:
                    cell.number_format = "₹#,##0.00"
                    cell.font = Font(name="Calibri", size=10, bold=True)
                elif col_idx == 20:
                    cell.number_format = "₹#,##0.00"
                    cell.fill = green_fill
                    cell.font = green_font
                elif col_idx in (21, 22, 23, 24):
                    cell.number_format = "₹#,##0.00"
                elif col_idx == 25:
                    cell.number_format = "0.00%"
                elif col_idx in (26, 27, 28, 29):
                    cell.number_format = "₹#,##0.00"
                elif col_idx == 30:
                    cell.number_format = "₹#,##0.00"
                    cell.font = amber_font

                if i % 2 == 0 and col_idx not in (17, 20, 30):
                    cell.fill = self.fill_zebra
                if col_idx == 17 and i % 2 == 0:
                    cell.fill = amber_fill

            ws.row_dimensions[i].height = 20

        end_row = start_row + len(self.opt_acts) - 1
        tot_row  = end_row + 1

        ws.cell(row=tot_row, column=2, value="TOTAL OPTIMIZED PROJECT COST").font = self.font_bold

        # Summed money columns: 8..17, 19, 21..23, 26..31
        for col_idx in [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19, 21, 22, 23, 26, 27, 28, 29, 30, 31]:
            letter = get_column_letter(col_idx)
            ws.cell(row=tot_row, column=col_idx,
                    value=f"=SUM({letter}{start_row}:{letter}{end_row})").number_format = "₹#,##0.00"
        # Surcharge as % of planned (derived, not a sum)
        ws.cell(row=tot_row, column=18,
                value=f"=IF(H{tot_row}=0,0,Q{tot_row}/H{tot_row}*100)").number_format = '0.00"%"'
        ws.cell(row=tot_row, column=25,
                value=f"=SUM(Y{start_row}:Y{end_row})").number_format = "0.00%"
        # Cols 5, 6, 7 and 20/24 are rates — not additive across mixed units.

        for c_idx in range(1, 32):
            cell = ws.cell(row=tot_row, column=c_idx)
            cell.font   = self.font_bold
            cell.border = self.border_double
            cell.fill   = self.fill_zebra
        ws.cell(row=tot_row, column=18).number_format = '0.00"%"'
        ws.cell(row=tot_row, column=25).number_format = "0.00%"
        ws.row_dimensions[tot_row].height = 24

        # ---- summary block -------------------------------------------------
        total_surcharge_cr = sum(
            act.cost.crash_total_surcharge for act in self.opt_acts
        ) / 1e7
        total_wastage_cr = sum(
            act.cost.wastage_cost for act in self.opt_acts
        ) / 1e7

        curr = tot_row + 2
        esc_cr = sum(act.cost.escalation_amount for act in self.opt_acts) / 1e7
        esc_enabled = bool(getattr(self.res, "activities", None)) and any(
            act.cost.escalation_amount > 0 for act in self.opt_acts)
        summary = [
            ("Total Planned Direct Cost (Qty x Planned Rate):",
             f"₹ {sum(act.quantity * act.cost.direct_unit_rate for act in self.opt_acts) / 1e7:,.2f} Crore"),
            ("Total Acceleration / Crashing Surcharge:",
             f"₹ {total_surcharge_cr:,.2f} Crore"),
            ("Total Wastage Cost (Material + Fuel/Plant/Labour Idle):",
             f"₹ {total_wastage_cr:,.2f} Crore"),
            ("Total Price Escalation (base date to mid-point of spend):",
             (f"₹ {esc_cr:,.2f} Crore" if esc_enabled else "₹ 0.00 Crore (disabled)")),
            ("Total Optimized Cost (incl. Wastage & Escalation; OH & CP embedded in rate):",
             f"₹ {self.res.total_calculated_cost_crore:,.2f} Crore"),
            ("Acceleration Uplift over Baseline:",
             f"₹ {self.res.additional_cost_crore:,.2f} Crore "
             f"({(self.res.additional_cost_crore / self.res.baseline_cost_crore * 100.0) if self.res.baseline_cost_crore else 0.0:,.2f}% of baseline)"),
            ("Duration Saved:",
             f"{self.res.duration_saved_days} Days "
             f"({self.res.baseline_duration_days} -> {self.res.calculated_duration_days} Days)"),
            ("Parallel Work Fronts Deployed:", f"{self.res.work_fronts} Fronts"),
            ("Maximum Budget (Constraint):",
             f"₹ {self.p_in.budget_crore:,.2f} Crore"),
        ]
        for i, (label, val) in enumerate(summary):
            lbl_cell = ws.cell(row=curr, column=2, value=label)
            lbl_cell.font = self.font_bold
            ws.merge_cells(start_row=curr, start_column=2, end_row=curr, end_column=4)
            val_cell = ws.cell(row=curr, column=5, value=val)
            val_cell.font = self.font_regular
            ws.merge_cells(start_row=curr, start_column=5, end_row=curr, end_column=10)
            # Highlight the wastage row in amber to match wastage columns 26-30
            if "Wastage" in label:
                for c in range(2, 11):
                    ws.cell(row=curr, column=c).fill = PatternFill("solid", fgColor="FFF2CC")
                lbl_cell.font = Font(bold=True, color="7B5800")
                val_cell.font = Font(color="7B5800")
            curr += 1

        duration_ok = self.res.calculated_duration_days <= self.p_in.duration_days
        budget_ok = self.res.total_calculated_cost_crore <= self.p_in.budget_crore
        for label, ok, pass_text, fail_text in (
            ("Duration Compliance Status:",
             duration_ok,
             f"PASS (Within {self.p_in.duration_days} Days)",
             f"FAIL (Exceeded by {self.res.calculated_duration_days - self.p_in.duration_days} Days)"),
            ("Budget Compliance Status:",
             budget_ok,
             "PASS (Within Budget)",
             "FAIL (Exceeded)"),
        ):
            ws.cell(row=curr, column=2, value=label).font = self.font_bold
            ws.merge_cells(start_row=curr, start_column=2, end_row=curr, end_column=4)
            status_cell = ws.cell(row=curr, column=5,
                                  value=pass_text if ok else fail_text)
            status_cell.font = self.font_pass if ok else self.font_critical
            status_cell.fill = self.fill_pass if ok else self.fill_crit
            ws.merge_cells(start_row=curr, start_column=5, end_row=curr, end_column=10)
            curr += 1
        curr += 1

        # ---- explanatory footnote -----------------------------------------
        ptype = getattr(self.p_in, "project_type", "road") or "road"
        embedded_wastage = any(
            getattr(a.cost, "wastage_in_base_rate", False) for a in self.opt_acts
        )
        rate_basis_text = (
            "The planned rate is the bare source/resource rate plus the configured "
            "overhead and contractor-profit on-costs. Wastage is calculated from "
            "the configured material/resource allowances as a separate direct-cost "
            "line for road, pipeline and building. "
        )
        if embedded_wastage:
            wastage_text = (
                "Execution wastage is EMBEDDED in the planned rate for this "
                "project type, so columns 26-30 are ₹0 by design and wastage is "
                "not charged a second time. "
            )
        else:
            wastage_text = (
                "Column 26 is the MATERIAL WASTAGE COST priced from first "
                "principles: each activity's recipe is resolved into real "
                "quantities, the material's own wastage factor is applied, and the "
                "incremental issued quantity is priced. Columns 27-29 are fuel, "
                "plant and labour idle losses; column 30 is their sum and is added "
                "to direct cost. "
            )
        footnote = (
            "HOW TO READ THIS SHEET: Columns 5, 6, 7 and 8 are the ORIGINAL plan and "
            f"never change — col 5 is the {self._source_rate_label()} (untouched "
            "reference), col 6 is the PLANNING VARIATION between the planned rate "
            "and that source rate, col 7 is the resulting Planned Unit Rate, and "
            "col 8 = quantity × col 7 is the Planned Direct Cost. "
            + rate_basis_text
            + "READ THE SIGNS: a positive variation means the planned rate is above "
            "the source rate; a negative variation means the source rate is above "
            "the planned build-up. "
            "Columns 9-16 are the ACCELERATION SURCHARGE: the extra cost of buying "
            "the duration reduction, broken out by cause (fleet mobilisation, "
            "shift/overtime premium, per-extra-front supervision, site congestion, "
            "per-extra-front site establishment and resident staff charged for every "
            "day the front is open, plant-hire minimum-month billing uplift, and "
            "super-linear gang interference). Column 17 is the sum of columns 9-16, "
            "expressed as a percentage of the planned cost in column 18 — this uplift "
            "is NOT a uniform percentage: it is time-driven and activity-dependent, "
            "so short, structure-heavy activities carry a very different surcharge "
            "per rupee of work from long linear fronts. Column 19 (Optimized Direct "
            "Cost) = column 8 + column 17, so column 20 (Effective Direct Unit Rate) "
            "= column 19 / quantity is a DERIVED rate, not a new source rate: neither "
            "the quantity nor the source rate has been altered; the effective rate is "
            "higher because the work is being executed faster with more gangs, more "
            "fronts and longer open-site time. Overhead (21) and profit (22) are "
            "EMBEDDED in the planned rate and shown as memo amounts; they are not "
            "added a second time. "
            + wastage_text
            + "Column 23 (Total Activity Cost) = col 19 (Optimized Direct Cost) + "
            "col 30 (wastage) + col 31 (price escalation). Column 31 is escalation to "
            "the mid-point of spend; it is zero when escalation is disabled in the "
            "assumptions. Col 24 is the resulting All-In Unit Rate."
        )
        ws.cell(row=curr, column=1, value="NOTE ON THE 'EFFECTIVE DIRECT UNIT RATE'").font = Font(
            size=10, bold=True, color=self.c_navy
        )
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=12)
        curr += 1
        fn = ws.cell(row=curr, column=1, value=footnote)
        fn.font = self.font_small
        fn.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=25)
        ws.row_dimensions[curr].height = 120

        self._auto_fit_columns(ws)
        # The description / footnote columns need a sane readable width
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 46

    # -------------------------------------------------------------
    # SHEET 11: BASELINE VS OPTIMIZED
    # -------------------------------------------------------------
    def _build_sheet_11_baseline_vs_optimized(self, wb):
        ws = wb.create_sheet(title="11_Baseline_vs_Optimized")
        self._create_banner(ws, "BASELINE VS. OPTIMIZED SCHEDULE COMPARISON", "Executive side-by-side presentation table comparing normal vs. constraint-crushed schedule", 6)
        
        ws.freeze_panes = "A6"
        ws.views.sheetView[0].showGridLines = True

        base_finish = self.p_in.start_date + timedelta(days=self.res.baseline_duration_days)
        opt_finish = self.p_in.start_date + timedelta(days=self.res.calculated_duration_days)
        target_finish = self.p_in.start_date + timedelta(days=self.p_in.duration_days)
        duration_ok = self.res.calculated_duration_days <= self.p_in.duration_days
        cost_ok = self.res.total_calculated_cost_crore <= self.p_in.budget_crore

        ws.cell(row=4, column=1, value="EXECUTIVE COMPARISON MATRIX").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells("A4:F4")

        headers = ["Metric", "BASELINE", "OPTIMIZED", "VARIANCE / SAVINGS", "TARGET CONSTRAINT", "STATUS"]
        self._apply_header(ws, 5, headers)

        if duration_ok and cost_ok:
            constraint_variance = "Both Constraints Satisfied"
            constraint_status = "PASS"
        elif not duration_ok and not cost_ok:
            constraint_variance = "Duration & Budget Constraints Not Met"
            constraint_status = "FAIL"
        elif not duration_ok:
            constraint_variance = "Duration Constraint Not Met"
            constraint_status = "FAIL"
        else:
            constraint_variance = "Budget Constraint Not Met"
            constraint_status = "FAIL"

        comparison_matrix = [
            ("Duration", f"{self.res.baseline_duration_days} Days", f"{self.res.calculated_duration_days} Days", f"-{self.res.duration_saved_days} Days Saved", f"<= {self.p_in.duration_days} Days", "PASS (FEASIBLE)" if duration_ok else "FAIL"),
            ("Cost", f"₹ {self.res.baseline_cost_crore:,.2f} Cr", f"₹ {self.res.total_calculated_cost_crore:,.2f} Cr", f"+₹ {self.res.additional_cost_crore:,.2f} Cr Acceleration (Excl. Escalation)", f"<= ₹ {self.p_in.budget_crore:,.2f} Cr", "PASS (FEASIBLE)" if cost_ok else "FAIL"),
            ("Resources", f"Normal ({self.res.baseline_work_fronts} Front{'s' if self.res.baseline_work_fronts != 1 else ''}, Standard Gangs)", f"Optimized ({self.res.work_fronts} Fronts, Multi-Gangs)", f"+{self.res.work_fronts - self.res.baseline_work_fronts} Parallel Fronts Deployed", "Dynamic Allocation", "ACTIVE"),
            ("Finish Date", str(base_finish), str(opt_finish), f"{self.res.duration_saved_days} Days Earlier", f"<= {target_finish}", "ON TIME" if opt_finish <= target_finish else "LATE"),
            ("Constraint Status", "Unconstrained Baseline", "Constraint-Based Plan", constraint_variance, "Duration & Budget Limits", constraint_status),
        ]

        curr = 6
        for row in comparison_matrix:
            for col_idx, val in enumerate(row, 1):
                cell = ws.cell(row=curr, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (2, 3, 4, 5, 6):
                    cell.alignment = Alignment(horizontal="center")
                if col_idx == 6:
                    if "PASS" in val or "ACTIVE" in val or "ON TIME" in val:
                        cell.fill = self.fill_pass
                        cell.font = Font(bold=True, color=self.c_pass_fg)
                    else:
                        cell.fill = self.fill_crit
                        cell.font = self.font_critical

            ws.row_dimensions[curr].height = 22
            curr += 1

        curr += 2
        ws.cell(row=curr, column=1, value="RESOURCE CHANGES & GANG ALLOCATION BREAKDOWN").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        curr += 1

        detail_headers = ["Activity ID", "Activity Name", "Baseline Gangs", "Optimized Gangs", "Days Saved", "Resource Changes Deployed"]
        self._apply_header(ws, curr, detail_headers)
        curr += 1

        base_map = {a.activity_id: a for a in self.base_acts}
        any_resourced = False
        for opt_act in self.opt_acts:
            b_act = base_map.get(opt_act.activity_id, opt_act)
            days_saved = b_act.duration_days - opt_act.duration_days
            desc = opt_act.resources_added_desc
            if not desc or desc == "-":
                # No gang was added, but the calendar may still have moved (the
                # staggered handover re-times a chain without adding resources).
                # Say so, otherwise the column reads as empty and the reader
                # cannot tell "nothing changed" from "nothing was recorded".
                if days_saved > 0:
                    desc = (f"Re-sequenced: staggered FS handover "
                            f"(no gang change) - {days_saved:.0f} d from re-timing")
                elif days_saved < 0:
                    desc = (f"Re-timed by the handover (no gang change) - "
                            f"{-days_saved:.0f} d longer on the calendar")
                else:
                    desc = "No change"
            else:
                any_resourced = True
            row_vals = [
                opt_act.activity_id,
                opt_act.name,
                b_act.crews,
                opt_act.crews,
                days_saved,
                desc
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=curr, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx in (3, 4, 5):
                    cell.alignment = Alignment(horizontal="center")
                if col_idx == 5 and days_saved > 0:
                    cell.font = Font(bold=True, color=self.c_pass_fg)

            if curr % 2 == 0:
                for c in range(1, 7):
                    ws.cell(row=curr, column=c).fill = self.fill_zebra

            ws.row_dimensions[curr].height = 20
            curr += 1

        # -------------------------------------------------------------
        # ACCELERATION COST WATERFALL (baseline -> optimized)
        # Mirrors the per-activity surcharge breakdown on sheet 10 so the
        # acceleration premium is visible and auditable in the comparison.
        # -------------------------------------------------------------
        curr += 1
        ws.cell(row=curr, column=1, value="ACCELERATION COST WATERFALL — WHY THE OPTIMIZED PLAN COSTS MORE").font = Font(
            size=11, bold=True, color=self.c_navy
        )
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        curr += 1

        opt_acts = self.opt_acts
        planned_direct = sum(act.quantity * act.cost.direct_unit_rate for act in opt_acts)

        def _sum(attr: str) -> float:
            return sum(getattr(act.cost, attr, 0.0) or 0.0 for act in opt_acts)

        waterfall = [
            ("Fleet Mobilisation (extra gangs x equipment)",
             _sum("crash_fleet_mobilisation")),
            ("Shift / Overtime Premium (extra shifts on labour + plant)",
             _sum("crash_shift_premium")),
            ("Front Supervision (per parallel work front)",
             _sum("crash_front_supervision")),
            ("Site Congestion (site density per extra gang)",
             _sum("crash_site_congestion")),
            (f"Extra-Front Site Establishment (fronts x days @ "
             f"₹{getattr(self.res, 'accel_site_establishment_rate', 0):,.0f}/day)",
             _sum("crash_site_establishment")),
            (f"Extra-Front Resident Staff (fronts x days @ "
             f"₹{getattr(self.res, 'accel_resident_staff_rate', 0):,.0f}/day)",
             _sum("crash_resident_staff")),
            ("Plant Hire Minimum Uplift (calendar-month billing on extra gangs)",
             _sum("crash_plant_hire_uplift")),
            ("Gang Interference Loss (super-linear in gang count)",
             _sum("crash_efficiency_loss")),
            ]

        wf_headers = ["Acceleration Cost Component", "Amount (₹)", "Amount (₹ Crore)",
                      "% of Planned Direct Cost", "% of Total Surcharge", "Driver"]
        self._apply_header(ws, curr, wf_headers)
        wf_header_row = curr
        curr += 1

        total_wf = sum(v for _, v in waterfall)
        for label, amount in waterfall:
            share = (amount / total_wf * 100.0) if total_wf > 0 else 0.0
            of_planned = (amount / planned_direct * 100.0) if planned_direct > 0 else 0.0
            driver = "Crew-scaling" if label.startswith(("Fleet", "Site Cong", "Gang")) else (
                "Time-driven" if label.startswith(("Extra-Front", "Plant Hire")) else "Front / shift scaling"
            )
            row_vals = [
                label,
                round(amount, 2),
                round(amount / 1e7, 4),
                round(of_planned, 4),
                round(share, 2),
                driver,
            ]
            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=curr, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                if col_idx == 2:
                    cell.number_format = "₹#,##0.00"
                elif col_idx == 3:
                    cell.number_format = "₹#,##0.0000 Cr"
                elif col_idx in (4, 5):
                    cell.number_format = '0.00"%"'
                    cell.alignment = Alignment(horizontal="center")
                elif col_idx == 6:
                    cell.alignment = Alignment(horizontal="center")
            if curr % 2 == 0:
                for c_idx in range(1, 7):
                    ws.cell(row=curr, column=c_idx).fill = self.fill_zebra
            ws.row_dimensions[curr].height = 20
            curr += 1

        # waterfall subtotal
        ws.cell(row=curr, column=1, value="TOTAL ACCELERATION SURCHARGE").font = self.font_bold
        ws.cell(row=curr, column=2,
                value=f"=SUM(B{wf_header_row + 1}:B{curr - 1})").number_format = "₹#,##0.00"
        ws.cell(row=curr, column=3, value=f"=B{curr}/10000000").number_format = "₹#,##0.0000 Cr"
        ws.cell(row=curr, column=4,
                value=(f"=IF({planned_direct}=0,0,B{curr}/{planned_direct}*100)")).number_format = '0.00"%"'
        ws.cell(row=curr, column=5, value=100.0 if total_wf > 0 else 0.0).number_format = '0.00"%"'
        for c_idx in range(1, 7):
            cell = ws.cell(row=curr, column=c_idx)
            cell.font = self.font_bold
            cell.border = self.border_double
            cell.fill = self.fill_zebra
        subtotal_row = curr
        curr += 2

        # baseline -> optimized cost bridge
        # Identical by construction, because the planned cost is quantity x the
        # same all-in rate in both plans (no on-cost pass, nothing time-based in
        # it). So the ONLY things that can move the total between the two plans
        # are the acceleration surcharge and the execution wastage, and both
        # deltas are shown so the bridge always reconciles to the rupee:
        #     baseline + d(surcharge) + d(wastage) == optimized
        # Showing a single "incremental compression cost" number instead would
        # not add up whenever re-timing moved wastage in the cheaper direction
        # (which it can: fewer gang-days = less idle-plant loss).
        base_sur = sum(a.cost.crash_total_surcharge for a in self.base_acts)
        opt_sur = sum(a.cost.crash_total_surcharge for a in self.opt_acts)
        base_wst = sum(a.cost.wastage_cost for a in self.base_acts)
        opt_wst = sum(a.cost.wastage_cost for a in self.opt_acts)
        base_esc = sum(a.cost.escalation_amount for a in self.base_acts)
        opt_esc = sum(a.cost.escalation_amount for a in self.opt_acts)
        embedded_wastage = any(
            getattr(a.cost, "wastage_in_base_rate", False)
            for a in self.opt_acts
        )
        wastage_bridge_label = (
            "(+/-) Change in execution wastage (embedded in planned rate - no separate charge)"
            if embedded_wastage else
            "(+/-) Change in execution wastage (gang-days moved by re-timing)"
        )
        bridge = [
            ("(=) BASELINE TOTAL COST", self.res.baseline_cost_crore),
            ("(+/-) Change in acceleration surcharge (extra gangs / re-sequencing)",
             (opt_sur - base_sur) / 1e7),
            (wastage_bridge_label, (opt_wst - base_wst) / 1e7),
            ("(+/-) Change in price escalation (money spent earlier/later)",
             (opt_esc - base_esc) / 1e7),
            ("(=) OPTIMIZED TOTAL COST", self.res.total_calculated_cost_crore),
        ]
        ws.cell(row=curr, column=1, value=(
            "COST BRIDGE (₹ Crore) — on-costs (OH+CP) are embedded in the "
            "planned unit rate")).font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        curr += 1
        bridge_headers = ["Cost Component", "₹ Crore", "Comment"]
        self._apply_header(ws, curr, bridge_headers)
        curr += 1
        for label, val in bridge:
            ws.cell(row=curr, column=1, value=label).font = self.font_regular
            ws.cell(row=curr, column=1).border = self.border_thin
            vc = ws.cell(row=curr, column=2, value=round(val, 4))
            vc.number_format = "₹#,##0.00"
            vc.border = self.border_thin
            cc = ws.cell(row=curr, column=3, value=(
                "Total surcharge (sheet 10, cols 9-16)" if "surcharge" in label.lower()
                else ("₹0 by design: wastage is embedded in the planned "
                      "rate, so re-timing cannot add a separate charge"
                      if (embedded_wastage and "wastage" in label.lower())
                      else ("Execution wastage delta (sheet 10, cols 26-30)"
                            if "wastage" in label.lower()
                            else ("Escalation delta (sheet 10, col 31; ₹0 when disabled)"
                                  if "escalation" in label.lower() else "")))
            ))
            cc.font = self.font_small
            cc.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=6)
            if label.startswith("(=)"):
                for c_idx in range(1, 7):
                    ws.cell(row=curr, column=c_idx).font = self.font_bold
            ws.row_dimensions[curr].height = 20
            curr += 1

        curr += 1
        note = (
            "The optimized plan costs more than the baseline because the schedule is bought with time-driven "
            f"acceleration premium. Quantity and the {self._source_rate_label().lower()} are identical in both "
            "cases; the extra cost is entirely the surcharge above, which scales with the number of gangs, the "
            "number of parallel work fronts, and — via per-front establishment, resident staff and calendar-month "
            "plant billing — with how many days each front is kept open. "
            + ("Execution wastage is embedded in the planned rate, so its bridge delta is ₹0 by design. "
               if embedded_wastage else "")
            + "Run with --no-accel-cost to zero the time-driven terms and see the legacy behaviour on identical input."
        )
        nc = ws.cell(row=curr, column=1, value=note)
        nc.font = self.font_small
        nc.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=6)
        ws.row_dimensions[curr].height = 70
        _ = subtotal_row  # retained for clarity / future formatting hooks

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 12: MONTHLY SCHEDULE (OPTIMIZED)
    # -------------------------------------------------------------
    def _build_sheet_12_monthly_schedule(self, wb):
        ws = wb.create_sheet(title="12_Monthly_Schedule")
        num_months = len(self.res.monthly_progress)
        self._create_banner(ws, "PERIODIC / MONTHLY PROGRESS SCHEDULE (OPTIMIZED)", f"Monthly planned progress matrix across the {self.res.calculated_duration_days}-day scheduled timeline", 4 + num_months)
        
        headers = ["Activity ID", "Activity Description", "Duration", "Start - Finish"] + [f"M{m.month_index}" for m in self.res.monthly_progress]
        self._apply_header(ws, 4, headers)
        ws.freeze_panes = "E5"

        from .progress_calculator import ProgressCalculator
        calc = ProgressCalculator(self.opt_acts, self.p_in, self.res.calculated_duration_days)
        _, matrix = calc.generate_monthly_schedule()

        for i, act in enumerate(self.opt_acts, 5):
            row_vals = [
                act.activity_id,
                act.name,
                f"{act.duration_days} d",
                f"Day {act.es} - {act.ef}"
            ]
            month_pcts = matrix.get(act.activity_id, [0.0] * num_months)

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular

            for m_idx, pct in enumerate(month_pcts, 5):
                cell = ws.cell(row=i, column=m_idx, value=pct / 100.0 if pct > 0 else "")
                cell.border = self.border_thin
                cell.alignment = Alignment(horizontal="center")
                if pct > 0:
                    cell.number_format = "0.0%"
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                    cell.font = Font(name="Calibri", size=9, bold=True, color="1F4E79")
            ws.row_dimensions[i].height = 20

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 13: S-CURVE (OPTIMIZED)
    # -------------------------------------------------------------
    def _build_sheet_13_scurve(self, wb):
        ws = wb.create_sheet(title="13_S_Curve")
        self._create_banner(ws, "PROJECT S-CURVE & PROGRESS TRAJECTORY (OPTIMIZED)", f"Cumulative progress trajectory and monthly expenditure curve for {self.res.calculated_duration_days} days", 6)
        
        headers = [
            "Month", "Timeline Window", "Planned Monthly Progress %", "Cumulative Progress %",
            "Monthly Cost (₹ Cr)", "Cumulative Cost (₹ Cr)"
        ]
        self._apply_header(ws, 4, headers)

        start_row = 5
        for i, item in enumerate(self.res.monthly_progress, start_row):
            row_vals = [
                item.month_name,
                f"Day {item.period_start_day} - {item.period_end_day}",
                item.period_planned_progress_pct / 100.0,
                item.cumulative_planned_progress_pct / 100.0,
                item.period_expenditure_crore,
                item.cumulative_expenditure_crore
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                
                if col_idx in (3, 4):
                    cell.number_format = "0.00%"
                elif col_idx in (5, 6):
                    cell.number_format = "₹#,##0.00"

                if i % 2 == 0:
                    cell.fill = self.fill_zebra
            ws.row_dimensions[i].height = 20

        end_row = start_row + len(self.res.monthly_progress) - 1

        chart = LineChart()
        chart.title = f"Optimized Cumulative Progress S-Curve ({self.res.calculated_duration_days} Days)"
        chart.style = 13
        chart.y_axis.title = "Cumulative Planned Progress (%)"
        chart.x_axis.title = "Project Period (Months)"
        chart.width = 18
        chart.height = 12

        data = Reference(ws, min_col=4, min_row=4, max_row=end_row)
        cats = Reference(ws, min_col=1, min_row=5, max_row=end_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        if chart.series:
            s1 = chart.series[0]
            s1.graphicalProperties.line.solidFill = "1B365D"
            s1.graphicalProperties.line.width = 25000
            s1.smooth = True

        ws.add_chart(chart, "H4")
        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 14: CRITICAL PATH
    # -------------------------------------------------------------
    def _build_sheet_14_critical_path(self, wb):
        ws = wb.create_sheet(title="14_Critical_Path")
        self._create_banner(ws, "CRITICAL PATH ANALYSIS (BASELINE VS. OPTIMIZED)", "Zero-float critical path sequences driving baseline vs constraint-optimized completion", 7)
        
        ws.cell(row=4, column=1, value="CRITICAL PATH SEQUENCES").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells("A4:G4")

        summary_rows = [
            ("Baseline Critical Path (Unconstrained)", " -> ".join(self.res.baseline_critical_path)),
            ("Optimized Critical Path (Constraint-Based)", " -> ".join(self.res.critical_path)),
            ("Baseline Duration", f"{self.res.baseline_duration_days} Days"),
            ("Optimized Scheduled Duration", f"{self.res.calculated_duration_days} Days (Met <= {self.p_in.duration_days} d)"),
            ("Duration Compression", f"{self.res.duration_saved_days} Days Saved"),
        ]

        curr = 5
        for label, val in summary_rows:
            c1 = ws.cell(row=curr, column=1, value=label)
            c1.font = self.font_bold
            c1.fill = self.fill_card
            c1.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=2)

            c2 = ws.cell(row=curr, column=3, value=val)
            c2.font = self.font_regular
            c2.border = self.border_thin
            ws.merge_cells(start_row=curr, start_column=3, end_row=curr, end_column=7)

            if "Optimized Scheduled Duration" in label and self.res.calculated_duration_days <= self.p_in.duration_days:
                c2.fill = self.fill_pass
                c2.font = Font(bold=True, color=self.c_pass_fg)
            elif "Duration Compression" in label:
                c2.fill = self.fill_pass
                c2.font = Font(bold=True, color=self.c_pass_fg)

            ws.row_dimensions[curr].height = 20
            curr += 1

        curr += 1
        ws.cell(row=curr, column=1, value="OPTIMIZED CRITICAL PATH ACTIVITIES IN DETAIL").font = Font(size=11, bold=True, color=self.c_navy)
        ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=7)
        curr += 1

        headers = ["Activity ID", "Activity Description", "Duration (Days)", "Start Day", "Finish Day", "Cost (₹ Crore)", "Weight %"]
        self._apply_header(ws, curr, headers)
        curr += 1

        crit_acts = [a for a in self.opt_acts if a.is_critical]
        for act in crit_acts:
            row_vals = [
                act.activity_id,
                act.name,
                act.duration_days,
                act.es,
                act.ef,
                round(act.cost.total_amount / 10000000.0, 2),
                act.weight_pct / 100.0
            ]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=curr, column=c_idx, value=val)
                cell.border = self.border_thin
                cell.font = self.font_regular
                cell.fill = self.fill_crit
                if c_idx in (3, 4, 5):
                    cell.alignment = Alignment(horizontal="center")
                elif c_idx == 7:
                    cell.number_format = "0.00%"
            ws.row_dimensions[curr].height = 20
            curr += 1

        self._auto_fit_columns(ws)
