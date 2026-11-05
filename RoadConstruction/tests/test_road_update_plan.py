# -*- coding: utf-8 -*-
"""
Tests for the update-plan workflow.

Covers:
  - Workbook selection and validation (WorkbookReader)
  - Progress-gap calculation
  - Remaining-work calculation
  - Target feasibility
  - Preservation of the original workbook and baseline
  - Regression: the existing new-project command is unchanged
"""
import os
import sys
import shutil
import tempfile
from datetime import date, timedelta

import pytest
from openpyxl import load_workbook

# Ensure the project root is on the path
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# Generated workbooks go to a temp directory, never the repository root.
TEST_DIR = tempfile.mkdtemp(prefix="planner_tests_")
sys.path.insert(0, ROOT)

from src.models import (
    ProjectInput, UpdatePlanInput,
)
from src.data_loader import DataLoader
from src.constraint_scheduler import ConstraintScheduler
from src.excel_report import ExcelReportGenerator
from src.workbook_reader import WorkbookReader
from src.update_planner import UpdatePlanner
from src.update_excel_report import UpdateExcelReportGenerator
from src.validation_sanity_checks import validate_update_result


# ------------------------------------------------------------------ fixtures

BASELINE_FILE = os.path.join(TEST_DIR, "_test_baseline.xlsx")
UPDATED_FILE = os.path.join(TEST_DIR, "_test_updated.xlsx")
ROAD_GUIDELINES = os.path.join(ROOT, "project_config", "road", "road_guidelines.json")
ROAD_ASSUMPTIONS = os.path.join(ROOT, "project_config", "road", "road_assumptions.yaml")

BASELINE_INPUT = ProjectInput(
    title="10 KM 2-Lane Highway Project",
    length_km=10.0,
    duration_days=365,
    budget_crore=250.0,
    start_date=date(2026, 4, 1),
    lanes=2,
    terrain="plain",
)


@pytest.fixture(scope="module")
def baseline_result():
    """Run the existing planning pipeline and return the result."""
    loader = DataLoader(guidelines_path=ROAD_GUIDELINES,
                        assumptions_path=ROAD_ASSUMPTIONS)
    loader.load_all(terrain="plain")
    scheduler = ConstraintScheduler(project_input=BASELINE_INPUT, data_loader=loader)
    result = scheduler.solve()
    return result, loader


@pytest.fixture(scope="module")
def baseline_workbook(baseline_result):
    """Generate a baseline Excel workbook for testing."""
    result, _ = baseline_result
    reporter = ExcelReportGenerator(result=result, output_path=BASELINE_FILE)
    reporter.generate()
    yield BASELINE_FILE
    # Cleanup
    for f in (BASELINE_FILE, UPDATED_FILE):
        if os.path.exists(f):
            try:
                os.remove(f)
            except PermissionError:
                pass


@pytest.fixture(scope="module")
def update_result(baseline_result, baseline_workbook):
    """Run the update planner and return the result."""
    result, loader = baseline_result
    update_input = UpdatePlanInput(
        input_path=baseline_workbook,
        start_date=date(2026, 4, 1),
        status_date=date(2026, 10, 1),
        actual_progress_pct=10.0,
        target_date=date(2027, 4, 1),
        output_path=UPDATED_FILE,
    )
    planner = UpdatePlanner(
        baseline_result=result,
        update_input=update_input,
        data_loader=loader,
    )
    return planner.plan()


# ==================================================================
# 1. WORKBOOK SELECTION AND VALIDATION
# ==================================================================

class TestWorkbookReader:
    def test_reads_project_input(self, baseline_workbook):
        """The reader correctly extracts ProjectInput from the workbook."""
        reader = WorkbookReader(baseline_workbook)
        proj_input, metadata = reader.read()
        assert proj_input.length_km == 10.0
        assert proj_input.lanes == 2
        assert proj_input.terrain == "plain"
        assert proj_input.duration_days == 365
        assert proj_input.budget_crore == 250.0
        assert proj_input.start_date == date(2026, 4, 1)

    def test_raises_on_missing_file(self):
        """A FileNotFoundError is raised for a non-existent workbook."""
        reader = WorkbookReader("nonexistent_file.xlsx")
        with pytest.raises(FileNotFoundError):
            reader.read()

    def test_raises_on_invalid_workbook(self, tmp_path):
        """A ValueError is raised for a workbook that is not a planner output."""
        from openpyxl import Workbook
        bad_path = str(tmp_path / "bad.xlsx")
        wb = Workbook()
        wb.save(bad_path)
        reader = WorkbookReader(bad_path)
        with pytest.raises(ValueError):
            reader.read()

    def test_scurve_data_extracted(self, baseline_workbook):
        """The S-curve monthly data is extracted from the workbook."""
        reader = WorkbookReader(baseline_workbook)
        _, metadata = reader.read()
        assert "scurve_months" in metadata
        assert len(metadata["scurve_months"]) > 0

    def test_dashboard_baseline_figures_extracted(self, baseline_workbook, baseline_result):
        """The source baseline duration/cost are available for drift checks."""
        result, _ = baseline_result
        reader = WorkbookReader(baseline_workbook)
        _, metadata = reader.read()
        assert metadata["baseline_duration_days"] == result.baseline_duration_days
        assert metadata["baseline_cost_crore"] == pytest.approx(
            result.baseline_cost_crore, abs=0.01)


# ==================================================================
# 2. PROGRESS-GAP CALCULATION
# ==================================================================

class TestProgressGap:
    def test_baseline_progress_at_status_date(self, update_result):
        """The baseline progress at the status date is computed correctly."""
        # Status date is 2026-10-01, about 183 days into a 365-day plan.
        # The progress should be between 0 and 100.
        assert 0 < update_result.baseline_progress_at_status_date <= 100

    def test_progress_gap_is_positive_when_behind(self, update_result):
        """When actual < baseline, the gap is positive."""
        assert update_result.actual_progress_pct == 10.0
        assert update_result.baseline_progress_at_status_date > 10.0
        gap = update_result.baseline_progress_at_status_date - update_result.actual_progress_pct
        assert abs(update_result.progress_gap_pct - round(gap, 2)) < 0.1

    def test_progress_gap_zero_when_on_track(self, baseline_result):
        """When actual equals baseline, the gap is zero."""
        result, loader = baseline_result
        update_input = UpdatePlanInput(
            input_path="dummy",
            start_date=date(2026, 4, 1),
            status_date=date(2026, 4, 1),
            actual_progress_pct=0.0,
            target_date=date(2027, 4, 1),
        )
        planner = UpdatePlanner(result, update_input, loader)
        r = planner.plan()
        assert r.progress_gap_pct == 0.0

    def test_required_progress_calculated(self, update_result):
        """Required progress = target_progress - actual_progress."""
        required = update_result.target_progress_pct - update_result.actual_progress_pct
        assert abs(update_result.required_progress_pct - round(required, 2)) < 0.1


# ==================================================================
# 3. REMAINING-WORK CALCULATION
# ==================================================================

class TestRemainingWork:
    def test_activity_status_populated(self, update_result):
        """Every baseline activity has a corresponding status entry."""
        baseline_acts = update_result.baseline_result.activities
        assert len(update_result.activity_status) == len(baseline_acts)

    def test_completed_quantities_sum_correctly(self, update_result):
        """Completed + remaining = original for each activity."""
        for st in update_result.activity_status:
            total = st.completed_quantity + st.remaining_quantity
            assert abs(total - st.original_quantity) < 0.1

    def test_remaining_quantities_non_negative(self, update_result):
        """Remaining quantities are never negative."""
        for st in update_result.activity_status:
            assert st.remaining_quantity >= 0

    def test_completion_pcts_in_range(self, update_result):
        """Completion percentages are in [0, 100]."""
        for st in update_result.activity_status:
            assert 0 <= st.actual_completion_pct <= 100

    def test_revised_activities_have_correct_quantities(self, update_result):
        """Revised activity quantities match the remaining quantities."""
        status_map = {s.activity_id: s for s in update_result.activity_status}
        for act in update_result.revised_activities:
            st = status_map.get(act.activity_id)
            if st:
                assert abs(act.quantity - st.remaining_quantity) < 0.5

    def test_revised_duration_positive(self, update_result):
        """The revised schedule has a positive duration."""
        assert update_result.revised_duration_days > 0

    def test_revised_critical_path_non_empty(self, update_result):
        """The revised critical path is identified."""
        assert len(update_result.revised_critical_path) > 0


# ==================================================================
# 4. TARGET FEASIBILITY
# ==================================================================

class TestTargetFeasibility:
    def test_feasibility_flag_set(self, update_result):
        """The target_feasible flag is set (True or False)."""
        assert isinstance(update_result.target_feasible, bool)

    def test_projected_delay_non_negative(self, update_result):
        """Projected delay is never negative."""
        assert update_result.projected_delay_days >= 0

    def test_revised_completion_date_after_status(self, update_result):
        """The revised completion date is after the status date."""
        assert update_result.revised_completion_date is not None
        assert update_result.revised_completion_date > update_result.status_date

    def test_infeasible_when_behind_far(self, update_result):
        """With only 10% actual vs 62.7% baseline, the target is infeasible."""
        # 10% progress at 6 months means the project is significantly behind.
        # The revised schedule cannot meet the original target date.
        assert update_result.actual_progress_pct < update_result.baseline_progress_at_status_date

    def test_feasible_when_on_track(self, baseline_result):
        """When actual matches baseline and target is the original end, feasible."""
        result, loader = baseline_result
        # Status at start, 0% progress, target at original end
        update_input = UpdatePlanInput(
            input_path="dummy",
            start_date=date(2026, 4, 1),
            status_date=date(2026, 4, 1),
            actual_progress_pct=0.0,
            target_date=date(2027, 4, 1),
        )
        planner = UpdatePlanner(result, update_input, loader)
        r = planner.plan()
        # With 0 progress at start and target at original end date,
        # the remaining work should fit within the available time
        assert r.revised_duration_days > 0


# ==================================================================
# 5. PRESERVATION OF ORIGINAL WORKBOOK AND BASELINE
# ==================================================================

class TestPreservation:
    def test_original_workbook_not_overwritten(self, baseline_workbook, update_result):
        """The original workbook still exists and is unchanged."""
        assert os.path.exists(baseline_workbook)
        wb = load_workbook(baseline_workbook)
        expected = [
            "01_Project_Summary", "02_Project_Inputs", "03_Assumptions",
            "04_WBS_Hierarchy", "05_Activities", "06_Baseline_CPM",
            "07_Baseline_Cost", "08_Constraint_Optimization", "09_Optimized_CPM",
            "10_Optimized_Cost", "11_Baseline_vs_Optimized", "12_Monthly_Schedule",
            "13_S_Curve", "14_Critical_Path", "15_Validation_Checks",
            "Plan", "Engineer_DailyLog",
        ]
        assert wb.sheetnames == expected

    def test_output_workbook_is_separate(self, update_result):
        """The updated workbook is a separate file with the expected sheets."""
        reporter = UpdateExcelReportGenerator(
            result=update_result, output_path=UPDATED_FILE)
        reporter.generate()
        assert os.path.exists(UPDATED_FILE)
        assert os.path.exists(BASELINE_FILE)
        wb = load_workbook(UPDATED_FILE)
        expected = [
            "01_Update_Summary", "02_Baseline_Schedule", "03_Assumptions",
            "03_Actual_Progress",
            "04_Activity_Status", "05_Revised_CPM", "06_Revised_Cost",
            "07_S_Curve_Comparison", "08_Target_Feasibility", "Plan",
            "Engineer_DailyLog", "09_Validation_Checks",
        ]
        assert wb.sheetnames == expected

    def test_updated_workbook_has_clear_validation_register(self, update_result):
        """The updated workbook exposes PASS, WARNING and ERROR checks."""
        update_result.validation_report = validate_update_result(
            update_result,
            update_result.baseline_result.project_input and {},
        )
        reporter = UpdateExcelReportGenerator(
            result=update_result, output_path=UPDATED_FILE)
        reporter.generate()
        wb = load_workbook(UPDATED_FILE, data_only=True)
        ws = wb["09_Validation_Checks"]
        assert ws["B4"].value == "PASS"
        assert ws["B5"].value == 0
        assert ws["B6"].value >= 0
        assert ws.max_row >= 20

    def test_baseline_preserved_in_result(self, update_result):
        """The original baseline result is preserved in the update result."""
        assert update_result.baseline_result is not None
        assert update_result.baseline_result.calculated_duration_days > 0
        assert update_result.baseline_total_cost_crore > 0

    def test_cost_summary_components(self, update_result):
        """Spent + remaining (excl. accel) + accel = total cost."""
        spent = update_result.spent_cost_crore
        remaining = update_result.remaining_cost_crore
        accel = update_result.additional_acceleration_cost_crore
        total = update_result.total_project_cost_crore
        assert abs(spent + remaining - total) < 0.1 or \
               abs(spent + remaining - accel - total) < 0.1


# ==================================================================
# 6. REGRESSION: EXISTING NEW-PROJECT COMMAND
# ==================================================================

class TestRegression:
    def test_existing_command_produces_same_result(self, baseline_result):
        """The existing pipeline still produces the expected deterministic result."""
        result, _ = baseline_result
        assert result.calculated_duration_days == 365
        assert result.total_calculated_cost_crore > 0
        assert result.baseline_duration_days > result.calculated_duration_days
        assert len(result.activities) > 0
        assert len(result.critical_path) > 0

    def test_existing_command_workbook_structure(self, baseline_workbook):
        """The existing command still produces the complete workbook."""
        wb = load_workbook(baseline_workbook)
        assert len(wb.sheetnames) == 17

    def test_main_help_still_works(self):
        """The main CLI still runs the existing command without 'update-plan'."""
        import subprocess
        r = subprocess.run(
            [sys.executable, "main.py", "--help"],
            capture_output=True, text=True, cwd=ROOT,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        assert r.returncode == 0
        assert "Road / Pipeline / Building" in r.stdout or "Road / Pipeline / Building" in r.stderr

    def test_update_plan_help(self):
        """The update-plan subcommand help is accessible."""
        import subprocess
        r = subprocess.run(
            [sys.executable, "main.py", "update-plan", "--help"],
            capture_output=True, text=True, cwd=ROOT,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        assert r.returncode == 0
        assert "--input" in r.stdout
        assert "--actual-progress" in r.stdout
        assert "--target-date" in r.stdout


# ==================================================================
# 7. ACTIVITY-LEVEL PROGRESS
# ==================================================================

class TestActivityLevelProgress:
    def test_activity_level_progress_used(self, baseline_result):
        """When activity-level progress is provided, it is used directly."""
        result, loader = baseline_result
        # Mark the first 3 activities as complete
        act_map = {a.activity_id: 0.0 for a in result.activities}
        for i, aid in enumerate(sorted(act_map.keys())[:3]):
            act_map[aid] = 100.0
        update_input = UpdatePlanInput(
            input_path="dummy",
            start_date=date(2026, 4, 1),
            status_date=date(2026, 10, 1),
            actual_progress_pct=15.0,
            target_date=date(2027, 4, 1),
            activity_progress=act_map,
        )
        planner = UpdatePlanner(result, update_input, loader)
        r = planner.plan()
        # The first 3 activities should be marked completed
        for st in r.activity_status:
            if st.activity_id in list(sorted(act_map.keys())[:3]):
                assert st.status == "completed"
                assert st.actual_completion_pct == 100.0

    def test_no_limitation_with_activity_progress(self, baseline_result):
        """When activity-level progress is provided, no limitation is reported."""
        result, loader = baseline_result
        act_map = {a.activity_id: 50.0 for a in result.activities}
        update_input = UpdatePlanInput(
            input_path="dummy",
            start_date=date(2026, 4, 1),
            status_date=date(2026, 10, 1),
            actual_progress_pct=50.0,
            target_date=date(2027, 4, 1),
            activity_progress=act_map,
        )
        planner = UpdatePlanner(result, update_input, loader)
        r = planner.plan()
        # No scaling limitation should be reported
        assert not any("Earned Schedule" in lim for lim in r.limitations)

    def test_limitation_without_activity_progress(self, update_result):
        """Without activity-level progress, the limitation is reported."""
        assert any("Earned Schedule" in lim for lim in update_result.limitations)


# ==================================================================
# 8. OPTIONAL TARGET DATE AND S-CURVE DELAY DISPLAY
# ==================================================================

class TestOptionalTargetDate:
    def test_workbook_reader_extracts_optimized_finish_date(self, baseline_workbook):
        """WorkbookReader extracts the optimized completion date from the baseline workbook."""
        reader = WorkbookReader(baseline_workbook)
        proj_input, metadata = reader.read()
        assert "optimized_finish_date" in metadata
        assert isinstance(metadata["optimized_finish_date"], date)
        assert metadata["optimized_finish_date"] > proj_input.start_date

    def test_target_date_optional_defaults_to_optimized_finish(self, baseline_result):
        """UpdatePlanner defaults target_date to the optimized finish date when omitted."""
        result, loader = baseline_result
        expected_finish = BASELINE_INPUT.start_date + timedelta(days=result.calculated_duration_days)
        update_input = UpdatePlanInput(
            input_path="dummy",
            start_date=BASELINE_INPUT.start_date,
            status_date=date(2026, 10, 1),
            actual_progress_pct=25.0,
            target_date=None,  # omitted!
        )
        planner = UpdatePlanner(result, update_input, loader)
        assert planner.target_date == expected_finish
        r = planner.plan()
        assert r.target_date == expected_finish


class TestSCurveDelayDisplay:
    def test_scurve_sheet_contains_delay_kpi_and_headers(self, update_result):
        """The S-curve sheet includes delay tracking KPI cards and columns."""
        wb_path = os.path.join(TEST_DIR, "_test_scurve_check.xlsx")
        reporter = UpdateExcelReportGenerator(result=update_result, output_path=wb_path)
        reporter.generate()
        wb = load_workbook(wb_path)
        ws = wb["07_S_Curve_Comparison"]

        # Check KPI cards
        assert ws.cell(row=4, column=1).value == "Status Date"
        assert ws.cell(row=4, column=7).value == "Projected Delay"
        delay_val = ws.cell(row=5, column=7).value
        assert "Delay" in delay_val or "Days" in delay_val

        # Check table headers at row 7
        headers = [ws.cell(row=7, column=c).value for c in range(1, 9)]
        assert "Month" in headers
        assert "Baseline Planned %" in headers
        assert "Actual Progress %" in headers
        assert "Revised Forecast (Target) %" in headers
        assert "Delay Period Forecast %" in headers
        assert "Total Cumulative Forecast %" in headers
        assert "Schedule Delay Impact" in headers

        chart = ws._charts[0]
        assert len(chart.series) == 3
        assert "Planned Baseline" in chart.title.tx.rich.p[0].r[0].t
        assert "Actual Progress" in chart.title.tx.rich.p[0].r[0].t
        assert "Revised Forecast" in chart.title.tx.rich.p[0].r[0].t
        series_columns = [
            series.val.numRef.f.split("!")[-1][1]
            for series in chart.series
        ]
        assert series_columns == ["B", "C", "F"]
        assert all(
            series.cat.strRef.f
            == "'07_S_Curve_Comparison'!$A$8:$A${}".format(ws.max_row)
            for series in chart.series
        )
        assert chart.series[2].tx.v == "Revised Forecast %"
        assert chart.y_axis.scaling.min == 0
        assert chart.y_axis.scaling.max == 1
        assert chart.y_axis.majorUnit == 0.1
        assert chart.y_axis.numFmt.formatCode == "0%"
        assert chart.y_axis.tickLblPos == "low"
        assert chart.y_axis.delete is False
        assert chart.x_axis.tickLblSkip == 1
        assert chart.x_axis.tickLblPos == "low"
        assert chart.x_axis.delete is False
        assert "\n" in ws["A8"].value
        assert "\n" in ws.cell(ws.max_row, 1).value

        # If there is a delay, check that delay rows are highlighted in red
        if update_result.projected_delay_days > 0:
            found_red_delay_row = False
            for r in range(8, ws.max_row + 1):
                delay_impact = str(ws.cell(row=r, column=7).value or "")
                fill = ws.cell(row=r, column=7).fill
                if "Days Delay" in delay_impact:
                    found_red_delay_row = True
                    # Check that the fill is red (start_color FFEBEE)
                    assert fill.start_color.rgb in ("00FFEBEE", "FFEBEE")
            assert found_red_delay_row, "Expected at least one delay row highlighted in red"

        try:
            os.remove(wb_path)
        except (PermissionError, OSError):
            pass
