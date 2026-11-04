"""Tests for shared schedule validation and workbook validation."""

import sys
from datetime import date
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.data_loader import DataLoader
from src.constraint_scheduler import ConstraintScheduler
from src.excel_report import ExcelReportGenerator
from src.models import ProjectInput
from src.validation_sanity_checks import validate_exported_workbook, validate_schedule_result
from src.duration_calculator import DurationCalculator


PIPELINE_GUIDELINES = ROOT / "project_config" / "pipeline" / "pipeline_guidelines.json"
PIPELINE_ASSUMPTIONS = ROOT / "project_config" / "pipeline" / "pipeline_assumptions.yaml"
ROAD_GUIDELINES = ROOT / "project_config" / "road" / "road_guidelines.json"
ROAD_ASSUMPTIONS = ROOT / "project_config" / "road" / "road_assumptions.yaml"
BUILDING_GUIDELINES = ROOT / "project_config" / "building" / "building_guidelines.json"
BUILDING_ASSUMPTIONS = ROOT / "project_config" / "building" / "building_assumptions.yaml"


def test_missing_productivity_assumption_is_rejected():
    with pytest.raises(ValueError, match="missing productivity_rate"):
        DurationCalculator(
            {"activity_master": [{
                "key": "earthwork",
                "productivity_unit": "m3/day",
                "base_crews": 1,
            }]},
            project_length_km=1.0,
        )


def test_loader_marks_missing_rate_as_hard_configuration_error():
    loader = DataLoader()
    loader.guidelines = {"data": {"activities": []}}
    loader.assumptions = {
        "activity_master": [{
            "key": "earthwork",
            "productivity_rate": 10,
            "productivity_unit": "m3/day",
            "base_crews": 1,
        }],
        "direct_unit_rates_inr": {},
    }
    loader.validate()
    assert any("no fallback rate is allowed" in error for error in loader.errors)


def test_hilly_road_uses_guidelines_hilly_catalogue():
    loader = DataLoader(str(ROAD_GUIDELINES), str(ROAD_ASSUMPTIONS))
    loader.load_all(terrain="hilly")
    assert loader.assumptions["selected_terrain"] == "hilly"
    catalogue = loader.assumptions["activity_master"]
    assert catalogue
    excavation = next(item for item in catalogue if item["key"] == "excavation")
    assert excavation["productivity_rate"] == pytest.approx(413.0)


def test_validation_report_checks_assumption_sources(pipeline_result):
    result, loader = pipeline_result
    result.activities[0].activity_key = "missing-assumption-key"
    report = validate_schedule_result(result, loader.assumptions)
    assert any(issue.code == "ASSUMPTION_SOURCE_MISSING" for issue in report.errors)


def _run_project(project, guidelines_path, assumptions_path, *, acceleration=True):
    """Load project data and return its solved schedule with the loader."""
    loader = DataLoader(str(guidelines_path), str(assumptions_path))
    loader.load_all(terrain=project.terrain)
    scheduler = ConstraintScheduler(
        project,
        loader,
        enable_acceleration_cost=acceleration,
    )
    return scheduler.solve(), loader


def _pipeline_project():
    """Create the smallest representative pipeline validation scenario."""
    return ProjectInput(
        project_type="pipeline",
        title="Validation corridor",
        length_km=10.0,
        pipe_diameter_mm=600,
        pipe_material="DI",
        trench_depth_m=2.0,
        number_of_work_fronts=2,
        duration_days=540,
        budget_crore=25.0,
        start_date=date(2026, 4, 1),
    )


@pytest.fixture
def pipeline_result():
    result, loader = _run_project(
        _pipeline_project(),
        PIPELINE_GUIDELINES,
        PIPELINE_ASSUMPTIONS,
    )
    return result, loader


def _road_result():
    """Create the representative road schedule used by the road regression."""
    project = ProjectInput(
        title="Road validation",
        length_km=10.0,
        lanes=2,
        terrain="plain",
        duration_days=365,
        budget_crore=250.0,
        start_date=date(2026, 4, 1),
    )
    return _run_project(
        project,
        ROAD_GUIDELINES,
        ROAD_ASSUMPTIONS,
        acceleration=False,
    )


def _building_result():
    project = ProjectInput(
        project_type="building",
        title="Building validation",
        length_km=0.0,
        built_up_area_sqm=2000.0,
        number_of_floors=4,
        building_type="residential",
        construction_type="rcc",
        quality_finish="standard",
        number_of_blocks=1,
        duration_days=540,
        budget_crore=6.0,
        start_date=date(2026, 4, 1),
    )
    return _run_project(project, BUILDING_GUIDELINES, BUILDING_ASSUMPTIONS)


def test_generated_pipeline_passes_validation(pipeline_result):
    result, loader = pipeline_result
    report = validate_schedule_result(result, loader.assumptions)
    assert report.passed, [issue.message for issue in report.errors]


def test_validation_detects_cost_tampering(pipeline_result):
    result, loader = pipeline_result
    result.activities[0].cost.total_amount += 1.0
    report = validate_schedule_result(result, loader.assumptions)
    assert any(issue.code == "TOTAL_COST_RECONCILIATION" for issue in report.errors)


def test_validation_detects_missing_predecessor(pipeline_result):
    result, loader = pipeline_result
    result.activities[0].predecessors = ["ACT-MISSING"]
    report = validate_schedule_result(result, loader.assumptions)
    assert any(issue.code == "MISSING_PREDECESSOR" for issue in report.errors)


def test_validation_results_are_exported_to_excel(pipeline_result, tmp_path):
    result, loader = pipeline_result
    result.validation_report = validate_schedule_result(result, loader.assumptions)
    output = tmp_path / "validated_pipeline.xlsx"
    ExcelReportGenerator(result, str(output)).generate()

    workbook = openpyxl.load_workbook(output, data_only=True)
    assert "15_Validation_Checks" in workbook.sheetnames
    sheet = workbook["15_Validation_Checks"]
    assert sheet["B4"].value == "PASS"
    assert sheet["B5"].value == 0
    assert sheet.max_row >= 11


def test_exported_workbook_reconciles_with_result(pipeline_result, tmp_path):
    result, loader = pipeline_result
    result.validation_report = validate_schedule_result(result, loader.assumptions)
    output = tmp_path / "reopened_pipeline.xlsx"
    ExcelReportGenerator(result, str(output)).generate()

    report = validate_exported_workbook(str(output), result)
    assert report.passed, [issue.message for issue in report.errors]


def test_exported_plan_detects_activity_input_tampering(pipeline_result, tmp_path):
    result, loader = pipeline_result
    result.validation_report = validate_schedule_result(result, loader.assumptions)
    output = tmp_path / "tampered_plan.xlsx"
    ExcelReportGenerator(result, str(output)).generate()

    workbook = openpyxl.load_workbook(output)
    plan = workbook["Plan"]
    activity_row = next(
        row for row in range(3, plan.max_row + 1)
        if plan.cell(row, 13).value == "Activity"
    )
    plan.cell(activity_row, 8).value = float(plan.cell(activity_row, 8).value) + 1.0
    workbook.save(output)

    report = validate_exported_workbook(str(output), result)
    assert any(issue.code == "WORKBOOK_PLAN_INPUT_MISMATCH" for issue in report.errors)


def test_road_validation_accepts_separate_wastage_and_zero_scope_rows():
    result, loader = _road_result()

    report = validate_schedule_result(result, loader.assumptions)
    assert report.passed, [issue.message for issue in report.errors]


@pytest.mark.parametrize("project_factory", [_road_result, _pipeline_project])
def test_resource_sheets_are_created_for_project_inputs(project_factory, tmp_path):
    if project_factory is _pipeline_project:
        result, loader = _run_project(
            project_factory(), PIPELINE_GUIDELINES, PIPELINE_ASSUMPTIONS
        )
    else:
        result, loader = project_factory()
    result.validation_report = validate_schedule_result(result, loader.assumptions)
    output = tmp_path / f"{result.project_input.project_type}_resources.xlsx"
    ExcelReportGenerator(result, str(output)).generate()
    workbook = openpyxl.load_workbook(output, data_only=True)
    assert {"Plan", "Engineer_DailyLog"} <= set(workbook.sheetnames)
    assert workbook["Plan"].max_row >= 3
    assert workbook["Engineer_DailyLog"].max_row >= 2


def test_building_resource_sheets_are_created(tmp_path):
    result, loader = _building_result()
    result.validation_report = validate_schedule_result(result, loader.assumptions)
    output = tmp_path / "building_resources.xlsx"
    ExcelReportGenerator(result, str(output)).generate()
    workbook = openpyxl.load_workbook(output, data_only=True)
    assert {"Plan", "Engineer_DailyLog"} <= set(workbook.sheetnames)


def test_resource_sheets_have_hierarchical_detail_and_day_counters(pipeline_result, tmp_path):
    result, loader = pipeline_result
    result.validation_report = validate_schedule_result(result, loader.assumptions)
    output = tmp_path / "detailed_resources.xlsx"
    ExcelReportGenerator(result, str(output)).generate()
    workbook = openpyxl.load_workbook(output, data_only=True)

    plan = workbook["Plan"]
    plan_headers = [cell.value for cell in plan[3]]
    assert plan_headers[:12] == [
        "Activity ID", "Activity", "Phase", "Milestone", "Start", "Finish",
        "Duration (wd)", "Quantity", "Unit", "Productivity", "Crew", "Status",
    ]
    assert "Subtotal" in [plan.cell(row, 13).value for row in range(1, plan.max_row + 1)]
    assert "PROJECT TOTALS" in [plan.cell(row, 1).value for row in range(1, plan.max_row + 1)]

    daily = workbook["Engineer_DailyLog"]
    headers = [cell.value for cell in daily[3]]
    assert headers[:7] == [
        "Project Day", "Activity Working Day", "Date", "Weekday",
        "Activity ID", "Activity Name", "Phase",
    ]
    assert daily.max_column == 20
