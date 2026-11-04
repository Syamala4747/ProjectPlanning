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


PIPELINE_GUIDELINES = ROOT / "project_config" / "pipeline" / "pipeline_guidelines.json"
PIPELINE_ASSUMPTIONS = ROOT / "project_config" / "pipeline" / "pipeline_assumptions.yaml"
ROAD_GUIDELINES = ROOT / "project_config" / "road" / "road_guidelines.json"
ROAD_ASSUMPTIONS = ROOT / "project_config" / "road" / "road_assumptions.yaml"


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


def test_road_validation_accepts_separate_wastage_and_zero_scope_rows():
    result, loader = _road_result()

    report = validate_schedule_result(result, loader.assumptions)
    assert report.passed, [issue.message for issue in report.errors]
