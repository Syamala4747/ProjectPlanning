"""Tests for fail-closed Excel-to-entity validation reporting."""

import json
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.entity_sanity_checks import (
    EntityCheck,
    EntityValidationReport,
    validate_entity_export,
)
from scripts.actual_to_entity import _apply_entity_type_reference
from scripts.push_to_db import validate_json_for_push


def test_missing_source_sheets_marks_every_check_incomplete(tmp_path):
    source = tmp_path / "incomplete.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "Summary"
    workbook.save(source)

    report = validate_entity_export(str(source), {})

    assert len(report.checks) == 26
    assert all(check.status == "INCOMPLETE" for check in report.checks)
    assert report.status == "INCOMPLETE"
    assert not report.passed


def test_incomplete_check_cannot_be_reported_as_pass():
    report = EntityValidationReport([
        EntityCheck(1, "source-backed check", "INCOMPLETE", "Source data unavailable."),
    ])

    assert report.status == "INCOMPLETE"
    assert report.to_dict()["status"] == "INCOMPLETE"
    assert not report.passed


def _workbook_with_required_sheets(path):
    workbook = openpyxl.Workbook()
    workbook.active.title = "Plan"
    workbook.active.append(["Activity ID"])
    workbook.create_sheet("Engineer_DailyLog").append(["Project Day"])
    workbook.create_sheet("03_Assumptions")
    workbook.save(path)


def test_mir_activity_date_coverage_passes_without_asset_ids(tmp_path):
    source = tmp_path / "activity_date_coverage.xlsx"
    _workbook_with_required_sheets(source)
    entities = {
        "engineering_daily_logs": [
            {"date": "2026-01-01", "activityId": "ACT001", "assetId": "ASSET001"},
        ],
        "material_issue_request": [
            {"date": "2026-01-01", "activityId": "ACT001", "materialId": "MAT001"},
            {"date": "2026-01-01", "activityId": "ACT001", "materialId": "MAT002"},
        ],
    }

    report = validate_entity_export(str(source), entities)
    check = report.checks[7]

    assert check.status == "PASS"
    assert "1 MIR activity/date groups" in check.details


def test_mir_activity_date_coverage_fails_without_matching_edl(tmp_path):
    source = tmp_path / "activity_date_mismatch.xlsx"
    _workbook_with_required_sheets(source)
    entities = {
        "engineering_daily_logs": [
            {"date": "2026-01-01", "activityId": "ACT001", "assetId": "ASSET001"},
        ],
        "material_issue_request": [
            {"date": "2026-01-02", "activityId": "ACT001", "materialId": "MAT001"},
        ],
    }

    report = validate_entity_export(str(source), entities)

    assert report.checks[7].status == "FAIL"
    assert "MIR has no matching EDL activity/date group" in report.checks[7].details


def _assert_reference_value(value, schema, location):
    schema_type = schema.get("type")
    allowed_types = schema_type if isinstance(schema_type, list) else [schema_type]
    if schema_type:
        valid = (
            ("string" in allowed_types and isinstance(value, str))
            or ("object" in allowed_types and isinstance(value, dict))
            or ("array" in allowed_types and isinstance(value, list))
            or ("boolean" in allowed_types and isinstance(value, bool))
            or ("integer" in allowed_types and isinstance(value, int) and not isinstance(value, bool))
            or ("number" in allowed_types and isinstance(value, (int, float)) and not isinstance(value, bool))
        )
        assert valid, f"{location} has type {type(value).__name__}, expected {schema_type}"
    if "enum" in schema:
        assert value in schema["enum"], f"{location} value {value!r} is outside its enum"
    if isinstance(value, dict):
        for field, child_schema in schema.get("properties", {}).items():
            assert field in value, f"{location}.{field} is missing"
            _assert_reference_value(value[field], child_schema, f"{location}.{field}")
    if isinstance(value, list) and schema.get("items"):
        for index, item in enumerate(value):
            _assert_reference_value(item, schema["items"], f"{location}[{index}]")


def test_generated_entity_shape_includes_all_reference_fields():
    reference = json.loads(
        (ROOT / "entityTypes_ref.json").read_text(encoding="utf-8")
    )
    slug_to_key = {
        "supervisor-update": "supervisor_update",
        "engineering-daily-logs": "engineering_daily_logs",
        "material-procurement-request": "material_procurement_request",
        "material-issue-request": "material_issue_request",
        "material-transaction": "material_transaction",
        "store-transaction": "store_transaction",
        "store-inventory": "store_inventory",
        "daily-log-summary": "daily_log_summary",
        "material-request": "material_request",
        "daily-log": "daily_log",
    }
    entities = {
        slug_to_key.get(entity["slug"], entity["slug"]): [{}]
        for entity in reference
    }
    entities["project"] = [{"projectType": "road"}]
    entities["asset"] = [{"assetId": "ASSET1", "assetName": "Excavator"}]
    entities["material"] = [{"materialId": "MAT1", "materialName": "Cement", "unit": "bag"}]
    entities["supervisor_update"] = [{
        "machinery": {"ASSET1": 2},
        "materialsReceived": {"Cement": 3},
        "actualProgress": 1.5,
        "unit": "tonne",
    }]
    _apply_entity_type_reference(entities)

    for entity_type in reference:
        key = slug_to_key.get(entity_type["slug"], entity_type["slug"])
        schema = entity_type["jsonSchema"]
        row = entities[key][0]
        assert set(schema["properties"]) <= set(row)
        assert set(schema.get("required", [])) <= set(row)
        _assert_reference_value(row, schema, entity_type["slug"])

    supervisor = entities["supervisor_update"][0]
    assert supervisor["machinery"] == [{"assetName": "Excavator", "count": 2}]
    assert supervisor["materialsReceived"] == [
        {"materialName": "Cement", "quantity": 3, "unit": "bag"}
    ]


def test_push_writes_incomplete_report_when_matching_workbook_is_missing(tmp_path):
    source_json = tmp_path / "project_actual_data.json"
    source_json.write_text("{}", encoding="utf-8")

    report = validate_json_for_push(str(source_json), {})
    report_path = tmp_path / "project_entity_validation.xlsx"
    saved_report = openpyxl.load_workbook(report_path, data_only=True)
    sheet = saved_report["Summary"]

    assert report.status == "INCOMPLETE"
    assert sheet["B1"].value == "INCOMPLETE"
    assert sheet.max_row == 6 + 26
    assert all(sheet.cell(row, 3).value == "INCOMPLETE" for row in range(7, 33))
    assert all(
        "Cannot open source workbook" in sheet.cell(row, 4).value
        for row in range(7, 33)
    )
    saved_report.close()


def test_validation_report_workbook_has_check_details_and_status_colors(tmp_path):
    from scripts.entity_sanity_checks import write_validation_workbook

    output = tmp_path / "report.xlsx"
    report = EntityValidationReport([
        EntityCheck(1, "verified check", "PASS", "All values reconcile."),
        EntityCheck(2, "failing check", "FAIL", "Inflow differs from outflow."),
        EntityCheck(3, "unverified check", "INCOMPLETE", "Source data unavailable."),
    ])
    write_validation_workbook(report, str(output))
    workbook = openpyxl.load_workbook(output, data_only=True)
    sheet = workbook["Summary"]

    assert sheet["A6"].value == "#"
    assert sheet["C7"].value == "PASS"
    assert sheet["D7"].value
    assert sheet["C8"].value == "FAIL"
    assert sheet["D8"].value == "Inflow differs from outflow."
    assert sheet["C9"].value == "INCOMPLETE"
    assert sheet.auto_filter.ref == "A6:D9"
    workbook.close()
