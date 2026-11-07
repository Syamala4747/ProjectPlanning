import pytest
import random
from datetime import date
from types import SimpleNamespace

from scripts.actual_to_entity import (
    _apply_entity_type_reference,
    _build_edl,
    _build_project,
)


@pytest.mark.parametrize(
    ("project_type", "expected_terrain"),
    [
        ("road", "plains"),
        ("pipeline", None),
        ("building", None),
    ],
)
def test_project_builder_only_sets_terrain_for_road(project_type, expected_terrain):
    project = _build_project(
        [],
        {"Terrain": "plain"},
        project_type,
        "sample-project",
    )

    assert project.get("terrain") == expected_terrain


@pytest.mark.parametrize(
    ("project_type", "expected_ui_type"),
    [
        ("road", "Highway"),
        ("pipeline", "Infrastructure"),
        ("building", "Building"),
    ],
)
def test_project_reference_does_not_add_terrain_to_nonroad_projects(
    project_type, expected_ui_type
):
    entities = {
        "project": [{
            "projectId": "sample-project",
            "projectType": project_type,
            "terrain": "plains",
        }],
    }

    _apply_entity_type_reference(entities)

    project = entities["project"][0]
    assert project["projectType"] == expected_ui_type
    if project_type == "road":
        assert project["terrain"] == "plains"
    else:
        assert "terrain" not in project


def test_edl_preserves_source_fuel_when_activity_has_no_operator():
    plan = SimpleNamespace(
        activity_id="ACT-01",
        labour_lines=[{"category": "unskilled", "workers": 2}],
        equipment_lines=[],
        equipment_cost=0.0,
    )
    log = SimpleNamespace(
        activity_id="ACT-01",
        date=date(2026, 11, 10),
        quantity=1.0,
        project_day=1,
        activity_working_day=1,
        unit="metre",
        equipment_cost=0.0,
        fuel_consumption=54.0,
        fuel_cost=5400.0,
        labour_cost=200.0,
    )

    edls, _ = _build_edl(
        [log],
        [plan],
        {},
        [],
        "sample-project",
        random.Random(42),
    )

    assert sum(row["fuelLitres"] for row in edls) == pytest.approx(54.0)
    assert sum(row["fuelCostINR"] for row in edls) == pytest.approx(5400.0)
    assert sum(row["fuelLitres"] > 0 for row in edls) == 1
