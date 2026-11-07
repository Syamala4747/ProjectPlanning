from datetime import date
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.push_to_db import apply_progress_snapshot, update_project_progress


def _entities():
    return {
        "project": [{
            "projectId": "north",
            "startDate": "2026-12-10",
            "overallActualProgress": 0.0,
            "lastProgressUpdateDate": "2026-12-10",
        }],
        "plan": [{"planId": "PLAN1", "plannedEnd": "2027-08-14"}],
        "engineering_daily_logs": [
            {"engineeringDailyLogId": "EDL1", "date": "2027-02-12"},
            {"engineeringDailyLogId": "EDL2", "date": "2027-02-13"},
        ],
        "supervisor_update": [{"date": "2027-02-12"}],
        "material_issue_request": [{"date": "2027-02-13"}],
        "material_procurement_request": [
            {"requestDate": "2027-02-11", "deliveryDate": "2027-02-12"},
            {"requestDate": "2027-02-12", "deliveryDate": "2027-02-13"},
        ],
        "logistics": [
            {"dispatchDate": "2027-02-11", "deliveryDate": "2027-02-12"},
            {"dispatchDate": "2027-02-12", "deliveryDate": "2027-02-13"},
        ],
        "material_transaction": [{"date": "2027-02-13"}],
    }


def test_progress_snapshot_keeps_full_plan_and_filters_events_inclusively():
    source = _entities()

    snapshot = apply_progress_snapshot(source, 30, date(2027, 2, 12))

    assert snapshot["project"][0]["overallActualProgress"] == 30
    assert snapshot["project"][0]["lastProgressUpdateDate"] == "2027-02-12"
    assert snapshot["plan"] == source["plan"]
    assert [row["engineeringDailyLogId"] for row in snapshot["engineering_daily_logs"]] == ["EDL1"]
    assert len(snapshot["supervisor_update"]) == 1
    assert snapshot["material_issue_request"] == []
    assert len(snapshot["material_procurement_request"]) == 1
    assert len(snapshot["logistics"]) == 1
    assert snapshot["material_transaction"] == []
    assert len(source["engineering_daily_logs"]) == 2


@pytest.mark.parametrize("progress", [-0.1, 100.1, float("nan"), float("inf")])
def test_progress_snapshot_rejects_invalid_progress(progress):
    with pytest.raises(ValueError, match="between 0 and 100"):
        apply_progress_snapshot(_entities(), progress, date(2027, 2, 12))


def test_progress_snapshot_rejects_date_before_project_start():
    with pytest.raises(ValueError, match="before project start"):
        apply_progress_snapshot(_entities(), 30, date(2026, 12, 9))


def test_progress_snapshot_rejects_undated_event_record():
    entities = _entities()
    entities["engineering_daily_logs"].append({"engineeringDailyLogId": "EDL3"})

    with pytest.raises(ValueError, match="Invalid or missing date"):
        apply_progress_snapshot(entities, 30, date(2027, 2, 12))


class _Cursor:
    def __init__(self, rows):
        self.rows = rows
        self.calls = []

    def execute(self, query, params):
        self.calls.append((query, params))

    def fetchall(self):
        return self.rows


def test_update_project_progress_updates_only_one_matching_project():
    cursor = _Cursor([("database-project-id",)])
    project = {
        "projectId": "north",
        "overallActualProgress": 30,
        "lastProgressUpdateDate": "2027-02-12",
    }

    updated = update_project_progress(
        cursor,
        {"project": "project-type-id"},
        "account-id",
        project,
    )

    assert updated is True
    assert len(cursor.calls) == 2
    assert cursor.calls[1][1][0:2] == (30, "2027-02-12")
    assert cursor.calls[1][1][-1] == "database-project-id"


def test_update_project_progress_does_not_update_when_project_is_new():
    cursor = _Cursor([])
    assert update_project_progress(
        cursor,
        {"project": "project-type-id"},
        "account-id",
        {
            "projectId": "north",
            "overallActualProgress": 30,
            "lastProgressUpdateDate": "2027-02-12",
        },
    ) is False
    assert len(cursor.calls) == 1


def test_update_project_progress_rejects_ambiguous_duplicate_project_records():
    cursor = _Cursor([("first",), ("second",)])

    with pytest.raises(RuntimeError, match="cannot safely apply"):
        update_project_progress(
            cursor,
            {"project": "project-type-id"},
            "account-id",
            {
                "projectId": "north",
                "overallActualProgress": 30,
                "lastProgressUpdateDate": "2027-02-12",
            },
        )
