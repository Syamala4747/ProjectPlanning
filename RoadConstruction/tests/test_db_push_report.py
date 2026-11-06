from pathlib import Path
import sys

from openpyxl import load_workbook
import psycopg2
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.push_to_db import (
    insert_entities,
    unsupported_entity_records,
    write_push_report,
)


class FakeCursor:
    def __init__(self, existing_rows=()):
        self._selected = list(existing_rows)
        self._inserted_id = None

    def execute(self, query, parameters):
        if query.lstrip().startswith("SELECT"):
            return
        else:
            self._inserted_id = parameters[0]

    def fetchall(self):
        return self._selected

    def fetchone(self):
        return (self._inserted_id,)


class FailingInsertCursor(FakeCursor):
    def execute(self, query, parameters):
        if query.lstrip().startswith("INSERT"):
            raise psycopg2.IntegrityError("test constraint failure")
        super().execute(query, parameters)


def test_insert_entities_skips_duplicate_business_ids_in_json():
    rows = [
        {"projectId": "project-a", "projectName": "Project A"},
        {"projectId": "project-a", "projectName": "Duplicate Project A"},
    ]

    stats, duplicates = insert_entities(
        FakeCursor(),
        {"project": rows},
        {"project": "entity-type-project"},
        "account-a",
    )

    assert stats["project"]["sent"] == 2
    assert stats["project"]["insert_attempted"] == 1
    assert stats["project"]["inserted"] == 1
    assert stats["project"]["duplicate_in_json"] == 1
    assert len(duplicates) == 1
    assert duplicates[0]["duplicate_kind"] == "repeated business ID in JSON"


def test_insert_entities_skips_records_already_in_database():
    rows = [{"projectId": "project-a", "projectName": "Project A"}]
    cursor = FakeCursor([("existing-db-id", "existing/path", "project-a")])

    stats, duplicates = insert_entities(
        cursor,
        {"project": rows},
        {"project": "entity-type-project"},
        "account-a",
    )

    assert stats["project"]["sent"] == 1
    assert stats["project"]["insert_attempted"] == 0
    assert stats["project"]["already_in_db"] == 1
    assert stats["project"]["inserted"] == 0
    assert duplicates[0]["duplicate_kind"] == "already exists in database; skipped"


def test_insert_error_is_reported_with_record_and_counted():
    stats = {}
    with pytest.raises(RuntimeError, match="project record 'project-a'"):
        insert_entities(
            FailingInsertCursor(),
            {"project": [{"projectId": "project-a"}]},
            {"project": "entity-type-project"},
            "account-a",
            stats,
            [],
        )

    assert stats["project"]["insert_attempted"] == 1
    assert stats["project"]["failed"] == 1


def test_push_report_marks_rolled_back_records_as_not_committed(tmp_path):
    json_file = tmp_path / "sample_actual_data.json"
    json_file.write_text("{}", encoding="utf-8")
    stats = {
        "project": {
            "sent": 2,
            "insert_attempted": 1,
            "inserted": 1,
            "already_in_db": 0,
            "duplicate_in_json": 1,
            "db_conflicts": 0,
        }
    }

    report_path = write_push_report(
        str(json_file),
        stats,
        [],
        "ROLLED_BACK",
        "insert failed",
    )

    workbook = load_workbook(report_path, data_only=True)
    summary = dict(workbook["Summary"].iter_rows(min_row=2, values_only=True))
    entity_counts = list(workbook["Entity Counts"].values)
    workbook.close()

    assert summary["Records in JSON (sent for processing)"] == 2
    assert summary["New records committed"] == 0
    assert summary["Commit status"] == "ROLLED_BACK"
    assert entity_counts[1][3] == 1
    assert entity_counts[1][4] == 0


def test_unhandled_entity_records_are_reported_but_empty_collections_are_allowed():
    assert unsupported_entity_records({
        "project": [{"projectId": "p1"}],
        "store": [],
        "store_inventory": [{"material": "cement"}],
    }) == {"store_inventory": 1}
