import sys
from pathlib import Path
from unittest.mock import patch

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.push_to_api import (
    HIERARCHY,
    BluConnApi,
    ApiUploadError,
    _multipart_body,
    _rows_to_xlsx,
    upload_entities,
    write_api_report,
)


class FakeBluConnApi:
    batch_size = 2

    def __init__(self, existing=None):
        self.types = []
        self.type_payloads = []
        self.relationships = []
        self.entities = existing or {}
        self.uploads = []
        self.next_id = 1

    def _account_path(self, suffix):
        return suffix

    def list_entity_types(self):
        return list(self.types)

    def post_json(self, path, payload):
        if path == "/entity-types":
            self.type_payloads.append(dict(payload))
            entity_type = {
                "id": f"type-{payload['slug']}",
                "slug": payload["slug"],
                "name": payload["name"],
            }
            self.types.append(entity_type)
            return entity_type
        if path == "/entity-type-relationships":
            self.relationships.append(payload)
            return payload
        raise AssertionError(f"Unexpected API path: {path}")

    def list_relationships(self):
        return list(self.relationships)

    def list_entities(self, entity_type_id, **_kwargs):
        return list(self.entities.get(entity_type_id, []))

    def bulk_upload(self, entity_type_id, parent_id, rows, slug, batch_number):
        self.uploads.append({
            "entity_type_id": entity_type_id,
            "parent_id": parent_id,
            "rows": list(rows),
            "slug": slug,
            "batch_number": batch_number,
        })
        ids = []
        stored = self.entities.setdefault(entity_type_id, [])
        for row in rows:
            remote_id = f"remote-{self.next_id}"
            self.next_id += 1
            ids.append(remote_id)
            stored.append({"id": remote_id, "data": dict(row)})
        return {
            "entitiesCreated": len(rows),
            "entitiesFailed": 0,
            "entityIds": ids,
        }


def _hierarchy_data():
    return {
        "project": [{"projectId": "p1"}],
        "boq": [
            {"boqId": "b1", "projectId": "p1"},
            {"boqId": "b2", "projectId": "p1"},
        ],
        "activity": [{"activityId": "a1", "boqId": "b1", "projectId": "p1"}],
        "plan": [{"planId": "plan1", "activityId": "a1", "projectId": "p1"}],
    }


def test_upload_entities_creates_types_relationships_and_uploads_parent_first():
    api = FakeBluConnApi()

    stats, duplicates, failures = upload_entities(api, _hierarchy_data())

    assert [item["slug"] for item in api.uploads] == [
        "project", "boq", "activity", "plan",
    ]
    assert api.uploads[1]["parent_id"] == "remote-1"
    assert len(api.uploads[1]["rows"]) == 2
    assert api.uploads[2]["parent_id"] in {"remote-2", "remote-3"}
    assert api.uploads[3]["parent_id"] == "remote-4"
    assert stats["project"]["inserted"] == 1
    assert stats["boq"]["inserted"] == 2
    assert stats["activity"]["inserted"] == 1
    assert stats["plan"]["inserted"] == 1
    assert duplicates == []
    assert failures == []
    assert len(api.relationships) == sum(
        1 for _, _, parent, _, _ in HIERARCHY if parent is not None
    )
    required_parent_by_slug = {
        payload["slug"]: payload["isParentRequired"]
        for payload in api.type_payloads
    }
    assert required_parent_by_slug["engineering-daily-logs"] is True
    assert required_parent_by_slug["labour-daily-logs"] is False
    assert required_parent_by_slug["material-procurement-request"] is True


def test_upload_entities_skips_existing_and_duplicate_json_ids():
    existing = {
        "type-project": [
            {"id": "existing-project", "data": {"projectId": "p1"}},
        ],
    }
    api = FakeBluConnApi(existing)
    data = _hierarchy_data()
    data["project"].append({"projectId": "p1", "projectName": "Duplicate"})

    stats, duplicates, failures = upload_entities(api, data)

    assert not any(item["slug"] == "project" for item in api.uploads)
    assert api.uploads[0]["parent_id"] == "existing-project"
    assert stats["project"]["already_in_db"] == 1
    assert stats["project"]["duplicate_in_json"] == 1
    assert len(duplicates) == 2
    assert failures == []


def test_duplicate_ids_are_scoped_by_project_and_parent_refs_use_matching_project():
    api = FakeBluConnApi()
    data = {
        "project": [
            {"projectId": "p1"},
            {"projectId": "p2"},
        ],
        "boq": [
            {"boqId": "same-boq", "projectId": "p1"},
            {"boqId": "same-boq", "projectId": "p2"},
        ],
    }

    stats, duplicates, failures = upload_entities(api, data)

    project_ids = {
        row["data"]["projectId"]: row["id"]
        for row in api.entities["type-project"]
    }
    boq_uploads = [item for item in api.uploads if item["slug"] == "boq"]
    assert len(boq_uploads) == 2
    assert {item["parent_id"] for item in boq_uploads} == set(project_ids.values())
    assert stats["boq"]["inserted"] == 2
    assert duplicates == []
    assert failures == []


def test_upload_entities_reports_parentless_rows_without_sending_them():
    api = FakeBluConnApi()
    data = {
        "boq": [{"boqId": "b1", "projectId": "missing-project"}],
    }

    stats, _, failures = upload_entities(api, data)

    boq_uploads = [item for item in api.uploads if item["slug"] == "boq"]
    assert boq_uploads == []
    assert stats["boq"]["inserted"] == 0
    assert stats["boq"]["failed"] == 1
    assert "Required parent 'project' could not be resolved" in failures[0]["reason"]


def test_upload_entities_routes_labour_logs_to_parentless_type():
    api = FakeBluConnApi()
    stats, _, failures = upload_entities(api, {
        "engineering_daily_logs": [{
            "engineeringDailyLogId": "EDL1",
            "projectId": "p1",
            "activityId": "a1",
            "employeeRole": "unskilled",
            "assetId": "",
        }],
    })

    log_uploads = [
        item for item in api.uploads
        if item["slug"] == "labour_daily_logs"
    ]
    assert len(log_uploads) == 1
    assert log_uploads[0]["parent_id"] is None
    assert log_uploads[0]["rows"][0]["engineeringDailyLogId"] == "EDL1"
    assert stats["labour-daily-logs"]["inserted"] == 1
    assert stats["engineering-daily-logs"]["sent"] == 0
    assert failures == []
    labour_type = next(
        row for row in api.type_payloads if row["slug"] == "labour-daily-logs"
    )
    assert labour_type["parentMode"] == "NONE"
    assert labour_type["isParentRequired"] is False


def test_asset_operator_without_asset_is_not_misclassified_as_labour():
    api = FakeBluConnApi()
    stats, _, failures = upload_entities(api, {
        "engineering_daily_logs": [{
            "engineeringDailyLogId": "EDL1",
            "projectId": "p1",
            "employeeRole": "equipment_operator",
            "assetId": "",
        }],
    })

    assert not any(row["slug"] == "labour_daily_logs" for row in api.uploads)
    assert stats["engineering-daily-logs"]["failed"] == 1
    assert "Required parent 'asset' could not be resolved" in failures[0]["reason"]


def test_existing_daily_log_type_can_require_asset_for_equipment_logs():
    api = FakeBluConnApi()
    api.types.append({
        "id": "type-engineering-daily-logs",
        "slug": "engineering-daily-logs",
        "isParentRequired": True,
    })

    stats, _, failures = upload_entities(api, {
        "asset": [{"assetId": "asset-1", "projectId": "p1"}],
        "engineering_daily_logs": [{
            "engineeringDailyLogId": "EDL1",
            "projectId": "p1",
            "assetId": "asset-1",
        }],
    })

    log_uploads = [
        row for row in api.uploads if row["slug"] == "engineering_daily_logs"
    ]
    assert len(log_uploads) == 1
    assert log_uploads[0]["parent_id"] == "remote-1"
    assert stats["engineering-daily-logs"]["inserted"] == 1
    assert failures == []


def test_report_contains_uploaded_skipped_duplicate_and_failed_counts(tmp_path):
    json_path = tmp_path / "demo_actual_data.json"
    json_path.write_text("{}", encoding="utf-8")
    stats = {
        "project": {
            "sent": 4,
            "insert_attempted": 1,
            "inserted": 1,
            "already_in_db": 1,
            "duplicate_in_json": 1,
            "db_conflicts": 0,
            "failed": 1,
            "skipped": 0,
        }
    }
    report = write_api_report(
        str(json_path),
        stats,
        [{"entity_type": "project", "entity_id": "p1"}],
        [{"entity_type": "project", "entity_id": "p4", "reason": "API failed"}],
        "API_PARTIAL_FAILURE",
    )

    from openpyxl import load_workbook

    workbook = load_workbook(report, data_only=True)
    summary = dict(workbook["Summary"].iter_rows(min_row=2, values_only=True))
    workbook.close()
    assert summary["Records uploaded"] == 1
    assert summary["Records skipped"] == 2
    assert summary["Duplicate records"] == 1
    assert summary["Failed records"] == 1


def test_xlsx_and_multipart_helpers_preserve_nested_entity_values():
    content = _rows_to_xlsx([
        {
            "entityId": "e1",
            "items": [{"name": "cement"}],
            "date": "2026-12-10",
            "timestamp": "2026-12-10T08:30:00Z",
        },
    ])
    workbook = __import__("openpyxl").load_workbook(
        __import__("io").BytesIO(content), data_only=True
    )
    assert workbook.active.cell(2, 2).value == '[{"name": "cement"}]'
    assert workbook.active.cell(2, 3).value == "2026-12-10T00:00:00Z"
    assert workbook.active.cell(2, 4).value == "2026-12-10T08:30:00Z"
    workbook.close()

    body = _multipart_body(
        "boundary",
        {"entityTypeId": "type-1", "parentId": "parent-1"},
        "file",
        "batch.xlsx",
        b"xlsx-bytes",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    assert b'name="entityTypeId"\r\n\r\ntype-1' in body
    assert b'name="parentId"\r\n\r\nparent-1' in body
    assert b'filename="batch.xlsx"' in body
    assert body.endswith(b"--boundary--\r\n")


def test_failed_api_bulk_upload_is_counted_and_recorded():
    class FailedBatchApi(FakeBluConnApi):
        def bulk_upload(self, *args, **kwargs):
            raise ApiUploadError("API unavailable", status=503)

    api = FailedBatchApi()
    stats, _, failures = upload_entities(
        api,
        {"project": [{"projectId": "p1"}]},
    )
    assert stats["project"]["failed"] == 1
    assert failures[0]["reason"] == "API unavailable"


def test_api_retries_transient_get_failures():
    class Response:
        status = 200

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def read(self):
            return b'{"ok": true}'

    api = BluConnApi(
        "https://example.invalid",
        "account",
        "Bearer token",
        max_retries=1,
        sleep=lambda _delay: None,
    )
    from urllib.error import HTTPError
    from email.message import Message

    transient = HTTPError(
        "https://example.invalid/resource", 503, "temporary", Message(), None
    )
    with patch(
        "scripts.push_to_api.urlopen",
        side_effect=[transient, Response()],
    ) as request:
        assert api._request("GET", "/resource") == {"ok": True}
        assert request.call_count == 2


def test_list_entities_uses_bluconn_cursor_pagination():
        api = BluConnApi("https://example.invalid", "account", "token")
        requested_paths = []
        responses = iter([
            {
                "entities": [{"id": "remote-1"}],
                "links": {"nextCursor": "cursor-2"},
            },
            {"entities": [{"id": "remote-2"}], "links": {"nextCursor": None}},
        ])

        def get(path):
            requested_paths.append(path)
            return next(responses)

        api.get = get

        assert api.list_entities("type-project") == [
            {"id": "remote-1"},
            {"id": "remote-2"},
        ]
        assert "page%5Bcursor%5D=cursor-2" in requested_paths[1]
        assert "page%5Bnumber%5D" not in requested_paths[1]


def test_list_entities_stops_when_all_target_ids_are_found():
    api = BluConnApi("https://example.invalid", "account", "token")
    requested_paths = []

    def get(path):
        requested_paths.append(path)
        return {
            "entities": [
                {"id": "remote-1", "data": {"projectId": "p1", "projectCode": "P-1"}},
            ],
            "links": {"nextCursor": "cursor-2"},
        }

    api.get = get

    entities = api.list_entities(
        "type-project",
        target_keys={("p1", "P-1")},
        id_field="projectCode",
    )

    assert [row["id"] for row in entities] == ["remote-1"]
    assert len(requested_paths) == 1


def test_list_entities_rejects_repeated_cursor():
        api = BluConnApi("https://example.invalid", "account", "token")
        responses = iter([
            {"entities": [{"id": "remote-1"}], "links": {"nextCursor": "cursor-2"}},
            {"entities": [{"id": "remote-2"}], "links": {"nextCursor": "cursor-2"}},
        ])
        api.get = lambda _path: next(responses)

        with pytest.raises(ApiUploadError, match="repeated an entity-list cursor"):
            api.list_entities("type-project")


def test_api_retries_tls_handshake_timeout_for_write():
    from urllib.error import URLError

    class Response:
            status = 200

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            def read(self):
                return b'{"ok": true}'

    api = BluConnApi(
            "https://example.invalid",
            "account",
            "token",
            max_retries=1,
            sleep=lambda _delay: None,
    )
    handshake_timeout = URLError("The handshake operation timed out")
    with patch(
            "scripts.push_to_api.urlopen",
            side_effect=[handshake_timeout, Response()],
    ) as request:
            assert api._request(
                "POST", "/resource", b"body", "application/json", retry_writes=True
            ) == {"ok": True}
            assert request.call_count == 2


def test_api_retries_disconnected_get_but_not_write():
    import http.client

    class Response:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def read(self):
            return b'{"ok": true}'

    api = BluConnApi(
        "https://example.invalid",
        "account",
        "token",
        max_retries=1,
        sleep=lambda _delay: None,
    )
    disconnected = http.client.RemoteDisconnected(
        "Remote end closed connection without response"
    )
    with patch(
        "scripts.push_to_api.urlopen",
        side_effect=[disconnected, Response()],
    ) as request:
        assert api._request("GET", "/resource") == {"ok": True}
        assert request.call_count == 2

    with patch(
        "scripts.push_to_api.urlopen",
        side_effect=http.client.RemoteDisconnected(
            "Remote end closed connection without response"
        ),
    ) as request:
        with pytest.raises(ApiUploadError, match="connection ended"):
            api._request(
                "POST", "/resource", b"body", "application/json", retry_writes=True
            )
        assert request.call_count == 1


def test_api_authentication_failure_is_not_retried():
    from urllib.error import HTTPError
    from email.message import Message

    api = BluConnApi(
        "https://example.invalid",
        "account",
        "token",
        max_retries=3,
        sleep=lambda _delay: None,
    )
    headers = Message()
    error = HTTPError(
        "https://example.invalid/resource", 401, "unauthorized", headers, None
    )
    with patch("scripts.push_to_api.urlopen", side_effect=error) as request:
        with pytest.raises(ApiUploadError, match="authentication/authorization failed"):
            api._request("GET", "/resource")
    assert request.call_count == 1
