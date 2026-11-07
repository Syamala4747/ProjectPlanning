"""Upload construction entities to BluConn through its API."""
from __future__ import annotations

import argparse
import http.client
import json
import os
import random
import sys
import time
import uuid
from collections import defaultdict
from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

import openpyxl

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))


# (API slug, JSON slug, parent API slug, entity ID field, parent ID field)
HIERARCHY = [
    ("project", "project", None, "projectId", None),
    ("assumption", "assumption", None, "assumptionId", None),
    ("asset", "asset", None, "assetId", None),
    ("material", "material", None, "materialId", None),
    ("subcontractor", "subcontractor", None, "subContractorId", None),
    ("logistics", "logistics", None, "logisticsId", None),
    ("boq", "boq", "project", "boqId", "projectId"),
    ("activity", "activity", "boq", "activityId", "boqId"),
    ("plan", "plan", "activity", "planId", "activityId"),
    ("supervisor-update", "supervisor_update", "activity", "supervisorUpdateId", "activityId"),
    ("engineering-daily-logs", "engineering_daily_logs", "asset", "engineeringDailyLogId", "assetId"),
    ("labour-daily-logs", "labour_daily_logs", None, "engineeringDailyLogId", None),
    ("material-procurement-request", "material_procurement_request", "material", "materialProcurementRequestId", "materialId"),
    ("material-issue-request", "material_issue_request", "material", "materialIssueRequestId", "materialId"),
    ("material-transaction", "material_transaction", "material", "materialTransactionId", "materialId"),
]

# Material procurement and equipment logs require actual parent instances in
# the BluConn account. Worker-only logs use a separate parentless type.
PARENT_REQUIRED_TYPES = {
    "engineering-daily-logs",
    "material-procurement-request",
}
ASSET_OPERATOR_ROLES = {"driver", "equipment_operator"}

ENTITY_NAMES = {
    "project": "Project",
    "assumption": "Assumption",
    "asset": "Asset",
    "material": "Material",
    "subcontractor": "SubContractor",
    "logistics": "Logistics",
    "boq": "BOQ",
    "activity": "Activity",
    "plan": "Plan",
    "supervisor-update": "Supervisor Update",
    "engineering-daily-logs": "Engineering Daily Logs",
    "labour-daily-logs": "Labour Daily Logs",
    "material-procurement-request": "Material Procurement Request",
    "material-issue-request": "Material Issue Request",
    "material-transaction": "Material Transaction",
}

_TRANSIENT_STATUSES = {429, 500, 502, 503, 504}
_RETRYABLE_WRITE_STATUSES = {429}


class ApiUploadError(RuntimeError):
    def __init__(
        self,
        message: str,
        status: Optional[int] = None,
        response_body: str = "",
    ):
        super().__init__(message)
        self.status = status
        self.response_body = response_body


def _load_api_environment() -> None:
    """Load only the API configuration keys from the local .env file."""
    env_path = _ROOT / ".env"
    if not env_path.is_file():
        return
    allowed = {
        "BLUCONN_API_URL",
        "BLUCONN_ACCOUNT_ID",
        "BLUCONN_AUTH_TOKEN",
    }
    for line in env_path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        if key not in allowed or key in os.environ:
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        os.environ[key] = value


class BluConnApi:
    def __init__(
        self,
        base_url: str,
        account_id: str,
        auth_token: str,
        timeout: int = 60,
        max_retries: int = 3,
        batch_size: int = 500,
        sleep=time.sleep,
    ):
        self.base_url = base_url.rstrip("/")
        self.account_id = quote(account_id, safe="")
        token = auth_token.strip()
        if token.lower().startswith("bearer "):
            token = token[7:].strip()
        self.auth_token = token
        self.timeout = timeout
        self.max_retries = max_retries
        self.batch_size = batch_size
        self._sleep = sleep

    def _account_path(self, suffix: str) -> str:
        return f"/accounts/{self.account_id}{suffix}"

    def _request(
        self,
        method: str,
        path: str,
        body: Optional[bytes] = None,
        content_type: Optional[str] = None,
        retry_writes: bool = False,
    ) -> Any:
        url = f"{self.base_url}{path}"
        for attempt in range(self.max_retries + 1):
            headers = {
                "Authorization": f"Bearer {self.auth_token}",
                "Accept": "application/json",
            }
            if content_type:
                headers["Content-Type"] = content_type
            request = Request(url, data=body, headers=headers, method=method)
            try:
                with urlopen(request, timeout=self.timeout) as response:
                    raw = response.read()
                    if not raw:
                        return {}
                    try:
                        return json.loads(raw.decode("utf-8"))
                    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
                        raise ApiUploadError(
                            f"{method} {path} returned invalid JSON."
                        ) from exc
            except HTTPError as exc:
                response_body = exc.read().decode("utf-8", errors="replace")[:2000]
                if exc.code in (401, 403):
                    raise ApiUploadError(
                        f"BluConn authentication/authorization failed (HTTP {exc.code}); "
                        "verify the API token and account ID.",
                        status=exc.code,
                    ) from exc
                retryable = (
                    exc.code in _TRANSIENT_STATUSES
                    and (method == "GET" or (
                        retry_writes and exc.code in _RETRYABLE_WRITE_STATUSES
                    ))
                )
                if retryable and attempt < self.max_retries:
                    self._wait_before_retry(attempt, exc.headers.get("Retry-After"))
                    continue
                message = f"{method} {path} failed with HTTP {exc.code}"
                if response_body:
                    message += f": {response_body}"
                raise ApiUploadError(
                    message, status=exc.code, response_body=response_body
                ) from exc
            except URLError as exc:
                handshake_timeout = (
                    "handshake operation timed out" in str(exc.reason).lower()
                )
                retryable = method == "GET" or (retry_writes and handshake_timeout)
                if retryable and attempt < self.max_retries:
                    self._wait_before_retry(attempt, None)
                    continue
                raise ApiUploadError(
                    f"{method} {path} could not reach BluConn: {exc.reason}"
                ) from exc
            except TimeoutError as exc:
                if method == "GET" and attempt < self.max_retries:
                    self._wait_before_retry(attempt, None)
                    continue
                raise ApiUploadError(f"{method} {path} timed out.") from exc
            except http.client.HTTPException as exc:
                if method == "GET" and attempt < self.max_retries:
                    self._wait_before_retry(attempt, None)
                    continue
                raise ApiUploadError(
                    f"{method} {path} connection ended before a complete "
                    f"HTTP response: {exc}"
                ) from exc
        raise ApiUploadError(f"{method} {path} exhausted its retry attempts.")

    def _wait_before_retry(self, attempt: int, retry_after: Optional[str]) -> None:
        try:
            delay = min(30.0, max(0.0, float(retry_after))) if retry_after else None
        except ValueError:
            delay = None
        if delay is None:
            delay = min(8.0, (2 ** attempt) + random.random())
        self._sleep(delay)

    def get(self, path: str) -> Any:
        return self._request("GET", path)

    def post_json(self, path: str, payload: Mapping[str, Any]) -> Any:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        return self._request(
            "POST", path, body, "application/json", retry_writes=True
        )

    def list_entity_types(self) -> List[Dict[str, Any]]:
        response = self.get(self._account_path("/entities/entity-types"))
        rows = response.get("entityTypes") if isinstance(response, dict) else response
        if not isinstance(rows, list):
            raise ApiUploadError("BluConn entity-type response has no entityTypes list.")
        return rows

    def list_relationships(self) -> List[Dict[str, Any]]:
        response = self.get(self._account_path("/entity-type-relationships"))
        rows = (
            response.get("relationships", response)
            if isinstance(response, dict)
            else response
        )
        if not isinstance(rows, list):
            raise ApiUploadError("BluConn relationship response is not a list.")
        return rows

    def list_entities(
        self,
        entity_type_id: str,
        target_keys: Optional[set[Tuple[str, str]]] = None,
        id_field: str = "",
        project_id_field: str = "projectId",
    ) -> List[Dict[str, Any]]:
        targets = set(target_keys or ())
        if target_keys is not None and not targets:
            return []

        all_rows: List[Dict[str, Any]] = []
        seen_pages: set[Tuple[str, ...]] = set()
        seen_cursors: set[str] = set()
        found_keys: set[Tuple[str, str]] = set()
        cursor = ""
        page_size = 50
        for page_number in range(1, 10001):
            if page_number == 1 or page_number % 10 == 0:
                print(
                    f"    Checking BluConn entities {entity_type_id}: "
                    f"page {page_number}, matched {len(found_keys)}/{len(targets)} IDs",
                    flush=True,
                )
            params = {
                "entityTypeId": entity_type_id,
                "page[size]": page_size,
            }
            if cursor:
                params["page[cursor]"] = cursor
            query = urlencode(params)
            response = self.get(
                self._account_path(f"/entities?{query}")
            )
            rows = response.get("entities") if isinstance(response, dict) else response
            if not isinstance(rows, list):
                raise ApiUploadError("BluConn entity-list response has no entities list.")
            if not rows:
                break

            signature = tuple(str(row.get("id", "")) for row in rows)
            if signature in seen_pages:
                raise ApiUploadError(
                    "BluConn repeated an entity-list page; duplicate detection "
                    "cannot safely continue."
                )
            seen_pages.add(signature)
            all_rows.extend(rows)
            for row in rows:
                row_data = _row_data(row)
                local_id = str(row_data.get(id_field) or "")
                project_id = str(row_data.get(project_id_field) or "")
                if (project_id, local_id) in targets:
                    found_keys.add((project_id, local_id))
            if targets and targets.issubset(found_keys):
                break

            links = response.get("links", {}) if isinstance(response, dict) else {}
            next_cursor = (
                links.get("nextCursor")
                if isinstance(links, dict)
                else None
            )
            if not next_cursor:
                break
            next_cursor = str(next_cursor)
            if next_cursor == cursor or next_cursor in seen_cursors:
                raise ApiUploadError(
                    "BluConn repeated an entity-list cursor; duplicate detection "
                    "cannot safely continue."
                )
            seen_cursors.add(next_cursor)
            cursor = next_cursor
        else:
            raise ApiUploadError(
                "BluConn entity listing exceeded 10,000 pages; refusing "
                "an incomplete duplicate check."
            )
        return all_rows

    def bulk_upload(
        self,
        entity_type_id: str,
        parent_id: Optional[str],
        rows: Sequence[Mapping[str, Any]],
        slug: str,
        batch_number: int,
    ) -> Dict[str, Any]:
        workbook_bytes = _rows_to_xlsx(rows)
        boundary = f"----BluConnBoundary{uuid.uuid4().hex}"
        payload = _multipart_body(
            boundary,
            {
                "entityTypeId": entity_type_id,
                **({"parentId": parent_id} if parent_id else {}),
            },
            "file",
            f"{slug}_{batch_number:04d}.xlsx",
            workbook_bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response = self._request(
            "POST",
            self._account_path("/entities/bulk-upload"),
            payload,
            f"multipart/form-data; boundary={boundary}",
            retry_writes=True,
        )
        if not isinstance(response, dict):
            raise ApiUploadError(
                f"BluConn bulk-upload response for {slug} is not an object."
            )
        return response


def _infer_property(value: Any, title: str) -> Dict[str, str]:
    if isinstance(value, bool):
        return {"type": "boolean", "title": title}
    if isinstance(value, (int, float)):
        return {"type": "number", "title": title}
    if isinstance(value, (dict, list)):
        return {"type": "string", "title": title}
    text = str(value or "")
    if len(text) >= 10 and text[4:5] == "-" and text[7:8] == "-":
        return {"type": "string", "format": "date-time", "title": title}
    return {"type": "string", "title": title}


def _schema_for_rows(
    rows: Sequence[Mapping[str, Any]], id_field: str
) -> Dict[str, Any]:
    properties: Dict[str, Any] = {}
    for row in rows:
        for key, value in row.items():
            if key not in properties or value not in (None, ""):
                if value not in (None, ""):
                    properties[key] = _infer_property(value, key)
                elif key not in properties:
                    properties[key] = _infer_property(value, key)
    return {
        "$schema": "http://json-schema.org/draft-07/schema#",
        "type": "object",
        "properties": properties,
        "required": [id_field] if id_field in properties else [],
    }


def _ensure_entity_types(
    api: BluConnApi,
    data: Mapping[str, Sequence[Mapping[str, Any]]],
) -> Dict[str, Dict[str, Any]]:
    account_path = api._account_path
    existing_rows = api.list_entity_types()
    existing_by_slug: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for entity_type in existing_rows:
        slug = entity_type.get("slug")
        if slug:
            existing_by_slug[str(slug)].append(entity_type)

    type_map: Dict[str, Dict[str, Any]] = {}
    for api_slug, json_slug, parent_slug, id_field, _ in HIERARCHY:
        aliases = {api_slug, json_slug}
        matches = [
            item
            for alias in aliases
            for item in existing_by_slug.get(alias, [])
        ]
        unique_matches = {str(item.get("id")): item for item in matches if item.get("id")}
        if len(unique_matches) > 1:
            raise ApiUploadError(
                f"Both canonical and legacy entity types exist for {api_slug!r}; "
                "resolve the duplicate types before uploading."
            )
        if unique_matches:
            entity_type = next(iter(unique_matches.values()))
            type_map[api_slug] = entity_type
            continue

        payload: Dict[str, Any] = {
            "name": ENTITY_NAMES[api_slug],
            "slug": api_slug,
            "parentMode": "RESTRICTED" if parent_slug else "NONE",
            "isParentRequired": api_slug in PARENT_REQUIRED_TYPES,
            "jsonSchema": _schema_for_rows(data.get(json_slug, []), id_field),
            "searchableFields": [id_field],
        }
        parent_type = type_map.get(parent_slug) if parent_slug else None
        if parent_type:
            payload["allowedParentTypeIds"] = [parent_type["id"]]
        try:
            response = api.post_json(account_path("/entity-types"), payload)
            entity_type = _unwrap_entity_type(response)
            if not entity_type.get("id"):
                raise ApiUploadError(
                    f"BluConn did not return an ID when creating {api_slug!r}."
                )
            type_map[api_slug] = entity_type
        except ApiUploadError as exc:
            if exc.status != 409:
                raise
            refreshed = api.list_entity_types()
            entity_type = next(
                (
                    item for item in refreshed
                    if item.get("slug") in aliases and item.get("id")
                ),
                None,
            )
            if entity_type is None:
                raise ApiUploadError(
                    f"BluConn reported a conflict creating {api_slug!r}, but the "
                    "existing type could not be resolved."
                ) from exc
            type_map[api_slug] = entity_type

    existing_relations = api.list_relationships()
    relation_keys = {
        (str(row.get("parentTypeId")), str(row.get("childTypeId")))
        for row in existing_relations
        if isinstance(row, dict)
    }
    for api_slug, _, parent_slug, _, _ in HIERARCHY:
        if parent_slug is None:
            continue
        parent_id = str(type_map[parent_slug]["id"])
        child_id = str(type_map[api_slug]["id"])
        if (parent_id, child_id) in relation_keys:
            continue
        try:
            api.post_json(
                account_path("/entity-type-relationships"),
                {"parentTypeId": parent_id, "childTypeId": child_id},
            )
        except ApiUploadError as exc:
            if exc.status != 409:
                raise
            refreshed = api.list_relationships()
            if not any(
                str(row.get("parentTypeId")) == parent_id
                and str(row.get("childTypeId")) == child_id
                for row in refreshed
            ):
                raise ApiUploadError(
                    f"BluConn reported a relationship conflict for "
                    f"{parent_slug!r} -> {api_slug!r}, but it could not be verified."
                ) from exc
        relation_keys.add((parent_id, child_id))
    return type_map


def _unwrap_entity_type(response: Any) -> Dict[str, Any]:
    if isinstance(response, dict):
        result = response.get("entityType", response)
        if isinstance(result, dict):
            return result
    return {}


def _row_data(entity: Mapping[str, Any]) -> Dict[str, Any]:
    value = entity.get("data", entity)
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            return {}
    return value if isinstance(value, dict) else {}


def _existing_id_map(
    api: BluConnApi,
    entity_type_id: str,
    id_field: str,
    api_slug: str,
    project_id_field: str,
    rows: Sequence[Mapping[str, Any]],
    duplicates: List[Dict[str, str]],
) -> Dict[Tuple[str, str], str]:
    mapping: Dict[Tuple[str, str], str] = {}
    target_keys = {
        (
            str(row.get(project_id_field) or ""),
            str(row.get(id_field) or ""),
        )
        for row in rows
        if row.get(id_field) not in (None, "")
    }
    if not target_keys:
        return mapping
    print(
        f"  Checking existing {api_slug} records "
        f"({len(target_keys)} distinct workbook IDs)...",
        flush=True,
    )
    for entity in api.list_entities(
        entity_type_id,
        target_keys=target_keys,
        id_field=id_field,
        project_id_field=project_id_field,
    ):
        remote_id = str(entity.get("id") or "")
        row_data = _row_data(entity)
        local_id = str(row_data.get(id_field) or "")
        project_id = str(row_data.get(project_id_field) or "")
        if not remote_id or not local_id:
            continue
        local_key = (project_id, local_id)
        if local_key in mapping:
            duplicates.append({
                "entity_type": api_slug,
                "project_id": project_id,
                "entity_id": local_id,
                "duplicate_kind": "multiple matching records already exist in API",
                "database_entity_ids": f"{mapping[local_key]}, {remote_id}",
            })
            continue
        mapping[local_key] = remote_id
    return mapping


def _rows_to_xlsx(rows: Sequence[Mapping[str, Any]]) -> bytes:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    headers: List[str] = []
    seen: set[str] = set()
    for row in rows:
        for key in row:
            if key not in seen:
                seen.add(key)
                headers.append(key)
    sheet.append(headers)
    for row in rows:
        values = []
        for key in headers:
            value = row.get(key)
            if isinstance(value, (dict, list)):
                value = json.dumps(value, ensure_ascii=False, default=str)
            elif value is None:
                value = ""
            elif isinstance(value, str) and len(value) == 10:
                try:
                    if date.fromisoformat(value).isoformat() == value:
                        value = f"{value}T00:00:00Z"
                except ValueError:
                    pass
            values.append(value)
        sheet.append(values)
    buffer = BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


def _multipart_body(
    boundary: str,
    fields: Mapping[str, str],
    file_field: str,
    filename: str,
    file_bytes: bytes,
    file_content_type: str,
) -> bytes:
    delimiter = f"--{boundary}\r\n".encode("ascii")
    chunks = []
    for key, value in fields.items():
        chunks.extend([
            delimiter,
            f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode("utf-8"),
            str(value).encode("utf-8"),
            b"\r\n",
        ])
    chunks.extend([
        delimiter,
        (
            f'Content-Disposition: form-data; name="{file_field}"; '
            f'filename="{filename}"\r\n'
            f"Content-Type: {file_content_type}\r\n\r\n"
        ).encode("utf-8"),
        file_bytes,
        b"\r\n",
        f"--{boundary}--\r\n".encode("ascii"),
    ])
    return b"".join(chunks)


def _new_entity_stats(sent: int) -> Dict[str, int]:
    return {
        "sent": sent,
        "insert_attempted": 0,
        "inserted": 0,
        "already_in_db": 0,
        "duplicate_in_json": 0,
        "db_conflicts": 0,
        "failed": 0,
        "skipped": 0,
    }


def upload_entities(
    api: BluConnApi,
    data: Mapping[str, Sequence[Mapping[str, Any]]],
) -> Tuple[Dict[str, Dict[str, int]], List[Dict[str, str]], List[Dict[str, str]]]:
    """Upload records in parent-before-child order, preserving API parent IDs."""
    prepared_data = _prepare_api_entities(data)
    stats = {
        api_slug: _new_entity_stats(len(prepared_data.get(json_slug, [])))
        for api_slug, json_slug, _, _, _ in HIERARCHY
    }
    duplicates: List[Dict[str, str]] = []
    failures: List[Dict[str, str]] = []
    print("  Checking BluConn entity types and relationships...", flush=True)
    type_map = _ensure_entity_types(api, prepared_data)
    uploaded_ids: Dict[str, Dict[Tuple[str, str], str]] = {}

    for api_slug, json_slug, parent_slug, id_field, parent_field in HIERARCHY:
        rows = list(prepared_data.get(json_slug, []))
        current = stats[api_slug]
        entity_type = type_map[api_slug]
        project_ids = _existing_id_map(
            api,
            str(entity_type["id"]),
            id_field,
            api_slug,
            "projectId",
            rows,
            duplicates,
        )
        uploaded_ids[api_slug] = project_ids

        unique_rows: List[Dict[str, Any]] = []
        seen: set[Tuple[str, str]] = set()
        for row in rows:
            local_id = str(row.get(id_field) or "")
            project_id = str(row.get("projectId") or "")
            local_key = (project_id, local_id)
            if not local_id:
                current["failed"] += 1
                failures.append({
                    "entity_type": api_slug,
                    "entity_id": "",
                    "reason": f"Missing required business ID field {id_field}.",
                })
            elif local_key in seen:
                current["duplicate_in_json"] += 1
                duplicates.append({
                    "entity_type": api_slug,
                    "project_id": project_id,
                    "entity_id": local_id,
                    "duplicate_kind": "repeated business ID in JSON; skipped",
                    "database_entity_ids": project_ids.get(local_key, ""),
                })
            elif local_key in project_ids:
                current["already_in_db"] += 1
                duplicates.append({
                    "entity_type": api_slug,
                    "project_id": project_id,
                    "entity_id": local_id,
                    "duplicate_kind": "already exists in API; skipped",
                    "database_entity_ids": project_ids[local_key],
                })
                seen.add(local_key)
            else:
                seen.add(local_key)
                unique_rows.append(dict(row))

        groups: Dict[Optional[str], List[Dict[str, Any]]] = defaultdict(list)
        for row in unique_rows:
            if parent_slug is None:
                groups[None].append(row)
                continue
            local_parent_id = str(row.get(parent_field) or "")
            project_id = str(row.get("projectId") or "")
            remote_parent_id = uploaded_ids.get(parent_slug, {}).get(
                (project_id, local_parent_id)
            )
            if not remote_parent_id:
                if (
                    not local_parent_id
                    and api_slug not in PARENT_REQUIRED_TYPES
                ):
                    groups[None].append(row)
                    continue
                current["failed"] += 1
                failures.append({
                    "entity_type": api_slug,
                    "entity_id": local_id,
                    "reason": (
                        f"Required parent {parent_slug!r} could not be resolved "
                        f"from {parent_field}={local_parent_id!r}; record was not sent."
                    ),
                })
                continue
            groups[remote_parent_id].append(row)

        for parent_id, group_rows in groups.items():
            for offset in range(0, len(group_rows), api.batch_size):
                batch = group_rows[offset:offset + api.batch_size]
                current["insert_attempted"] += len(batch)
                batch_number = offset // api.batch_size + 1
                print(
                    f"  Uploading {api_slug} batch {batch_number} "
                    f"({len(batch)} records)...",
                    flush=True,
                )
                try:
                    response = api.bulk_upload(
                        str(entity_type["id"]),
                        parent_id,
                        batch,
                        api_slug.replace("-", "_"),
                        batch_number,
                    )
                except ApiUploadError as exc:
                    current["failed"] += len(batch)
                    failures.append({
                        "entity_type": api_slug,
                        "entity_id": ", ".join(
                            str(row.get(id_field) or "") for row in batch[:10]
                        ),
                        "reason": str(exc),
                    })
                    continue

                try:
                    created = _response_count(response, "entitiesCreated", 0)
                    failed = _response_count(response, "entitiesFailed", 0)
                    skipped = _response_count(
                        response,
                        "entitiesSkipped",
                        _response_count(response, "entitiesDuplicate", 0),
                    )
                except ApiUploadError as exc:
                    current["failed"] += len(batch)
                    failures.append({
                        "entity_type": api_slug,
                        "entity_id": ", ".join(
                            str(row.get(id_field) or "") for row in batch[:10]
                        ),
                        "reason": str(exc),
                    })
                    continue
                if created + failed + skipped > len(batch):
                    current["failed"] += len(batch)
                    failures.append({
                        "entity_type": api_slug,
                        "entity_id": ", ".join(
                            str(row.get(id_field) or "") for row in batch[:10]
                        ),
                        "reason": (
                            "BluConn response counts exceeded the number of "
                            "submitted records; batch counts are invalid."
                        ),
                    })
                    continue
                if created + failed + skipped < len(batch):
                    failed += len(batch) - created - failed - skipped
                    failures.append({
                        "entity_type": api_slug,
                        "entity_id": ", ".join(
                            str(row.get(id_field) or "")
                            for row in batch[:10]
                        ),
                        "reason": (
                            "API response counts did not account for every "
                            "record in the batch; unaccounted records marked failed."
                        ),
                    })
                current["inserted"] += created
                current["failed"] += failed
                current["skipped"] += skipped
                current["db_conflicts"] += skipped

                detail_rows = response.get("failedEntities") or response.get("errors") or []
                if isinstance(detail_rows, list):
                    for item in detail_rows:
                        failures.append({
                            "entity_type": api_slug,
                            "entity_id": str(
                                item.get(id_field)
                                or item.get("entityId")
                                or item.get("id")
                                or ""
                            ) if isinstance(item, dict) else "",
                            "reason": (
                                json.dumps(item, ensure_ascii=False, default=str)
                                if isinstance(item, dict) else str(item)
                            ),
                        })

                returned_ids = response.get("entityIds", [])
                if not isinstance(returned_ids, list):
                    returned_ids = []
                if (
                    created == len(batch)
                    and failed == 0
                    and skipped == 0
                    and len(returned_ids) == len(batch)
                ):
                    for row, remote_id in zip(batch, returned_ids):
                        if remote_id not in (None, ""):
                            project_ids[(
                                str(row.get("projectId") or ""),
                                str(row[id_field]),
                            )] = str(remote_id)
                elif created:
                    # A partial response cannot safely map server IDs to input
                    # rows by position. Resolve successful rows by business ID.
                    refreshed = _existing_id_map(
                        api,
                        str(entity_type["id"]),
                        id_field,
                        api_slug,
                        "projectId",
                        unique_rows,
                        duplicates,
                    )
                    project_ids.update(refreshed)

        # If the API accepted every item but omitted entityIds, resolve IDs so
        # later child batches still receive the correct parentId.
        if current["inserted"] and any(
            (
                str(row.get("projectId") or ""),
                str(row.get(id_field) or ""),
            ) not in project_ids
            for row in unique_rows
        ):
            refreshed = _existing_id_map(
                api,
                str(entity_type["id"]),
                id_field,
                api_slug,
                "projectId",
                unique_rows,
                duplicates,
            )
            project_ids.update(refreshed)

    return stats, duplicates, failures


def _prepare_api_entities(
    data: Mapping[str, Sequence[Mapping[str, Any]]],
) -> Dict[str, Sequence[Mapping[str, Any]]]:
    """Keep equipment EDLs asset-parented; route labour logs to their own type."""
    prepared_data: Dict[str, Sequence[Mapping[str, Any]]] = dict(data)
    engineering_logs = list(data.get("engineering_daily_logs", []))
    prepared_data["engineering_daily_logs"] = [
        row
        for row in engineering_logs
        if row.get("assetId")
        or str(row.get("employeeRole") or "").strip().lower()
        in ASSET_OPERATOR_ROLES
    ]
    prepared_data["labour_daily_logs"] = [
        *data.get("labour_daily_logs", []),
        *(
            row for row in engineering_logs
            if not row.get("assetId")
            and str(row.get("employeeRole") or "").strip().lower()
            not in ASSET_OPERATOR_ROLES
        ),
    ]
    return prepared_data


def _response_count(response: Mapping[str, Any], key: str, default: int) -> int:
    value = response.get(key, default)
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise ApiUploadError(
            f"BluConn bulk-upload response has invalid {key}: {value!r}."
        ) from exc
    if result < 0:
        raise ApiUploadError(
            f"BluConn bulk-upload response has negative {key}: {result}."
        )
    return result


def write_api_report(
    json_path: str,
    stats: Mapping[str, Mapping[str, int]],
    duplicates: Sequence[Mapping[str, str]],
    failures: Sequence[Mapping[str, str]],
    status: str,
    error_message: str = "",
) -> str:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    source = Path(json_path).resolve()
    suffix = "_actual_data.json"
    stem = source.name[:-len(suffix)] if source.name.endswith(suffix) else source.stem
    report_path = source.with_name(f"{stem}_api_upload_report.xlsx")
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"

    sent = sum(values.get("sent", 0) for values in stats.values())
    attempted = sum(values.get("insert_attempted", 0) for values in stats.values())
    inserted = sum(values.get("inserted", 0) for values in stats.values())
    skipped = sum(
        values.get("already_in_db", 0)
        + values.get("duplicate_in_json", 0)
        + values.get("skipped", 0)
        for values in stats.values()
    )
    failed = sum(values.get("failed", 0) for values in stats.values())
    summary.append(["BluConn API Upload Report", ""])
    summary.append(["Entity JSON", source.name])
    summary.append(["Upload status", status])
    summary.append(["Records in JSON", sent])
    summary.append(["Records attempted", attempted])
    summary.append(["Records uploaded", inserted])
    summary.append(["Records skipped", skipped])
    summary.append(["Duplicate records", len(duplicates)])
    summary.append(["Failed records", failed])
    if error_message:
        summary.append(["Error", error_message])

    counts = workbook.create_sheet("Entity Counts")
    counts.append([
        "Entity type", "JSON records", "Upload attempts", "Uploaded",
        "Already in API", "Duplicate JSON", "API duplicate/skipped",
        "Skipped (missing parent)", "Failed",
    ])
    for api_slug, values in stats.items():
        counts.append([
            api_slug,
            values.get("sent", 0),
            values.get("insert_attempted", 0),
            values.get("inserted", 0),
            values.get("already_in_db", 0),
            values.get("duplicate_in_json", 0),
            values.get("db_conflicts", 0),
            values.get("skipped", 0),
            values.get("failed", 0),
        ])

    duplicate_sheet = workbook.create_sheet("Duplicates")
    duplicate_sheet.append([
        "Entity type", "Project ID", "Entity ID", "Finding", "API entity IDs",
    ])
    for duplicate in duplicates:
        duplicate_sheet.append([
            duplicate.get("entity_type", ""),
            duplicate.get("project_id", ""),
            duplicate.get("entity_id", ""),
            duplicate.get("duplicate_kind", ""),
            duplicate.get("database_entity_ids", ""),
        ])
    if not duplicates:
        duplicate_sheet.append(["", "", "", "No duplicates detected", ""])

    failure_sheet = workbook.create_sheet("Failures")
    failure_sheet.append(["Entity type", "Entity ID", "Reason"])
    for failure in failures:
        failure_sheet.append([
            failure.get("entity_type", ""),
            failure.get("entity_id", ""),
            failure.get("reason", ""),
        ])
    if not failures:
        failure_sheet.append(["", "", "No failed records"])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill
        for column_cells in sheet.columns:
            width = max(
                (len(str(cell.value or "")) for cell in column_cells),
                default=10,
            )
            for cell in column_cells:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            sheet.column_dimensions[
                get_column_letter(column_cells[0].column)
            ].width = min(max(width + 2, 12), 72)
    workbook.save(report_path)
    workbook.close()
    return str(report_path)


def _load_and_validate(
    json_path: str,
    workbook_path: str,
    generated_data: Optional[Dict[str, List[Dict[str, Any]]]] = None,
) -> Dict[str, List[Dict[str, Any]]]:
    from scripts.entity_sanity_checks import (
        validate_entity_export,
        write_validation_workbook,
    )

    if generated_data is None:
        with Path(json_path).open(encoding="utf-8") as source:
            generated_data = json.load(source)
    if not isinstance(generated_data, dict):
        raise ValueError("Entity JSON must be an object keyed by entity type.")

    supported = {json_slug for _, json_slug, _, _, _ in HIERARCHY}
    unsupported = {
        slug: len(rows)
        for slug, rows in generated_data.items()
        if slug not in supported and rows
    }
    if unsupported:
        details = ", ".join(
            f"{slug}={count}" for slug, count in sorted(unsupported.items())
        )
        raise ValueError(f"Unsupported entity collections: {details}")

    report = validate_entity_export(workbook_path, generated_data)
    report_path = Path(json_path).with_name(
        f"{Path(json_path).name.removesuffix('_actual_data.json')}"
        "_entity_validation.xlsx"
    )
    write_validation_workbook(report, str(report_path))
    print(f"  Validation report: {report_path}")
    print(f"  Validation status: {report.status}")
    for check in report.checks:
        print(f"    [{check.status}] {check.number:02d}. {check.name}: {check.details}")
    if report.status == "FAIL":
        raise ValueError(
            f"Entity validation failed; API upload stopped. See {report_path}"
        )
    if report.status == "INCOMPLETE":
        print("  WARNING: Some validation checks were incomplete.")
    return generated_data


def _build_from_workbook(xlsx_path: str, project_type: str) -> Tuple[str, Dict[str, List[Dict[str, Any]]]]:
    from scripts.actual_to_entity import build_entities

    workbook_path = Path(xlsx_path).resolve()
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel workbook not found: {workbook_path}")
    generated = build_entities(str(workbook_path), project_type, 42)
    json_path = workbook_path.with_name(f"{workbook_path.stem}_actual_data.json")
    supported = {json_slug for _, json_slug, _, _, _ in HIERARCHY}
    data = {
        slug: rows
        for slug, rows in generated.items()
        if slug in supported or rows
    }
    with json_path.open("w", encoding="utf-8") as output:
        json.dump(data, output, indent=2, ensure_ascii=False, default=str)
    return str(json_path), data


def _source_workbook_for_json(json_path: Path, source_xlsx: str) -> Path:
    if source_xlsx:
        return Path(source_xlsx).resolve()
    name = json_path.name
    suffix = "_actual_data.json"
    if name.endswith(suffix):
        name = name[:-len(suffix)] + ".xlsx"
    else:
        name = f"{json_path.stem}.xlsx"
    return json_path.with_name(name)


def main(argv: Optional[Sequence[str]] = None) -> int:
    _load_api_environment()
    parser = argparse.ArgumentParser(
        description="Validate and upload construction entities through the BluConn API."
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--xlsx", help="Build and validate entities from this workbook.")
    source.add_argument("--json", help="Upload an existing *_actual_data.json file.")
    parser.add_argument("--project-type", default="road",
                        choices=["road", "pipeline", "building"])
    parser.add_argument(
        "--source-xlsx",
        default="",
        help="Workbook to validate when --json is used (defaults to the matching .xlsx).",
    )
    parser.add_argument("--batch-size", type=int, default=500)
    parser.add_argument("--timeout", type=int, default=60)
    parser.add_argument("--retries", type=int, default=3)
    args = parser.parse_args(argv)
    if args.batch_size < 1:
        parser.error("--batch-size must be at least 1")
    if args.timeout < 1:
        parser.error("--timeout must be at least 1 second")
    if args.retries < 0:
        parser.error("--retries cannot be negative")

    base_url = os.environ.get("BLUCONN_API_URL", "").strip()
    account_id = os.environ.get("BLUCONN_ACCOUNT_ID", "").strip()
    auth_token = os.environ.get("BLUCONN_AUTH_TOKEN", "").strip()
    missing = [
        key for key, value in (
            ("BLUCONN_API_URL", base_url),
            ("BLUCONN_ACCOUNT_ID", account_id),
            ("BLUCONN_AUTH_TOKEN", auth_token),
        ) if not value
    ]
    if missing:
        print(
            "Missing API configuration in the environment or RoadConstruction/.env: "
            + ", ".join(missing),
            file=sys.stderr,
        )
        return 2
    if not base_url.startswith(("https://", "http://")):
        print("BLUCONN_API_URL must be an http:// or https:// URL.", file=sys.stderr)
        return 2

    try:
        if args.xlsx:
            json_path, data = _build_from_workbook(args.xlsx, args.project_type)
            workbook_path = Path(args.xlsx).resolve()
        else:
            json_path = str(Path(args.json).resolve())
            if not Path(json_path).is_file():
                raise FileNotFoundError(f"Entity JSON not found: {json_path}")
            workbook_path = _source_workbook_for_json(
                Path(json_path), args.source_xlsx
            )
            if not workbook_path.is_file():
                raise FileNotFoundError(
                    f"Validation workbook not found: {workbook_path}. "
                    "Provide --source-xlsx."
                )
            with Path(json_path).open(encoding="utf-8") as source_file:
                data = json.load(source_file)
        data = _prepare_api_entities(
            _load_and_validate(json_path, str(workbook_path), data)
        )
        api = BluConnApi(
            base_url,
            account_id,
            auth_token,
            timeout=args.timeout,
            max_retries=args.retries,
            batch_size=args.batch_size,
        )
        print(f"  Entity JSON: {json_path}")
        print("  Destination: BluConn API")
        stats: Dict[str, Dict[str, int]] = {
            api_slug: _new_entity_stats(len(data.get(json_slug, [])))
            for api_slug, json_slug, _, _, _ in HIERARCHY
        }
        duplicates: List[Dict[str, str]] = []
        failures: List[Dict[str, str]] = []
        status = "API_FAILED"
        error_message = ""
        try:
            stats, duplicates, failures = upload_entities(api, data)
            status = "API_PARTIAL_FAILURE" if failures else "API_COMPLETED"
        except ApiUploadError as exc:
            error_message = str(exc)
            print(f"  API setup/upload failed: {error_message}", file=sys.stderr)
            failures.append({
                "entity_type": "API",
                "entity_id": "",
                "reason": error_message,
            })
            for api_slug, json_slug, _, _, _ in HIERARCHY:
                stats[api_slug]["failed"] += len(data.get(json_slug, []))

        report_path = write_api_report(
            json_path, stats, duplicates, failures, status, error_message
        )
        print(f"  API upload report: {report_path}")
        print(
            "  Summary: "
            f"uploaded={sum(item['inserted'] for item in stats.values())}, "
            f"skipped={sum(item['already_in_db'] + item['duplicate_in_json'] + item['skipped'] for item in stats.values())}, "
            f"failed={sum(item['failed'] for item in stats.values())}"
        )
        return 1 if failures or error_message else 0
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"API upload stopped: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
