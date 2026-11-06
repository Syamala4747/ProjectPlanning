"""
push_to_db.py
=============
Reads *_actual_data.json and inserts all 14 entity types directly into
the BluConn PostgreSQL database.

Tables written:
    entity_types              — one row per entity type (14 total)
    entity_type_relationships — parent→child type links
    entities                  — one row per instance (data = jsonb)

account_id is read from the accounts table automatically.
If accounts is empty, a new account row is created.

Usage
-----
    python scripts/push_to_db.py
    python scripts/push_to_db.py --json output/road_actual_data.json
    python scripts/push_to_db.py --wipe   # clears existing entities first

DATABASE_URL must be in .env or set as an environment variable.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import uuid
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ── load .env ────────────────────────────────────────────────────────────────
_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))
_env  = _root / ".env"
if _env.exists():
    for _line in _env.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if not _line or _line.startswith("#"):
            continue
        if "=" not in _line:
            continue
        k, _, v = _line.partition("=")
        k = k.strip(); v = v.strip().strip('"').strip("'")
        if k not in os.environ:
            os.environ[k] = v

try:
    import psycopg2
    import psycopg2.extras
except ImportError:
    print("psycopg2 not installed. Run:  pip install psycopg2-binary")
    sys.exit(1)


# ── helpers ───────────────────────────────────────────────────────────────────

def _uid() -> str:
    return str(uuid.uuid4())

def _now() -> datetime:
    return datetime.now(timezone.utc)

def _next_entity_timestamp(previous: Optional[datetime]) -> datetime:
    """Return a UTC timestamp at least one millisecond after the prior row."""
    now = _now()
    if previous is None:
        return now
    return max(now, previous + timedelta(milliseconds=1))

def _entity_event_date(row: Dict[str, Any]) -> Optional[date]:
    """Find the record's business date, preferring its actual event/request date."""
    for field in (
        "date", "requestDate", "plannedStart", "startDate", "dispatchDate",
        "deliveryDate", "awardDate",
    ):
        value = row.get(field)
        if value in (None, ""):
            continue
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        try:
            return date.fromisoformat(str(value).strip()[:10])
        except ValueError as exc:
            raise ValueError(
                f"Invalid {field} for entity timestamp: {value!r}"
            ) from exc
    return None

def _entity_timestamp(
    row: Dict[str, Any],
    date_offsets: Dict[date, int],
    previous_undated: Optional[datetime],
) -> tuple[datetime, Optional[datetime]]:
    """Use the entity's business date at UTC midnight, unique per day to 1 ms."""
    event_date = _entity_event_date(row)
    if event_date is None:
        timestamp = _next_entity_timestamp(previous_undated)
        return timestamp, timestamp

    offset = date_offsets.get(event_date, 0)
    date_offsets[event_date] = offset + 1
    timestamp = datetime.combine(
        event_date, time.min, tzinfo=timezone.utc
    ) + timedelta(milliseconds=offset)
    return timestamp, previous_undated

def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(text).lower()).strip("_")


# ── hierarchy definition ─────────────────────────────────────────────────────
# (database_slug, json_slug, parent_slug, id_field, parent_field_in_data)
HIERARCHY = [
    ("project",                       "project",                       None,                        "projectId",                    None),
    ("assumption",                    "assumption",                    None,                        "assumptionId",                 None),
    ("asset",                         "asset",                         None,                        "assetId",                      None),
    ("material",                      "material",                      None,                        "materialId",                   None),
    ("subcontractor",                 "subcontractor",                 None,                        "subContractorId",              None),
    ("logistics",                     "logistics",                     None,                        "logisticsId",                  None),
    ("boq",                           "boq",                           "project",                   "boqId",                        "projectId"),
    ("activity",                      "activity",                      "boq",                       "activityId",                   "boqId"),
    ("plan",                          "plan",                          "activity",                  "planId",                       "activityId"),
    ("supervisor-update",             "supervisor_update",             "activity",                  "supervisorUpdateId",           "activityId"),
    ("engineering-daily-logs",        "engineering_daily_logs",        "asset",                     "engineeringDailyLogId",        "assetId"),
    ("material-procurement-request",  "material_procurement_request",  "material",                  "materialProcurementRequestId", "materialId"),
    ("material-issue-request",        "material_issue_request",        "material",                  "materialIssueRequestId",       "materialId"),
    ("material-transaction",          "material_transaction",          "material",                  "materialTransactionId",        "materialId"),
]

SLUG_LABELS = {
    "project":                       "Project",
    "assumption":                    "Assumption",
    "asset":                         "Asset",
    "material":                      "Material",
    "subcontractor":                 "SubContractor",
    "logistics":                     "Logistics",
    "boq":                           "BOQ",
    "activity":                      "Activity",
    "plan":                          "Plan",
    "supervisor-update":             "Supervisor Update",
    "engineering-daily-logs":        "Engineering Daily Logs",
    "material-procurement-request":  "Material Procurement Request",
    "material-issue-request":        "Material Issue Request",
    "material-transaction":          "Material Transaction",
}


# ── DB connection ─────────────────────────────────────────────────────────────

def connect() -> psycopg2.extensions.connection:
    url = os.environ.get("DATABASE_URL", "")
    if not url:
        print("DATABASE_URL not set. Add it to .env")
        sys.exit(1)
    try:
        conn = psycopg2.connect(url)
        conn.autocommit = False
        return conn
    except Exception as e:
        print(f"Could not connect to database: {e}")
        sys.exit(1)


# ── account_id ────────────────────────────────────────────────────────────────

def get_or_create_account(cur) -> str:
    """Read the first account_id from the accounts table, or create one."""
    # Try reading existing
    try:
        cur.execute("SELECT id FROM accounts ORDER BY created_at LIMIT 1")
        row = cur.fetchone()
        if row:
            aid = str(row[0])
            print(f"  Using existing account_id: {aid}")
            return aid
    except Exception:
        pass  # accounts table may have different structure

    # Try with name column
    try:
        cur.execute("SELECT id FROM accounts LIMIT 1")
        row = cur.fetchone()
        if row:
            aid = str(row[0])
            print(f"  Using existing account_id: {aid}")
            return aid
    except Exception:
        pass

    # Create a new account row
    aid = _uid()
    now = _now()
    try:
        cur.execute("""
            INSERT INTO accounts (id, name, created_at, updated_at)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (id) DO NOTHING
        """, (aid, "BluConn Construction", now, now))
        print(f"  Created new account_id: {aid}")
    except Exception as e:
        # accounts table might have required columns we don't know about
        # fall back to a deterministic ID based on the DB URL
        db_url = os.environ.get("DATABASE_URL", "fallback")
        import hashlib
        aid = str(uuid.UUID(hashlib.md5(db_url.encode()).hexdigest()))
        print(f"  Using derived account_id: {aid} (could not insert into accounts: {e})")
    return aid


# ── entity_types ──────────────────────────────────────────────────────────────

def build_json_schema(slug: str, rows: List[Dict]) -> Dict:
    """Derive a minimal JSON schema from the data rows."""
    sample = rows[0] if rows else {}
    props: Dict[str, Any] = {}
    for k, v in sample.items():
        if isinstance(v, bool):
            t = "boolean"
        elif isinstance(v, (int, float)):
            t = "number"
        elif isinstance(v, (dict, list)):
            t = "string"  # stored as JSON string in the schema
        else:
            sv = str(v) if v is not None else ""
            t = "string"
            if re.match(r"^\d{4}-\d{2}-\d{2}", sv):
                props[k] = {"type": "string", "format": "date-time", "title": k}
                continue
        props[k] = {"type": t, "title": k}

    id_col = next((h[3] for h in HIERARCHY if h[0] == slug), None)
    return {
        "$schema": "http://json-schema.org/draft-07/schema#",
        "type": "object",
        "properties": props,
        "required": [id_col] if id_col and id_col in props else [],
    }


def upsert_entity_types(cur, data: Dict, account_id: str) -> Dict[str, str]:
    """
    Insert entity_types rows. Returns {slug: entity_type_db_id}.
    Uses ON CONFLICT (account_id, slug) DO UPDATE so re-runs are safe.
    """
    type_id_map: Dict[str, str] = {}
    now = _now()

    for db_slug, json_slug, parent_slug, id_col, _ in HIERARCHY:
        rows = data.get(json_slug, [])
        slug = db_slug
        label = SLUG_LABELS.get(slug, slug)
        schema = build_json_schema(slug, rows)
        parent_mode = "RESTRICTED" if parent_slug else "NONE"
        is_parent_required = bool(parent_slug)
        et_id = _uid()

        cur.execute("""
            SELECT id, slug
            FROM entity_types
            WHERE account_id = %s AND slug IN (%s, %s)
        """, (account_id, db_slug, json_slug))
        existing_types = cur.fetchall()
        if len(existing_types) > 1:
            raise RuntimeError(
                f"Both legacy and UI-compatible entity types exist for {slug!r}; "
                "resolve the duplicate types before pushing."
            )
        if existing_types and existing_types[0][1] != db_slug:
            cur.execute(
                "UPDATE entity_types SET slug = %s WHERE account_id = %s AND id = %s",
                (db_slug, account_id, existing_types[0][0]),
            )

        cur.execute("""
            INSERT INTO entity_types
                (id, account_id, name, slug, description,
                 json_schema, searchable_fields,
                 is_parent_required, parent_mode,
                 created_at, updated_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (account_id, slug)
                DO UPDATE SET
                    name              = EXCLUDED.name,
                    json_schema       = EXCLUDED.json_schema,
                    updated_at        = EXCLUDED.updated_at
            RETURNING id
        """, (
            et_id, account_id, label, db_slug,
            f"Construction planning entity: {label}",
            json.dumps(schema), json.dumps([id_col] if id_col else []),
            is_parent_required, parent_mode,
            now, now,
        ))
        row = cur.fetchone()
        actual_id = str(row[0]) if row else et_id
        type_id_map[slug] = actual_id

    return type_id_map


def upsert_relationships(cur, type_id_map: Dict[str, str],
                          account_id: str) -> None:
    """Insert entity_type_relationships rows."""
    now = _now()
    for slug, _, parent_slug, _, _ in HIERARCHY:
        if not parent_slug:
            continue
        child_id  = type_id_map.get(slug)
        parent_id = type_id_map.get(parent_slug)
        if not child_id or not parent_id:
            continue
        cur.execute("""
            INSERT INTO entity_type_relationships
                (id, account_id, parent_type_id, child_type_id,
                 created_at, updated_at)
            VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (account_id, parent_type_id, child_type_id)
                DO NOTHING
        """, (_uid(), account_id, parent_id, child_id, now, now))


# ── entities ──────────────────────────────────────────────────────────────────

def insert_entities(cur, data: Dict, type_id_map: Dict[str, str],
                    account_id: str,
                    stats: Optional[Dict[str, Dict[str, int]]] = None,
                    duplicates: Optional[List[Dict[str, str]]] = None
                    ) -> tuple[Dict[str, Dict[str, int]], List[Dict[str, str]]]:
    """
    Insert entity rows in hierarchy order so parents exist before children.

    local_id_map:  local data id (e.g. "ACT-01")  →  DB entity row id (UUID)
    ancestry_path: slash-joined ancestor DB ids, e.g. "proj-uuid/boq-uuid"
    """
    local_to_db: Dict[tuple[str, str], str] = {}
    local_to_path: Dict[tuple[str, str], str] = {}
    date_timestamp_offsets: Dict[date, int] = {}
    last_undated_timestamp: Optional[datetime] = None
    stats = {} if stats is None else stats
    duplicates = [] if duplicates is None else duplicates

    for slug, json_slug, parent_slug, id_col, parent_field in HIERARCHY:
        rows = data.get(json_slug, [])
        entity_stats = {
            "sent": len(rows),
            "insert_attempted": 0,
            "inserted": 0,
            "already_in_db": 0,
            "duplicate_in_json": 0,
            "db_conflicts": 0,
            "failed": 0,
        }
        stats[slug] = entity_stats
        et_id = type_id_map.get(slug)
        if not et_id:
            print(f"    [{slug}] no entity_type id — skipping")
            continue

        rows_by_project: Dict[str, List[Dict[str, Any]]] = {}
        for row in rows:
            project_id = str(row.get("projectId") or "")
            rows_by_project.setdefault(project_id, []).append(row)

        existing_by_key: Dict[tuple[str, str], tuple[str, str]] = {}
        for project_id, project_rows in rows_by_project.items():
            if not project_id or not id_col:
                continue
            incoming_ids = list(dict.fromkeys(
                str(row.get(id_col)) for row in project_rows
                if row.get(id_col) not in (None, "")
            ))
            if not incoming_ids:
                continue
            cur.execute("""
                SELECT id, ancestry_path, data ->> %s AS local_id
                FROM entities
                WHERE account_id = %s
                  AND entity_type_id = %s
                  AND data ->> 'projectId' = %s
                  AND data ->> %s = ANY(%s)
                ORDER BY created_at, id
            """, (id_col, account_id, et_id, project_id, id_col, incoming_ids))
            for existing_id, ancestry_path, existing_local_id in cur.fetchall():
                key = (project_id, str(existing_local_id))
                if key in existing_by_key:
                    duplicates.append({
                        "entity_type": slug,
                        "project_id": project_id,
                        "entity_id": key[1],
                        "duplicate_kind": "multiple matching records already in database",
                        "database_entity_ids": f"{existing_by_key[key][0]}, {existing_id}",
                    })
                    continue
                existing_by_key[key] = (
                    str(existing_id),
                    str(ancestry_path or ""),
                )

        seen_input: set[tuple[str, str]] = set()
        for row in rows:
            row_timestamp, last_undated_timestamp = _entity_timestamp(
                row, date_timestamp_offsets, last_undated_timestamp
            )
            local_id = str(row.get(id_col) or "") if id_col else _uid()
            if not local_id:
                local_id = _uid()
            project_id = str(row.get("projectId") or "")
            local_key = (project_id, local_id)

            if id_col and local_key in seen_input:
                entity_stats["duplicate_in_json"] += 1
                mapped_id = local_to_db.get(local_key)
                duplicates.append({
                    "entity_type": slug,
                    "project_id": project_id,
                    "entity_id": local_id,
                    "duplicate_kind": "repeated business ID in JSON",
                    "database_entity_ids": mapped_id or "",
                })
                continue
            if id_col:
                seen_input.add(local_key)

            existing = existing_by_key.get(local_key)
            if existing:
                existing_id, existing_path = existing
                local_to_db[local_key] = existing_id
                local_to_path[local_key] = existing_path
                entity_stats["already_in_db"] += 1
                duplicates.append({
                    "entity_type": slug,
                    "project_id": project_id,
                    "entity_id": local_id,
                    "duplicate_kind": "already exists in database; skipped",
                    "database_entity_ids": existing_id,
                })
                continue

            db_id    = _uid()

            # Resolve parent
            db_parent_id:   Optional[str] = None
            db_ancestry:    str           = ""

            if parent_field:
                local_parent_id = str(row.get(parent_field, ""))
                parent_key = (project_id, local_parent_id)
                db_parent_id = local_to_db.get(parent_key)
                if db_parent_id:
                    parent_path = local_to_path.get(parent_key, "")
                    db_ancestry  = f"{parent_path}/{db_parent_id}".lstrip("/")

            # Serialise nested structures inside data as JSON strings
            # so the jsonb column stays flat enough for the UI to render
            clean_data: Dict[str, Any] = {}
            for k, v in row.items():
                if isinstance(v, (dict, list)):
                    clean_data[k] = json.dumps(v, ensure_ascii=False, default=str)
                elif v is None:
                    clean_data[k] = None
                else:
                    clean_data[k] = v

            entity_stats["insert_attempted"] += 1
            try:
                cur.execute("""
                    INSERT INTO entities
                        (id, account_id, entity_type_id,
                         parent_id, ancestry_path,
                         data, version,
                         created_at, updated_at)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (id) DO NOTHING
                    RETURNING id
                """, (
                    db_id, account_id, et_id,
                    db_parent_id, db_ancestry,
                    json.dumps(clean_data, ensure_ascii=False, default=str),
                    1, row_timestamp, row_timestamp,
                ))
            except psycopg2.Error as exc:
                entity_stats["failed"] += 1
                raise RuntimeError(
                    f"Database insert failed for {slug} record {local_id!r} "
                    f"in project {project_id!r}: {exc}"
                ) from exc
            inserted_row = cur.fetchone()
            if not inserted_row:
                entity_stats["db_conflicts"] += 1
                continue
            local_to_db[local_key] = str(inserted_row[0])
            local_to_path[local_key] = db_ancestry
            entity_stats["inserted"] += 1

        print(
            f"    {slug:<40} sent={entity_stats['sent']:>6}  "
            f"inserted={entity_stats['inserted']:>6}  "
            f"already_in_db={entity_stats['already_in_db']:>6}  "
            f"duplicate_in_json={entity_stats['duplicate_in_json']:>4}"
        )
    return stats, duplicates


def write_push_report(
    json_path: str,
    stats: Dict[str, Dict[str, int]],
    duplicates: List[Dict[str, str]],
    commit_status: str,
    error_message: str = "",
) -> str:
    """Write a per-entity insert and duplicate reconciliation workbook."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    json_file = Path(json_path).resolve()
    report_path = json_file.with_name(
        f"{json_file.name.removesuffix('_actual_data.json')}_db_push_report.xlsx"
    )
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    total_sent = sum(item["sent"] for item in stats.values())
    total_attempted = sum(item["insert_attempted"] for item in stats.values())
    total_inserted = sum(item["inserted"] for item in stats.values())
    total_existing = sum(item["already_in_db"] for item in stats.values())
    total_json_duplicates = sum(item["duplicate_in_json"] for item in stats.values())
    total_conflicts = sum(item["db_conflicts"] for item in stats.values())
    total_failures = sum(item.get("failed", 0) for item in stats.values())
    committed = total_inserted if commit_status == "COMMITTED" else 0
    summary.append(["Database Push Report", ""])
    summary.append(["JSON file", json_file.name])
    summary.append(["Commit status", commit_status])
    summary.append(["Records in JSON (sent for processing)", total_sent])
    summary.append(["New records attempted", total_attempted])
    summary.append(["New records committed", committed])
    summary.append(["Already in database (skipped)", total_existing])
    summary.append(["Duplicate rows in JSON (skipped)", total_json_duplicates])
    summary.append(["Database ID conflicts", total_conflicts])
    summary.append(["Failed inserts", total_failures])
    summary.append(["Duplicate findings", len(duplicates)])
    if error_message:
        summary.append(["Error", error_message])

    entity_sheet = workbook.create_sheet("Entity Counts")
    entity_sheet.append([
        "Entity type", "JSON records", "Insert attempts", "Inserted in transaction",
        "Committed", "Already in database", "Duplicate rows in JSON", "ID conflicts",
        "Failed inserts",
    ])
    for slug, values in stats.items():
        entity_sheet.append([
            slug,
            values["sent"],
            values["insert_attempted"],
            values["inserted"],
            values["inserted"] if commit_status == "COMMITTED" else 0,
            values["already_in_db"],
            values["duplicate_in_json"],
            values["db_conflicts"],
            values.get("failed", 0),
        ])

    duplicate_sheet = workbook.create_sheet("Duplicates")
    duplicate_sheet.append([
        "Entity type", "Project ID", "Entity ID", "Finding", "Database entity IDs",
    ])
    for duplicate in duplicates:
        duplicate_sheet.append([
            duplicate["entity_type"],
            duplicate["project_id"],
            duplicate["entity_id"],
            duplicate["duplicate_kind"],
            duplicate["database_entity_ids"],
        ])
    if not duplicates:
        duplicate_sheet.append(["", "", "", "No duplicate records detected", ""])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2" if sheet.title != "Summary" else "A2"
        sheet.auto_filter.ref = sheet.dimensions if sheet.title != "Summary" else "A1:B1"
        for cell in sheet[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill
        for column_cells in sheet.columns:
            width = max(
                (len(str(cell.value or "")) for cell in column_cells),
                default=10,
            )
            sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = min(
                max(width + 2, 12), 60
            )
    workbook.save(report_path)
    return str(report_path)


# ── wipe ──────────────────────────────────────────────────────────────────────

def wipe(cur, account_id: str) -> None:
    print("  Wiping existing data for this account...")
    cur.execute("DELETE FROM entities                WHERE account_id = %s", (account_id,))
    cur.execute("DELETE FROM entity_type_relationships WHERE account_id = %s", (account_id,))
    cur.execute("DELETE FROM entity_types            WHERE account_id = %s", (account_id,))
    print("  Wiped.")


# ── main ──────────────────────────────────────────────────────────────────────

def find_json() -> str:
    """Auto-find the most recently written *_actual_data.json in output/."""
    out = _root / "output"
    if not out.exists():
        return ""
    files = sorted(
        out.glob("*_actual_data.json"),
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )
    return str(files[0]) if files else ""


def unsupported_entity_records(data: Dict[str, List[Dict]]) -> Dict[str, int]:
    """Return non-empty JSON entity collections not handled by the DB pusher."""
    supported = {json_slug for _, json_slug, _, _, _ in HIERARCHY}
    return {
        slug: len(rows)
        for slug, rows in data.items()
        if slug not in supported and rows
    }


def export_excel_to_json(
    xlsx_path: str,
    project_type: str = "road",
    seed: int = 42,
) -> str:
    """Build supported entity records from a workbook and save adjacent JSON."""
    from scripts.actual_to_entity import build_entities

    workbook_path = Path(xlsx_path).resolve()
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel workbook not found: {workbook_path}")

    generated = build_entities(str(workbook_path), project_type, seed)
    supported = {json_slug for _, json_slug, _, _, _ in HIERARCHY}
    data = {
        slug: rows
        for slug, rows in generated.items()
        if slug in supported or rows
    }
    unsupported = unsupported_entity_records(data)
    if unsupported:
        details = ", ".join(
            f"{slug}={count}" for slug, count in sorted(unsupported.items())
        )
        raise ValueError(
            "The workbook generated entity types this database pusher does not "
            f"support: {details}"
        )

    json_path = workbook_path.with_name(f"{workbook_path.stem}_actual_data.json")
    with json_path.open("w", encoding="utf-8") as output_file:
        json.dump(data, output_file, indent=2, ensure_ascii=False, default=str)
    return str(json_path)


def validate_json_for_push(json_path: str, data: Dict[str, List[Dict]]):
    """Write the source-backed Excel sanity report beside the JSON being pushed."""
    from scripts.entity_sanity_checks import (
        validate_entity_export,
        write_validation_workbook,
    )

    json_file = Path(json_path).resolve()
    suffix = "_actual_data.json"
    stem = (
        json_file.name[:-len(suffix)]
        if json_file.name.endswith(suffix)
        else json_file.stem
    )
    workbook_path = json_file.with_name(f"{stem}.xlsx")
    report_path = json_file.with_name(f"{stem}_entity_validation.xlsx")

    report = validate_entity_export(str(workbook_path), data)
    write_validation_workbook(report, str(report_path))

    print(f"\n  Validation report: {report_path}")
    for check in report.checks:
        print(f"    [{check.status}] {check.number:02d}. {check.name}: {check.details}")
    print(
        f"  Validation status: {report.status} "
        f"(PASS={sum(check.status == 'PASS' for check in report.checks)}, "
        f"FAIL={sum(check.status == 'FAIL' for check in report.checks)}, "
        f"INCOMPLETE={sum(check.status == 'INCOMPLETE' for check in report.checks)})"
    )
    if report.status == "FAIL":
        raise ValueError(f"Sanity checks failed; database push stopped. See {report_path}")
    if report.status == "INCOMPLETE":
        print("  WARNING: Checks could not all be verified; see the report for details.")
    return report


def main() -> None:
    p = argparse.ArgumentParser(description="Push entity JSON to PostgreSQL")
    source_group = p.add_mutually_exclusive_group()
    source_group.add_argument(
        "--json",
        default="",
        help="Path to *_actual_data.json (auto-detected if omitted)",
    )
    source_group.add_argument(
        "--xlsx",
        default="",
        help="Generate entities from this workbook, then validate and push them",
    )
    p.add_argument(
        "--project-type",
        default="road",
        choices=["road", "pipeline", "building"],
        help="Project type when using --xlsx (default: road)",
    )
    p.add_argument("--wipe",  action="store_true", help="Delete existing entities for this account first")
    p.add_argument(
        "--schema-only",
        action="store_true",
        help="Update entity types and relationships without inserting entity records",
    )
    args = p.parse_args()
    if args.wipe and args.schema_only:
        p.error("--wipe cannot be used with --schema-only")

    if args.xlsx:
        try:
            json_path = export_excel_to_json(args.xlsx, args.project_type)
        except (FileNotFoundError, ValueError, OSError) as exc:
            print(f"Could not export entities from Excel: {exc}")
            sys.exit(1)
        print(f"  Excel : {args.xlsx}")
        print(f"  JSON  : {json_path}")
    else:
        json_path = args.json or find_json()
    if not json_path or not os.path.exists(json_path):
        print("No input file found. Generate an Excel plan or provide --json/--xlsx.")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  PUSH TO DATABASE")
    print("=" * 60)
    print(f"  JSON   : {json_path}")

    with open(json_path, encoding="utf-8") as f:
        data: Dict[str, List[Dict]] = json.load(f)

    unsupported = unsupported_entity_records(data)
    if unsupported:
        details = ", ".join(
            f"{slug}={count}" for slug, count in sorted(unsupported.items())
        )
        print(
            "Unsupported entity collections contain records and cannot be pushed "
            f"by this script: {details}. No database changes were made."
        )
        sys.exit(1)

    total = sum(len(v) for v in data.values())
    print(f"  Records: {total:,} across {len(data)} entity types")
    try:
        validate_json_for_push(json_path, data)
    except ValueError as exc:
        print(f"\n  ✗ {exc}")
        sys.exit(1)

    print("\n  Connecting to database...")
    conn = connect()
    cur  = conn.cursor()
    print("  Connected.")
    push_stats: Dict[str, Dict[str, int]] = {
        slug: {
            "sent": len(data.get(json_slug, [])),
            "insert_attempted": 0,
            "inserted": 0,
            "already_in_db": 0,
            "duplicate_in_json": 0,
            "db_conflicts": 0,
            "failed": 0,
        }
        for slug, json_slug, _, _, _ in HIERARCHY
    }
    duplicates: List[Dict[str, str]] = []
    commit_status = "ROLLED_BACK"

    try:
        # 1. Get or create account
        print("\n  Resolving account...")
        account_id = get_or_create_account(cur)

        # 2. Wipe if requested
        if args.wipe:
            wipe(cur, account_id)

        # 3. Upsert entity_types (14 rows)
        print("\n  Upserting entity_types...")
        type_id_map = upsert_entity_types(cur, data, account_id)
        print(f"  {len(type_id_map)} entity types ready.")

        # 4. Upsert relationships
        print("\n  Upserting entity_type_relationships...")
        upsert_relationships(cur, type_id_map, account_id)
        child_count = sum(1 for _, _, ps, _, _ in HIERARCHY if ps)
        print(f"  {child_count} relationships ready.")

        # 5. Insert entities unless only synchronizing schema/UI-compatible slugs.
        if args.schema_only:
            print("\n  Schema-only mode: entity records were not modified.")
        else:
            print("\n  Inserting entities...")
            push_stats, duplicates = insert_entities(
                cur, data, type_id_map, account_id, push_stats, duplicates
            )

        conn.commit()
        commit_status = "SCHEMA_ONLY" if args.schema_only else "COMMITTED"
        print("\n  ✔ All committed.")

    except Exception as e:
        conn.rollback()
        print(f"\n  ✗ Error — rolled back: {e}")
        try:
            report_path = write_push_report(
                json_path, push_stats, duplicates, commit_status, str(e)
            )
            print(f"  Push report (transaction rolled back): {report_path}")
        except Exception as report_error:
            print(f"  Could not write push report: {report_error}")
        import traceback; traceback.print_exc()
        sys.exit(1)
    finally:
        cur.close()
        conn.close()

    report_path = write_push_report(json_path, push_stats, duplicates, commit_status)
    total_inserted = sum(values["inserted"] for values in push_stats.values())
    total_existing = sum(values["already_in_db"] for values in push_stats.values())
    total_json_duplicates = sum(
        values["duplicate_in_json"] for values in push_stats.values()
    )
    print("\n  PUSH RECONCILIATION")
    print(f"    Records in JSON:             {total:,}")
    print(f"    Newly inserted and committed: {total_inserted:,}")
    print(f"    Existing records skipped:     {total_existing:,}")
    print(f"    Duplicate JSON rows skipped:  {total_json_duplicates:,}")
    print(f"    Duplicate findings:           {len(duplicates):,}")
    print(f"    Excel report: {report_path}")

    print("\n" + "=" * 60)
    print("  DONE — open the BluConn UI to see your data")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
