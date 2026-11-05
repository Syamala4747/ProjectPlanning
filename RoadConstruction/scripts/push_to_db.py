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
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

# ── load .env ────────────────────────────────────────────────────────────────
_root = Path(__file__).resolve().parent.parent
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

def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(text).lower()).strip("_")


# ── hierarchy definition ─────────────────────────────────────────────────────
# Matches push-entities.mjs HIERARCHY exactly.
# (slug, parent_slug, id_field, parent_field_in_data)
HIERARCHY = [
    ("project",                       None,         "projectId",                    None),
    ("assumption",                    None,         "assumptionId",                 None),
    ("asset",                         None,         "assetId",                      None),
    ("material",                      None,         "materialId",                   None),
    ("subcontractor",                 None,         "subContractorId",              None),
    ("logistics",                     None,         "logisticsId",                  None),
    ("boq",                           "project",    "boqId",                        "projectId"),
    ("activity",                      "boq",        "activityId",                   "boqId"),
    ("plan",                          "activity",   "planId",                       "activityId"),
    ("supervisor_update",             "activity",   "supervisorUpdateId",           "activityId"),
    ("engineering_daily_logs",        "asset",      "engineeringDailyLogId",        "assetId"),
    ("material_procurement_request",  "material",   "materialProcurementRequestId", "materialId"),
    ("material_issue_request",        "material",   "materialIssueRequestId",       "materialId"),
    ("material_transaction",          "material",   "materialTransactionId",        "materialId"),
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
    "supervisor_update":             "Supervisor Update",
    "engineering_daily_logs":        "Engineering Daily Logs",
    "material_procurement_request":  "Material Procurement Request",
    "material_issue_request":        "Material Issue Request",
    "material_transaction":          "Material Transaction",
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

    id_col = next((h[2] for h in HIERARCHY if h[0] == slug), None)
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

    for slug, parent_slug, id_col, _ in HIERARCHY:
        rows = data.get(slug, [])
        label = SLUG_LABELS.get(slug, slug)
        schema = build_json_schema(slug, rows)
        parent_mode = "RESTRICTED" if parent_slug else "NONE"
        is_parent_required = bool(parent_slug)
        et_id = _uid()

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
            et_id, account_id, label, slug,
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
    for slug, parent_slug, _, _ in HIERARCHY:
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
                    account_id: str) -> None:
    """
    Insert entity rows in hierarchy order so parents exist before children.

    local_id_map:  local data id (e.g. "ACT-01")  →  DB entity row id (UUID)
    ancestry_path: slash-joined ancestor DB ids, e.g. "proj-uuid/boq-uuid"
    """
    local_to_db: Dict[str, str] = {}   # local_id → DB uuid
    local_to_path: Dict[str, str] = {} # local_id → ancestry_path
    now = _now()

    for slug, parent_slug, id_col, parent_field in HIERARCHY:
        rows = data.get(slug, [])
        et_id = type_id_map.get(slug)
        if not et_id:
            print(f"    [{slug}] no entity_type id — skipping")
            continue

        inserted = 0
        skipped  = 0

        for row in rows:
            local_id = str(row.get(id_col, "")) if id_col else _uid()
            db_id    = _uid()

            # Resolve parent
            db_parent_id:   Optional[str] = None
            db_ancestry:    str           = ""

            if parent_field:
                local_parent_id = str(row.get(parent_field, ""))
                db_parent_id    = local_to_db.get(local_parent_id)
                if db_parent_id:
                    parent_path  = local_to_path.get(local_parent_id, "")
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

            try:
                cur.execute("""
                    INSERT INTO entities
                        (id, account_id, entity_type_id,
                         parent_id, ancestry_path,
                         data, version,
                         created_at, updated_at)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (id) DO NOTHING
                """, (
                    db_id, account_id, et_id,
                    db_parent_id, db_ancestry,
                    json.dumps(clean_data, ensure_ascii=False, default=str),
                    1, now, now,
                ))
                local_to_db[local_id]   = db_id
                local_to_path[local_id] = db_ancestry
                inserted += 1
            except Exception as e:
                skipped += 1
                if skipped == 1:
                    print(f"    [{slug}] first insert error: {e}")

        print(f"    {slug:<40} inserted={inserted:>6}  skipped={skipped}")


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


def main() -> None:
    p = argparse.ArgumentParser(description="Push entity JSON to PostgreSQL")
    p.add_argument("--json",  default="",    help="Path to *_actual_data.json (auto-detected if omitted)")
    p.add_argument("--wipe",  action="store_true", help="Delete existing entities for this account first")
    args = p.parse_args()

    json_path = args.json or find_json()
    if not json_path or not os.path.exists(json_path):
        print("No *_actual_data.json found. Run the planner first.")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  PUSH TO DATABASE")
    print("=" * 60)
    print(f"  JSON   : {json_path}")

    with open(json_path, encoding="utf-8") as f:
        data: Dict[str, List[Dict]] = json.load(f)

    total = sum(len(v) for v in data.values())
    print(f"  Records: {total:,} across {len(data)} entity types")

    print("\n  Connecting to database...")
    conn = connect()
    cur  = conn.cursor()
    print("  Connected.")

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
        child_count = sum(1 for _, ps, _, _ in HIERARCHY if ps)
        print(f"  {child_count} relationships ready.")

        # 5. Insert entities
        print("\n  Inserting entities...")
        insert_entities(cur, data, type_id_map, account_id)

        conn.commit()
        print("\n  ✔ All committed.")

    except Exception as e:
        conn.rollback()
        print(f"\n  ✗ Error — rolled back: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)
    finally:
        cur.close()
        conn.close()

    print("\n" + "=" * 60)
    print("  DONE — open the BluConn UI to see your data")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
