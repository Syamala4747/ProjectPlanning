# Setup Guide — DB Push (for friend's machine)

This guide is for pushing the generated construction project entities
into the BluConn PostgreSQL database.

Your machine has the BluConn app running locally with Docker.
This script connects to that DB and inserts all 14 entity types.

---

## Prerequisites

- Python 3.10+
- The BluConn app running locally via Docker
- Git

---

## Step 1 — Clone the repo

```bash
git clone <repo-url>
cd RoadConstruction
```

---

## Step 2 — Install Python dependencies

```bash
pip install -r requirements.txt
pip install psycopg2-binary
```

---

## Step 3 — Find your local DB connection string

Run this in the BluConn project folder:

```bash
docker compose exec app env | grep DATABASE_URL
```

Or check the `.env` file in the BluConn project root.

It will look like:
```
DATABASE_URL=postgresql://user:password@postgres:5432/dbname
```

---

## Step 4 — Create your .env file

Create a file called `.env` in the RoadConstruction project root
(this file is gitignored — never commit it):

```
DATABASE_URL=postgresql://user:password@localhost:5432/dbname
```

> **Note:** Replace `postgres` (Docker hostname) with `localhost` if
> you are running the script from outside Docker.
> If running from inside Docker, use the original hostname.

---

## Step 5 — Generate the entity JSON

Run the planner for whichever project type you want:

```bash
# Road
python main.py road --title "10 KM 4-Lane Highway" --duration "365 days" --budget 300 --start-date 2026-01-01

# Pipeline
python main.py pipeline --title "Hyderabad Water Pipeline" --length 10 --diameter 600 --material DI --trench-depth 2 --fronts 2 --duration "18 months" --budget 15 --start-date 2026-01-01

# Building
python main.py building --title "ABC Residential Building" --area 10000 --floors 4 --quality standard --duration "24 months" --budget 22 --start-date 2026-01-01
```

This creates a file like:
```
output/10_km_4_lane_highway_actual_data.json
```

---

## Step 6 — Push to the database

```bash
# Auto-detects the most recently generated JSON in output/
python scripts/push_to_db.py

# Or specify a specific file
python scripts/push_to_db.py --json output/10_km_4_lane_highway_actual_data.json

# To clear existing data first (clean push)
python scripts/push_to_db.py --wipe
```

Expected output:
```
============================================================
  PUSH TO DATABASE
============================================================
  JSON   : output/10_km_4_lane_highway_actual_data.json
  Records: 62,654 across 14 entity types

  Connecting to database...
  Connected.

  Resolving account...
  Using existing account_id: abc-123...

  Upserting entity_types...
  14 entity types ready.

  Upserting entity_type_relationships...
  8 relationships ready.

  Inserting entities...
    project                                  inserted=      1
    assumption                               inserted=    176
    asset                                    inserted=     17
    material                                 inserted=     54
    subcontractor                            inserted=     24
    logistics                                inserted=   3893
    boq                                      inserted=     19
    activity                                 inserted=     19
    plan                                     inserted=     19
    supervisor_update                        inserted=    939
    engineering_daily_logs                   inserted=  37494
    material_procurement_request             inserted=   3843
    material_issue_request                   inserted=   5648
    material_transaction                     inserted=   9491

  ✔ All committed.

============================================================
  DONE — open the BluConn UI to see your data
============================================================
```

---

## Troubleshooting

**"could not translate host name ... to address"**
→ The DB hostname is a Docker internal name. Either:
- Run the script from inside Docker: `docker compose exec app python scripts/push_to_db.py`
- Or expose the DB port in docker-compose.yml and use `localhost` in DATABASE_URL

**"psycopg2 not installed"**
→ Run: `pip install psycopg2-binary`

**"DATABASE_URL not set"**
→ Make sure `.env` exists in the project root with the DATABASE_URL line

**"permission denied for table entity_types"**
→ The DB user needs INSERT rights. Ask the BluConn team or use the admin credentials.

---

## What gets pushed

| Entity Type | Records (road example) |
|---|---|
| project | 1 |
| assumption | 176 |
| asset | 17 |
| material | 54 |
| subcontractor | 24 |
| boq | 19 |
| activity | 19 |
| plan | 19 |
| supervisor_update | 939 |
| engineering_daily_logs | 37,494 |
| material_procurement_request | 3,843 |
| material_issue_request | 5,648 |
| material_transaction | 9,491 |
| logistics | 3,893 |
| **TOTAL** | **~62,600** |

After the push, open the BluConn UI dashboard and you will see
all the project data populated.
