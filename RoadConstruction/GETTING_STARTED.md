# Create, Update, and Push a Project

Run these commands in PowerShell from the `RoadConstruction` folder.

## One-time setup

Use Python 3.12 and install the required packages:

```powershell
py -3.12 -m venv .venv312
.\.venv312\Scripts\Activate.ps1
python -m pip install openpyxl pyyaml psycopg2-binary
```

For direct PostgreSQL pushes, configure `DATABASE_URL` in `.env`. For BluConn
API pushes, configure `BLUCONN_API_URL`, `BLUCONN_ACCOUNT_ID`, and
`BLUCONN_AUTH_TOKEN` there instead. Keep all credentials private; each uploader
uses only its own route's configuration.

## Road project

### Create a new road Excel

```powershell
python main.py road `
  --title "South Road" `
  --length 12 `
  --lanes 4 `
  --terrain plain `
  --duration "480 days" `
  --budget 450 `
  --start-date 2026-11-10 `
  --output "output\south_road.xlsx"
```

This command creates only `output\south_road.xlsx`.

### Update the road plan

Use the baseline Excel as `--input`:

```powershell
python main.py update-plan `
  --input "output\south_road.xlsx" `
  --status-date 2027-01-10 `
  --actual-progress 8.5 `
  --target-date 2028-04-15 `
  --target-progress 100 `
  --output "output\south_road_updated.xlsx"
```

### Push a road Excel to the database

Push either the new or updated workbook by setting its path:

```powershell
python scripts/push_to_db.py --xlsx "output\south_road.xlsx" --project-type road
```

For the updated workbook, use:

```powershell
python scripts/push_to_db.py --xlsx "output\south_road_updated.xlsx" --project-type road
```

### Push a progress snapshot

To push the project's cumulative actual progress as of a date, provide the
original workbook, progress percentage, and inclusive status date:

```powershell
python scripts/push_to_db.py `
  --xlsx "output\north.xlsx" `
  --project-type road `
  --progress 30 `
  --status-date 2027-02-12
```

This writes a separate progress-snapshot JSON, records the percentage and date
on the project entity, and excludes generated daily-log, supervisor, material,
logistics, and transaction records dated after the status date. The original
planned schedule remains unchanged and included. This is a project-level
snapshot; it does not infer actual progress for individual activities.

The database pusher skips entity IDs that already exist. In progress-snapshot
mode, it updates the existing project's progress and status date, but it does
not delete future-dated records from an earlier full push. The workbook's
generated daily-log data is simulated planning data, not verified field
activity.

### Push a road Excel through the BluConn API

The API route is independent from the direct PostgreSQL command above. It
uses the `BLUCONN_API_URL`, `BLUCONN_ACCOUNT_ID`, and `BLUCONN_AUTH_TOKEN`
environment variables and does not connect to PostgreSQL:

```powershell
python scripts/push_to_api.py --xlsx "output\south_road.xlsx" --project-type road
```

For an existing entity JSON file, provide its source workbook so the same
source-backed validation can run:

```powershell
python scripts/push_to_api.py `
  --json "output\south_road_actual_data.json" `
  --source-xlsx "output\south_road.xlsx"
```

The API route creates/uses entity types and relationships, checks existing
business IDs, batches uploads in parent-before-child order, and writes a
`*_api_upload_report.xlsx` beside the entity JSON. Configure the environment
variables in `.env` or the process environment; credentials are never passed
as command-line arguments.

## Pipeline project

### Create a new pipeline Excel

```powershell
python main.py pipeline `
  --title "South Water Pipeline" `
  --length 12 `
  --diameter 300 `
  --material DI `
  --pipeline-type water `
  --duration "480 days" `
  --budget 450 `
  --start-date 2026-11-10 `
  --output "output\south_pipeline.xlsx"
```

This command creates only `output\south_pipeline.xlsx`.

### Update the pipeline plan

```powershell
python main.py update-plan `
  --input "output\south_pipeline.xlsx" `
  --status-date 2027-01-10 `
  --actual-progress 8.5 `
  --target-date 2028-04-15 `
  --target-progress 100 `
  --output "output\south_pipeline_updated.xlsx"
```

### Push a pipeline Excel to the database

```powershell
python scripts/push_to_db.py --xlsx "output\south_pipeline.xlsx" --project-type pipeline
```

For the updated workbook, use:

```powershell
python scripts/push_to_db.py --xlsx "output\south_pipeline_updated.xlsx" --project-type pipeline
```

### Push a pipeline Excel through the BluConn API

This is a separate route from PostgreSQL; it uses the `BLUCONN_*` settings
configured above:

```powershell
python scripts/push_to_api.py --xlsx "output\south_pipeline.xlsx" --project-type pipeline
```

## Building project

### Create a new building Excel

```powershell
python main.py building `
  --title "South Residential Building" `
  --area 10000 `
  --floors 5 `
  --type residential `
  --construction rcc `
  --quality standard `
  --duration "480 days" `
  --budget 100 `
  --start-date 2026-11-10 `
  --output "output\south_building.xlsx"
```

This command creates only `output\south_building.xlsx`.

### Update the building plan

```powershell
python main.py update-plan `
  --input "output\south_building.xlsx" `
  --status-date 2027-01-10 `
  --actual-progress 8.5 `
  --target-date 2028-04-15 `
  --target-progress 100 `
  --output "output\south_building_updated.xlsx"
```

### Push a building Excel to the database

```powershell
python scripts/push_to_db.py --xlsx "output\south_building.xlsx" --project-type building
```

For the updated workbook, use:

```powershell
python scripts/push_to_db.py --xlsx "output\south_building_updated.xlsx" --project-type building
```

### Push a building Excel through the BluConn API

```powershell
python scripts/push_to_api.py --xlsx "output\south_building.xlsx" --project-type building
```

## What the push command does

Each push command creates entity JSON from the selected Excel and validates it
before upload. The PostgreSQL command and BluConn API command are independent;
choose one route for a project rather than sending the same records through
both. Each route writes an entity-validation report and a route-specific upload
report. Review the validation result and upload counts, then confirm the
project in the destination.

The generated daily logs are simulated planning data, not verified field
activity. An interrupted API write can have an ambiguous outcome if BluConn
accepted it before the connection closed; check the destination before retrying.

The legacy JSON push command is still available:

```powershell
python scripts/push_to_db.py --json "output\project_actual_data.json"
```
