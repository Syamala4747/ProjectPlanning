# Create, Update, and Push a Project

Run these commands in PowerShell from the `RoadConstruction` folder.

## One-time setup

Use Python 3.12 and install the required packages:

```powershell
py -3.12 -m venv .venv312
.\.venv312\Scripts\Activate.ps1
python -m pip install openpyxl pyyaml psycopg2-binary
```

For database pushes, configure `DATABASE_URL` in the `.env` file. Keep the
database password private.

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

## What the push command does

The push command creates the entity JSON from the selected Excel, validates it,
and only sends records to the database if validation allows it. It also writes
the entity-validation and database-push Excel reports beside the workbook.
Review the validation report and push summary after running the command.

The legacy JSON push command is still available:

```powershell
python scripts/push_to_db.py --json "output\project_actual_data.json"
```
