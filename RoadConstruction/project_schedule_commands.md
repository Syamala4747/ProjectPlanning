# Project Schedule Commands

This guide shows how to:

1. Prepare input values for a new schedule.
2. Generate a new Excel schedule for a road, pipeline, or building project.
3. Update an existing Excel schedule for each project type.

Run all commands from:

```powershell
cd C:\Users\syama\JAVA\Desktop\road\Roadgit\ProjectPlanning\RoadConstruction
```

Use the project virtual environment:

```powershell
& ".\.venv-pipeline\Scripts\Activate.ps1"
```

If PowerShell blocks activation, run every command with the environment's
Python executable instead:

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py ...
```

## 1. New schedule input templates

### 1.1 Road project inputs

| Input | Meaning | Example |
|---|---|---|
| `--title` | Project name | `"Hyderabad Highway"` |
| `--length` | Road length in kilometres | `10` |
| `--lanes` | Number of lanes | `2` |
| `--terrain` | `plain` or `hilly` | `plain` |
| `--duration` | Planned duration | `"365 days"` |
| `--budget` | Budget in crore | `250` |
| `--start-date` | Start date in `YYYY-MM-DD` format | `2026-04-01` |
| `--output` | Output Excel file | `road_schedule.xlsx` |

Road command template:

```powershell
python main.py road `
  --title "<ROAD PROJECT TITLE>" `
  --length <LENGTH_KM> `
  --lanes <NUMBER_OF_LANES> `
  --terrain plain `
  --duration "<DURATION> days" `
  --budget <BUDGET_CRORE> `
  --start-date "<YYYY-MM-DD>" `
  --output "<ROAD_OUTPUT_FILE>.xlsx"
```

Use `--terrain hilly` for a hilly project. Use `--no-accel-cost` only when
acceleration cost should intentionally be disabled.

### 1.2 Pipeline project inputs

| Input | Meaning | Example |
|---|---|---|
| `--title` | Project name | `"Hyderabad Water Pipeline"` |
| `--length` | Pipeline length in kilometres | `10` |
| `--diameter` | Nominal pipe diameter in millimetres | `600` |
| `--material` | `DI`, `HDPE`, `MS`, or `RCC` | `DI` |
| `--trench-depth` | Average trench depth in metres | `2` |
| `--fronts` | Simultaneous work fronts | `2` |
| `--pipeline-type` | `water`, `sewerage`, or `gas` | `water` |
| `--duration` | Planned duration | `"365 days"` |
| `--budget` | Budget in crore | `25` |
| `--start-date` | Start date in `YYYY-MM-DD` format | `2026-04-01` |
| `--output` | Output Excel file | `pipeline_schedule.xlsx` |

Optional pipeline inputs include `--villages`, `--ohts`, `--wtp-mld`,
`--ibps`, `--fhtc`, `--road-crossings`, `--highway-crossings`,
`--rail-crossings`, `--hdd-crossings`, and `--river-crossings`.

Pipeline command template:

```powershell
python main.py pipeline `
  --title "<PIPELINE PROJECT TITLE>" `
  --length <LENGTH_KM> `
  --diameter <DIAMETER_MM> `
  --material <DI|HDPE|MS|RCC> `
  --trench-depth <DEPTH_METRES> `
  --fronts <NUMBER_OF_FRONTS> `
  --pipeline-type <water|sewerage|gas> `
  --duration "<DURATION> days" `
  --budget <BUDGET_CRORE> `
  --start-date "<YYYY-MM-DD>" `
  --output "<PIPELINE_OUTPUT_FILE>.xlsx"
```

### 1.3 Building project inputs

| Input | Meaning | Example |
|---|---|---|
| `--title` | Project name | `"Hyderabad Residential Block"` |
| `--area` | Total built-up area in square metres | `2000` |
| `--floors` | Total floors, including ground floor | `4` |
| `--type` | `residential` or `commercial` | `residential` |
| `--construction` | `rcc`, `load_bearing`, or `steel` | `rcc` |
| `--quality` | `basic`, `standard`, or `premium` | `standard` |
| `--blocks` | Number of blocks; `0` means auto-derive | `1` |
| `--duration` | Planned duration | `"540 days"` |
| `--budget` | Budget in crore | `6` |
| `--start-date` | Start date in `YYYY-MM-DD` format | `2026-04-01` |
| `--output` | Output Excel file | `building_schedule.xlsx` |

Use `--sqft` only when the value passed to `--area` is in square feet.
Use `--basement` when a basement is included.

Building command template:

```powershell
python main.py building `
  --title "<BUILDING PROJECT TITLE>" `
  --area <BUILT_UP_AREA_SQM> `
  --floors <NUMBER_OF_FLOORS> `
  --type <residential|commercial> `
  --construction <rcc|load_bearing|steel> `
  --quality <basic|standard|premium> `
  --blocks <NUMBER_OF_BLOCKS> `
  --duration "<DURATION> days" `
  --budget <BUDGET_CRORE> `
  --start-date "<YYYY-MM-DD>" `
  --output "<BUILDING_OUTPUT_FILE>.xlsx"
```

## 2. New schedule commands with sample input

### 2.1 Road

```powershell
python main.py road `
  --title "Hyderabad Highway 10 km" `
  --length 10 `
  --lanes 2 `
  --terrain plain `
  --duration "365 days" `
  --budget 250 `
  --start-date "2026-04-01" `
  --output "road_project_schedule.xlsx"
```

### 2.2 Pipeline

```powershell
python main.py pipeline `
  --title "Hyderabad Water Pipeline 10 km" `
  --length 10 `
  --diameter 600 `
  --material DI `
  --trench-depth 2 `
  --fronts 2 `
  --pipeline-type water `
  --duration "365 days" `
  --budget 25 `
  --start-date "2026-04-01" `
  --output "pipeline_project_schedule.xlsx"
```

### 2.3 Building

```powershell
python main.py building `
  --title "Hyderabad Residential Block" `
  --area 2000 `
  --floors 4 `
  --type residential `
  --construction rcc `
  --quality standard `
  --blocks 1 `
  --duration "540 days" `
  --budget 6 `
  --start-date "2026-04-01" `
  --output "building_project_schedule.xlsx"
```

## 3. Updating an existing Excel schedule

The update command is the same for road, pipeline, and building workbooks.
The application reads the project type from the existing workbook and
automatically loads the matching guidelines and assumptions.

Required update inputs:

| Input | Meaning | Example |
|---|---|---|
| `--input` | Existing generated Excel schedule | `road_project_schedule.xlsx` |
| `--status-date` | Date on which progress is reported | `2026-10-01` |
| `--actual-progress` | Actual overall progress from 0 to 100 | `42` |
| `--output` | Updated Excel schedule | `road_project_schedule_updated.xlsx` |

Optional update inputs:

| Input | Meaning |
|---|---|
| `--start-date` | Override the original project start date |
| `--target-date` | Date by which the target should be achieved |
| `--target-progress` | Target progress percentage at the target date |
| `--activity-progress` | JSON file containing activity-level progress |
| `--no-accel-cost` | Disable acceleration cost intentionally |

### 3.1 Update command template

```powershell
python main.py update-plan `
  --input "<EXISTING_SCHEDULE>.xlsx" `
  --status-date "<STATUS_DATE>" `
  --actual-progress <ACTUAL_PROGRESS_PERCENT> `
  --target-date "<TARGET_DATE>" `
  --target-progress <TARGET_PROGRESS_PERCENT> `
  --output "<UPDATED_SCHEDULE>.xlsx"
```

`--target-date` and `--target-progress` are optional. If `--target-date` is
not supplied, the command uses the optimized finish date already stored in the
existing workbook when available.

## 4. Update commands with sample input

### 4.1 Update a road schedule

```powershell
python main.py update-plan `
  --input "road_project_schedule.xlsx" `
  --status-date "2026-10-01" `
  --actual-progress 42 `
  --target-date "2027-03-31" `
  --target-progress 100 `
  --output "road_project_schedule_updated.xlsx"
```

### 4.2 Update a pipeline schedule

```powershell
python main.py update-plan `
  --input "pipeline_project_schedule.xlsx" `
  --status-date "2026-10-01" `
  --actual-progress 35 `
  --target-date "2027-03-31" `
  --target-progress 100 `
  --output "pipeline_project_schedule_updated.xlsx"
```

### 4.3 Update a building schedule

```powershell
python main.py update-plan `
  --input "building_project_schedule.xlsx" `
  --status-date "2026-10-01" `
  --actual-progress 28 `
  --target-date "2027-09-30" `
  --target-progress 100 `
  --output "building_project_schedule_updated.xlsx"
```

## 5. Activity-level progress template

Overall progress is enough for a basic update. For a more accurate update,
create a JSON file with activity IDs and their completion percentages:

```json
{
  "ACT-01": 100,
  "ACT-02": 100,
  "ACT-03": 65,
  "ACT-04": 0
}
```

The IDs must exist in the original workbook. Values must be between `0` and
`100`.

Use it with any project type:

```powershell
python main.py update-plan `
  --input "pipeline_project_schedule.xlsx" `
  --status-date "2026-10-01" `
  --actual-progress 35 `
  --activity-progress "pipeline_activity_progress.json" `
  --output "pipeline_project_schedule_updated.xlsx"
```

## 6. What the commands produce

For a new project, the generated workbook contains the complete schedule,
quantities, costs, CPM dates, monthly progress, S-curve, and validation
register.

For an updated project, the generated workbook contains the baseline-versus-
actual comparison, remaining activities, revised CPM, revised cost, revised
monthly progress, feasibility information, and update validation.

Always review the validation sheet:

- New workbook: `15_Validation_Checks`
- Updated workbook: `09_Validation_Checks`

The validation sheet should show `PASS`, `WARNING`, or `ERROR` for each check.
An `ERROR` means the workbook should not be accepted until the issue is
understood and corrected.

## 7. Check available options

```powershell
python main.py road --help
python main.py pipeline --help
python main.py building --help
python main.py update-plan --help
```

