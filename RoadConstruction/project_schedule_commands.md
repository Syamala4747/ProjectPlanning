# Project Schedule Command Templates

This document is the command reference for generating and updating Road,
Pipeline, and Building Excel schedules.

Run commands from:

```powershell
cd C:\Users\syama\JAVA\Desktop\road\Roadgit\ProjectPlanning\RoadConstruction
```

All generated Excel workbooks are stored in:

```text
C:\Users\syama\JAVA\Desktop\road\Roadgit\ProjectPlanning\RoadConstruction\output\
```

When `--output "file.xlsx"` is used, the application writes
`output\file.xlsx`. The `output` folder is created automatically. An absolute
path may be supplied when a different location is intentionally required.

Use the project virtual environment directly:

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py <command>
```

## Field labels used in this document

- **MANDATORY**: provide this value for a reproducible project schedule.
- **OPTIONAL**: provide it only when that scope or behavior exists.
- **DEFAULT**: the application supplies this value when omitted.
- **DO NOT ADD**: do not use this option unless you intentionally want to
  override the normal behavior.

Although some CLI fields have prompts or defaults, provide all **MANDATORY**
fields explicitly in production commands. Do not rely on interactive prompts.

## 1. Road project

### 1.1 New road workbook fields

| Field | Status | Meaning | Example |
|---|---|---|---|
| `--title` | MANDATORY | Project name | `"NH Link Road 10 km"` |
| `--length` | MANDATORY | Length in kilometres | `10` |
| `--lanes` | MANDATORY | Number of lanes | `4` |
| `--terrain` | MANDATORY | `plain` or `hilly` | `plain` |
| `--duration` | MANDATORY | Maximum duration | `"300 days"` |
| `--budget` | MANDATORY | Maximum budget in crore | `250` |
| `--start-date` | MANDATORY | Start date, `YYYY-MM-DD` | `"2026-11-04"` |
| `--output` | MANDATORY | New Excel filename | `"road_schedule.xlsx"` |
| `--service-road-sides` | OPTIONAL | Service-road sides; `0` means none | `2` |
| `--toll-plazas` | OPTIONAL | Number of toll plazas | `1` |
| `--atms` | OPTIONAL | Number of ATMS scopes | `1` |
| `--bridges` | OPTIONAL | Number of bridges | `1` |
| `--flyovers` | OPTIONAL | Number of flyovers | `0` |
| `--underpasses` | OPTIONAL | Number of underpasses | `1` |
| `--retaining-walls-sqm` | OPTIONAL | Retaining-wall area in sqm | `5000` |
| `--no-accel-cost` | OPTIONAL | Disable acceleration cost | flag only |

Optional structure activities are generated only when their input is greater
than zero. Do not add optional fields with invented quantities.

### 1.2 Road new-workbook template

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py road `
  --title "<ROAD PROJECT TITLE>" `
  --length <LENGTH_KM> `
  --lanes <2_OR_4> `
  --terrain <plain_OR_hilly> `
  --duration "<DURATION days>" `
  --budget <BUDGET_CRORE> `
  --start-date "<YYYY-MM-DD>" `
  --output "<ROAD_NEW_FILE>.xlsx" `
  --service-road-sides <0_OR_1_OR_2> `
  --toll-plazas <0_OR_MORE> `
  --atms <0_OR_MORE> `
  --bridges <0_OR_MORE> `
  --flyovers <0_OR_MORE> `
  --underpasses <0_OR_MORE> `
  --retaining-walls-sqm <0_OR_MORE>
```

### 1.3 Road update-workbook template

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py update-plan `
  --input "<ROAD_NEW_OR_EXISTING_FILE>.xlsx" `
  --status-date "<YYYY-MM-DD>" `
  --actual-progress <0_TO_100> `
  --output "<ROAD_UPDATED_FILE>.xlsx"
```

`--input`, `--status-date`, `--actual-progress`, and `--output` are mandatory.
The update command reads the original target date from the workbook. Do not
add `--target-date` or `--target-progress` unless a new contractual target is
intentionally being tested.

Optional update fields:

```text
--start-date <YYYY-MM-DD>
--activity-progress <activity-progress.json>
--no-accel-cost
```

## 2. Pipeline project

### 2.1 New pipeline workbook fields

| Field | Status | Meaning | Example |
|---|---|---|---|
| `--title` | MANDATORY | Project name | `"City Water Main"` |
| `--length` | MANDATORY | Pipeline length in kilometres | `10` |
| `--diameter` | MANDATORY | Pipe diameter in millimetres | `600` |
| `--material` | MANDATORY | `DI`, `HDPE`, `MS`, or `RCC` | `DI` |
| `--trench-depth` | MANDATORY | Average depth in metres | `2` |
| `--fronts` | MANDATORY | Simultaneous work fronts | `2` |
| `--duration` | MANDATORY | Maximum duration | `"365 days"` |
| `--budget` | MANDATORY | Maximum budget in crore | `25` |
| `--start-date` | MANDATORY | Start date, `YYYY-MM-DD` | `"2026-04-01"` |
| `--output` | MANDATORY | New Excel filename | `"pipeline_schedule.xlsx"` |
| `--pipeline-type` | OPTIONAL | `water`, `sewerage`, or `gas` | `water` |
| `--trench-width` | OPTIONAL | Width in metres; `0` derives it | `0` |
| `--installation-method` | OPTIONAL | `open_trench` or `trenchless` | `open_trench` |
| `--terrain` | OPTIONAL | `plain` or `hilly` | `plain` |
| `--villages` | OPTIONAL | Villages served | `0` |
| `--ohts` | OPTIONAL | OHT/ESR count | `0` |
| `--wtp-mld` | OPTIONAL | WTP capacity in MLD | `0` |
| `--ibps` | OPTIONAL | Intake/booster stations | `0` |
| `--fhtc` | OPTIONAL | Household tap connections | `0` |
| `--road-crossings` | OPTIONAL | Declared road crossings | `0` |
| `--highway-crossings` | OPTIONAL | Declared highway crossings | `0` |
| `--rail-crossings` | OPTIONAL | Declared railway crossings | `0` |
| `--hdd-crossings` | OPTIONAL | Declared HDD crossings | `0` |
| `--river-crossings` | OPTIONAL | Declared river crossings | `0` |

Use declared crossing counts. Do not enter crossings that are not confirmed by
the route survey or engineering scope.

### 2.2 Pipeline new-workbook template

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py pipeline `
  --title "<PIPELINE PROJECT TITLE>" `
  --length <LENGTH_KM> `
  --diameter <DIAMETER_MM> `
  --material <DI_OR_HDPE_OR_MS_OR_RCC> `
  --trench-depth <DEPTH_METRES> `
  --fronts <NUMBER_OF_FRONTS> `
  --duration "<DURATION days>" `
  --budget <BUDGET_CRORE> `
  --start-date "<YYYY-MM-DD>" `
  --output "<PIPELINE_NEW_FILE>.xlsx" `
  --pipeline-type <water_OR_sewerage_OR_gas> `
  --road-crossings <0_OR_MORE> `
  --highway-crossings <0_OR_MORE> `
  --rail-crossings <0_OR_MORE> `
  --hdd-crossings <0_OR_MORE> `
  --river-crossings <0_OR_MORE>
```

### 2.3 Pipeline update-workbook template

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py update-plan `
  --input "<PIPELINE_NEW_OR_EXISTING_FILE>.xlsx" `
  --status-date "<YYYY-MM-DD>" `
  --actual-progress <0_TO_100> `
  --output "<PIPELINE_UPDATED_FILE>.xlsx"
```

Do not add `--target-date` or `--target-progress` for the normal update
workflow. The stored workbook target is used automatically.

## 3. Building project

### 3.1 New building workbook fields

| Field | Status | Meaning | Example |
|---|---|---|---|
| `--title` | MANDATORY | Project name | `"Residential Block A"` |
| `--area` | MANDATORY | Built-up area in sqm | `2000` |
| `--floors` | MANDATORY | Total floors including ground; G+3 = `4` | `4` |
| `--type` | MANDATORY | `residential` or `commercial` | `residential` |
| `--construction` | MANDATORY | `rcc`, `load_bearing`, or `steel` | `rcc` |
| `--quality` | MANDATORY | `basic`, `standard`, or `premium` | `standard` |
| `--foundation-type` | MANDATORY | `pile`, `raft`, or `isolated_footing` | `pile` |
| `--duration` | MANDATORY | Maximum duration | `"540 days"` |
| `--budget` | MANDATORY | Maximum budget in crore | `6` |
| `--start-date` | MANDATORY | Start date, `YYYY-MM-DD` | `"2026-04-01"` |
| `--output` | MANDATORY | New Excel filename | `"building_schedule.xlsx"` |
| `--blocks` | OPTIONAL | Block count; `0` auto-derives | `1` |
| `--sqft` | OPTIONAL | Interpret `--area` as square feet | flag only |
| `--basement` | OPTIONAL | Include basement | flag only |
| `--floor-wise-structural` | OPTIONAL | Explicit structural activity per floor | flag only |
| `--no-accel-cost` | OPTIONAL | Disable acceleration cost | flag only |

Use `--floor-wise-structural` only when the building activity catalogue has
productivity assumptions for every generated floor activity.

### 3.2 Building new-workbook template

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py building `
  --title "<BUILDING PROJECT TITLE>" `
  --area <BUILT_UP_AREA_SQM> `
  --floors <TOTAL_FLOORS> `
  --type <residential_OR_commercial> `
  --construction <rcc_OR_load_bearing_OR_steel> `
  --quality <basic_OR_standard_OR_premium> `
  --foundation-type <pile_OR_raft_OR_isolated_footing> `
  --duration "<DURATION days>" `
  --budget <BUDGET_CRORE> `
  --start-date "<YYYY-MM-DD>" `
  --output "<BUILDING_NEW_FILE>.xlsx" `
  --blocks <0_OR_NUMBER_OF_BLOCKS>
```

### 3.3 Building update-workbook template

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py update-plan `
  --input "<BUILDING_NEW_OR_EXISTING_FILE>.xlsx" `
  --status-date "<YYYY-MM-DD>" `
  --actual-progress <0_TO_100> `
  --output "<BUILDING_UPDATED_FILE>.xlsx"
```

Do not add `--target-date` or `--target-progress` for the normal update
workflow.

## 4. Real project command set

The following six commands create one complete example pair for each project
type. Replace `--actual-progress 42` with the measured progress from the
project engineer.

### 4.1 Real road project: 10 km plain 4-lane highway

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py road `
  --title "Hyderabad Outer Link Plain 4-Lane Highway 10 km" `
  --length 10 `
  --lanes 4 `
  --terrain plain `
  --duration "300 days" `
  --budget 250 `
  --start-date "2026-11-04" `
  --service-road-sides 2 `
  --bridges 1 `
  --underpasses 1 `
  --retaining-walls-sqm 5000 `
  --output "hyderabad_plain_4lane_10km_schedule.xlsx"

& ".\.venv-pipeline\Scripts\python.exe" main.py update-plan `
  --input "hyderabad_plain_4lane_10km_schedule.xlsx" `
  --status-date "2026-11-04" `
  --actual-progress 0 `
  --output "hyderabad_plain_4lane_10km_schedule_updated.xlsx"
```

### 4.2 Real pipeline project: 10 km DN600 DI water main

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py pipeline `
  --title "Hyderabad City DN600 DI Water Main 10 km" `
  --length 10 `
  --diameter 600 `
  --material DI `
  --trench-depth 2 `
  --fronts 2 `
  --pipeline-type water `
  --duration "365 days" `
  --budget 25 `
  --start-date "2026-04-01" `
  --road-crossings 12 `
  --highway-crossings 2 `
  --rail-crossings 1 `
  --hdd-crossings 2 `
  --river-crossings 0 `
  --output "hyderabad_dn600_pipeline_10km_schedule.xlsx"

& ".\.venv-pipeline\Scripts\python.exe" main.py update-plan `
  --input "hyderabad_dn600_pipeline_10km_schedule.xlsx" `
  --status-date "2026-11-04" `
  --actual-progress 42 `
  --output "hyderabad_dn600_pipeline_10km_schedule_updated.xlsx"
```

### 4.3 Real building project: 2,000 sqm G+3 residential block

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py building `
  --title "Hyderabad Residential Block G+3" `
  --area 2000 `
  --floors 4 `
  --type residential `
  --construction rcc `
  --quality standard `
  --foundation-type pile `
  --blocks 1 `
  --duration "540 days" `
  --budget 6 `
  --start-date "2026-04-01" `
  --output "hyderabad_residential_g3_schedule.xlsx"

& ".\.venv-pipeline\Scripts\python.exe" main.py update-plan `
  --input "hyderabad_residential_g3_schedule.xlsx" `
  --status-date "2026-11-04" `
  --actual-progress 42 `
  --output "hyderabad_residential_g3_schedule_updated.xlsx"
```

## 5. Activity-level progress

Overall progress is acceptable for a basic update. For a more accurate update,
create a JSON file whose activity IDs exist in the original workbook:

```json
{
  "ACT-01": 100,
  "ACT-02": 100,
  "ACT-03": 65,
  "ACT-04": 0
}
```

Then add this optional argument:

```powershell
--activity-progress "activity_progress.json"
```

All values must be between `0` and `100`.

## 6. Output and validation

New workbooks contain the `15_Validation_Checks` sheet. Updated workbooks
contain `09_Validation_Checks`.

- `PASS`: arithmetic, schedule, or reconciliation check passed.
- `WARNING`: planning or engineering review is required; totals are not
  automatically invalid.
- `ERROR`: hard validation failure; investigate before accepting the workbook.

If the requested output workbook is open in Excel, the application writes a
safe alternative filename such as `_new.xlsx` in the same `output` folder.
Close the workbook and use the actual filename printed by the command.

To inspect all available options:

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py road --help
& ".\.venv-pipeline\Scripts\python.exe" main.py pipeline --help
& ".\.venv-pipeline\Scripts\python.exe" main.py building --help
& ".\.venv-pipeline\Scripts\python.exe" main.py update-plan --help
```
