# Construction Project Planning System
## Current Architecture, Inputs, Calculations, Workflows, and Excel Outputs

**Repository:** `RoadConstruction`  
**Document purpose:** Describe what the current application actually does, what
each file is responsible for, what inputs are required for road, pipeline, and
building projects, and what each generated Excel workbook contains.

This document describes the implementation currently present in the
repository. It does not claim that a feature exists merely because a file or
class name suggests it. In particular, the current application generates
planned resource and cost information; it does not yet provide a complete
site-actual attendance, consumption, expenditure, or enterprise-integration
workflow.

---

## 1. Executive summary

The application is a Python command-line planning system for:

1. Road construction projects.
2. Water, sewerage, or gas pipeline projects.
3. Residential or commercial building projects.
4. Updating an existing generated Excel schedule using reported progress.

The main workflow is:

```text
CLI input
  -> ProjectInput model
  -> DataLoader
  -> project guidelines JSON + assumptions YAML
  -> quantity calculation
  -> productivity and duration calculation
  -> CPM dependency scheduling
  -> constraint/resource optimization
  -> cost calculation and escalation
  -> monthly progress and S-curve
  -> validation
  -> Excel workbook
```

The update workflow is:

```text
Existing Excel workbook
  -> WorkbookReader
  -> reconstruct ProjectInput
  -> reload current configuration
  -> freshness comparison
  -> re-derive baseline
  -> compare reported progress
  -> build remaining work
  -> revised CPM and optional acceleration
  -> revised cost and S-curve
  -> validation
  -> updated Excel workbook
```

The project is therefore best described as a **constraint-based construction
schedule, quantity, cost, and reporting engine**.

It is **not yet** a full enterprise construction execution platform because it
does not have a database, API integration, site attendance import, actual
equipment-hours import, actual material-consumption ledger, or actual
expenditure ledger.

---

## 2. Repository structure

```text
RoadConstruction/
├── main.py
├── README.md
├── project_schedule_commands.md
├── audit_calculations.py
├── verify.py
├── project_config/
│   ├── road/
│   │   ├── road_guidelines.json
│   │   └── road_assumptions.yaml
│   ├── pipeline/
│   │   ├── pipeline_guidelines.json
│   │   └── pipeline_assumptions.yaml
│   └── building/
│       ├── building_guidelines.json
│       └── building_assumptions.yaml
├── src/
│   ├── __init__.py
│   ├── models.py
│   ├── data_loader.py
│   ├── quantity_calculator.py
│   ├── duration_calculator.py
│   ├── project_calendar.py
│   ├── dependency_manager.py
│   ├── cpm_engine.py
│   ├── constraint_scheduler.py
│   ├── cost_calculator.py
│   ├── escalation.py
│   ├── progress_calculator.py
│   ├── wbs_generator.py
│   ├── plan_generator.py
│   ├── engineer_daily_log.py
│   ├── excel_report.py
│   ├── workbook_reader.py
│   ├── update_planner.py
│   ├── update_excel_report.py
│   ├── validation_sanity_checks.py
│   └── workbook_metadata.py
└── tests/
    ├── test_building_audit.py
    ├── test_cli_parsers.py
    ├── test_pipeline_audit.py
    ├── test_road_update_plan.py
    └── test_validation_sanity_checks.py
```

Generated `.xlsx` files are output artifacts. They are not required as Python
source dependencies, although they can be retained as reference workbooks or
test fixtures.

---

## 3. Application entry points

The entry point is `main.py`.

### 3.1 New road schedule

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py road `
  --title "Hyderabad Highway 10 km" `
  --length 10 `
  --lanes 2 `
  --terrain plain `
  --duration "365 days" `
  --budget 250 `
  --start-date "2026-04-01" `
  --output "road_project_schedule.xlsx"
```

### 3.2 New pipeline schedule

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py pipeline `
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

### 3.3 New building schedule

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py building `
  --title "Hyderabad Residential Block" `
  --area 2000 `
  --floors 4 `
  --type residential `
  --construction rcc `
  --quality standard `
  --blocks 1 `
  --foundation-type pile `
  --floor-wise-structural `
  --duration "540 days" `
  --budget 6 `
  --start-date "2026-04-01" `
  --output "building_project_schedule.xlsx"
```

### 3.4 Update an existing schedule

```powershell
& ".\.venv-pipeline\Scripts\python.exe" main.py update-plan `
  --input "road_project_schedule.xlsx" `
  --status-date "2026-10-01" `
  --actual-progress 42 `
  --target-date "2027-03-31" `
  --target-progress 100 `
  --output "road_project_schedule_updated.xlsx"
```

The same `update-plan` command can update road, pipeline, or building
workbooks. The project type is read from the existing workbook and used to
select the corresponding configuration.

---

## 4. File-by-file architecture

### 4.1 Root application files

| File | Responsibility |
|---|---|
| `main.py` | Parses CLI commands, builds `ProjectInput`, selects configuration, runs scheduling, validates results, and calls the appropriate Excel generator. Also owns the update-plan orchestration. |
| `README.md` | User-facing overview, installation notes, historical architecture description, and examples. Some historical sheet/module descriptions may be older than the current implementation; the actual current sheet list is defined by `excel_report.py` and `update_excel_report.py`. |
| `project_schedule_commands.md` | Command templates and input explanations for new and update workflows. |
| `audit_calculations.py` | Standalone calculation audit utility. It is not imported by the normal CLI workflow. |
| `verify.py` | Standalone verification gate for reference workbooks. It is not imported by the normal CLI workflow. |
| `.gitignore` | Excludes Python caches, virtual environments, temporary Excel files, and test outputs. |

### 4.2 Core data and configuration files

| File | Responsibility |
|---|---|
| `src/models.py` | Defines `ProjectInput`, `RoadActivity`, `ActivityReference`, `ActivityCostBreakdown`, `ProjectScheduleResult`, `UpdatePlanInput`, and `UpdatePlanResult`. |
| `src/data_loader.py` | Loads JSON guidelines and YAML assumptions, selects plain/hilly catalogues, validates rate references, and emits configuration warnings. |
| `src/workbook_metadata.py` | Stores application version, project type, input summary, configuration hash, and generation timestamp in Excel document properties. Also checks freshness during update. |
| `src/__init__.py` | Re-exports commonly used models and services from the `src` package. |

### 4.3 Calculation and scheduling files

| File | Responsibility |
|---|---|
| `src/quantity_calculator.py` | Routes quantity calculation to road, pipeline, or building logic. Handles geometry, optional scope, crossings, foundation type, and floor-wise structural quantities. |
| `src/duration_calculator.py` | Converts quantity, productivity, crews, and shifts into effort duration. |
| `src/project_calendar.py` | Converts effort days into calendar days and applies working-week, holiday, monsoon, terrain, and night-shift productivity rules. |
| `src/dependency_manager.py` | Checks duplicate activity IDs, missing predecessors, and circular dependencies. |
| `src/cpm_engine.py` | Performs topological scheduling, forward pass, backward pass, float calculation, critical-path detection, and predecessor lag handling. |
| `src/constraint_scheduler.py` | Main solver. Creates activities, calculates baseline schedule/cost, tests additional fronts/crews/shifts, applies optimization, calculates escalation, and returns `ProjectScheduleResult`. |
| `src/cost_calculator.py` | Builds planned rates from source rates and resource assumptions; calculates labour, equipment, fuel, material, wastage, acceleration, overhead/profit treatment, and escalation inputs. |
| `src/escalation.py` | Applies time-based price escalation to the relevant cost components while keeping escalation separately auditable. |
| `src/progress_calculator.py` | Creates monthly planned progress and expenditure values for the optimized schedule. |

### 4.4 Reporting and update files

| File | Responsibility |
|---|---|
| `src/wbs_generator.py` | Builds project, work-package, and activity WBS nodes with cost, duration, dates, weight, and criticality roll-ups. |
| `src/plan_generator.py` | Builds activity-level planned resource rows from the solved activity cost/resource output. It does not re-price or re-quantify activities. |
| `src/engineer_daily_log.py` | Distributes planned activity quantity and planned cost over scheduled working days using calendar productivity weights. Despite the name, the current rows are planned values, not site actuals. |
| `src/excel_report.py` | Generates a new-project workbook containing summary, inputs, assumptions, WBS, activities, CPM, cost, optimization, monthly schedule, S-curve, validation, Plan, and Engineer_DailyLog sheets. |
| `src/workbook_reader.py` | Reads an existing generated workbook, reconstructs project inputs and metadata, reads summary dates/costs, and reports missing or inconsistent workbook data. |
| `src/update_planner.py` | Estimates progress at a status date, builds remaining activities, reruns CPM, applies acceleration if necessary, calculates revised cost, and produces revised monthly progress. |
| `src/update_excel_report.py` | Generates the updated-plan workbook: update summary, baseline schedule, actual progress input, activity status, revised CPM, revised cost, S-curve comparison, feasibility, and validation. |
| `src/validation_sanity_checks.py` | Validates in-memory schedule results and exported workbooks, including quantities, rates, costs, CPM, Plan totals, DailyLog totals, S-curve totals, and update-plan consistency. |

### 4.5 Tests

| File | Coverage focus |
|---|---|
| `tests/test_cli_parsers.py` | CLI argument parsing and project-input construction. |
| `tests/test_pipeline_audit.py` | Pipeline quantities, crossings, rates, costs, and schedule behavior. |
| `tests/test_building_audit.py` | Building quantities, foundation selection, floor-wise structural activities, rates, and escalation behavior. |
| `tests/test_road_update_plan.py` | Road generation, workbook generation, update planning, and updated workbook behavior. |
| `tests/test_validation_sanity_checks.py` | Arithmetic, workbook, Plan, DailyLog, cost, CPM, and reconciliation validations. |

---

## 5. Project configuration architecture

Each project type has two configuration files:

```text
project_config/<project_type>/<project_type>_guidelines.json
project_config/<project_type>/<project_type>_assumptions.yaml
```

### Guidelines JSON

The guidelines file supplies activity catalogue information such as:

- activity identifiers;
- activity names and descriptions;
- WBS/work-package information;
- units;
- specification or standard references;
- predecessor relationships;
- activity-level rate or reference metadata where configured.

### Assumptions YAML

The assumptions file supplies calculation inputs such as:

- activity master;
- productivity rates;
- crew composition;
- equipment profiles;
- equipment hire rates;
- labour rates;
- fuel rates and consumption;
- material rates and recipes;
- wastage;
- calendar;
- escalation;
- optimization limits;
- cost shares and rate build-up settings.

The active duration and cost calculations use the
`activity_master[].productivity_rate` values. A top-level `productivity` block,
where present, is treated by the loader as reference-only and does not
override the activity master.

---

## 6. Required inputs by project type

All projects require:

| Input | Meaning | Example |
|---|---|---|
| `--title` | Project name | `Hyderabad Highway 10 km` |
| `--duration` | Required target duration; accepts days or months | `365 days` |
| `--budget` | Maximum budget in crore | `250` |
| `--start-date` | Planned start date | `2026-04-01` |
| `--output` | Target Excel workbook | `road_project_schedule.xlsx` |

Duration entered in months is converted using approximately 30 days per month.
Budget values entered in lakh/lac are converted into crore.

### 6.1 Road inputs

| CLI input | Required or default | Purpose |
|---|---|---|
| `--length` | Required unless extracted from title | Road length in kilometres. |
| `--lanes` | Default 4; may be extracted from title | Number of lanes. |
| `--terrain` | `plain` or `hilly`; default can be inferred from title | Selects calendar/productivity and, for hilly roads, the hilly activity catalogue. |
| `--service-road-sides` | Default 0 | Optional service-road scope. |
| `--toll-plazas` | Default 0 | Optional toll-plaza count. |
| `--atms` | Default 0 | Optional ATMS count. |
| `--bridges` | Default 0 | Optional bridge count. |
| `--flyovers` | Default 0 | Optional flyover count. |
| `--underpasses` | Default 0 | Optional underpass count. |
| `--retaining-walls-sqm` | Default 0 | Optional retaining-wall area. |
| `--guidelines` | Project default | Override road guidelines JSON. |
| `--assumptions` | Project default | Override road assumptions YAML. |
| `--no-accel-cost` | Flag | Disables time-driven acceleration surcharge intentionally. |

Road-specific calculated scope includes carriageway/formation dimensions,
earthwork, pavement layers, drainage, structures, utilities, and selected
optional structures.

### 6.2 Pipeline inputs

| CLI input | Required or default | Purpose |
|---|---|---|
| `--length` | Required unless extracted from title | Pipeline length in kilometres. |
| `--diameter` | Default 300 mm | Nominal pipe diameter. |
| `--material` | Default `DI`; choices `DI`, `HDPE`, `MS`, `RCC` | Pipe material. |
| `--trench-depth` | Default 2.0 m | Average trench depth. |
| `--trench-width` | Default 0; derived when zero | Excavated trench width. |
| `--installation-method` | Default `open_trench`; choices `open_trench`, `trenchless` | Installation method. |
| `--fronts` | Default 1 | Simultaneous work fronts. |
| `--pipeline-type` | Default `water`; choices `water`, `sewerage`, `gas` | Pipeline service type. |
| `--terrain` | Default `plain` | Calendar/productivity terrain. |
| `--villages` | Default 0 | Village count; can derive scheme-level facilities. |
| `--ohts` | Default 0 | OHT/ESR count. |
| `--wtp-mld` | Default 0 | WTP capacity in MLD. |
| `--ibps` | Default 0 | Intake/booster pumping-station count. |
| `--fhtc` | Default 0 | Functional household tap-connection count. |
| `--road-crossings` | Default 0 | Declared open-cut road crossings. |
| `--highway-crossings` | Default 0 | Declared NH/SH crossings. |
| `--rail-crossings` | Default 0 | Declared railway crossings. |
| `--hdd-crossings` | Default 0 | Declared HDD/trenchless crossings. |
| `--river-crossings` | Default 0 | Declared river crossings. |
| `--guidelines` | Project default | Override pipeline guidelines JSON. |
| `--assumptions` | Project default | Override pipeline assumptions YAML. |
| `--no-accel-cost` | Flag | Disables time-driven acceleration surcharge intentionally. |

Crossing and scheme-level quantities are only meaningful when the corresponding
scope is declared or when the configuration explicitly defines a documented
default/derivation rule.

### 6.3 Building inputs

| CLI input | Required or default | Purpose |
|---|---|---|
| `--area` | Required; total built-up area | Area in square metres by default. |
| `--sqft` | Flag | Treats `--area` as square feet and converts to square metres. |
| `--floors` | Default 4 | Total floors, including ground floor. |
| `--type` | Default `residential`; choices residential/commercial | Building use. |
| `--construction` | Default `rcc`; choices rcc/load_bearing/steel | Structural construction type. |
| `--quality` | Default `standard`; choices basic/standard/premium | Quality/finish level. |
| `--basement` | Flag | Includes basement scope. |
| `--blocks` | Default 0, auto-derived | Number of blocks/towers. |
| `--foundation-type` | Default `pile`; choices pile/raft/isolated_footing | Foundation activity selection. |
| `--floor-wise-structural` | Flag | Generates explicit floor-wise structural activities. |
| `--guidelines` | Project default | Override building guidelines JSON. |
| `--assumptions` | Project default | Override building assumptions YAML. |
| `--no-accel-cost` | Flag | Disables time-driven acceleration surcharge intentionally. |

Building quantities are planning allowances unless engineering drawings and
design take-offs are supplied. Foundation type changes the generated
foundation-related activities. Floor-wise mode makes structural activities
explicit for each floor rather than hiding them in an aggregate.

---

## 7. New-project processing in detail

### Step 1: Parse and normalize input

`main.py` builds a `ProjectInput` object. `ProjectInput.__post_init__`:

- normalizes project type;
- validates project type;
- normalizes road terrain;
- validates optional scope counts;
- validates pipeline and building-specific fields;
- converts or stores project-specific values.

Road titles can provide lane and kilometre values. Pipeline titles can provide
length. Building area is converted from square feet when `--sqft` is supplied.

### Step 2: Load configuration

`DataLoader.load_all()` loads:

```text
guidelines JSON
assumptions YAML
selected terrain catalogue
```

It validates references, missing rates, productivity-source ambiguity, and
configuration warnings before scheduling proceeds.

### Step 3: Calculate quantities

`QuantityCalculator` selects the project-specific calculator:

- `RoadQuantityCalculator`
- `PipelineQuantityCalculator`
- `BuildingQuantityCalculator`

The result is a list of `RoadActivity` objects. The class name
`RoadActivity` is a shared activity model and is used for all project types.

### Step 4: Calculate durations and dates

Productivity, crews, shifts, and quantity determine effort days. The calendar
then accounts for:

- working days per week;
- rest days;
- configured/fixed holidays;
- monsoon productivity;
- terrain productivity;
- night-shift productivity.

CPM calculates early start, early finish, late start, late finish, total float,
free float, and critical status.

### Step 5: Optimize constraints

`ConstraintScheduler` compares the baseline against:

- required duration;
- maximum budget;
- available work fronts;
- crew/resource scaling;
- shift changes;
- fast-tracking/lead-lag options where configured.

The scheduler records optimization iterations and feasibility status rather than
silently returning an infeasible result.

### Step 6: Calculate costs

The cost engine maintains separate components for:

```text
labour
material
equipment
fuel
wastage
acceleration
escalation
```

The active project configurations keep escalation and wastage auditable rather
than hiding them inside the source rate. The pipeline DN600 laying rate remains
configuration-driven and is not changed automatically by the application.

### Step 7: Generate planned reporting rows

`build_plan()` creates activity-level planned resource rows.
`build_daily_log()` distributes active activities over actual working calendar
days. Daily quantities and costs are weighted by calendar productivity factors,
and the daily quantity totals reconcile to the activity quantity.

These are **planned** records. They are not field actuals.

### Step 8: Validate and export

The application validates:

- activity and quantity integrity;
- rates and cost reconciliation;
- predecessor network;
- CPM;
- Plan rows;
- daily log rows;
- monthly progress;
- S-curve totals;
- exported workbook consistency.

Workbook generation stops when hard validation errors are found.

---

## 8. New-project Excel workbook

The current new-project workbook is generated by
`src/excel_report.py`.

The current generated workbook contains these sheets:

| Sheet | Contents |
|---|---|
| `01_Project_Summary` | Project identity, baseline duration/cost, optimized duration/cost, target constraints, finish dates, feasibility, and warnings. |
| `02_Project_Inputs` | Normalized user inputs and project-specific geometry/scope. |
| `03_Assumptions` | Flattened assumptions snapshot, rates, productivity, resource profiles, calendar, escalation, and source references. |
| `04_WBS_Hierarchy` | Project root, work packages, activity rows, cost roll-ups, dates, durations, weights, and criticality. |
| `05_Activities` | Activity catalogue, WBS, quantity, unit, productivity, baseline/optimized resources, dates, costs, and references. |
| `06_Baseline_CPM` | Baseline CPM dates, predecessors, float, and critical-path information. |
| `07_Baseline_Cost` | Baseline activity cost build-up and total. |
| `08_Constraint_Optimization` | Optimization iterations, constraint status, resource changes, duration changes, and cost changes. |
| `09_Optimized_CPM` | Final optimized CPM dates, float, predecessors, and critical status. |
| `10_Optimized_Cost` | Final cost build-up including resource costs, wastage, acceleration, escalation, overhead/profit treatment, and totals. |
| `11_Baseline_vs_Optimized` | Baseline/optimized duration, cost, resource, and schedule comparison. |
| `12_Monthly_Schedule` | Monthly planned progress and activity distribution. |
| `13_S_Curve` | Planned monthly progress, cumulative progress, monthly expenditure, cumulative expenditure, and chart. |
| `14_Critical_Path` | Baseline and optimized critical-path sequences and critical activity details. |
| `15_Validation_Checks` | PASS, WARNING, and ERROR checks persisted into the workbook. |
| `Plan` | Activity-wise planned labour, equipment, fuel, material, productivity, and cost information. |
| `Engineer_DailyLog` | Calendar-day planned quantities, planned resource descriptions, planned costs, and cumulative planned cost. |

### Important distinction

`Engineer_DailyLog` is currently a planned execution log. Its header and
implementation explicitly identify the values as planned and not actual field
consumption.

It does not currently provide complete columns for:

- actual quantity;
- attendance by labour category;
- actual equipment hours;
- actual material issue/consumption;
- actual fuel;
- actual daily expenditure;
- planned-versus-actual variance.

---

## 9. Update-plan Excel workbook

The updated-plan workbook is generated by
`src/update_excel_report.py`.

| Sheet | Contents |
|---|---|
| `01_Update_Summary` | Project identity, status date, actual progress, baseline/revised duration and cost, target date, and feasibility. |
| `02_Baseline_Schedule` | Original activity schedule used as the update reference. |
| `03_Actual_Progress` | Reported aggregate progress and activity-level progress values supplied to the update command. |
| `04_Activity_Status` | Estimated completed, in-progress, and remaining status for each activity. |
| `05_Revised_CPM` | CPM schedule for remaining/revised activities. |
| `06_Revised_Cost` | Remaining and revised cost build-up. |
| `07_S_Curve_Comparison` | Baseline/revised monthly and cumulative progress/expenditure comparison. |
| `08_Target_Feasibility` | Target-date and target-progress feasibility analysis. |
| `09_Validation_Checks` | Update-plan validation results. |

The update command accepts either:

```text
--actual-progress <overall percentage>
```

or an optional JSON activity-progress mapping:

```json
{
  "ACT-01": 100,
  "ACT-02": 65,
  "ACT-03": 20
}
```

This is progress-based replanning. It is not a full daily actual resource and
cost ledger.

---

## 10. Internal data model

### ProjectInput

Stores normalized project scope:

- title;
- project type;
- length/area;
- duration;
- budget;
- start date;
- road lane/terrain/optional scope;
- pipeline diameter/material/trench/crossings/scheme facilities;
- building area/floors/type/construction/quality/basement/blocks/foundation.

### RoadActivity

Represents every calculated activity for every project type:

- activity ID and WBS;
- work package;
- description;
- unit and quantity;
- productivity;
- crew/equipment descriptions;
- duration;
- predecessors;
- CPM dates and float;
- planned dates;
- optimization fields;
- cost breakdown;
- rate/source metadata.

### ActivityCostBreakdown

Stores:

- labour;
- material;
- equipment;
- fuel;
- direct cost;
- planned/source rates;
- overhead/profit treatment;
- wastage;
- acceleration;
- escalation;
- material recipe lines.

### ProjectScheduleResult

Stores:

- project input;
- optimized activities;
- baseline activities;
- baseline and optimized duration/cost;
- critical paths;
- monthly progress;
- feasibility;
- optimization iterations;
- warnings;
- validation report;
- source configuration paths;
- pipeline effective scope when applicable.

### UpdatePlanResult

Stores:

- original baseline result;
- actual aggregate progress;
- baseline progress at status date;
- progress gap;
- activity status;
- revised activities;
- revised duration/cost;
- target feasibility;
- revised monthly progress;
- warnings and limitations.

---

## 11. What the current system does well

The current implementation provides:

- one common CLI for three construction domains;
- project-specific assumptions and activity catalogues;
- quantity calculations for road, pipeline, and building scopes;
- optional road scope activation;
- pipeline crossing inputs and scheme-level facilities;
- building foundation selection;
- explicit floor-wise building structural activities;
- productivity-based duration calculation;
- calendar-aware scheduling;
- dependency and cycle validation;
- baseline CPM and optimized CPM;
- resource/crew/front/shift optimization;
- cost-component traceability;
- escalation calculation;
- WBS roll-ups;
- monthly schedule and S-curve;
- workbook provenance metadata;
- workbook freshness warnings during update;
- workbook-level validation;
- update-plan re-estimation and revised schedule generation.

The existing regression suite currently passes with:

```text
75 passed
```

### Assumption-source enforcement

The runtime now rejects two previously silent engineering fallbacks:

- an activity without `activity_master[].productivity_rate`,
  `productivity_unit`, or `base_crews`;
- an activity catalogue entry without a quantity produced by the selected
  quantity calculator.

These values must come from the selected project assumptions/catalogue or from
an explicit quantity formula. The application no longer substitutes
`100 units/day` or `100 units` for missing engineering values. User-input
defaults and derived geometric/CPM values remain valid: user inputs are
identified in `02_Project_Inputs`, while derived values are calculated from
those inputs and the loaded assumptions.

`DataLoader.errors` separately records missing/invalid activity productivity
and direct rates. The CLI stops before scheduling when these hard configuration
errors exist. Other configuration observations, such as a reference-only
legacy block or a missing citation register, remain visible warnings and are
shown in the console/workbook rather than being silently ignored.

---

## 12. Current limitations against a full execution-management requirement

The following are not fully implemented:

1. Persistent daily actuals storage.
2. Actual labour attendance by trade and date.
3. Actual equipment usage and operating hours by date.
4. Actual material consumption by activity and date.
5. Actual fuel receipts/consumption by date.
6. Actual daily labour/equipment/material/fuel expenditure.
7. Daily planned-versus-actual comparison for every resource.
8. Actual productivity calculated from measured field quantities.
9. Dedicated actuals Excel input sheet with controlled data-entry columns.
10. Duplicate actual-record detection using a project/activity/date/resource key.
11. Enterprise workflow adapters.
12. Database or API persistence.
13. Idempotent import/retry processing.
14. Push integration into an enterprise platform.

The current `update-plan` workflow should therefore be described as:

> Progress-based schedule updating and remaining-work replanning.

It should not be described as:

> Complete actual resource and cost tracking.

---

## 13. Recommended demonstration sequence

For a project demonstration:

1. Show `main.py` and the four commands.
2. Show the selected project configuration directory.
3. Run a road example.
4. Run a pipeline example with declared crossings.
5. Run a building example with foundation and floor-wise options.
6. Open `01_Project_Summary` and explain baseline versus optimized results.
7. Open `02_Project_Inputs` and `03_Assumptions`.
8. Open `05_Activities`, `06_Baseline_CPM`, and `09_Optimized_CPM`.
9. Open `07_Baseline_Cost` and `10_Optimized_Cost`.
10. Open `12_Monthly_Schedule` and `13_S_Curve`.
11. Open `15_Validation_Checks`.
12. Open `Plan` and explain planned resource rows.
13. Open `Engineer_DailyLog` and clearly state that it is planned, not actual.
14. Run `update-plan` using an overall progress percentage and optionally an
    activity-progress JSON file.
15. Show the revised workbook and explain target feasibility.

---

## 14. Final architecture verdict

The current project is a functioning multi-domain construction planning
application with:

```text
input normalization
 + configuration-driven quantity calculation
 + productivity/duration calculation
 + calendar-aware CPM
 + constraint optimization
 + cost/escalation calculation
 + monthly progress
 + Excel reporting
 + update-plan replanning
```

Its current architectural boundary ends at **planned schedule/report
generation and progress-based update planning**. To become a complete
construction execution monitoring platform, the next architectural addition
should be an explicit actuals subsystem rather than a rewrite of the current
planning engine:

```text
planned schedule
  -> actuals input model
  -> validated daily actual records
  -> planned-vs-actual reconciliation
  -> actual cost/productivity analysis
  -> update planner
  -> Excel/database/API outputs
```

The existing calculation and workbook layers can be reused as the planning
baseline for that future actuals subsystem.
