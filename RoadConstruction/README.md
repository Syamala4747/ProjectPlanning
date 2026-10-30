# Road Construction Project Planning & Constraint-Based Scheduling System

A Python-based engineering scheduling and cost estimation system for **Road & Highway Construction Projects** conforming to **MoRTH Specifications (5th Revision)** and **MoRTH Standard Data Book (2019)** with **Constraint-Based CPM Scheduling & Resource Optimization**.

---

## Architecture & Data Flow

```
User Input (Length km, Required Duration, Max Budget, Start Date)
   │
   ▼
DataLoader (guidelines.json + assumptions.yaml)
   │
   ▼
QuantityCalculator (IRC:73 / MoRTH Cross-section Quantities)
   │
   ▼
ConstraintScheduler (Baseline Plan Evaluation)
   ├── Calculate Normal Duration & Cost (Standard Gangs, 1 Front)
   ├── Run CPM Engine (Baseline Critical Path & Floats)
   │
   ▼
Iterative Resource Optimization Loop (If Baseline Duration > Required Duration)
   ├── Step A: Scale Parallel Work Fronts (2, 3, 4, 5+ Fronts)
   ├── Step B: Scale Equipment & Gangs on Critical Bottlenecks (Embankment, Subgrade, DBM, BC, Drainage)
   ├── Step C: Staggered Section Handover (Fast-Tracking / Lead-Lag on Linear Chain)
   ├── Step D: Extended / Double Shifts (1.5 - 2.0 shifts with Overtime Premiums)
   ├── Recalculate: Duration, Cost, CPM, Total Duration, Total Cost
   └── Stop condition: Duration <= Target AND Cost <= Budget (or report Infeasibility)
   │
   ▼
ProgressCalculator (Optimized Monthly Distribution & Planned S-Curve)
   │
   ▼
WBSGenerator (3-Level Tree: Project Root -> Work Packages -> Activities with Roll-ups)
   │
   ▼
ExcelReportGenerator (14-Sheet Executive Workbook with openpyxl Charts & Comparison Tables)
   │
   ▼
road_project_schedule.xlsx
```

---

## Key Features

1. **Hard Constraint Enforcement**: Duration and budget are treated as rigid constraints (maximum limits). The engine iteratively compresses the schedule to meet target deadlines while verifying financial viability.
2. **MoRTH SDB Traceability**: Every activity is linked to official MoRTH specification sections (e.g., Section 201, 301, 305, 401, 406, 502, 505, 801), Standard Data Book item numbers, and page references.
3. **Deterministic CPM Engine with Fast-Tracking**: Performs topological sorting, forward pass, backward pass, float calculations, and zero-float critical path sequence extraction with relationship lags.
4. **Physically Grounded Cost Impact**: Every resource addition (extra equipment fleets, parallel work fronts, mobilization costs, overtime wage premiums, supervision) strictly and realistically increases cost without artificial inflation.
5. **Dynamic 3-Level WBS Hierarchy**: Generates a standardized Work Breakdown Structure (`1.0 Project` -> `1.1 Work Package` -> `1.1.1 Activity`) with bottom-up roll-up calculations for costs, durations, and weights.
6. **Clear Infeasibility Reporting**: If no configuration can satisfy both constraints, clearly outputs: `"No feasible schedule found within the given budget and duration."` with specific constraint violation details.
7. **14-Sheet Excel Workbook Structure**:
   - `01_Project_Summary`: Project Info, BASELINE Section (Duration, Cost, Finish, Status), and OPTIMIZED / CONSTRAINT-BASED Section (Required Dur, Opt Dur, Max Budget, Opt Cost, Statuses).
   - `02_Project_Inputs`: User Constraints clearly labeled (`Required Duration = CONSTRAINT`, `Maximum Budget = CONSTRAINT`) and engineering geometry inputs.
   - `03_Assumptions`: Normative baseline productivities, normal crews, equipment sets, and planned unit rates.
   - `04_WBS_Hierarchy`: Dedicated 3-level Work Breakdown Structure tree with dynamic package roll-up costs, durations, and weights.
   - `05_Activities`: Master activity sheet with WBS Code & Work Package columns and side-by-side Baseline vs. Optimized crews, durations, costs, duration saved, and additional costs.
   - `06_Baseline_CPM`: ONLY the baseline schedule using default resources and normal productivity.
   - `07_Baseline_Cost`: Baseline rate analysis and total baseline cost with WBS alignment (unforced against budget).
   - `08_Constraint_Optimization`: The optimization log showing Iteration 0 (Baseline) through Final Iteration with PASS/FAIL constraint statuses.
   - `09_Optimized_CPM`: The final constraint-based schedule (ES, EF, LS, LF, Total Float, Free Float, Critical Path).
   - `10_Optimized_Cost`: Optimized rate analysis and itemized direct cost, overheads, contractor profit, and weights for the final crashed schedule with WBS alignment.
   - `11_Baseline_vs_Optimized`: Executive comparison table and detailed resource change breakdown.
   - `12_Monthly_Schedule`: Monthly Gantt progress matrix generated from the OPTIMIZED schedule.
   - `13_S_Curve`: Planned monthly and cumulative progress trajectory with embedded line chart.
   - `14_Critical_Path`: Baseline vs. Optimized critical path sequences and critical activities table.


---

## Installation & Requirements

Ensure you have Python 3.8+ installed with the following packages:

```bash
pip install pyyaml openpyxl
```

---

## How to Run

### 1. Interactive Mode:
```bash
python main.py
```
You will be prompted to enter:
```
Project Title: 10 KM 2-Lane Highway Project
# Output: Automatically detects 2-Lane Highway (IRC:SP:73 standard) and 10 km length!
```

### 2. Command-Line Argument Mode:
```bash
# 2-Lane Highway:
python main.py --title "10 KM 2-Lane Highway Project" --duration "300 days" --budget 250 --output 10km_2lane.xlsx

# 6-Lane Expressway:
python main.py --title "50 KM 6-Lane Mega Expressway" --duration "730 days" --budget 3000 --output 50km_6lane.xlsx

# 7-Lane Highway Project:
python main.py --title "10 KM 7-Lane Highway Project" --duration "365 days" --budget 700 --output 10km_7lane.xlsx

# Explicit hilly-terrain model (requires a complete hilly_terrain_standards
# activity catalog in assumptions.yaml):
python main.py --title "25 KM 4-Lane Highway Project" --terrain hilly --duration "800 days" --budget 1000 --output hilly_25km.xlsx

# A title containing "hilly", "hill", or "mountain" also selects hilly.
# Otherwise the planner uses the plain-terrain activity_master by default.
```

---

## Project Structure

```
RoadConstruction/
│
├── main.py                          # Application CLI entry point (new-project + update-plan)
├── assumptions.yaml                 # Baseline rates, norms & calendar assumptions
├── guidelines.json                  # MoRTH Standard Data Book 2019 specification
│
├── src/
│   ├── __init__.py
│   ├── models.py                    # Data models & typed structures
│   ├── data_loader.py               # JSON & YAML parser
│   ├── quantity_calculator.py       # Geometric cross-section quantity engine
│   ├── cost_calculator.py           # Unit rate analysis & resource crashing cost calculator
│   ├── duration_calculator.py       # Productivity-based duration engine
│   ├── cpm_engine.py                # Critical Path Method engine with lead/lag support
│   ├── constraint_scheduler.py      # Core Constraint-Based Optimization Engine
│   ├── dependency_manager.py        # Precedence & cycle validation
│   ├── activity_generator.py        # Scope extraction & activity generator
│   ├── wbs_generator.py             # 3-level hierarchical WBS builder
│   ├── progress_calculator.py       # Monthly period distribution engine
│   ├── s_curve.py                   # S-Curve progression aggregator
│   ├── excel_report.py              # 14-sheet openpyxl Excel workbook builder (new project)
│   ├── workbook_reader.py           # Reads existing workbook to reconstruct project state
│   ├── update_planner.py            # Update engine: compare progress & replan remaining work
│   └── update_excel_report.py       # 8-sheet Excel workbook builder (updated plan)
│
├── tests/
│   └── test_update_plan.py          # Tests for the update-plan workflow
│
└── README.md
```

---

## Update-Plan Command: Revising an Existing Plan with Actual Progress

### Overview

The `update-plan` command lets you revise an already-generated project plan without
regenerating the entire project from scratch. It reads the existing Excel workbook,
compares actual progress with the original baseline, identifies delays, and generates
a revised plan for the remaining work.

The existing new-project command (`python main.py --title ...`) remains unchanged and
continues to produce the same 14-sheet workbook.

### Usage

```bash
python main.py update-plan \
  --input "road_project_schedule.xlsx" \
  --start-date "2026-04-01" \
  --status-date "2026-10-01" \
  --actual-progress 10 \
  --target-date "2027-04-01" \
  --output "road_project_schedule_updated.xlsx"
```

### Required Inputs

| Argument | Description |
|---|---|
| `--input` | Path to the existing Excel project plan workbook |
| `--status-date` | Date on which actual progress is reported (YYYY-MM-DD) |
| `--actual-progress` | Actual cumulative project progress as a percentage (0–100) |
| `--output` | Filename for the updated Excel workbook |

### Optional Inputs

| Argument | Description |
|---|---|
| `--target-date` | Future date by which the target should be achieved (YYYY-MM-DD). **Optional**: If omitted, the optimized completion date from the previous Excel workbook is automatically used. |
| `--start-date` | Original project start date. If omitted, the date from the workbook is used. |
| `--target-progress` | Cumulative progress target at the target date. If omitted, the baseline target for that date is used. |
| `--activity-progress` | JSON file with activity-level progress: `{"ACT-01": 100.0, "ACT-02": 50.0, ...}` |
| `--guidelines` | Path to guidelines.json (default: `guidelines.json`) |
| `--assumptions` | Path to assumptions.yaml (default: `assumptions.yaml`) |
| `--no-accel-cost` | Disable time-driven acceleration cost terms |

### How It Works

1. **Read workbook** — Extracts the project inputs (length, lanes, terrain, start date,
   duration, budget) from the existing workbook's Project Inputs sheet.
2. **Re-derive baseline** — Re-runs the deterministic `ConstraintScheduler` with the same
   inputs to reconstruct the full baseline schedule, activities, costs, and S-curve.
3. **Compare progress** — Computes the baseline cumulative progress at the status date
   and compares it with the actual progress to identify the progress gap.
4. **Estimate activity completion** — If activity-level progress is provided, uses it
   directly. Otherwise, scales each activity's baseline completion by the overall
   actual/baseline ratio and reports this as a limitation.
5. **Replan remaining work** — Builds a remaining-work activity list (completed
   activities get zero duration; in-progress activities get duration proportional to
   remaining quantity) and re-runs the existing CPM engine from the status date.
6. **Check feasibility & accelerate** — If the revised schedule exceeds the target date,
   applies the same least-cost crashing logic (add crews/shifts on critical activities)
   to compress the remaining schedule. Reports whether the target is achievable.
7. **Generate output** — Produces a separate 8-sheet workbook (the original is never
   overwritten).

### Updated Excel Output (8 Sheets)

| Sheet | Content |
|---|---|
| `01_Update_Summary` | Executive dashboard: progress gap, revised duration, feasibility, costs |
| `02_Baseline_Schedule` | Original optimized CPM schedule from the existing workbook |
| `03_Actual_Progress` | Per-activity baseline vs actual completion comparison |
| `04_Activity_Status` | Completed / remaining quantities for each activity |
| `05_Revised_CPM` | Revised CPM schedule for remaining work (durations, dates, critical path) |
| `06_Revised_Cost` | Remaining-work cost, acceleration surcharge, and total project cost |
| `07_S_Curve_Comparison` | Baseline, actual, and revised forecast S-curves with chart (delay days highlighted in distinct red) |
| `08_Target_Feasibility` | Feasibility verdict, projected delay, and resource changes applied |

### Activity-Level Progress

For accurate activity-level replanning, provide a JSON file mapping activity IDs to
completion percentages:

```json
{
  "ACT-01": 100.0,
  "ACT-02": 100.0,
  "ACT-05": 50.0,
  "ACT-06": 0.0
}
```

When this is omitted, the system estimates per-activity completion by scaling baseline
completions proportionally and reports the limitation in the output.

### Running Tests

```bash
python -m pytest tests/test_update_plan.py -v
```
