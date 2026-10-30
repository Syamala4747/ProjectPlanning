"""
Main Entry Point for Road Construction Project Planning & Constraint-Based Scheduling.
Enforces hard duration and budget constraints through iterative resource optimization.
"""
import os
import sys
import re
import json
import argparse
from datetime import date, datetime

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from src.models import (ProjectInput, ProjectScheduleResult,
                        parse_lanes_from_title, parse_length_from_title,
                        parse_terrain_from_title,
                        UpdatePlanInput)
from src.data_loader import DataLoader
from src.dependency_manager import DependencyManager
from src.constraint_scheduler import ConstraintScheduler
from src.excel_report import ExcelReportGenerator
from src.workbook_reader import WorkbookReader
from src.update_planner import UpdatePlanner
from src.update_excel_report import UpdateExcelReportGenerator

def parse_duration_input(raw_input: str) -> int:
    """Parses duration string like '365', '365 days', '12 months' into days."""
    s = str(raw_input).strip().lower()
    match_months = re.search(r"(\d+(\.\d+)?)\s*(months?|m)", s)
    if match_months:
        months_val = float(match_months.group(1))
        return int(round(months_val * 30))  # 30 days per month planning convention
    
    match_days = re.search(r"(\d+(\.\d+)?)", s)
    if match_days:
        return int(round(float(match_days.group(1))))
    
    raise ValueError(f"Unable to parse duration from: '{raw_input}'")

def parse_budget_input(raw_input: str) -> float:
    """Parses budget string like '300', '300 crore', '300 cr' into float in Crore."""
    s = str(raw_input).strip().lower()
    match = re.search(r"(\d+(\.\d+)?)", s)
    if match:
        return float(match.group(1))
    raise ValueError(f"Unable to parse budget from: '{raw_input}'")

def parse_start_date(raw_input: str) -> date:
    """Parses date string in YYYY-MM-DD format, or defaults to today."""
    s = str(raw_input).strip()
    if not s:
        return date.today()
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return date.today()

def run_project_planning(
    project_input: ProjectInput,
    guidelines_path: str = "road_guidelines.json",
    assumptions_path: str = "road_assumptions.yaml",
    output_path: str = "road_project_schedule.xlsx",
    enable_acceleration_cost: bool = True
) -> ProjectScheduleResult:
    """
    Executes the constraint-based planning pipeline:
    1. Loads MoRTH guidelines and assumptions.
    2. Runs ConstraintScheduler to compute baseline and iteratively optimize to meet duration and budget.
    3. Generates the executive-grade 14-sheet Excel report.

    `enable_acceleration_cost=False` forces every NEW time-driven acceleration
    term to zero (CLI: --no-accel-cost), reproducing the legacy cost behaviour
    on identical input so the two cost models can be compared directly.
    """
    # 1. Load reference files
    print("⏳ [1/5] Loading MoRTH guidelines and assumptions...")
    loader = DataLoader(guidelines_path=guidelines_path, assumptions_path=assumptions_path)
    loader.load_all(terrain=project_input.terrain)
    print("    ✔ Guidelines & assumptions loaded successfully.")
    if loader.warnings:
        print("    ⚠ DATA-INTEGRITY WARNINGS:")
        for w in loader.warnings:
            print(f"      • {w}")

    if not enable_acceleration_cost:
        print("    ⚠ --no-accel-cost: time-driven acceleration surcharges are DISABLED "
              "(legacy cost behaviour).")

    # 2. Run Constraint-Based Optimizer
    print("⏳ [2/5] Running Constraint-Based Optimizer & CPM Calculation...")
    scheduler = ConstraintScheduler(
        project_input=project_input,
        data_loader=loader,
        enable_acceleration_cost=enable_acceleration_cost
    )
    result = scheduler.solve()

    print("================================================================================")
    print(" 🛣️  ROAD CONSTRUCTION CONSTRAINT-BASED SCHEDULING PIPELINE")
    print("================================================================================")
    print(f"Project Title:              {project_input.title}")
    print(f"Terrain:                    {project_input.terrain.title()}")
    print(f"Highway Configuration:      {project_input.lanes}-Lane Highway ({project_input.carriageway_width_m:.1f}m Carriageway, {project_input.formation_width_m:.1f}m Formation)")
    print(f"Project Length:             {project_input.length_km:.2f} km")
    print(f"Required Target Duration:   {project_input.duration_days} days (Hard Maximum Constraint)")
    print(f"Approved Maximum Budget:    ₹ {project_input.budget_crore:,.2f} Crore (Hard Ceiling Constraint)")
    print(f"Project Start Date:         {project_input.start_date}")
    print(" Calendar rules in force (road_assumptions.yaml -> calendar):")
    for _line in scheduler.calendar.describe():
        print(f"   • {_line}")
    print("--------------------------------------------------------------------------------")

    # 3. Validate network dependencies
    print("⏳ [3/5] Validating optimized activity precedence network...")
    errors = DependencyManager.validate_network(result.activities)
    if errors:
        for err in errors:
            print(f"    ❌ Network Error: {err}")
        raise ValueError("Network dependency validation failed.")
    print("    ✔ Dependency network is cycle-free and fully connected.")

    # 4. Display baseline vs optimized audit summary
    print("\n--------------------------------------------------------------------------------")
    print(" 📊 CONSTRAINT-BASED SCHEDULING RESULTS & AUDIT TRAIL")
    print("--------------------------------------------------------------------------------")
    print(f" Highway Type:                   {project_input.lanes}-Lane Highway Standard")
    print(f" Baseline Duration (Normal):     {result.baseline_duration_days} Days")
    print(f" Baseline Cost:                  ₹ {result.baseline_cost_crore:,.2f} Crore")
    print(f" Optimized Duration:             {result.calculated_duration_days} Days (Target: <= {project_input.duration_days} Days)")
    print(f" Optimized Total Cost:           ₹ {result.total_calculated_cost_crore:,.2f} Crore (Budget: <= ₹ {project_input.budget_crore:,.2f} Crore)")
    print(f" Duration Saved:                 {result.duration_saved_days} Days")
    print(f" Additional Optimization Cost:   ₹ {result.additional_cost_crore:,.2f} Crore")
    print(f" Parallel Work Fronts Deployed:  {result.work_fronts} Simultaneous Fronts")
    print(f" Optimized Critical Path:        {' -> '.join(result.critical_path)}")

    if result.is_feasible:
        print("\n ✅ STATUS: FEASIBLE SCHEDULE FOUND WITHIN GIVEN BUDGET AND DURATION.")
        print(f"    {result.feasibility_message}")
    else:
        print("\n ❌ STATUS: No feasible schedule found within the given budget and duration.")
        print("    Constraint Violations:")
        for w in result.warnings:
            print(f"    • {w}")

    # Reality checks (benchmark / Bromilow / monsoon exposure) are reported even
    # for a feasible plan - a schedule that meets both constraints can still be
    # wildly out of line with what this kind of road actually costs and takes.
    if result.validation_warnings:
        print("\n ⚠ VALIDATION WARNINGS:")
        for w in result.validation_warnings:
            print(f"    • {w}")

    # 5. Generate Excel Report
    print(f"\n⏳ [4/5] Generating 14-sheet Excel report: {output_path}...")
    reporter = ExcelReportGenerator(result=result, output_path=output_path)
    final_file = reporter.generate()
    print(f"    ✔ Excel report generated successfully at: {os.path.abspath(final_file)}")
    print("================================================================================")
    print(" ✅ CONSTRAINT-BASED PLANNING COMPLETED SUCCESSFULLY")
    print(f" Output File: {os.path.abspath(final_file)}")
    print("================================================================================")

    return result


# ======================================================================
# UPDATE-PLAN COMMAND
# ======================================================================

def _load_activity_progress(path: str) -> dict:
    """Load activity-level progress from a JSON file: {activity_id: pct}."""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError(
            f"Activity progress file '{path}' must be a JSON object "
            f"mapping activity IDs to completion percentages.")
    return {str(k): float(v) for k, v in data.items()}


def run_update_plan(args) -> None:
    """
    Executes the update-plan pipeline:
    1. Reads the existing Excel workbook to reconstruct the original project.
    2. Re-derives the baseline by running the deterministic ConstraintScheduler.
    3. Compares actual progress with the baseline.
    4. Replans the remaining work using the existing CPM / cost / calendar engines.
    5. Checks target-date feasibility and applies acceleration if needed.
    6. Generates the updated Excel workbook.
    """
    # 1. Read the existing workbook
    print(f"  [1/6] Reading existing workbook: {args.input}...")
    reader = WorkbookReader(args.input)
    proj_input, metadata = reader.read()
    if reader.warnings:
        for w in reader.warnings:
            print(f"      - {w}")

    # Override start date if provided on CLI
    if args.start_date:
        proj_input.start_date = parse_start_date(args.start_date)

    print(f"      Project: {proj_input.title}")
    print(f"      Length:  {proj_input.length_km:.1f} km, {proj_input.lanes}-Lane, {proj_input.terrain}")
    print(f"      Start:   {proj_input.start_date}")

    # 2. Load guidelines and assumptions
    print("  [2/6] Loading MoRTH guidelines and assumptions...")
    loader = DataLoader(guidelines_path=args.guidelines, assumptions_path=args.assumptions)
    loader.load_all(terrain=proj_input.terrain)
    if loader.warnings:
        for w in loader.warnings:
            print(f"      - {w}")

    # 3. Re-derive the baseline
    print("  [3/6] Re-deriving original baseline schedule...")
    scheduler = ConstraintScheduler(
        project_input=proj_input,
        data_loader=loader,
        enable_acceleration_cost=not args.no_accel_cost,
    )
    baseline_result = scheduler.solve()
    print(f"      Baseline: {baseline_result.baseline_duration_days} d, "
          f"Rs {baseline_result.baseline_cost_crore:,.2f} Cr")
    print(f"      Optimized: {baseline_result.calculated_duration_days} d, "
          f"Rs {baseline_result.total_calculated_cost_crore:,.2f} Cr")

    # 4. Parse activity-level progress if provided
    activity_progress = None
    if args.activity_progress:
        activity_progress = _load_activity_progress(args.activity_progress)
        print(f"      Loaded activity-level progress for {len(activity_progress)} activities.")

    # 5. Build update input and run the planner
    status_date = parse_start_date(args.status_date)
    target_date = None
    if getattr(args, "target_date", None) and str(args.target_date).strip():
        target_date = parse_start_date(args.target_date)
        print(f"      Target Date (CLI):                          {target_date}")
    elif metadata.get("optimized_finish_date"):
        target_date = metadata["optimized_finish_date"]
        print(f"      Target Date (from previous Excel optimized): {target_date}")
    else:
        target_date = proj_input.start_date + timedelta(days=baseline_result.calculated_duration_days)
        print(f"      Target Date (from re-derived baseline/opt): {target_date}")

    update_input = UpdatePlanInput(
        input_path=args.input,
        start_date=proj_input.start_date,
        status_date=status_date,
        actual_progress_pct=float(args.actual_progress),
        target_date=target_date,
        target_progress_pct=float(args.target_progress) if args.target_progress else None,
        output_path=args.output,
        activity_progress=activity_progress,
        guidelines_path=args.guidelines,
        assumptions_path=args.assumptions,
        enable_acceleration_cost=not args.no_accel_cost,
    )

    print("  [4/6] Comparing progress and replanning remaining work...")
    planner = UpdatePlanner(
        baseline_result=baseline_result,
        update_input=update_input,
        data_loader=loader,
    )
    result = planner.plan()

    # Print update summary
    print("================================================================================")
    print("  ROAD CONSTRUCTION PROJECT PLAN UPDATE")
    print("================================================================================")
    print(f"  Project:                    {proj_input.title}")
    print(f"  Status Date:                {status_date}")
    print(f"  Target Date:                {target_date}")
    print(f"  Actual Progress:            {result.actual_progress_pct:.1f}%")
    print(f"  Baseline Progress at Status:{result.baseline_progress_at_status_date:.1f}%")
    print(f"  Progress Gap (Behind):     {result.progress_gap_pct:.1f} pp")
    print(f"  Target Progress:            {result.target_progress_pct:.1f}%")
    print(f"  Required Progress (Remain): {result.required_progress_pct:.1f} pp")
    print(f"  Revised Duration:          {result.revised_duration_days} days (from status date)")
    print(f"  Revised Completion Date:   {result.revised_completion_date}")
    print(f"  Target Feasible:            {'YES' if result.target_feasible else 'NO'}")
    if result.projected_delay_days > 0:
        print(f"  Projected Delay:            {result.projected_delay_days} days")
    print(f"  Spent Cost:                 Rs {result.spent_cost_crore:,.2f} Cr")
    print(f"  Remaining Cost:             Rs {result.remaining_cost_crore:,.2f} Cr")
    if result.additional_acceleration_cost_crore > 0:
        print(f"  Acceleration Cost:          Rs {result.additional_acceleration_cost_crore:,.2f} Cr")
    print(f"  Total Revised Cost:         Rs {result.total_project_cost_crore:,.2f} Cr")
    print(f"  Baseline Total Cost:        Rs {result.baseline_total_cost_crore:,.2f} Cr")
    print(f"  Revised Critical Path:     {' -> '.join(result.revised_critical_path)}")

    if result.limitations:
        print("\n  LIMITATIONS:")
        for lim in result.limitations:
            print(f"    - {lim}")
    if result.warnings:
        print("\n  WARNINGS:")
        for w in result.warnings:
            print(f"    - {w}")

    print(f"\n  {result.feasibility_message}")

    # 6. Generate the updated Excel report
    print(f"\n  [5/6] Generating updated Excel report: {args.output}...")
    reporter = UpdateExcelReportGenerator(result=result, output_path=args.output)
    final_file = reporter.generate()

    # Validate the revised network
    print("  [6/6] Validating revised dependency network...")
    errors = DependencyManager.validate_network(result.revised_activities)
    if errors:
        for err in errors:
            print(f"      - Network Error: {err}")
    else:
        print("      - Dependency network is cycle-free and fully connected.")

    print("================================================================================")
    print("  PROJECT PLAN UPDATE COMPLETED SUCCESSFULLY")
    print(f"  Output File: {os.path.abspath(final_file)}")
    print("================================================================================")


def _build_update_plan_parser() -> argparse.ArgumentParser:
    """Build the argument parser for the update-plan subcommand."""
    p = argparse.ArgumentParser(
        prog="python main.py update-plan",
        description="Update an existing project plan based on actual progress.",
    )
    p.add_argument("--input", type=str, required=True,
                   help="Path to the existing Excel project plan workbook")
    p.add_argument("--start-date", type=str, default="",
                   help="Original project start date (YYYY-MM-DD). "
                        "If omitted, the date from the workbook is used.")
    p.add_argument("--status-date", type=str, required=True,
                   help="Date on which actual progress is reported (YYYY-MM-DD)")
    p.add_argument("--actual-progress", type=float, required=True,
                   help="Actual cumulative project progress as a percentage (0-100)")
    p.add_argument("--target-date", type=str, default="",
                   help="Future date by which the target should be achieved (YYYY-MM-DD). "
                        "Optional: if omitted, the optimized final date from the previous Excel workbook is used.")
    p.add_argument("--target-progress", type=float, default=None,
                   help="Optional cumulative progress target at the target date. "
                        "If omitted, the baseline target for that date is used.")
    p.add_argument("--output", type=str, default="road_project_schedule_updated.xlsx",
                   help="Filename for the updated Excel workbook")
    p.add_argument("--activity-progress", type=str, default=None,
                   help="Optional JSON file with activity-level progress: "
                        '{"ACT-01": 100.0, "ACT-02": 50.0, ...}')
    p.add_argument("--guidelines", type=str, default="road_guidelines.json",
                   help="Path to road_guidelines.json")
    p.add_argument("--assumptions", type=str, default="road_assumptions.yaml",
                   help="Path to road_assumptions.yaml")
    p.add_argument("--no-accel-cost", dest="no_accel_cost", action="store_true",
                   help="Disable time-driven acceleration cost terms")
    return p


def main():
    # Dispatch to update-plan subcommand if the first argument is "update-plan"
    if len(sys.argv) > 1 and sys.argv[1] == "update-plan":
        parser = _build_update_plan_parser()
        args = parser.parse_args(sys.argv[2:])
        run_update_plan(args)
        return

    parser = argparse.ArgumentParser(description="Road Construction Project Planner & Constraint-Based Scheduler")
    parser.add_argument("--title", type=str, default="", help="Project Title (e.g. '10 KM 2-Lane Highway Project')")
    parser.add_argument("--length", type=float, default=0.0, help="Project Length in km")
    parser.add_argument("--duration", type=str, default="", help="Required Project Duration (e.g. 365 days or 12 months)")
    parser.add_argument("--budget", type=str, default="", help="Maximum Project Budget in Crore (e.g. 300)")
    parser.add_argument("--start-date", type=str, default="", help="Project Start Date (YYYY-MM-DD)")
    parser.add_argument("--terrain", type=str, choices=("plain", "hilly"), default="",
                        help="Terrain model; defaults to plain unless the title says hilly, hill, or mountain")
    parser.add_argument("--guidelines", type=str, default="road_guidelines.json", help="Path to road_guidelines.json")
    parser.add_argument("--assumptions", type=str, default="road_assumptions.yaml", help="Path to road_assumptions.yaml")
    parser.add_argument("--output", type=str, default="road_project_schedule.xlsx", help="Output Excel filename")
    parser.add_argument("--lanes", type=int, default=0, help="Number of Highway Lanes (auto-detected from title if omitted)")
    parser.add_argument(
        "--no-accel-cost",
        dest="no_accel_cost",
        action="store_true",
        help="Disable all NEW time-driven acceleration cost terms (extra-front site establishment, "
             "resident staff, plant hire minimum uplift, gang interference). Reproduces the legacy "
             "cost behaviour so both cost models can be compared on identical input."
    )
    args = parser.parse_args()

    # Interactive input if CLI arguments are not fully provided
    title = args.title.strip()
    if not title:
        try:
            title = input("Project Title (e.g. 25 KM 4-Lane Highway Project or 10 KM 2-Lane Road): ").strip()
        except EOFError:
            title = "25 KM 4-Lane Highway Project"
        if not title:
            title = "25 KM 4-Lane Highway Project"

    terrain = args.terrain.strip().lower() if args.terrain else parse_terrain_from_title(title)
    print(f"    ✔ Selected terrain model: {terrain}")

    # Automatically extract lane count from project title
    title_lanes = parse_lanes_from_title(title, default_lanes=0)
    if title_lanes > 0:
        lanes = title_lanes
        print(f"    ✔ Detected {lanes}-Lane Highway configuration from project title: '{title}'")
    elif args.lanes > 0:
        lanes = args.lanes
    else:
        lanes = 4

    # Automatically extract length from title if available and not specified on CLI
    title_length = parse_length_from_title(title, default_length=0.0)
    length_km = args.length
    if length_km <= 0:
        if title_length > 0:
            length_km = title_length
            print(f"    ✔ Detected {length_km:.1f} km corridor length from project title: '{title}'")
        else:
            try:
                raw_len = input("Project Length in km (e.g. 25): ").strip()
                length_km = float(raw_len) if raw_len else 25.0
            except (EOFError, ValueError):
                length_km = 25.0

    duration_str = args.duration.strip()
    if not duration_str:
        try:
            duration_str = input("Required Project Duration (e.g. 365 days or 12 months): ").strip()
        except EOFError:
            duration_str = "365 days"
        if not duration_str:
            duration_str = "365 days"
    duration_days = parse_duration_input(duration_str)

    budget_str = args.budget.strip()
    if not budget_str:
        try:
            budget_str = input("Maximum Project Budget in Crore (e.g. 1000 or 880): ").strip()
        except EOFError:
            budget_str = "1000"
        if not budget_str:
            budget_str = "1000"
    budget_crore = parse_budget_input(budget_str)

    start_date_str = args.start_date.strip()
    if not start_date_str:
        start_date = date.today()
    else:
        start_date = parse_start_date(start_date_str)

    # Validation
    if length_km <= 0:
        print("Error: Project length must be greater than 0.")
        sys.exit(1)
    if duration_days <= 0:
        print("Error: Project duration must be greater than 0.")
        sys.exit(1)
    if budget_crore <= 0:
        print("Error: Project budget must be greater than 0.")
        sys.exit(1)

    proj_input = ProjectInput(
        title=title,
        length_km=length_km,
        duration_days=duration_days,
        budget_crore=budget_crore,
        start_date=start_date,
        lanes=lanes,
        terrain=terrain
    )

    run_project_planning(
        project_input=proj_input,
        guidelines_path=args.guidelines,
        assumptions_path=args.assumptions,
        output_path=args.output,
        enable_acceleration_cost=not args.no_accel_cost
    )

if __name__ == "__main__":
    main()
