"""
Main Entry Point for Construction Project Planning & Constraint-Based Scheduling.
Supports road, pipeline, and building project types.
"""
import os
import sys
import re
import json
import argparse
from datetime import date, datetime, timedelta

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
from src.workbook_metadata import freshness_warnings
from src.validation_sanity_checks import (
    format_validation_report,
    validate_schedule_result,
    validate_exported_workbook,
    validate_update_result,
    validate_exported_update_workbook,
)

_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
_OUTPUT_DIR = os.path.join(_PROJECT_ROOT, "output")


def _output_file_path(path: str, default_name: str) -> str:
    """Keep CLI-generated workbooks in the project output directory."""
    requested = str(path or "").strip() or default_name
    if os.path.isabs(requested):
        return requested
    normalized = os.path.normpath(requested)
    if normalized == "output" or normalized.startswith(f"output{os.sep}"):
        return normalized
    return os.path.join("output", normalized)


def _input_file_path(path: str) -> str:
    """Resolve a workbook name from the output directory when appropriate."""
    requested = str(path or "").strip()
    if os.path.exists(requested) or os.path.isabs(requested):
        return requested
    candidate = os.path.join("output", requested)
    return candidate if os.path.exists(candidate) else requested


# Default config paths per project type
_CONFIG = {
    "road": {
        "guidelines": "project_config/road/road_guidelines.json",
        "assumptions": "project_config/road/road_assumptions.yaml",
        "output": "output/road_project_schedule.xlsx",
        "update_output": "output/road_project_schedule_updated.xlsx",
    },
    "pipeline": {
        "guidelines": "project_config/pipeline/pipeline_guidelines.json",
        "assumptions": "project_config/pipeline/pipeline_assumptions.yaml",
        "output": "output/pipeline_project_schedule.xlsx",
        "update_output": "output/pipeline_project_schedule_updated.xlsx",
    },
    "building": {
        "guidelines": "project_config/building/building_guidelines.json",
        "assumptions": "project_config/building/building_assumptions.yaml",
        "output": "output/building_project_schedule.xlsx",
        "update_output": "output/building_project_schedule_updated.xlsx",
    },
}


def parse_duration_input(raw_input: str) -> int:
    """Parses duration string like '365', '365 days', '12 months' into days."""
    s = str(raw_input).strip().lower().replace(",", "")
    match_months = re.search(r"(\d+(?:\.\d+)?)\s*(months?|mo|m)\b", s)
    if match_months:
        months_val = float(match_months.group(1))
        return int(round(months_val * 30))
    match_days = re.search(r"(\d+(?:\.\d+)?)", s)
    if match_days:
        return int(round(float(match_days.group(1))))
    raise ValueError(f"Unable to parse duration from: '{raw_input}'")


def parse_budget_input(raw_input: str) -> float:
    """Parses a budget string into Crore.

    Accepts plain numbers ('1000'), currency formatting ('₹ 1,000 Cr') and
    lakh/lac units ('150 lakh' -> 1.5 Cr). Crore is the model's native unit.
    """
    s = str(raw_input).strip().lower().replace(",", "")
    match = re.search(r"(\d+(?:\.\d+)?)", s)
    if not match:
        raise ValueError(f"Unable to parse budget from: '{raw_input}'")
    value = float(match.group(1))
    if re.search(r"\b(lakh|lakhs|lac|lacs)\b", s):
        return value / 100.0
    return value


def parse_start_date(raw_input: str) -> date:
    """Parses YYYY-MM-DD or defaults to today for an empty value."""
    s = str(raw_input).strip()
    if not s:
        return date.today()
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(
            f"Invalid date '{s}'. Use YYYY-MM-DD (for example 2026-11-03)."
        ) from exc


def _project_type_banner(project_type: str) -> str:
    return {
        "road": "🛣️  ROAD CONSTRUCTION CONSTRAINT-BASED SCHEDULING PIPELINE",
        "pipeline": "🚰 WATER PIPELINE CONSTRAINT-BASED SCHEDULING PIPELINE",
        "building": "🏗️  BUILDING CONSTRUCTION CONSTRAINT-BASED SCHEDULING PIPELINE",
    }.get(project_type, "⚙️  CONSTRUCTION PROJECT SCHEDULING PIPELINE")


def run_project_planning(
    project_input: ProjectInput,
    guidelines_path: str,
    assumptions_path: str,
    output_path: str,
    enable_acceleration_cost: bool = True
) -> ProjectScheduleResult:
    """
    Executes the constraint-based planning pipeline for any project type.
    """
    ptype = project_input.project_type

    print(f"⏳ [1/5] Loading guidelines and assumptions for {ptype.upper()} project...")
    loader = DataLoader(guidelines_path=guidelines_path, assumptions_path=assumptions_path)
    loader.load_all(terrain=project_input.terrain)
    print("    ✔ Guidelines & assumptions loaded successfully.")
    if loader.errors:
        for error in loader.errors:
            print(f"      ✖ {error}")
        raise ValueError(
            "Required engineering assumptions are missing or invalid; "
            "workbook generation stopped."
        )
    if loader.warnings:
        print("    ⚠ DATA-INTEGRITY WARNINGS:")
        for w in loader.warnings:
            print(f"      • {w}")

    if not enable_acceleration_cost:
        print("    ⚠ --no-accel-cost: time-driven acceleration surcharges are DISABLED.")

    print(f"⏳ [2/5] Running Constraint-Based Optimizer & CPM Calculation...")
    scheduler = ConstraintScheduler(
        project_input=project_input,
        data_loader=loader,
        enable_acceleration_cost=enable_acceleration_cost
    )
    result = scheduler.solve()
    result.guidelines_path = guidelines_path
    result.assumptions_path = assumptions_path

    validation = validate_schedule_result(
        result,
        loader.assumptions,
        calendar=scheduler.calendar,
    )
    for warning in result.validation_warnings:
        validation.add("WARNING", "SCHEDULER_SANITY", warning)
    # Keep the structured report on the result so the Excel exporter can show
    # the exact checks that were run, including warnings.
    result.validation_report = validation
    if validation.errors:
        print(format_validation_report(validation))
        raise ValueError("Schedule validation failed; workbook generation stopped.")
    for issue in validation.warnings:
        warning = f"{issue.code}: {issue.message}"
        if warning not in result.warnings:
            result.warnings.append(warning)

    print("=" * 80)
    print(f" {_project_type_banner(ptype)}")
    print("=" * 80)
    print(f"Project Title:              {project_input.title}")
    if ptype == "road":
        print(f"Terrain:                    {project_input.terrain.title()}")
        print(f"Highway Configuration:      {project_input.lanes}-Lane Highway ({project_input.carriageway_width_m:.1f}m Carriageway)")
        print(f"Project Length:             {project_input.length_km:.2f} km")
    elif ptype == "pipeline":
        print(f"Pipeline Type:              {getattr(project_input, 'pipeline_type', 'water').title()}")
        print(f"Pipe Specification:         DN{int(project_input.pipe_diameter_mm)} {project_input.pipe_material}")
        print(f"Total Length:               {project_input.length_km:.2f} km")
        print(f"Installation Method:        {project_input.installation_method.replace('_', ' ').title()}")
        print(f"Trench:                     {project_input.trench_depth_m:.1f}m deep × {project_input.trench_width_m:.2f}m wide")
        print(f"Work Fronts:                {project_input.number_of_work_fronts}")
    elif ptype == "building":
        print(f"Building Type:              {getattr(project_input, 'building_type', 'residential').title()}")
        print(f"Built-up Area:              {getattr(project_input, 'built_up_area_sqm', 0):,.0f} sqm  ({getattr(project_input, 'number_of_floors', 1)} floor(s))")
        print(f"Construction Type:          {getattr(project_input, 'construction_type', 'RCC').upper()}")
        print(f"Quality Finish:             {getattr(project_input, 'quality_finish', 'standard').title()}")
        print(f"Basement:                   {'Yes' if getattr(project_input, 'basement', False) else 'No'}")
    print(f"Required Target Duration:   {project_input.duration_days} days (Hard Maximum)")
    print(f"Approved Maximum Budget:    ₹ {project_input.budget_crore:,.2f} Crore (Hard Ceiling)")
    print(f"Project Start Date:         {project_input.start_date}")
    print(f"Calendar rules in force:")
    for _line in scheduler.calendar.describe():
        print(f"   • {_line}")
    print("-" * 80)

    print(f"⏳ [3/5] Validating optimized activity precedence network...")
    errors = DependencyManager.validate_network(result.activities)
    if errors:
        for err in errors:
            print(f"    ❌ Network Error: {err}")
        raise ValueError("Network dependency validation failed.")
    print("    ✔ Dependency network is cycle-free and fully connected.")

    print("\n" + "-" * 80)
    print(" 📊 CONSTRAINT-BASED SCHEDULING RESULTS")
    print("-" * 80)
    print(f" Baseline Duration (Normal):     {result.baseline_duration_days} Days")
    print(f" Baseline Cost:                  ₹ {result.baseline_cost_crore:,.2f} Crore")
    print(f" Optimized Duration:             {result.calculated_duration_days} Days (Target: <= {project_input.duration_days} Days)")
    print(f" Optimized Total Cost:           ₹ {result.total_calculated_cost_crore:,.2f} Crore (Budget: <= ₹ {project_input.budget_crore:,.2f} Crore)")
    print(f" Duration Saved:                 {result.duration_saved_days} Days")
    print(f" Additional Optimization Cost:   ₹ {result.additional_cost_crore:,.2f} Crore")
    print(f" Parallel Work Fronts Deployed:  {result.work_fronts}")
    print(f" Optimized Critical Path:        {' -> '.join(result.critical_path)}")

    if result.is_feasible:
        print("\n ✅ STATUS: FEASIBLE SCHEDULE FOUND WITHIN GIVEN BUDGET AND DURATION.")
        print(f"    {result.feasibility_message}")
    else:
        print("\n ❌ STATUS: No feasible schedule found within the given budget and duration.")
        for w in result.warnings:
            print(f"    • {w}")

    if result.validation_warnings:
        print("\n ⚠ VALIDATION WARNINGS:")
        for w in result.validation_warnings:
            print(f"    • {w}")

    print(f"\n⏳ [4/5] Generating Excel report: {output_path}...")
    reporter = ExcelReportGenerator(result=result, output_path=output_path)
    final_file = reporter.generate()
    workbook_validation = validate_exported_workbook(final_file, result)
    if workbook_validation.errors:
        print(format_validation_report(workbook_validation))
        raise ValueError("Exported workbook validation failed.")
    print(f"    ✔ Excel report generated successfully at: {os.path.abspath(final_file)}")

    # ── [5/5] Entity export: SQLite + JSON ───────────────────────────────────
    print(f"⏳ [5/5] Exporting entities...")
    out_dir = os.path.dirname(os.path.abspath(final_file)) or "."
    try:
        from scripts.actual_to_entity import build_entities as _build_entities
        _project_slug = re.sub(r"[^a-z0-9]+", "_",
                               project_input.title.lower()).strip("_")[:40]
        _actual_json = os.path.join(out_dir, f"{_project_slug}_actual_data.json")
        entities = _build_entities(
            xlsx_path=final_file,
            project_type=project_input.project_type,
            seed=42,
        )
        with open(_actual_json, "w", encoding="utf-8") as _f:
            import json as _json
            _json.dump(entities, _f, indent=2, ensure_ascii=False, default=str)
        total = sum(len(v) for v in entities.values())
        print(f"    ✔ {total:,} entity records written to: {os.path.abspath(_actual_json)}")
        print(f"    → Run: node push-entities.mjs --json={os.path.abspath(_actual_json)}")
    except Exception as _exc:
        print(f"    ⚠ Entity export failed (non-fatal): {_exc}")

    print("=" * 80)
    print(" ✅ CONSTRAINT-BASED PLANNING COMPLETED SUCCESSFULLY")
    print(f" Output File: {os.path.abspath(final_file)}")
    print("=" * 80)

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
    input_path = _input_file_path(args.input)
    print(f"  [1/6] Reading existing workbook: {input_path}...")
    reader = WorkbookReader(input_path)
    proj_input, metadata = reader.read()
    if reader.warnings:
        for w in reader.warnings:
            print(f"      - {w}")

    # Override start date if provided on CLI
    if args.start_date:
        proj_input.start_date = parse_start_date(args.start_date)

    ptype = str(getattr(proj_input, "project_type", "road") or "road").lower()
    print(f"      Project: {proj_input.title}")
    if ptype == "building":
        print(f"      Built-up area: {proj_input.built_up_area_sqm:,.1f} sqm, "
              f"G+{max(0, proj_input.number_of_floors - 1)}, "
              f"{proj_input.number_of_blocks} block(s)")
    elif ptype == "pipeline":
        print(f"      Length:  {proj_input.length_km:.1f} km, DN"
              f"{int(proj_input.pipe_diameter_mm)} {proj_input.pipe_material}, "
              f"{proj_input.number_of_work_fronts} front(s)")
    else:
        print(f"      Length:  {proj_input.length_km:.1f} km, "
              f"{proj_input.lanes}-Lane, {proj_input.terrain}")
    print(f"      Start:   {proj_input.start_date}")

    # 2. Load guidelines and assumptions
    print("  [2/6] Loading guidelines and assumptions...")
    guidelines_path, assumptions_path = _auto_config(
        ptype, getattr(args, "guidelines", "") or "", getattr(args, "assumptions", "") or "")
    loader = DataLoader(guidelines_path=guidelines_path, assumptions_path=assumptions_path)
    loader.load_all(terrain=proj_input.terrain)
    if loader.errors:
        for error in loader.errors:
            print(f"      - CONFIGURATION ERROR: {error}")
        raise ValueError(
            "Required engineering assumptions are missing or invalid; "
            "update-plan stopped."
        )
    freshness = freshness_warnings(
        metadata.get("workbook_provenance", {}),
        proj_input,
        guidelines_path,
        assumptions_path,
    )
    for warning in freshness:
        reader.warnings.append("FRESHNESS: " + warning)
    for warning in freshness:
        print(f"      - FRESHNESS: {warning}")
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
    baseline_result.guidelines_path = guidelines_path
    baseline_result.assumptions_path = assumptions_path
    print(f"      Baseline: {baseline_result.baseline_duration_days} d, "
          f"Rs {baseline_result.baseline_cost_crore:,.2f} Cr")
    print(f"      Optimized: {baseline_result.calculated_duration_days} d, "
          f"Rs {baseline_result.total_calculated_cost_crore:,.2f} Cr")

    # Cross-check the re-derived baseline against the figures printed in the
    # source workbook. A mismatch means the assumptions/config have changed
    # since that workbook was generated, so the update is not a like-for-like
    # replan of the original baseline.
    source_days = metadata.get("baseline_duration_days")
    source_cost = metadata.get("baseline_cost_crore")
    if source_days is not None and abs(
            baseline_result.baseline_duration_days - source_days) > 1:
        print(f"      WARNING: source workbook baseline is {source_days} d, "
              f"but re-derived baseline is {baseline_result.baseline_duration_days} d.")
    if source_cost is not None and abs(
            baseline_result.baseline_cost_crore - source_cost) > 0.05:
        print(f"      WARNING: source workbook baseline cost is "
              f"Rs {source_cost:,.2f} Cr, but re-derived baseline cost is "
              f"Rs {baseline_result.baseline_cost_crore:,.2f} Cr.")

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
        input_path=input_path,
        start_date=proj_input.start_date,
        status_date=status_date,
        actual_progress_pct=float(args.actual_progress),
        target_date=target_date,
        target_progress_pct=float(args.target_progress) if args.target_progress else None,
        output_path=_output_file_path(
            args.output,
            _CONFIG.get(ptype, _CONFIG["road"])["update_output"],
        ),
        activity_progress=activity_progress,
        guidelines_path=guidelines_path,
        assumptions_path=assumptions_path,
        enable_acceleration_cost=not args.no_accel_cost,
    )

    print("  [4/6] Comparing progress and replanning remaining work...")
    planner = UpdatePlanner(
        baseline_result=baseline_result,
        update_input=update_input,
        data_loader=loader,
    )
    result = planner.plan()
    update_validation = validate_update_result(result, loader.assumptions)
    result.validation_report = update_validation
    if update_validation.errors:
        print(format_validation_report(update_validation))
        raise ValueError("Updated schedule validation failed; workbook generation stopped.")

    # Print update summary
    print("================================================================================")
    print(f"  {ptype.upper()} PROJECT PLAN UPDATE")
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
    print(f"\n  [5/6] Generating updated Excel report: {update_input.output_path}...")
    reporter = UpdateExcelReportGenerator(
        result=result,
        output_path=update_input.output_path,
    )
    final_file = reporter.generate()
    workbook_validation = validate_exported_update_workbook(final_file, result)
    if workbook_validation.errors:
        print(format_validation_report(workbook_validation))
        raise ValueError("Exported updated workbook validation failed.")

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
                   help="Original project start date (YYYY-MM-DD).")
    p.add_argument("--status-date", type=str, required=True,
                   help="Date on which actual progress is reported (YYYY-MM-DD)")
    p.add_argument("--actual-progress", type=float, required=True,
                   help="Actual cumulative project progress as a percentage (0-100)")
    p.add_argument("--target-date", type=str, default="",
                   help="Future date by which the target should be achieved (YYYY-MM-DD).")
    p.add_argument("--target-progress", type=float, default=None)
    p.add_argument("--output", type=str, default="")
    p.add_argument("--activity-progress", type=str, default=None)
    p.add_argument("--guidelines", type=str, default="",
                   help="Path to guidelines JSON (auto-detected from workbook project_type if omitted)")
    p.add_argument("--assumptions", type=str, default="",
                   help="Path to assumptions YAML (auto-detected from workbook project_type if omitted)")
    p.add_argument("--no-accel-cost", dest="no_accel_cost", action="store_true")
    return p


def _auto_config(project_type: str, guidelines: str, assumptions: str):
    """Return (guidelines, assumptions) falling back to defaults for the project type."""
    cfg = _CONFIG.get(str(project_type or "road").lower(), _CONFIG["road"])
    g = guidelines if guidelines else cfg["guidelines"]
    a = assumptions if assumptions else cfg["assumptions"]
    return g, a


def _build_road_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python main.py road",
        description="Plan a road construction project.",
    )
    p.add_argument("--title", type=str, default="")
    p.add_argument("--length", type=float, default=0.0)
    p.add_argument("--duration", type=str, default="")
    p.add_argument("--budget", type=str, default="")
    p.add_argument("--start-date", type=str, default="")
    p.add_argument("--terrain", type=str, choices=("plain", "hilly"), default="")
    p.add_argument("--lanes", type=int, default=0)
    p.add_argument("--service-road-sides", type=int, default=0,
                   help="Optional service-road sides; 0 means out of scope")
    p.add_argument("--toll-plazas", type=int, default=0)
    p.add_argument("--atms", type=int, default=0)
    p.add_argument("--bridges", type=int, default=0)
    p.add_argument("--flyovers", type=int, default=0)
    p.add_argument("--underpasses", type=int, default=0)
    p.add_argument("--retaining-walls-sqm", type=float, default=0.0)
    p.add_argument("--guidelines", type=str, default="")
    p.add_argument("--assumptions", type=str, default="")
    p.add_argument("--output", type=str, default="")
    p.add_argument("--no-accel-cost", dest="no_accel_cost", action="store_true")
    return p


def _build_pipeline_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python main.py pipeline",
        description="Plan a water pipeline construction project.",
    )
    p.add_argument("--title", type=str, default="")
    p.add_argument("--length", type=float, default=0.0,
                   help="Total pipeline length in km")
    p.add_argument("--diameter", type=float, default=300.0,
                   help="Nominal pipe diameter in mm (e.g. 300, 400, 600)")
    p.add_argument("--material", type=str, default="DI",
                   choices=("DI", "HDPE", "MS", "RCC"),
                   help="Pipe material: DI / HDPE / MS / RCC")
    p.add_argument("--trench-depth", type=float, default=2.0,
                   help="Average trench depth in metres")
    p.add_argument("--trench-width", type=float, default=0.0,
                   help="Excavated trench width in metres (0 = derive from pipe diameter)")
    p.add_argument("--installation-method", type=str, default="open_trench",
                   choices=("open_trench", "trenchless"),
                   help="Open-trench or trenchless (HDD) installation")
    p.add_argument("--fronts", type=int, default=1,
                   help="Number of simultaneous work fronts")
    p.add_argument("--pipeline-type", type=str, default="water",
                   choices=("water", "sewerage", "gas"))
    p.add_argument("--villages", type=int, default=0,
                   help="Number of villages served (derives OHT/FHTC/WTP defaults)")
    p.add_argument("--ohts", type=int, default=0,
                   help="Number of overhead service reservoirs (OHT/ESR)")
    p.add_argument("--wtp-mld", type=float, default=0.0,
                   help="Water treatment plant capacity in MLD (0 = not in scope)")
    p.add_argument("--ibps", type=int, default=0,
                   help="Number of intake / booster pumping stations")
    p.add_argument("--fhtc", type=int, default=0,
                   help="Number of functional household tap connections")
    p.add_argument("--road-crossings", type=int, default=0,
                   help="Declared open-cut road crossings (0 = derive from length)")
    p.add_argument("--highway-crossings", type=int, default=0,
                   help="Declared NH/SH crossings")
    p.add_argument("--rail-crossings", type=int, default=0,
                   help="Declared rail crossings")
    p.add_argument("--hdd-crossings", type=int, default=0,
                   help="Declared HDD/trenchless crossings")
    p.add_argument("--river-crossings", type=int, default=0,
                   help="Declared river crossings")
    p.add_argument("--duration", type=str, default="")
    p.add_argument("--budget", type=str, default="")
    p.add_argument("--start-date", type=str, default="")
    p.add_argument("--terrain", type=str, choices=("plain", "hilly"), default="plain")
    p.add_argument("--guidelines", type=str, default="")
    p.add_argument("--assumptions", type=str, default="")
    p.add_argument("--output", type=str, default="")
    p.add_argument("--no-accel-cost", dest="no_accel_cost", action="store_true")
    return p


def _build_building_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python main.py building",
        description="Plan a residential/commercial building construction project.",
    )
    p.add_argument("--title", type=str, default="")
    p.add_argument("--area", type=float, default=0.0,
                   help="Total built-up area in sqm (or sqft with --sqft flag)")
    p.add_argument("--sqft", action="store_true",
                   help="Interpret --area as square feet instead of sqm")
    p.add_argument("--floors", type=int, default=4,
                   help="Number of floors including ground (e.g. G+3 = 4)")
    p.add_argument("--type", dest="building_type", type=str, default="residential",
                   choices=("residential", "commercial"))
    p.add_argument("--construction", dest="construction_type", type=str, default="rcc",
                   choices=("rcc", "load_bearing", "steel"))
    p.add_argument("--quality", type=str, default="standard",
                   choices=("basic", "standard", "premium"))
    p.add_argument("--basement", action="store_true",
                   help="Include basement in scope")
    p.add_argument("--blocks", type=int, default=0,
                   help="Number of blocks/towers (0 = auto-derive from area)")
    p.add_argument("--foundation-type", type=str, default="pile",
                   choices=("pile", "raft", "isolated_footing"))
    p.add_argument("--floor-wise-structural", action="store_true",
                   help="Create explicit structural activities for every floor")
    p.add_argument("--duration", type=str, default="")
    p.add_argument("--budget", type=str, default="")
    p.add_argument("--start-date", type=str, default="")
    p.add_argument("--guidelines", type=str, default="")
    p.add_argument("--assumptions", type=str, default="")
    p.add_argument("--output", type=str, default="")
    p.add_argument("--no-accel-cost", dest="no_accel_cost", action="store_true")
    return p


def run_road(args) -> None:
    title = args.title.strip()
    if not title:
        try:
            title = input("Project Title (e.g. 25 KM 4-Lane Highway): ").strip()
        except EOFError:
            title = "25 KM 4-Lane Highway Project"
        if not title:
            title = "25 KM 4-Lane Highway Project"

    terrain = args.terrain.strip().lower() if args.terrain else parse_terrain_from_title(title)
    title_lanes = parse_lanes_from_title(title, default_lanes=0)
    lanes = title_lanes if title_lanes > 0 else (args.lanes if args.lanes > 0 else 4)
    title_length = parse_length_from_title(title, default_length=0.0)
    length_km = args.length if args.length > 0 else (title_length if title_length > 0 else 0.0)
    if length_km <= 0:
        try:
            raw_len = input("Project Length in km: ").strip()
            length_km = float(raw_len) if raw_len else 25.0
        except (EOFError, ValueError):
            length_km = 25.0

    duration_str = args.duration.strip() or _prompt("Required Duration (e.g. 365 days or 12 months): ", "365 days")
    budget_str = args.budget.strip() or _prompt("Maximum Budget in Crore: ", "1000")
    duration_days = parse_duration_input(duration_str)
    budget_crore = parse_budget_input(budget_str)
    start_date = parse_start_date(args.start_date.strip()) if args.start_date.strip() else date.today()

    proj_input = ProjectInput(
        title=title,
        length_km=length_km,
        duration_days=duration_days,
        budget_crore=budget_crore,
        start_date=start_date,
        lanes=lanes,
        terrain=terrain,
        project_type="road",
        service_road_sides=args.service_road_sides,
        toll_plazas=args.toll_plazas,
        atms=args.atms,
        bridges=args.bridges,
        flyovers=args.flyovers,
        underpasses=args.underpasses,
        retaining_walls_sqm=args.retaining_walls_sqm,
    )
    guidelines, assumptions = _auto_config("road", args.guidelines, args.assumptions)
    output = _output_file_path(args.output, _CONFIG["road"]["output"])
    run_project_planning(proj_input, guidelines, assumptions, output, not args.no_accel_cost)


def run_pipeline(args) -> None:
    title = args.title.strip()
    if not title:
        try:
            title = input("Pipeline Project Title (e.g. Hyderabad Water Pipeline 10 km): ").strip()
        except EOFError:
            title = "Water Pipeline Project"
        if not title:
            title = "Water Pipeline Project"

    length_km = args.length if args.length > 0 else 0.0
    if length_km <= 0:
        tl = parse_length_from_title(title, 0.0)
        if tl > 0:
            length_km = tl
        else:
            try:
                raw = input("Pipeline length in km: ").strip()
                length_km = float(raw) if raw else 10.0
            except (EOFError, ValueError):
                length_km = 10.0

    duration_str = args.duration.strip() or _prompt("Required Duration (e.g. 8 months): ", "8 months")
    budget_str = args.budget.strip() or _prompt("Maximum Budget in Crore: ", "5")
    duration_days = parse_duration_input(duration_str)
    budget_crore = parse_budget_input(budget_str)
    start_date = parse_start_date(args.start_date.strip()) if args.start_date.strip() else date.today()

    proj_input = ProjectInput(
        title=title,
        length_km=length_km,
        duration_days=duration_days,
        budget_crore=budget_crore,
        start_date=start_date,
        project_type="pipeline",
        pipe_diameter_mm=args.diameter,
        pipe_material=args.material.upper(),
        trench_depth_m=args.trench_depth,
        trench_width_m=args.trench_width,
        installation_method=args.installation_method,
        number_of_work_fronts=args.fronts,
        pipeline_type=args.pipeline_type,
        terrain=args.terrain or "plain",
        num_villages=args.villages,
        num_oht=args.ohts,
        wtp_capacity_mld=args.wtp_mld,
        num_ibps=args.ibps,
        num_fhtc=args.fhtc,
        number_of_road_crossings=args.road_crossings,
        number_of_highway_crossings=args.highway_crossings,
        number_of_rail_crossings=args.rail_crossings,
        number_of_hdd_crossings=args.hdd_crossings,
        number_of_river_crossings=args.river_crossings,
    )
    guidelines, assumptions = _auto_config("pipeline", args.guidelines, args.assumptions)
    output = _output_file_path(args.output, _CONFIG["pipeline"]["output"])
    run_project_planning(proj_input, guidelines, assumptions, output, not args.no_accel_cost)


def run_building(args) -> None:
    title = args.title.strip()
    if not title:
        try:
            title = input("Building Project Title (e.g. ABC Residential Building): ").strip()
        except EOFError:
            title = "Residential Building Project"
        if not title:
            title = "Residential Building Project"

    area_sqm = args.area
    if args.sqft and area_sqm > 0:
        area_sqm = round(area_sqm / 10.764, 1)
    if area_sqm <= 0:
        try:
            raw = input("Built-up area in sqm (or sqft if --sqft): ").strip()
            area_sqm = float(raw) if raw else 1000.0
            if args.sqft:
                area_sqm = round(area_sqm / 10.764, 1)
        except (EOFError, ValueError):
            area_sqm = 1000.0

    floors = args.floors if args.floors > 0 else 4
    blocks = args.blocks if args.blocks > 0 else 0  # 0 → auto-derive

    duration_str = args.duration.strip() or _prompt("Required Duration (e.g. 12 months): ", "12 months")
    budget_str = args.budget.strip() or _prompt("Maximum Budget in Crore: ", "2")
    duration_days = parse_duration_input(duration_str)
    budget_crore = parse_budget_input(budget_str)
    start_date = parse_start_date(args.start_date.strip()) if args.start_date.strip() else date.today()

    proj_input = ProjectInput(
        title=title,
        length_km=area_sqm,  # length_km carries area for routing; ProjectInput normalizes
        duration_days=duration_days,
        budget_crore=budget_crore,
        start_date=start_date,
        project_type="building",
        built_up_area_sqm=area_sqm,
        number_of_floors=floors,
        building_type=args.building_type,
        construction_type=args.construction_type,
        quality_finish=args.quality,
        basement=args.basement,
        number_of_blocks=blocks,
        foundation_type=args.foundation_type,
        floor_wise_structural_detail=args.floor_wise_structural,
    )
    guidelines, assumptions = _auto_config("building", args.guidelines, args.assumptions)
    output = _output_file_path(args.output, _CONFIG["building"]["output"])
    run_project_planning(proj_input, guidelines, assumptions, output, not args.no_accel_cost)


def _prompt(prompt_text: str, default: str) -> str:
    try:
        val = input(prompt_text).strip()
        return val if val else default
    except EOFError:
        return default


def main():
    # ---- subcommand dispatch ----
    if len(sys.argv) > 1:
        cmd = sys.argv[1]

        if cmd == "pipeline":
            parser = _build_pipeline_parser()
            args = parser.parse_args(sys.argv[2:])
            run_pipeline(args)
            return

        if cmd == "building":
            parser = _build_building_parser()
            args = parser.parse_args(sys.argv[2:])
            run_building(args)
            return

        if cmd == "update-plan":
            parser = _build_update_plan_parser()
            args = parser.parse_args(sys.argv[2:])
            run_update_plan(args)
            return

        if cmd == "road":
            parser = _build_road_parser()
            args = parser.parse_args(sys.argv[2:])
            run_road(args)
            return

    # ---- legacy / default: road planning (backward compat) ----
    parser = argparse.ArgumentParser(description="Construction Project Planner — Road / Pipeline / Building")
    parser.add_argument("--title", type=str, default="")
    parser.add_argument("--length", type=float, default=0.0)
    parser.add_argument("--duration", type=str, default="")
    parser.add_argument("--budget", type=str, default="")
    parser.add_argument("--start-date", type=str, default="")
    parser.add_argument("--terrain", type=str, choices=("plain", "hilly"), default="")
    parser.add_argument("--guidelines", type=str, default="",
                        help="Path to guidelines JSON (defaults to project_config/road/road_guidelines.json)")
    parser.add_argument("--assumptions", type=str, default="",
                        help="Path to assumptions YAML (defaults to project_config/road/road_assumptions.yaml)")
    parser.add_argument("--output", type=str, default="")
    parser.add_argument("--lanes", type=int, default=0)
    parser.add_argument("--no-accel-cost", dest="no_accel_cost", action="store_true")
    args = parser.parse_args()

    # Wrap into a Namespace the road runner expects
    args.terrain = args.terrain or ""
    run_road(args)


if __name__ == "__main__":
    try:
        main()
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
