"""
actual_to_entity.py
===================
Reads the Plan sheet and Engineer_DailyLog sheet from any generated
Excel workbook (road / pipeline / building — same sheet structure
regardless of project type) and builds all 14 entity-type instances.

The Excel structure is always identical:

    Plan sheet:
        Row 1  = banner
        Row 2  = subtitle
        Row 3  = column headers  (Activity ID, Activity, Phase, Type, ...)
        Row 4+ = data rows

    Engineer_DailyLog sheet:
        Row 1  = banner
        Row 2  = subtitle
        Row 3  = column headers  (Project Day, Date, Activity ID, Crew, ...)
        Row 4+ = data rows

Only the data values differ between road/pipeline/building.
The project_type field tells the API which activity catalog applies.

Usage
-----
    # From main.py after generate:
    from scripts.actual_to_entity import build_entities
    entities = build_entities(xlsx_path, project_type="road")
    # entities is a dict: { slug: [instance, ...] }

    # CLI:
    python scripts/actual_to_entity.py \\
        --xlsx output/road_project_schedule.xlsx \\
        --project-type road \\
        --output output/road_actual_data.json \\
        [--seed 42]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import random
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import openpyxl


# ─────────────────────────────────────────────────────────────────────────────
# helpers
# ─────────────────────────────────────────────────────────────────────────────

def _rng(project_id: str, seed: int = 42) -> random.Random:
    raw = hashlib.sha256(f"{project_id}:{seed}".encode()).hexdigest()
    return random.Random(int(raw[:16], 16))


def _f(v) -> float:
    try:
        return float(v) if v is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


def _i(v, floor: int = 0) -> int:
    try:
        return max(floor, int(float(v)) if v is not None else floor)
    except (TypeError, ValueError):
        return floor


def _date(v) -> Optional[date]:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(v).strip()[:10], fmt).date()
        except ValueError:
            pass
    return None


def _iso(d) -> Optional[str]:
    if d is None:
        return None
    return d.isoformat() if isinstance(d, (date, datetime)) else str(d)


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(text).lower()).strip("_")


def _vary(rng: random.Random, val: float, pct: float = 0.05) -> float:
    return max(0.0, val * (1 + rng.uniform(-pct, pct)))


# ─────────────────────────────────────────────────────────────────────────────
# Excel readers  (same logic for every project type)
# ─────────────────────────────────────────────────────────────────────────────

def _header_row(ws, anchor: str) -> int:
    """Find the row where column A equals anchor (within first 8 rows)."""
    for i in range(1, 9):
        if ws.cell(i, 1).value == anchor:
            return i
    return 3  # default: our sheets always put headers at row 3


def _col_map(ws, header_row: int) -> Dict[str, int]:
    """Return {header_text: column_number} for the header row."""
    m: Dict[str, int] = {}
    for c in range(1, 50):
        v = ws.cell(header_row, c).value
        if v:
            m[str(v).strip()] = c
    return m


# ── Plan sheet ────────────────────────────────────────────────────────────────

@dataclass
class _PlanRow:
    activity_id: str
    activity_name: str
    phase: str
    milestone: str
    start: Optional[date]
    finish: Optional[date]
    duration_days: int
    quantity: float
    unit: str
    productivity: float
    crew_multiplier: int
    status: str
    source_unit_rate: float
    planned_unit_rate: float
    labour_cost: float
    equipment_cost: float
    fuel_cost: float
    material_cost: float
    other_cost: float
    total_cost: float
    billed: float
    labour_lines: List[Dict]   = field(default_factory=list)
    equipment_lines: List[Dict] = field(default_factory=list)
    material_lines: List[Dict] = field(default_factory=list)


def _read_plan(wb) -> List[_PlanRow]:
    ws  = wb["Plan"]
    hr  = _header_row(ws, "Activity ID")
    col = _col_map(ws, hr)

    rows: List[_PlanRow] = []
    cur:  Optional[_PlanRow] = None

    _no_mat_labels = (
        "no material", "no purchasable", "unavailable",
        "rate apportionment", "pure labour",
    )

    for rn in range(hr + 1, ws.max_row + 1):
        rtype  = ws.cell(rn, col.get("Type", 13)).value
        act_id = ws.cell(rn, 1).value

        if rtype == "Activity":
            cur = _PlanRow(
                activity_id    = str(act_id or ""),
                activity_name  = str(ws.cell(rn, col.get("Activity", 2)).value or ""),
                phase          = str(ws.cell(rn, col.get("Phase", 3)).value or ""),
                milestone      = str(ws.cell(rn, col.get("Milestone", 4)).value or ""),
                start          = _date(ws.cell(rn, col.get("Start", 5)).value),
                finish         = _date(ws.cell(rn, col.get("Finish", 6)).value),
                duration_days  = _i(ws.cell(rn, col.get("Duration (wd)", 7)).value, 1),
                quantity       = _f(ws.cell(rn, col.get("Quantity", 8)).value),
                unit           = str(ws.cell(rn, col.get("Unit", 9)).value or ""),
                productivity   = _f(ws.cell(rn, col.get("Productivity", 10)).value),
                crew_multiplier= _i(ws.cell(rn, col.get("Crew", 11)).value, 1),
                status         = str(ws.cell(rn, col.get("Status", 12)).value or "Active"),
                source_unit_rate  = _f(ws.cell(rn, col.get("Source Unit Rate (Rs)", 33)).value),
                planned_unit_rate = _f(ws.cell(rn, col.get("Planned Unit Rate (Rs)", 35)).value),
                labour_cost    = 0.0, equipment_cost = 0.0, fuel_cost   = 0.0,
                material_cost  = 0.0, other_cost     = 0.0, total_cost  = 0.0, billed = 0.0,
            )
            rows.append(cur)

        elif rtype == "Labour" and cur:
            cur.labour_lines.append({
                "category"  : str(ws.cell(rn, col.get("Item", 14)).value or ""),
                "workers"   : _i(ws.cell(rn, col.get("Qty/Count", 16)).value),
                "daily_rate": _f(ws.cell(rn, col.get("Rate", 18)).value),
                "working_days": _i(ws.cell(rn, col.get("Working Days", 19)).value),
                "total_cost": _f(ws.cell(rn, col.get("Labour Cost", 24)).value),
            })

        elif rtype == "Equipment" and cur:
            cur.equipment_lines.append({
                "name"          : str(ws.cell(rn, col.get("Item", 14)).value or ""),
                "units"         : _i(ws.cell(rn, col.get("Qty/Count", 16)).value),
                "hire_rate"     : _f(ws.cell(rn, col.get("Rate", 18)).value),
                "total_hours"   : _f(ws.cell(rn, col.get("Total Hours", 21)).value),
                "fuel_lph"      : _f(ws.cell(rn, col.get("Fuel L/hr", 22)).value),
                "fuel_litres"   : _f(ws.cell(rn, col.get("Total Fuel Qty", 23)).value),
                "fuel_cost"     : _f(ws.cell(rn, col.get("Fuel Cost", 26)).value),
                "total_cost"    : _f(ws.cell(rn, col.get("Equipment Cost", 25)).value),
            })

        elif rtype == "Material" and cur:
            name = str(ws.cell(rn, col.get("Item", 14)).value or "")
            if not any(lbl in name.lower() for lbl in _no_mat_labels):
                cur.material_lines.append({
                    "name"      : name,
                    "quantity"  : _f(ws.cell(rn, col.get("Qty/Count", 16)).value),
                    "unit"      : str(ws.cell(rn, col.get("Unit", 17)).value or ""),
                    "unit_rate" : _f(ws.cell(rn, col.get("Rate", 18)).value),
                    "total_cost": _f(ws.cell(rn, col.get("Material Cost", 28)).value),
                })

        elif rtype == "Subtotal" and cur:
            cur.labour_cost    = _f(ws.cell(rn, col.get("Labour Cost",    24)).value)
            cur.equipment_cost = _f(ws.cell(rn, col.get("Equipment Cost", 25)).value)
            cur.fuel_cost      = _f(ws.cell(rn, col.get("Fuel Cost",      26)).value)
            cur.material_cost  = _f(ws.cell(rn, col.get("Material Cost",  28)).value)
            cur.other_cost     = _f(ws.cell(rn, col.get("Other Cost",     29)).value)
            cur.total_cost     = _f(ws.cell(rn, col.get("Cost (Rs)",      30)).value)
            cur.billed         = _f(ws.cell(rn, col.get("Billed (Rs)",    31)).value)

    return [r for r in rows if r.activity_id]


# ── Engineer_DailyLog sheet ───────────────────────────────────────────────────

@dataclass
class _LogRow:
    project_day: int
    activity_working_day: int
    date: Optional[date]
    weekday: str
    activity_id: str
    activity_name: str
    phase: str
    quantity: float
    unit: str
    crew: str
    equipment: str
    materials: str
    labour_cost: float
    equipment_cost: float
    fuel_consumption: float
    fuel_cost: float
    material_cost: float
    other_cost: float
    total_cost: float
    cumulative_cost: float


def _read_log(wb) -> List[_LogRow]:
    ws  = wb["Engineer_DailyLog"]
    hr  = _header_row(ws, "Project Day")
    rows: List[_LogRow] = []

    for rn in range(hr + 1, ws.max_row + 1):
        vals = [ws.cell(rn, c).value for c in range(1, 21)]
        if not any(v is not None for v in vals):
            continue
        d = _date(vals[2])
        if d is None:
            continue
        rows.append(_LogRow(
            project_day          = _i(vals[0], 1),
            activity_working_day = _i(vals[1], 1),
            date                 = d,
            weekday              = str(vals[3] or d.strftime("%A")),
            activity_id          = str(vals[4] or ""),
            activity_name        = str(vals[5] or ""),
            phase                = str(vals[6] or ""),
            quantity             = _f(vals[7]),
            unit                 = str(vals[8] or ""),
            crew                 = str(vals[9] or ""),
            equipment            = str(vals[10] or ""),
            materials            = str(vals[11] or ""),
            labour_cost          = _f(vals[12]),
            equipment_cost       = _f(vals[13]),
            fuel_consumption     = _f(vals[14]),
            fuel_cost            = _f(vals[15]),
            material_cost        = _f(vals[16]),
            other_cost           = _f(vals[17]),
            total_cost           = _f(vals[18]),
            cumulative_cost      = _f(vals[19]),
        ))
    return rows


def _read_inputs(wb) -> Dict[str, Any]:
    """Read project inputs from Sheet 02."""
    ws  = wb["02_Project_Inputs"]
    hr  = _header_row(ws, "Parameter")
    out: Dict[str, Any] = {}
    for rn in range(hr + 1, ws.max_row + 1):
        k = ws.cell(rn, 1).value
        v = ws.cell(rn, 2).value
        if k:
            out[str(k).strip()] = v
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Entity builders
# ─────────────────────────────────────────────────────────────────────────────

def _build_project(plan_rows, inputs, project_type, project_id):
    total_planned = sum(r.billed      for r in plan_rows)
    total_exec    = sum(r.total_cost  for r in plan_rows)
    start  = min((r.start  for r in plan_rows if r.start),  default=date.today())
    finish = max((r.finish for r in plan_rows if r.finish), default=date.today())
    return {
        "projectId":               project_id,
        "projectName":             str(inputs.get("Project Length", project_id)),
        "projectType":             project_type,
        "projectStatus":           "In Progress",
        "startDate":               _iso(start),
        "estimatedCompletionDate": _iso(finish),
        "durationDays":            (finish - start).days + 1,
        "budgetINR":               _f(inputs.get("Maximum Available Budget", 0)) * 1e7,
        "contractValueINR":        total_planned,
        "executionCostINR":        total_exec,
        "residualMarginINR":       total_planned - total_exec,
        "overallActualProgress":   0.0,
        "costVarianceINR":         0.0,
        "scheduleVarianceDays":    0,
        "currency":                "INR",
    }


def _build_boq(plan_rows, project_id):
    return [{
        "boqId":          f"BOQ-{r.activity_id}",
        "projectId":      project_id,
        "boqCode":        r.activity_id,
        "description":    r.activity_name,
        "phase":          r.phase,
        "milestone":      r.milestone,
        "unit":           r.unit,
        "quantity":       r.quantity,
        "unitRate":       r.planned_unit_rate,
        "sourceUnitRate": r.source_unit_rate,
        "totalAmount":    r.billed,
        "billedAmount":   0.0,
        "amountReceived": 0.0,
    } for r in plan_rows]


def _build_activities(plan_rows, project_id):
    return [{
        "activityId":          r.activity_id,
        "projectId":           project_id,
        "boqId":               f"BOQ-{r.activity_id}",
        "activityName":        r.activity_name,
        "phase":               r.phase,
        "milestone":           r.milestone,
        "plannedStart":        _iso(r.start),
        "plannedEnd":          _iso(r.finish),
        "actualStart":         None,
        "actualEnd":           None,
        "durationDays":        r.duration_days,
        "quantity":            r.quantity,
        "unit":                r.unit,
        "productivity":        r.productivity,
        "crewMultiplier":      r.crew_multiplier,
        "plannedCostINR":      r.total_cost,
        "billedCostINR":       r.billed,
        "actualCostINR":       0.0,
        "percentageComplete":  0.0,
        "earnedValue":         0.0,
        "status":              "Not Started",
    } for r in plan_rows]


def _build_plans(plan_rows, project_id):
    out = []
    for idx, r in enumerate(plan_rows, 1):
        out.append({
            "planId":                f"PLAN{str(idx).zfill(4)}",
            "activityId":            r.activity_id,
            "projectId":             project_id,
            "plannedStart":          _iso(r.start),
            "plannedEnd":            _iso(r.finish),
            "durationDays":          r.duration_days,
            "quantity":              r.quantity,
            "unit":                  r.unit,
            "productivity":          r.productivity,
            "crewMultiplier":        r.crew_multiplier,
            "labourCostINR":         r.labour_cost,
            "equipmentCostINR":      r.equipment_cost,
            "fuelCostINR":           r.fuel_cost,
            "materialCostINR":       r.material_cost,
            "otherCostINR":          r.other_cost,
            "totalExecutionCostINR": r.total_cost,
            "labourDetail": [{
                "trade": ll["category"], "workers": ll["workers"],
                "workingDays": ll["working_days"], "dailyRate": ll["daily_rate"],
                "totalCost": ll["total_cost"],
            } for ll in r.labour_lines],
            "equipmentDetail": [{
                "name": el["name"], "units": el["units"],
                "hireRate": el["hire_rate"], "totalHours": el["total_hours"],
                "fuelLph": el["fuel_lph"], "fuelLitres": el["fuel_litres"],
                "totalCost": el["total_cost"],
            } for el in r.equipment_lines],
            "materialDetail": [{
                "name": ml["name"], "quantity": ml["quantity"],
                "unit": ml["unit"], "unitRate": ml["unit_rate"],
                "totalCost": ml["total_cost"],
            } for ml in r.material_lines],
        })
    return out


def _build_assets(plan_rows, project_id):
    """One asset per unique equipment name across all plan rows."""
    seen: Dict[str, str] = {}
    assets = []
    idx = 1
    _skip = ("Unavailable", "Plant/fuel", "Fuel allow")
    for r in plan_rows:
        for eq in r.equipment_lines:
            name = eq["name"]
            if not name or any(name.startswith(s) for s in _skip):
                continue
            if name in seen:
                continue
            aid = f"ASSET{str(idx).zfill(4)}"
            seen[name] = aid
            idx += 1
            assets.append({
                "assetId":                   aid,
                "projectId":                 project_id,
                "assetCode":                 aid,
                "assetName":                 name.replace("_", " ").title(),
                "equipmentType":             name,
                "plannedTotalHours":         eq["total_hours"],
                "actualUtilisationHrsPerDay": 0.0,
                "hoursLoggedToDate":         0.0,
                "odometerOpening":           0.0,
                "odometerClosing":           0.0,
                "status":                    "Active",
            })
    return assets, seen          # also return name→id map


def _build_materials(plan_rows, project_id):
    """One material per unique material name across all plan rows."""
    seen: Dict[str, str] = {}
    materials = []
    idx = 1
    for r in plan_rows:
        for m in r.material_lines:
            name = m["name"]
            if not name or name in seen:
                continue
            mid = f"MAT{str(idx).zfill(4)}"
            seen[name] = mid
            idx += 1
            materials.append({
                "materialId":       mid,
                "projectId":        project_id,
                "materialCode":     mid,
                "materialName":     name.replace("_", " ").title(),
                "unit":             m["unit"] or "unit",
                "unitCost":         m["unit_rate"],
                "plannedQuantity":  m["quantity"],
                "currentStock":     0.0,
                "procuredQuantity": 0.0,
                "issuedQuantity":   0.0,
                "totalCostToDate":  0.0,
            })
    # Always include diesel
    if "diesel" not in seen:
        mid = f"MAT{str(idx).zfill(4)}"
        seen["diesel"] = mid
        materials.append({
            "materialId": mid, "projectId": project_id,
            "materialCode": mid, "materialName": "Diesel",
            "unit": "litre", "unitCost": 100.0,
            "plannedQuantity": 0.0, "currentStock": 0.0,
            "procuredQuantity": 0.0, "issuedQuantity": 0.0, "totalCostToDate": 0.0,
        })
    return materials, seen       # also return name→id map


def _build_subcontractors(project_id, rng):
    """Generate realistic subcontractor + worker records."""
    subs = []
    idx  = 1
    labour_cos  = ["Sharma Construction Labour", "Reddy Workforce Services",
                   "Hyderabad Labour Suppliers", "AP Manpower Agency"]
    equip_cos   = ["Venkat Equipment Hire", "SR Plant & Machinery",
                   "HCL Earthmovers", "Deccan Construction Equipment"]
    mat_cos     = ["Vijay Building Materials", "Sri Sai Cement Traders",
                   "Southern Aggregates", "Metro Steel Suppliers"]
    worker_ids: List[str] = []

    for co in labour_cos:
        sub_id = f"SUB{str(idx).zfill(4)}"; idx += 1
        subs.append({"subContractorId": sub_id, "projectId": project_id,
                     "companyName": co, "subContractorType": "workforce",
                     "status": "Active"})
        for _ in range(3):
            wid = f"WK{str(len(worker_ids)+1).zfill(3)}"
            worker_ids.append(wid)
            subs.append({"subContractorId": wid, "projectId": project_id,
                         "companyName": co, "subContractorType": "worker",
                         "parentSubContractorId": sub_id, "status": "Active"})

    for co in equip_cos + mat_cos:
        sub_id = f"SUB{str(idx).zfill(4)}"; idx += 1
        stype = "equipment" if co in equip_cos else "material_supplier"
        subs.append({"subContractorId": sub_id, "projectId": project_id,
                     "companyName": co, "subContractorType": stype, "status": "Active"})

    return subs, worker_ids


def _build_assumptions(wb, project_id):
    """Read assumptions directly from the 03_Assumptions sheet."""
    if "03_Assumptions" not in wb.sheetnames:
        return []
    ws  = wb["03_Assumptions"]
    out = []
    idx = 1
    for row in ws.iter_rows(min_row=4, values_only=True):
        k = row[0]; v = row[1]
        if k is None:
            continue
        key = str(k).strip()
        if not key or key in ("Assumption", "ASSUMPTIONS USED BY PLAN"):
            continue
        out.append({
            "assumptionId":   f"ASMP{str(idx).zfill(4)}",
            "projectId":      project_id,
            "parameterKey":   key,
            "parameterValue": str(v) if v is not None else "",
            "unit":           str(row[2]) if len(row) > 2 and row[2] else "",
            "sourceNote":     str(row[3]) if len(row) > 3 and row[3] else "assumptions YAML",
        })
        idx += 1
    return out


# ── Operational (simulated) entities ─────────────────────────────────────────

_MACHINE_ODO = {  # km/hr per machine type
    "excavator_0_9cum": 1.2, "jcb_3dx": 2.5, "dozer_d6": 1.5,
    "motor_grader_120hp": 3.0, "tipper_10t": 18.0, "transit_mixer_6cum": 12.0,
    "vibratory_roller_8_10t": 2.0, "paver_finisher": 1.5,
    "water_tanker_12kl": 15.0, "crawler_crane_50t": 0.3,
    "light_crane_3t": 1.0, "bitumen_sprayer_4000l": 5.0,
}
_SUP_ROLES = {"supervisor", "site_engineer", "project_manager",
              "qc_technician", "safety_officer", "surveyor"}
_OPR_ROLES = {"equipment_operator", "driver"}


def _build_edl(log_rows, plan_rows, asset_name_to_id, worker_ids,
               project_id, rng):
    plan_map = {r.activity_id: r for r in plan_rows}
    emp_pool: List[Dict] = []
    emp_idx = [1]

    def _emp(role):
        eid = f"EMP{str(emp_idx[0]).zfill(4)}"
        emp_idx[0] += 1
        sub = worker_ids[emp_idx[0] % len(worker_ids)] if (emp_idx[0] % 7 == 0 and worker_ids) else None
        return {"employeeId": eid, "role": role, "subContractorId": sub}

    # Pre-build a small pool per role from plan
    by_role: Dict[str, List[Dict]] = defaultdict(list)
    for pr in plan_rows:
        for ll in pr.labour_lines:
            for _ in range(max(1, ll["workers"])):
                e = _emp(ll["category"])
                by_role[ll["category"]].append(e)
    for v in by_role.values():
        emp_pool.extend(v)

    odo: Dict[str, float] = {}
    odo_rate: Dict[str, float] = {}
    for name, aid in asset_name_to_id.items():
        speed = _MACHINE_ODO.get(name, 0.0)
        odo_rate[aid] = speed * (0.9 + rng.random() * 0.2)
        odo[aid] = 0.0

    edl_list: List[Dict] = []
    sup_map: Dict[Tuple, List[str]] = defaultdict(list)
    edl_idx = 1

    for log in log_rows:
        pr = plan_map.get(log.activity_id)
        if not pr:
            continue
        total_w  = max(1, sum(ll["workers"] for ll in pr.labour_lines))
        eq_names = [eq["name"] for eq in pr.equipment_lines
                    if not eq["name"].startswith(("Unavailable", "Plant/fuel", "Fuel allow"))]
        date_iso = _iso(log.date)

        for ll in pr.labour_lines:
            role = ll["category"]
            if role in _SUP_ROLES:
                continue
            workers = max(1, ll["workers"])
            pool    = by_role.get(role, [_emp(role)])
            share   = ll["workers"] / total_w / max(1, workers)

            for w in range(min(workers, max(1, len(pool)))):
                emp      = pool[w % len(pool)]
                is_opr   = role in _OPR_ROLES
                asset_id = ""
                odo_o = odo_c = 0.0

                if is_opr and eq_names:
                    eq_name  = eq_names[w % len(eq_names)]
                    asset_id = asset_name_to_id.get(eq_name, "")
                    if asset_id:
                        spd   = odo_rate.get(asset_id, 0.0)
                        km    = spd * log.equipment_cost / max(1, pr.equipment_cost) * 8 \
                                if pr.equipment_cost > 0 else 0.0
                        odo_o = odo.get(asset_id, 0.0)
                        odo_c = odo_o + km
                        odo[asset_id] = odo_c
                        key = (date_iso, log.activity_id, asset_id)
                        sups = by_role.get("supervisor", [])
                        if sups:
                            sid = sups[edl_idx % len(sups)]["employeeId"]
                            if sid not in sup_map[key]:
                                sup_map[key].append(sid)

                edl_list.append({
                    "engineeringDailyLogId": f"EDL{str(edl_idx).zfill(5)}",
                    "projectId":             project_id,
                    "activityId":            log.activity_id,
                    "assetId":               asset_id,
                    "employeeId":            emp["employeeId"],
                    "employeeRole":          role,
                    "subContractorId":       emp.get("subContractorId") or "",
                    "date":                  date_iso,
                    "dayNumber":             log.project_day,
                    "activityWorkingDay":    log.activity_working_day,
                    "shift":                 "shift-single",
                    "workDone":              round(_vary(rng, log.quantity * share, 0.03), 4),
                    "unit":                  log.unit,
                    "labourCostINR":         round(_vary(rng, log.labour_cost   * share, 0.03), 2),
                    "equipmentCostINR":      round(_vary(rng, log.equipment_cost * share, 0.03), 2) if is_opr else 0.0,
                    "fuelCostINR":           round(_vary(rng, log.fuel_cost      * share, 0.03), 2) if is_opr else 0.0,
                    "fuelLitres":            round(_vary(rng, log.fuel_consumption * share, 0.03), 3) if is_opr else 0.0,
                    "odometerOpening":       round(odo_o, 1),
                    "odometerClosing":       round(odo_c, 1),
                })
                edl_idx += 1

    return edl_list, dict(sup_map)


def _build_mir(log_rows, plan_rows, mat_id_map, worker_ids, project_id, rng):
    plan_map   = {r.activity_id: r for r in plan_rows}
    diesel_id  = mat_id_map.get("diesel", "MAT9999")
    mir_list: List[Dict] = []
    idx = 1

    for log in log_rows:
        pr = plan_map.get(log.activity_id)
        if not pr:
            continue
        date_iso  = _iso(log.date)
        requestor = worker_ids[idx % len(worker_ids)] if worker_ids else "WK001"

        for ml in pr.material_lines:
            mid = mat_id_map.get(ml["name"])
            if not mid:
                continue
            qty = _vary(rng, ml["quantity"] / max(1, pr.duration_days), 0.05)
            if qty <= 0:
                continue
            mir_list.append({
                "materialIssueRequestId": f"MIR{str(idx).zfill(5)}",
                "projectId":   project_id,
                "materialId":  mid,
                "activityId":  log.activity_id,
                "date":        date_iso,
                "materialName": ml["name"].replace("_", " ").title(),
                "quantity":    round(qty, 4),
                "unit":        ml["unit"] or "unit",
                "unitCost":    ml["unit_rate"],
                "totalCost":   round(qty * ml["unit_rate"], 2),
                "status":      "Issued",
                "personRequested": requestor,
            })
            idx += 1

        if log.fuel_cost > 0:
            fuel_qty = _vary(rng, log.fuel_consumption, 0.05)
            mir_list.append({
                "materialIssueRequestId": f"MIR{str(idx).zfill(5)}",
                "projectId":   project_id,
                "materialId":  diesel_id,
                "activityId":  log.activity_id,
                "date":        date_iso,
                "materialName": "Diesel",
                "quantity":    round(fuel_qty, 3),
                "unit":        "litre",
                "unitCost":    100.0,
                "totalCost":   round(fuel_qty * 100.0, 2),
                "status":      "Issued",
                "personRequested": worker_ids[idx % len(worker_ids)] if worker_ids else "WK001",
            })
            idx += 1

    return mir_list


def _build_mpr(mir_list, mat_id_map, project_id, rng):
    groups: Dict[Tuple, float] = defaultdict(float)
    meta: Dict[str, Dict]      = {}

    for mir in mir_list:
        key = (mir["materialId"], mir["date"])
        groups[key] += mir["quantity"]
        meta[mir["materialId"]] = {
            "name": mir["materialName"], "unit": mir["unit"], "rate": mir["unitCost"]}

    vendors    = ["Sri Sai Suppliers", "Vijay Traders", "AP Materials", "Southern Supplies"]
    mpr_list: List[Dict] = []
    idx = 1

    for (mid, cons_date_str), qty in sorted(groups.items()):
        cons_date = date.fromisoformat(cons_date_str)
        proc_date = cons_date - timedelta(days=rng.randint(5, 6))
        while proc_date.weekday() == 6:
            proc_date -= timedelta(days=1)
        proc_qty  = qty * 1.05
        rate      = meta.get(mid, {}).get("rate", 100.0)
        mpr_list.append({
            "materialProcurementRequestId": f"MPR{str(idx).zfill(5)}",
            "projectId":   project_id,
            "materialId":  mid,
            "requestNumber": f"MPR-{str(idx).zfill(5)}",
            "materialName": meta.get(mid, {}).get("name", mid),
            "quantity":    round(proc_qty, 4),
            "unit":        meta.get(mid, {}).get("unit", "unit"),
            "requestDate": proc_date.isoformat(),
            "deliveryDate": cons_date.isoformat(),
            "status":      "Delivered",
            "unitRate":    rate,
            "totalCost":   round(proc_qty * rate, 2),
            "vendorName":  rng.choice(vendors),
            "logisticsId": f"LOG{str(idx).zfill(5)}",
        })
        idx += 1

    return mpr_list


def _build_logistics(mpr_list, mir_list, project_id, rng):
    transporters = ["Ravi Logistics", "Sai Transport", "SR Carriers", "AP Freight"]
    vehicles     = ["AP09AB1234", "TS07CD5678", "KA03EF9012", "MH12GH3456"]
    drivers      = ["Suresh D.", "Raju T.", "Mohan K.", "Anil V."]
    logi = []
    idx  = 1

    for mpr in mpr_list:
        trips = rng.randint(2, 6)
        logi.append({
            "logisticsId":   mpr.get("logisticsId", f"LOG{str(idx).zfill(5)}"),
            "projectId":     project_id,
            "materialProcurementRequestId": mpr["materialProcurementRequestId"],
            "materialIssueRequestId": None,
            "movementType":  "INFLOW",
            "materialName":  mpr["materialName"],
            "totalQuantity": mpr["quantity"],
            "unit":          mpr["unit"],
            "transporter":   rng.choice(transporters),
            "vehicle":       rng.choice(vehicles),
            "driver":        rng.choice(drivers),
            "sourceAddress": "Supplier Warehouse, Hyderabad",
            "deliveryAddress": "Construction Site",
            "dispatchDate":  mpr["requestDate"],
            "deliveryDate":  mpr["deliveryDate"],
            "tripCount":     trips,
            "status":        "Delivered",
        })
        idx += 1

    for mir in mir_list[:50]:   # sample of outflow logistics
        logi.append({
            "logisticsId":   f"LOG{str(idx).zfill(5)}",
            "projectId":     project_id,
            "materialProcurementRequestId": None,
            "materialIssueRequestId": mir["materialIssueRequestId"],
            "movementType":  "OUTFLOW",
            "materialName":  mir["materialName"],
            "totalQuantity": mir["quantity"],
            "unit":          mir["unit"],
            "transporter":   "Internal",
            "vehicle":       "Site Vehicle",
            "driver":        rng.choice(drivers),
            "sourceAddress": "Site Store",
            "deliveryAddress": "Working Site",
            "dispatchDate":  mir["date"],
            "deliveryDate":  mir["date"],
            "tripCount":     1,
            "status":        "Issued",
        })
        idx += 1

    return logi


def _build_transactions(mpr_list, mir_list, project_id):
    txns = []
    idx  = 1
    for mpr in mpr_list:
        txns.append({
            "materialTransactionId": f"MTX{str(idx).zfill(5)}",
            "projectId":    project_id,
            "materialId":   mpr["materialId"],
            "materialName": mpr["materialName"],
            "transactionType": "INFLOW",
            "quantity":     mpr["quantity"],
            "unit":         mpr["unit"],
            "unitCost":     mpr["unitRate"],
            "totalCost":    mpr["totalCost"],
            "date":         mpr["deliveryDate"],
            "referenceId":  mpr["materialProcurementRequestId"],
            "referenceType": "MaterialProcurementRequest",
            "runningStock": 0.0,
        })
        idx += 1
    for mir in mir_list:
        txns.append({
            "materialTransactionId": f"MTX{str(idx).zfill(5)}",
            "projectId":    project_id,
            "materialId":   mir["materialId"],
            "materialName": mir["materialName"],
            "transactionType": "OUTFLOW",
            "quantity":     mir["quantity"],
            "unit":         mir["unit"],
            "unitCost":     mir["unitCost"],
            "totalCost":    mir["totalCost"],
            "date":         mir["date"],
            "referenceId":  mir["materialIssueRequestId"],
            "referenceType": "MaterialIssueRequest",
            "runningStock": 0.0,
        })
        idx += 1

    txns.sort(key=lambda t: (t["date"], 0 if t["transactionType"] == "INFLOW" else 1))
    stock: Dict[str, float] = defaultdict(float)
    for t in txns:
        mid = t["materialId"]
        if t["transactionType"] == "INFLOW":
            stock[mid] += t["quantity"]
        else:
            stock[mid] = max(0.0, stock[mid] - t["quantity"])
        t["runningStock"] = round(stock[mid], 4)

    return txns


def _build_supervisor_updates(log_rows, plan_rows, edl_list, mir_list,
                               project_id, rng):
    plan_map = {r.activity_id: r for r in plan_rows}
    remarks  = ["No issues", "Delayed due to weather", "Progressing on schedule",
                "Material delivery pending", "Equipment breakdown resolved"]

    edl_by = defaultdict(list)
    for e in edl_list:
        edl_by[(e["date"], e["activityId"])].append(e)

    mir_by = defaultdict(list)
    for m in mir_list:
        mir_by[(m["date"], m["activityId"])].append(m)

    sup_list: List[Dict] = []
    idx = 1

    for log in log_rows:
        pr = plan_map.get(log.activity_id)
        if not pr:
            continue
        date_iso = _iso(log.date)
        key      = (date_iso, log.activity_id)
        day_edls = edl_by.get(key, [])
        day_mirs = mir_by.get(key, [])
        if not day_edls:
            continue

        role_counts: Dict[str, int] = defaultdict(int)
        for e in day_edls:
            role_counts[e["employeeRole"]] += 1
        for ll in pr.labour_lines:
            if ll["category"] in _SUP_ROLES:
                role_counts[ll["category"]] += max(1, ll["workers"])

        machinery: Dict[str, int] = defaultdict(int)
        for e in day_edls:
            if e.get("assetId"):
                machinery[e["assetId"]] += 1

        mat_recv: Dict[str, float] = defaultdict(float)
        for m in day_mirs:
            mat_recv[m["materialName"]] += m["quantity"]

        sup_list.append({
            "supervisorUpdateId":      f"SUP{str(idx).zfill(5)}",
            "projectId":               project_id,
            "activityId":              log.activity_id,
            "boqId":                   f"BOQ-{log.activity_id}",
            "date":                    date_iso,
            "dayNumber":               log.project_day,
            "totalSkilled":            sum(v for k, v in role_counts.items() if k not in _SUP_ROLES),
            "totalUnskilled":          role_counts.get("unskilled", 0),
            "totalSupervisors":        sum(v for k, v in role_counts.items() if k in _SUP_ROLES),
            "manpower":                dict(role_counts),
            "machinery":               dict(machinery),
            "materialsReceived":       dict(mat_recv),
            "plannedProgress":         round(log.quantity, 4),
            "actualProgress":          round(_vary(rng, log.quantity, 0.05), 4),
            "unit":                    log.unit,
            "labourCostINR":           round(_vary(rng, log.labour_cost, 0.03), 2),
            "equipmentCostINR":        round(_vary(rng, log.equipment_cost, 0.03), 2),
            "materialCostINR":         round(_vary(rng, log.material_cost, 0.05), 2),
            "totalCostINR":            round(_vary(rng, log.total_cost, 0.03), 2),
            "engineeringDailyLogIds":  [e["engineeringDailyLogId"] for e in day_edls],
            "materialIssueRequestIds": [m["materialIssueRequestId"] for m in day_mirs],
            "remarks":                 rng.choice(remarks),
        })
        idx += 1

    return sup_list


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def build_entities(xlsx_path: str,
                   project_type: str = "road",
                   seed: int = 42) -> Dict[str, List[Dict]]:
    """
    Read Plan + Engineer_DailyLog sheets from xlsx_path and return all
    14 entity types as a dict  { slug: [instance_dict, ...] }.

    Parameters
    ----------
    xlsx_path    : path to the Excel workbook (any project type)
    project_type : "road" | "pipeline" | "building"
    seed         : RNG seed for reproducible IDs (default 42)
    """
    wb         = openpyxl.load_workbook(xlsx_path, data_only=True)
    plan_rows  = _read_plan(wb)
    log_rows   = _read_log(wb)
    inputs     = _read_inputs(wb)

    project_id = _slug(
        str(inputs.get("Project Length", "") or inputs.get("title", "project"))
    )[:60] or "project_001"
    r = _rng(project_id, seed)

    project    = _build_project(plan_rows, inputs, project_type, project_id)
    boq        = _build_boq(plan_rows, project_id)
    activities = _build_activities(plan_rows, project_id)
    plans      = _build_plans(plan_rows, project_id)
    assumptions= _build_assumptions(wb, project_id)
    assets, asset_map = _build_assets(plan_rows, project_id)
    materials, mat_map = _build_materials(plan_rows, project_id)
    subcontractors, worker_ids = _build_subcontractors(project_id, r)

    edl_list, sup_map = _build_edl(log_rows, plan_rows, asset_map, worker_ids, project_id, r)
    mir_list   = _build_mir(log_rows, plan_rows, mat_map, worker_ids, project_id, r)
    mpr_list   = _build_mpr(mir_list, mat_map, project_id, r)
    logistics  = _build_logistics(mpr_list, mir_list, project_id, r)
    txns       = _build_transactions(mpr_list, mir_list, project_id)
    sup_updates= _build_supervisor_updates(log_rows, plan_rows, edl_list, mir_list, project_id, r)

    return {
        "project":                      [project],
        "assumption":                   assumptions,
        "asset":                        assets,
        "material":                     materials,
        "subcontractor":                subcontractors,
        "boq":                          boq,
        "activity":                     activities,
        "plan":                         plans,
        "engineering_daily_logs":       edl_list,
        "material_issue_request":       mir_list,
        "material_procurement_request": mpr_list,
        "logistics":                    logistics,
        "material_transaction":         txns,
        "supervisor_update":            sup_updates,
    }


# ─────────────────────────────────────────────────────────────────────────────
# CLI  (for standalone use)
# ─────────────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(
        description="Read Plan+DailyLog sheets → 14 entity-type JSON")
    p.add_argument("--xlsx",         required=True,
                   help="Path to generated Excel workbook")
    p.add_argument("--project-type", default="road",
                   choices=["road", "pipeline", "building"])
    p.add_argument("--output",       required=True,
                   help="Output JSON file path")
    p.add_argument("--seed",         type=int, default=42)
    args = p.parse_args()

    print(f"\nReading: {args.xlsx}  (type={args.project_type})")
    entities = build_entities(args.xlsx, args.project_type, args.seed)

    print("Entity counts:")
    for slug, rows in entities.items():
        print(f"  {slug:<40} {len(rows):>6}")

    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(entities, f, indent=2, ensure_ascii=False, default=str)
    print(f"\nOutput: {os.path.abspath(args.output)}\n")


if __name__ == "__main__":
    main()
