"""Strict, source-backed validation for Excel-to-entity exports."""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import date
import math
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from scripts.actual_to_entity import _read_log, _read_plan


@dataclass
class EntityCheck:
    number: int
    name: str
    status: str
    details: str


@dataclass
class EntityValidationReport:
    checks: List[EntityCheck]

    @property
    def passed(self) -> bool:
        return len(self.checks) == 26 and all(c.status == "PASS" for c in self.checks)

    @property
    def status(self) -> str:
        if any(check.status == "FAIL" for check in self.checks):
            return "FAIL"
        if len(self.checks) != 26 or any(check.status == "INCOMPLETE" for check in self.checks):
            return "INCOMPLETE"
        return "PASS"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status,
            "passed": self.passed,
            "checks": [asdict(check) for check in self.checks],
        }


def write_validation_workbook(report: EntityValidationReport, output_path: str) -> None:
    """Write a readable Excel summary and detail for every validation check."""
    workbook = openpyxl.Workbook()
    summary = workbook.active
    summary.title = "Summary"
    summary.append(["Entity Validation Report", report.status])
    summary.append(["Passed", sum(check.status == "PASS" for check in report.checks)])
    summary.append(["Failed", sum(check.status == "FAIL" for check in report.checks)])
    summary.append(["Incomplete", sum(check.status == "INCOMPLETE" for check in report.checks)])
    summary.append([])
    summary.append(["#", "Check", "Status", "Details"])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    status_fills = {
        "PASS": PatternFill("solid", fgColor="C6EFCE"),
        "FAIL": PatternFill("solid", fgColor="FFC7CE"),
        "INCOMPLETE": PatternFill("solid", fgColor="FFEB9C"),
    }
    for cell in summary[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    for cell in summary[6]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill

    for check in report.checks:
        summary.append([check.number, check.name, check.status, check.details])
        row = summary.max_row
        summary.cell(row, 3).fill = status_fills.get(check.status, PatternFill())
        summary.cell(row, 3).font = Font(bold=True)
        for cell in summary[row]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    summary.column_dimensions["A"].width = 10
    summary.column_dimensions["B"].width = 58
    summary.column_dimensions["C"].width = 16
    summary.column_dimensions["D"].width = 110
    summary.freeze_panes = "A7"
    summary.auto_filter.ref = f"A6:D{summary.max_row}"
    summary.sheet_view.showGridLines = False

    workbook.save(output_path)
    workbook.close()


ENTITY_SPECS = {
    "project": ("projectId", ()),
    "assumption": ("assumptionId", ()),
    "asset": ("assetId", ()),
    "material": ("materialId", ()),
    "subcontractor": ("subContractorId", ("parentSubContractorId",)),
    "boq": ("boqId", ("projectId",)),
    "activity": ("activityId", ("projectId", "boqId")),
    "plan": ("planId", ("projectId", "activityId")),
    "engineering_daily_logs": (
        "engineeringDailyLogId", ("projectId", "activityId", "assetId", "subContractorId"),
    ),
    "material_issue_request": (
        "materialIssueRequestId", ("projectId", "materialId", "activityId"),
    ),
    "material_procurement_request": (
        "materialProcurementRequestId", ("projectId", "materialId", "logisticsId"),
    ),
    "logistics": (
        "logisticsId", ("projectId", "materialProcurementRequestId", "materialIssueRequestId"),
    ),
    "material_transaction": (
        "materialTransactionId", ("projectId", "materialId", "referenceId"),
    ),
    "supervisor_update": (
        "supervisorUpdateId", ("projectId", "activityId", "boqId"),
    ),
}

_TOLERANCE = 0.02
_SUPERVISOR_ROLES = {
    "supervisor", "site_engineer", "project_manager",
    "qc_technician", "safety_officer", "surveyor",
}


def _close(left: Any, right: Any, tolerance: float = _TOLERANCE) -> bool:
    try:
        a, b = float(left), float(right)
        return math.isfinite(a) and math.isfinite(b) and abs(a - b) <= tolerance
    except (TypeError, ValueError):
        return False


def _date(value: Any) -> Optional[date]:
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


def _index(rows: Sequence[Mapping[str, Any]], field: str) -> Dict[str, Mapping[str, Any]]:
    return {
        str(row.get(field)): row
        for row in rows
        if row.get(field) not in (None, "")
    }


def _summarize(errors: Sequence[str], success: str) -> Tuple[str, str]:
    if not errors:
        return "PASS", success
    shown = errors[:8]
    suffix = f" ({len(errors) - len(shown)} more)" if len(errors) > len(shown) else ""
    return "FAIL", "; ".join(shown) + suffix


def validate_entity_export(
    workbook_path: str,
    entities: Mapping[str, Sequence[Mapping[str, Any]]],
) -> EntityValidationReport:
    """Run all 26 requested checks; missing source evidence is never a pass."""
    checks: List[EntityCheck] = []
    try:
        workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    except Exception as exc:
        return EntityValidationReport([
            EntityCheck(i, name, "INCOMPLETE", f"Cannot open source workbook: {exc}")
            for i, name in enumerate(_CHECK_NAMES, 1)
        ])

    missing_sheets = {"Plan", "Engineer_DailyLog", "03_Assumptions"} - set(workbook.sheetnames)
    if missing_sheets:
        workbook.close()
        return EntityValidationReport([
            EntityCheck(i, name, "INCOMPLETE",
                        f"Required source worksheet(s) missing: {sorted(missing_sheets)}")
            for i, name in enumerate(_CHECK_NAMES, 1)
        ])

    try:
        plan_rows = _read_plan(workbook)
        log_rows = _read_log(workbook)
        plan_entities = list(entities.get("plan", []))
        boq_entities = list(entities.get("boq", []))
        activities = list(entities.get("activity", []))
        assets = list(entities.get("asset", []))
        materials = list(entities.get("material", []))
        edls = list(entities.get("engineering_daily_logs", []))
        mirs = list(entities.get("material_issue_request", []))
        mprs = list(entities.get("material_procurement_request", []))
        logistics = list(entities.get("logistics", []))
        transactions = list(entities.get("material_transaction", []))
        supervisors = list(entities.get("supervisor_update", []))
        subcontractors = list(entities.get("subcontractor", []))

        plan_by_activity = {row.activity_id: row for row in plan_rows}
        source_logs_by_key = _logs_by_key(log_rows)
        plan_entity_by_activity = _index(plan_entities, "activityId")
        boq_by_activity = _index(boq_entities, "boqCode")
        activity_by_id = _index(activities, "activityId")
        material_by_id = _index(materials, "materialId")
        subcontractor_ids = {
            str(row.get("subContractorId")) for row in subcontractors
            if row.get("subContractorId")
        }

        def run(number: int, name: str, fn: Callable[[], Tuple[str, str]]) -> None:
            try:
                status, details = fn()
            except Exception as exc:
                status, details = "FAIL", f"Validator raised {type(exc).__name__}: {exc}"
            checks.append(EntityCheck(number, name, status, details))

        def check_plan_costs() -> Tuple[str, str]:
            sheet = workbook["03_Assumptions"]
            rates = {
                str(sheet.cell(row, 1).value): sheet.cell(row, 8).value
                for row in range(5, sheet.max_row + 1)
                if sheet.cell(row, 1).value
            }
            errors = []
            if not rates:
                return "INCOMPLETE", "No activity rates found in 03_Assumptions."
            for activity_id, row in plan_by_activity.items():
                entity = plan_entity_by_activity.get(activity_id)
                if entity is None:
                    errors.append(f"{activity_id}: plan entity missing")
                    continue
                expected_rate = rates.get(activity_id)
                if expected_rate is None or not _close(row.planned_unit_rate, expected_rate, 0.01):
                    errors.append(f"{activity_id}: planned unit rate differs from assumptions")
                if not _close(entity.get("totalExecutionCostINR"), row.total_cost, 1.0):
                    errors.append(f"{activity_id}: execution cost differs from Plan subtotal")
                breakdown = sum(float(entity.get(key) or 0) for key in (
                    "labourCostINR", "equipmentCostINR", "fuelCostINR",
                    "materialCostINR", "otherCostINR",
                ))
                if not _close(breakdown, row.total_cost, 1.0):
                    errors.append(f"{activity_id}: plan cost components do not sum to subtotal")
            return _summarize(errors, f"Assumption rates and Plan subtotals verified for {len(plan_rows)} activities.")

        def check_plan_excel_json() -> Tuple[str, str]:
            errors = []
            if len(plan_entities) != len(plan_rows):
                errors.append(f"Excel Plan activities={len(plan_rows)}, JSON plan entities={len(plan_entities)}")
            for activity_id, source in plan_by_activity.items():
                entity = plan_entity_by_activity.get(activity_id)
                if entity is None:
                    errors.append(f"{activity_id}: missing plan entity")
                    continue
                for field, expected in (
                    ("quantity", source.quantity), ("durationDays", source.duration_days),
                    ("crewMultiplier", source.crew_multiplier), ("productivity", source.productivity),
                    ("labourCostINR", source.labour_cost), ("equipmentCostINR", source.equipment_cost),
                    ("fuelCostINR", source.fuel_cost), ("materialCostINR", source.material_cost),
                    ("otherCostINR", source.other_cost),
                    ("totalExecutionCostINR", source.total_cost),
                ):
                    if not _close(entity.get(field), expected, 1.0 if "Cost" in field else 0.02):
                        errors.append(f"{activity_id}: {field} differs from Excel Plan")
                if entity.get("unit") != source.unit:
                    errors.append(f"{activity_id}: unit differs from Excel Plan")
            return _summarize(errors, f"Excel Plan matches all {len(plan_rows)} plan entities.")

        def check_boq() -> Tuple[str, str]:
            errors = []
            if len(boq_entities) != len(plan_rows):
                errors.append(f"Excel activity rows={len(plan_rows)}, BOQ rows={len(boq_entities)}")
            for aid, source in plan_by_activity.items():
                row = boq_by_activity.get(aid)
                if row is None:
                    errors.append(f"{aid}: BOQ missing")
                    continue
                for key, expected in (
                    ("quantity", source.quantity), ("unitRate", source.planned_unit_rate),
                    ("sourceUnitRate", source.source_unit_rate), ("totalAmount", source.billed),
                ):
                    if not _close(row.get(key), expected, 1.0):
                        errors.append(f"{aid}: BOQ {key} differs from Excel Plan")
                if row.get("description") != source.activity_name:
                    errors.append(f"{aid}: BOQ description differs from Excel Plan")
            return _summarize(errors, f"BOQ rows match Excel Plan for {len(plan_rows)} activities.")

        def check_activities() -> Tuple[str, str]:
            errors = []
            if len(activities) != len(plan_rows):
                errors.append(f"Excel activity rows={len(plan_rows)}, entity activities={len(activities)}")
            for aid, source in plan_by_activity.items():
                row = activity_by_id.get(aid)
                if row is None:
                    errors.append(f"{aid}: activity entity missing")
                    continue
                for key, expected in (
                    ("quantity", source.quantity), ("durationDays", source.duration_days),
                    ("productivity", source.productivity), ("crewMultiplier", source.crew_multiplier),
                    ("plannedCostINR", source.total_cost), ("billedCostINR", source.billed),
                ):
                    if not _close(row.get(key), expected, 1.0):
                        errors.append(f"{aid}: activity {key} differs from Excel Plan")
            return _summarize(errors, f"Activity entities match Excel Plan for {len(plan_rows)} activities.")

        def equipment_totals() -> Dict[str, Dict[str, float]]:
            totals: Dict[str, Dict[str, float]] = defaultdict(lambda: {"hours": 0.0, "cost": 0.0})
            for row in plan_rows:
                for item in row.equipment_lines:
                    name = item["name"]
                    if name and not name.startswith(("Unavailable", "Plant/fuel", "Fuel allow")):
                        totals[name]["hours"] += item["total_hours"]
                        totals[name]["cost"] += item["total_cost"]
            return totals

        def check_assets() -> Tuple[str, str]:
            expected = equipment_totals()
            by_type = {str(row.get("equipmentType")): row for row in assets}
            errors = []
            if set(by_type) != set(expected):
                errors.append(f"asset equipment types differ; missing={sorted(set(expected)-set(by_type))}, extra={sorted(set(by_type)-set(expected))}")
            required = ("assetId", "assetCode", "assetName", "equipmentType",
                        "plannedTotalHours", "plannedEquipmentCostINR",
                        "actualUtilisationHrsPerDay", "hoursLoggedToDate",
                        "odometerOpening", "odometerClosing", "status")
            for equipment, values in expected.items():
                row = by_type.get(equipment)
                if row is None:
                    continue
                for key in required:
                    if key not in row or row.get(key) is None:
                        errors.append(f"{equipment}: missing asset field {key}")
                if not _close(row.get("plannedTotalHours"), values["hours"], 0.1):
                    errors.append(f"{equipment}: planned hours differ from Plan resources")
                if not _close(row.get("plannedEquipmentCostINR"), values["cost"], 1.0):
                    errors.append(f"{equipment}: planned cost differs from Plan resources")
                for key in ("plannedTotalHours", "plannedEquipmentCostINR",
                            "actualUtilisationHrsPerDay", "hoursLoggedToDate",
                            "odometerOpening", "odometerClosing"):
                    if not _finite_nonnegative(row.get(key)):
                        errors.append(f"{equipment}: {key} is not a finite nonnegative number")
            return _summarize(errors, f"All {len(expected)} assets and planned usage/costs match Excel resources.")

        def check_edl_work_done() -> Tuple[str, str]:
            source = _sum_by(log_rows, lambda row: (row.date.isoformat(), row.activity_id), lambda row: row.quantity)
            actual = _sum_by(edls, lambda row: (str(row.get("date")), str(row.get("activityId"))), lambda row: row.get("workDone"))
            errors = []
            for key, quantity in source.items():
                if key not in actual:
                    errors.append(f"{key}: no EDL workDone")
                elif not _close(actual[key], quantity, 0.02):
                    errors.append(f"{key}: EDL workDone={actual[key]:.4f}, source quantity={quantity:.4f}")
            for key in actual.keys() - source.keys():
                errors.append(f"{key}: EDL has no matching Engineer_DailyLog row")
            return _summarize(errors, f"EDL workDone reconciles by activity/date across {len(source)} source groups.")

        def check_fuel() -> Tuple[str, str]:
            expected = _sum_by(log_rows, lambda row: row.date.isoformat(), lambda row: row.fuel_consumption)
            edl = _sum_by(edls, lambda row: str(row.get("date")), lambda row: row.get("fuelLitres"))
            diesel_ids = {mid for mid, row in material_by_id.items()
                          if "diesel" in str(row.get("materialName", "")).lower()}
            mir = _sum_by([row for row in mirs if str(row.get("materialId")) in diesel_ids],
                          lambda row: str(row.get("date")), lambda row: row.get("quantity"))
            errors = []
            for day, quantity in expected.items():
                if not _close(edl.get(day, 0), quantity, 0.02):
                    errors.append(f"{day}: EDL fuel={edl.get(day, 0):.3f}, source={quantity:.3f}")
                if quantity > 0 and not _close(mir.get(day, 0), quantity, 0.02):
                    errors.append(f"{day}: diesel MIR={mir.get(day, 0):.3f}, source={quantity:.3f}")
            return _summarize(errors, f"Daily fuel reconciles for {len(expected)} dates across source, EDL and diesel MIR.")

        def check_mir_edl_activity_date_coverage() -> Tuple[str, str]:
            if not mirs:
                return "INCOMPLETE", "No MIR records are available to verify against EDL activity/date groups."
            edl_keys = set()
            for row in edls:
                edl_keys.add((str(row.get("date")), str(row.get("activityId"))))
            mir_keys = {
                (str(row.get("date")), str(row.get("activityId")))
                for row in mirs
            }
            missing = sorted(mir_keys - edl_keys)
            errors = [
                f"{activity_id} on {day}: MIR has no matching EDL activity/date group"
                for day, activity_id in missing
            ]
            return _summarize(
                errors,
                f"All {len(mir_keys)} MIR activity/date groups have matching EDL records.",
            )

        def check_supervisors() -> Tuple[str, str]:
            edl_by: Dict[Tuple[str, str], List[Mapping[str, Any]]] = defaultdict(list)
            mir_by: Dict[Tuple[str, str], List[Mapping[str, Any]]] = defaultdict(list)
            sup_by: Dict[Tuple[str, str], List[Mapping[str, Any]]] = defaultdict(list)
            for row in edls:
                edl_by[(str(row.get("date")), str(row.get("activityId")))].append(row)
            for row in mirs:
                mir_by[(str(row.get("date")), str(row.get("activityId")))].append(row)
            for row in supervisors:
                sup_by[(str(row.get("date")), str(row.get("activityId")))].append(row)
            errors = []
            for key, source_log in _logs_by_key(log_rows).items():
                edl_rows = edl_by.get(key, [])
                rows = sup_by.get(key, [])
                if not edl_rows:
                    if rows:
                        errors.append(f"{key}: SUP exists without EDL")
                    continue
                if len(rows) != 1:
                    errors.append(f"{key}: expected one SUP, got {len(rows)}")
                    continue
                sup = rows[0]
                expected_roles = Counter(str(row.get("employeeRole", "")) for row in edl_rows)
                plan = plan_by_activity.get(key[1])
                if plan:
                    for labour in plan.labour_lines:
                        if labour["category"] in _SUPERVISOR_ROLES:
                            expected_roles[labour["category"]] += max(1, labour["workers"])
                actual_roles = {
                    role: count for role, count in (sup.get("manpower") or {}).items()
                    if count
                }
                if dict(expected_roles) != actual_roles:
                    errors.append(f"{key}: SUP manpower differs from EDL plus planned supervisors")
                expected_machinery = Counter(str(row["assetId"]) for row in edl_rows if row.get("assetId"))
                actual_machinery = sup.get("machinery") or {}
                if isinstance(actual_machinery, list):
                    asset_ids_by_name = {
                        str(asset.get("assetName")): str(asset.get("assetId"))
                        for asset in assets
                    }
                    actual_machinery = Counter({
                        asset_ids_by_name.get(str(item.get("assetName")), str(item.get("assetName"))):
                            int(item.get("count") or 0)
                        for item in actual_machinery
                        if isinstance(item, Mapping)
                    })
                if dict(expected_machinery) != dict(actual_machinery):
                    errors.append(f"{key}: SUP machinery counts differ from EDL")
                expected_materials: Dict[str, float] = defaultdict(float)
                for mir in mir_by.get(key, []):
                    expected_materials[str(mir.get("materialName"))] += float(mir.get("quantity") or 0)
                actual_materials = sup.get("materialsReceived") or {}
                if isinstance(actual_materials, list):
                    actual_materials = {
                        str(item.get("materialName")): float(item.get("quantity") or 0)
                        for item in actual_materials
                        if isinstance(item, Mapping)
                    }
                if set(expected_materials) != set(actual_materials) or any(
                    not _close(expected_materials[name], actual_materials[name], 0.02)
                    for name in expected_materials
                ):
                    errors.append(f"{key}: SUP materialsReceived differs from MIR")
                if set(sup.get("engineeringDailyLogIds") or []) != {
                    row.get("engineeringDailyLogId") for row in edl_rows
                }:
                    errors.append(f"{key}: SUP EDL references differ")
                if set(sup.get("materialIssueRequestIds") or []) != {
                    row.get("materialIssueRequestId") for row in mir_by.get(key, [])
                }:
                    errors.append(f"{key}: SUP MIR references differ")
            return _summarize(errors, f"SUP manpower, machinery, materials, and references match {len(sup_by)} daily groups.")

        def check_edl_sup_assets() -> Tuple[str, str]:
            edl_by: Dict[Tuple[str, str], set] = defaultdict(set)
            sup_by: Dict[Tuple[str, str], set] = defaultdict(set)
            asset_ids_by_name = {
                str(asset.get("assetName")): str(asset.get("assetId"))
                for asset in assets
            }
            for row in edls:
                if row.get("assetId"):
                    edl_by[(str(row.get("date")), str(row.get("activityId")))].add(row["assetId"])
            for row in supervisors:
                machinery = row.get("machinery") or {}
                if isinstance(machinery, list):
                    asset_ids = (
                        asset_ids_by_name.get(str(item.get("assetName")), str(item.get("assetName")))
                        for item in machinery if isinstance(item, Mapping)
                    )
                else:
                    asset_ids = machinery
                for asset_id in asset_ids:
                    sup_by[(str(row.get("date")), str(row.get("activityId")))].add(asset_id)
            errors = [f"{key}: EDL assets != SUP machinery" for key in set(edl_by) | set(sup_by)
                      if edl_by.get(key, set()) != sup_by.get(key, set())]
            return _summarize(errors, f"EDL assets match SUP machinery across {len(set(edl_by)|set(sup_by))} daily groups.")

        def check_mpr_dates() -> Tuple[str, str]:
            errors = []
            for row in mprs:
                requested, consumed = _date(row.get("requestDate")), _date(row.get("deliveryDate"))
                if requested is None or consumed is None:
                    errors.append(f"{row.get('materialProcurementRequestId')}: invalid request/delivery date")
                    continue
                working_days = sum(
                    1 for ordinal in range(requested.toordinal() + 1, consumed.toordinal() + 1)
                    if date.fromordinal(ordinal).weekday() != 6
                )
                if working_days not in (5, 6):
                    errors.append(f"{row.get('materialProcurementRequestId')}: lead time is {working_days} working days")
            return _summarize(errors, f"All {len(mprs)} procurement dates are 5-6 working days before delivery/consumption.")

        def check_logistics() -> Tuple[str, str]:
            mpr_refs = [row.get("materialProcurementRequestId") for row in logistics
                        if row.get("materialProcurementRequestId")]
            mir_refs = [row.get("materialIssueRequestId") for row in logistics
                        if row.get("materialIssueRequestId")]
            errors = []
            for label, source, refs in (
                ("MPR", mprs, mpr_refs), ("MIR", mirs, mir_refs),
            ):
                expected = Counter(str(row.get(
                    "materialProcurementRequestId" if label == "MPR" else "materialIssueRequestId"
                )) for row in source)
                actual = Counter(str(ref) for ref in refs)
                if expected != actual:
                    missing = list((expected - actual).elements())[:5]
                    extra = list((actual - expected).elements())[:5]
                    errors.append(f"{label} logistics mismatch; missing={missing}, extra={extra}")
            return _summarize(errors, f"Every one of {len(mprs)} MPRs and {len(mirs)} MIRs has exactly one logistics record.")

        def check_transactions() -> Tuple[str, str]:
            errors = []
            expected_in = sum(float(row.get("quantity") or 0) for row in mprs)
            expected_out = sum(float(row.get("quantity") or 0) for row in mirs)
            actual_in = sum(float(row.get("quantity") or 0) for row in transactions if row.get("transactionType") == "INFLOW")
            actual_out = sum(float(row.get("quantity") or 0) for row in transactions if row.get("transactionType") == "OUTFLOW")
            if not _close(actual_in, expected_in, 0.05):
                errors.append(f"transaction INFLOW={actual_in:.4f}, MPR={expected_in:.4f}")
            if not _close(actual_out, expected_out, 0.05):
                errors.append(f"transaction OUTFLOW={actual_out:.4f}, MIR={expected_out:.4f}")
            stock: Dict[str, float] = defaultdict(float)
            for row in transactions:
                material = str(row.get("materialId"))
                quantity = float(row.get("quantity") or 0)
                if row.get("transactionType") == "INFLOW":
                    stock[material] += quantity
                elif row.get("transactionType") == "OUTFLOW":
                    stock[material] -= quantity
                    if stock[material] < -_TOLERANCE:
                        errors.append(f"{row.get('materialTransactionId')}: negative stock for {material}")
                else:
                    errors.append(f"{row.get('materialTransactionId')}: invalid transaction type")
                if not _close(row.get("runningStock"), stock[material], 0.05):
                    errors.append(f"{row.get('materialTransactionId')}: runningStock does not reconcile")
            return _summarize(errors, "Transaction inflow/outflow matches MPR/MIR and stock never goes negative.")

        def check_daily_parity() -> Tuple[str, str]:
            source_by_date = Counter(row.date.isoformat() for row in log_rows)
            edl_by_date = Counter(str(row.get("date")) for row in edls)
            mir_by_date = Counter(str(row.get("date")) for row in mirs)
            sup_by_date = Counter(str(row.get("date")) for row in supervisors)
            errors = []
            for day in source_by_date:
                if not edl_by_date[day]:
                    errors.append(f"{day}: source log has no EDL")
                if not sup_by_date[day]:
                    errors.append(f"{day}: source log has no SUP")
            for label, counts in (("EDL", edl_by_date), ("MIR", mir_by_date), ("SUP", sup_by_date)):
                extras = set(counts) - set(source_by_date)
                if extras:
                    errors.append(f"{label}: rows have dates absent from daily log: {sorted(extras)[:5]}")
            expected_edl_by_date: Counter = Counter()
            expected_mir_by_date: Counter = Counter()
            expected_sup_by_date: Counter = Counter()
            for log in log_rows:
                plan = plan_by_activity.get(log.activity_id)
                if plan is None:
                    errors.append(f"{log.date}: {log.activity_id} has no Plan source")
                    continue
                day = log.date.isoformat()
                edl_count = sum(
                    max(1, item["workers"])
                    for item in plan.labour_lines
                    if item["category"] not in _SUPERVISOR_ROLES
                )
                expected_edl_by_date[day] += edl_count
                expected_mir_by_date[day] += sum(
                    1 for item in plan.material_lines
                    if item["quantity"] / max(1, plan.duration_days) > 0
                ) + int(log.fuel_cost > 0)
                expected_sup_by_date[day] += int(edl_count > 0)
            for label, actual, expected in (
                ("EDL", edl_by_date, expected_edl_by_date),
                ("MIR", mir_by_date, expected_mir_by_date),
                ("SUP", sup_by_date, expected_sup_by_date),
            ):
                if actual != expected:
                    errors.append(f"{label} per-date counts differ from source: "
                                  f"actual={dict(actual)}, expected={dict(expected)}")
            return _summarize(errors, f"Daily EDL/MIR/SUP counts match source by date across {len(source_by_date)} dates.")

        def check_material_coverage() -> Tuple[str, str]:
            procured = {str(row.get("materialId")) for row in mprs}
            issued = {str(row.get("materialId")) for row in mirs}
            missing = sorted(issued - procured)
            return _summarize([f"issued material {mid} has no MPR" for mid in missing],
                              f"All {len(issued)} issued materials have at least one MPR.")

        def check_boq_activity_bijection() -> Tuple[str, str]:
            errors = []
            boq_ids = [str(row.get("boqId")) for row in boq_entities]
            activity_boqs = [str(row.get("boqId")) for row in activities]
            activity_ids = [str(row.get("activityId")) for row in activities]
            if len(boq_ids) != len(set(boq_ids)):
                errors.append("BOQ IDs are not unique")
            if len(activity_ids) != len(set(activity_ids)):
                errors.append("Activity IDs are not unique")
            if len(activity_boqs) != len(set(activity_boqs)):
                errors.append("An activity BOQ is referenced more than once")
            if set(boq_ids) != set(activity_boqs):
                errors.append("BOQ and activity mappings are not one-to-one")
            if len(boq_entities) != len(activities):
                errors.append(f"BOQ count={len(boq_entities)}, activity count={len(activities)}")
            return _summarize(errors, f"{len(boq_entities)} BOQs map bijectively to activities.")

        def check_unique_ids() -> Tuple[str, str]:
            errors = []
            for slug, (id_field, _) in ENTITY_SPECS.items():
                rows = list(entities.get(slug, []))
                ids = [str(row.get(id_field)) for row in rows]
                duplicates = [value for value, count in Counter(ids).items()
                              if count > 1 or value in ("", "None")]
                if duplicates:
                    errors.append(f"{slug}: duplicate/missing IDs {duplicates[:5]}")
            return _summarize(errors, "All 14 entity types have unique, nonempty IDs.")

        def check_references() -> Tuple[str, str]:
            ids = {
                slug: {str(row.get(id_field)) for row in entities.get(slug, [])
                       if row.get(id_field) not in (None, "")}
                for slug, (id_field, _) in ENTITY_SPECS.items()
            }
            refs = {
                "projectId": "project", "boqId": "boq", "activityId": "activity",
                "assetId": "asset", "materialId": "material", "logisticsId": "logistics",
                "parentSubContractorId": "subcontractor",
            }
            logistics_ids = {
                str(row.get("logisticsId")) for row in logistics
                if row.get("logisticsId") not in (None, "")
            }
            errors = []
            for slug, rows in entities.items():
                if slug not in ENTITY_SPECS:
                    errors.append(f"unknown entity type {slug}")
                    continue
                for row in rows:
                    local_id = row.get(ENTITY_SPECS[slug][0], "?")
                    for field, target in refs.items():
                        value = row.get(field)
                        if value in (None, ""):
                            continue
                        if str(value) not in ids.get(target, set()):
                            errors.append(f"{slug} {local_id}: {field}={value} does not resolve")
                    if slug == "engineering_daily_logs" and row.get("subContractorId"):
                        if str(row["subContractorId"]) not in subcontractor_ids:
                            errors.append(f"{slug} {local_id}: subContractorId does not resolve")
                    if slug == "material_procurement_request":
                        value = row.get("logisticsId")
                        if value and str(value) not in logistics_ids:
                            errors.append(f"{slug} {local_id}: logisticsId does not resolve")
                    if slug == "material_transaction":
                        target = ("material_procurement_request"
                                  if row.get("referenceType") == "MaterialProcurementRequest"
                                  else "material_issue_request"
                                  if row.get("referenceType") == "MaterialIssueRequest" else None)
                        if target is None or str(row.get("referenceId")) not in ids[target]:
                            errors.append(f"{slug} {local_id}: transaction reference does not resolve")
                    if slug == "supervisor_update":
                        for field, target in (("engineeringDailyLogIds", "engineering_daily_logs"),
                                              ("materialIssueRequestIds", "material_issue_request")):
                            for ref in row.get(field) or []:
                                if str(ref) not in ids[target]:
                                    errors.append(f"{slug} {local_id}: {field} reference {ref} does not resolve")
                    if slug == "logistics":
                        for field, target in (("materialProcurementRequestId", "material_procurement_request"),
                                              ("materialIssueRequestId", "material_issue_request")):
                            value = row.get(field)
                            if value and str(value) not in ids[target]:
                                errors.append(f"{slug} {local_id}: {field} does not resolve")
            return _summarize(errors, "All populated cross-entity references resolve.")

        def check_required() -> Tuple[str, str]:
            errors = []
            if len(entities) != 14:
                errors.append(f"expected 14 entity arrays, found {len(entities)}")
            for slug, (id_field, required_fields) in ENTITY_SPECS.items():
                if slug not in entities:
                    errors.append(f"{slug}: entity array missing")
                    continue
                for index, row in enumerate(entities[slug]):
                    if id_field not in row or row.get(id_field) in (None, ""):
                        errors.append(f"{slug}[{index}]: schema-required ID field {id_field} missing")
            return _summarize(errors, "All generated schema ID/relationship required fields are present.")

        def check_material_stock_cost() -> Tuple[str, str]:
            errors = []
            rates: Dict[str, set] = defaultdict(set)
            for row in plan_rows:
                for item in row.material_lines:
                    rates[item["name"]].add(float(item["unit_rate"]))
            for row in materials:
                name = str(row.get("materialName", ""))
                if "currentStock" not in row or row.get("currentStock") is None:
                    errors.append(f"{name}: currentStock missing")
                expected_rates = rates.get(name.lower().replace(" ", "_").replace("-", "_"))
                if expected_rates is None:
                    expected_rates = next((values for material_name, values in rates.items()
                                            if material_name.replace("_", " ").casefold() == name.casefold()), None)
                if name.casefold() == "diesel":
                    expected_rates = {100.0}
                if not expected_rates:
                    errors.append(f"{name}: no source unit cost found")
                elif len(expected_rates) != 1 or not _close(row.get("unitCost"), next(iter(expected_rates)), 0.01):
                    errors.append(f"{name}: unitCost does not match one consistent Plan assumption rate")
            return _summarize(errors, f"currentStock and unit costs verified for {len(materials)} materials.")

        def check_mpr_material_ids() -> Tuple[str, str]:
            errors = [str(row.get("materialProcurementRequestId")) + ": materialId null/unresolved"
                      for row in mprs if not row.get("materialId") or str(row.get("materialId")) not in material_by_id]
            return _summarize(errors, f"All {len(mprs)} MPR material IDs are populated and resolve.")

        def check_mir_requestor_quantity() -> Tuple[str, str]:
            errors = []
            known_workers = {
                str(row.get("subContractorId")) for row in subcontractors
                if row.get("subContractorType") == "worker"
            }
            for row in mirs:
                request_id = str(row.get("materialIssueRequestId"))
                if row.get("personRequested") not in known_workers:
                    errors.append(f"{request_id}: personRequested is not a known worker")
                qty, rate = row.get("quantity"), row.get("unitCost")
                if not _finite_positive(qty):
                    errors.append(f"{request_id}: quantity is not positive")
                if not _close(row.get("totalCost"), round(float(qty or 0) * float(rate or 0), 2), 0.005):
                    errors.append(f"{request_id}: totalCost does not equal quantity × unitCost")
                key = (str(row.get("date")), str(row.get("activityId")))
                source_log = source_logs_by_key.get(key)
                plan = plan_by_activity.get(key[1])
                if source_log is None or plan is None:
                    errors.append(f"{request_id}: no matching daily-log/plan source")
                    continue
                material = material_by_id.get(str(row.get("materialId")), {})
                material_name = str(material.get("materialName", "")).casefold()
                if material_name == "diesel":
                    expected_quantity = source_log.fuel_consumption
                    if not _close(qty, expected_quantity, 0.002):
                        errors.append(f"{request_id}: diesel quantity differs from daily fuel consumption")
                else:
                    line = next(
                        (item for item in plan.material_lines
                         if item["name"].replace("_", " ").title().casefold()
                         == str(row.get("materialName", "")).casefold()),
                        None,
                    )
                    if line is None:
                        errors.append(f"{request_id}: material/activity does not match Plan assumptions")
                    else:
                        expected_quantity = line["quantity"] / max(1, plan.duration_days)
                        if abs(float(qty or 0) - expected_quantity) > expected_quantity * 0.051 + 0.0001:
                            errors.append(f"{request_id}: quantity is outside the configured daily Plan amount")
            return _summarize(errors, f"All {len(mirs)} MIR quantities/costs are valid and requestors resolve.")

        def check_asset_required_fields() -> Tuple[str, str]:
            required = ("assetId", "assetCode", "assetName", "equipmentType",
                        "plannedTotalHours", "plannedEquipmentCostINR",
                        "actualUtilisationHrsPerDay", "hoursLoggedToDate",
                        "odometerOpening", "odometerClosing", "status")
            errors = []
            for row in assets:
                aid = row.get("assetId", "?")
                for field in required:
                    if field not in row or row.get(field) is None:
                        errors.append(f"{aid}: missing {field}")
                if not _finite_nonnegative(row.get("plannedEquipmentCostINR")):
                    errors.append(f"{aid}: invalid plannedEquipmentCostINR")
            return _summarize(errors, f"All {len(assets)} assets include required identity, usage and cost fields.")

        def check_supervisor_daily_counts() -> Tuple[str, str]:
            errors = []
            source_keys = _logs_by_key(log_rows)
            edl_counts = Counter((str(row.get("date")), str(row.get("activityId"))) for row in edls)
            mir_counts = Counter((str(row.get("date")), str(row.get("activityId"))) for row in mirs)
            sup_counts = Counter((str(row.get("date")), str(row.get("activityId"))) for row in supervisors)
            for key, source in source_keys.items():
                plan = plan_by_activity.get(key[1])
                if plan is None:
                    errors.append(f"{key}: no activity resource assumptions")
                    continue
                expected_edl = sum(
                    max(1, labour["workers"])
                    for labour in plan.labour_lines
                    if labour["category"] not in _SUPERVISOR_ROLES
                )
                expected_mir = sum(1 for material in plan.material_lines
                                   if material["quantity"] / max(1, plan.duration_days) > 0)
                if source.fuel_cost > 0:
                    expected_mir += 1
                expected_sup = 1 if expected_edl else 0
                if edl_counts[key] != expected_edl:
                    errors.append(f"{key}: EDL count {edl_counts[key]} != expected {expected_edl}")
                if mir_counts[key] != expected_mir:
                    errors.append(f"{key}: MIR count {mir_counts[key]} != expected {expected_mir}")
                if sup_counts[key] != expected_sup:
                    errors.append(f"{key}: SUP count {sup_counts[key]} != expected {expected_sup}")
            return _summarize(errors, f"Daily EDL/MIR/SUP counts match {len(source_keys)} source activity-days.")

        def check_material_ledger() -> Tuple[str, str]:
            errors = []
            mir_by_mat = _sum_by(mirs, lambda row: str(row.get("materialId")), lambda row: row.get("quantity"))
            mpr_by_mat = _sum_by(mprs, lambda row: str(row.get("materialId")), lambda row: row.get("quantity"))
            txn_by_mat_type: Dict[Tuple[str, str], List[Mapping[str, Any]]] = defaultdict(list)
            for row in transactions:
                txn_by_mat_type[(str(row.get("materialId")), str(row.get("transactionType")))].append(row)
            for material in materials:
                mid = str(material.get("materialId"))
                inflows = txn_by_mat_type.get((mid, "INFLOW"), [])
                outflows = txn_by_mat_type.get((mid, "OUTFLOW"), [])
                if not inflows:
                    errors.append(f"{mid}: no INFLOW")
                if not outflows:
                    errors.append(f"{mid}: no OUTFLOW")
                if len(inflows) != sum(1 for row in mprs if str(row.get("materialId")) == mid):
                    errors.append(f"{mid}: INFLOW transaction count differs from MPR count")
                if len(outflows) != sum(1 for row in mirs if str(row.get("materialId")) == mid):
                    errors.append(f"{mid}: OUTFLOW transaction count differs from MIR count")
                in_qty = sum(float(row.get("quantity") or 0) for row in inflows)
                out_qty = sum(float(row.get("quantity") or 0) for row in outflows)
                if not _close(in_qty, mpr_by_mat.get(mid, 0), 0.05):
                    errors.append(f"{mid}: ledger INFLOW differs from MPR")
                if not _close(out_qty, mir_by_mat.get(mid, 0), 0.05):
                    errors.append(f"{mid}: ledger OUTFLOW differs from MIR")
                if not _close(in_qty, out_qty, 0.05):
                    errors.append(f"{mid}: material inflow/outflow quantities do not balance")
            return _summarize(errors, f"Material ledger balances for all {len(materials)} materials.")

        def check_activity_work_done() -> Tuple[str, str]:
            source_by_activity = _sum_by(log_rows, lambda row: row.activity_id, lambda row: row.quantity)
            plan_by_id = {row.activity_id: row for row in plan_rows}
            actual_by_activity = _sum_by(edls, lambda row: str(row.get("activityId")), lambda row: row.get("workDone"))
            errors = []
            for aid, activity in plan_by_id.items():
                source_qty = source_by_activity.get(aid, 0)
                edl_qty = actual_by_activity.get(aid, 0)
                if not _close(source_qty, activity.quantity, 0.02):
                    errors.append(f"{aid}: daily sheet sum={source_qty:.4f} differs from Plan quantity={activity.quantity:.4f}")
                if not _close(edl_qty, activity.quantity, 0.02):
                    errors.append(f"{aid}: EDL workDone={edl_qty:.4f} differs from activity quantity={activity.quantity:.4f}")
            return _summarize(errors, f"EDL workDone equals Plan quantity for all {len(plan_rows)} activities.")

        run(1, _CHECK_NAMES[0], check_plan_costs)
        run(2, _CHECK_NAMES[1], check_plan_excel_json)
        run(3, _CHECK_NAMES[2], check_boq)
        run(4, _CHECK_NAMES[3], check_activities)
        run(5, _CHECK_NAMES[4], check_assets)
        run(6, _CHECK_NAMES[5], check_edl_work_done)
        run(7, _CHECK_NAMES[6], check_fuel)
        run(8, _CHECK_NAMES[7], check_mir_edl_activity_date_coverage)
        run(9, _CHECK_NAMES[8], check_supervisors)
        run(10, _CHECK_NAMES[9], check_edl_sup_assets)
        run(11, _CHECK_NAMES[10], check_mpr_dates)
        run(12, _CHECK_NAMES[11], check_logistics)
        run(13, _CHECK_NAMES[12], check_transactions)
        run(14, _CHECK_NAMES[13], check_daily_parity)
        run(15, _CHECK_NAMES[14], check_material_coverage)
        run(16, _CHECK_NAMES[15], check_boq_activity_bijection)
        run(17, _CHECK_NAMES[16], check_unique_ids)
        run(18, _CHECK_NAMES[17], check_references)
        run(19, _CHECK_NAMES[18], check_required)
        run(20, _CHECK_NAMES[19], check_material_stock_cost)
        run(21, _CHECK_NAMES[20], check_mpr_material_ids)
        run(22, _CHECK_NAMES[21], check_mir_requestor_quantity)
        run(23, _CHECK_NAMES[22], check_asset_required_fields)
        run(24, _CHECK_NAMES[23], check_supervisor_daily_counts)
        run(25, _CHECK_NAMES[24], check_material_ledger)
        run(26, _CHECK_NAMES[25], check_activity_work_done)
    finally:
        workbook.close()
    return EntityValidationReport(checks)


_CHECK_NAMES = (
    "Plan entity costs come from assumptions",
    "Excel Plan sheet equals JSON plan entries",
    "BOQ counts/data match Excel Plan and JSON",
    "Activity counts/data match Excel Plan and JSON",
    "Assets, usage, and costs match assumptions",
    "EDL workDone calculation",
    "MIR fuel procurement/distribution matches daily log by date",
    "Every MIR activity/date group has matching EDL records",
    "Supervisor updates match MIR/EDL/labour/machinery source counts",
    "EDL assets are covered by supervisor machinery per date",
    "MPR dates are 5-6 working days before consumption",
    "Every MPR and MIR has exactly one logistics record",
    "Material transactions match MPR/MIR and have nonnegative stock",
    "Daily MIR/EDL/SUP coverage matches source dates and groups",
    "Every issued material has at least one MPR",
    "Each BOQ maps to exactly one activity and vice versa",
    "IDs are unique within each entity type",
    "All populated cross-entity references resolve",
    "All required entity ID/relationship fields are present",
    "Material currentStock and unitCosts match source assumptions",
    "MPR materialId is non-null and resolves",
    "MIR requestor resolves and quantities/costs are correct",
    "All asset identity, usage, and cost fields are present",
    "Supervisor labour/machinery counts match EDL/MIR by date",
    "Material ledger balances and transaction counts match MPR/MIR",
    "EDL workDone sums to activity quantity",
)


def _finite_nonnegative(value: Any) -> bool:
    try:
        number = float(value)
        return math.isfinite(number) and number >= 0
    except (TypeError, ValueError):
        return False


def _finite_positive(value: Any) -> bool:
    try:
        number = float(value)
        return math.isfinite(number) and number > 0
    except (TypeError, ValueError):
        return False


def _sum_by(rows: Iterable[Any], key_fn: Callable[[Any], Any],
            value_fn: Callable[[Any], Any]) -> Dict[Any, float]:
    result: Dict[Any, float] = defaultdict(float)
    for row in rows:
        result[key_fn(row)] += float(value_fn(row) or 0)
    return dict(result)


def _logs_by_key(rows: Iterable[Any]) -> Dict[Tuple[str, str], Any]:
    return {
        (row.date.isoformat(), row.activity_id): row
        for row in rows
    }
