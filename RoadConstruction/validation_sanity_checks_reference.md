# Validation and Sanity Checks Reference

This document lists the checks applied to generated schedules and exported
Excel workbooks. The implementation is in
[`src/validation_sanity_checks.py`](src/validation_sanity_checks.py), and the
results are written to `15_Validation_Checks` for new workbooks and
`09_Validation_Checks` for updated workbooks.

## How to read the result

| Result | Meaning |
|---|---|
| `PASS` | The check completed successfully. |
| `WARNING` | The calculation is internally usable, but an assumption, planning risk, or engineering decision requires review. |
| `ERROR` | A hard reconciliation or integrity failure exists. The workbook must not be accepted until corrected. |

The validation report is not an engineering design approval. A passing
calculation does not verify drawings, quantities, rates, site conditions,
contract clauses, or statutory approvals.

## 1. Checks shared by Road, Pipeline, and Building

These checks apply to all three project types.

### Schedule and activity integrity

| Code | Check |
|---|---|
| `NO_ACTIVITIES` | The schedule contains at least one activity. |
| `ACTIVITY_MASTER` | Scheduled activity keys exist in the selected activity catalogue. |
| `UNKNOWN_ACTIVITY` | No scheduled activity exists outside the activity master. |
| `ZERO_ACTIVE_QUANTITY` | Active activities have positive quantities. Catalogue-only zero rows may be warnings for road/building; pipeline zero active quantities are errors. |
| `MISSING_PREDECESSOR` | Every predecessor ID exists in the active schedule. |
| `CIRCULAR_DEPENDENCY` | The predecessor network contains no cycle. |
| `PRECEDENCE_VIOLATION` | Each activity starts after its predecessor and lag requirement. |
| `INVALID_ACTIVITY_DATES` | Activity dates are valid and ordered. |
| `INVALID_CREW_COUNT` | Each active activity has a valid crew count. |
| `INVALID_SHIFT_COUNT` | Shift allocation is valid. |
| `CREW_CAP_EXCEEDED` | Configured crew/resource caps are not exceeded. |

### Quantity, productivity, duration, and CPM

| Code | Check |
|---|---|
| `NEGATIVE_QUANTITY` | Quantities cannot be negative. |
| `INVALID_PRODUCTIVITY` | Productivity must be positive. |
| `INVALID_DURATION` | A non-zero activity must have at least one day. |
| `PRODUCTIVITY_DURATION_MISMATCH` | Effort equals `ceil(quantity / productivity / crews / effective shifts)`. |
| `CALENDAR_DURATION_MISMATCH` | Calendar-adjusted duration matches working-day, holiday, monsoon, terrain, and shift rules. |
| `CPM_EF_IDENTITY` | `EF = ES + duration`. |
| `CPM_LF_IDENTITY` | `LF = LS + duration`. |
| `CPM_FLOAT_IDENTITY` | `Total float = LS - ES`. |
| `NEGATIVE_FLOAT` | CPM float values are not negative. |
| `PROJECT_DURATION_RECONCILIATION` | Project finish matches the activity network. |
| `FEASIBILITY_DURATION_MISMATCH` | Optimized duration satisfies the requested duration ceiling. |
| `FEASIBILITY_BUDGET_MISMATCH` | Optimized cost satisfies the budget ceiling. |

### Rate, cost, wastage, and escalation

| Code | Check |
|---|---|
| `ASSUMPTION_SOURCE_MISSING` | Scheduled values resolve to an activity assumption. |
| `ASSUMPTION_RATE_MISSING` | Every scheduled activity has a configured rate. |
| `ASSUMPTION_RATE_INVALID` | Configured rates are numeric and positive. |
| `RATE_METADATA_MISSING` | Rate source metadata is present. |
| `RATE_NOT_VERIFIED` | Rate metadata is still marked unverified/provisional. This is a review warning, not an arithmetic failure. |
| `INVALID_PLANNED_RATE` | Planned unit rate is positive. |
| `UNIT_RATE_BUILDUP_MISMATCH` | Planned rate matches the configured source/resource/on-cost build-up. |
| `DIRECT_COST_RECONCILIATION` | Direct cost matches quantity × planned rate + wastage/surcharge treatment. |
| `TOTAL_COST_RECONCILIATION` | Total cost equals direct cost + configured overhead/profit/escalation treatment. |
| `NEGATIVE_COST` | Cost values are not negative. |
| `EMBEDDED_ONCOST_DUPLICATION` | Overhead and profit are not charged twice. |
| `WASTAGE_RECONCILIATION` | Total wastage equals material, fuel, plant, and labour loss components. |
| `ESCALATION_OUTPUT_INVALID` | Escalation factor is at least 1 and escalation amount is non-negative. |
| `PROVISIONAL_ESCALATION` | Escalation source or contract clause remains provisional. |

### Plan, resource, and progress reconciliation

| Code | Check |
|---|---|
| `PLAN_DETAIL_TOTAL_MISMATCH` | Plan detail rows equal activity summary values. |
| `MATERIAL_DETAIL_UNAVAILABLE` | A material-cost allocation has no physical material quantity/recipe detail; it is reported as an unitemized cost balance and requires review. |
| `PLAN_ACTIVITY_SET_MISMATCH` | Plan activity rows match the calculated activity set. |
| `PLAN_PROJECT_EXECUTION_MISMATCH` | Plan execution total equals activity subtotals. |
| `PLAN_PROJECT_BILLED_MISMATCH` | Plan billed total equals activity subtotals. |
| `MONTHLY_COST_RECONCILIATION` | Monthly costs equal project cost. |
| `PROGRESS_NOT_COMPLETE` | Monthly progress reaches 100%. |
| `RESOURCE_PLAN_DAILY_LOG_MISMATCH` | Resource plan and Engineer_DailyLog values reconcile. |
| `SCHEDULER_SANITY` | Planning-risk heuristics such as Bromilow duration and monsoon exposure are reported. |

### Exported workbook checks

| Code | Check |
|---|---|
| `WORKBOOK_ACTIVITY_SET_MISMATCH` | Workbook activity set matches the in-memory schedule. |
| `WORKBOOK_SCURVE_PROGRESS_MISMATCH` | Workbook S-curve progress reconciles. |
| `WORKBOOK_SCURVE_COST_MISMATCH` | Workbook S-curve costs reconcile. |
| `WORKBOOK_PLAN_INPUT_MISMATCH` | Workbook Plan inputs match the schedule result. |
| `WORKBOOK_PLAN_PROJECT_EXECUTION_MISMATCH` | Exported Plan execution total reconciles. |
| `WORKBOOK_PLAN_PROJECT_BILLED_MISMATCH` | Exported Plan billed total reconciles. |
| `WORKBOOK_SHEETS_MISSING` | Required workbook sheets exist. |

## 2. Road-specific checks

Road schedules use the shared checks plus road-specific geometry and planning
rules:

- Lane-dependent widths are loaded from `lane_standards`.
- Quantities for formation, excavation, embankment, subgrade, GSB, WMM,
  pavement layers, drainage, and road furniture are checked through the
  quantity and cost reconciliation checks.
- Optional activities are checked against declared scope:
  service roads, toll plazas, ATMS, bridges, flyovers, underpasses, and
  retaining walls must not appear as active scope when their input is zero.
- Hilly schedules must use the hilly catalogue from road guidelines and must
  not silently copy plain-terrain productivity.
- Road rate source metadata is checked through
  `ASSUMPTION_RATE_MISSING`, `RATE_METADATA_MISSING`, and
  `RATE_NOT_VERIFIED`.
- Parallel critical-chain and monsoon exposure messages are reported through
  `SCHEDULER_SANITY`.

The independent road audit can also be run with:

```powershell
& ".\.venv-pipeline\Scripts\python.exe" -m src.audit_calculations
```

That audit independently recalculates representative road quantities, rates,
CPM identities, wastage, escalation, and totals.

## 3. Pipeline-specific checks

Pipeline schedules use the shared checks plus:

| Code | Pipeline-specific rule |
|---|---|
| `SOIL_SHARES_NOT_100` | Ordinary soil, moorum, soft rock, and hard rock shares total 100%. |
| `QUANTITY_FORMULA_MISMATCH` | Excavation quantities match length × trench width × depth × strata share. |
| `EXCAVATION_PARTITION_MISMATCH` | Soil-strata quantities partition the total trench volume. |
| `RATE_SOURCE_MISSING` | A source direct rate exists for the activity. |
| `BACKFILL_BEFORE_LAYING` | Backfill/test work follows pipe laying. |
| `COMMISSIONING_BEFORE_TEST` | Commissioning follows the required test/backfill sequence. |
| `ENGINEERING_HANDOVER_REVIEW` | Excavation/bedding-to-laying handover is explicit or requires engineering review. |

Pipeline planned unit rates come from `direct_unit_rates_inr` in the
assumptions file. The configured statutory overhead and contractor-profit
factor is applied to that source rate; the activity's planned direct amount is
then based on activity quantity × planned unit rate, with configured wastage
and acceleration costs treated separately.

The source rate may include material cost even when no physical material recipe
is configured for an activity. In that case the estimate still uses the
activity quantity and planned unit rate, but the Plan sheet does not present
the allocated amount as a material cost with quantity `0`. It places the
unitemized amount under `Other Cost` and emits
`MATERIAL_DETAIL_UNAVAILABLE`. Add a physical material recipe and quantities
to obtain auditable material quantities and costs.

Pipeline inputs reviewed include:

- Pipe diameter, material, length, trench width/depth
- Work fronts and installation method
- Road, highway, railway, HDD, and river crossings
- Villages, OHTs, WTP, IBPS, and FHTC scope where declared

Crossing quantities must come from declared crossing inputs. A passing
quantity check confirms arithmetic; it does not confirm that the route survey
identified every real crossing.

## 4. Building-specific checks

Building schedules use the shared checks plus building scope and catalogue
controls:

- Foundation selection is limited to `pile`, `raft`, or
  `isolated_footing`.
- Only activities associated with the selected foundation are allowed to be
  active.
- Structural activities must use explicit floor-wise quantities when the
  floor-wise option is selected.
- Structural quantities remain planning assumptions until drawings and design
  take-offs are provided.
- MEP quantities are planning allowances unless design take-offs are loaded.
- A zero-quantity catalogue row may remain visible for audit/catalogue
  visibility, producing `ZERO_ACTIVE_QUANTITY` as a warning rather than a
  cost calculation.
- Building rates and escalation remain subject to source and contract review;
  passing arithmetic does not make an indicative rate tender-verified.

## 5. Update-plan checks

Updated workbooks use `validate_update_result()` and add:

| Code | Check |
|---|---|
| `UPDATE_NO_REVISED_ACTIVITIES` | Status rows cannot exist without revised activities when work remains. |
| `UPDATE_ACTIVITY_STATUS_SET_MISMATCH` | Status rows match the baseline activity set. |
| `UPDATE_REVISED_ACTIVITY_SET_MISMATCH` | Every revised activity has a status row. |
| `UPDATE_PROGRESS_OUT_OF_RANGE` | Actual activity progress is 0–100%. |
| `UPDATE_QUANTITY_RECONCILIATION` | Completed + remaining quantity equals original quantity. |
| `UPDATE_CPM_EF_IDENTITY` | Revised `EF = ES + duration`. |
| `UPDATE_CPM_LF_IDENTITY` | Revised `LF = LS + duration`. |
| `UPDATE_CPM_FLOAT_IDENTITY` | Revised float equals `LS - ES`. |
| `UPDATE_MISSING_PREDECESSOR` | Revised predecessors exist. |
| `UPDATE_PRECEDENCE_VIOLATION` | Revised dates respect predecessor dates and lags. |
| `UPDATE_DIRECT_COST_RECONCILIATION` | Revised direct cost reconciles. |
| `UPDATE_TOTAL_COST_RECONCILIATION` | Revised total cost reconciles. |
| `UPDATE_WASTAGE_RECONCILIATION` | Revised wastage components reconcile. |
| `UPDATE_ESCALATION_OUTPUT_INVALID` | Revised escalation values are valid. |
| `UPDATE_DURATION_RECONCILIATION` | Revised duration equals the latest activity finish. |
| `UPDATE_REMAINING_COST_RECONCILIATION` | Remaining cost equals revised activity totals. |
| `UPDATE_PROJECT_COST_RECONCILIATION` | Spent + remaining equals revised project cost. |
| `UPDATE_SCURVE_COST_RECONCILIATION` | Revised S-curve expenditure equals remaining cost. |
| `UPDATE_COMPLETION_DATE_MISSING` | Revised completion date is present. |
| `UPDATE_NEGATIVE_DELAY` | Projected delay is not negative. |
| `UPDATE_PLANNING_WARNING` | Planner-generated delay or feasibility warning is preserved. |

## 6. Important limitation

The checks prove internal consistency and assumption traceability. They cannot
prove the engineering correctness of an assumption that is itself wrong or
unapproved. Human review is still required for drawings, quantities, route
surveys, geotechnical conditions, rates, escalation clauses, scope, and
construction feasibility.
