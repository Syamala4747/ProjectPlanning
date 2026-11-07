# How the Construction Planning System Works

This application turns project details into an Excel construction plan. It can
plan roads, pipelines, and buildings, update an existing plan, and—when
requested separately—send the plan's records to a database.

For setup instructions and copy-ready commands, see
[GETTING_STARTED.md](GETTING_STARTED.md).

## The whole process in one view

**Project details → schedule and cost planning → Excel workbook → optional
progress update → entity records → validation → one database route**

Creating or updating an Excel workbook does **not** connect to a database.
Database upload is a separate action. The Excel workbook itself is never sent
to the database; the uploader reads information from it and sends the records
it creates from that information.

## 1. Project details and configuration

The user describes the project. The details depend on the kind of work:

- **Road:** title, length, lanes, terrain, target duration, budget, and start
  date.
- **Pipeline:** title, length, pipe size and material, installation method,
  work fronts, target duration, budget, and start date.
- **Building:** title, built-up area, floors, building and construction type,
  quality, target duration, budget, and start date.

The application selects the matching project configuration. Each project type
has guideline data and assumptions, such as activity catalogs, production
rates, resource requirements, calendars, and costs. The application checks
these inputs before planning. Missing or invalid required configuration stops
workbook generation; configuration warnings are shown for the user to review.

## 2. How the plan is calculated

The planning engine turns the project details and configuration into work:

1. It calculates quantities and selects the activities that apply to the
   project.
2. It estimates activity durations and costs from quantities, productivity,
   crews, equipment, and other assumptions.
3. It schedules the activities according to their dependencies using the
   Critical Path Method (CPM).
4. It compares the baseline schedule with the user's duration and budget
   limits. When needed, it evaluates resource increases, work fronts, shifts,
   or overlapping work to reduce the duration while accounting for added
   costs.
5. It calculates the final schedule, critical path, costs, work breakdown,
   monthly progress, and planned progress curve.

The result may be feasible or infeasible. An infeasible result explains which
target could not be met; it is not silently presented as meeting the target.

## 3. The Excel workbook

Before saving, the application checks the schedule and dependencies. It then
creates the project's Excel workbook with the summary, inputs, assumptions,
activities, baseline and optimized schedules and costs, work breakdown,
monthly schedule, progress curve, and critical-path details.

Workbook generation ends here. It creates the Excel file only: it does not
create the database JSON, run the entity checks, or connect to PostgreSQL or
the BluConn API.

## 4. Updating a plan

An update starts from an existing workbook; it does not replace the baseline.
The application reads the original project details and reconstructs its
baseline using the applicable configuration. It compares reported progress
at a status date with planned progress, estimates which work remains, and
recalculates the schedule and feasibility for the target date.

If activity-by-activity progress is not supplied, the application estimates
activity progress from the overall reported progress and notes that limitation.
The result is saved as a separate updated Excel workbook. Updating a plan does
not upload it to the database.

## 5. Turning the workbook into database records

When the user chooses to push a workbook, the Python entity builder reads its
plan, daily-log, input, and assumption sheets. It creates structured records
for projects, activities, plans, materials, equipment, daily logs, requests,
transactions, supervisor updates, and related entity types.

The builder adds and conforms fields using the entity-type reference file.
It writes the generated records as a JSON file beside the workbook. This JSON
is an intermediate representation for the upload process; it is not the Excel
file and is not itself proof that records reached the database.

The entity builder creates records from workbook data and the configured
entity schema. Generated daily logs and related progress records are based on
the plan and represent simulated activity, not verified field work. For the
BluConn API route, worker-only engineering logs that have no asset are mapped
to a separate `labour-daily-logs` entity type; equipment/operator logs retain
their Asset relationship. Confirm that the receiving application supports this
additional entity type before relying on its UI to display worker logs.

## 6. Validation before upload

The entity validator compares the generated records with the source workbook.
Its checks cover counts, costs, activity and BOQ relationships, daily work,
materials, equipment, requests, transactions, dates, and cross-references.
It produces an Excel validation report beside the workbook.

- **PASS:** the checks succeeded.
- **FAIL:** the upload stops until the reported problems are addressed.
- **INCOMPLETE:** some evidence could not be checked; the report explains why.

Validation checks the generated records against the workbook. It does not
confirm database connectivity or prove that a later database/API request was
successful.

## 7. Choose one of the two database routes

Both routes start with entity records, but they reach the database differently.
Choose one route for a project; do not send the same project through both.

### Route A: Python connects directly to PostgreSQL

The Python uploader can take the workbook or an already-generated JSON file.
For a workbook, it builds the JSON, runs validation, and proceeds only when
validation allows the push. It then reads the database connection details from
the local `.env` configuration and connects directly from the user's computer
to PostgreSQL using a database driver.

It prepares the entity types and their relationships, then inserts records in
parent-before-child order so each record can be linked to its parent. It checks
for matching project and entity IDs already in the database and skips those
records rather than inserting them again. Successful work is committed as a
database transaction; an insertion error causes the transaction to roll back.
The uploader writes a database-push Excel report with sent, inserted,
already-existing, duplicate, and failure counts.

For an as-of progress snapshot, push a workbook with `--progress` and
`--status-date`, for example:

```powershell
python scripts/push_to_db.py --xlsx "output\north.xlsx" --project-type road --progress 30 --status-date 2027-02-12
```

The status date is inclusive. This writes a separate snapshot JSON, records
the project-level actual progress and status date, and omits generated
daily-log, supervisor, material, logistics, and transaction records after that
date. The full planned schedule remains included; activity-level completion is
not inferred from an overall percentage. Existing project progress is updated
when the project is already in the database, but future records from an earlier
full push are not deleted. Generated daily-log data is simulated and is not
confirmation of field activity.

Because this route connects directly to PostgreSQL, the database host and port
must be reachable from the computer running the Python script. A Docker
service name is normally visible only inside Docker's network; a script running
on Windows needs the database's host-published address and port.

### Route B: Python sends requests to the BluConn API

The separate `scripts/push_to_api.py` command reads the same workbook or
entity JSON, reuses the existing entity builder and source-backed workbook
validation, and uploads through BluConn's account-scoped API. Configure
`BLUCONN_API_URL`, `BLUCONN_ACCOUNT_ID`, and `BLUCONN_AUTH_TOKEN` in the
environment or local `.env` file. The configured bearer token is sent in the
Authorization header. This route does not import the PostgreSQL uploader, use
`DATABASE_URL`, or connect directly to PostgreSQL.

Example:

```powershell
python scripts/push_to_api.py --xlsx "output\south_road.xlsx" --project-type road
```

An existing JSON file can also be uploaded; provide its source workbook for
validation:

```powershell
python scripts/push_to_api.py --json "output\south_road_actual_data.json" --source-xlsx "output\south_road.xlsx"
```

The API route creates or reuses entity types and type relationships, checks
existing business IDs (scoped by project), and uploads in parent-before-child
order. Child records are batched by their resolved API parent ID. Authentication
or authorization failures stop the request; retryable reads and rate-limited
writes are retried, while ambiguous server/network failures are reported
rather than blindly resubmitted. Each run writes an
`*_api_upload_report.xlsx` with uploaded, skipped, duplicate, and failed counts
plus per-record details where the API provides them. Date-only values are sent
as midnight UTC timestamps when required by the API schema. Equipment/operator
Engineering Daily Logs remain children of their Asset. Generated worker-only
logs without an asset are uploaded to a separate parentless `labour-daily-logs`
entity type, so the script does not invent an equipment relationship or require
changing the existing BluConn Engineering Daily Logs type. The receiving
application must support this additional entity type to display those worker
logs in its UI. The Project `terrain` field is populated only for road projects;
pipeline and building project records omit it because terrain is not applicable.

The direct PostgreSQL and API routes use separate configuration and commands.
The API route does not delete or update previously uploaded records; matching
business IDs are skipped. The BluConn backend handles persisting accepted
records.

API credentials and database credentials are different. The API uploader
sends the configured token as a bearer token; renew it when it expires. Never
put credentials in command-line arguments, source files, or shared reports.

## 8. How to tell what happened

There are three different outcomes to distinguish:

1. **Workbook created:** the planning or update step completed.
2. **Records prepared and validated:** entity JSON was created and the
   validation report shows its result.
3. **Records reached the database:** the selected uploader connected to its
   destination and reported successful insert/API counts.

A workbook or a passing validation report alone does not mean data was
uploaded. A failed or interrupted write can have an ambiguous outcome if the
server accepted it before the connection closed; reconcile the destination
before retrying. After a successful push, review the push summary/report and
confirm the project in the BluConn application.

## Main parts of the application

- **Project configuration:** project-specific guidelines and assumptions for
  roads, pipelines, and buildings.
- **Planning engine:** quantity, cost, duration, calendar, resource, and CPM
  calculations.
- **Excel reporting:** creates baseline and updated project workbooks.
- **Entity builder and validator:** converts workbook content into structured
  records and checks them against the workbook.
- **Python database uploader:** connects directly to PostgreSQL.
- **BluConn API uploader:** sends authenticated requests to the BluConn API.
