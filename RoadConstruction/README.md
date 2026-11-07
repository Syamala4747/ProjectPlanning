# How the Construction Planning System Works

This application turns project details into an Excel construction plan. It can
plan roads, pipelines, and buildings, update an existing plan, and—when
requested separately—send the plan's records to a database.

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

The standard builder creates records for the original 14 entity types. The
one-command BluConn API flow also adds six **simulated demo types**:
stores, store inventory, material requests, store transactions, daily logs,
and daily-log summaries. Store and inventory records are derived from supported
materials and the generated procurement/issue requests. Daily logs and
summaries are based on the workbook's planned daily schedule. These rows are
marked as simulated and are not actual field activity or confirmed inventory.
Use them for demo/testing only; do not treat them as operational records.

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

### Route B: Node sends requests to the BluConn API

The Node uploader can take an Excel workbook, convert it to entity JSON, and
send the records to the BluConn API in one run. It uses the API URL, account
ID, and bearer token configured in the local `.env` file. It can also upload
an existing entity JSON file. For a workbook, it includes the six simulated
demo types described above and resolves their API-generated entity references
before sending dependent records. It maps JSON keys to the API's canonical
hyphenated entity-type slugs and reuses matching existing project, asset, and
material IDs so child records can still be linked on a repeat upload. Since
the API already has logistics IDs with the same short values, the API route
prefixes logistics IDs with the project ID and updates their procurement
references; the source JSON and direct PostgreSQL route remain unchanged. It
does not sign in through Cognito, use a user ID, or read or require
`DATABASE_URL`.

For flat entity rows, the script converts groups of JSON records into temporary
Excel files in memory and sends them through the API bulk-upload endpoint.
When an entity schema has nested objects or arrays, the script sends JSON
records directly to the API so those values remain objects and arrays rather
than becoming spreadsheet text. Rows that require a parent but have no parent
reference (for example, an engineering log without an asset) are attempted
without a parent, matching the direct PostgreSQL uploader. If the API requires
a parent or rejects the record for another reason, the failure is reported.
The BluConn backend handles writing accepted records to its database. The API
route does not automatically run the Python database uploader's workbook
validation.

Before each API upload, the Node uploader applies the API entity schema:
date-only values are converted when a date-time is required, blank optional
values that violate a format or pattern are omitted, and required fields,
enums, patterns, and date-time values are checked locally. A missing required
value stops that upload with the field path instead of sending a row the API
will reject.

API credentials and database credentials are different. Keep the bearer token
Bearer tokens expire and must be replaced with a fresh token when needed. The
current API uploader does not send a separate user-ID header.

## 8. How to tell what happened

There are three different outcomes to distinguish:

1. **Workbook created:** the planning or update step completed.
2. **Records prepared and validated:** entity JSON was created and the
   validation report shows its result.
3. **Records reached the database:** the selected uploader connected to its
   destination and reported successful insert/API counts.

A workbook or a passing validation report alone does not mean data was
uploaded. A connection error means the push did not reach the database. After
a successful push, review the push summary/report and confirm the project in
the BluConn application.

## Main parts of the application

- **Project configuration:** project-specific guidelines and assumptions for
  roads, pipelines, and buildings.
- **Planning engine:** quantity, cost, duration, calendar, resource, and CPM
  calculations.
- **Excel reporting:** creates baseline and updated project workbooks.
- **Entity builder and validator:** converts workbook content into structured
  records and checks them against the workbook.
- **Python database uploader:** connects directly to PostgreSQL.
- **Node API uploader:** sends authenticated requests to the BluConn API.
