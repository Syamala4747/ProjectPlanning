/**
 * push-entities.mjs
 * =================
 * Reads the *_actual_data.json produced by actual_to_entity.py and
 * pushes all 14 entity types to the BluConn API.
 *
 * The JSON has the shape:  { slug: [instance_dict, ...] }
 * This script converts each slug's array into a temporary in-memory
 * Excel buffer (row 1 = headers, row 2+ = data) and calls the
 * bulk-upload endpoint — one call per entity type, in hierarchy order.
 *
 * Usage
 * -----
 *   node push-entities.mjs [--json=<path>] [--wipe] [--schema-only]
 *                          [--account=<id>] [--api=<url>]
 *
 * If --json is omitted the script auto-finds the most recently written
 * *_actual_data.json in ./output/.
 *
 * Credentials come from .env:
 *   ACCOUNT_ID, BASE_URL, COGNITO_USER_POOL_ID, COGNITO_CLIENT_ID,
 *   PUSH_EMAIL, PUSH_PASSWORD
 */

import { CognitoIdentityProviderClient, InitiateAuthCommand } from "@aws-sdk/client-cognito-identity-provider";
import ExcelJS from "exceljs";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── .env loader ───────────────────────────────────────────────────────────────
const envPath = path.resolve(__dirname, ".env");
if (fs.existsSync(envPath)) {
  for (const line of fs.readFileSync(envPath, "utf-8").split("\n")) {
    const t = line.trim();
    if (!t || t.startsWith("#")) continue;
    const eq = t.indexOf("=");
    if (eq === -1) continue;
    const key = t.slice(0, eq).trim();
    let val = t.slice(eq + 1).trim();
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'")))
      val = val.slice(1, -1);
    if (!(key in process.env)) process.env[key] = val;
  }
}

// ── config ────────────────────────────────────────────────────────────────────
const ACCOUNT_ID = process.env.ACCOUNT_ID
  || process.argv.find(a => a.startsWith("--account="))?.split("=")[1]
  || "";

const BASE_URL = process.env.BASE_URL
  || process.argv.find(a => a.startsWith("--api="))?.split("=")[1]
  || "https://remote-dev.bluconn.ai/sandilya_m_bluconn/api";

const EMAIL    = process.env.PUSH_EMAIL    || "";
const PASSWORD = process.env.PUSH_PASSWORD || "";
const WIPE        = process.argv.includes("--wipe");
const SCHEMA_ONLY = process.argv.includes("--schema-only");
const CHUNK_SIZE  = 5000;

// ── find JSON file ────────────────────────────────────────────────────────────
function findJsonFile() {
  const arg = process.argv.find(a => a.startsWith("--json="))?.split("=")[1];
  if (arg) return arg;
  // Auto-find most recent *_actual_data.json in ./output/
  const outDir = path.join(__dirname, "output");
  if (!fs.existsSync(outDir)) throw new Error(`output/ directory not found. Run the planner first.`);
  const files = fs.readdirSync(outDir)
    .filter(f => f.endsWith("_actual_data.json"))
    .map(f => ({ f, t: fs.statSync(path.join(outDir, f)).mtime }))
    .sort((a, b) => b.t - a.t);
  if (!files.length) throw new Error(`No *_actual_data.json found in output/. Run the planner first.`);
  return path.join(outDir, files[0].f);
}

// ── hierarchy: slug → parent field in the data ───────────────────────────────
// Defines upload order (parents before children) and how to find the parent ID.
const HIERARCHY = [
  // standalone / root
  { slug: "project",                       parentSlug: null,         idCol: "projectId",                    parentField: null,         label: "PROJECT" },
  { slug: "assumption",                    parentSlug: null,         idCol: "assumptionId",                 parentField: null,         label: "ASSUMPTION" },
  { slug: "asset",                         parentSlug: null,         idCol: "assetId",                      parentField: null,         label: "ASSET" },
  { slug: "material",                      parentSlug: null,         idCol: "materialId",                   parentField: null,         label: "MATERIAL" },
  { slug: "subcontractor",                 parentSlug: null,         idCol: "subContractorId",              parentField: null,         label: "SUBCONTRACTOR" },
  { slug: "logistics",                     parentSlug: null,         idCol: "logisticsId",                  parentField: null,         label: "LOGISTICS" },
  // children
  { slug: "boq",                           parentSlug: "project",    idCol: "boqId",                        parentField: "projectId",  label: "BOQ" },
  { slug: "activity",                      parentSlug: "boq",        idCol: "activityId",                   parentField: "boqId",      label: "ACTIVITY" },
  { slug: "plan",                          parentSlug: "activity",   idCol: "planId",                       parentField: "activityId", label: "PLAN" },
  { slug: "supervisor_update",             parentSlug: "activity",   idCol: "supervisorUpdateId",           parentField: "activityId", label: "SUPERVISOR UPDATE" },
  { slug: "engineering_daily_logs",        parentSlug: "asset",      idCol: "engineeringDailyLogId",        parentField: "assetId",    label: "ENGINEERING DAILY LOGS" },
  { slug: "material_procurement_request",  parentSlug: "material",   idCol: "materialProcurementRequestId", parentField: "materialId", label: "MATERIAL PROCUREMENT REQUEST" },
  { slug: "material_issue_request",        parentSlug: "material",   idCol: "materialIssueRequestId",       parentField: "materialId", label: "MATERIAL ISSUE REQUEST" },
  { slug: "material_transaction",          parentSlug: "material",   idCol: "materialTransactionId",        parentField: "materialId", label: "MATERIAL TRANSACTION" },
];

// ── auth ──────────────────────────────────────────────────────────────────────
let authToken = "";

async function authenticate() {
  if (!EMAIL || !PASSWORD) throw new Error("Set PUSH_EMAIL and PUSH_PASSWORD in .env");
  console.log("Authenticating...");
  const userPoolId = process.env.COGNITO_USER_POOL_ID;
  const clientId   = process.env.COGNITO_CLIENT_ID;
  if (!userPoolId || !clientId)
    throw new Error("Set COGNITO_USER_POOL_ID and COGNITO_CLIENT_ID in .env");
  const region = userPoolId.split("_")[0];
  const client = new CognitoIdentityProviderClient({ region });
  const resp = await client.send(new InitiateAuthCommand({
    AuthFlow: "USER_PASSWORD_AUTH",
    ClientId: clientId,
    AuthParameters: { USERNAME: EMAIL, PASSWORD: PASSWORD },
  }));
  if (!resp.AuthenticationResult?.IdToken) throw new Error("Auth failed");
  authToken = resp.AuthenticationResult.IdToken;
  console.log("Authenticated.\n");
}

// ── API helpers ───────────────────────────────────────────────────────────────
const hdrs = () => ({ Authorization: `Bearer ${authToken}` });

async function apiGet(p)       { const r = await fetch(`${BASE_URL}${p}`, { headers: hdrs() }); if (!r.ok) throw new Error(`GET ${p} (${r.status}): ${await r.text()}`); return r.json(); }
async function apiPost(p, b)   { const r = await fetch(`${BASE_URL}${p}`, { method:"POST",   headers:{...hdrs(),"Content-Type":"application/json"}, body:JSON.stringify(b) }); if (!r.ok) throw new Error(`POST ${p} (${r.status}): ${await r.text()}`); return r.json(); }
async function apiPatch(p, b)  { const r = await fetch(`${BASE_URL}${p}`, { method:"PATCH",  headers:{...hdrs(),"Content-Type":"application/json"}, body:JSON.stringify(b) }); if (!r.ok) throw new Error(`PATCH ${p} (${r.status}): ${await r.text()}`); return r.json(); }
async function apiDelete(p)    { const r = await fetch(`${BASE_URL}${p}`, { method:"DELETE", headers: hdrs() }); if (!r.ok && r.status !== 404) throw new Error(`DELETE ${p} (${r.status})`); }
async function apiMultipart(p, fd) { const r = await fetch(`${BASE_URL}${p}`, { method:"POST", headers: hdrs(), body: fd }); if (!r.ok) throw new Error(`POST ${p} (${r.status}): ${await r.text()}`); return r.json(); }

// ── JSON → Excel buffer (in-memory, no files written) ────────────────────────
function flatten(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return v;
}

async function toExcelBuffer(rows) {
  if (!rows || rows.length === 0) return null;
  // Collect all headers from the union of all row keys
  const headerSet = new Set();
  for (const row of rows) Object.keys(row).forEach(k => headerSet.add(k));
  const headers = [...headerSet];

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.addRow(headers);
  for (const row of rows) ws.addRow(headers.map(h => flatten(row[h])));
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ── bulk upload helper ────────────────────────────────────────────────────────
async function bulkUpload(entityTypeId, parentId, rows, slugLabel) {
  if (!rows || rows.length === 0) return { entitiesCreated: 0, entitiesFailed: 0, entityIds: [] };

  let totalCreated = 0, totalFailed = 0;
  const allIds = [];

  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    const buf = await toExcelBuffer(chunk);
    if (!buf) continue;

    const fd = new FormData();
    fd.append("file", new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
              `${slugLabel}_${Math.floor(i / CHUNK_SIZE) + 1}.xlsx`);
    fd.append("entityTypeId", entityTypeId);
    if (parentId) fd.append("parentId", parentId);

    let retries = 0;
    while (retries <= 3) {
      try {
        const res = await apiMultipart(`/accounts/${ACCOUNT_ID}/entities/bulk-upload`, fd);
        totalCreated += res.entitiesCreated || 0;
        totalFailed  += res.entitiesFailed  || 0;
        if (res.entityIds) allIds.push(...res.entityIds);
        break;
      } catch (err) {
        retries++;
        if (retries > 3) { console.error(`   chunk ${Math.floor(i/CHUNK_SIZE)+1} failed: ${err.message}`); totalFailed += chunk.length; break; }
        await new Promise(r => setTimeout(r, retries * 2000));
      }
    }
  }
  return { entitiesCreated: totalCreated, entitiesFailed: totalFailed, entityIds: allIds };
}

// ── schema type inference ─────────────────────────────────────────────────────
function inferType(sample) {
  if (sample === null || sample === undefined || sample === "") return { type: "string" };
  if (typeof sample === "number")  return { type: "number" };
  if (typeof sample === "boolean") return { type: "boolean" };
  const s = String(sample);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return { type: "string", format: "date-time" };
  return { type: "string" };
}

function buildJsonSchema(rows, idCol) {
  const sample = rows[0] || {};
  const props = {};
  for (const [k, v] of Object.entries(sample)) {
    props[k] = { ...inferType(v), title: k };
    if (props[k].type === "object" || props[k].type === "array") props[k].type = "string";
  }
  return {
    "$schema": "http://json-schema.org/draft-07/schema#",
    type: "object",
    properties: props,
    required: idCol ? [idCol] : [],
  };
}

// ── wipe ──────────────────────────────────────────────────────────────────────
async function wipeAll() {
  console.log("=== WIPING EXISTING DATA ===\n");
  const types = await apiGet(`/accounts/${ACCOUNT_ID}/entities/entity-types`);
  for (const et of types.entityTypes) {
    let del = 0, more = true;
    while (more) {
      const res = await apiGet(`/accounts/${ACCOUNT_ID}/entities?entityTypeId=${et.id}&page[size]=50`);
      if (!res.entities?.length) break;
      for (const e of res.entities) { try { await apiDelete(`/accounts/${ACCOUNT_ID}/entities/${e.id}`); del++; } catch {} }
      more = res.entities.length === 50;
    }
    if (del) console.log(`  Deleted ${del} × ${et.name}`);
  }
  console.log();
}

// ── create / update entity types ──────────────────────────────────────────────
async function ensureEntityTypes(data) {
  console.log("=== ENSURING ENTITY TYPES ===\n");
  const existing = await apiGet(`/accounts/${ACCOUNT_ID}/entities/entity-types`);
  const typeMap  = new Map(existing.entityTypes.map(et => [et.slug, et]));

  for (const step of HIERARCHY) {
    const rows = data[step.slug] || [];
    if (!rows.length && !typeMap.has(step.slug)) { console.log(`  SKIP (no data): ${step.label}`); continue; }

    if (typeMap.has(step.slug)) {
      console.log(`  EXISTS: ${step.label.padEnd(38)} id=${typeMap.get(step.slug).id}`);
      continue;
    }

    const parentType = step.parentSlug ? typeMap.get(step.parentSlug) : null;
    const schema     = buildJsonSchema(rows, step.idCol);
    const payload    = {
      name:             step.label,
      slug:             step.slug,
      parentMode:       step.parentSlug ? "RESTRICTED" : "NONE",
      isParentRequired: Boolean(step.parentSlug),
      jsonSchema:       schema,
      searchableFields: step.idCol ? [step.idCol] : [],
      ...(parentType ? { allowedParentTypeIds: [parentType.id] } : {}),
    };

    try {
      const res = await apiPost(`/accounts/${ACCOUNT_ID}/entity-types`, payload);
      typeMap.set(step.slug, res);
      console.log(`  CREATED: ${step.label.padEnd(38)} id=${res.id}`);
    } catch (err) {
      if (err.message.includes("409")) {
        const fresh = await apiGet(`/accounts/${ACCOUNT_ID}/entities/entity-types`);
        const found = fresh.entityTypes.find(et => et.slug === step.slug);
        if (found) { typeMap.set(step.slug, found); console.log(`  EXISTS: ${step.label.padEnd(38)} id=${found.id} (409 resolved)`); }
      } else {
        console.error(`  ERROR ${step.label}: ${err.message}`);
      }
    }
  }

  // Ensure parent-child relationships
  console.log();
  const rels = await apiGet(`/accounts/${ACCOUNT_ID}/entity-type-relationships`);
  const relSet = new Set(rels.map(r => `${r.parentTypeId}:${r.childTypeId}`));
  for (const step of HIERARCHY) {
    if (!step.parentSlug) continue;
    const child  = typeMap.get(step.slug);
    const parent = typeMap.get(step.parentSlug);
    if (!child || !parent) continue;
    const key = `${parent.id}:${child.id}`;
    if (relSet.has(key)) continue;
    try {
      await apiPost(`/accounts/${ACCOUNT_ID}/entity-type-relationships`, { parentTypeId: parent.id, childTypeId: child.id });
      console.log(`  REL: ${step.parentSlug} → ${step.slug}`);
    } catch (err) {
      if (!err.message.includes("409")) console.error(`  REL ERROR: ${step.parentSlug}→${step.slug}: ${err.message}`);
    }
  }

  return typeMap;
}

// ── upload data ───────────────────────────────────────────────────────────────
async function uploadAll(data, typeMap) {
  console.log("\n=== UPLOADING DATA ===\n");
  const idMap = new Map();   // slug → { localId → serverEntityId }

  for (const step of HIERARCHY) {
    const rows = data[step.slug] || [];
    const et   = typeMap.get(step.slug);
    console.log(`>> ${step.label}  (${rows.length} records)`);
    if (!rows.length || !et) { idMap.set(step.slug, new Map()); console.log(`   skipped\n`); continue; }

    const mapping = new Map();

    if (!step.parentSlug) {
      // Root / standalone — upload all at once
      const res = await bulkUpload(et.id, null, rows, step.slug);
      res.entityIds.forEach((eid, i) => {
        if (rows[i]) mapping.set(String(rows[i][step.idCol]), eid);
      });
      console.log(`   created=${res.entitiesCreated}  failed=${res.entitiesFailed}  mapped=${mapping.size}\n`);
    } else {
      // Child — group by parent ID, upload per parent
      const parentMap = idMap.get(step.parentSlug) || new Map();
      const groups    = new Map();
      const orphans   = [];

      for (const row of rows) {
        const localParentId = row[step.parentField] ? String(row[step.parentField]) : null;
        const serverParentId = localParentId ? parentMap.get(localParentId) : null;
        if (!serverParentId) { orphans.push(row); continue; }
        if (!groups.has(serverParentId)) groups.set(serverParentId, []);
        groups.get(serverParentId).push(row);
      }

      if (orphans.length) console.log(`   ${orphans.length} orphans (parent not found in previous upload)`);
      console.log(`   ${groups.size} parent group(s)`);

      let gc = 0, gf = 0, gi = 0;
      for (const [pid, groupRows] of groups) {
        gi++;
        const res = await bulkUpload(et.id, pid, groupRows, `${step.slug}_g${gi}`);
        gc += res.entitiesCreated; gf += res.entitiesFailed;
        res.entityIds.forEach((eid, i) => { if (groupRows[i]) mapping.set(String(groupRows[i][step.idCol]), eid); });
        if (gi % 20 === 0 || groups.size <= 5) console.log(`   [${gi}/${groups.size}] created=${res.entitiesCreated} failed=${res.entitiesFailed}`);
      }

      // Upload orphans without parent
      if (orphans.length) {
        const res = await bulkUpload(et.id, null, orphans, `${step.slug}_orphans`);
        gc += res.entitiesCreated; gf += res.entitiesFailed;
        res.entityIds.forEach((eid, i) => { if (orphans[i]) mapping.set(String(orphans[i][step.idCol]), eid); });
      }

      console.log(`   TOTAL: created=${gc}  failed=${gf}  mapped=${mapping.size}\n`);
    }

    idMap.set(step.slug, mapping);
  }

  return idMap;
}

// ── main ──────────────────────────────────────────────────────────────────────
async function main() {
  const jsonPath = findJsonFile();

  console.log("=== push-entities ===\n");
  console.log(`JSON    : ${jsonPath}`);
  console.log(`Account : ${ACCOUNT_ID}`);
  console.log(`API     : ${BASE_URL}`);
  console.log(`Wipe    : ${WIPE}`);
  console.log(`Schema  : ${SCHEMA_ONLY ? "only" : "full upload"}\n`);

  if (!ACCOUNT_ID) throw new Error("ACCOUNT_ID is not set. Add it to .env");

  const data = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));

  console.log("Entity counts in JSON:");
  for (const [slug, rows] of Object.entries(data))
    console.log(`  ${slug.padEnd(38)} ${rows.length}`);
  console.log();

  await authenticate();
  if (WIPE) await wipeAll();

  const typeMap = await ensureEntityTypes(data);
  if (SCHEMA_ONLY) { console.log("\nSchema-only mode — skipping data upload.\n"); return; }

  const idMap = await uploadAll(data, typeMap);

  console.log("=== SUMMARY ===\n");
  for (const step of HIERARCHY) {
    const m = idMap.get(step.slug);
    console.log(`  ${step.label.padEnd(38)} : ${m?.size ?? 0} pushed`);
  }
  console.log("\n=== DONE ===\n");
}

main().catch(err => { console.error("\nFATAL:", err.message || err); process.exit(1); });
