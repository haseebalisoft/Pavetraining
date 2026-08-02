/**
 * One-by-one Graph vs SharePoint schema audit + Outlook/login readiness.
 *
 *   node --env-file=.env.local scripts/audit-graph-sharepoint-schema.mjs
 */
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

function loadEnv() {
  try {
    for (const line of readFileSync(resolve(ROOT, ".env.local"), "utf8").split(
      /\r?\n/,
    )) {
      const t = line.trim();
      if (!t || t.startsWith("#")) continue;
      const i = t.indexOf("=");
      if (i < 0) continue;
      const k = t.slice(0, i).trim();
      let v = t.slice(i + 1).trim();
      if (
        (v.startsWith('"') && v.endsWith('"')) ||
        (v.startsWith("'") && v.endsWith("'"))
      )
        v = v.slice(1, -1);
      if (!process.env[k]) process.env[k] = v;
    }
  } catch {
    // ignore
  }
}
loadEnv();

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing env ${name}`);
  return value;
}
function hasEnv(name) {
  return Boolean(process.env[name]?.trim());
}

function sitePath() {
  const siteId = String(requireEnv("SHAREPOINT_SITE_ID")).replace(/\/+$/, "");
  if (siteId.includes(":/")) {
    return `/sites/${siteId.endsWith(":") ? siteId : `${siteId}:`}`;
  }
  return `/sites/${siteId}`;
}

async function getToken() {
  const body = new URLSearchParams({
    client_id: requireEnv("AZURE_CLIENT_ID"),
    client_secret: requireEnv("AZURE_CLIENT_SECRET"),
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(
    `https://login.microsoftonline.com/${requireEnv("AZURE_TENANT_ID")}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
  );
  if (!res.ok) throw new Error(`Token ${res.status}: ${await res.text()}`);
  return (await res.json()).access_token;
}

async function graph(token, path, init = {}) {
  const url = path.startsWith("http")
    ? path
    : `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      "Content-Type": "application/json",
      ...(init.headers || {}),
    },
  });
  const text = await res.text();
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = { raw: text.slice(0, 200) };
  }
  return { ok: res.ok, status: res.status, json, text };
}

function parseSchemaLists() {
  const text = readFileSync(
    resolve(ROOT, "src/lib/schema/sharepointSchema.ts"),
    "utf8",
  );
  const lists = [];
  // Match list blocks: key: { key: "...", listName: "...", listIdEnvVar: "...", fields: xxxFields
  const blockRe =
    /(\w+):\s*\{\s*key:\s*"(\w+)",\s*listName:\s*"([^"]+)",\s*displayName:\s*"([^"]+)",\s*listIdEnvVar:\s*"([^"]+)",\s*fields:\s*(\w+)/g;
  let m;
  while ((m = blockRe.exec(text))) {
    lists.push({
      objectKey: m[1],
      key: m[2],
      listName: m[3],
      displayName: m[4],
      listIdEnvVar: m[5],
      fieldsConst: m[6],
    });
  }

  // Parse each fields const
  const fieldsByConst = {};
  const fieldsRe = /const (\w+Fields) = \{([\s\S]*?)\} as const;/g;
  while ((m = fieldsRe.exec(text))) {
    const name = m[1];
    const body = m[2];
    const map = {};
    const entryRe = /(\w+):\s*"([^"]+)"/g;
    let e;
    while ((e = entryRe.exec(body))) {
      map[e[1]] = e[2];
    }
    fieldsByConst[name] = map;
  }

  for (const list of lists) {
    list.fields = fieldsByConst[list.fieldsConst] || {};
  }
  return lists;
}

const rows = [];
function record(area, name, ok, detail) {
  rows.push({ area, name, ok, detail });
  const mark = ok ? "PASS" : "FAIL";
  console.log(`${mark}  [${area}] ${name}${detail ? ` — ${detail}` : ""}`);
}

const token = await getToken();
const root = sitePath();
console.log("=== Graph token OK ===");
console.log(`Site root: ${root}\n`);

// Site resolve
{
  const site = await graph(token, `/sites/${requireEnv("SHAREPOINT_SITE_ID")}`);
  record(
    "site",
    "Resolve SharePoint site",
    site.ok,
    site.ok
      ? site.json?.displayName || site.json?.name || "ok"
      : `HTTP ${site.status}`,
  );
}

const lists = parseSchemaLists();
console.log(`\n=== Lists (${lists.length}) one-by-one ===\n`);

for (const list of lists) {
  const listId = process.env[list.listIdEnvVar];
  if (!listId) {
    record("list", list.listName, false, `Missing env ${list.listIdEnvVar}`);
    continue;
  }

  const meta = await graph(token, `${root}/lists/${listId}`);
  if (!meta.ok) {
    record(
      "list",
      list.listName,
      false,
      `Cannot read list (${list.listIdEnvVar}): HTTP ${meta.status}`,
    );
    continue;
  }

  const liveName = meta.json?.displayName || meta.json?.name || "";
  const nameOk =
    liveName.toLowerCase() === list.listName.toLowerCase() ||
    liveName.toLowerCase().includes(list.displayName.toLowerCase()) ||
    list.listName.toLowerCase().includes(liveName.toLowerCase());

  // Columns
  let cols = [];
  let next = `${root}/lists/${listId}/columns?$top=200`;
  while (next) {
    const page = await graph(token, next);
    if (!page.ok) break;
    cols.push(...(page.json?.value || []));
    next = page.json?.["@odata.nextLink"] || null;
  }
  const colNames = new Set(cols.map((c) => String(c.name || "")));
  // Also accept LookupId companions as present if base exists
  const required = Object.entries(list.fields).filter(
    ([k, v]) => k !== "id" && v !== "ID",
  );
  const missing = [];
  for (const [, internal] of required) {
    if (colNames.has(internal)) continue;
    // Graph companions like XxxLookupId are not real columns
    if (internal.endsWith("LookupId")) continue;
    missing.push(internal);
  }

  // Sample items
  const items = await graph(
    token,
    `${root}/lists/${listId}/items?$expand=fields&$top=5`,
  );
  const itemCount = items.ok ? (items.json?.value || []).length : -1;

  const ok = meta.ok && missing.length === 0;
  record(
    "list",
    `${list.key} / ${list.listName}`,
    ok,
    `live="${liveName}" nameMatch=${nameOk} cols=${cols.length} missingFields=${missing.length}${
      missing.length ? ` [${missing.slice(0, 8).join(", ")}${missing.length > 8 ? "…" : ""}]` : ""
    } sampleItems=${itemCount}`,
  );

  // Extra: Training Matrix Update should have many N### columns
  if (list.key === "trainingMatrixExample") {
    const nCols = cols.filter((c) =>
      /^N\d+/i.test(String(c.displayName || "")),
    ).length;
    record(
      "matrix",
      "Training Matrix Update N### columns",
      nCols >= 100,
      `${nCols} N-code display columns`,
    );
  }
}

// Permissions seed / RoleType choices
console.log("\n=== Permissions / login data ===\n");
{
  const listId = process.env.SHAREPOINT_PERMISSIONS_LIST_ID;
  if (listId) {
    const items = await graph(
      token,
      `${root}/lists/${listId}/items?$expand=fields&$top=50`,
    );
    const values = items.json?.value || [];
    const activeAdmins = values.filter((i) => {
      const f = i.fields || {};
      return (
        String(f.RoleType || "").toLowerCase() === "admin" &&
        String(f.Status || "").toLowerCase() === "active"
      );
    });
    const customers = values.filter((i) => {
      const f = i.fields || {};
      return (
        String(f.RoleType || "").toLowerCase() === "customer" &&
        String(f.Status || "").toLowerCase() === "active"
      );
    });
    record(
      "auth-data",
      "Permissions List readable",
      items.ok,
      `${values.length} rows, ${activeAdmins.length} active Admin, ${customers.length} active Customer`,
    );
  } else {
    record("auth-data", "Permissions List", false, "Missing list id env");
  }
}

// Auth env (login)
console.log("\n=== Login / Auth env ===\n");
for (const key of [
  "AUTH_SECRET",
  "AUTH_URL",
  "AUTH_MICROSOFT_ENTRA_ID_ID",
  "AUTH_MICROSOFT_ENTRA_ID_SECRET",
  "AUTH_MICROSOFT_ENTRA_ID_ISSUER",
  "AZURE_TENANT_ID",
  "AZURE_CLIENT_ID",
  "AZURE_CLIENT_SECRET",
]) {
  record("login", key, hasEnv(key), hasEnv(key) ? "set" : "MISSING");
}

// Outlook
console.log("\n=== Outlook / Events ===\n");
const outlookUser = hasEnv("OUTLOOK_USER_ID") || hasEnv("OUTLOOK_MAILBOX");
record(
  "outlook",
  "OUTLOOK_USER_ID configured",
  outlookUser,
  outlookUser
    ? "set — Outlook sync can run"
    : "NOT SET in .env.local — Events sync to Outlook will be skipped until configured",
);
record(
  "outlook",
  "OUTLOOK_CALENDAR_ID optional",
  true,
  hasEnv("OUTLOOK_CALENDAR_ID")
    ? "set"
    : "not set (defaults to mailbox calendar)",
);

{
  const listId = process.env.SHAREPOINT_EVENTS_LIST_ID;
  if (listId) {
    const items = await graph(
      token,
      `${root}/lists/${listId}/items?$expand=fields&$top=10`,
    );
    const values = items.json?.value || [];
    const withOutlook = values.filter((i) => i.fields?.OutlookEventId);
    const syncFields = [
      "OutlookEventId",
      "SyncStatus",
      "DoNotSync",
      "LastSyncSource",
      "SyncHash",
    ];
    const cols = [];
    let next = `${root}/lists/${listId}/columns?$top=200`;
    while (next) {
      const page = await graph(token, next);
      if (!page.ok) break;
      cols.push(...(page.json?.value || []));
      next = page.json?.["@odata.nextLink"] || null;
    }
    const names = new Set(cols.map((c) => c.name));
    const missingSync = syncFields.filter((f) => !names.has(f));
    record(
      "events",
      "Events list + sync columns",
      items.ok && missingSync.length === 0,
      `${values.length} sample rows, ${withOutlook.length} with OutlookEventId, missingSyncCols=${missingSync.join(",") || "none"}`,
    );

    // If Outlook configured, try a dry read of mailbox calendars
    if (outlookUser) {
      const userId = process.env.OUTLOOK_USER_ID || process.env.OUTLOOK_MAILBOX;
      const cal = await graph(
        token,
        `/users/${encodeURIComponent(userId)}/calendar`,
      );
      record(
        "outlook",
        "Read Outlook calendar (Graph)",
        cal.ok,
        cal.ok
          ? `calendar name=${cal.json?.name || "Calendar"}`
          : `HTTP ${cal.status} — check Calendars.ReadWrite app permission + mailbox`,
      );
    }
  }
}

// Client-critical reads: matrix, registers, workforce, company
console.log("\n=== Client-critical data samples ===\n");
async function sampleList(label, envVar, pick) {
  const listId = process.env[envVar];
  if (!listId) {
    record("data", label, false, `Missing ${envVar}`);
    return;
  }
  const items = await graph(
    token,
    `${root}/lists/${listId}/items?$expand=fields&$top=3`,
  );
  if (!items.ok) {
    record("data", label, false, `HTTP ${items.status}`);
    return;
  }
  const values = items.json?.value || [];
  const preview = values.map(pick).filter(Boolean).slice(0, 3).join(" | ");
  record(
    "data",
    label,
    values.length > 0,
    values.length
      ? `${values.length} rows e.g. ${preview}`
      : "0 items (empty list)",
  );
}

await sampleList(
  "Company List",
  "SHAREPOINT_COMPANY_LIST_ID",
  (i) => i.fields?.CompanyName || i.fields?.Title,
);
await sampleList(
  "Workforce List",
  "SHAREPOINT_WORKFORCE_LIST_ID",
  (i) => i.fields?.CandidateName,
);
await sampleList(
  "NPORS Register",
  "SHAREPOINT_NPORS_REGISTER_LIST_ID",
  (i) =>
    `id=${i.id} outcome=${i.fields?.TrainingOutcome || "—"} candLookup=${i.fields?.CandidateNameLookupId || "—"}`,
);
await sampleList(
  "EUSR Register",
  "SHAREPOINT_EUSR_REGISTER_LIST_ID",
  (i) => `id=${i.id} candLookup=${i.fields?.CandidateNameLookupId || "—"}`,
);
await sampleList(
  "Streetworks (NRSWA)",
  "SHAREPOINT_NRSWA_REGISTER_LIST_ID",
  (i) => `id=${i.id} candLookup=${i.fields?.CandidateNameLookupId || "—"}`,
);
await sampleList(
  "In-House Certificates",
  "SHAREPOINT_IN_HOUSE_CERTIFICATES_LIST_ID",
  (i) => `id=${i.id} candLookup=${i.fields?.CandidateNameLookupId || "—"}`,
);
await sampleList(
  "Training Matrix Update",
  "SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID",
  (i) => i.fields?.Title || i.fields?.Name,
);
await sampleList(
  "NVQ Register",
  "SHAREPOINT_NVQ_REGISTER_LIST_ID",
  (i) => `id=${i.id} title=${i.fields?.NvqTitle || "—"}`,
);
await sampleList(
  "Offers / Promotions",
  "SHAREPOINT_OFFERS_PROMOTIONS_LIST_ID",
  (i) => i.fields?.Title,
);
await sampleList(
  "Departments",
  "SHAREPOINT_DEPARTMENTS_LIST_ID",
  (i) => i.fields?.Name || i.fields?.Title,
);

// Production login page
console.log("\n=== Production ===\n");
try {
  const res = await fetch("https://pave-training-portal-nu.vercel.app/login", {
    redirect: "manual",
  });
  record(
    "prod",
    "Login page",
    res.status >= 200 && res.status < 400,
    `HTTP ${res.status}`,
  );
} catch (error) {
  record(
    "prod",
    "Login page",
    false,
    error instanceof Error ? error.message : String(error),
  );
}

console.log("\n=== SUMMARY ===");
const byArea = {};
for (const r of rows) {
  byArea[r.area] ||= { pass: 0, fail: 0 };
  if (r.ok) byArea[r.area].pass += 1;
  else byArea[r.area].fail += 1;
}
for (const [area, s] of Object.entries(byArea)) {
  console.log(`${area}: ${s.pass} pass, ${s.fail} fail`);
}
const failed = rows.filter((r) => !r.ok);
console.log(`\nTOTAL fail: ${failed.length}`);
if (failed.length) {
  console.log("\nFailures:");
  for (const f of failed) {
    console.log(` - [${f.area}] ${f.name}: ${f.detail}`);
  }
  process.exitCode = 1;
} else {
  console.log("\nAll checks passed.");
}
