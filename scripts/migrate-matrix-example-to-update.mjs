/**
 * Copy rows from legacy "Training matrix example" (Excel serials / field_*)
 * into "Training Matrix Update" (real Date columns).
 *
 *   node --env-file=.env scripts/migrate-matrix-example-to-update.mjs
 *   node --env-file=.env scripts/migrate-matrix-example-to-update.mjs --dry-run
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const DRY_RUN = process.argv.includes("--dry-run");
const ROOT = process.cwd();

const SOURCE_LIST_ID =
  process.env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID_PREVIOUS?.trim() ||
  "48d230f0-bcca-46d3-994b-3946750eade5";
const TARGET_LIST_ID = process.env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID;

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing env ${name}`);
  return value;
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
  if (!res.ok) throw new Error(await res.text());
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
      "Content-Type": "application/json",
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      ...(init.headers || {}),
    },
  });
  if (!res.ok) throw new Error(`${res.status} ${url}: ${await res.text()}`);
  if (res.status === 204) return null;
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

async function listAll(token, first) {
  const items = [];
  let path = first;
  while (path) {
    const page = await graph(token, path);
    items.push(...(page.value || []));
    path = page["@odata.nextLink"] || null;
  }
  return items;
}

function normalize(header) {
  return String(header || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function collapse(header) {
  return normalize(header)
    .toLowerCase()
    .replace(/[–—]/g, "-")
    .replace(/[°º]/g, "")
    .replace(/[^a-z0-9]+/g, "");
}

function extractCode(header) {
  const match = normalize(header).match(/^(N\d+[A-Z]?)\b/i);
  return match ? match[1].toUpperCase() : null;
}

function excelSerialToIso(value) {
  if (value == null || value === "") return null;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed || /^(—|–|-|n\/?a|null|none|0)$/i.test(trimmed)) return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) return trimmed.slice(0, 10);
    const asNum = Number(trimmed);
    if (!Number.isNaN(asNum) && asNum > 20000) return excelSerialToIso(asNum);
    const parsed = Date.parse(trimmed);
    if (!Number.isNaN(parsed)) return new Date(parsed).toISOString().slice(0, 10);
    return null;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value) || value === 0) return null;
    const ms = Date.UTC(1899, 11, 30) + Math.round(value) * 86_400_000;
    return new Date(ms).toISOString().slice(0, 10);
  }
  return null;
}

function loadTemplateHeaders() {
  const src = readFileSync(
    resolve(ROOT, "src/lib/services/bulkUpload/clientTemplateHeaders.ts"),
    "utf8",
  );
  const meta = [
    ...src
      .match(/export const CLIENT_MATRIX_META_HEADERS[\s\S]*?\] as const;/)[0]
      .matchAll(/"([^"]+)"/g),
  ].map((m) => m[1]);
  const cats = [...src.matchAll(/"header":\s*"([^"]+)"/g)].map((m) => m[1]);
  return [...meta, ...cats];
}

function buildIndex(columns) {
  const writable = columns.filter(
    (c) =>
      !c.readOnly &&
      ![
        "ContentType",
        "Attachments",
        "Edit",
        "LinkTitleNoMenu",
        "LinkTitle",
      ].includes(c.name),
  );
  const byDisplay = new Map(
    writable.map((c) => [normalize(c.displayName).toLowerCase(), c]),
  );
  const byCollapsed = new Map(
    writable.map((c) => [collapse(c.displayName), c]),
  );
  const byCode = new Map();
  for (const c of writable) {
    const code = extractCode(c.displayName);
    if (!code) continue;
    const existing = byCode.get(code);
    if (
      !existing ||
      (/2$/.test(existing.displayName) && !/2$/.test(c.displayName))
    ) {
      byCode.set(code, c);
    }
  }
  return { writable, byDisplay, byCollapsed, byCode };
}

function resolveHeader(index, header) {
  if (header === "Name") {
    return index.writable.find((c) => c.name === "Title") || null;
  }
  return (
    index.byDisplay.get(normalize(header).toLowerCase()) ||
    index.byCollapsed.get(collapse(header)) ||
    (extractCode(header) ? index.byCode.get(extractCode(header)) : null) ||
    null
  );
}

if (!TARGET_LIST_ID?.trim()) {
  throw new Error("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID is required");
}

const token = await getToken();
const root = sitePath();
const template = loadTemplateHeaders();

console.log("Source (legacy example):", SOURCE_LIST_ID);
console.log("Target (Training Matrix Update):", TARGET_LIST_ID);
console.log(DRY_RUN ? "Mode: DRY RUN" : "Mode: LIVE");

const [sourceCols, targetCols, sourceItems, targetItems] = await Promise.all([
  listAll(token, `${root}/lists/${SOURCE_LIST_ID}/columns?$top=200`),
  listAll(token, `${root}/lists/${TARGET_LIST_ID}/columns?$top=200`),
  listAll(
    token,
    `${root}/lists/${SOURCE_LIST_ID}/items?$expand=fields&$top=200`,
  ),
  listAll(
    token,
    `${root}/lists/${TARGET_LIST_ID}/items?$expand=fields&$top=200`,
  ),
]);

const sourceIndex = buildIndex(sourceCols);
const targetIndex = buildIndex(targetCols);

const existingByName = new Map();
for (const item of targetItems) {
  const title = String(item.fields?.Title || "").trim().toLowerCase();
  if (title) existingByName.set(title, item);
}

console.log(`Source rows: ${sourceItems.length}`);
console.log(`Target rows before: ${targetItems.length}`);

let created = 0;
let updated = 0;
let skipped = 0;
let errors = 0;

for (const item of sourceItems) {
  const fieldsIn = item.fields || {};
  const name = String(fieldsIn.Title || fieldsIn.LinkTitle || "").trim();
  if (!name) {
    skipped += 1;
    continue;
  }

  const out = { Title: name };
  let datesWritten = 0;

  for (const header of template) {
    if (header === "Name") continue;
    const sourceCol = resolveHeader(sourceIndex, header);
    const targetCol = resolveHeader(targetIndex, header);
    if (!sourceCol || !targetCol) continue;

    const iso = excelSerialToIso(fieldsIn[sourceCol.name]);
    if (!iso) continue;
    out[targetCol.name] = iso;
    datesWritten += 1;
  }

  const key = name.toLowerCase();
  const existing = existingByName.get(key);

  console.log(
    `${existing ? "UPDATE" : "CREATE"} ${name} (${datesWritten} dates)`,
  );

  if (DRY_RUN) {
    if (existing) updated += 1;
    else created += 1;
    continue;
  }

  try {
    if (existing) {
      await graph(
        token,
        `${root}/lists/${TARGET_LIST_ID}/items/${existing.id}/fields`,
        {
          method: "PATCH",
          body: JSON.stringify(out),
        },
      );
      updated += 1;
    } else {
      const createdItem = await graph(
        token,
        `${root}/lists/${TARGET_LIST_ID}/items`,
        {
          method: "POST",
          body: JSON.stringify({ fields: out }),
        },
      );
      existingByName.set(key, createdItem);
      created += 1;
    }
  } catch (error) {
    errors += 1;
    console.error(`  FAIL ${name}: ${error.message}`);
  }
}

const after = await listAll(
  token,
  `${root}/lists/${TARGET_LIST_ID}/items?$select=id&$top=200`,
);

console.log("\n=== Migration summary ===");
console.log(`Created: ${created}`);
console.log(`Updated: ${updated}`);
console.log(`Skipped: ${skipped}`);
console.log(`Errors: ${errors}`);
console.log(`Target rows after: ${after.length}`);

if (errors > 0) process.exit(1);
console.log(DRY_RUN ? "\nDRY RUN complete." : "\nMigration complete.");
