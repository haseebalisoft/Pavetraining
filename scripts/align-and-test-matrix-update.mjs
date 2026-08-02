/**
 * Align Training Matrix Update columns with Training matrix example.xlsx,
 * then upload ISO dates and live-fetch to verify admin-facing headers.
 *
 *   node --env-file=.env scripts/align-and-test-matrix-update.mjs
 *   node --env-file=.env scripts/align-and-test-matrix-update.mjs --dry-run
 */
import { createRequire } from "node:module";
import { resolve, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const require = createRequire(import.meta.url);
const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DRY_RUN = process.argv.includes("--dry-run");

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
      "Content-Type": "application/json",
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      ...(init.headers || {}),
    },
  });
  if (!res.ok) {
    throw new Error(`Graph ${res.status} ${url}: ${await res.text()}`);
  }
  if (res.status === 204) return null;
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

async function listAll(token, firstPath) {
  const items = [];
  let path = firstPath;
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
  const m = normalize(header).match(/^(N\d+[A-Z]?)\b/i);
  return m ? m[1].toUpperCase() : null;
}

function internalFromHeader(header) {
  if (header === "DOB") return "DOB";
  if (header === "Face ift") return "Faceift";
  const code = extractCode(header);
  if (!code) {
    return header.replace(/[^A-Za-z0-9]/g, "").slice(0, 32);
  }
  // SharePoint-style encoded name from display (best-effort for create)
  const rest = normalize(header)
    .replace(/^[Nn]\d+[A-Za-z]?\s*-\s*/, "")
    .replace(/[^A-Za-z0-9]/g, "");
  return `${code}_x002d_${rest}`.slice(0, 32);
}

const headersMod = await import(
  pathToFileURL(
    resolve(ROOT, "src/lib/services/bulkUpload/clientTemplateHeaders.ts"),
  ).href
).catch(() => null);

// Fallback: parse headers from the TS source without TS loader
function loadTemplateHeadersFromSource() {
  const fs = require("fs");
  const src = fs.readFileSync(
    resolve(ROOT, "src/lib/services/bulkUpload/clientTemplateHeaders.ts"),
    "utf8",
  );
  const meta = [
    ...src
      .match(
        /export const CLIENT_MATRIX_META_HEADERS[\s\S]*?\] as const;/,
      )[0]
      .matchAll(/"([^"]+)"/g),
  ].map((m) => m[1]);
  const cats = [...src.matchAll(/"header":\s*"([^"]+)"/g)].map((m) => m[1]);
  return [...meta, ...cats];
}

const TEMPLATE_HEADERS =
  headersMod?.CLIENT_MATRIX_DISPLAY_HEADERS ?? loadTemplateHeadersFromSource();

const listId = requireEnv("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID");
const token = await getToken();
const root = sitePath();

console.log("List:", listId);
console.log("Template headers:", TEMPLATE_HEADERS.length);
console.log(DRY_RUN ? "Mode: DRY RUN" : "Mode: LIVE");

const columns = await listAll(token, `${root}/lists/${listId}/columns?$top=200`);
const writable = columns.filter(
  (c) =>
    !c.readOnly &&
    c.name !== "ContentType" &&
    c.name !== "Attachments" &&
    c.name !== "Edit" &&
    c.name !== "LinkTitleNoMenu" &&
    c.name !== "LinkTitle",
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
  if (!existing || /2$/.test(existing.displayName) && !/2$/.test(c.displayName)) {
    byCode.set(code, c);
  }
}

function resolveTemplateHeader(header) {
  if (header === "Name") {
    return writable.find((c) => c.name === "Title") || null;
  }
  return (
    byDisplay.get(normalize(header).toLowerCase()) ||
    byCollapsed.get(collapse(header)) ||
    (extractCode(header) ? byCode.get(extractCode(header)) : null) ||
    null
  );
}

const missing = [];
const renameNeeded = [];
for (const header of TEMPLATE_HEADERS) {
  if (header === "Name") continue;
  const col = resolveTemplateHeader(header);
  if (!col) {
    missing.push(header);
    continue;
  }
  if (normalize(col.displayName) !== normalize(header)) {
    renameNeeded.push({
      header,
      current: normalize(col.displayName),
      columnId: col.id,
      name: col.name,
    });
  }
}

console.log("\n=== Missing columns ===");
console.log(missing.length ? missing.join("\n") : "(none)");
console.log("\n=== Display names to align with Excel ===");
console.log(
  renameNeeded.length
    ? renameNeeded.map((r) => `${r.current}  →  ${r.header}`).join("\n")
    : "(none)",
);

// Create missing Date columns
for (const header of missing) {
  const payload = {
    displayName: header,
    description: `Aligned from Training matrix example.xlsx`,
    dateTime: {
      format: "dateOnly",
    },
  };
  // Prefer a stable name when possible
  payload.name = internalFromHeader(header);
  console.log(`\nCreate column: ${header} (name=${payload.name})`);
  if (DRY_RUN) continue;
  try {
    const created = await graph(token, `${root}/lists/${listId}/columns`, {
      method: "POST",
      body: JSON.stringify(payload),
    });
    console.log(`  OK id=${created.id} name=${created.name}`);
    writable.push(created);
    byDisplay.set(normalize(created.displayName).toLowerCase(), created);
    byCollapsed.set(collapse(created.displayName), created);
    const code = extractCode(created.displayName);
    if (code) byCode.set(code, created);
  } catch (error) {
    // Retry without explicit name (let SharePoint generate)
    console.warn(`  Named create failed: ${error.message}`);
    try {
      const created = await graph(token, `${root}/lists/${listId}/columns`, {
        method: "POST",
        body: JSON.stringify({
          displayName: header,
          dateTime: { format: "dateOnly" },
        }),
      });
      console.log(`  OK (auto-name) id=${created.id} name=${created.name}`);
      writable.push(created);
      byDisplay.set(normalize(created.displayName).toLowerCase(), created);
      byCollapsed.set(collapse(created.displayName), created);
      const code = extractCode(created.displayName);
      if (code) byCode.set(code, created);
    } catch (error2) {
      console.error(`  FAILED: ${error2.message}`);
    }
  }
}

// Rename mismatched display names to Excel headers
for (const item of renameNeeded) {
  console.log(`\nRename ${item.name}: "${item.current}" → "${item.header}"`);
  if (DRY_RUN) continue;
  try {
    await graph(token, `${root}/lists/${listId}/columns/${item.columnId}`, {
      method: "PATCH",
      body: JSON.stringify({ displayName: item.header }),
    });
    console.log("  OK");
  } catch (error) {
    console.error(`  FAILED: ${error.message}`);
  }
}

// Refresh columns after mutations
const refreshed = await listAll(
  token,
  `${root}/lists/${listId}/columns?$top=200`,
);
const refreshedWritable = refreshed.filter(
  (c) =>
    !c.readOnly &&
    c.name !== "ContentType" &&
    c.name !== "Attachments" &&
    c.name !== "Edit" &&
    c.name !== "LinkTitleNoMenu" &&
    c.name !== "LinkTitle",
);

const resolve2 = (header) => {
  if (header === "Name") {
    return refreshedWritable.find((c) => c.name === "Title") || null;
  }
  const byD = new Map(
    refreshedWritable.map((c) => [normalize(c.displayName).toLowerCase(), c]),
  );
  const byC = new Map(
    refreshedWritable.map((c) => [collapse(c.displayName), c]),
  );
  const byCode2 = new Map();
  for (const c of refreshedWritable) {
    const code = extractCode(c.displayName);
    if (!code) continue;
    const existing = byCode2.get(code);
    if (
      !existing ||
      (/2$/.test(existing.displayName) && !/2$/.test(c.displayName))
    ) {
      byCode2.set(code, c);
    }
  }
  return (
    byD.get(normalize(header).toLowerCase()) ||
    byC.get(collapse(header)) ||
    (extractCode(header) ? byCode2.get(extractCode(header)) : null) ||
    null
  );
};

const stillMissing = TEMPLATE_HEADERS.filter(
  (h) => h !== "Name" && !resolve2(h),
);
console.log("\n=== After align: still missing ===");
console.log(stillMissing.length ? stillMissing.join("\n") : "(none)");

// Upload test row with ISO dates on key + previously-missing columns
const marker = `Matrix Align Test ${new Date().toISOString().slice(0, 19)}`;
const testDates = {
  DOB: "1988-04-12",
  "CSCS Expiry": "2028-03-01",
  "N001 - Ind FLT": "2027-06-15",
  "N020 - Tiltrotator System": "2027-07-20",
  "N114 - Overhead Container Gantry Crane": "2027-08-25",
  "N202 - Excavator 360°": "2027-09-10",
  "N216 - Road Planer": "2027-10-05",
};

const fields = { Title: marker };
for (const [header, iso] of Object.entries(testDates)) {
  const col = resolve2(header);
  if (!col) {
    console.warn(`Skip write ${header} — column not found`);
    continue;
  }
  fields[col.name] = iso;
  console.log(`Will write ${header} → ${col.name} = ${iso}`);
}

if (DRY_RUN) {
  console.log("\nDRY RUN complete — no item created.");
  process.exit(0);
}

console.log("\nCreating test item…");
const createdItem = await graph(token, `${root}/lists/${listId}/items`, {
  method: "POST",
  body: JSON.stringify({ fields }),
});
const itemId = createdItem.id;
console.log("Created item", itemId);

const read = await graph(
  token,
  `${root}/lists/${listId}/items/${itemId}?$expand=fields`,
);
const f = read.fields || {};

console.log("\n=== Live fetch (admin-facing headers) ===");
let failures = 0;
for (const [header, expected] of Object.entries(testDates)) {
  const col = resolve2(header);
  if (!col) {
    console.log(`FAIL  ${header}: column missing`);
    failures += 1;
    continue;
  }
  const raw = f[col.name];
  const got = raw == null ? null : String(raw).slice(0, 10);
  const ok = got === expected;
  console.log(
    `${ok ? "OK  " : "FAIL"} ${header}\n      SP field ${col.name} = ${raw} (expect ${expected})`,
  );
  if (!ok) failures += 1;
}

// Also verify portal-style columnValues map for all template headers
console.log("\n=== Portal header coverage on this item ===");
let mapped = 0;
for (const header of TEMPLATE_HEADERS) {
  if (header === "Name") {
    mapped += 1;
    continue;
  }
  if (resolve2(header)) mapped += 1;
}
console.log(
  `Resolvable template headers: ${mapped}/${TEMPLATE_HEADERS.length}`,
);

console.log("\nDeleting test item…");
await graph(token, `${root}/lists/${listId}/items/${itemId}`, {
  method: "DELETE",
});

if (failures > 0 || stillMissing.length > 0) {
  console.error(
    `\nRESULT: FAIL (date mismatches=${failures}, missing columns=${stillMissing.length})`,
  );
  process.exit(1);
}

console.log("\nRESULT: PASS — columns aligned, ISO dates round-trip OK.");
