/**
 * Compare Excel headers vs SharePoint Training Matrix Update columns,
 * and sample Zeeshan's stored field count.
 *
 *   node scripts/diagnose-matrix-column-coverage.mjs
 */
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { resolve } from "node:path";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");
const { ClientSecretCredential } = require("@azure/identity");

function loadEnv() {
  const env = {};
  for (const line of readFileSync(".env.local", "utf8").split(/\r?\n/)) {
    if (!line || line.startsWith("#")) continue;
    const i = line.indexOf("=");
    if (i < 0) continue;
    let v = line.slice(i + 1).trim();
    if (
      (v.startsWith('"') && v.endsWith('"')) ||
      (v.startsWith("'") && v.endsWith("'"))
    ) {
      v = v.slice(1, -1);
    }
    env[line.slice(0, i).trim()] = v;
  }
  return env;
}

function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function headerLookupKey(value) {
  return normalizeHeader(value)
    .toLowerCase()
    .replace(/[–—]/g, "-")
    .replace(/[°º]/g, "");
}
function siteApiRoot(siteIdRaw) {
  let siteId = String(siteIdRaw).replace(/\/+$/, "");
  if (siteId.includes(":/") && !siteId.endsWith(":")) siteId = `${siteId}:`;
  return `/sites/${siteId}`;
}
async function fetchRetry(url, options = {}, attempts = 5) {
  let lastErr;
  for (let i = 1; i <= attempts; i += 1) {
    try {
      return await fetch(url, options);
    } catch (error) {
      lastErr = error;
      await new Promise((r) => setTimeout(r, 400 * 2 ** (i - 1)));
    }
  }
  throw lastErr;
}

const env = loadEnv();
const wb = XLSX.read(readFileSync(resolve("Training matrix example.xlsx")), {
  type: "buffer",
  raw: true,
});
const matrix = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
  header: 1,
  defval: null,
  raw: true,
  blankrows: false,
});
const excelHeaders = matrix[0]
  .map((h) => (h == null ? "" : String(h).trim()))
  .filter(Boolean);

const credential = new ClientSecretCredential(
  env.AZURE_TENANT_ID,
  env.AZURE_CLIENT_ID,
  env.AZURE_CLIENT_SECRET,
);
const token = await credential.getToken(
  "https://graph.microsoft.com/.default",
);
const auth = { Authorization: `Bearer ${token.token}` };
const root = `https://graph.microsoft.com/v1.0${siteApiRoot(env.SHAREPOINT_SITE_ID)}`;
const listId = env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID;

const colMap = new Map(); // lookupKey -> {name, displayName, storage}
let url = `${root}/lists/${listId}/columns?$top=200`;
while (url) {
  const res = await fetchRetry(url, { headers: auth });
  const json = await res.json();
  for (const col of json.value ?? []) {
    if (col.readOnly || !col.name || !col.displayName) continue;
    if (col.name === "ContentType" || col.name === "Attachments") continue;
    const info = {
      name: col.name,
      displayName: col.displayName,
      storage: col.dateTime ? "dateTime" : col.number ? "number" : "other",
    };
    colMap.set(headerLookupKey(col.displayName), info);
  }
  url = json["@odata.nextLink"] ?? "";
}

const matched = [];
const missingInSp = [];
for (const h of excelHeaders) {
  if (/^name$/i.test(h)) continue;
  const hit = colMap.get(headerLookupKey(h));
  if (hit) matched.push({ excel: h, sp: hit.displayName, storage: hit.storage, field: hit.name });
  else missingInSp.push(h);
}

console.log({
  excelHeaders: excelHeaders.length,
  spWritableCols: colMap.size,
  matched: matched.length,
  missingInSp: missingInSp.length,
});
console.log("missing sample:", missingInSp.slice(0, 20));
console.log(
  "storage mix:",
  matched.reduce((a, m) => {
    a[m.storage] = (a[m.storage] || 0) + 1;
    return a;
  }, {}),
);

// Zeeshan row filled non-null field count on SP
let itemsUrl = `${root}/lists/${listId}/items?$expand=fields&$filter=fields/Title eq 'Zeeshan'&$top=5`;
// filter may need Prefer header
const itemsRes = await fetchRetry(itemsUrl, {
  headers: { ...auth, Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" },
});
const itemsJson = await itemsRes.json();
let items = itemsJson.value ?? [];
if (!items.length) {
  // fallback scan
  itemsUrl = `${root}/lists/${listId}/items?$expand=fields&$top=200`;
  while (itemsUrl) {
    const res = await fetchRetry(itemsUrl, { headers: auth });
    const json = await res.json();
    for (const it of json.value ?? []) {
      if (String(it.fields?.Title || "").toLowerCase() === "zeeshan") {
        items.push(it);
      }
    }
    itemsUrl = json["@odata.nextLink"] ?? "";
    if (items.length) break;
  }
}

const item = items[0];
if (!item) {
  console.log("Zeeshan not found in list");
  process.exit(0);
}
const fields = item.fields ?? {};
const fieldNames = new Set(Object.keys(fields));
let nonEmptyMatched = 0;
let emptyMatched = 0;
const emptySample = [];
for (const m of matched) {
  const v = fields[m.field];
  if (v == null || v === "" || v === 0) {
    emptyMatched += 1;
    if (emptySample.length < 12) emptySample.push(m.excel);
  } else {
    nonEmptyMatched += 1;
  }
}
console.log({
  zeeshanId: item.id,
  graphFieldKeys: fieldNames.size,
  matchedNonEmpty: nonEmptyMatched,
  matchedEmptyOrZero: emptyMatched,
  emptySample,
  cscs: fields[matched.find((m) => /cscs/i.test(m.excel))?.field],
  n001: fields[matched.find((m) => /n001/i.test(m.excel))?.field],
});
