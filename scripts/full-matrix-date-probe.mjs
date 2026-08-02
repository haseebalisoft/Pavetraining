/**
 * Upload ISO dates across all Training Matrix Update date columns and
 * verify live fetch maps back to Excel/admin header names.
 *
 *   node --env-file=.env scripts/full-matrix-date-probe.mjs
 */
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { resolve as pathResolve } from "node:path";

const require = createRequire(import.meta.url);
const ROOT = process.cwd();

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
  if (!res.ok) throw new Error(`${res.status} ${await res.text()}`);
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

const src = readFileSync(
  pathResolve(ROOT, "src/lib/services/bulkUpload/clientTemplateHeaders.ts"),
  "utf8",
);
const meta = [
  ...src
    .match(/export const CLIENT_MATRIX_META_HEADERS[\s\S]*?\] as const;/)[0]
    .matchAll(/"([^"]+)"/g),
].map((m) => m[1]);
const cats = [...src.matchAll(/"header":\s*"([^"]+)"/g)].map((m) => m[1]);
const TEMPLATE = [...meta, ...cats];

const token = await getToken();
const root = sitePath();
const listId = requireEnv("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID");
const cols = (
  await listAll(token, `${root}/lists/${listId}/columns?$top=200`)
).filter(
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
  cols.map((c) => [normalize(c.displayName).toLowerCase(), c]),
);
const byCollapsed = new Map(cols.map((c) => [collapse(c.displayName), c]));
const byCode = new Map();
for (const c of cols) {
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

function resolveHeader(header) {
  if (header === "Name") return cols.find((c) => c.name === "Title") || null;
  return (
    byDisplay.get(normalize(header).toLowerCase()) ||
    byCollapsed.get(collapse(header)) ||
    (extractCode(header) ? byCode.get(extractCode(header)) : null) ||
    null
  );
}

const unresolved = TEMPLATE.filter((h) => h !== "Name" && !resolveHeader(h));
console.log(
  "Unresolved headers:",
  unresolved.length ? unresolved.join(", ") : "(none)",
);
console.log(
  "Coverage:",
  TEMPLATE.length - unresolved.length,
  "/",
  TEMPLATE.length,
);

const marker = `Full Matrix Probe ${new Date().toISOString().slice(0, 16)}`;
const fields = { Title: marker };
const expected = {};
let i = 0;
for (const header of TEMPLATE) {
  if (header === "Name") continue;
  const col = resolveHeader(header);
  if (!col?.dateTime) continue;
  const day = String((i % 28) + 1).padStart(2, "0");
  const month = String((i % 12) + 1).padStart(2, "0");
  const iso = `2030-${month}-${day}`;
  fields[col.name] = iso;
  expected[header] = { name: col.name, iso };
  i += 1;
}
console.log("Writing date fields:", Object.keys(expected).length);

const created = await graph(token, `${root}/lists/${listId}/items`, {
  method: "POST",
  body: JSON.stringify({ fields }),
});
const itemId = created.id;
const read = await graph(
  token,
  `${root}/lists/${listId}/items/${itemId}?$expand=fields`,
);
const f = read.fields || {};

const columnValues = { Name: f.Title };
let fail = 0;
let ok = 0;
for (const [header, metaInfo] of Object.entries(expected)) {
  const raw = f[metaInfo.name];
  const got = raw == null ? null : String(raw).slice(0, 10);
  columnValues[header] = got;
  if (got !== metaInfo.iso) {
    fail += 1;
    console.log("FAIL", header, "got", got, "expected", metaInfo.iso);
  } else {
    ok += 1;
  }
}

console.log(`Round-trip OK=${ok} FAIL=${fail}`);
const sample = [
  "DOB",
  "CSCS Expiry",
  "N001 - Ind FLT",
  "N020 - Tiltrotator System",
  "N114 - Overhead Container Gantry Crane",
  "N202 - Excavator 360°",
  "N216 - Road Planer",
];
console.log(
  "Admin sample:",
  sample.map((h) => `${h}=${columnValues[h]}`).join(" | "),
);

await graph(token, `${root}/lists/${listId}/items/${itemId}`, {
  method: "DELETE",
});

if (fail || unresolved.length) {
  process.exit(1);
}
console.log(
  "\nPASS — all Excel/admin headers resolve; ISO dates round-trip cleanly.",
);
