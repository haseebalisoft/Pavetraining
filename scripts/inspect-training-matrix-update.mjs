/**
 * Discover "Training matrix update" SharePoint list without @azure/identity
 * (uses client-credentials token via fetch).
 *
 *   node --env-file=.env scripts/inspect-training-matrix-update.mjs
 */
import { createRequire } from "node:module";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const EXCEL_PATH = resolve(ROOT, "Training matrix example.xlsx");

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
  const tenant = requireEnv("AZURE_TENANT_ID");
  const clientId = requireEnv("AZURE_CLIENT_ID");
  const clientSecret = requireEnv("AZURE_CLIENT_SECRET");
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(
    `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
  );
  if (!res.ok) {
    throw new Error(`Token failed: ${res.status} ${await res.text()}`);
  }
  const json = await res.json();
  return json.access_token;
}

async function graph(token, path) {
  const url = path.startsWith("http")
    ? path
    : `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    throw new Error(`Graph ${res.status} ${url}: ${await res.text()}`);
  }
  return res.json();
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

function excelHeaders(path) {
  const wb = XLSX.readFile(path, { cellDates: true, raw: false });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: null,
    raw: false,
    blankrows: true,
  });
  let headerIdx = 0;
  for (let i = 0; i < Math.min(rows.length, 15); i += 1) {
    const row = (rows[i] || []).map(normalize).filter(Boolean);
    if (
      row.includes("Name") ||
      row.includes("DOB") ||
      row.some((h) => /Expiry/i.test(h))
    ) {
      headerIdx = i;
      break;
    }
  }
  return (rows[headerIdx] || []).map(normalize).filter(Boolean);
}

function columnType(col) {
  if (col.dateTime) return "dateTime";
  if (col.text) return "text";
  if (col.number) return "number";
  if (col.boolean) return "boolean";
  if (col.choice) return "choice";
  if (col.lookup) return "lookup";
  if (col.calculated) return "calculated";
  return "other";
}

const token = await getToken();
const root = sitePath();

const lists = await listAll(
  token,
  `${root}/lists?$select=id,displayName,name,list&$top=200`,
);

const matrixLists = lists.filter(
  (l) =>
    /training\s*matrix/i.test(l.displayName || "") ||
    /training\s*matrix/i.test(l.name || ""),
);

console.log("=== Training Matrix lists ===");
for (const l of matrixLists) {
  console.log(
    `- ${l.displayName} | id=${l.id} | template=${l.list?.template ?? "?"}`,
  );
}

const target =
  matrixLists.find((l) =>
    /training\s*matrix\s*update/i.test(l.displayName || ""),
  ) || matrixLists.find((l) => /update/i.test(l.displayName || ""));

if (!target) {
  console.error("\nCould not find list named 'Training matrix update'.");
  process.exit(1);
}

console.log(`\nTARGET: ${target.displayName}`);
console.log(`LIST_ID=${target.id}`);

const cols = await listAll(token, `${root}/lists/${target.id}/columns?$top=200`);
const writable = cols.filter(
  (c) =>
    !c.readOnly &&
    c.name !== "ContentType" &&
    c.name !== "Attachments" &&
    c.name !== "Edit" &&
    c.name !== "LinkTitleNoMenu" &&
    c.name !== "LinkTitle",
);

console.log(`\nWritable columns: ${writable.length}`);
for (const c of writable) {
  console.log(
    `  [${columnType(c)}] display="${normalize(c.displayName)}" name="${c.name}"`,
  );
}

const excel = excelHeaders(EXCEL_PATH);
console.log(`\nExcel headers (${EXCEL_PATH}): ${excel.length}`);

const byDisplay = new Map(
  writable.map((c) => [normalize(c.displayName), c]),
);
const titleCol = writable.find((c) => c.name === "Title");

const missingInSp = [];
for (const header of excel) {
  if (header === "Name") {
    if (!byDisplay.has("Name") && !titleCol) missingInSp.push(header);
    continue;
  }
  if (!byDisplay.has(header)) missingInSp.push(header);
}

const excelSet = new Set(excel);
const extraInSp = [];
for (const c of writable) {
  const display = normalize(c.displayName);
  if (c.name === "Title" && (excelSet.has("Name") || excelSet.has("Title"))) {
    continue;
  }
  if (!excelSet.has(display) && display !== "Title") {
    extraInSp.push(`${display} (${c.name}, ${columnType(c)})`);
  }
}

console.log("\n=== Missing in SharePoint (present in Excel) ===");
console.log(missingInSp.length ? missingInSp.join("\n") : "(none)");

console.log("\n=== Extra in SharePoint (not in Excel) ===");
console.log(extraInSp.length ? extraInSp.join("\n") : "(none)");

const dateIssues = [];
for (const c of writable) {
  const display = normalize(c.displayName);
  if (
    /expiry/i.test(display) ||
    display === "DOB" ||
    /date of birth/i.test(display)
  ) {
    if (!c.dateTime) {
      dateIssues.push(
        `${display} is ${columnType(c)} (name=${c.name}) — expected dateTime`,
      );
    }
  }
}

console.log("\n=== Date type check (Expiry / DOB) ===");
console.log(
  dateIssues.length ? dateIssues.join("\n") : "(all look like dateTime)",
);

const items = await graph(
  token,
  `${root}/lists/${target.id}/items?$expand=fields&$top=3`,
);
console.log(`\nSample items on page: ${items.value?.length ?? 0}`);
if (items.value?.[0]?.fields) {
  const keys = Object.keys(items.value[0].fields).filter(
    (k) => !k.startsWith("@"),
  );
  console.log("Sample field keys:", keys.slice(0, 40).join(", "));
}

console.log("\n=== .env suggestion ===");
console.log(
  `SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID=${target.id}  # Training matrix update`,
);
console.log(
  `# previous example list: ${process.env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID || "(unset)"}`,
);

// Write machine-readable summary for follow-up
const summary = {
  listId: target.id,
  displayName: target.displayName,
  writableCount: writable.length,
  excelHeaderCount: excel.length,
  missingInSp,
  extraInSp,
  dateIssues,
  columns: writable.map((c) => ({
    displayName: normalize(c.displayName),
    name: c.name,
    type: columnType(c),
  })),
  excel,
};
const outPath = resolve(ROOT, "scripts/training-matrix-update-inspect.json");
require("fs").writeFileSync(outPath, JSON.stringify(summary, null, 2));
console.log(`\nWrote ${outPath}`);
