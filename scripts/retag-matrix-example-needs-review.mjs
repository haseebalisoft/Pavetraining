/**
 * Re-tag Training Matrix Update rows from Training matrix example.xlsx
 * as MatrixLinkStatus = "Needs Review" when they are not Workforce-linked.
 *
 *   node scripts/retag-matrix-example-needs-review.mjs
 */
import { readFileSync, existsSync } from "node:fs";
import { resolve } from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");
const { ClientSecretCredential } = require("@azure/identity");

function loadEnvLocal() {
  const path = resolve(process.cwd(), ".env.local");
  const env = {};
  for (const line of readFileSync(path, "utf8").split(/\r?\n/)) {
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

function nameKey(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
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
      const wait = 500 * 2 ** (i - 1);
      console.warn(`fetch retry ${i}/${attempts} in ${wait}ms:`, error.message);
      await new Promise((r) => setTimeout(r, wait));
    }
  }
  throw lastErr;
}

const env = loadEnvLocal();
const filePath = resolve(process.cwd(), "Training matrix example.xlsx");
if (!existsSync(filePath)) {
  console.error("Missing Training matrix example.xlsx");
  process.exit(1);
}

const wb = XLSX.read(readFileSync(filePath), {
  type: "buffer",
  cellDates: false,
  raw: true,
});
const sheet = wb.Sheets[wb.SheetNames[0]];
const matrix = XLSX.utils.sheet_to_json(sheet, {
  header: 1,
  defval: null,
  raw: true,
  blankrows: false,
});
const headers = (matrix[0] ?? []).map((c) =>
  c == null ? "" : String(c).trim(),
);
const nameIdx = headers.findIndex((h) => /^name$/i.test(h));
if (nameIdx < 0) {
  console.error("No Name column in spreadsheet");
  process.exit(1);
}

const targetNames = new Set();
for (let i = 1; i < matrix.length; i += 1) {
  const name = matrix[i]?.[nameIdx];
  const key = nameKey(name);
  if (key) targetNames.add(key);
}
console.log("Target names from xlsx:", targetNames.size);

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

// Resolve MatrixLinkStatus internal name
const colsUrl = `${root}/lists/${listId}/columns?$top=200`;
const colsRes = await fetchRetry(colsUrl, { headers: auth });
const colsJson = await colsRes.json();
const linkCol = (colsJson.value ?? []).find(
  (c) => String(c.displayName || "").toLowerCase() === "matrixlinkstatus",
);
if (!linkCol?.name) {
  console.error(
    "MatrixLinkStatus column not found on Training Matrix Update list.",
  );
  process.exit(1);
}
console.log("MatrixLinkStatus field:", linkCol.name);

let url = `${root}/lists/${listId}/items?$expand=fields&$top=200`;
const items = [];
while (url) {
  const res = await fetchRetry(url, { headers: auth });
  const json = await res.json();
  if (!res.ok) {
    console.error("List items failed", res.status, json);
    process.exit(1);
  }
  items.push(...(json.value ?? []));
  url = json["@odata.nextLink"] ?? "";
}
console.log("Matrix rows loaded:", items.length);

let updated = 0;
let skippedLinked = 0;
let skippedOther = 0;

for (const item of items) {
  const fields = item.fields ?? {};
  const title = String(fields.Title ?? "").trim();
  const key = nameKey(title);
  if (!targetNames.has(key)) {
    skippedOther += 1;
    continue;
  }
  const wfId = fields.WorkforceItemId ?? fields.workforceItemId;
  if (wfId != null && String(wfId).trim() !== "") {
    skippedLinked += 1;
    continue;
  }
  const current = String(fields[linkCol.name] ?? "").trim();
  if (current === "Needs Review") {
    skippedOther += 1;
    continue;
  }
  const patchUrl = `${root}/lists/${listId}/items/${item.id}/fields`;
  const patchRes = await fetchRetry(patchUrl, {
    method: "PATCH",
    headers: { ...auth, "Content-Type": "application/json" },
    body: JSON.stringify({ [linkCol.name]: "Needs Review" }),
  });
  if (!patchRes.ok) {
    const err = await patchRes.text();
    console.error("PATCH failed", item.id, title, patchRes.status, err);
    continue;
  }
  updated += 1;
  console.log("  tagged Needs Review:", title, `#${item.id}`);
}

console.log("\nDone:", {
  updated,
  skippedLinked,
  skippedOther,
  targetNames: targetNames.size,
});
