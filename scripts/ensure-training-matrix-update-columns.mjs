/**
 * Ensure Training Matrix Update has every category column from the portal template.
 *
 *   node --env-file=.env.local scripts/ensure-training-matrix-update-columns.mjs
 *   node --env-file=.env.local scripts/ensure-training-matrix-update-columns.mjs --dry-run
 */
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DRY_RUN = process.argv.includes("--dry-run");

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
      ) {
        v = v.slice(1, -1);
      }
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
      ...(init.headers || {}),
    },
  });
  if (!res.ok) {
    throw new Error(`Graph ${res.status} ${url}: ${await res.text()}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

function extractCode(header) {
  const m = String(header).trim().match(/^(N\d+[A-Z]?)\b/i);
  return m ? m[1].toUpperCase() : null;
}

function codeToInternal(code) {
  return `${code}Expiry`;
}

// Load headers from generated TS module via eval of the array export is hard;
// read clientTemplateHeaders by requiring compiled? Use dynamic import of JSON-like
// by parsing the file for "header": lines under CLIENT_MATRIX_CATEGORY_COLUMNS.
function loadCategoryHeaders() {
  const text = readFileSync(
    resolve(ROOT, "src/lib/services/bulkUpload/clientTemplateHeaders.ts"),
    "utf8",
  );
  const headers = [];
  const re = /"header":\s*"([^"]+)"/g;
  let match;
  while ((match = re.exec(text))) {
    const header = match[1];
    if (/^N\d+/i.test(header) || /Expiry$/i.test(header)) {
      headers.push(header);
    }
  }
  // Also meta from CLIENT_MATRIX_META if present as string array
  const meta = [
    "CSCS Expiry",
    "SSSTS Expiry",
    "SMSTS Expiry",
    "NRSWA Expiry",
    "EUSR Expiry",
  ];
  return [...new Set([...meta, ...headers])];
}

const token = await getToken();
const root = sitePath();
const listId = requireEnv("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID");
const wanted = loadCategoryHeaders();

console.log(
  `Ensuring ${wanted.length} columns on Training Matrix Update (${listId})${
    DRY_RUN ? " [dry-run]" : ""
  }`,
);

let cols = [];
let next = `${root}/lists/${listId}/columns?$top=200`;
while (next) {
  const page = await graph(token, next);
  cols.push(...(page.value || []));
  next = page["@odata.nextLink"] || null;
}

const byDisplay = new Map(
  cols.map((c) => [String(c.displayName || "").trim().toLowerCase(), c]),
);
const byName = new Map(cols.map((c) => [String(c.name || "").toLowerCase(), c]));

let created = 0;
let skipped = 0;
let failed = 0;

for (const header of wanted) {
  const key = header.toLowerCase();
  if (byDisplay.has(key)) {
    skipped += 1;
    continue;
  }
  const code = extractCode(header);
  const internal = code
    ? codeToInternal(code)
    : header.replace(/[^A-Za-z0-9]+/g, "").slice(0, 32) || "ExpiryCol";
  if (byName.has(internal.toLowerCase())) {
    skipped += 1;
    continue;
  }

  console.log(`${DRY_RUN ? "Would create" : "Create"}: ${header} → ${internal}`);
  if (DRY_RUN) {
    created += 1;
    continue;
  }
  try {
    await graph(token, `${root}/lists/${listId}/columns`, {
      method: "POST",
      body: JSON.stringify({
        name: internal,
        displayName: header,
        dateTime: {
          format: "dateOnly",
        },
      }),
    });
    created += 1;
  } catch (error) {
    failed += 1;
    console.error(`  FAIL ${header}:`, error instanceof Error ? error.message : error);
  }
}

// ManualOverrides — pipe-separated headers set by admin direct edit.
const manualName = "ManualOverrides";
if (
  !byName.has(manualName.toLowerCase()) &&
  !byDisplay.has("manual overrides")
) {
  console.log(
    `${DRY_RUN ? "Would create" : "Create"}: ManualOverrides (text)`,
  );
  if (!DRY_RUN) {
    try {
      await graph(token, `${root}/lists/${listId}/columns`, {
        method: "POST",
        body: JSON.stringify({
          name: manualName,
          displayName: "Manual Overrides",
          text: {},
        }),
      });
      created += 1;
    } catch (error) {
      failed += 1;
      console.error(
        `  FAIL ManualOverrides:`,
        error instanceof Error ? error.message : error,
      );
    }
  } else {
    created += 1;
  }
} else {
  skipped += 1;
  console.log("ManualOverrides already present.");
}

console.log(
  `\nDone. created/planned=${created} alreadyPresent=${skipped} failed=${failed}`,
);
if (failed) process.exitCode = 1;
