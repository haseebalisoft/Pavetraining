/**
 * Ensure the "Training Matrix Update" list has the strong Workforce↔Matrix link
 * columns. Idempotent: only missing columns are created (matched by display
 * name), so it is safe to re-run.
 *
 *   node --env-file=.env.local scripts/ensure-matrix-link-columns.mjs
 *   node --env-file=.env.local scripts/ensure-matrix-link-columns.mjs --dry-run
 *
 * Columns created (display name → type):
 *   WorkforceItemId   Number   (SharePoint Workforce item id — the strong key)
 *   WorkforceNumber   Text     (human Workforce number, e.g. W00001)
 *   CompanyItemId     Number   (Company lookup item id)
 *   CompanyNumber     Text     (human Company number, e.g. C00002)
 *   CandidateName     Text     (candidate name mirror; Title stays authoritative)
 *   MatrixLinkStatus  Choice   (Linked | Orphan | Needs Review)
 *
 * NOTE: creating list columns needs an app role such as Sites.Manage.All or
 * Sites.FullControl.All. If the app only has Sites.ReadWrite.All the POSTs 403;
 * the script then reports exactly which columns + types to add by hand in the
 * SharePoint list settings. Nothing else breaks — the app writes each link
 * column only when it exists, so it degrades gracefully until the columns land.
 */
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DRY_RUN = process.argv.includes("--dry-run");

const EXAMPLE_LIST_DISPLAY_NAME = "Training Matrix Update";
const EXAMPLE_LIST_ENV = "SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID";

/** Display name → Graph column facet. Types the app matches on by display name. */
const LINK_COLUMNS = [
  { displayName: "WorkforceItemId", facet: { number: { decimalPlaces: "none" } } },
  { displayName: "WorkforceNumber", facet: { text: {} } },
  { displayName: "CompanyItemId", facet: { number: { decimalPlaces: "none" } } },
  { displayName: "CompanyNumber", facet: { text: {} } },
  { displayName: "CandidateName", facet: { text: {} } },
  {
    displayName: "MatrixLinkStatus",
    facet: { choice: { choices: ["Linked", "Orphan", "Needs Review"] } },
  },
];

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
    // ignore — env may be provided by --env-file
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

async function resolveListId(token, root) {
  const fromEnv = process.env[EXAMPLE_LIST_ENV]?.trim();
  if (fromEnv) return fromEnv;
  const escaped = EXAMPLE_LIST_DISPLAY_NAME.replace(/'/g, "''");
  const res = await graph(
    token,
    `${root}/lists?$filter=displayName eq '${escaped}'&$select=id,displayName&$top=5`,
  );
  const match = (res.value || []).find(
    (list) =>
      String(list.displayName || "").trim().toLowerCase() ===
      EXAMPLE_LIST_DISPLAY_NAME.toLowerCase(),
  );
  if (!match?.id) {
    throw new Error(
      `Could not resolve list "${EXAMPLE_LIST_DISPLAY_NAME}". Set ${EXAMPLE_LIST_ENV}.`,
    );
  }
  return match.id;
}

const token = await getToken();
const root = sitePath();
const listId = await resolveListId(token, root);

console.log(
  `Ensuring ${LINK_COLUMNS.length} link columns on ${EXAMPLE_LIST_DISPLAY_NAME} (${listId})${
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
const manualFallback = [];

for (const column of LINK_COLUMNS) {
  const key = column.displayName.toLowerCase();
  if (byDisplay.has(key) || byName.has(key)) {
    skipped += 1;
    console.log(`Skip (exists): ${column.displayName}`);
    continue;
  }

  const typeLabel = Object.keys(column.facet)[0];
  console.log(
    `${DRY_RUN ? "Would create" : "Create"}: ${column.displayName} (${typeLabel})`,
  );
  if (DRY_RUN) {
    created += 1;
    continue;
  }

  try {
    await graph(token, `${root}/lists/${listId}/columns`, {
      method: "POST",
      body: JSON.stringify({
        // displayName == internal name: all are alphanumeric, so the app's
        // display-name lookup and SharePoint's internal name line up.
        name: column.displayName,
        displayName: column.displayName,
        ...column.facet,
      }),
    });
    created += 1;
  } catch (error) {
    failed += 1;
    manualFallback.push(`${column.displayName} (${typeLabel})`);
    console.error(
      `  FAIL ${column.displayName}:`,
      error instanceof Error ? error.message : error,
    );
  }
}

console.log(
  `\nDone. created/planned=${created} alreadyPresent=${skipped} failed=${failed}`,
);

if (manualFallback.length) {
  console.log(
    "\nColumn creation failed (likely missing Sites.Manage.All / Sites.FullControl.All).",
  );
  console.log(
    `Add these columns by hand in ${EXAMPLE_LIST_DISPLAY_NAME} → List settings → Create column:`,
  );
  for (const item of manualFallback) console.log(`  • ${item}`);
  console.log(
    "  MatrixLinkStatus choices: Linked, Orphan, Needs Review (Choice column).",
  );
}

if (failed) process.exitCode = 1;
