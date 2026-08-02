/**
 * Smoke-test write/read against Training Matrix Update (ISO Date columns).
 *
 *   node --env-file=.env scripts/smoke-training-matrix-update.mjs
 */
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

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
  return res.json();
}

const listId = requireEnv("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID");
const token = await getToken();
const root = sitePath();

console.log("List id:", listId);

const colsPage = await graph(token, `${root}/lists/${listId}/columns?$top=200`);
const cols = colsPage.value || [];
const title = cols.find((c) => c.name === "Title");
const dob = cols.find((c) => (c.displayName || "").trim() === "DOB");
const cscs = cols.find((c) => (c.displayName || "").trim() === "CSCS Expiry");
const n001 = cols.find((c) =>
  String(c.displayName || "").startsWith("N001"),
);

console.log("Title:", title?.name);
console.log("DOB:", dob?.name, dob?.dateTime ? "dateTime" : "NOT date");
console.log("CSCS:", cscs?.name, cscs?.dateTime ? "dateTime" : "NOT date");
console.log("N001:", n001?.name, n001?.dateTime ? "dateTime" : "NOT date");

if (!dob?.dateTime || !cscs?.dateTime) {
  throw new Error("Expected DOB/CSCS to be dateTime columns on Training Matrix Update");
}

const marker = `Portal Smoke ${new Date().toISOString().slice(0, 19)}`;
const fields = {
  Title: marker,
  [dob.name]: "1990-05-15",
  [cscs.name]: "2027-01-31",
};
if (n001?.name) fields[n001.name] = "2026-12-01";

console.log("Creating item with ISO dates…");
const created = await graph(token, `${root}/lists/${listId}/items`, {
  method: "POST",
  body: JSON.stringify({ fields }),
});
const itemId = created.id;
console.log("Created item", itemId);

const read = await graph(
  token,
  `${root}/lists/${listId}/items/${itemId}?$expand=fields`,
);
const f = read.fields || {};
console.log("Read back Title:", f.Title);
console.log("Read back DOB:", f[dob.name]);
console.log("Read back CSCS:", f[cscs.name]);
if (n001?.name) console.log("Read back N001:", f[n001.name]);

const okDob = String(f[dob.name] || "").startsWith("1990-05-15");
const okCscs = String(f[cscs.name] || "").startsWith("2027-01-31");
if (!okDob || !okCscs) {
  throw new Error("ISO date round-trip failed — check column types / write format");
}

console.log("Deleting smoke item…");
await graph(token, `${root}/lists/${listId}/items/${itemId}`, {
  method: "DELETE",
});

console.log("\nSMOKE PASS — Training Matrix Update accepts ISO dates.");

// Compare template headers vs list (from inspect json if present)
try {
  const insp = JSON.parse(
    readFileSync(
      resolve(ROOT, "scripts/training-matrix-update-inspect.json"),
      "utf8",
    ),
  );
  console.log(
    `\nInspect snapshot: ${insp.displayName} writable=${insp.writableCount}`,
  );
} catch {
  /* optional */
}
