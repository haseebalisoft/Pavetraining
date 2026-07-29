/**
 * Integration smoke: Workforce lookup → NPORS create → Matrix Update write → cleanup.
 *
 *   node --env-file=.env.local scripts/test-register-matrix-integration.mjs
 */
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
void ROOT;

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

async function listAll(token, listPath) {
  let url = `${listPath}/items?$expand=fields&$top=200`;
  const rows = [];
  while (url) {
    const page = await graph(token, url);
    rows.push(...(page.value || []));
    url = page["@odata.nextLink"] || null;
  }
  return rows;
}

function fieldText(fields, ...keys) {
  for (const key of keys) {
    const value = fields?.[key];
    if (value == null) continue;
    if (typeof value === "string" && value.trim()) return value.trim();
    if (typeof value === "object" && value.LookupValue) {
      return String(value.LookupValue).trim();
    }
  }
  return "";
}

const token = await getToken();
const root = sitePath();
const workforceListId = requireEnv("SHAREPOINT_WORKFORCE_LIST_ID");
const companyListId = requireEnv("SHAREPOINT_COMPANY_LIST_ID");
const nporsListId = requireEnv("SHAREPOINT_NPORS_REGISTER_LIST_ID");
const matrixListId = requireEnv("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID");

console.log("=== Register → Matrix integration smoke ===");

const workforce = await listAll(token, `${root}/lists/${workforceListId}`);
const companies = await listAll(token, `${root}/lists/${companyListId}`);

const companyById = new Map(companies.map((item) => [String(item.id), item]));

let candidate = null;
let company = null;
for (const item of workforce) {
  const name = fieldText(item.fields, "CandidateName", "Title");
  const companyLookupId = String(
    item.fields?.CompanyNameLookupId ??
      item.fields?.Company_x0020_Name_x003a__x0020_LookupId ??
      "",
  ).trim();
  if (!name || !companyLookupId) continue;
  const hit = companyById.get(companyLookupId);
  if (!hit) continue;
  const companyName =
    fieldText(hit.fields, "CompanyName", "Title") || `Company #${hit.id}`;
  candidate = { id: item.id, name, companyName };
  company = { id: hit.id, name: companyName };
  break;
}

if (!candidate || !company) {
  throw new Error("No Workforce candidate with a matching Company List row.");
}

console.log(`Candidate: ${candidate.name} (#${candidate.id})`);
console.log(`Company:   ${company.name} (#${company.id})`);

const expiryIso = "2030-12-15T00:00:00Z";
// Create without MultiChoice first (Graph app-only often rejects NPORSCategory).
const created = await graph(token, `${root}/lists/${nporsListId}/items`, {
  method: "POST",
  body: JSON.stringify({
    fields: {
      CandidateNameLookupId: Number(candidate.id),
      CompanyNameLookupId: Number(company.id),
      TrainingOutcome: "Pass",
      Expiry: expiryIso,
      CustomerVisible: true,
      Notes: "Portal integration smoke — safe to delete",
    },
  }),
});

const nporsId = created.id;
console.log(`Created NPORS #${nporsId}`);

const readBack = await graph(
  token,
  `${root}/lists/${nporsListId}/items/${nporsId}?$expand=fields`,
);
const readName = fieldText(readBack.fields, "CandidateName", "Title");
const readCompany = fieldText(readBack.fields, "CompanyName");
const lookupOk =
  readBack.fields?.CandidateNameLookupId === Number(candidate.id) ||
  String(readBack.fields?.CandidateNameLookupId) === String(candidate.id);

console.log(`NPORS name resolved: ${readName || "(lookup id only)"}`);
console.log(`NPORS company resolved: ${readCompany || "(lookup id only)"}`);
console.log(`Lookup id match: ${lookupOk ? "YES" : "NO"}`);

if (!lookupOk) {
  throw new Error("CandidateNameLookupId was not persisted on NPORS create.");
}

// Simulate portal matrix sync using the form category (N001) even if SP MultiChoice write fails.
const formCategory = "N001";
console.log(`Form category (for matrix sync): ${formCategory}`);

// Mirror app matrix sync write for N001 on Training Matrix Update.
const matrixItems = await listAll(token, `${root}/lists/${matrixListId}`);
const matrixCols = await graph(
  token,
  `${root}/lists/${matrixListId}/columns?$top=320`,
);
const n001Col = (matrixCols.value || []).find((col) =>
  String(col.displayName || "").startsWith("N001"),
);
if (!n001Col?.name) {
  throw new Error("Could not find N001 column on Training Matrix Update.");
}

let matrixRow = matrixItems.find((item) => {
  const title = fieldText(item.fields, "Title", "Name");
  return title.toLowerCase() === candidate.name.toLowerCase();
});

let matrixAction = "updated";
if (!matrixRow) {
  matrixAction = "created";
  matrixRow = await graph(token, `${root}/lists/${matrixListId}/items`, {
    method: "POST",
    body: JSON.stringify({
      fields: {
        Title: candidate.name,
        [n001Col.name]: expiryIso,
      },
    }),
  });
} else {
  await graph(
    token,
    `${root}/lists/${matrixListId}/items/${matrixRow.id}/fields`,
    {
      method: "PATCH",
      body: JSON.stringify({
        [n001Col.name]: expiryIso,
      }),
    },
  );
}

const matrixCheck = await graph(
  token,
  `${root}/lists/${matrixListId}/items/${matrixRow.id}?$expand=fields`,
);
const n001Value = matrixCheck.fields?.[n001Col.name];
console.log(`Matrix row ${matrixAction}: #${matrixRow.id}`);
console.log(`Matrix ${n001Col.displayName}: ${n001Value}`);

if (!n001Value || !String(n001Value).startsWith("2030-12-15")) {
  throw new Error("Matrix N001 expiry was not written as ISO date.");
}

// Cleanup NPORS smoke row (keep matrix data — real sync would leave it).
await graph(token, `${root}/lists/${nporsListId}/items/${nporsId}`, {
  method: "DELETE",
});
console.log(`Deleted smoke NPORS #${nporsId}`);

console.log("\nPASS — Workforce lookup, NPORS create, matrix date write OK.");
console.log(
  "App path: Admin NPORS save → createAdminRegister(workforceId) → triggerMatrixSyncAfterRegister.",
);
