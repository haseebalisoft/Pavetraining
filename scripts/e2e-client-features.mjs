/**
 * End-to-end smoke against live SharePoint (app-only Graph).
 *
 * Covers:
 * 1) Workforce lookup → NPORS create → matrix N001 write (sync path)
 * 2) Company logo upload (Graph drive + CompanyLogo thumbnail JSON)
 * 3) Candidate photo upload (Graph drive + Photo thumbnail JSON)
 *
 *   node --env-file=.env.local scripts/e2e-client-features.mjs
 */
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

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
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      ...(init.body instanceof Uint8Array || Buffer.isBuffer(init.body)
        ? {}
        : { "Content-Type": "application/json" }),
      ...(init.headers || {}),
    },
  });
  if (!res.ok) {
    throw new Error(`Graph ${res.status} ${url}: ${await res.text()}`);
  }
  if (res.status === 204) return null;
  const ct = res.headers.get("content-type") || "";
  if (ct.includes("application/json")) return res.json();
  return res.arrayBuffer();
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

function parseThumb(value) {
  if (!value) return null;
  if (typeof value === "string" && value.trim().startsWith("{")) {
    try {
      return JSON.parse(value);
    } catch {
      return null;
    }
  }
  if (typeof value === "object") return value;
  return null;
}

const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64",
);

const results = [];
function pass(name, detail) {
  results.push({ name, ok: true, detail });
  console.log(`PASS  ${name}${detail ? ` — ${detail}` : ""}`);
}
function fail(name, detail) {
  results.push({ name, ok: false, detail });
  console.error(`FAIL  ${name} — ${detail}`);
}

const token = await getToken();
const root = sitePath();
const workforceListId = requireEnv("SHAREPOINT_WORKFORCE_LIST_ID");
const companyListId = requireEnv("SHAREPOINT_COMPANY_LIST_ID");
const nporsListId = requireEnv("SHAREPOINT_NPORS_REGISTER_LIST_ID");
const matrixListId = requireEnv("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID");

console.log("=== E2E client features ===\n");

// ─── 1) Resolve workforce + company ─────────────────────────────────────────
const workforce = await listAll(token, `${root}/lists/${workforceListId}`);
const companies = await listAll(token, `${root}/lists/${companyListId}`);
const companyById = new Map(companies.map((c) => [String(c.id), c]));

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
  candidate = {
    id: item.id,
    name,
    companyName:
      fieldText(hit.fields, "CompanyName", "Title") || `Company #${hit.id}`,
  };
  company = { id: hit.id, name: candidate.companyName };
  break;
}

if (!candidate || !company) {
  fail("resolve-workforce", "No Workforce row with matching Company List");
  process.exit(1);
}
pass(
  "resolve-workforce",
  `${candidate.name} (#${candidate.id}) @ ${company.name} (#${company.id})`,
);

// ─── 2) NPORS create with CandidateNameLookupId (auto-name path) ────────────
let nporsId = null;
try {
  const expiryIso = "2031-06-15T00:00:00Z";
  const created = await graph(token, `${root}/lists/${nporsListId}/items`, {
    method: "POST",
    body: JSON.stringify({
      fields: {
        CandidateNameLookupId: Number(candidate.id),
        CompanyNameLookupId: Number(company.id),
        TrainingOutcome: "Pass",
        Expiry: expiryIso,
        CustomerVisible: true,
        Notes: "E2E portal smoke — delete me",
      },
    }),
  });
  nporsId = created.id;
  const readBack = await graph(
    token,
    `${root}/lists/${nporsListId}/items/${nporsId}?$expand=fields`,
  );
  const lookupOk =
    String(readBack.fields?.CandidateNameLookupId) === String(candidate.id);
  if (!lookupOk) throw new Error("CandidateNameLookupId not persisted");
  pass("npors-create-lookup", `NPORS #${nporsId} linked to workforce #${candidate.id}`);

  // Simulate portal matrix sync write for form category N001
  const matrixCols = await graph(
    token,
    `${root}/lists/${matrixListId}/columns?$top=320`,
  );
  const n001Col = (matrixCols.value || []).find((col) =>
    String(col.displayName || "").startsWith("N001"),
  );
  if (!n001Col?.name) throw new Error("N001 column missing on Training Matrix Update");

  const matrixItems = await listAll(token, `${root}/lists/${matrixListId}`);
  let matrixRow = matrixItems.find((item) => {
    const title = fieldText(item.fields, "Title", "Name");
    return title.toLowerCase() === candidate.name.toLowerCase();
  });

  if (!matrixRow) {
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
        body: JSON.stringify({ [n001Col.name]: expiryIso }),
      },
    );
  }

  const matrixCheck = await graph(
    token,
    `${root}/lists/${matrixListId}/items/${matrixRow.id}?$expand=fields`,
  );
  const n001Value = String(matrixCheck.fields?.[n001Col.name] || "");
  if (!n001Value.startsWith("2031-06-15")) {
    throw new Error(`Matrix N001 not updated: ${n001Value}`);
  }
  pass(
    "matrix-sync-n001",
    `Matrix #${matrixRow.id} ${n001Col.displayName}=${n001Value.slice(0, 10)}`,
  );
} catch (error) {
  fail(
    "npors-matrix-flow",
    error instanceof Error ? error.message : String(error),
  );
} finally {
  if (nporsId) {
    try {
      await graph(token, `${root}/lists/${nporsListId}/items/${nporsId}`, {
        method: "DELETE",
      });
      pass("npors-cleanup", `Deleted NPORS #${nporsId}`);
    } catch (error) {
      fail(
        "npors-cleanup",
        error instanceof Error ? error.message : String(error),
      );
    }
  }
}

// ─── 3) Company logo upload ─────────────────────────────────────────────────
try {
  const drive = await graph(token, `${root}/drive`);
  const driveId = drive.id;
  const fileName = `${Date.now()}-e2e-logo.png`;
  const path = `PortalMedia/company-${company.id}/${fileName}`;
  const encoded = path
    .split("/")
    .map((p) => encodeURIComponent(p))
    .join("/");
  const uploaded = await graph(
    token,
    `/drives/${driveId}/root:/${encoded}:/content`,
    {
      method: "PUT",
      headers: { "Content-Type": "image/png" },
      body: PNG,
    },
  );
  const serverRelativeUrl = decodeURIComponent(
    String(uploaded.webUrl || "").replace(
      "https://pavetraining.sharepoint.com",
      "",
    ),
  );
  const meta = {
    type: "thumbnail",
    fileName,
    fieldName: "CompanyLogo",
    serverUrl: "https://pavetraining.sharepoint.com",
    serverRelativeUrl,
    id: uploaded.id,
    driveId,
  };
  await graph(
    token,
    `${root}/lists/${companyListId}/items/${company.id}/fields`,
    {
      method: "PATCH",
      body: JSON.stringify({ CompanyLogo: JSON.stringify(meta) }),
    },
  );
  const companyRead = await graph(
    token,
    `${root}/lists/${companyListId}/items/${company.id}?$expand=fields`,
  );
  const thumb = parseThumb(companyRead.fields?.CompanyLogo);
  if (!thumb?.serverRelativeUrl && !thumb?.id) {
    throw new Error("CompanyLogo thumbnail JSON not readable after upload");
  }
  const bytes = await graph(
    token,
    `/drives/${driveId}/items/${uploaded.id}/content`,
  );
  if (!(bytes instanceof ArrayBuffer) || bytes.byteLength < 10) {
    throw new Error("Logo binary could not be downloaded via Graph");
  }
  pass(
    "company-logo-upload",
    `Company #${company.id} logo ${bytes.byteLength} bytes`,
  );
} catch (error) {
  fail(
    "company-logo-upload",
    error instanceof Error ? error.message : String(error),
  );
}

// ─── 4) Candidate photo upload ──────────────────────────────────────────────
try {
  const drive = await graph(token, `${root}/drive`);
  const driveId = drive.id;
  const fileName = `${Date.now()}-e2e-photo.png`;
  const path = `PortalMedia/workforce-${candidate.id}/${fileName}`;
  const encoded = path
    .split("/")
    .map((p) => encodeURIComponent(p))
    .join("/");
  const uploaded = await graph(
    token,
    `/drives/${driveId}/root:/${encoded}:/content`,
    {
      method: "PUT",
      headers: { "Content-Type": "image/png" },
      body: PNG,
    },
  );
  const serverRelativeUrl = decodeURIComponent(
    String(uploaded.webUrl || "").replace(
      "https://pavetraining.sharepoint.com",
      "",
    ),
  );
  const meta = {
    type: "thumbnail",
    fileName,
    fieldName: "Photo",
    serverUrl: "https://pavetraining.sharepoint.com",
    serverRelativeUrl,
    id: uploaded.id,
    driveId,
  };
  await graph(
    token,
    `${root}/lists/${workforceListId}/items/${candidate.id}/fields`,
    {
      method: "PATCH",
      body: JSON.stringify({ Photo: JSON.stringify(meta) }),
    },
  );
  const wfRead = await graph(
    token,
    `${root}/lists/${workforceListId}/items/${candidate.id}?$expand=fields`,
  );
  const thumb = parseThumb(wfRead.fields?.Photo);
  if (!thumb?.serverRelativeUrl && !thumb?.id) {
    throw new Error("Photo thumbnail JSON not readable after upload");
  }
  const bytes = await graph(
    token,
    `/drives/${driveId}/items/${uploaded.id}/content`,
  );
  if (!(bytes instanceof ArrayBuffer) || bytes.byteLength < 10) {
    throw new Error("Photo binary could not be downloaded via Graph");
  }
  pass(
    "candidate-photo-upload",
    `Workforce #${candidate.id} photo ${bytes.byteLength} bytes`,
  );
} catch (error) {
  fail(
    "candidate-photo-upload",
    error instanceof Error ? error.message : String(error),
  );
}

// ─── 5) Matrix column coverage ──────────────────────────────────────────────
try {
  const colsPage = await graph(
    token,
    `${root}/lists/${matrixListId}/columns?$top=320`,
  );
  const cols = colsPage.value || [];
  const nCodes = cols.filter((c) =>
    /^N\d+/i.test(String(c.displayName || "")),
  );
  const hasN001 = nCodes.some((c) =>
    String(c.displayName || "").startsWith("N001"),
  );
  if (!hasN001) throw new Error("N001 missing");
  pass(
    "matrix-category-columns",
    `${nCodes.length} N### display columns on Training Matrix Update`,
  );
} catch (error) {
  fail(
    "matrix-category-columns",
    error instanceof Error ? error.message : String(error),
  );
}

console.log("\n=== Summary ===");
const failed = results.filter((r) => !r.ok);
const passed = results.filter((r) => r.ok);
console.log(`Passed: ${passed.length}`);
console.log(`Failed: ${failed.length}`);
if (failed.length) {
  for (const row of failed) console.log(` - ${row.name}: ${row.detail}`);
  process.exitCode = 1;
} else {
  console.log("All E2E checks passed.");
}
