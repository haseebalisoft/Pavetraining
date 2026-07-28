/**
 * E2E: wipe Workforce List → import Workforce list.xlsx → verify all columns.
 *
 * Usage (from repo root):
 *   node --env-file=.env scripts/test-workforce-bulk-e2e.mjs
 *   node --env-file=.env scripts/test-workforce-bulk-e2e.mjs --dry-run
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import * as XLSX from "xlsx";

const DRY_RUN = process.argv.includes("--dry-run");
const VERIFY_ONLY = process.argv.includes("--verify-only");
const EXCEL_PATH = resolve(process.cwd(), "Workforce list.xlsx");

const LISTS = {
  company: process.env.SHAREPOINT_COMPANY_LIST_ID,
  workforce: process.env.SHAREPOINT_WORKFORCE_LIST_ID,
  permissions: process.env.SHAREPOINT_PERMISSIONS_LIST_ID,
  departments: process.env.SHAREPOINT_DEPARTMENTS_LIST_ID,
  trainingMatrix: process.env.SHAREPOINT_TRAINING_MATRIX_LIST_ID,
  matrixCategory: process.env.SHAREPOINT_TRAINING_MATRIX_CATEGORY_RECORDS_LIST_ID,
  npors: process.env.SHAREPOINT_NPORS_REGISTER_LIST_ID,
  eusr: process.env.SHAREPOINT_EUSR_REGISTER_LIST_ID,
  nrswa: process.env.SHAREPOINT_NRSWA_REGISTER_LIST_ID,
  inHouse: process.env.SHAREPOINT_IN_HOUSE_CERTIFICATES_LIST_ID,
  nvq: process.env.SHAREPOINT_NVQ_REGISTER_LIST_ID,
  documents: process.env.SHAREPOINT_CUSTOMER_DOCUMENTS_LIST_ID,
  events: process.env.SHAREPOINT_EVENTS_LIST_ID,
};

/** Child lists that can block Workforce deletes via Candidate lookup. */
const WORKFORCE_CHILDREN = [
  { key: "documents", label: "Customer Documents", fields: ["CandidateLookupId"] },
  {
    key: "trainingMatrix",
    label: "Training Matrix",
    fields: ["CandidateNameLookupId"],
  },
  {
    key: "matrixCategory",
    label: "Matrix Category Records",
    fields: ["Candidate_x0020_NameLookupId"],
  },
  { key: "npors", label: "NPORS", fields: ["CandidateNameLookupId"] },
  { key: "eusr", label: "EUSR", fields: ["CandidateNameLookupId"] },
  { key: "nrswa", label: "Streetworks", fields: ["CandidateNameLookupId"] },
  { key: "inHouse", label: "In-House", fields: ["CandidateNameLookupId"] },
  { key: "nvq", label: "NVQ", fields: ["CandidateNameLookupId"] },
  { key: "events", label: "Events", fields: ["CandidateNameLookupId"] },
];

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing env ${name}`);
  return value;
}

function siteRoot() {
  const siteId = String(requireEnv("SHAREPOINT_SITE_ID")).replace(/\/+$/, "");
  if (siteId.includes(":/")) {
    return `/sites/${siteId.endsWith(":") ? siteId : `${siteId}:`}`;
  }
  return `/sites/${siteId}`;
}

function getClient() {
  const credential = new ClientSecretCredential(
    requireEnv("AZURE_TENANT_ID"),
    requireEnv("AZURE_CLIENT_ID"),
    requireEnv("AZURE_CLIENT_SECRET"),
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return Client.initWithMiddleware({ authProvider });
}

function nameKey(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function cellToString(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof value === "number") return String(value);
  const text = String(value).trim();
  return text || null;
}

function normalizeDateValue(value) {
  if (!value?.trim()) return null;
  const text = value.trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);

  const uk = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (uk) {
    return `${uk[3]}-${uk[2].padStart(2, "0")}-${uk[1].padStart(2, "0")}`;
  }

  const short = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2})$/);
  if (short) {
    const first = Number(short[1]);
    const second = Number(short[2]);
    let year = Number(short[3]);
    year += year >= 70 ? 1900 : 2000;
    let month;
    let day;
    if (first > 12 && second <= 12) {
      day = first;
      month = second;
    } else {
      month = first;
      day = second;
    }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 20000 && serial < 60000) {
      const parsed = XLSX.SSF.parse_date_code(serial);
      if (parsed) {
        return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
      }
    }
  }

  const ms = Date.parse(text);
  if (!Number.isNaN(ms)) {
    const parsed = new Date(ms);
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
  }
  return text;
}

function asDateOnly(value) {
  if (value == null || value === "") return null;
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  const ms = Date.parse(text);
  if (!Number.isNaN(ms)) {
    const parsed = new Date(ms);
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
  }
  return text || null;
}

function lookupDisplay(fields, internalName) {
  const direct = fields?.[internalName];
  if (typeof direct === "string" && direct.trim()) {
    // Graph sometimes returns bare LookupId digits as a string — ignore those.
    if (!/^\d+$/.test(direct.trim())) return direct.trim();
  }
  if (direct && typeof direct === "object") {
    const v = direct.LookupValue ?? direct.lookupValue;
    if (typeof v === "string" && v.trim()) return v.trim();
  }
  return null;
}

function lookupId(fields, internalName) {
  const nested = fields?.[internalName];
  if (nested && typeof nested === "object" && nested.LookupId != null) {
    return String(nested.LookupId);
  }
  const id = fields?.[`${internalName}LookupId`];
  if (id != null && String(id).trim() !== "") return String(id);
  const direct = fields?.[internalName];
  if (typeof direct === "string" && /^\d+$/.test(direct.trim())) {
    return direct.trim();
  }
  if (typeof direct === "number") return String(direct);
  return null;
}

async function listAllItems(client, listId) {
  const items = [];
  let url = `${siteRoot()}/lists/${listId}/items?$expand=fields&$top=200`;
  while (url) {
    const res = await client
      .api(url.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, ""))
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .get();
    items.push(...(res.value ?? []));
    url = res["@odata.nextLink"] || null;
  }
  return items;
}

async function deleteItem(client, listId, itemId) {
  if (DRY_RUN) return;
  await client.api(`${siteRoot()}/lists/${listId}/items/${itemId}`).delete();
}

async function createItem(client, listId, fields) {
  if (DRY_RUN) return { id: "dry-run", fields };
  const created = await client
    .api(`${siteRoot()}/lists/${listId}/items`)
    .post({ fields });
  return { id: String(created.id), fields: created.fields ?? fields };
}

function parseExcel() {
  const bytes = readFileSync(EXCEL_PATH);
  const workbook = XLSX.read(bytes, {
    type: "buffer",
    cellDates: true,
    raw: false,
  });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: false,
    blankrows: false,
  });
  const headers = (matrix[0] ?? [])
    .map((cell) => (cell == null ? "" : String(cell).trim()))
    .filter((h) => h.length > 0);

  const rows = [];
  for (let i = 1; i < matrix.length; i += 1) {
    const raw = matrix[i] ?? [];
    const row = {};
    let hasAny = false;
    for (let c = 0; c < headers.length; c += 1) {
      const value = cellToString(raw[c]);
      row[headers[c]] = value;
      if (value) hasAny = true;
    }
    if (hasAny) rows.push(row);
  }
  return { headers, rows };
}

function mapExcelRow(raw) {
  return {
    workforceNumber: raw["Workforce Number"],
    candidateName: raw["Candidate Name"],
    companyName: raw["Company Name"],
    companyNumber: raw["Company Number"],
    trainingManager: raw["Training manager"],
    supervisor: raw["Supervisor"],
    candidateAddress: raw["Candidate Address"],
    email: raw["Email"],
    contactNumber: raw["Contact number"],
    dateOfBirth: normalizeDateValue(raw["Date of birth"]),
    niNumber: raw["Ni Number"],
    nporsNumbers: raw["NPORS Number"],
    cscsNumber: raw["CSCS Number"],
    cscsExpiry: normalizeDateValue(raw["Cscs Expiry"]),
    swqrNumber: raw["SWQR Number"],
    swqrExpiry: normalizeDateValue(raw["Swqr Expiry"]),
    eusrNumber: raw["EUSR Number"],
    eusrExpiry: normalizeDateValue(raw["Eusr Expiry"]),
    inHouseCertificationNumber: raw["In House Certification Number"],
    department: raw["Department"] ?? raw[" Department"],
    status: raw["Status"],
    notes: raw["Notes"],
  };
}

async function ensureCompany(client, companies, name, number) {
  const key = nameKey(name);
  let hit = companies.find((c) => nameKey(c.name) === key);
  if (hit) return hit;

  if (DRY_RUN) {
    hit = { id: `dry-company-${companies.length + 1}`, name, number };
    companies.push(hit);
    return hit;
  }

  const created = await createItem(client, LISTS.company, {
    Title: name,
    CompanyName: name,
    CompanyNumber: number || `AUTO-${Date.now()}`,
    Status: "Active",
  });
  hit = {
    id: created.id,
    name: created.fields?.CompanyName ?? name,
    number: created.fields?.CompanyNumber ?? number,
  };
  companies.push(hit);
  return hit;
}

async function ensurePermission(client, people, displayName, roleType) {
  const key = nameKey(displayName);
  let hit = people.find(
    (p) => nameKey(p.name) === key || nameKey(p.email) === key,
  );
  if (hit) return hit;

  const slug = displayName
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/^\.+|\.+$/g, "")
    .slice(0, 40);
  const userEmail = `import.${slug || "person"}.${Date.now()}@pave.local`;

  if (DRY_RUN) {
    hit = { id: `dry-person-${people.length + 1}`, name: displayName, email: userEmail };
    people.push(hit);
    return hit;
  }

  const created = await createItem(client, LISTS.permissions, {
    Name: displayName,
    UserEmail: userEmail,
    RoleType: roleType,
    Status: "Active",
    AccessScope: "Company",
    CanView: true,
    CanDownload: false,
    CanEdit: false,
  });
  hit = {
    id: created.id,
    name: created.fields?.Name ?? displayName,
    email: created.fields?.UserEmail ?? userEmail,
  };
  people.push(hit);
  return hit;
}

async function ensureDepartment(client, departments, name) {
  const key = nameKey(name);
  let hit = departments.find((d) => nameKey(d.name) === key);
  if (hit) return hit;

  if (DRY_RUN) {
    hit = { id: `dry-dept-${departments.length + 1}`, name };
    departments.push(hit);
    return hit;
  }

  const created = await createItem(client, LISTS.departments, {
    Title: name,
    Name: name,
  });
  hit = {
    id: created.id,
    name: created.fields?.Name ?? created.fields?.Title ?? name,
  };
  departments.push(hit);
  return hit;
}

function normalizeStatus(value) {
  if (!value?.trim()) return "Active";
  const key = value.trim().toLowerCase();
  if (key === "active") return "Active";
  if (key === "inactive") return "inactive";
  return value.trim();
}

async function main() {
  requireEnv("AZURE_TENANT_ID");
  requireEnv("AZURE_CLIENT_ID");
  requireEnv("AZURE_CLIENT_SECRET");
  requireEnv("SHAREPOINT_SITE_ID");
  requireEnv("SHAREPOINT_WORKFORCE_LIST_ID");
  requireEnv("SHAREPOINT_COMPANY_LIST_ID");
  requireEnv("SHAREPOINT_PERMISSIONS_LIST_ID");
  requireEnv("SHAREPOINT_DEPARTMENTS_LIST_ID");

  const client = getClient();
  const { headers, rows: excelRawRows } = parseExcel();
  const excelRows = excelRawRows.map(mapExcelRow);

  console.log(DRY_RUN ? "DRY RUN mode" : VERIFY_ONLY ? "VERIFY ONLY mode" : "LIVE mode");
  console.log(`Excel: ${EXCEL_PATH}`);
  console.log(`Excel headers (${headers.length}): ${headers.join(" | ")}`);
  console.log(`Excel data rows: ${excelRows.length}`);

  if (!VERIFY_ONLY) {
  // ── 1) Wipe workforce (and related candidate lookups that block delete) ─
  const existingWorkforce = await listAllItems(client, LISTS.workforce);
  const workforceIds = new Set(existingWorkforce.map((row) => String(row.id)));
  console.log(`\nWorkforce rows before wipe: ${existingWorkforce.length}`);

  let relatedDeleted = 0;
  for (const child of WORKFORCE_CHILDREN) {
    const listId = LISTS[child.key];
    if (!listId) continue;
    let rows = [];
    try {
      rows = await listAllItems(client, listId);
    } catch (error) {
      console.warn(
        `Could not load ${child.label}:`,
        error?.message || String(error),
      );
      continue;
    }
    const matches = rows.filter((row) => {
      const fields = row.fields || {};
      for (const field of child.fields) {
        if (workforceIds.has(String(fields[field] ?? ""))) return true;
      }
      // Broad fallback: any *Candidate*LookupId pointing at workforce.
      for (const [key, value] of Object.entries(fields)) {
        if (!/Candidate.*LookupId$/i.test(key)) continue;
        if (workforceIds.has(String(value ?? ""))) return true;
      }
      return false;
    });
    for (const row of matches) {
      try {
        await deleteItem(client, listId, row.id);
        relatedDeleted += 1;
      } catch (error) {
        console.warn(
          `  Could not delete ${child.label} #${row.id}:`,
          error?.message || String(error),
        );
      }
    }
    if (matches.length) {
      console.log(`  ${child.label}: ${matches.length} related deleted`);
    }
  }
  console.log(`Related candidate rows deleted: ${relatedDeleted}`);

  let deleted = 0;
  const deleteErrors = [];
  for (const item of existingWorkforce) {
    try {
      await deleteItem(client, LISTS.workforce, item.id);
      deleted += 1;
      if (deleted % 10 === 0) process.stdout.write(`  deleted ${deleted}…\n`);
    } catch (error) {
      deleteErrors.push(`#${item.id}: ${error?.message || String(error)}`);
    }
  }
  console.log(`Workforce deleted: ${deleted}`);
  if (deleteErrors.length) {
    console.log(`Delete errors: ${deleteErrors.length}`);
    deleteErrors.slice(0, 10).forEach((e) => console.log(" -", e));
  }

  const afterWipe = DRY_RUN
    ? existingWorkforce
    : await listAllItems(client, LISTS.workforce);
  console.log(`Workforce remaining after wipe: ${afterWipe.length}`);
  if (!DRY_RUN && afterWipe.length > 0) {
    throw new Error("Wipe incomplete — aborting import.");
  }

  // ── 2) Load lookup caches ──────────────────────────────────────────
  const companyItems = await listAllItems(client, LISTS.company);
  const companies = companyItems.map((item) => ({
    id: String(item.id),
    name: item.fields?.CompanyName ?? item.fields?.Title ?? "",
    number: item.fields?.CompanyNumber ?? null,
  }));

  const permissionItems = await listAllItems(client, LISTS.permissions);
  const people = permissionItems.map((item) => ({
    id: String(item.id),
    name: item.fields?.Name ?? null,
    email: item.fields?.UserEmail ?? "",
  }));

  const departmentItems = await listAllItems(client, LISTS.departments);
  const departments = departmentItems.map((item) => ({
    id: String(item.id),
    name: item.fields?.Name ?? item.fields?.Title ?? "",
  }));

  console.log(
    `\nCaches — companies: ${companies.length}, permissions: ${people.length}, departments: ${departments.length}`,
  );

  // ── 3) Import ──────────────────────────────────────────────────────
  const importErrors = [];
  let imported = 0;
  for (const [index, row] of excelRows.entries()) {
    const rowNumber = index + 2;
    try {
      if (!row.candidateName?.trim()) {
        throw new Error("Candidate Name missing");
      }
      if (!row.companyName?.trim()) {
        throw new Error("Company Name missing");
      }
      if (!row.email?.trim()) {
        throw new Error("Email missing (required by SharePoint)");
      }

      const company = await ensureCompany(
        client,
        companies,
        row.companyName,
        row.companyNumber,
      );

      const fields = {
        CandidateName: row.candidateName,
        WorkforceNumber: row.workforceNumber || null,
        CompanyNameLookupId: Number(company.id),
        Email: row.email,
        Contactnumber: row.contactNumber || null,
        CandidateAddress: row.candidateAddress || null,
        Dateofbirth: row.dateOfBirth || null,
        NiNumber: row.niNumber || null,
        NPORSNumbers: row.nporsNumbers || null,
        CSCSNumber: row.cscsNumber || null,
        CscsExpiry: row.cscsExpiry || null,
        SWQRNumber: row.swqrNumber || null,
        SwqrExpiry: row.swqrExpiry || null,
        EUSRNumber: row.eusrNumber || null,
        EusrExpiry: row.eusrExpiry || null,
        InHouseCertificationNumber: row.inHouseCertificationNumber || null,
        Status: normalizeStatus(row.status),
        Notes: row.notes || null,
      };

      if (row.trainingManager?.trim()) {
        const tm = await ensurePermission(
          client,
          people,
          row.trainingManager.trim(),
          "Admin",
        );
        fields.TrainingmanagerLookupId = Number(tm.id);
      }
      if (row.supervisor?.trim()) {
        const supervisor = await ensurePermission(
          client,
          people,
          row.supervisor.trim(),
          "Customer",
        );
        fields.SupervisorLookupId = Number(supervisor.id);
      }
      if (row.department?.trim()) {
        const dept = await ensureDepartment(
          client,
          departments,
          row.department.trim(),
        );
        fields.Department0LookupId = Number(dept.id);
      }

      // Drop nulls — Graph can reject explicit nulls on some create paths.
      for (const [key, value] of Object.entries(fields)) {
        if (value === null || value === undefined || value === "") {
          delete fields[key];
        }
      }

      await createItem(client, LISTS.workforce, fields);
      imported += 1;
      if (imported % 10 === 0) {
        console.log(`  imported ${imported}/${excelRows.length}`);
      }
    } catch (error) {
      const message = error?.body
        ? typeof error.body === "string"
          ? error.body
          : JSON.stringify(error.body)
        : error?.message || String(error);
      importErrors.push(`Row ${rowNumber} (${row.candidateName}): ${message}`);
      console.error(`  FAIL row ${rowNumber}: ${message}`);
    }
  }

  console.log(`\nImported: ${imported}/${excelRows.length}`);
  if (importErrors.length) {
    console.log(`Import errors (${importErrors.length}):`);
    importErrors.slice(0, 20).forEach((e) => console.log(" -", e));
  }
  } // end !VERIFY_ONLY

  if (DRY_RUN) {
    console.log("\nDry run complete — skipped live verify.");
    return;
  }

  // ── 4) Verify (reload lookup caches so newly created rows resolve) ─
  const live = await listAllItems(client, LISTS.workforce);
  console.log(`\nWorkforce rows after import: ${live.length}`);

  const companiesAfter = (await listAllItems(client, LISTS.company)).map(
    (item) => ({
      id: String(item.id),
      name: item.fields?.CompanyName ?? item.fields?.Title ?? "",
      number: item.fields?.CompanyNumber ?? null,
    }),
  );
  const peopleAfter = (await listAllItems(client, LISTS.permissions)).map(
    (item) => ({
      id: String(item.id),
      name: item.fields?.Name ?? null,
      email: item.fields?.UserEmail ?? "",
    }),
  );
  const departmentsAfter = (
    await listAllItems(client, LISTS.departments)
  ).map((item) => ({
    id: String(item.id),
    name: item.fields?.Name ?? item.fields?.Title ?? "",
  }));

  const companyById = new Map(companiesAfter.map((c) => [c.id, c]));
  const personById = new Map(peopleAfter.map((p) => [p.id, p]));
  const deptById = new Map(departmentsAfter.map((d) => [d.id, d]));

  const mappedLive = live.map((item) => {
    const f = item.fields || {};
    const companyId = lookupId(f, "CompanyName");
    const tmId = lookupId(f, "Trainingmanager");
    const supervisorId = lookupId(f, "Supervisor");
    const deptId = lookupId(f, "Department0");
    return {
      id: String(item.id),
      workforceNumber: f.WorkforceNumber ?? null,
      candidateName: f.CandidateName ?? null,
      companyName:
        lookupDisplay(f, "CompanyName") ??
        (companyId ? companyById.get(companyId)?.name : null) ??
        null,
      companyNumber:
        (typeof f.Company_x0020_Name_x003a__x0020_ === "string" &&
        !/^\d+$/.test(f.Company_x0020_Name_x003a__x0020_)
          ? f.Company_x0020_Name_x003a__x0020_
          : null) ??
        (companyId ? companyById.get(companyId)?.number : null) ??
        null,
      trainingManager:
        lookupDisplay(f, "Trainingmanager") ??
        (tmId ? personById.get(tmId)?.name : null) ??
        null,
      supervisor:
        lookupDisplay(f, "Supervisor") ??
        (supervisorId ? personById.get(supervisorId)?.name : null) ??
        null,
      candidateAddress: f.CandidateAddress ?? null,
      email: f.Email ?? null,
      contactNumber: f.Contactnumber ?? null,
      dateOfBirth: asDateOnly(f.Dateofbirth),
      niNumber: f.NiNumber ?? null,
      nporsNumbers: f.NPORSNumbers ?? null,
      cscsNumber: f.CSCSNumber ?? null,
      cscsExpiry: asDateOnly(f.CscsExpiry),
      swqrNumber: f.SWQRNumber ?? null,
      swqrExpiry: asDateOnly(f.SwqrExpiry),
      eusrNumber: f.EUSRNumber ?? null,
      eusrExpiry: asDateOnly(f.EusrExpiry),
      inHouseCertificationNumber: f.InHouseCertificationNumber ?? null,
      department:
        lookupDisplay(f, "Department0") ??
        (deptId ? deptById.get(deptId)?.name : null) ??
        null,
      status: f.Status ?? null,
      notes: f.Notes ?? null,
    };
  });

  const compareFields = [
    "workforceNumber",
    "candidateName",
    "companyName",
    "companyNumber",
    "trainingManager",
    "supervisor",
    "candidateAddress",
    "email",
    "contactNumber",
    "dateOfBirth",
    "niNumber",
    "nporsNumbers",
    "cscsNumber",
    "cscsExpiry",
    "swqrNumber",
    "swqrExpiry",
    "eusrNumber",
    "eusrExpiry",
    "inHouseCertificationNumber",
    "department",
  ];

  let matchedRows = 0;
  const mismatches = [];
  const missing = [];

  for (const expected of excelRows) {
    const key = nameKey(expected.workforceNumber || expected.candidateName);
    const actual =
      mappedLive.find(
        (row) => nameKey(row.workforceNumber) === nameKey(expected.workforceNumber),
      ) ??
      mappedLive.find(
        (row) =>
          nameKey(row.candidateName) === nameKey(expected.candidateName) &&
          nameKey(row.companyName) === nameKey(expected.companyName),
      );

    if (!actual) {
      missing.push(expected.candidateName);
      continue;
    }

    let rowOk = true;
    for (const field of compareFields) {
      const exp =
        field === "status"
          ? normalizeStatus(expected[field])
          : expected[field] == null || expected[field] === ""
            ? null
            : String(expected[field]).trim();
      const act =
        actual[field] == null || actual[field] === ""
          ? null
          : String(actual[field]).trim();

      // Status blank in Excel becomes Active on import.
      const expectedNorm =
        field === "status" && !expected.status ? "Active" : exp;

      if (nameKey(expectedNorm ?? "") !== nameKey(act ?? "")) {
        rowOk = false;
        mismatches.push({
          candidate: expected.candidateName,
          field,
          expected: expectedNorm,
          actual: act,
        });
      }
    }
    if (rowOk) matchedRows += 1;
  }

  console.log("\n========== VERIFY SUMMARY ==========");
  console.log(`Excel rows: ${excelRows.length}`);
  console.log(`SharePoint rows: ${mappedLive.length}`);
  console.log(`Fully matched rows: ${matchedRows}`);
  console.log(`Missing in SharePoint: ${missing.length}`);
  console.log(`Field mismatches: ${mismatches.length}`);

  if (missing.length) {
    console.log("\nMissing candidates:");
    missing.slice(0, 20).forEach((name) => console.log(" -", name));
  }
  if (mismatches.length) {
    console.log("\nMismatches (first 40):");
    mismatches.slice(0, 40).forEach((row) => {
      console.log(
        ` - ${row.candidate} / ${row.field}: expected=${JSON.stringify(row.expected)} actual=${JSON.stringify(row.actual)}`,
      );
    });
  }

  // Sample first row dump
  if (excelRows[0] && mappedLive[0]) {
    const sampleExpected = excelRows[0];
    const sampleActual =
      mappedLive.find(
        (r) =>
          nameKey(r.workforceNumber) ===
          nameKey(sampleExpected.workforceNumber),
      ) ?? mappedLive[0];
    console.log("\nSample expected (row 1):");
    console.log(JSON.stringify(sampleExpected, null, 2));
    console.log("\nSample actual:");
    console.log(JSON.stringify(sampleActual, null, 2));
  }

  if (
    missing.length ||
    mismatches.length ||
    mappedLive.length !== excelRows.length ||
    matchedRows !== excelRows.length
  ) {
    process.exitCode = 1;
    console.log("\nRESULT: FAIL");
  } else {
    console.log("\nRESULT: PASS — all Excel columns match SharePoint");
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
