/**
 * E2E: cascade-wipe Company List → import Company list.xlsx → verify columns.
 *
 * WARNING: cascade deletes related Workforce / Matrix / registers / docs for
 * those companies (same as Admin company delete).
 *
 * Usage:
 *   node --env-file=.env scripts/test-company-bulk-e2e.mjs
 *   node --env-file=.env scripts/test-company-bulk-e2e.mjs --verify-only
 *   node --env-file=.env scripts/test-company-bulk-e2e.mjs --dry-run
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import * as XLSX from "xlsx";

const DRY_RUN = process.argv.includes("--dry-run");
const VERIFY_ONLY = process.argv.includes("--verify-only");
const EXCEL_PATH = resolve(process.cwd(), "Company list.xlsx");

const LISTS = {
  company: process.env.SHAREPOINT_COMPANY_LIST_ID,
  workforce: process.env.SHAREPOINT_WORKFORCE_LIST_ID,
  trainingMatrix: process.env.SHAREPOINT_TRAINING_MATRIX_LIST_ID,
  npors: process.env.SHAREPOINT_NPORS_REGISTER_LIST_ID,
  eusr: process.env.SHAREPOINT_EUSR_REGISTER_LIST_ID,
  nrswa: process.env.SHAREPOINT_NRSWA_REGISTER_LIST_ID,
  inHouse: process.env.SHAREPOINT_IN_HOUSE_CERTIFICATES_LIST_ID,
  nvq: process.env.SHAREPOINT_NVQ_REGISTER_LIST_ID,
  documents: process.env.SHAREPOINT_CUSTOMER_DOCUMENTS_LIST_ID,
  events: process.env.SHAREPOINT_EVENTS_LIST_ID,
  permissions: process.env.SHAREPOINT_PERMISSIONS_LIST_ID,
  logs: process.env.SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID,
  matrixCategory: process.env.SHAREPOINT_TRAINING_MATRIX_CATEGORY_RECORDS_LIST_ID,
};

const CASCADE = [
  { key: "documents", label: "Customer Documents", fields: ["CompanyLookupId"] },
  {
    key: "trainingMatrix",
    label: "Training Matrix",
    fields: ["MatrixCompanyLookupId", "Company_x0020_NameLookupId"],
  },
  {
    key: "matrixCategory",
    label: "Matrix Category Records",
    fields: ["Company_x0020_NameLookupId"],
  },
  { key: "npors", label: "NPORS", fields: ["CompanyNameLookupId"] },
  { key: "eusr", label: "EUSR", fields: ["CompanyNameLookupId"] },
  { key: "nrswa", label: "Streetworks", fields: ["CompanyNameLookupId"] },
  { key: "inHouse", label: "In-House", fields: ["CompanyNameLookupId"] },
  {
    key: "nvq",
    label: "NVQ",
    fields: ["NVQCompanyLookupId", "Company_x0020_NameLookupId"],
  },
  { key: "events", label: "Events", fields: ["EventCompanyLookupId"] },
  { key: "workforce", label: "Workforce", fields: ["CompanyNameLookupId"] },
  { key: "permissions", label: "Permissions", fields: ["CompanyLookupId"] },
];

/** Clear LookupId then delete — needed when SharePoint blocks child deletes. */
async function clearLookupAndDelete(client, listId, itemId, lookupFields) {
  if (DRY_RUN) return;
  const patch = {};
  for (const field of lookupFields) {
    patch[field] = null;
  }
  if (Object.keys(patch).length) {
    try {
      await client
        .api(`${siteRoot()}/lists/${listId}/items/${itemId}/fields`)
        .patch(patch);
    } catch {
      // continue — delete may still work
    }
  }
  await client.api(`${siteRoot()}/lists/${listId}/items/${itemId}`).delete();
}

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
  const text = String(value).trim();
  return text || null;
}

function normalizeStatus(value) {
  if (!value?.trim()) return "Active";
  const key = value.trim().toLowerCase();
  if (key === "active") return "Active";
  if (key === "inactive") return "Inactive";
  if (key === "on hold" || key === "onhold" || key === "hold") return "Inactive";
  return "Active";
}

function normalizeSize(value) {
  if (!value?.trim()) return null;
  const key = value.trim().toLowerCase();
  if (key === "small") return "Small";
  if (key === "medium") return "Medium";
  if (key === "large") return "Large";
  if (key === "enterprise") return "Enterprise";
  return value.trim();
}

async function listAllItems(client, listId) {
  if (!listId) return [];
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
    companyNumber: raw["Company Number"],
    companyName: raw["Company Name"],
    companySize: normalizeSize(raw["Company Size"]),
    registeredAddress: raw["Registered Address"],
    companyRegNumber: raw["Company Reg Number"],
    vatNo: raw["VAT No"],
    telNo: raw["Tel No"],
    email: raw["Email"],
    mainContact: raw["Main Contact"],
    accountsContactName: raw["Accounts Contact Name"],
    accountsAddress: raw["Accounts address"],
    accountsContactNumber: raw["Accounts Contact number"],
    accountsEmail: raw["Accounts email"],
    notesPricesAgreed: raw["Notes prices agreed"],
    status: normalizeStatus(raw["Status"]),
  };
}

function matchesCompany(fields, companyId, fieldNames) {
  for (const name of fieldNames) {
    if (String(fields?.[name] ?? "") === String(companyId)) return true;
  }
  const fallbacks = [
    "CompanyLookupId",
    "CompanyNameLookupId",
    "EventCompanyLookupId",
    "MatrixCompanyLookupId",
    "NVQCompanyLookupId",
    "Company_x0020_NameLookupId",
  ];
  for (const name of fallbacks) {
    if (String(fields?.[name] ?? "") === String(companyId)) return true;
  }
  return false;
}

async function main() {
  requireEnv("AZURE_TENANT_ID");
  requireEnv("AZURE_CLIENT_ID");
  requireEnv("AZURE_CLIENT_SECRET");
  requireEnv("SHAREPOINT_SITE_ID");
  requireEnv("SHAREPOINT_COMPANY_LIST_ID");

  const client = getClient();
  const { headers, rows: excelRaw } = parseExcel();
  const excelRows = excelRaw.map(mapExcelRow);

  console.log(
    DRY_RUN ? "DRY RUN mode" : VERIFY_ONLY ? "VERIFY ONLY mode" : "LIVE mode",
  );
  console.log(`Excel: ${EXCEL_PATH}`);
  console.log(`Excel headers (${headers.length}): ${headers.join(" | ")}`);
  console.log(`Excel data rows: ${excelRows.length}`);

  if (!VERIFY_ONLY) {
    const companies = await listAllItems(client, LISTS.company);
    console.log(`\nCompanies before wipe: ${companies.length}`);

    const childCaches = {};
    for (const target of CASCADE) {
      const listId = LISTS[target.key];
      if (!listId) continue;
      try {
        childCaches[target.key] = await listAllItems(client, listId);
        console.log(`Loaded ${target.label}: ${childCaches[target.key].length}`);
      } catch (error) {
        console.warn(
          `Could not load ${target.label}:`,
          error?.message || String(error),
        );
        childCaches[target.key] = [];
      }
    }
    if (LISTS.logs) {
      try {
        childCaches.logs = await listAllItems(client, LISTS.logs);
      } catch {
        childCaches.logs = [];
      }
    }

    let relatedDeleted = 0;
    let companiesDeleted = 0;
    const errors = [];

    for (const company of companies) {
      const companyId = String(company.id);
      const name =
        company.fields?.CompanyName || company.fields?.Title || `#${companyId}`;

      for (const target of CASCADE) {
        const listId = LISTS[target.key];
        const rows = childCaches[target.key] || [];
        if (!listId) continue;
        const matches = rows.filter((row) =>
          matchesCompany(row.fields || {}, companyId, target.fields),
        );
        for (const row of matches) {
          try {
            if (target.key === "permissions") {
              await clearLookupAndDelete(client, listId, row.id, [
                "CompanyLookupId",
              ]);
            } else {
              await deleteItem(client, listId, row.id);
            }
            relatedDeleted += 1;
            childCaches[target.key] = childCaches[target.key].filter(
              (r) => r.id !== row.id,
            );
          } catch (error) {
            errors.push(
              `${target.label} #${row.id}: ${error?.message || String(error)}`,
            );
          }
        }
      }

      if (LISTS.logs && name && !String(name).startsWith("#")) {
        const nameLower = String(name).trim().toLowerCase();
        const logMatches = (childCaches.logs || []).filter(
          (row) =>
            String(row.fields?.Company || "")
              .trim()
              .toLowerCase() === nameLower,
        );
        for (const row of logMatches) {
          try {
            await deleteItem(client, LISTS.logs, row.id);
            relatedDeleted += 1;
            childCaches.logs = childCaches.logs.filter((r) => r.id !== row.id);
          } catch (error) {
            errors.push(`Logs #${row.id}: ${error?.message || String(error)}`);
          }
        }
      }

      try {
        await deleteItem(client, LISTS.company, companyId);
        companiesDeleted += 1;
      } catch (error) {
        errors.push(`Company #${companyId}: ${error?.message || String(error)}`);
      }
    }

    console.log(`Companies deleted: ${companiesDeleted}`);
    console.log(`Related rows deleted: ${relatedDeleted}`);
    if (errors.length) {
      console.log(`Delete errors (${errors.length}):`);
      errors.slice(0, 15).forEach((e) => console.log(" -", e));
    }

    const remaining = DRY_RUN
      ? companies
      : await listAllItems(client, LISTS.company);
    console.log(`Companies remaining after wipe: ${remaining.length}`);
    if (!DRY_RUN && remaining.length > 0) {
      throw new Error("Company wipe incomplete — aborting import.");
    }

    // Import
    let imported = 0;
    const importErrors = [];
    for (const [index, row] of excelRows.entries()) {
      const rowNumber = index + 2;
      try {
        if (!row.companyName?.trim()) throw new Error("Company Name missing");
        if (!row.companyNumber?.trim()) throw new Error("Company Number missing");
        if (!row.email?.trim()) throw new Error("Email missing (required)");
        if (!row.registeredAddress?.trim()) {
          throw new Error("Registered Address missing (required)");
        }
        if (!row.companySize?.trim()) {
          throw new Error("Company Size missing (required)");
        }

        const fields = {
          Title: row.companyName,
          CompanyName: row.companyName,
          CompanyNumber: row.companyNumber,
          CompanySize: row.companySize,
          RegisteredAddress: row.registeredAddress,
          CompanyRegNumber: row.companyRegNumber || null,
          VATNo: row.vatNo || null,
          TelNo: row.telNo || null,
          Email: row.email,
          MainContact: row.mainContact || null,
          AccountsContactName: row.accountsContactName || null,
          Accountsaddress: row.accountsAddress || null,
          AccountsContactnumber: row.accountsContactNumber || null,
          Accountsemail: row.accountsEmail || null,
          Notespricesagreed: row.notesPricesAgreed || null,
          Status: row.status || "Active",
        };
        for (const [key, value] of Object.entries(fields)) {
          if (value === null || value === undefined || value === "") {
            delete fields[key];
          }
        }
        await createItem(client, LISTS.company, fields);
        imported += 1;
        if (imported % 5 === 0) {
          console.log(`  imported ${imported}/${excelRows.length}`);
        }
      } catch (error) {
        const message = error?.body
          ? typeof error.body === "string"
            ? error.body
            : JSON.stringify(error.body)
          : error?.message || String(error);
        importErrors.push(`Row ${rowNumber} (${row.companyName}): ${message}`);
        console.error(`  FAIL row ${rowNumber}: ${message}`);
      }
    }

    console.log(`\nImported: ${imported}/${excelRows.length}`);
    if (importErrors.length) {
      console.log(`Import errors (${importErrors.length}):`);
      importErrors.slice(0, 20).forEach((e) => console.log(" -", e));
    }
  }

  if (DRY_RUN) {
    console.log("\nDry run complete — skipped live verify.");
    return;
  }

  const live = await listAllItems(client, LISTS.company);
  console.log(`\nCompanies after import: ${live.length}`);

  const mappedLive = live.map((item) => {
    const f = item.fields || {};
    return {
      id: String(item.id),
      companyNumber: f.CompanyNumber ?? null,
      companyName: f.CompanyName ?? f.Title ?? null,
      companySize: f.CompanySize ?? null,
      registeredAddress: f.RegisteredAddress ?? null,
      companyRegNumber: f.CompanyRegNumber ?? null,
      vatNo: f.VATNo ?? null,
      telNo: f.TelNo ?? null,
      email: f.Email ?? null,
      mainContact: f.MainContact ?? null,
      accountsContactName: f.AccountsContactName ?? null,
      accountsAddress: f.Accountsaddress ?? null,
      accountsContactNumber: f.AccountsContactnumber ?? null,
      accountsEmail: f.Accountsemail ?? null,
      notesPricesAgreed: f.Notespricesagreed ?? null,
      status: f.Status ?? null,
    };
  });

  const compareFields = [
    "companyNumber",
    "companyName",
    "companySize",
    "registeredAddress",
    "companyRegNumber",
    "vatNo",
    "telNo",
    "email",
    "mainContact",
    "accountsContactName",
    "accountsAddress",
    "accountsContactNumber",
    "accountsEmail",
    "notesPricesAgreed",
    "status",
  ];

  let matchedRows = 0;
  const mismatches = [];
  const missing = [];

  for (const expected of excelRows) {
    const actual =
      mappedLive.find(
        (row) =>
          nameKey(row.companyNumber) === nameKey(expected.companyNumber),
      ) ??
      mappedLive.find(
        (row) => nameKey(row.companyName) === nameKey(expected.companyName),
      );

    if (!actual) {
      missing.push(expected.companyName);
      continue;
    }

    let rowOk = true;
    for (const field of compareFields) {
      const exp =
        expected[field] == null || expected[field] === ""
          ? null
          : String(expected[field]).trim();
      const act =
        actual[field] == null || actual[field] === ""
          ? null
          : String(actual[field]).trim();
      if (nameKey(exp ?? "") !== nameKey(act ?? "")) {
        rowOk = false;
        mismatches.push({
          company: expected.companyName,
          field,
          expected: exp,
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
    console.log("\nMissing companies:");
    missing.slice(0, 20).forEach((name) => console.log(" -", name));
  }
  if (mismatches.length) {
    console.log("\nMismatches (first 40):");
    mismatches.slice(0, 40).forEach((row) => {
      console.log(
        ` - ${row.company} / ${row.field}: expected=${JSON.stringify(row.expected)} actual=${JSON.stringify(row.actual)}`,
      );
    });
  }

  if (excelRows[0]) {
    const sampleExpected = excelRows[0];
    const sampleActual =
      mappedLive.find(
        (r) =>
          nameKey(r.companyNumber) === nameKey(sampleExpected.companyNumber),
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
