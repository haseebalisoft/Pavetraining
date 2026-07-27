/**
 * Wipe ALL Company List rows and related company-scoped data via Graph.
 *
 * Usage (from repo root):
 *   node --env-file=.env.local scripts/wipe-company-list.mjs
 *
 * Optional dry run:
 *   node --env-file=.env.local scripts/wipe-company-list.mjs --dry-run
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const DRY_RUN = process.argv.includes("--dry-run");

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

/** Child lists + Graph field names that may hold the company lookup id. */
const CASCADE = [
  { key: "documents", label: "Customer Documents", fields: ["CompanyLookupId"] },
  {
    key: "trainingMatrix",
    label: "Training Matrix",
    fields: ["MatrixCompanyLookupId", "Company_x0020_NameLookupId"],
  },
  { key: "matrixCategory", label: "Matrix Category Records", fields: ["Company_x0020_NameLookupId"] },
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

function itemCompanyIds(fields, fieldNames) {
  const ids = new Set();
  for (const name of fieldNames) {
    const v = fields?.[name];
    if (v != null && String(v).trim() !== "") ids.add(String(v));
  }
  // Nested lookup object fallback
  for (const key of Object.keys(fields || {})) {
    if (!/LookupId$/i.test(key) && key !== "Company" && key !== "CompanyName") {
      continue;
    }
    if (fieldNames.includes(key)) continue;
  }
  return ids;
}

function matchesCompany(fields, companyId, fieldNames) {
  for (const name of fieldNames) {
    if (String(fields?.[name] ?? "") === String(companyId)) return true;
  }
  // Broad fallbacks used by some lists
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
  console.log(DRY_RUN ? "DRY RUN — no deletes will be sent." : "LIVE WIPE starting…");

  const companies = await listAllItems(client, LISTS.company);
  console.log(`Companies found: ${companies.length}`);

  // Preload child lists once (faster than per-company Graph filters).
  const childCaches = {};
  for (const target of CASCADE) {
    const listId = LISTS[target.key];
    if (!listId) {
      console.log(`Skip ${target.label}: list id not configured`);
      continue;
    }
    try {
      childCaches[target.key] = await listAllItems(client, listId);
      console.log(`Loaded ${target.label}: ${childCaches[target.key].length} rows`);
    } catch (error) {
      console.warn(
        `Could not load ${target.label}:`,
        error?.message || String(error),
      );
      childCaches[target.key] = [];
    }
  }

  // Training Manager Logs (text Company field) — load once.
  if (LISTS.logs) {
    try {
      childCaches.logs = await listAllItems(client, LISTS.logs);
      console.log(`Loaded Training Manager Logs: ${childCaches.logs.length} rows`);
    } catch (error) {
      console.warn("Could not load logs:", error?.message || String(error));
      childCaches.logs = [];
    }
  }

  let relatedDeleted = 0;
  let companiesDeleted = 0;
  const errors = [];

  for (const company of companies) {
    const companyId = String(company.id);
    const name =
      company.fields?.CompanyName ||
      company.fields?.Title ||
      `#${companyId}`;
    console.log(`\n→ Company ${companyId}: ${name}`);

    for (const target of CASCADE) {
      const listId = LISTS[target.key];
      const rows = childCaches[target.key] || [];
      if (!listId) continue;

      const matches = rows.filter((row) =>
        matchesCompany(row.fields || {}, companyId, target.fields),
      );

      for (const row of matches) {
        try {
          await deleteItem(client, listId, row.id);
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
      if (matches.length) {
        console.log(`  ${target.label}: ${matches.length} deleted`);
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
      if (logMatches.length) {
        console.log(`  Training Manager Logs: ${logMatches.length} deleted`);
      }
    }

    try {
      await deleteItem(client, LISTS.company, companyId);
      companiesDeleted += 1;
      console.log(`  Company deleted`);
    } catch (error) {
      errors.push(`Company #${companyId}: ${error?.message || String(error)}`);
      console.error(`  FAILED company delete:`, error?.message || error);
    }
  }

  // Verify remaining
  const remaining = await listAllItems(client, LISTS.company);
  console.log("\n========== SUMMARY ==========");
  console.log(`Mode: ${DRY_RUN ? "DRY RUN" : "LIVE"}`);
  console.log(`Companies deleted: ${companiesDeleted}`);
  console.log(`Related rows deleted: ${relatedDeleted}`);
  console.log(`Companies still remaining: ${remaining.length}`);
  if (errors.length) {
    console.log(`Errors (${errors.length}):`);
    for (const err of errors.slice(0, 20)) console.log(" -", err);
  }
  if (remaining.length > 0) {
    console.log("Remaining company names:");
    for (const row of remaining.slice(0, 30)) {
      console.log(
        ` - #${row.id} ${row.fields?.CompanyName || row.fields?.Title || ""}`,
      );
    }
    process.exitCode = 1;
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
