/**
 * Diagnose / verify customer Events filtering for Murphy plant.
 *
 * Usage:
 *   node --env-file=.env.local scripts/test-customer-events.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

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

function asBoolean(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    return value.toLowerCase() === "true" || value === "1" || value === "Yes";
  }
  if (typeof value === "number") return value === 1;
  return false;
}

function asString(value) {
  if (typeof value === "string" && value.trim()) return value.trim();
  if (typeof value === "number") return String(value);
  return undefined;
}

function resolveEventCompanyId(fields) {
  return (
    asString(fields.EventCompanyLookupId) ??
    asString(fields.EventCompanyId) ??
    (fields.EventCompany && typeof fields.EventCompany === "object"
      ? asString(fields.EventCompany.LookupId)
      : undefined) ??
    null
  );
}

function mapCustomerEvent(id, fields, companyId, companyName) {
  const visible =
    asBoolean(fields.Customer_x0020_Visible) ||
    asBoolean(fields.CustomerVisible);
  if (!visible) return null;

  const eventCompanyId = resolveEventCompanyId(fields);
  const idMatches =
    Boolean(eventCompanyId) &&
    String(eventCompanyId).trim() === String(companyId).trim();
  if (!idMatches) return null;

  const title = asString(fields.Title);
  if (!title) return null;

  return {
    id,
    title,
    eventDate: fields.EventDate ?? null,
    company: companyName,
    eventCompanyId,
  };
}

async function main() {
  const client = getClient();
  const eventsList = requireEnv("SHAREPOINT_EVENTS_LIST_ID");
  const companyList = requireEnv("SHAREPOINT_COMPANY_LIST_ID");
  const root = `${siteRoot()}/lists/${eventsList}/items`;
  const prefer = "HonorNonIndexedQueriesWarningMayFailRandomly";
  const companyId = "27";

  const company = await client
    .api(`${siteRoot()}/lists/${companyList}/items/${companyId}`)
    .expand("fields")
    .get();
  const companyName =
    company.fields?.CompanyName || company.fields?.Title || "Murphy plant";
  console.log("Customer company:", companyId, companyName);

  // Prove broken AND filter (old app behavior).
  try {
    const broken = await client
      .api(root)
      .expand("fields")
      .filter(
        "fields/EventCompanyLookupId eq 27 and fields/Customer_x0020_Visible eq true",
      )
      .top(50)
      .header("Prefer", prefer)
      .get();
    console.log(
      "OLD combined Graph filter count:",
      (broken.value || []).length,
      "(bug if 0 while item 7 is visible+Murphy)",
    );
  } catch (error) {
    console.log("OLD combined Graph filter failed:", error.message);
  }

  // Fixed app behavior: company lookup filter only, visibility in memory.
  const byCompany = await client
    .api(root)
    .expand("fields")
    .filter("fields/EventCompanyLookupId eq 27")
    .top(50)
    .header("Prefer", prefer)
    .get();

  const mapped = (byCompany.value || [])
    .map((item) =>
      mapCustomerEvent(item.id, item.fields || {}, companyId, companyName),
    )
    .filter(Boolean);

  console.log(
    "FIXED path Graph company filter count:",
    (byCompany.value || []).length,
  );
  console.log("FIXED path after visibility map:", mapped.length);
  for (const row of mapped) {
    console.log(" -", row.id, row.title, row.eventDate);
  }

  // Fast must not see Murphy events when mapping with Fast company id against Murphy rows.
  const fastMapped = (byCompany.value || [])
    .map((item) =>
      mapCustomerEvent(item.id, item.fields || {}, "34", "Fast - Zohaib Rashid"),
    )
    .filter(Boolean);
  console.log(
    "Fast company seeing Murphy-filtered rows (expect 0):",
    fastMapped.length,
  );

  if (mapped.length === 0) {
    console.error("FAIL: Murphy customer still gets 0 events after fix path.");
    process.exit(1);
  }
  if (fastMapped.length !== 0) {
    console.error("FAIL: Fast incorrectly matched Murphy company filter rows.");
    process.exit(1);
  }

  console.log("\nPASS: Customer Events path returns Murphy events.");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
