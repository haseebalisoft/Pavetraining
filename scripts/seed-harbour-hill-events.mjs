/**
 * Creates three customer-visible Events for Harbour & Hill Property Ltd.
 *
 * Usage (from repo root):
 *   node --env-file=.env.local scripts/seed-harbour-hill-events.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const LISTS = {
  company: process.env.SHAREPOINT_COMPANY_LIST_ID,
  events: process.env.SHAREPOINT_EVENTS_LIST_ID,
};

const SITE = process.env.SHAREPOINT_SITE_ID;
const COMPANY_NAME = "Harbour & Hill Property Ltd";
const TRAINING_ADDRESS = "73 Seaview Terrace, Brighton, BN2 5LS";

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing env ${name}`);
  return value;
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

function siteRoot() {
  const siteId = String(SITE).replace(/\/+$/, "");
  if (siteId.includes(":/")) {
    const withTransition = siteId.endsWith(":") ? siteId : `${siteId}:`;
    return `/sites/${withTransition}`;
  }
  return `/sites/${siteId}`;
}

async function listItems(client, listId, filter, top = 200) {
  let request = client
    .api(`${siteRoot()}/lists/${listId}/items`)
    .expand("fields")
    .top(top)
    .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
  if (filter) request = request.filter(filter);
  const res = await request.get();
  return res.value ?? [];
}

async function createItem(client, listId, fields) {
  try {
    const created = await client
      .api(`${siteRoot()}/lists/${listId}/items`)
      .post({ fields });
    return { id: String(created.id), fields: created.fields ?? fields };
  } catch (error) {
    const body =
      error?.body ?? error?.message ?? JSON.stringify(error, null, 2);
    const enriched = new Error(
      `createItem failed: ${typeof body === "string" ? body : JSON.stringify(body)}`,
    );
    enriched.cause = error;
    throw enriched;
  }
}

async function updateItemFields(client, listId, itemId, fields) {
  await client
    .api(`${siteRoot()}/lists/${listId}/items/${itemId}/fields`)
    .patch(fields);
}

function hoursFromNow(hours) {
  return new Date(Date.now() + hours * 60 * 60 * 1000).toISOString();
}

async function findCompanyByName(client, name) {
  // Avoid OData filter with `&` — it can split the query string.
  const all = await listItems(client, LISTS.company, undefined, 5000);
  const match = all.find(
    (item) =>
      String(item.fields?.CompanyName ?? "")
        .trim()
        .toLowerCase() === name.trim().toLowerCase(),
  );
  if (!match) return null;
  return { id: String(match.id), name: match.fields?.CompanyName ?? name };
}

async function ensureEvent(client, row) {
  const all = await listItems(client, LISTS.events, undefined, 2000);
  const existing = all.find(
    (item) =>
      String(item.fields?.Title ?? "").trim().toLowerCase() ===
      row.title.trim().toLowerCase(),
  );
  const fields = {
    Title: row.title,
    EventCompanyLookupId: Number(row.companyId),
    Customer_x0020_Visible: true,
    EventDate: row.eventDate,
    EndDate: row.endDate,
    Location: row.location,
    TrainingAddress: row.trainingAddress,
    Description: row.description,
    // Allow Outlook calendar sync (admin create flow default).
    DoNotSync: "No",
    SyncStatus: "Pending",
    SyncDirection: "SharePointToOutlook",
    LastSyncSource: "SharePoint",
  };
  if (existing) {
    await updateItemFields(client, LISTS.events, existing.id, fields);
    console.log(`  event updated: ${row.title} (#${existing.id})`);
    return String(existing.id);
  }
  try {
    const created = await createItem(client, LISTS.events, fields);
    console.log(`  event created: ${row.title} (#${created.id})`);
    return created.id;
  } catch (error) {
    const msg = String(error?.message ?? error);
    if (/unique constraints/i.test(msg)) {
      throw new Error(
        'SharePoint "Event Company" has Enforce unique values enabled, so only one event can use each company. Turn that off on the Events list Event Company column, then re-run.',
      );
    }
    console.warn(`  event create failed, retrying lean:`, msg);
    const created = await createItem(client, LISTS.events, {
      Title: row.title,
      EventCompanyLookupId: Number(row.companyId),
      Customer_x0020_Visible: true,
      EventDate: row.eventDate,
      EndDate: row.endDate,
      TrainingAddress: row.trainingAddress,
    });
    console.log(`  event created (lean): ${row.title} (#${created.id})`);
    return created.id;
  }
}

async function main() {
  requireEnv("SHAREPOINT_SITE_ID");
  requireEnv("SHAREPOINT_COMPANY_LIST_ID");
  requireEnv("SHAREPOINT_EVENTS_LIST_ID");

  const client = getClient();

  console.log(`Looking up company: ${COMPANY_NAME}`);
  const company = await findCompanyByName(client, COMPANY_NAME);
  if (!company) {
    throw new Error(`Company not found in SharePoint: ${COMPANY_NAME}`);
  }
  console.log(`  found: ${company.name} (#${company.id})`);

  const events = [
    {
      title: "Harbour & Hill — Plant induction",
      location: "Brighton yard",
      description: "Plant department induction / toolbox talk.",
      eventDate: hoursFromNow(24 * 7),
      endDate: hoursFromNow(24 * 7 + 3),
    },
    {
      title: "Harbour & Hill — Transport briefing",
      location: "Brighton yard",
      description: "Transport department safety briefing.",
      eventDate: hoursFromNow(24 * 10),
      endDate: hoursFromNow(24 * 10 + 2),
    },
    {
      title: "Harbour & Hill — Combined Plant & Transport day",
      location: "Brighton yard",
      description: "Joint Plant and Transport training day.",
      eventDate: hoursFromNow(24 * 14),
      endDate: hoursFromNow(24 * 14 + 8),
    },
  ];

  console.log("\nCreating / updating events…");
  const ids = [];
  for (const event of events) {
    const id = await ensureEvent(client, {
      ...event,
      companyId: company.id,
      trainingAddress: TRAINING_ADDRESS,
    });
    ids.push({ title: event.title, id });
  }

  console.log("\nDone.");
  console.log(
    JSON.stringify(
      {
        companyId: company.id,
        companyName: company.name,
        events: ids,
        note: "Customer-visible; DoNotSync=Yes. Supervisor with Department Only (Transport, Plant) sees all company-visible events.",
      },
      null,
      2,
    ),
  );
}

main().catch((error) => {
  console.error("\nFailed:", error?.message ?? error);
  if (error?.body) console.error(error.body);
  process.exit(1);
});
