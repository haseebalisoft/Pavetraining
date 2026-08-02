/**
 * Diagnose why SharePoint Events calendar view may be empty.
 * Usage: node --env-file=.env.local scripts/diagnose-sp-events-calendar.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

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
  let siteId = String(process.env.SHAREPOINT_SITE_ID).replace(/\/+$/, "");
  if (siteId.includes(":/") && !siteId.endsWith(":")) siteId += ":";
  return `/sites/${siteId}`;
}

async function main() {
  const client = getClient();
  const listId = requireEnv("SHAREPOINT_EVENTS_LIST_ID");
  const root = siteRoot();

  const list = await client.api(`${root}/lists/${listId}`).get();
  console.log("List:", list.displayName, "| template:", list.list?.template);

  const cols = await client
    .api(`${root}/lists/${listId}/columns`)
    .top(200)
    .get();
  const interesting = (cols.value || [])
    .filter((c) =>
      /date|event|end|allday|start|title|company|visible/i.test(c.name),
    )
    .map((c) => `${c.name} (${c.displayName}) type=${c.type ?? "?"}`);
  console.log("\nRelevant columns:");
  for (const line of interesting) console.log(" ", line);

  console.log("\nItems #29–31:");
  for (const id of ["29", "30", "31"]) {
    try {
      const item = await client
        .api(`${root}/lists/${listId}/items/${id}`)
        .expand("fields")
        .get();
      const f = item.fields || {};
      console.log(`#${id} ${f.Title}`);
      console.log(`  EventDate=${f.EventDate}`);
      console.log(`  EndDate=${f.EndDate}`);
      console.log(`  fAllDayEvent=${f.fAllDayEvent}`);
      console.log(
        `  EventCompanyLookupId=${f.EventCompanyLookupId} Visible=${f.Customer_x0020_Visible ?? f.CustomerVisible}`,
      );
    } catch (error) {
      console.log(`#${id} ERROR:`, error?.message ?? error);
    }
  }

  const all = await client
    .api(`${root}/lists/${listId}/items`)
    .expand("fields")
    .top(200)
    .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
    .get();

  const harbour = (all.value || []).filter((i) =>
    /Harbour/i.test(String(i.fields?.Title || "")),
  );
  console.log(`\nAll Harbour-titled items: ${harbour.length}`);
  for (const i of harbour) {
    console.log(
      `  #${i.id} ${i.fields.Title} | EventDate=${i.fields.EventDate} | EndDate=${i.fields.EndDate}`,
    );
  }

  // Items with any August 2026 date
  const aug = (all.value || []).filter((i) => {
    const d = String(i.fields?.EventDate || "");
    return d.startsWith("2026-08");
  });
  console.log(`\nAll items with EventDate in Aug 2026: ${aug.length}`);
  for (const i of aug) {
    console.log(`  #${i.id} ${i.fields.Title} | ${i.fields.EventDate}`);
  }

  console.log(`\nTotal items fetched: ${(all.value || []).length}`);
}

main().catch((error) => {
  console.error("Failed:", error?.message ?? error);
  if (error?.body) console.error(error.body);
  process.exit(1);
});
