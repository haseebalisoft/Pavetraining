/**
 * Inspect full field bag for Events #29 and compare to a healthy calendar row if any.
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

  const cols = await client.api(`${root}/lists/${listId}/columns`).top(300).get();
  console.log("All column names:");
  console.log((cols.value || []).map((c) => c.name).sort().join(", "));

  const item = await client
    .api(`${root}/lists/${listId}/items/29`)
    .expand("fields")
    .get();

  const keys = Object.keys(item.fields || {}).sort();
  console.log("\nFields on #29:");
  for (const k of keys) {
    const v = item.fields[k];
    if (v === null || v === undefined || v === "") continue;
    console.log(`  ${k}=`, typeof v === "object" ? JSON.stringify(v) : v);
  }

  // Try setting EventType=0 (single occurrence) — classic calendar key
  const candidates = {
    EventType: 0,
    fAllDayEvent: false,
    fRecurrence: false,
  };

  console.log("\nTrying EventType=0 patch…");
  try {
    await client
      .api(`${root}/lists/${listId}/items/29/fields`)
      .patch(candidates);
    console.log("EventType patch OK on #29");
  } catch (error) {
    console.warn("EventType patch failed:", error?.message ?? error);
    if (error?.body) console.warn(error.body);
  }

  // Create a test event via Graph with EventType to see if it appears
  console.log("\nCreating calendar test item for today…");
  const start = new Date();
  start.setUTCHours(13, 0, 0, 0);
  const end = new Date(start.getTime() + 2 * 60 * 60 * 1000);
  try {
    const created = await client.api(`${root}/lists/${listId}/items`).post({
      fields: {
        Title: "CALENDAR VISIBILITY TEST — delete me",
        EventDate: start.toISOString(),
        EndDate: end.toISOString(),
        fAllDayEvent: false,
        fRecurrence: false,
        EventType: 0,
        Customer_x0020_Visible: true,
        Category: "Meeting",
      },
    });
    console.log("Created test item #", created.id);
  } catch (error) {
    console.warn("Create with EventType/Category failed:", error?.message ?? error);
    try {
      const created = await client.api(`${root}/lists/${listId}/items`).post({
        fields: {
          Title: "CALENDAR VISIBILITY TEST — delete me",
          EventDate: start.toISOString(),
          EndDate: end.toISOString(),
          fAllDayEvent: false,
          fRecurrence: false,
          Customer_x0020_Visible: true,
        },
      });
      console.log("Created lean test item #", created.id);
    } catch (error2) {
      console.error("Lean create failed:", error2?.message ?? error2);
    }
  }
}

main().catch((e) => {
  console.error(e?.message ?? e);
  process.exit(1);
});
