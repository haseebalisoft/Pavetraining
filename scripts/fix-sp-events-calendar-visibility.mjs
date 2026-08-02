/**
 * Fix SharePoint Events calendar visibility for Harbour items.
 * Classic calendar views often ignore Graph-created rows unless Event
 * content type + fAllDayEvent / fRecurrence are set.
 *
 * Usage: node --env-file=.env.local scripts/fix-sp-events-calendar-visibility.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const EVENT_IDS = ["29", "30", "31"];

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

  let contentTypes = [];
  try {
    const res = await client
      .api(`${root}/lists/${listId}/contentTypes`)
      .get();
    contentTypes = res.value || [];
  } catch (error) {
    console.warn("Could not list content types:", error?.message ?? error);
  }

  console.log("Content types:");
  for (const ct of contentTypes) {
    console.log(`  ${ct.id} | ${ct.name}`);
  }

  const eventCt =
    contentTypes.find((ct) => /^event$/i.test(ct.name || "")) ||
    contentTypes.find((ct) => /event/i.test(ct.name || ""));

  for (const id of EVENT_IDS) {
    console.log(`\n#${id}`);
    const item = await client
      .api(`${root}/lists/${listId}/items/${id}`)
      .expand("fields")
      .get();
    const f = item.fields || {};
    console.log(`  Title: ${f.Title}`);
    console.log(`  ContentType: ${f.ContentType}`);
    console.log(`  ContentTypeId: ${f.ContentTypeId}`);
    console.log(`  EventDate: ${f.EventDate}`);
    console.log(`  EndDate: ${f.EndDate}`);
    console.log(`  fAllDayEvent: ${f.fAllDayEvent}`);
    console.log(`  fRecurrence: ${f.fRecurrence}`);

    const patch = {
      fAllDayEvent: false,
      fRecurrence: false,
      // Re-write dates so classic calendar picks them up
      EventDate: f.EventDate,
      EndDate: f.EndDate,
    };
    if (eventCt?.name) {
      patch.ContentType = eventCt.name;
    }

    try {
      await client
        .api(`${root}/lists/${listId}/items/${id}/fields`)
        .patch(patch);
      console.log(`  patched OK:`, patch);
    } catch (error) {
      console.warn(`  full patch failed:`, error?.message ?? error);
      // Lean retry without ContentType
      const lean = {
        fAllDayEvent: false,
        fRecurrence: false,
        EventDate: f.EventDate,
        EndDate: f.EndDate,
      };
      try {
        await client
          .api(`${root}/lists/${listId}/items/${id}/fields`)
          .patch(lean);
        console.log(`  lean patch OK:`, lean);
      } catch (error2) {
        console.error(`  lean patch failed:`, error2?.message ?? error2);
      }
    }

    const refreshed = await client
      .api(`${root}/lists/${listId}/items/${id}`)
      .expand("fields")
      .get();
    const rf = refreshed.fields || {};
    console.log(
      `  after: ContentType=${rf.ContentType} fAllDay=${rf.fAllDayEvent} fRecur=${rf.fRecurrence}`,
    );
  }

  console.log(
    "\nDone. Hard-refresh the SharePoint calendar page (Ctrl+F5).",
  );
  console.log(
    "Also try List ribbon → All Events / Current Events if Calendar is still empty.",
  );
}

main().catch((error) => {
  console.error("Failed:", error?.message ?? error);
  if (error?.body) console.error(error.body);
  process.exit(1);
});
