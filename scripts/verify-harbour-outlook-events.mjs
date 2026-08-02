/**
 * Verify Harbour & Hill events exist on the Outlook mailbox without signing in.
 * Usage: node --env-file=.env.local scripts/verify-harbour-outlook-events.mjs
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
  const userId = requireEnv("OUTLOOK_USER_ID").trim();
  const start = new Date().toISOString();
  const end = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();

  console.log(`Mailbox: ${userId}`);
  console.log(`Looking ahead 30 days (${start} → ${end})\n`);

  const res = await client
    .api(`/users/${encodeURIComponent(userId)}/calendarView`)
    .query({ startDateTime: start, endDateTime: end })
    .select("id,subject,start,end,location,categories")
    .orderby("start/dateTime")
    .top(50)
    .header("Prefer", 'outlook.timezone="UTC"')
    .get();

  const all = res.value ?? [];
  const harbour = all.filter((e) =>
    /Harbour|PAVE Training Portal/i.test(
      `${e.subject ?? ""} ${(e.categories ?? []).join(" ")}`,
    ),
  );

  console.log(`Harbour / PAVE calendar events: ${harbour.length}`);
  for (const e of harbour) {
    console.log(`\n• ${e.subject}`);
    console.log(`  start: ${e.start?.dateTime} ${e.start?.timeZone ?? ""}`);
    console.log(`  end:   ${e.end?.dateTime} ${e.end?.timeZone ?? ""}`);
    console.log(`  loc:   ${e.location?.displayName || "—"}`);
  }

  if (harbour.length === 0) {
    console.log("\nNo Harbour matches. First 10 upcoming on that mailbox:");
    for (const e of all.slice(0, 10)) {
      console.log(`- ${e.start?.dateTime}  ${e.subject}`);
    }
  }

  console.log("\n--- SharePoint Events #29–31 ---");
  for (const id of ["29", "30", "31"]) {
    const item = await client
      .api(
        `${siteRoot()}/lists/${requireEnv("SHAREPOINT_EVENTS_LIST_ID")}/items/${id}`,
      )
      .expand("fields")
      .get();
    const f = item.fields ?? {};
    const oid = f.OutlookEventId ? String(f.OutlookEventId) : "";
    console.log(`#${id} ${f.Title}`);
    console.log(
      `  SyncStatus=${f.SyncStatus}  DoNotSync=${f.DoNotSync}  OutlookId=${oid ? oid.slice(0, 36) + "…" : "(none)"}`,
    );
  }
}

main().catch((error) => {
  console.error("Failed:", error?.message ?? error);
  if (error?.body) console.error(error.body);
  process.exit(1);
});
