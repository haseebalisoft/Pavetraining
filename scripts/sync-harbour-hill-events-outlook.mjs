/**
 * Push Harbour & Hill events #29–31 to Outlook and clear DoNotSync.
 *
 * Usage:
 *   node --env-file=.env.local scripts/sync-harbour-hill-events-outlook.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const EVENT_IDS = ["29", "30", "31"];
const COMPANY_NAME = "Harbour & Hill Property Ltd";

const LISTS = {
  events: process.env.SHAREPOINT_EVENTS_LIST_ID,
};

const SITE = process.env.SHAREPOINT_SITE_ID;

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

function outlookConfig() {
  const groupId = process.env.OUTLOOK_GROUP_ID?.trim();
  const userId = process.env.OUTLOOK_USER_ID?.trim();
  const calendarId = process.env.OUTLOOK_CALENDAR_ID?.trim() || null;

  if (groupId) return { mode: "group", groupId, calendarId };
  if (userId) return { mode: "user", userId, calendarId };
  throw new Error(
    "Outlook calendar is not configured. Set OUTLOOK_GROUP_ID or OUTLOOK_USER_ID.",
  );
}

function outlookEventsPath(config) {
  if (config.mode === "group") {
    if (config.calendarId) {
      return `/groups/${encodeURIComponent(config.groupId)}/calendars/${encodeURIComponent(config.calendarId)}/events`;
    }
    return `/groups/${encodeURIComponent(config.groupId)}/events`;
  }
  if (config.calendarId) {
    return `/users/${encodeURIComponent(config.userId)}/calendars/${encodeURIComponent(config.calendarId)}/events`;
  }
  return `/users/${encodeURIComponent(config.userId)}/events`;
}

function outlookEventItemPath(config, outlookEventId) {
  if (config.mode === "group") {
    return `/groups/${encodeURIComponent(config.groupId)}/events/${encodeURIComponent(outlookEventId)}`;
  }
  return `/users/${encodeURIComponent(config.userId)}/events/${encodeURIComponent(outlookEventId)}`;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function toGraphDateTime(iso) {
  const trimmed = String(iso).trim().replace(/Z$/i, "");
  return {
    dateTime: trimmed.includes(".") ? trimmed.replace(/\.\d+$/, "") : trimmed,
    timeZone: "UTC",
  };
}

function buildOutlookBody(fields) {
  const title = String(fields.Title ?? "PAVE Training Event").trim();
  const description = String(fields.Description ?? "").trim();
  const trainingAddress = String(fields.TrainingAddress ?? "").trim();
  const location = String(fields.Location ?? "").trim();
  const start = fields.EventDate;
  const end = fields.EndDate || fields.EventDate;

  const parts = [];
  if (description) {
    parts.push(`<p>${escapeHtml(description).replace(/\n/g, "<br/>")}</p>`);
  }
  parts.push(`<p><strong>Company:</strong> ${escapeHtml(COMPANY_NAME)}</p>`);
  if (trainingAddress) {
    parts.push(
      `<p><strong>Training address:</strong> ${escapeHtml(trainingAddress)}</p>`,
    );
  }
  parts.push(
    `<p><em>Synced from PAVE Training Portal (SharePoint Events).</em></p>`,
  );

  return {
    subject: `${COMPANY_NAME} — ${title}`,
    body: { contentType: "HTML", content: parts.join("\n") },
    start: toGraphDateTime(start),
    end: toGraphDateTime(end),
    location: location ? { displayName: location } : undefined,
    showAs: "busy",
    categories: ["PAVE Training Portal"],
  };
}

async function patchFields(client, id, fields) {
  await client
    .api(`${siteRoot()}/lists/${LISTS.events}/items/${id}/fields`)
    .patch(fields);
}

async function main() {
  requireEnv("SHAREPOINT_SITE_ID");
  requireEnv("SHAREPOINT_EVENTS_LIST_ID");

  const client = getClient();
  const config = outlookConfig();
  const collection = outlookEventsPath(config);
  console.log(`Outlook target: ${collection}`);
  console.log("(Note: LastSyncedAt is not provisioned on this list — skipping it.)");

  const results = [];

  for (const id of EVENT_IDS) {
    console.log(`\nEvent #${id}…`);
    const item = await client
      .api(`${siteRoot()}/lists/${LISTS.events}/items/${id}`)
      .expand("fields")
      .get();

    const fields = item.fields ?? {};
    const title = fields.Title ?? "(untitled)";
    console.log(`  title: ${title}`);

    if (!fields.EventDate) {
      throw new Error(`Event #${id} has no EventDate`);
    }

    const body = buildOutlookBody(fields);
    let outlookId = fields.OutlookEventId
      ? String(fields.OutlookEventId)
      : null;

    if (outlookId) {
      console.log(`  updating Outlook ${outlookId}`);
      try {
        await client.api(outlookEventItemPath(config, outlookId)).patch(body);
      } catch (error) {
        console.warn(`  update failed, creating new:`, error?.message ?? error);
        const created = await client.api(collection).post(body);
        outlookId = created.id;
        console.log(`  created Outlook: ${outlookId}`);
      }
    } else {
      const created = await client.api(collection).post(body);
      outlookId = created.id;
      console.log(`  created Outlook: ${outlookId}`);
    }

    // Minimal SharePoint write — avoid LastSyncedAt (column missing on live list).
    await patchFields(client, id, {
      DoNotSync: "No",
      OutlookEventId: String(outlookId),
    });
    console.log(`  saved OutlookEventId + DoNotSync=No`);

    try {
      await patchFields(client, id, {
        SyncStatus: "Synced",
        SyncDirection: "SharePointToOutlook",
        LastSyncSource: "SharePoint",
        SyncError: "",
      });
      console.log(`  SyncStatus=Synced`);
    } catch (error) {
      console.warn(
        `  SyncStatus patch skipped:`,
        error?.message ?? error,
      );
    }

    results.push({ sharePointId: id, title, outlookEventId: outlookId });
  }

  console.log("\nDone — check calendar for info@pavetraining.co.uk:");
  console.log(JSON.stringify(results, null, 2));
}

main().catch((error) => {
  console.error("\nFailed:", error?.message ?? error);
  if (error?.body) console.error(error.body);
  process.exit(1);
});
