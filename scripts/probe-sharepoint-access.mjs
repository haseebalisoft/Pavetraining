/**
 * Probe Graph app-only access to key SharePoint lists.
 *   node --env-file=.env.local scripts/probe-sharepoint-access.mjs
 */
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

function requireEnv(name) {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Missing env ${name}`);
  return value;
}

function siteRoot() {
  const siteId = requireEnv("SHAREPOINT_SITE_ID").replace(/\/+$/, "");
  if (siteId.includes(":/")) {
    return `/sites/${siteId.endsWith(":") ? siteId : `${siteId}:`}`;
  }
  return `/sites/${siteId}`;
}

async function main() {
  const credential = new ClientSecretCredential(
    requireEnv("AZURE_TENANT_ID"),
    requireEnv("AZURE_CLIENT_ID"),
    requireEnv("AZURE_CLIENT_SECRET"),
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });
  const client = Client.initWithMiddleware({ authProvider });
  const root = siteRoot();

  const rawSite = requireEnv("SHAREPOINT_SITE_ID");
  console.log("clientId:", process.env.AZURE_CLIENT_ID);
  console.log("siteId len:", rawSite.length);
  console.log("siteId starts:", rawSite.slice(0, 20));
  console.log("siteId ends:", rawSite.slice(-20));
  console.log("siteId has :/", rawSite.includes(":/"));
  console.log("siteRoot:", root);

  try {
    const site = await client.api(root).get();
    console.log("SITE OK:", site.displayName || site.name || site.id);
  } catch (error) {
    console.log(
      "SITE FAIL:",
      error?.statusCode,
      error?.message || String(error),
    );
  }

  const lists = {
    company: process.env.SHAREPOINT_COMPANY_LIST_ID,
    workforce: process.env.SHAREPOINT_WORKFORCE_LIST_ID,
    permissions: process.env.SHAREPOINT_PERMISSIONS_LIST_ID,
    departments: process.env.SHAREPOINT_DEPARTMENTS_LIST_ID,
    events: process.env.SHAREPOINT_EVENTS_LIST_ID,
    trainingManagerLogs:
      process.env.SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID ||
      process.env.SHAREPOINT_AUDIT_LOGS_LIST_ID,
  };

  for (const [name, id] of Object.entries(lists)) {
    if (!id?.trim()) {
      console.log(`${name}: NOT SET`);
      continue;
    }
    try {
      const meta = await client.api(`${root}/lists/${id}`).get();
      console.log(`${name}: META OK (${meta.displayName || meta.name})`);
      try {
        const items = await client
          .api(`${root}/lists/${id}/items`)
          .top(1)
          .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
          .get();
        console.log(
          `${name}: READ OK (sample items=${(items.value || []).length})`,
        );
      } catch (error) {
        console.log(
          `${name}: READ FAIL ${error?.statusCode} ${error?.message || error}`,
        );
      }
    } catch (error) {
      console.log(
        `${name}: META FAIL ${error?.statusCode} ${error?.message || error}`,
      );
    }
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
