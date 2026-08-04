/**
 * Wipe ALL Workforce List rows via Graph.
 *
 * Usage (from repo root):
 *   node --env-file=.env.local scripts/wipe-workforce-list.mjs
 *
 * Optional dry run:
 *   node --env-file=.env.local scripts/wipe-workforce-list.mjs --dry-run
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const DRY_RUN = process.argv.includes("--dry-run");
const CONCURRENCY = 8;

function requireEnv(name) {
  const value = process.env[name]?.trim();
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
  const items = [];
  let url = `${siteRoot()}/lists/${listId}/items?$top=200`;
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

async function mapPool(items, concurrency, mapper) {
  let next = 0;
  const workers = Array.from(
    { length: Math.min(concurrency, items.length || 1) },
    async () => {
      while (true) {
        const i = next;
        next += 1;
        if (i >= items.length) return;
        await mapper(items[i], i);
      }
    },
  );
  await Promise.all(workers);
}

async function main() {
  requireEnv("AZURE_TENANT_ID");
  requireEnv("AZURE_CLIENT_ID");
  requireEnv("AZURE_CLIENT_SECRET");
  requireEnv("SHAREPOINT_SITE_ID");
  const listId = requireEnv("SHAREPOINT_WORKFORCE_LIST_ID");

  const client = getClient();
  console.log(
    DRY_RUN
      ? "DRY RUN — no deletes will be sent."
      : "LIVE WIPE of Workforce List starting…",
  );

  const items = await listAllItems(client, listId);
  console.log(`Workforce rows found: ${items.length}`);

  let deleted = 0;
  let failed = 0;
  const errors = [];

  await mapPool(items, CONCURRENCY, async (item) => {
    try {
      if (!DRY_RUN) {
        await client
          .api(`${siteRoot()}/lists/${listId}/items/${item.id}`)
          .delete();
      }
      deleted += 1;
      if (deleted % 25 === 0 || deleted === items.length) {
        console.log(`Progress: ${deleted}/${items.length}`);
      }
    } catch (error) {
      failed += 1;
      errors.push(`#${item.id}: ${error?.message || String(error)}`);
    }
  });

  const remaining = DRY_RUN ? items.length : (await listAllItems(client, listId)).length;

  console.log("\n========== SUMMARY ==========");
  console.log(`Mode: ${DRY_RUN ? "DRY RUN" : "LIVE"}`);
  console.log(`Deleted: ${deleted}`);
  console.log(`Failed: ${failed}`);
  console.log(`Remaining: ${remaining}`);
  if (errors.length) {
    console.log(`Errors (first 20):`);
    for (const err of errors.slice(0, 20)) console.log(" -", err);
  }
  if (!DRY_RUN && remaining > 0) {
    process.exitCode = 1;
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
