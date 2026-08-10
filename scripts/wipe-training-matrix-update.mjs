/**
 * Wipe ALL rows from the "Training Matrix Update" list via Graph.
 *
 * Only this one list is touched — Companies, Permissions and Workforce are left
 * alone. To also clear Workforce, run scripts/wipe-workforce-list.mjs.
 *
 * Usage (from repo root):
 *   node --env-file=.env.local scripts/wipe-training-matrix-update.mjs --dry-run
 *   node --env-file=.env.local scripts/wipe-training-matrix-update.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const DRY_RUN = process.argv.includes("--dry-run");
const CONCURRENCY = 8;
const LIST_DISPLAY_NAME = "Training Matrix Update";

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

async function resolveListId(client) {
  const fromEnv = process.env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID?.trim();
  if (fromEnv) return { id: fromEnv, source: "env" };

  const res = await client.api(`${siteRoot()}/lists?$top=200`).get();
  const match = (res.value ?? []).find(
    (list) =>
      String(list.displayName ?? "").trim().toLowerCase() ===
      LIST_DISPLAY_NAME.toLowerCase(),
  );
  if (!match) {
    throw new Error(
      `List "${LIST_DISPLAY_NAME}" not found. Set SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID.`,
    );
  }
  return { id: match.id, source: "display name" };
}

async function listAllItems(client, listId) {
  const items = [];
  let url = `${siteRoot()}/lists/${listId}/items?$top=200&$expand=fields($select=Title)`;
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

  const client = getClient();
  const { id: listId, source } = await resolveListId(client);
  console.log(`List: ${LIST_DISPLAY_NAME} (${listId}) — resolved by ${source}`);
  console.log(
    DRY_RUN
      ? "DRY RUN — no deletes will be sent."
      : `LIVE WIPE of "${LIST_DISPLAY_NAME}" starting…`,
  );

  const items = await listAllItems(client, listId);
  console.log(`Rows found: ${items.length}`);
  if (DRY_RUN) {
    for (const item of items.slice(0, 10)) {
      console.log(` - #${item.id} ${item.fields?.Title ?? "(no Title)"}`);
    }
    if (items.length > 10) console.log(` … and ${items.length - 10} more`);
  }

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

  const remaining = DRY_RUN
    ? items.length
    : (await listAllItems(client, listId)).length;

  console.log("\n========== SUMMARY ==========");
  console.log(`Mode: ${DRY_RUN ? "DRY RUN" : "LIVE"}`);
  console.log(`Deleted: ${deleted}`);
  console.log(`Failed: ${failed}`);
  console.log(`Remaining: ${remaining}`);
  if (errors.length) {
    console.log("Errors (first 20):");
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
