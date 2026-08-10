/**
 * Wipe Training Matrix SharePoint list rows via Graph.
 * Targets the portal matrix list (EXAMPLE) plus optional legacy/category lists.
 *
 * Usage (from repo root):
 *   node --env-file=.env.local scripts/wipe-training-matrix-list.mjs
 *
 * Optional dry run:
 *   node --env-file=.env.local scripts/wipe-training-matrix-list.mjs --dry-run
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const DRY_RUN = process.argv.includes("--dry-run");
/** Only wipe Training Matrix Update (portal). Skips legacy + category lists. */
const EXAMPLE_ONLY = process.argv.includes("--example-only");
const CONCURRENCY = 3;

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

async function withRetry(label, fn, attempts = 5) {
  let lastErr;
  for (let i = 1; i <= attempts; i += 1) {
    try {
      return await fn();
    } catch (error) {
      lastErr = error;
      const msg = error?.message || String(error);
      const wait = 500 * 2 ** (i - 1);
      console.warn(`[retry ${i}/${attempts}] ${label}: ${msg} → wait ${wait}ms`);
      await new Promise((r) => setTimeout(r, wait));
    }
  }
  throw lastErr;
}

async function listAllItems(client, listId) {
  const items = [];
  let url = `${siteRoot()}/lists/${listId}/items?$top=200`;
  while (url) {
    const res = await withRetry(`list ${listId}`, () =>
      client
        .api(url.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, ""))
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .get(),
    );
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

async function wipeList(client, label, listId) {
  if (!listId) {
    console.log(`\nSkip ${label}: list id not set`);
    return { label, deleted: 0, failed: 0, remaining: 0, skipped: true };
  }

  console.log(
    `\n${DRY_RUN ? "DRY RUN" : "LIVE WIPE"} — ${label} (${listId})`,
  );
  const items = await listAllItems(client, listId);
  console.log(`Rows found: ${items.length}`);

  let deleted = 0;
  let failed = 0;
  const errors = [];

  await mapPool(items, CONCURRENCY, async (item) => {
    try {
      if (!DRY_RUN) {
        await withRetry(`delete #${item.id}`, () =>
          client
            .api(`${siteRoot()}/lists/${listId}/items/${item.id}`)
            .delete(),
        );
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

  if (errors.length) {
    console.log(`Errors (first 20):`);
    for (const err of errors.slice(0, 20)) console.log(" -", err);
  }

  return { label, deleted, failed, remaining, skipped: false };
}

async function main() {
  requireEnv("AZURE_TENANT_ID");
  requireEnv("AZURE_CLIENT_ID");
  requireEnv("AZURE_CLIENT_SECRET");
  requireEnv("SHAREPOINT_SITE_ID");

  const targets = [
    {
      label: "Training Matrix (portal / EXAMPLE)",
      listId: process.env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID?.trim(),
    },
    ...(!EXAMPLE_ONLY
      ? [
          {
            label: "Training Matrix (legacy)",
            listId: process.env.SHAREPOINT_TRAINING_MATRIX_LIST_ID?.trim(),
          },
          {
            label: "Matrix Category Records",
            listId:
              process.env.SHAREPOINT_TRAINING_MATRIX_CATEGORY_RECORDS_LIST_ID?.trim(),
          },
        ]
      : []),
  ];

  if (EXAMPLE_ONLY) {
    console.log("Mode: --example-only (Training Matrix Update list only)");
  }

  if (!targets[0].listId) {
    throw new Error(
      "SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID is required (portal matrix list)",
    );
  }

  const client = getClient();
  const results = [];
  for (const target of targets) {
    results.push(await wipeList(client, target.label, target.listId));
  }

  console.log("\n========== SUMMARY ==========");
  console.log(`Mode: ${DRY_RUN ? "DRY RUN" : "LIVE"}`);
  let hardFail = false;
  for (const r of results) {
    if (r.skipped) {
      console.log(`${r.label}: skipped`);
      continue;
    }
    console.log(
      `${r.label}: deleted=${r.deleted} failed=${r.failed} remaining=${r.remaining}`,
    );
    if (!DRY_RUN && (r.failed > 0 || r.remaining > 0)) hardFail = true;
  }
  if (hardFail) process.exitCode = 1;
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
