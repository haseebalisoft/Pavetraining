/**
 * Wipe operational SharePoint list items. Keeps:
 *   - Permissions row for wayne.curry@pavetraining.co.uk (hardcoded admin)
 *   - NPORS Categories + Training Course Categories (form catalogs)
 *
 * Usage:
 *   node --env-file=.env.local scripts/wipe-all-sharepoint-data.mjs --dry-run
 *   node --env-file=.env.local scripts/wipe-all-sharepoint-data.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const DRY_RUN = process.argv.includes("--dry-run");
const CONCURRENCY = 6;
const KEEP_PERMISSION_EMAILS = new Set(["wayne.curry@pavetraining.co.uk"]);

function requireEnv(name) {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Missing env ${name}`);
  return value;
}

function envId(name) {
  return process.env[name]?.trim() || "";
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

function errorMessage(error) {
  return error?.message || String(error);
}

function isGone(error) {
  const msg = errorMessage(error).toLowerCase();
  return (
    msg.includes("not found") ||
    msg.includes("item does not exist") ||
    /\b404\b/.test(msg)
  );
}

function isRetryable(error) {
  if (isGone(error)) return false;
  const msg = errorMessage(error).toLowerCase();
  return (
    msg.includes("throttl") ||
    msg.includes("too many") ||
    msg.includes("timeout") ||
    msg.includes("429") ||
    msg.includes("503") ||
    msg.includes("504")
  );
}

async function withRetry(label, fn, attempts = 5) {
  let lastErr;
  for (let i = 1; i <= attempts; i += 1) {
    try {
      return await fn();
    } catch (error) {
      lastErr = error;
      if (isGone(error)) return { gone: true };
      if (!isRetryable(error) || i === attempts) throw error;
      const wait = 400 * 2 ** (i - 1);
      console.warn(
        `[retry ${i}/${attempts}] ${label}: ${errorMessage(error)} → wait ${wait}ms`,
      );
      await new Promise((r) => setTimeout(r, wait));
    }
  }
  throw lastErr;
}

async function listAllItems(client, listId, expandFields = false) {
  const items = [];
  let url = `${siteRoot()}/lists/${listId}/items?$top=200${
    expandFields ? "&$expand=fields" : ""
  }`;
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
    { length: Math.min(concurrency, Math.max(items.length, 1)) },
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

async function wipeList(client, label, listId, options = {}) {
  if (!listId) {
    console.log(`\nSkip ${label}: list id not set`);
    return { label, found: 0, deleted: 0, failed: 0, remaining: 0, skipped: true };
  }

  console.log(`\n${DRY_RUN ? "DRY RUN" : "LIVE WIPE"} — ${label}`);
  let items = await listAllItems(
    client,
    listId,
    Boolean(options.keep) || Boolean(options.filesBeforeFolders),
  );
  if (typeof options.keep === "function") {
    items = items.filter((item) => !options.keep(item));
  }
  if (options.filesBeforeFolders) {
    items = [...items].sort((a, b) => {
      const aFolder = Number(a.fields?.FSObjType ?? 0) === 1 ? 1 : 0;
      const bFolder = Number(b.fields?.FSObjType ?? 0) === 1 ? 1 : 0;
      if (aFolder !== bFolder) return aFolder - bFolder;
      const aPath = String(a.fields?.FileDirRef ?? a.fields?.FileRef ?? "");
      const bPath = String(b.fields?.FileDirRef ?? b.fields?.FileRef ?? "");
      return bPath.length - aPath.length;
    });
  } else if (typeof options.sort === "function") {
    items = [...items].sort(options.sort);
  }
  console.log(`Rows to delete: ${items.length}`);

  let deleted = 0;
  let failed = 0;
  const errors = [];

  await mapPool(items, CONCURRENCY, async (item) => {
    try {
      if (!DRY_RUN) {
        await withRetry(`delete ${label} #${item.id}`, () =>
          client.api(`${siteRoot()}/lists/${listId}/items/${item.id}`).delete(),
        );
      }
      deleted += 1;
      if (deleted % 50 === 0 || deleted === items.length) {
        console.log(`  ${label}: ${deleted}/${items.length}`);
      }
    } catch (error) {
      failed += 1;
      errors.push(`#${item.id}: ${error?.message || String(error)}`);
    }
  });

  const remainingItems = DRY_RUN
    ? items
    : await listAllItems(client, listId, Boolean(options.keep));
  const remaining = typeof options.keep === "function"
    ? remainingItems.filter((item) => !options.keep(item)).length
    : remainingItems.length;

  if (errors.length) {
    console.log(`  Errors (${errors.length}), first 10:`);
    for (const err of errors.slice(0, 10)) console.log(`   - ${err}`);
  }
  console.log(`  deleted=${deleted} failed=${failed} remaining=${remaining}`);
  return { label, found: items.length, deleted, failed, remaining, skipped: false };
}

async function unlinkKeptPermissionsCompany(client, listId) {
  if (DRY_RUN || !listId) return;
  const items = await listAllItems(client, listId, true);
  for (const item of items) {
    const email = String(item.fields?.UserEmail ?? "")
      .trim()
      .toLowerCase();
    if (!KEEP_PERMISSION_EMAILS.has(email)) continue;
    const companyId = item.fields?.CompanyLookupId;
    if (companyId == null || String(companyId).trim() === "") continue;
    try {
      await client.api(`${siteRoot()}/lists/${listId}/items/${item.id}/fields`).patch({
        CompanyLookupId: null,
      });
      console.log(`Unlinked company lookup on kept permission #${item.id} (${email})`);
    } catch (error) {
      console.warn(
        `Could not unlink company on permission #${item.id}:`,
        error?.message || String(error),
      );
    }
  }
}

function keepPermission(item) {
  const email = String(item.fields?.UserEmail ?? "")
    .trim()
    .toLowerCase();
  return KEEP_PERMISSION_EMAILS.has(email);
}

async function main() {
  requireEnv("AZURE_TENANT_ID");
  requireEnv("AZURE_CLIENT_ID");
  requireEnv("AZURE_CLIENT_SECRET");
  requireEnv("SHAREPOINT_SITE_ID");

  const client = getClient();
  console.log(
    DRY_RUN
      ? "DRY RUN — counting only, no deletes."
      : "LIVE WIPE of SharePoint operational data.",
  );
  console.log(
    `Keeping Permissions: ${[...KEEP_PERMISSION_EMAILS].join(", ")}`,
  );
  console.log("Keeping catalogs: NPORS Categories, Training Course Categories");

  const permissionsId = envId("SHAREPOINT_PERMISSIONS_LIST_ID");

  // Child / transactional lists first so Restrict Delete does not block.
  const order = [
    [
      "Customer Documents",
      envId("SHAREPOINT_CUSTOMER_DOCUMENTS_LIST_ID"),
      { filesBeforeFolders: true },
    ],
    ["Training Matrix (legacy)", envId("SHAREPOINT_TRAINING_MATRIX_LIST_ID")],
    [
      "Training Matrix Category Records",
      envId("SHAREPOINT_TRAINING_MATRIX_CATEGORY_RECORDS_LIST_ID"),
    ],
    ["Training Matrix Update", envId("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID")],
    ["NPORS Register", envId("SHAREPOINT_NPORS_REGISTER_LIST_ID")],
    ["EUSR Register", envId("SHAREPOINT_EUSR_REGISTER_LIST_ID")],
    ["Streetworks Register", envId("SHAREPOINT_NRSWA_REGISTER_LIST_ID")],
    ["In-House Certificates", envId("SHAREPOINT_IN_HOUSE_CERTIFICATES_LIST_ID")],
    ["NVQ Register", envId("SHAREPOINT_NVQ_REGISTER_LIST_ID")],
    ["Events", envId("SHAREPOINT_EVENTS_LIST_ID")],
    ["Offers / Promotions", envId("SHAREPOINT_OFFERS_PROMOTIONS_LIST_ID")],
    ["Training Manager Logs", envId("SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID")],
    ["Workforce", envId("SHAREPOINT_WORKFORCE_LIST_ID")],
    ["Departments", envId("SHAREPOINT_DEPARTMENTS_LIST_ID")],
  ];

  const results = [];
  for (const [label, listId, opts] of order) {
    results.push(await wipeList(client, label, listId, opts || {}));
  }

  await unlinkKeptPermissionsCompany(client, permissionsId);
  results.push(
    await wipeList(client, "Permissions (except Wayne Curry)", permissionsId, {
      keep: keepPermission,
    }),
  );
  results.push(
    await wipeList(client, "Companies", envId("SHAREPOINT_COMPANY_LIST_ID")),
  );

  // Second pass for anything Restrict Delete blocked the first time.
  console.log("\n--- retry pass ---");
  for (const [label, listId, opts] of [
    ...order,
    ["Permissions (except Wayne Curry)", permissionsId, { keep: keepPermission }],
    ["Companies", envId("SHAREPOINT_COMPANY_LIST_ID")],
  ]) {
    const again = await wipeList(client, `${label} (retry)`, listId, opts || {});
    results.push(again);
  }

  console.log("\n========== SUMMARY ==========");
  console.log(`Mode: ${DRY_RUN ? "DRY RUN" : "LIVE"}`);
  let leftover = 0;
  for (const row of results.filter((r) => !String(r.label).includes("(retry)"))) {
    leftover += row.remaining;
    console.log(
      ` - ${row.label}: deleted ${row.deleted}, remaining ${row.remaining}`,
    );
  }
  if (!DRY_RUN) {
    const kept = await listAllItems(client, permissionsId, true);
    console.log("\nPermissions still present:");
    for (const item of kept) {
      console.log(
        ` - #${item.id} ${item.fields?.Name || ""} <${item.fields?.UserEmail || ""}> ${item.fields?.RoleType || ""} ${item.fields?.Status || ""}`,
      );
    }
  }
  if (leftover > 0 && !DRY_RUN) {
    console.log(`\nWARNING: ${leftover} operational rows still remaining.`);
    process.exitCode = 1;
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
