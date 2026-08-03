/**
 * Expand NPORS Register → NPORS Category MultiChoice to match the Training Matrix
 * category codes (short labels live in the app; SharePoint stores the code only).
 *
 * App-only Graph usually returns 403 on column schema updates. This script uses
 * YOUR login (device code) as Site Owner.
 *
 * Prerequisites:
 *   - Entra app has delegated Sites.FullControl.All (admin consent)
 *   - You sign in as a SharePoint Site Owner
 *
 * Usage:
 *   node --env-file=.env.local scripts/update-npors-category-choices.mjs --dry-run
 *   node --env-file=.env.local scripts/update-npors-category-choices.mjs
 *
 * Optional: also try app-only first (will fall back to device code on 403):
 *   node --env-file=.env.local scripts/update-npors-category-choices.mjs --try-app
 */
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { DeviceCodeCredential, ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DRY_RUN = process.argv.includes("--dry-run");
const TRY_APP = process.argv.includes("--try-app");

function loadEnv() {
  try {
    for (const line of readFileSync(resolve(ROOT, ".env.local"), "utf8").split(
      /\r?\n/,
    )) {
      const t = line.trim();
      if (!t || t.startsWith("#")) continue;
      const i = t.indexOf("=");
      if (i < 0) continue;
      const k = t.slice(0, i).trim();
      let v = t.slice(i + 1).trim();
      if (
        (v.startsWith('"') && v.endsWith('"')) ||
        (v.startsWith("'") && v.endsWith("'"))
      ) {
        v = v.slice(1, -1);
      }
      if (!process.env[k]) process.env[k] = v;
    }
  } catch {
    // ignore
  }
}

loadEnv();

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

/** Codes from Training Matrix template (same source as admin NPORS dropdown). */
function loadTargetCodes() {
  const text = readFileSync(
    resolve(ROOT, "src/lib/services/bulkUpload/clientTemplateHeaders.ts"),
    "utf8",
  );
  const codes = [];
  const re = /"code":\s*"([^"]+)"/g;
  let match;
  while ((match = re.exec(text))) {
    const code = match[1].trim().toUpperCase();
    if (/^N\d+[A-Z]?$/i.test(code) || /^S\d+$/i.test(code)) {
      codes.push(code);
    }
  }
  return Array.from(new Set(codes)).sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true }),
  );
}

async function listAll(client, firstPath) {
  const items = [];
  let path = firstPath;
  while (path) {
    const page = await client.api(path).get();
    items.push(...(page.value || []));
    path = page["@odata.nextLink"]
      ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
      : null;
  }
  return items;
}

function makeClient(credential, scopes) {
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes,
  });
  return Client.initWithMiddleware({ authProvider });
}

async function findCategoryColumn(client, listId) {
  const columns = await listAll(
    client,
    `${siteRoot()}/lists/${listId}/columns?$top=200`,
  );
  return (
    columns.find(
      (c) =>
        String(c.name || "").toLowerCase() === "nporscategory" ||
        /npors\s*category/i.test(c.displayName || ""),
    ) || null
  );
}

async function patchChoices(client, listId, column, choices) {
  return client.api(`${siteRoot()}/lists/${listId}/columns/${column.id}`).patch({
    choice: {
      ...(column.choice || {}),
      allowTextEntry: false,
      choices,
    },
  });
}

async function runWithClient(client, label) {
  const listId = requireEnv("SHAREPOINT_NPORS_REGISTER_LIST_ID");
  const target = loadTargetCodes();
  console.log(`\nAuth: ${label}`);
  console.log(`NPORS Register list: ${listId}`);
  console.log(`Target codes: ${target.length}`);

  const column = await findCategoryColumn(client, listId);
  if (!column) {
    throw new Error('Could not find "NPORS Category" column on NPORS Register.');
  }

  const current = (column.choice?.choices || [])
    .map((c) => String(c).trim())
    .filter(Boolean);
  const currentClean = current.filter(
    (c) => !/^choice\s*\d+$/i.test(c) && c.toUpperCase() !== "OTHER",
  );

  console.log(`Column: ${column.displayName} (${column.name})`);
  console.log(`Current choices (${current.length}): ${current.join(", ") || "(none)"}`);
  console.log(`Target choices (${target.length}): ${target.join(", ")}`);

  const same =
    currentClean.length === target.length &&
    target.every((code, i) => currentClean[i] === code);

  if (same && current.length === currentClean.length) {
    console.log("\nAlready aligned — nothing to do.");
    return { ok: true, updated: false };
  }

  if (DRY_RUN) {
    console.log(
      `\n[dry-run] Would replace choices with ${target.length} NPORS codes (junk like "Choice 8" removed).`,
    );
    return { ok: true, updated: false };
  }

  await patchChoices(client, listId, column, target);
  console.log(`\nUpdated NPORS Category choices (${target.length} codes).`);
  return { ok: true, updated: true };
}

async function main() {
  const listId = requireEnv("SHAREPOINT_NPORS_REGISTER_LIST_ID");
  void listId;

  if (TRY_APP) {
    console.log("Trying app-only credentials first…");
    try {
      const appCred = new ClientSecretCredential(
        requireEnv("AZURE_TENANT_ID"),
        requireEnv("AZURE_CLIENT_ID"),
        requireEnv("AZURE_CLIENT_SECRET"),
      );
      const appClient = makeClient(appCred, [
        "https://graph.microsoft.com/.default",
      ]);
      await runWithClient(appClient, "app-only");
      return;
    } catch (error) {
      const message = error?.message || String(error);
      console.warn(`App-only failed: ${message}`);
      if (!/403|accessDenied|Access denied/i.test(message)) {
        throw error;
      }
      console.warn("Falling back to device-code (Site Owner sign-in)…");
    }
  }

  console.log(
    "\nSign in as a SharePoint Site Owner when prompted (device code).\n",
  );

  const credential = new DeviceCodeCredential({
    tenantId: requireEnv("AZURE_TENANT_ID"),
    clientId: requireEnv("AZURE_CLIENT_ID"),
    clientSecret: process.env.AZURE_CLIENT_SECRET?.trim() || undefined,
    userPromptCallback: (info) => {
      console.log("\n========================================");
      console.log(info.message);
      console.log("========================================\n");
    },
  });

  const client = makeClient(credential, [
    "https://graph.microsoft.com/Sites.FullControl.All",
  ]);

  try {
    await runWithClient(client, "delegated (device code)");
  } catch (error) {
    console.error("\nFailed to update NPORS Category choices:");
    console.error(error?.message || error);
    console.error(
      "\nChecks:\n" +
        "  1. Sign in as Site Owner on PaveTrainingOperationAdmin\n" +
        "  2. Entra app has delegated Sites.FullControl.All + admin consent\n" +
        "  3. Or set choices manually: NPORS Register → List settings → NPORS Category",
    );
    process.exit(1);
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
