/**
 * Create a new SharePoint list "Training Matrix" with ALL columns from
 * Training matrix example.xlsx (exact display names).
 *
 * Usage:
 *   node --env-file=.env scripts/create-training-matrix-list.mjs
 *   node --env-file=.env scripts/create-training-matrix-list.mjs --dry-run
 *   node --env-file=.env scripts/create-training-matrix-list.mjs --name "Training Matrix"
 */
import { createRequire } from "node:module";
import { resolve, dirname } from "node:path";
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const EXCEL_PATH = resolve(ROOT, "Training matrix example.xlsx");
const DRY_RUN = process.argv.includes("--dry-run");

function argValue(flag, fallback) {
  const idx = process.argv.indexOf(flag);
  if (idx >= 0 && process.argv[idx + 1]) return process.argv[idx + 1];
  return fallback;
}

const LIST_DISPLAY_NAME = argValue("--name", "Training Matrix");

function requireEnv(name) {
  const value = process.env[name];
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

function parseHeaders() {
  const wb = XLSX.readFile(EXCEL_PATH, { cellDates: true, raw: false });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: null,
    raw: false,
  });
  return (rows[0] || [])
    .map((h) => (h == null ? "" : String(h).replace(/\u00a0/g, " ").trim()))
    .filter(Boolean);
}

/** Safe SharePoint internal column name from template header. */
function internalName(header) {
  if (header === "Name") return "CandidateNameText";
  if (header === "DOB") return "DOB";
  if (header === "Face ift") return "FaceFitExpiry";
  const meta = {
    "CSCS Expiry": "CSCSExpiry",
    "SSSTS Expiry": "SSSTSExpiry",
    "SMSTS Expiry": "SMSTSExpiry",
    "NRSWA Expiry": "NRSWAExpiry",
    "EUSR Expiry": "EUSRExpiry",
  };
  if (meta[header]) return meta[header];
  const code = header.match(/^(N\d+[A-Z]?)\b/i)?.[1];
  if (code) return `${code.toUpperCase()}Expiry`;
  return header.replace(/[^A-Za-z0-9]/g, "").slice(0, 32) || "Col";
}

function columnDefinition(header) {
  const name = internalName(header);
  if (header === "Name") {
    return {
      name,
      displayName: header,
      text: { allowMultipleLines: false },
    };
  }
  // DOB + all expiry columns are date-only
  return {
    name,
    displayName: header,
    dateTime: { format: "dateOnly" },
  };
}

async function listExisting(client) {
  const res = await client.api(`${siteRoot()}/lists?$top=200`).get();
  return res.value || [];
}

async function createColumns(client, listId, headers) {
  const created = [];
  const failed = [];
  for (const header of headers) {
    // Title already exists on generic lists — map Name onto Title display later if needed.
    // We still create CandidateNameText for an explicit "Name" column.
    const body = columnDefinition(header);
    process.stdout.write(`  + ${header} (${body.name}) … `);
    if (DRY_RUN) {
      console.log("dry-run");
      created.push(header);
      continue;
    }
    try {
      await client.api(`${siteRoot()}/lists/${listId}/columns`).post(body);
      console.log("ok");
      created.push(header);
    } catch (err) {
      console.log("FAIL", err.message || err);
      failed.push({ header, error: err.message || String(err) });
    }
  }
  return { created, failed };
}

function updateEnvListId(listId) {
  const envPath = resolve(ROOT, ".env");
  let env = readFileSync(envPath, "utf8");
  if (/SHAREPOINT_TRAINING_MATRIX_LIST_ID=/.test(env)) {
    env = env.replace(
      /SHAREPOINT_TRAINING_MATRIX_LIST_ID=.*/,
      `SHAREPOINT_TRAINING_MATRIX_LIST_ID=${listId}`,
    );
  } else {
    env += `\nSHAREPOINT_TRAINING_MATRIX_LIST_ID=${listId}\n`;
  }
  // Keep previous list id for reference
  if (!/SHAREPOINT_TRAINING_MATRIX_LIST_ID_PREVIOUS=/.test(env)) {
    const prev = process.env.SHAREPOINT_TRAINING_MATRIX_LIST_ID;
    if (prev && prev !== listId) {
      env += `SHAREPOINT_TRAINING_MATRIX_LIST_ID_PREVIOUS=${prev}\n`;
    }
  }
  writeFileSync(envPath, env);
}

async function main() {
  const headers = parseHeaders();
  console.log(
    DRY_RUN ? "DRY RUN" : "LIVE",
    `— create list "${LIST_DISPLAY_NAME}" with ${headers.length} columns from Training matrix example.xlsx`,
  );

  const client = getClient();
  const existing = await listExisting(client);
  const clash = existing.find(
    (l) =>
      String(l.displayName || "").trim().toLowerCase() ===
      LIST_DISPLAY_NAME.trim().toLowerCase(),
  );
  if (clash) {
    console.log(
      `Note: a list named "${clash.displayName}" already exists (id=${clash.id}). Creating a new list anyway with a unique internal name.`,
    );
  }

  const uniqueSuffix = new Date()
    .toISOString()
    .replace(/[-:TZ.]/g, "")
    .slice(0, 12);
  const listPayload = {
    displayName: LIST_DISPLAY_NAME,
    description:
      "Full training matrix matching Training matrix example.xlsx column headers.",
    list: { template: "genericList" },
  };

  // Graph uses displayName; if duplicate display names are awkward, append marker only when clash.
  if (clash) {
    listPayload.displayName = `${LIST_DISPLAY_NAME} (${uniqueSuffix})`;
    console.log(`Using display name: ${listPayload.displayName}`);
  }

  let listId;
  if (DRY_RUN) {
    listId = "dry-run-list-id";
    console.log("Would create list:", listPayload.displayName);
  } else {
    try {
      const created = await client.api(`${siteRoot()}/lists`).post(listPayload);
      listId = created.id;
      console.log(`Created list id=${listId} name="${created.displayName}"`);
    } catch (err) {
      console.error("Failed to create list:", err.message || err);
      console.error(
        "The app likely needs Sites.FullControl.All (or Manage Lists) on this site.",
      );
      process.exit(1);
    }
  }

  console.log("\nAdding columns…");
  const { created, failed } = await createColumns(client, listId, headers);

  if (!DRY_RUN && listId && created.length) {
    updateEnvListId(listId);
    console.log(
      `\nUpdated .env SHAREPOINT_TRAINING_MATRIX_LIST_ID=${listId}`,
    );
  }

  console.log("\n========== CREATE SUMMARY ==========");
  console.log(`List: ${listPayload.displayName}`);
  console.log(`List ID: ${listId}`);
  console.log(`Columns created: ${created.length}/${headers.length}`);
  console.log(`Columns failed: ${failed.length}`);
  if (failed.length) {
    for (const f of failed.slice(0, 30)) {
      console.log(`  - ${f.header}: ${f.error}`);
    }
    process.exitCode = 1;
  } else {
    console.log(
      "RESULT: PASS — new Training Matrix list created with all template columns",
    );
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
