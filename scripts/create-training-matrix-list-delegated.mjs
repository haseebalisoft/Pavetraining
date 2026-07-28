/**
 * Create SharePoint "Training Matrix" list with all columns from
 * Training matrix example.xlsx, using YOUR login (device code).
 *
 * App-only auth cannot create lists (403). This uses delegated auth.
 *
 * Usage (in zsh/bash — no PowerShell needed):
 *   node --env-file=.env scripts/create-training-matrix-list-delegated.mjs
 */
import { createRequire } from "node:module";
import { resolve, dirname } from "node:path";
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { DeviceCodeCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const EXCEL_PATH = resolve(ROOT, "Training matrix example.xlsx");
const LIST_DISPLAY_NAME = "Training Matrix";

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
    return { name, displayName: header, text: { allowMultipleLines: false } };
  }
  return { name, displayName: header, dateTime: { format: "dateOnly" } };
}

function updateEnvListId(listId) {
  const envPath = resolve(ROOT, ".env");
  let env = readFileSync(envPath, "utf8");
  const prev = process.env.SHAREPOINT_TRAINING_MATRIX_LIST_ID?.trim();
  if (/SHAREPOINT_TRAINING_MATRIX_LIST_ID=/.test(env)) {
    env = env.replace(
      /SHAREPOINT_TRAINING_MATRIX_LIST_ID=.*/,
      `SHAREPOINT_TRAINING_MATRIX_LIST_ID=${listId}`,
    );
  } else {
    env += `\nSHAREPOINT_TRAINING_MATRIX_LIST_ID=${listId}\n`;
  }
  if (
    prev &&
    prev !== listId &&
    !/SHAREPOINT_TRAINING_MATRIX_LIST_ID_PREVIOUS=/.test(env)
  ) {
    env += `SHAREPOINT_TRAINING_MATRIX_LIST_ID_PREVIOUS=${prev}\n`;
  }
  writeFileSync(envPath, env);
}

async function main() {
  const headers = parseHeaders();
  console.log(
    `Creating list "${LIST_DISPLAY_NAME}" with ${headers.length} columns…`,
  );
  console.log(
    "\nSign in as a Site Owner when prompted (device code).\n",
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

  // Delegated scopes — user must be Site Owner / have Manage Lists
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/Sites.FullControl.All"],
  });
  const client = Client.initWithMiddleware({ authProvider });

  const existing = await client.api(`${siteRoot()}/lists?$top=200`).get();
  const clash = (existing.value || []).find(
    (l) =>
      String(l.displayName || "").trim().toLowerCase() ===
      LIST_DISPLAY_NAME.toLowerCase(),
  );

  const stamp = new Date()
    .toISOString()
    .replace(/[-:TZ.]/g, "")
    .slice(0, 12);
  const displayName = clash
    ? `${LIST_DISPLAY_NAME} (${stamp})`
    : LIST_DISPLAY_NAME;
  if (clash) {
    console.log(
      `Existing "${LIST_DISPLAY_NAME}" found — creating as "${displayName}"`,
    );
  }

  let created;
  try {
    created = await client.api(`${siteRoot()}/lists`).post({
      displayName,
      description:
        "Full training matrix matching Training matrix example.xlsx headers.",
      list: { template: "genericList" },
    });
  } catch (err) {
    console.error("\nFailed to create list:", err.message || err);
    console.error(
      "Make sure you signed in as Site Owner, and the Entra app has",
    );
    console.error(
      "delegated permission Sites.FullControl.All (admin consent).",
    );
    process.exit(1);
  }

  const listId = created.id;
  console.log(`\nCreated list id=${listId} name="${created.displayName}"`);
  console.log("Adding columns…\n");

  let ok = 0;
  const failed = [];
  for (const header of headers) {
    const body = columnDefinition(header);
    process.stdout.write(`  + ${header} … `);
    try {
      await client.api(`${siteRoot()}/lists/${listId}/columns`).post(body);
      console.log("ok");
      ok += 1;
    } catch (err) {
      console.log("FAIL", err.message || err);
      failed.push({ header, error: err.message || String(err) });
    }
  }

  updateEnvListId(listId);
  console.log(`\nUpdated .env SHAREPOINT_TRAINING_MATRIX_LIST_ID=${listId}`);

  console.log("\n========== CREATE SUMMARY ==========");
  console.log(`List: ${displayName}`);
  console.log(`List ID: ${listId}`);
  console.log(`Columns: ${ok}/${headers.length}`);
  console.log(`Failed: ${failed.length}`);
  if (failed.length) {
    for (const f of failed.slice(0, 20)) console.log(`  - ${f.header}: ${f.error}`);
    process.exitCode = 1;
  } else {
    console.log("RESULT: PASS");
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
