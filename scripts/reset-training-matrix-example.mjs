/**
 * Wipe Training matrix example → convert Number date columns to DateTime →
 * re-upload from Excel with real dates → verify admin-visible rows.
 *
 * Usage:
 *   node --env-file=.env.local scripts/reset-training-matrix-example.mjs
 *   node --env-file=.env.local scripts/reset-training-matrix-example.mjs --limit=30
 *   node --env-file=.env.local scripts/reset-training-matrix-example.mjs --verify-only
 *   node --env-file=.env.local scripts/reset-training-matrix-example.mjs --skip-convert
 */

import { createRequire } from "node:module";
import { resolve } from "node:path";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const VERIFY_ONLY = process.argv.includes("--verify-only");
const SKIP_CONVERT = process.argv.includes("--skip-convert");
const limitArg = process.argv.find((a) => a.startsWith("--limit="));
const ROW_LIMIT = limitArg ? Number(limitArg.split("=")[1]) : null;

const EXCEL_PATH = resolve(process.cwd(), "Training matrix example.xlsx");

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

function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function safeInternalName(header) {
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

function excelSerialToIso(value) {
  if (value == null || value === "") return null;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed || /^(—|–|-|n\/?a|null|none|0)$/i.test(trimmed)) return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) return trimmed.slice(0, 10);
    const asNum = Number(trimmed);
    if (!Number.isNaN(asNum) && asNum > 20000) return excelSerialToIso(asNum);
    const parsed = Date.parse(trimmed);
    if (!Number.isNaN(parsed)) return new Date(parsed).toISOString().slice(0, 10);
    return null;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value) || value === 0) return null;
    const ms = Date.UTC(1899, 11, 30) + Math.round(value) * 86_400_000;
    return new Date(ms).toISOString().slice(0, 10);
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  return null;
}

function isoToSharePointDateTime(iso) {
  if (!iso?.trim()) return null;
  const text = iso.trim().slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  return `${text}T00:00:00Z`;
}

function looksLikeSerialNumber(value) {
  if (typeof value === "number") return value > 20000 && value < 80000;
  if (typeof value === "string") {
    const t = value.trim();
    if (/^\d{5}(\.\d+)?$/.test(t)) {
      const n = Number(t);
      return n > 20000 && n < 80000;
    }
  }
  return false;
}

function looksLikeDateTime(value) {
  if (value == null) return false;
  return /^\d{4}-\d{2}-\d{2}/.test(String(value));
}

async function listAllItems(client, listId) {
  const items = [];
  let url = `${siteRoot()}/lists/${listId}/items?$expand=fields&$top=200`;
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

async function listColumns(client, listId) {
  const cols = [];
  let url = `${siteRoot()}/lists/${listId}/columns?$top=200`;
  while (url) {
    const page = await client
      .api(url.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, ""))
      .get();
    cols.push(...(page.value ?? []));
    url = page["@odata.nextLink"]
      ? String(page["@odata.nextLink"]).replace(
          /^https:\/\/graph\.microsoft\.com\/v1\.0/i,
          "",
        )
      : null;
  }
  return cols;
}

function buildColumnMap(columns) {
  const map = new Map();
  for (const col of columns) {
    if (col.readOnly || !col.name || !col.displayName) continue;
    if (col.name === "ContentType" || col.name === "Attachments") continue;
    map.set(normalizeHeader(col.displayName).toLowerCase(), {
      name: col.name,
      id: col.id,
      isDateTime: Boolean(col.dateTime),
      isNumber: Boolean(col.number),
      displayName: col.displayName,
    });
    if (col.name === "Title") {
      map.set("name", {
        name: col.name,
        id: col.id,
        isDateTime: false,
        isNumber: false,
        displayName: col.displayName,
      });
    }
  }
  return map;
}

async function wipeList(client, listId) {
  const items = await listAllItems(client, listId);
  console.log(`  deleting ${items.length} existing rows…`);
  let deleted = 0;
  let errors = 0;
  for (const item of items) {
    try {
      await client.api(`${siteRoot()}/lists/${listId}/items/${item.id}`).delete();
      deleted += 1;
      if (deleted % 25 === 0) console.log(`    deleted ${deleted}/${items.length}`);
    } catch (error) {
      errors += 1;
      console.warn(`    delete #${item.id} failed:`, error?.message || error);
    }
  }
  return { deleted, errors, before: items.length };
}

/**
 * Number (Excel serial) columns cannot store calendar dates in SharePoint UI.
 * Delete each Number date column and recreate as DateTime with the same display name.
 */
async function convertNumberColumnsToDateTime(client, listId, columns) {
  const skip = new Set([
    "Title",
    "ContentType",
    "Attachments",
    "Edit",
    "LinkTitle",
    "LinkTitleNoMenu",
    "ID",
  ]);

  const targets = columns.filter(
    (col) =>
      col.number &&
      !col.readOnly &&
      col.name &&
      !skip.has(col.name) &&
      /^field_\d+$/i.test(col.name),
  );

  console.log(`  converting ${targets.length} Number columns → DateTime…`);
  let converted = 0;
  let failed = 0;
  const failures = [];

  for (const col of targets) {
    const displayName = normalizeHeader(col.displayName);
    const newName = safeInternalName(displayName);
    process.stdout.write(`    ${displayName} (${col.name} → ${newName}) … `);
    try {
      await client
        .api(`${siteRoot()}/lists/${listId}/columns/${col.id}`)
        .delete();
      await client.api(`${siteRoot()}/lists/${listId}/columns`).post({
        name: newName,
        displayName,
        dateTime: { format: "dateOnly" },
      });
      converted += 1;
      console.log("ok");
    } catch (error) {
      failed += 1;
      const msg = error?.message || String(error);
      failures.push(`${displayName}: ${msg}`);
      console.log("FAIL", msg);
      // Best effort: try recreate even if delete partially succeeded
      try {
        await client.api(`${siteRoot()}/lists/${listId}/columns`).post({
          name: newName,
          displayName,
          dateTime: { format: "dateOnly" },
        });
        converted += 1;
        failed -= 1;
        console.log("      recovered recreate ok");
      } catch {
        // leave failed
      }
    }
  }

  return { converted, failed, failures };
}

function parseExcelRows() {
  const wb = XLSX.readFile(EXCEL_PATH, { cellDates: true, raw: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: true,
    blankrows: false,
  });

  let headerIdx = -1;
  for (let i = 0; i < matrix.length; i += 1) {
    const cells = (matrix[i] || []).map((c) =>
      c == null ? "" : String(c).replace(/\u00a0/g, " ").trim(),
    );
    const lower = new Set(cells.map((c) => c.toLowerCase()));
    const hasNCode = cells.some((c) => /^N\d{3}/i.test(c));
    if (lower.has("name") && (lower.has("dob") || hasNCode)) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) {
    throw new Error("Could not find Name/DOB header row in Excel.");
  }

  const headers = (matrix[headerIdx] || []).map((h) =>
    h == null ? "" : String(h).replace(/\u00a0/g, " ").trim(),
  );

  const rows = [];
  for (let i = headerIdx + 1; i < matrix.length; i += 1) {
    const raw = matrix[i] || [];
    const name = raw[0] == null ? "" : String(raw[0]).trim();
    if (!name) continue;
    if (name.toLowerCase() === "name") break;

    const source = {};
    for (let c = 0; c < headers.length; c += 1) {
      const header = headers[c];
      if (!header) continue;
      source[header] = raw[c] ?? null;
    }
    rows.push(source);
  }
  return { headers: headers.filter(Boolean), rows };
}

function buildFields(source, columnMap) {
  const name = String(source.Name ?? source.Title ?? "").trim();
  const fields = { Title: name };
  let dateFields = 0;

  for (const [header, raw] of Object.entries(source)) {
    if (normalizeHeader(header).toLowerCase() === "name") continue;
    const col = columnMap.get(normalizeHeader(header).toLowerCase());
    if (!col?.name) continue;
    const iso = excelSerialToIso(raw);
    const spDate = isoToSharePointDateTime(iso);
    if (col.isNumber && !col.isDateTime) {
      // Still Number column — write serial so upload doesn't break
      fields[col.name] = iso
        ? Math.round(
            (Date.parse(`${iso}T00:00:00Z`) - Date.UTC(1899, 11, 30)) /
              86_400_000,
          )
        : null;
    } else {
      fields[col.name] = spDate;
    }
    if (spDate || fields[col.name]) dateFields += 1;
  }
  return { fields, dateFields };
}

async function uploadRows(client, listId, rows, columnMap) {
  let created = 0;
  let errors = 0;
  let withDates = 0;
  const sampleFailures = [];
  const toUpload = ROW_LIMIT ? rows.slice(0, ROW_LIMIT) : rows;
  console.log(`  uploading ${toUpload.length} rows (of ${rows.length} in Excel)…`);

  for (const source of toUpload) {
    const name = String(source.Name ?? "").trim();
    if (!name) continue;
    try {
      const { fields, dateFields } = buildFields(source, columnMap);
      if (dateFields > 0) withDates += 1;
      await client.api(`${siteRoot()}/lists/${listId}/items`).post({ fields });
      created += 1;
      if (created % 25 === 0) {
        console.log(`    created ${created}/${toUpload.length}`);
      }
    } catch (error) {
      errors += 1;
      if (sampleFailures.length < 8) {
        sampleFailures.push(`${name}: ${error?.message || error}`);
      }
    }
  }
  return { created, errors, withDates, sampleFailures, attempted: toUpload.length };
}

async function verify(client, listId, columnMap) {
  const items = await listAllItems(client, listId);
  console.log(`\n=== Verify live SharePoint (${items.length} rows) ===`);

  const dateCols = [...columnMap.values()].filter(
    (c) => c.name !== "Title" && (c.isDateTime || c.isNumber),
  );

  let serialHits = 0;
  let isoHits = 0;
  const samples = [];

  for (const item of items.slice(0, 12)) {
    const f = item.fields ?? {};
    const title = f.Title || f.LinkTitle || "";
    const dateSample = {};
    let rowHasSerial = false;

    for (const col of dateCols) {
      const value = f[col.name];
      if (value == null || value === "" || value === 0) continue;
      if (looksLikeSerialNumber(value)) {
        serialHits += 1;
        rowHasSerial = true;
      }
      if (looksLikeDateTime(value)) isoHits += 1;
      if (Object.keys(dateSample).length < 5) {
        dateSample[`${col.displayName}(${col.name})`] = value;
      }
    }

    samples.push({ id: item.id, title, serial: rowHasSerial, dates: dateSample });
  }

  console.log("Sample rows:");
  for (const s of samples.slice(0, 8)) {
    console.log(
      `  #${s.id} ${s.title} serial?=${s.serial} dates=${JSON.stringify(s.dates)}`,
    );
  }

  const titled = items.filter(
    (i) => String(i.fields?.Title || i.fields?.LinkTitle || "").trim().length > 0,
  );

  const dateTimeCols = dateCols.filter((c) => c.isDateTime).length;
  const numberCols = dateCols.filter((c) => c.isNumber && !c.isDateTime).length;

  console.log("\nAdmin Training Matrix visibility:");
  console.log(`  rows with Name/Title: ${titled.length}`);
  console.log(`  DateTime columns: ${dateTimeCols}`);
  console.log(`  remaining Number date columns: ${numberCols}`);
  console.log(`  ISO-like date hits (sample): ${isoHits}`);
  console.log(`  serial-number hits (sample): ${serialHits}`);

  const ok = titled.length > 0 && serialHits === 0 && dateTimeCols > 0;
  return {
    ok,
    count: titled.length,
    serialHits,
    isoHits,
    dateTimeCols,
    numberCols,
    sampleNames: samples.map((s) => s.title).filter(Boolean),
  };
}

async function main() {
  const listId = requireEnv("SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID");
  const client = getClient();

  console.log("=== Training matrix example reset ===");
  console.log(`List ID: ${listId}`);
  console.log(`Excel:   ${EXCEL_PATH}`);

  let columns = await listColumns(client, listId);
  let columnMap = buildColumnMap(columns);
  console.log(`Columns: ${columns.length} (mapped ${columnMap.size})`);

  if (VERIFY_ONLY) {
    const result = await verify(client, listId, columnMap);
    process.exitCode = result.ok ? 0 : 1;
    return;
  }

  console.log("\n1) Wipe existing rows…");
  const wiped = await wipeList(client, listId);
  console.log(`  deleted=${wiped.deleted} errors=${wiped.errors}`);

  let converted = { converted: 0, failed: 0, failures: [] };
  if (!SKIP_CONVERT) {
    console.log("\n2) Convert Number → DateTime columns…");
    columns = await listColumns(client, listId);
    converted = await convertNumberColumnsToDateTime(client, listId, columns);
    console.log(
      `  converted=${converted.converted} failed=${converted.failed}`,
    );
    columns = await listColumns(client, listId);
    columnMap = buildColumnMap(columns);
  } else {
    console.log("\n2) Skipping column convert (--skip-convert)");
  }

  console.log("\n3) Parse Excel…");
  const { headers, rows } = parseExcelRows();
  console.log(`  headers=${headers.length} dataRows=${rows.length}`);

  console.log("\n4) Upload with real DateTime values…");
  const uploaded = await uploadRows(client, listId, rows, columnMap);
  console.log(
    `  created=${uploaded.created} errors=${uploaded.errors} rowsWithDates=${uploaded.withDates}`,
  );
  if (uploaded.sampleFailures.length) {
    console.log("  sample failures:");
    for (const msg of uploaded.sampleFailures) console.log(`   - ${msg}`);
  }

  console.log("\n5) Verify live data + admin visibility…");
  columns = await listColumns(client, listId);
  columnMap = buildColumnMap(columns);
  const result = await verify(client, listId, columnMap);

  console.log("\n========== SUMMARY ==========");
  console.log(
    JSON.stringify(
      {
        wiped,
        converted,
        uploaded: {
          created: uploaded.created,
          errors: uploaded.errors,
          withDates: uploaded.withDates,
          attempted: uploaded.attempted,
        },
        verify: result,
      },
      null,
      2,
    ),
  );

  if (!result.ok || uploaded.errors > 0 || converted.failed > 0) {
    process.exitCode = 1;
  } else {
    console.log(
      "PASS: SharePoint shows DateTime dates; admin matrix can load named rows.",
    );
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
