/**
 * Sync SharePoint Training Matrix list columns to match the template headers in
 * "Sharepoint list _ training matrix example 11-07-26.xlsx".
 *
 * - Creates missing Date columns (internal {CODE}Expiry, display = exact template header)
 * - Renames display names of existing expiry columns to match template spelling
 * - Regenerates schema + CLIENT matrixField mappings
 * - Does NOT wipe data (run wipe/import separately)
 *
 * Usage:
 *   node --env-file=.env scripts/sync-training-matrix-columns.mjs
 *   node --env-file=.env scripts/sync-training-matrix-columns.mjs --dry-run
 */
import { createRequire } from "node:module";
import { resolve, dirname } from "node:path";
import { writeFileSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const DRY_RUN = process.argv.includes("--dry-run");
const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const EXCEL_PATH = resolve(
  ROOT,
  "Sharepoint list _ training matrix example 11-07-26.xlsx",
);

const META = [
  { header: "CSCS Expiry", code: "CSCS", appKey: "cscsExpiry", sp: "CSCSExpiry" },
  { header: "SSSTS Expiry", code: "SSSTS", appKey: "ssstsExpiry", sp: "SSSTSExpiry" },
  { header: "SMSTS Expiry", code: "SMSTS", appKey: "smstsExpiry", sp: "SMSTSExpiry" },
  { header: "NRSWA Expiry", code: "NRSWA", appKey: "nrswaExpiry", sp: "NRSWAExpiry" },
  { header: "EUSR Expiry", code: "EUSR", appKey: "eusrExpiry", sp: "EUSRExpiry" },
];

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

function extractCode(header) {
  const m = String(header).trim().match(/^(N\d+[A-Z]?)\b/i);
  return m ? m[1].toUpperCase() : null;
}

function parseTemplateHeaders() {
  const wb = XLSX.readFile(EXCEL_PATH, { cellDates: true, raw: false });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: null,
    raw: false,
    blankrows: true,
  });
  let headerIdx = -1;
  for (let i = 0; i < rows.length; i += 1) {
    const cells = (rows[i] || []).map((c) =>
      c == null ? "" : String(c).replace(/\u00a0/g, " ").trim(),
    );
    const lower = new Set(cells.map((c) => c.toLowerCase()));
    if (
      lower.has("name") &&
      lower.has("dob") &&
      cells.some((c) => /^N\d{3}/i.test(c))
    ) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) throw new Error("Matrix header row not found in template.");
  return (rows[headerIdx] || [])
    .map((h) => (h == null ? "" : String(h).replace(/\u00a0/g, " ").trim()))
    .filter(Boolean);
}

async function listAllColumns(client, listId) {
  const columns = [];
  let url = `${siteRoot()}/lists/${listId}/columns?$top=200`;
  while (url) {
    const page = await client.api(url).get();
    columns.push(...(page.value || []));
    url = page["@odata.nextLink"]
      ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
      : null;
  }
  return columns;
}

async function main() {
  const headers = parseTemplateHeaders();
  const categoryHeaders = headers.filter(
    (h) => !["Name", "DOB"].includes(h) && !META.some((m) => m.header === h),
  );

  const defs = [
    ...META.map((m) => ({
      header: m.header,
      code: m.code,
      appKey: m.appKey,
      spField: m.sp,
      name: m.code,
    })),
    ...categoryHeaders.map((header) => {
      const code = extractCode(header);
      if (!code) throw new Error(`No N-code in header: ${header}`);
      const name = header.includes(" - ")
        ? header.split(" - ").slice(1).join(" - ").trim()
        : code;
      return {
        header,
        code,
        appKey: `${code.toLowerCase()}Expiry`,
        spField: `${code}Expiry`,
        name,
      };
    }),
  ];

  console.log(
    DRY_RUN ? "DRY RUN" : "LIVE",
    `— syncing ${defs.length} expiry columns to Training Matrix`,
  );

  const client = getClient();
  const listId = requireEnv("SHAREPOINT_TRAINING_MATRIX_LIST_ID");
  const existing = await listAllColumns(client, listId);
  const byName = new Map(existing.map((c) => [c.name, c]));
  const byDisplay = new Map(
    existing.map((c) => [String(c.displayName || "").trim().toLowerCase(), c]),
  );

  let created = 0;
  let renamed = 0;
  let skipped = 0;
  const errors = [];
  const actuallyCreatedOrRenamed = [];

  for (const def of defs) {
    // Prefer exact internal name match — never rename Workforce lookup projections.
    let col = byName.get(def.spField) || null;
    if (!col) {
      // Match only writable dateTime columns by display name.
      const byDisp = byDisplay.get(def.header.toLowerCase());
      if (byDisp?.dateTime && !String(byDisp.name).includes("_x003a_")) {
        col = byDisp;
      }
    }
    if (!col) {
      const legacy = byDisplay.get(`${def.code} Expiry`.toLowerCase());
      if (
        legacy?.dateTime &&
        !String(legacy.name).includes("_x003a_") &&
        (legacy.name === def.spField || /^N\d+Expiry$/i.test(legacy.name))
      ) {
        col = legacy;
      }
    }

    if (!col) {
      console.log(`  CREATE ${def.spField}  display="${def.header}"`);
      if (!DRY_RUN) {
        try {
          const createdCol = await client
            .api(`${siteRoot()}/lists/${listId}/columns`)
            .post({
              name: def.spField,
              displayName: def.header,
              dateTime: { format: "dateOnly" },
            });
          byName.set(createdCol.name, createdCol);
          byDisplay.set(
            String(createdCol.displayName || "").trim().toLowerCase(),
            createdCol,
          );
          created += 1;
          actuallyCreatedOrRenamed.push(def);
        } catch (err) {
          errors.push(`${def.spField}: ${err.message || err}`);
          console.error(`    FAIL ${def.spField}:`, err.message || err);
        }
      } else {
        created += 1;
      }
      continue;
    }

    const currentDisplay = String(col.displayName || "").trim();
    if (currentDisplay !== def.header) {
      console.log(
        `  RENAME ${col.name}: "${currentDisplay}" → "${def.header}"`,
      );
      if (!DRY_RUN) {
        try {
          await client
            .api(`${siteRoot()}/lists/${listId}/columns/${col.id}`)
            .patch({ displayName: def.header });
          renamed += 1;
          actuallyCreatedOrRenamed.push(def);
        } catch (err) {
          errors.push(`rename ${col.name}: ${err.message || err}`);
          console.error(`    FAIL rename ${col.name}:`, err.message || err);
        }
      } else {
        renamed += 1;
      }
    } else {
      skipped += 1;
      actuallyCreatedOrRenamed.push(def);
    }
  }

  // Only regenerate schema from columns that already exist OR were created successfully.
  const existingWritable = defs.filter((d) => {
    const col = byName.get(d.spField);
    return Boolean(col?.dateTime) || actuallyCreatedOrRenamed.includes(d);
  });
  // Fallback to known 8 N-code columns if create failed (403).
  const KNOWN = new Set([
    "N001",
    "N003",
    "N004",
    "N010",
    "N020",
    "N021",
    "N027",
    "N100",
  ]);
  const schemaDefs = defs.filter(
    (d) => KNOWN.has(d.code) || byName.get(d.spField)?.dateTime,
  );

  if (errors.length && !schemaDefs.length) {
    console.error(
      "\nNo column create permission (Sites.Manage / FullControl needed).",
    );
    console.error(
      "Schema left unchanged. Add columns as Site Owner, then re-run this script.",
    );
    console.error(
      "Checklist: scripts/training-matrix-columns-to-add.csv",
    );
    process.exitCode = 1;
    return;
  }

  // Regenerate CLIENT matrixField values — writable only when SP column exists.
  const headersPath = resolve(
    ROOT,
    "src/lib/services/bulkUpload/clientTemplateHeaders.ts",
  );
  let headersSrc = readFileSync(headersPath, "utf8");
  const writableCodes = new Set(schemaDefs.map((d) => d.code));
  for (const def of defs.filter((d) => d.code.startsWith("N") && /^N\d/i.test(d.code))) {
    const re = new RegExp(
      `("code":\\s*"${def.code}"[\\s\\S]*?"matrixField":\\s*)([^,\\n]+)`,
      "m",
    );
    if (!re.test(headersSrc)) {
      errors.push(`CLIENT header missing code ${def.code}`);
      continue;
    }
    const val = writableCodes.has(def.code) ? `"${def.appKey}"` : "null";
    headersSrc = headersSrc.replace(re, `$1${val}`);
  }
  if (!DRY_RUN) writeFileSync(headersPath, headersSrc);
  console.log(
    DRY_RUN
      ? "Would update clientTemplateHeaders.ts matrixField mappings"
      : "Updated clientTemplateHeaders.ts matrixField mappings",
  );

  // Regenerate schema expiry fields block (writable columns only)
  const schemaPath = resolve(ROOT, "src/lib/schema/sharepointSchema.ts");
  let schemaSrc = readFileSync(schemaPath, "utf8");

  const fieldLines = [
    "const trainingMatrixFields = {",
    '  id: "ID",',
    '  candidateName: "CandidateName",',
    '  matrixCompany: "MatrixCompany",',
    '  companyName: "Company_x0020_Name",',
    '  department: "Department",',
    '  overallStatus: "OverallStatus",',
    '  needsReview: "NeedsReview",',
    '  matrixNotes: "MatrixNotes",',
    '  nextExpiryDate: "NextExpiryDate",',
    ...schemaDefs
      .filter((d) => /^N\d/i.test(d.code))
      .map((d) => `  ${d.appKey}: "${d.spField}",`),
    "} as const;",
  ].join("\n");

  const labelLines = [
    "    labels: {",
    '      id: "ID",',
    '      candidateName: "Candidate name",',
    '      matrixCompany: "Matrix company",',
    '      companyName: "Company name",',
    '      department: "Department",',
    '      overallStatus: "Overall status",',
    '      needsReview: "Needs review",',
    '      matrixNotes: "Matrix notes",',
    '      nextExpiryDate: "Next expiry date",',
    ...schemaDefs
      .filter((d) => /^N\d/i.test(d.code))
      .map((d) => `      ${d.appKey}: ${JSON.stringify(d.header)},`),
    "    },",
  ].join("\n");

  schemaSrc = schemaSrc.replace(
    /const trainingMatrixFields = \{[\s\S]*?\} as const;/,
    fieldLines,
  );
  schemaSrc = schemaSrc.replace(
    /(trainingMatrix: \{[\s\S]*?)labels: \{[\s\S]*?\},(\n  \} satisfies SharePointListDefinition<typeof trainingMatrixFields>)/,
    `$1${labelLines}$2`,
  );

  if (!DRY_RUN) writeFileSync(schemaPath, schemaSrc);
  console.log(
    DRY_RUN
      ? "Would update sharepointSchema.ts trainingMatrix fields/labels"
      : "Updated sharepointSchema.ts trainingMatrix fields/labels",
  );

  // Write mapping helper used by importer/e2e
  const mapPath = resolve(
    ROOT,
    "src/lib/services/bulkUpload/matrixExpiryFieldMap.ts",
  );
  const mapSrc = `/* Auto-generated by scripts/sync-training-matrix-columns.mjs — do not edit by hand. */
export const MATRIX_META_EXPIRY_COLUMNS = ${JSON.stringify(
    META.map((m) => ({
      header: m.header,
      code: m.code,
      appKey: m.appKey,
      spField: m.sp,
    })),
    null,
    2,
  )} as const;

export const MATRIX_CATEGORY_EXPIRY_COLUMNS = ${JSON.stringify(
    defs
      .filter((d) => d.code.startsWith("N"))
      .map((d) => ({
        header: d.header,
        code: d.code,
        name: d.name,
        appKey: d.appKey,
        spField: d.spField,
      })),
    null,
    2,
  )} as const;

export const ALL_MATRIX_EXPIRY_APP_KEYS = [
  ...MATRIX_META_EXPIRY_COLUMNS.map((c) => c.appKey),
  ...MATRIX_CATEGORY_EXPIRY_COLUMNS.map((c) => c.appKey),
] as const;
`;
  if (!DRY_RUN) writeFileSync(mapPath, mapSrc);
  console.log(
    DRY_RUN
      ? "Would write matrixExpiryFieldMap.ts"
      : "Wrote matrixExpiryFieldMap.ts",
  );

  console.log("\n========== SYNC SUMMARY ==========");
  console.log(`Template headers: ${headers.length}`);
  console.log(`Expiry columns targeted: ${defs.length}`);
  console.log(`Created: ${created}`);
  console.log(`Renamed display: ${renamed}`);
  console.log(`Already matched: ${skipped}`);
  console.log(`Errors: ${errors.length}`);
  if (errors.length) {
    for (const e of errors.slice(0, 20)) console.log(`  - ${e}`);
    process.exitCode = 1;
  } else {
    console.log("RESULT: PASS — Training Matrix columns aligned to template");
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
