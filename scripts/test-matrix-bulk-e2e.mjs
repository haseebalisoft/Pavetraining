/**
 * E2E: Training Matrix from SharePoint list export
 *   "Sharepoint list _ training matrix example 11-07-26.xlsx"
 *
 * - Does NOT wipe Company / Workforce
 * - Ensures missing matrix candidates exist in Workforce
 * - Wipes Training Matrix + Category Records only
 * - Imports matrix section + N-code category expiries
 * - Verifies against the template section
 *
 * Usage:
 *   node --env-file=.env scripts/test-matrix-bulk-e2e.mjs
 *   node --env-file=.env scripts/test-matrix-bulk-e2e.mjs --verify-only
 */
import { createRequire } from "node:module";
import { resolve } from "node:path";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const DRY_RUN = process.argv.includes("--dry-run");
const VERIFY_ONLY = process.argv.includes("--verify-only");
const EXCEL_PATH = resolve(
  process.cwd(),
  "Sharepoint list _ training matrix example 11-07-26.xlsx",
);

const LISTS = {
  company: process.env.SHAREPOINT_COMPANY_LIST_ID,
  workforce: process.env.SHAREPOINT_WORKFORCE_LIST_ID,
  permissions: process.env.SHAREPOINT_PERMISSIONS_LIST_ID,
  departments: process.env.SHAREPOINT_DEPARTMENTS_LIST_ID,
  trainingMatrix: process.env.SHAREPOINT_TRAINING_MATRIX_LIST_ID,
  matrixCategory: process.env.SHAREPOINT_TRAINING_MATRIX_CATEGORY_RECORDS_LIST_ID,
  nporsCategories: process.env.SHAREPOINT_NPORS_CATEGORIES_LIST_ID,
};

const MATRIX_FIELD_BY_CODE = {
  N001: "N001Expiry",
  N003: "N003Expiry",
  N004: "N004Expiry",
  N010: "N010Expiry",
  N020: "N020Expiry",
  N021: "N021Expiry",
  N027: "N027Expiry",
  N100: "N100Expiry",
};

const META_CODES = [
  { code: "CSCS", header: "CSCS Expiry" },
  { code: "SSSTS", header: "SSSTS Expiry" },
  { code: "SMSTS", header: "SMSTS Expiry" },
  { code: "NRSWA", header: "NRSWA Expiry" },
  { code: "EUSR", header: "EUSR Expiry" },
  { code: "FACEFIT", header: "Face ift" },
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

function nameKey(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function cellToString(value) {
  if (value == null) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  const text = String(value).trim();
  if (!text || /^(—|–|-|n\/?a|null|none)$/i.test(text)) return null;
  return text;
}

function normalizeDateValue(value) {
  if (!value?.trim()) return null;
  const text = value.trim();
  if (/^(—|–|-|n\/?a|null|none)$/i.test(text)) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);

  const uk = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (uk) {
    return `${uk[3]}-${uk[2].padStart(2, "0")}-${uk[1].padStart(2, "0")}`;
  }

  const short = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2})$/);
  if (short) {
    const first = Number(short[1]);
    const second = Number(short[2]);
    let year = Number(short[3]);
    year += year >= 70 ? 1900 : 2000;
    let month;
    let day;
    if (first > 12 && second <= 12) {
      day = first;
      month = second;
    } else {
      month = first;
      day = second;
    }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  const ms = Date.parse(text);
  if (!Number.isNaN(ms)) {
    const parsed = new Date(ms);
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
  }
  return text;
}

function extractCategoryCode(title) {
  if (!title?.trim()) return null;
  const text = title.trim();
  const upper = text.toUpperCase();
  for (const code of ["CSCS", "SSSTS", "SMSTS", "NRSWA", "EUSR", "FACEFIT"]) {
    if (upper === code || upper.startsWith(`${code} `) || upper.startsWith(`${code}-`)) {
      return code;
    }
  }
  return text.match(/^(N\d+[A-Z]?)/i)?.[1]?.toUpperCase() ?? null;
}

async function listAllItems(client, listId) {
  if (!listId) return [];
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

async function deleteItem(client, listId, itemId) {
  if (DRY_RUN) return;
  await client.api(`${siteRoot()}/lists/${listId}/items/${itemId}`).delete();
}

async function createItem(client, listId, fields) {
  if (DRY_RUN) return { id: `dry-${Date.now()}`, fields };
  const created = await client
    .api(`${siteRoot()}/lists/${listId}/items`)
    .post({ fields });
  return { id: String(created.id), fields: created.fields ?? fields };
}

async function updateItem(client, listId, itemId, fields) {
  if (DRY_RUN) return;
  await client
    .api(`${siteRoot()}/lists/${listId}/items/${itemId}/fields`)
    .patch(fields);
}

async function mapPool(items, concurrency, worker) {
  const results = [];
  let index = 0;
  async function run() {
    while (index < items.length) {
      const current = index;
      index += 1;
      results[current] = await worker(items[current], current);
    }
  }
  await Promise.all(
    Array.from({ length: Math.min(concurrency, items.length) }, () => run()),
  );
  return results;
}

function parseMatrixSection() {
  const wb = XLSX.readFile(EXCEL_PATH, { cellDates: true, raw: false });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: null,
    raw: false,
    blankrows: true,
  });

  // SharePoint export may omit the section title cell; find matrix headers by columns.
  let headerIdx = -1;
  for (let i = 0; i < rows.length; i += 1) {
    const cells = (rows[i] || []).map((c) =>
      c == null ? "" : String(c).replace(/\u00a0/g, " ").trim(),
    );
    const lower = new Set(cells.map((c) => c.toLowerCase()));
    const hasNCode = cells.some((c) => /^N\d{3}/i.test(c));
    if (
      lower.has("name") &&
      lower.has("dob") &&
      (hasNCode || lower.has("cscs expiry"))
    ) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) {
    throw new Error(
      "Could not find Training Matrix header row (Name/DOB/N-codes) in file.",
    );
  }

  const headerRow = rows[headerIdx] || [];
  const headers = headerRow
    .map((h) => (h == null ? "" : String(h).replace(/\u00a0/g, " ").trim()))
    .filter(Boolean);

  const data = [];
  for (let i = headerIdx + 1; i < rows.length; i += 1) {
    const raw = rows[i] || [];
    const nonEmpty = raw.filter((x) => cellToString(x));
    if (!nonEmpty.length) continue;
    const first = cellToString(raw[0]);
    if (first.toLowerCase() === "name") break;
    // New section title (single label cell)
    if (
      nonEmpty.length === 1 &&
      raw[0] &&
      !headers.includes(String(raw[0]).trim())
    ) {
      break;
    }
    const row = {};
    for (let c = 0; c < headers.length; c += 1) {
      row[headers[c]] = cellToString(raw[c]);
    }
    if (row.Name) data.push(row);
  }

  return { headers, rows: data, titleRow: headerIdx };
}

function categoryWritesFromRow(row, headers) {
  const writes = [];
  for (const header of headers) {
    if (["Name", "DOB"].includes(header)) continue;
    const value = normalizeDateValue(row[header]);
    if (!value) continue;

    const meta = META_CODES.find((m) => m.header === header);
    if (meta) {
      writes.push({ code: meta.code, name: meta.code, expiryDate: value });
      continue;
    }
    const code = extractCategoryCode(header);
    if (!code) continue;
    const name = header.includes(" - ")
      ? header.split(" - ").slice(1).join(" - ").trim()
      : code;
    writes.push({ code, name, expiryDate: value });
  }
  return writes;
}

async function main() {
  requireEnv("AZURE_TENANT_ID");
  requireEnv("AZURE_CLIENT_ID");
  requireEnv("AZURE_CLIENT_SECRET");
  requireEnv("SHAREPOINT_SITE_ID");
  requireEnv("SHAREPOINT_TRAINING_MATRIX_LIST_ID");
  requireEnv("SHAREPOINT_TRAINING_MATRIX_CATEGORY_RECORDS_LIST_ID");
  requireEnv("SHAREPOINT_WORKFORCE_LIST_ID");
  requireEnv("SHAREPOINT_COMPANY_LIST_ID");
  requireEnv("SHAREPOINT_NPORS_CATEGORIES_LIST_ID");

  const client = getClient();
  const { headers, rows: excelRows } = parseMatrixSection();

  console.log(
    DRY_RUN ? "DRY RUN mode" : VERIFY_ONLY ? "VERIFY ONLY mode" : "LIVE mode",
  );
  console.log(`Excel: ${EXCEL_PATH}`);
  console.log(`Matrix headers: ${headers.length}`);
  console.log(`Matrix data rows: ${excelRows.length}`);
  console.log(
    `Has Face ift column: ${headers.includes("Face ift") ? "yes" : "no (optional)"}`,
  );

  const companies = await listAllItems(client, LISTS.company);
  if (!companies.length) {
    throw new Error("No companies in SharePoint — import Company list first.");
  }
  const defaultCompany = companies[0];
  const companyById = new Map(companies.map((c) => [String(c.id), c]));
  const companyByName = new Map(
    companies.map((c) => [
      nameKey(c.fields?.CompanyName || c.fields?.Title),
      c,
    ]),
  );

  let workforce = await listAllItems(client, LISTS.workforce);
  const workforceByName = new Map();
  for (const item of workforce) {
    const name = item.fields?.CandidateName || item.fields?.Title;
    if (name) workforceByName.set(nameKey(name), item);
  }

  if (!VERIFY_ONLY) {
    // Ensure workforce candidates exist for matrix names (do not wipe companies/workforce).
    console.log("\nEnsuring Workforce candidates for matrix names…");
    for (const row of excelRows) {
      const key = nameKey(row.Name);
      if (workforceByName.has(key)) continue;
      const email = `matrix.${key.replace(/[^a-z0-9]+/g, ".")}.${Date.now()}@pave.local`;
      const fields = {
        Title: row.Name,
        CandidateName: row.Name,
        CompanyNameLookupId: Number(defaultCompany.id),
        Email: email,
        Dateofbirth: normalizeDateValue(row.DOB),
        Status: "Active",
        Department0LookupId: null,
      };
      for (const [k, v] of Object.entries(fields)) {
        if (v == null || v === "") delete fields[k];
      }
      const created = await createItem(client, LISTS.workforce, fields);
      workforceByName.set(key, created);
      workforce.push(created);
      console.log(
        `  created workforce ${row.Name} → company ${defaultCompany.fields?.CompanyName}`,
      );
    }

    // Wipe matrix + category records only
    const existingMatrix = await listAllItems(client, LISTS.trainingMatrix);
    const existingCats = await listAllItems(client, LISTS.matrixCategory);
    console.log(
      `\nWiping matrix=${existingMatrix.length}, category records=${existingCats.length}`,
    );
    let wipedCats = 0;
    await mapPool(existingCats, 12, async (item) => {
      try {
        await deleteItem(client, LISTS.matrixCategory, item.id);
        wipedCats += 1;
        if (wipedCats % 100 === 0) {
          console.log(`  wiped categories ${wipedCats}/${existingCats.length}`);
        }
      } catch (error) {
        console.warn(`  cat #${item.id}:`, error?.message || String(error));
      }
    });
    await mapPool(existingMatrix, 8, async (item) => {
      try {
        await deleteItem(client, LISTS.trainingMatrix, item.id);
      } catch (error) {
        console.warn(
          `  matrix #${item.id}:`,
          error?.message || String(error),
        );
      }
    });

    const afterMatrix = DRY_RUN
      ? existingMatrix
      : await listAllItems(client, LISTS.trainingMatrix);
    const afterCats = DRY_RUN
      ? existingCats
      : await listAllItems(client, LISTS.matrixCategory);
    console.log(
      `After wipe matrix=${afterMatrix.length}, categories=${afterCats.length}`,
    );
    if (!DRY_RUN && (afterMatrix.length || afterCats.length)) {
      throw new Error("Matrix wipe incomplete — aborting.");
    }

    // Load NPORS category lookup
    const nporsCats = await listAllItems(client, LISTS.nporsCategories);
    const categoryByCode = new Map();
    for (const item of nporsCats) {
      const code = extractCategoryCode(item.fields?.Title);
      if (code && !categoryByCode.has(code)) categoryByCode.set(code, item.id);
    }
    async function ensureCategory(code, name) {
      if (categoryByCode.has(code)) return categoryByCode.get(code);
      const title = /^N\d+/i.test(code) ? `${code} - ${name}` : code;
      const created = await createItem(client, LISTS.nporsCategories, {
        Title: title,
        Category_x0020_Code: code,
        Active: true,
        Customer_x0020_Visible: true,
      });
      categoryByCode.set(code, created.id);
      return created.id;
    }

    // Import matrix rows
    let imported = 0;
    let categoriesWritten = 0;
    const importErrors = [];

    for (const [index, row] of excelRows.entries()) {
      const rowNumber = index + 2;
      try {
        const wf = workforceByName.get(nameKey(row.Name));
        if (!wf) throw new Error(`Workforce missing for ${row.Name}`);
        const companyId =
          String(wf.fields?.CompanyNameLookupId ?? "") ||
          String(defaultCompany.id);
        const company = companyById.get(String(companyId)) || defaultCompany;

        const matrixFields = {
          CandidateNameLookupId: Number(wf.id),
          MatrixCompanyLookupId: Number(company.id),
          NeedsReview: false,
        };

        const writes = categoryWritesFromRow(row, headers);
        // earliest next expiry
        const times = writes
          .map((w) => new Date(w.expiryDate).getTime())
          .filter((t) => !Number.isNaN(t));
        if (times.length) {
          matrixFields.NextExpiryDate = new Date(Math.min(...times))
            .toISOString()
            .slice(0, 10);
        }
        for (const write of writes) {
          const field = MATRIX_FIELD_BY_CODE[write.code];
          if (field) matrixFields[field] = write.expiryDate;
        }

        // Sync CSCS / EUSR / NRSWA onto Workforce so matrix lookup columns populate.
        const wfPatch = {};
        const cscs = normalizeDateValue(row["CSCS Expiry"]);
        const eusr = normalizeDateValue(row["EUSR Expiry"]);
        const nrswa = normalizeDateValue(row["NRSWA Expiry"]);
        if (cscs) wfPatch.CscsExpiry = cscs;
        if (eusr) wfPatch.EusrExpiry = eusr;
        if (nrswa) wfPatch.SwqrExpiry = nrswa;
        if (Object.keys(wfPatch).length) {
          await updateItem(client, LISTS.workforce, wf.id, wfPatch);
        }

        await createItem(client, LISTS.trainingMatrix, matrixFields);

        // Category records (parallel-ish)
        await mapPool(writes, 6, async (write) => {
          const categoryId = await ensureCategory(write.code, write.name);
          const payload = {
            Title: `${row.Name} · ${write.code}`.slice(0, 240),
            Candidate_x0020_NameLookupId: Number(wf.id),
            Company_x0020_NameLookupId: Number(company.id),
            Category_x0020_CodeLookupId: Number(categoryId),
            Category_x0020_Name: write.name,
            Expiry_x0020_Date: write.expiryDate,
            Status: "Active",
            Customer_x0020_Visible: true,
          };
          await createItem(client, LISTS.matrixCategory, payload);
          categoriesWritten += 1;
        });

        imported += 1;
        console.log(
          `  imported ${imported}/${excelRows.length} ${row.Name} (${writes.length} categories)`,
        );
      } catch (error) {
        const message = error?.body
          ? typeof error.body === "string"
            ? error.body
            : JSON.stringify(error.body)
          : error?.message || String(error);
        importErrors.push(`Row ${rowNumber} (${row.Name}): ${message}`);
        console.error(`  FAIL ${row.Name}: ${message}`);
      }
    }

    console.log(`\nMatrix imported: ${imported}/${excelRows.length}`);
    console.log(`Category records written: ${categoriesWritten}`);
    if (importErrors.length) {
      console.log(`Import errors (${importErrors.length}):`);
      importErrors.slice(0, 20).forEach((e) => console.log(" -", e));
    }
  }

  if (DRY_RUN) {
    console.log("Dry run complete.");
    return;
  }

  // Verify
  workforce = await listAllItems(client, LISTS.workforce);
  const matrix = await listAllItems(client, LISTS.trainingMatrix);
  const cats = await listAllItems(client, LISTS.matrixCategory);
  const nporsCats = await listAllItems(client, LISTS.nporsCategories);
  const codeByCatId = new Map();
  for (const item of nporsCats) {
    const code = extractCategoryCode(item.fields?.Title);
    if (code) codeByCatId.set(String(item.id), code);
  }
  const wfById = new Map(
    workforce.map((w) => [
      String(w.id),
      w.fields?.CandidateName || w.fields?.Title || "",
    ]),
  );

  console.log(
    `\nAfter import: matrix=${matrix.length}, categories=${cats.length}, workforce=${workforce.length}`,
  );

  const byCandidate = new Map();
  for (const item of cats) {
    const candId = String(item.fields?.Candidate_x0020_NameLookupId ?? "");
    const catId = String(item.fields?.Category_x0020_CodeLookupId ?? "");
    const code = codeByCatId.get(catId);
    const expiry = normalizeDateValue(
      cellToString(item.fields?.Expiry_x0020_Date),
    );
    const candName = wfById.get(candId);
    if (!candName || !code || !expiry) continue;
    const bucket = byCandidate.get(nameKey(candName)) ?? new Map();
    bucket.set(code, expiry);
    byCandidate.set(nameKey(candName), bucket);
  }

  let matchedRows = 0;
  const mismatches = [];
  const missing = [];

  for (const expected of excelRows) {
    const matrixHit = matrix.find((item) => {
      const candId = String(item.fields?.CandidateNameLookupId ?? "");
      return nameKey(wfById.get(candId)) === nameKey(expected.Name);
    });
    if (!matrixHit) {
      missing.push(expected.Name);
      continue;
    }

    const expectedWrites = categoryWritesFromRow(expected, headers);
    const actualCodes = byCandidate.get(nameKey(expected.Name)) ?? new Map();
    let rowOk = true;

    // Wide matrix N-code fields
    for (const write of expectedWrites) {
      const field = MATRIX_FIELD_BY_CODE[write.code];
      if (field) {
        const actual = normalizeDateValue(
          cellToString(matrixHit.fields?.[field]),
        );
        if (actual !== write.expiryDate) {
          rowOk = false;
          mismatches.push({
            candidate: expected.Name,
            field,
            expected: write.expiryDate,
            actual,
          });
        }
      }
      const catExpiry = actualCodes.get(write.code) ?? null;
      if (catExpiry !== write.expiryDate) {
        rowOk = false;
        mismatches.push({
          candidate: expected.Name,
          field: `category:${write.code}`,
          expected: write.expiryDate,
          actual: catExpiry,
        });
      }
    }

    if (rowOk) matchedRows += 1;
  }

  console.log("\n========== VERIFY SUMMARY ==========");
  console.log(`Excel matrix rows: ${excelRows.length}`);
  console.log(`SharePoint matrix rows: ${matrix.length}`);
  console.log(`Fully matched rows: ${matchedRows}`);
  console.log(`Missing matrix rows: ${missing.length}`);
  console.log(`Field/category mismatches: ${mismatches.length}`);

  if (missing.length) {
    console.log("Missing:", missing.join(", "));
  }
  if (mismatches.length) {
    console.log("Mismatches (first 30):");
    mismatches.slice(0, 30).forEach((m) => {
      console.log(
        ` - ${m.candidate} / ${m.field}: expected=${m.expected} actual=${m.actual}`,
      );
    });
  }

  const companiesStill = await listAllItems(client, LISTS.company);
  console.log(
    `\nSafety: companies still ${companiesStill.length} (not wiped)`,
  );

  if (
    missing.length ||
    mismatches.length ||
    matchedRows !== excelRows.length ||
    matrix.length < excelRows.length
  ) {
    process.exitCode = 1;
    console.log("\nRESULT: FAIL");
  } else {
    console.log("\nRESULT: PASS — SharePoint matrix template imported cleanly");
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
