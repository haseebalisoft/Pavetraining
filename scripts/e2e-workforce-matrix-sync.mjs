/**
 * LIVE end-to-end test of two-way Workforce <-> Training Matrix sync through the
 * REAL bulk-upload pipeline (previewBulkUpload + commitBulkUpload — the exact
 * code the admin UI calls).
 *
 * Round A: wipe -> upload Workforce -> upload Training Matrix
 * Round B: wipe -> upload Training Matrix FIRST -> upload Workforce
 *
 * Companies, Departments and Permissions are NEVER touched.
 *
 * Usage (from repo root):
 *   node --env-file=.env.local --conditions=react-server \
 *     --import ./scripts/_register-e2e-hook.mjs \
 *     scripts/e2e-workforce-matrix-sync.mjs --dry-run
 *
 *   node --env-file=.env.local --conditions=react-server \
 *     --import ./scripts/_register-e2e-hook.mjs \
 *     scripts/e2e-workforce-matrix-sync.mjs
 *
 * Flags:
 *   --dry-run      report current row counts and exit (no writes, no deletes)
 *   --round=a|b    run only one round
 *   --no-wipe      skip the wipes (upload onto whatever is there)
 */
import { readFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const DRY_RUN = process.argv.includes("--dry-run");
const NO_WIPE = process.argv.includes("--no-wipe");
const ROUND_ARG = process.argv.find((arg) => arg.startsWith("--round="));
const ONLY_ROUND = ROUND_ARG ? ROUND_ARG.split("=")[1]?.toLowerCase() : null;

const WORKFORCE_FILE = "Workforce list.xlsx";
const MATRIX_FILE = "Training matrix example.xlsx";

// ---------------------------------------------------------------------------
// Graph plumbing (wipe + independent verification, separate from the app code)
// ---------------------------------------------------------------------------

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

const credential = new ClientSecretCredential(
  requireEnv("AZURE_TENANT_ID"),
  requireEnv("AZURE_CLIENT_ID"),
  requireEnv("AZURE_CLIENT_SECRET"),
);
const graph = Client.initWithMiddleware({
  authProvider: new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  }),
});

const WORKFORCE_LIST_ID = requireEnv("SHAREPOINT_WORKFORCE_LIST_ID");
const MATRIX_LIST_ID = requireEnv(
  "SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID",
);
const CATEGORY_LIST_ID =
  process.env.SHAREPOINT_TRAINING_MATRIX_CATEGORY_RECORDS_LIST_ID?.trim() || null;

async function listItemIds(listId) {
  const ids = [];
  let url = `${siteRoot()}/lists/${listId}/items?$top=200&$select=id`;
  while (url) {
    const res = await graph
      .api(url.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, ""))
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .get();
    ids.push(...(res.value ?? []).map((item) => item.id));
    url = res["@odata.nextLink"] || null;
  }
  return ids;
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

/**
 * Deletes every row it can. SharePoint refuses to delete a row that another list
 * still references through an enforced lookup ("related to another item in
 * another list") — those survive as a baseline rather than aborting the test.
 */
async function wipeList(label, listId) {
  const ids = await listItemIds(listId);
  if (!ids.length) {
    console.log(`  ${label}: already empty`);
    return { deleted: 0, remainingIds: [] };
  }
  let failed = 0;
  const errors = [];
  await mapPool(ids, 8, async (id) => {
    try {
      await graph.api(`${siteRoot()}/lists/${listId}/items/${id}`).delete();
    } catch (error) {
      failed += 1;
      errors.push(`#${id}: ${error?.message ?? String(error)}`);
    }
  });
  const remainingIds = await listItemIds(listId);
  console.log(
    `  ${label}: deleted ${ids.length - failed}/${ids.length}, remaining ${remainingIds.length}`,
  );
  for (const err of errors.slice(0, 3)) {
    console.log(`    kept ${err}`);
  }
  if (errors.length > 3) console.log(`    … and ${errors.length - 3} more`);
  return { deleted: ids.length - failed, remainingIds };
}

// ---------------------------------------------------------------------------
// App pipeline
// ---------------------------------------------------------------------------

const { previewBulkUpload, commitBulkUpload } = await import(
  "@/lib/services/bulkUpload/bulkUploadService"
);
const { listTrainingMatrixExampleRows } = await import(
  "@/lib/services/bulkUpload/trainingMatrixExampleService"
);
const { listAdminWorkforce, listAdminMatrix } = await import(
  "@/lib/services/adminCrudService"
);
const { candidateNameKey, isoDateKey } = await import(
  "@/lib/services/bulkUpload/workforceMatrixSync"
);

function loadFile(fileName) {
  const bytes = readFileSync(resolvePath(process.cwd(), fileName));
  return new File([bytes], fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

/** Mirrors /api/admin/bulk-upload/commit: rowNumber + fields + source. */
function toCommitRows(previewRows) {
  return previewRows.map((row) => ({
    rowNumber: row.rowNumber,
    fields: row.fields,
    source: row.source,
  }));
}

function summarize(label, result) {
  const s = result.summary;
  console.log(
    `  ${label}: total=${s.totalRows} ready=${s.readyRows} warn=${s.warningRows} dup=${s.duplicateRows} err=${s.errorRows} skip=${s.skippedRows} imported=${s.importedRows}`,
  );
  const byOutcome = new Map();
  for (const row of result.rows) {
    const key = row.linkOutcome ?? "(none)";
    byOutcome.set(key, (byOutcome.get(key) ?? 0) + 1);
  }
  if (byOutcome.size) {
    console.log(
      `    link: ${[...byOutcome.entries()].map(([k, v]) => `${k}=${v}`).join(" ")}`,
    );
  }
  return s;
}

function reportProblems(label, rows, statuses = ["Error"]) {
  const bad = rows.filter((row) => statuses.includes(row.status));
  if (!bad.length) {
    console.log(`    no ${statuses.join("/")} rows ✓`);
    return [];
  }
  console.log(`    ${bad.length} ${statuses.join("/")} row(s) in ${label}:`);
  for (const row of bad.slice(0, 12)) {
    console.log(
      `      row ${row.rowNumber} [${row.status}] ${row.fields.candidateName ?? ""} — ${row.messages.join(" | ")}`,
    );
  }
  if (bad.length > 12) console.log(`      … and ${bad.length - 12} more`);
  return bad;
}

function seconds(startedAt) {
  return `${((Date.now() - startedAt) / 1000).toFixed(1)}s`;
}

async function runUpload(importType, fileName, { duplicateMode = "update" } = {}) {
  const file = loadFile(fileName);

  const previewStart = Date.now();
  const preview = await previewBulkUpload({
    importType,
    file,
    autoCreateMissingCompanies: true,
  });
  const previewMs = seconds(previewStart);
  summarize(`preview ${importType}`, preview);
  console.log(`    preview took ${previewMs}`);
  const previewErrors = reportProblems(`preview ${importType}`, preview.rows);

  const commitStart = Date.now();
  const commit = await commitBulkUpload({
    importType,
    fileName,
    duplicateMode,
    suppressNotifications: true,
    autoCreateMissingCompanies: true,
    rows: toCommitRows(preview.rows),
  });
  const commitMs = seconds(commitStart);
  summarize(`commit  ${importType}`, commit);
  console.log(`    commit took ${commitMs}`);
  const commitErrors = reportProblems(`commit ${importType}`, commit.rows);

  // Blank trailing rows in the client export must be skipped, never errors.
  const skipped = commit.rows.filter((row) => row.status === "Skipped").length;
  if (skipped) console.log(`    skipped (blank/duplicate) rows: ${skipped}`);

  return { preview, commit, previewErrors, commitErrors };
}

// ---------------------------------------------------------------------------
// Verification against the live lists
// ---------------------------------------------------------------------------

const failures = [];
function check(condition, message) {
  if (condition) {
    console.log(`    PASS  ${message}`);
  } else {
    console.log(`    FAIL  ${message}`);
    failures.push(message);
  }
}

async function verifyLinked({ expectAllLinked, allowUnlinked = 0 }) {
  const [rows, workforce] = await Promise.all([
    listTrainingMatrixExampleRows(),
    listAdminWorkforce(),
  ]);
  console.log(
    `  live: workforce=${workforce.length} matrixRows=${rows.length}`,
  );

  const linked = rows.filter((row) => row.matrixLinkStatus === "Linked");
  const needsReview = rows.filter(
    (row) => row.matrixLinkStatus === "Needs Review",
  );
  const other = rows.filter(
    (row) =>
      row.matrixLinkStatus !== "Linked" &&
      row.matrixLinkStatus !== "Needs Review",
  );
  console.log(
    `  status: Linked=${linked.length} NeedsReview=${needsReview.length} other=${other.length}`,
  );

  // No duplicate matrix rows for one candidate.
  const keys = rows.map(
    (row) =>
      `${candidateNameKey(row.candidateName)}|${isoDateKey(row.dateOfBirth)}`,
  );
  const dupKeys = keys.filter((key, index) => keys.indexOf(key) !== index);
  check(dupKeys.length === 0, `no duplicate Name+DOB matrix rows (found ${dupKeys.length})`);
  if (dupKeys.length) {
    console.log(`      e.g. ${[...new Set(dupKeys)].slice(0, 5).join(", ")}`);
  }

  // One matrix row per workforce owner.
  const owners = rows
    .map((row) => String(row.workforceItemId ?? "").trim())
    .filter(Boolean);
  check(
    new Set(owners).size === owners.length,
    `each WorkforceItemId owns at most one matrix row (${owners.length} owned, ${new Set(owners).size} distinct)`,
  );

  if (expectAllLinked) {
    check(
      needsReview.length <= allowUnlinked,
      `no Needs Review rows left (found ${needsReview.length}, allowance ${allowUnlinked})`,
    );
    check(
      linked.length >= rows.length - allowUnlinked,
      `every matrix row is Linked (${linked.length}/${rows.length})`,
    );
    check(
      owners.length >= rows.length - allowUnlinked,
      `every matrix row stores WorkforceItemId (${owners.length}/${rows.length})`,
    );

    // Every owner id is a real workforce record.
    const wfIds = new Set(workforce.map((row) => String(row.id)));
    const dangling = owners.filter((id) => !wfIds.has(id));
    check(dangling.length === 0, `no dangling WorkforceItemId (found ${dangling.length})`);

    // Link fields populated.
    const missingNumber = rows.filter((row) => !row.workforceNumber?.trim());
    check(
      missingNumber.length <= allowUnlinked,
      `every row stores WorkforceNumber (missing ${missingNumber.length}, allowance ${allowUnlinked})`,
    );
    const missingCompany = rows.filter(
      (row) => !String(row.companyItemId ?? "").trim(),
    );
    check(
      missingCompany.length <= allowUnlinked,
      `every row stores CompanyItemId (missing ${missingCompany.length}, allowance ${allowUnlinked})`,
    );
  }

  // Uploaded expiry data actually landed.
  const withExpiry = rows.filter((row) =>
    Object.entries(row.columnValues).some(
      ([header, value]) =>
        header !== "Name" && header !== "DOB" && value?.trim(),
    ),
  );
  console.log(`  rows carrying at least one expiry value: ${withExpiry.length}`);

  return { rows, workforce, linked, needsReview, withExpiry };
}

async function verifyAdminMatrixVisibility(expectedHiddenAtLeast) {
  const visible = await listAdminMatrix();
  const all = await listAdminMatrix(null, { includeUnlinked: true });
  console.log(
    `  Admin Matrix: default=${visible.length} showAll=${all.length}`,
  );
  if (expectedHiddenAtLeast > 0) {
    check(
      all.length - visible.length >= expectedHiddenAtLeast,
      `unlinked rows hidden by default (hidden=${all.length - visible.length}, expected >= ${expectedHiddenAtLeast})`,
    );
  }
  return { visible, all };
}

// ---------------------------------------------------------------------------
// Rounds
// ---------------------------------------------------------------------------

/**
 * Deletes rows from EXACTLY two lists: Workforce List and Training Matrix
 * Update. Companies, Departments, Permissions and Matrix Category Records are
 * never touched — Category Records is only reported, so its backup rows survive.
 */
async function wipeBoth() {
  console.log("\n-- wipe (Workforce + Training Matrix Update only) --");
  const matrix = await wipeList("Training Matrix Update", MATRIX_LIST_ID);
  const workforce = await wipeList("Workforce List", WORKFORCE_LIST_ID);
  if (CATEGORY_LIST_ID) {
    const kept = (await listItemIds(CATEGORY_LIST_ID)).length;
    console.log(`  Matrix Category Records: left untouched (${kept} rows)`);
  }
  const baseline = {
    workforce: workforce.remainingIds.length,
    matrix: matrix.remainingIds.length,
  };
  if (baseline.workforce || baseline.matrix) {
    console.log(
      `  baseline after wipe — workforce=${baseline.workforce} matrix=${baseline.matrix} (SharePoint refused to delete these)`,
    );
  }

  // This wipe deletes via raw Graph calls, bypassing the app's
  // unstable_cache/revalidateTag path entirely. The app's cached workforce/
  // matrix list reads (45s TTL) can still be serving the pre-wipe snapshot to
  // the very next preview/commit call, which misreports fresh creates as
  // "duplicate updates". Wait out the TTL so the app's next read is live.
  const CACHE_SETTLE_MS = 46_000;
  console.log(`  waiting ${CACHE_SETTLE_MS / 1000}s for the app read-cache to expire...`);
  await new Promise((resolve) => setTimeout(resolve, CACHE_SETTLE_MS));

  return baseline;
}

async function roundA() {
  console.log("\n=================================================");
  console.log("ROUND A — Workforce first, then Training Matrix");
  console.log("=================================================");
  const baseline = NO_WIPE
    ? { workforce: 0, matrix: 0 }
    : await wipeBoth();

  console.log("\n-- upload 1/2: Workforce --");
  const wf = await runUpload("workforce", WORKFORCE_FILE);
  check(
    wf.commit.summary.errorRows === 0,
    `Workforce commit had no Error rows (${wf.commit.summary.errorRows})`,
  );
  check(
    wf.commit.summary.importedRows > 0,
    `Workforce imported rows > 0 (${wf.commit.summary.importedRows})`,
  );

  console.log("\n-- after Workforce upload --");
  const afterWf = await verifyLinked({
    expectAllLinked: true,
    allowUnlinked: baseline.matrix,
  });
  check(
    afterWf.rows.length === wf.commit.summary.importedRows + baseline.matrix,
    `one matrix row per imported candidate (${afterWf.rows.length} rows / ${wf.commit.summary.importedRows} imported + ${baseline.matrix} baseline)`,
  );

  console.log("\n-- upload 2/2: Training Matrix --");
  const mx = await runUpload("trainingMatrix", MATRIX_FILE);
  check(
    mx.preview.summary.errorRows === 0,
    `Matrix preview had no Error rows (${mx.preview.summary.errorRows})`,
  );
  check(
    mx.commit.summary.errorRows === 0,
    `Matrix commit had no Error rows (${mx.commit.summary.errorRows})`,
  );
  check(
    mx.commit.summary.warningRows === 0,
    `Matrix commit reported no spurious warnings (${mx.commit.summary.warningRows})`,
  );

  console.log("\n-- after Matrix upload --");
  const afterMx = await verifyLinked({
    expectAllLinked: true,
    allowUnlinked: baseline.matrix,
  });
  check(
    afterMx.rows.length === afterWf.rows.length,
    `Matrix upload created NO new rows (${afterWf.rows.length} -> ${afterMx.rows.length})`,
  );
  check(
    afterMx.withExpiry.length >= afterWf.withExpiry.length,
    `expiry data present after matrix upload (${afterWf.withExpiry.length} -> ${afterMx.withExpiry.length})`,
  );

  console.log("\n-- re-upload Matrix (idempotency) --");
  const mx2 = await runUpload("trainingMatrix", MATRIX_FILE);
  check(
    mx2.commit.summary.errorRows === 0,
    `Matrix re-upload had no Error rows (${mx2.commit.summary.errorRows})`,
  );
  const afterMx2 = await verifyLinked({
    expectAllLinked: true,
    allowUnlinked: baseline.matrix,
  });
  check(
    afterMx2.rows.length === afterMx.rows.length,
    `Matrix re-upload created NO new rows (${afterMx.rows.length} -> ${afterMx2.rows.length})`,
  );

  await verifyAdminMatrixVisibility(0);
  return afterMx2;
}

async function roundB() {
  console.log("\n=================================================");
  console.log("ROUND B — Training Matrix FIRST, then Workforce");
  console.log("=================================================");
  const baseline = NO_WIPE
    ? { workforce: 0, matrix: 0 }
    : await wipeBoth();

  console.log("\n-- upload 1/2: Training Matrix (empty Workforce) --");
  const mx = await runUpload("trainingMatrix", MATRIX_FILE);
  check(
    mx.preview.summary.errorRows === 0,
    `Matrix-first preview had no Error rows (${mx.preview.summary.errorRows})`,
  );
  check(
    mx.commit.summary.errorRows === 0,
    `Matrix-first commit had no Error rows (${mx.commit.summary.errorRows})`,
  );
  const needsReviewRows = mx.commit.rows.filter(
    (row) => row.linkOutcome === "needsReviewNoMatch",
  );
  const populatedRows = mx.commit.rows.filter(
    (row) => row.status !== "Skipped",
  ).length;
  check(
    needsReviewRows.length === populatedRows,
    `every populated matrix row reported "no Workforce match" (${needsReviewRows.length}/${populatedRows})`,
  );

  console.log("\n-- after Matrix-first upload --");
  const afterMx = await verifyLinked({ expectAllLinked: false });
  check(
    afterMx.needsReview.length >= afterMx.rows.length - baseline.matrix,
    `all rows stored as Needs Review (${afterMx.needsReview.length}/${afterMx.rows.length})`,
  );
  check(
    afterMx.withExpiry.length >= afterMx.rows.length - baseline.matrix,
    `uploaded expiry data preserved on unlinked rows (${afterMx.withExpiry.length}/${afterMx.rows.length})`,
  );
  await verifyAdminMatrixVisibility(1);

  console.log("\n-- upload 2/2: Workforce (must ADOPT the Needs Review rows) --");
  const wf = await runUpload("workforce", WORKFORCE_FILE);
  check(
    wf.commit.summary.errorRows === 0,
    `Workforce commit had no Error rows (${wf.commit.summary.errorRows})`,
  );
  const adopted = wf.commit.rows.filter(
    (row) =>
      row.linkOutcome === "linkedExistingNeedsReview" ||
      row.linkOutcome === "linkedNameDob" ||
      row.linkOutcome === "linkedCompanyNameDob",
  );
  const createdNew = wf.commit.rows.filter(
    (row) => row.linkOutcome === "createdLinked",
  );
  const skipped = wf.commit.rows.filter(
    (row) => row.linkOutcome === "skippedAmbiguous",
  );
  console.log(
    `    adopted=${adopted.length} createdNew=${createdNew.length} skippedAmbiguous=${skipped.length}`,
  );
  check(
    adopted.length >= afterMx.rows.length,
    `every Needs Review row was adopted (adopted=${adopted.length}, unlinked rows were ${afterMx.rows.length})`,
  );
  check(skipped.length === 0, `no candidate skipped as ambiguous (${skipped.length})`);

  console.log("\n-- after Workforce upload --");
  const afterWf = await verifyLinked({
    expectAllLinked: true,
    allowUnlinked: baseline.matrix,
  });
  check(
    afterWf.rows.length === wf.commit.summary.importedRows + baseline.matrix,
    `no duplicate rows created (${afterWf.rows.length} rows / ${wf.commit.summary.importedRows} candidates + ${baseline.matrix} baseline)`,
  );
  check(
    afterWf.withExpiry.length >= afterMx.withExpiry.length,
    `matrix expiry data survived the adoption (${afterMx.withExpiry.length} -> ${afterWf.withExpiry.length})`,
  );
  await verifyAdminMatrixVisibility(0);
  return afterWf;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  console.log("Workforce list:", WORKFORCE_LIST_ID);
  console.log("Training Matrix Update:", MATRIX_LIST_ID);
  console.log("Matrix Category Records:", CATEGORY_LIST_ID ?? "(not set)");

  const [wfIds, mxIds] = await Promise.all([
    listItemIds(WORKFORCE_LIST_ID),
    listItemIds(MATRIX_LIST_ID),
  ]);
  console.log(
    `\nCurrent rows — Workforce: ${wfIds.length}, Training Matrix Update: ${mxIds.length}`,
  );
  if (CATEGORY_LIST_ID) {
    console.log(
      `Current rows — Matrix Category Records: ${(await listItemIds(CATEGORY_LIST_ID)).length}`,
    );
  }

  if (DRY_RUN) {
    console.log(
      "\nDRY RUN — nothing deleted, nothing uploaded. Re-run without --dry-run to execute.",
    );
    return;
  }

  if (!ONLY_ROUND || ONLY_ROUND === "a") await roundA();
  if (!ONLY_ROUND || ONLY_ROUND === "b") await roundB();

  console.log("\n================ RESULT ================");
  if (failures.length) {
    console.log(`${failures.length} CHECK(S) FAILED:`);
    for (const failure of failures) console.log(` - ${failure}`);
    process.exitCode = 1;
  } else {
    console.log("ALL CHECKS PASSED");
  }
}

main().catch((error) => {
  console.error("\nE2E ABORTED:", error?.stack || error);
  process.exitCode = 1;
});
