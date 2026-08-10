/**
 * Local preview test for Training Matrix upsert + logging.
 *
 *   node --env-file=.env.local --experimental-strip-types --import ./scripts/register-server-only-stub.mjs scripts/test-matrix-upload-preview.mjs
 */
import { readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { resolve } from "node:path";

process.env.BULK_UPLOAD_LOGS = process.env.BULK_UPLOAD_LOGS || "verbose";

const fileArg = process.argv.find((a) => a.startsWith("--file="));
const filePath = resolve(
  process.cwd(),
  fileArg ? fileArg.slice("--file=".length) : "Training matrix example.xlsx",
);

function loadSrc(rel) {
  return import(pathToFileURL(resolve(process.cwd(), "src", rel)).href);
}

const bytes = readFileSync(filePath);

const { parseSpreadsheetBuffer } = await loadSrc(
  "lib/services/bulkUpload/parseSpreadsheet.ts",
);
const { previewMatrixImport } = await loadSrc(
  "lib/services/bulkUpload/matrixImporter.ts",
);
const { summarizeBulkRows } = await loadSrc(
  "lib/services/bulkUpload/candidateImporter.ts",
);

const spreadsheet = parseSpreadsheetBuffer(bytes, filePath);
console.log(
  `\nParsed ${spreadsheet.rows.length} data rows from ${filePath}`,
);
console.log("Headers (first 12):", spreadsheet.headers.slice(0, 12).join(" | "));

const first = spreadsheet.rows[0] ?? {};
console.log("\nFirst row Name/DOB/CSCS raw:", {
  Name: first.Name ?? first["Candidate Name"],
  DOB: first.DOB,
  "CSCS Expiry": first["CSCS Expiry"],
});

const rows = await previewMatrixImport(spreadsheet);
const summary = summarizeBulkRows(rows);

console.log("\n========== PREVIEW SUMMARY ==========");
console.log(summary);
console.log(
  "Override (matchedEntityId set):",
  rows.filter((r) => r.matchedEntityId).length,
);
console.log(
  "Create (no match):",
  rows.filter((r) => !r.matchedEntityId && r.status !== "Error").length,
);

const errors = rows.filter((r) => r.status === "Error").slice(0, 10);
if (errors.length) {
  console.log("\nFirst errors:");
  for (const row of errors) {
    console.log(`  row ${row.rowNumber}: ${row.messages.join(" | ")}`);
  }
}

console.log("\nFirst 5 row statuses:");
for (const row of rows.slice(0, 5)) {
  console.log(
    `  #${row.rowNumber} ${row.status} name=${row.fields.candidateName} dob=${row.fields.dateOfBirth} :: ${row.messages[0]}`,
  );
}

if (summary.errorRows === summary.totalRows && summary.totalRows > 0) {
  console.error("\nFAIL: every row is Error — Confirm import would stay blocked.");
  process.exitCode = 1;
} else if (summary.readyRows + summary.warningRows === 0) {
  console.error("\nFAIL: no Ready/Warning rows.");
  process.exitCode = 1;
} else {
  console.log(
    "\nOK: importable rows present (Ready/Warning). Confirm import should enable.",
  );
}
