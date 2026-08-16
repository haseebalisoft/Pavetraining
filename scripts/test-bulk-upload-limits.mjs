/**
 * Unit tests for the bulk-upload row-limit helpers.
 *
 * Runs with the built-in Node test runner + native TypeScript type-stripping
 * (Node >= 22); bulkUploadLimits.ts has only a type-only import, so it loads
 * standalone with no Graph and no "@/" resolver hook:
 *
 *   node --test scripts/test-bulk-upload-limits.mjs
 *
 * Workforce (and every other type) is unlimited today. The helper still
 * enforces a cap if one is configured, and blank rows never count.
 */
import test from "node:test";
import assert from "node:assert/strict";

const BASE = new URL("../src/lib/", import.meta.url).pathname;
const {
  bulkUploadRowLimit,
  bulkUploadRowLimitError,
  checkBulkUploadRowLimit,
  countPopulatedRows,
  isEmptyFieldSet,
} = await import(BASE + "services/bulkUpload/bulkUploadLimits.ts");

/** Build N populated commit-style rows, then M blank ones. */
function rows(populated, blank = 0) {
  const out = [];
  for (let i = 0; i < populated; i += 1) {
    out.push({ rowNumber: i + 1, fields: { candidateName: `Person ${i + 1}` } });
  }
  for (let i = 0; i < blank; i += 1) {
    out.push({
      rowNumber: populated + i + 1,
      fields: { candidateName: null, company: "   ", workforceNumber: "" },
    });
  }
  return out;
}

// --- Limit configuration --------------------------------------------------

test("workforce and company have no row cap", () => {
  assert.equal(bulkUploadRowLimit("workforce"), null);
  assert.equal(bulkUploadRowLimit("company"), null);
});

test("error helper still formats a clear message if a cap is configured later", () => {
  assert.equal(
    bulkUploadRowLimitError("workforce", 50),
    "Workforce bulk upload supports a maximum of 50 records per upload.",
  );
});

// --- Unlimited: any populated count is accepted --------------------------

test("50 workforce rows are accepted", () => {
  const check = checkBulkUploadRowLimit("workforce", 50);
  assert.equal(check.ok, true);
  assert.equal(check.error, null);
  assert.equal(check.limit, null);
});

test("51 workforce rows are accepted (no cap)", () => {
  const check = checkBulkUploadRowLimit("workforce", 51);
  assert.equal(check.ok, true);
  assert.equal(check.error, null);
});

test("large workforce batches are accepted (no cap)", () => {
  const check = checkBulkUploadRowLimit("workforce", 1000);
  assert.equal(check.ok, true);
  assert.equal(check.error, null);
});

// --- Empty rows do not count ---------------------------------------------

test("isEmptyFieldSet treats null/blank/whitespace-only field sets as empty", () => {
  assert.equal(isEmptyFieldSet(null), true);
  assert.equal(isEmptyFieldSet({}), true);
  assert.equal(isEmptyFieldSet({ a: null, b: "", c: "   " }), true);
  assert.equal(isEmptyFieldSet({ a: null, b: "x" }), false);
});

test("blank rows do not count toward populated-row totals", () => {
  const count = countPopulatedRows(rows(50, 20));
  assert.equal(count, 50);
  assert.equal(checkBulkUploadRowLimit("workforce", count).ok, true);
});

test("populated-row count ignores blank padding", () => {
  assert.equal(countPopulatedRows(rows(51, 0)), 51);
  assert.equal(countPopulatedRows(rows(51, 30)), 51);
});

// --- Company stays unlimited ---------------------------------------------

test("company upload is unlimited", () => {
  assert.equal(checkBulkUploadRowLimit("company", 500).ok, true);
  assert.equal(checkBulkUploadRowLimit("company", 5000).ok, true);
});
