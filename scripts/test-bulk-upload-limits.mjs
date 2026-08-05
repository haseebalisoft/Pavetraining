/**
 * Unit tests for the bulk-upload hard row limits.
 *
 * Runs with the built-in Node test runner + native TypeScript type-stripping
 * (Node >= 22); bulkUploadLimits.ts has only a type-only import, so it loads
 * standalone with no Graph and no "@/" resolver hook:
 *
 *   node --test scripts/test-bulk-upload-limits.mjs
 *
 * Wayne's requirement: workforce bulk upload accepts at most 50 records per
 * upload; blank rows and the header do not count; the error is clear; and
 * company can carry its own separate limit (or none) without affecting this.
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

const WORKFORCE_ERROR =
  "Workforce bulk upload supports a maximum of 50 records per upload.";

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

test("workforce limit is 50; company has its own (no) limit", () => {
  assert.equal(bulkUploadRowLimit("workforce"), 50);
  assert.equal(bulkUploadRowLimit("company"), null);
});

test("error message matches the required client wording exactly", () => {
  assert.equal(bulkUploadRowLimitError("workforce", 50), WORKFORCE_ERROR);
});

// --- Accept / reject boundary --------------------------------------------

test("50 workforce rows are accepted", () => {
  const check = checkBulkUploadRowLimit("workforce", 50);
  assert.equal(check.ok, true);
  assert.equal(check.error, null);
  assert.equal(check.limit, 50);
});

test("51 workforce rows are rejected with the clear error", () => {
  const check = checkBulkUploadRowLimit("workforce", 51);
  assert.equal(check.ok, false);
  assert.equal(check.error, WORKFORCE_ERROR);
  assert.equal(check.limit, 50);
});

test("far-over (e.g. 1000) workforce rows are rejected", () => {
  const check = checkBulkUploadRowLimit("workforce", 1000);
  assert.equal(check.ok, false);
  assert.equal(check.error, WORKFORCE_ERROR);
});

// --- Empty rows do not count ---------------------------------------------

test("isEmptyFieldSet treats null/blank/whitespace-only field sets as empty", () => {
  assert.equal(isEmptyFieldSet(null), true);
  assert.equal(isEmptyFieldSet({}), true);
  assert.equal(isEmptyFieldSet({ a: null, b: "", c: "   " }), true);
  assert.equal(isEmptyFieldSet({ a: null, b: "x" }), false);
});

test("blank rows do not count toward the limit (50 populated + 20 blank = OK)", () => {
  const count = countPopulatedRows(rows(50, 20));
  assert.equal(count, 50);
  assert.equal(checkBulkUploadRowLimit("workforce", count).ok, true);
});

test("51 populated rows are rejected even with no blank rows", () => {
  const count = countPopulatedRows(rows(51, 0));
  assert.equal(count, 51);
  assert.equal(checkBulkUploadRowLimit("workforce", count).ok, false);
});

test("51 populated + blank padding still counts 51 and is rejected", () => {
  const count = countPopulatedRows(rows(51, 30));
  assert.equal(count, 51);
  const check = checkBulkUploadRowLimit("workforce", count);
  assert.equal(check.ok, false);
  assert.equal(check.error, WORKFORCE_ERROR);
});

// --- Company keeps its own behaviour (separate limit / no cap) ------------

test("company upload is not limited by the workforce cap", () => {
  assert.equal(checkBulkUploadRowLimit("company", 500).ok, true);
  assert.equal(checkBulkUploadRowLimit("company", 5000).ok, true);
});
