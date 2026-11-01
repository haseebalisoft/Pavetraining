/**
 * Overall "EUSR Expiry" Matrix column retired: EUSR is shown per course, and
 * a candidate with no course dates still sees their EUSR card on the profile.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-eusr-retired.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const { CLIENT_MATRIX_DISPLAY_HEADERS, CLIENT_MATRIX_META_HEADERS } = await import(
  "@/lib/services/bulkUpload/clientTemplateHeaders"
);
const { MATRIX_EDITOR_DATE_HEADERS, isMatrixExpiryColumnKey } = await import(
  "@/lib/training/matrixEditorFields"
);
const { buildCandidateCategoryRows } = await import("@/lib/training/candidateCategories");

test("Matrix table/editor no longer show the overall EUSR Expiry column", () => {
  assert.ok(!CLIENT_MATRIX_DISPLAY_HEADERS.includes("EUSR Expiry"));
  assert.ok(!MATRIX_EDITOR_DATE_HEADERS.includes("EUSR Expiry"));
  // Each EUSR course keeps its own column.
  for (const course of ["EUSR - SHEA Gas", "EUSR - Water Hygiene"]) {
    assert.ok(MATRIX_EDITOR_DATE_HEADERS.includes(course), course);
  }
});

test("upload template still accepts EUSR Expiry (it fills the Workforce card)", () => {
  assert.ok(CLIENT_MATRIX_META_HEADERS.includes("EUSR Expiry"));
});

test("a stored overall EUSR date no longer drives Next expiry / colours", () => {
  assert.equal(isMatrixExpiryColumnKey("EUSR Expiry"), false);
  assert.equal(isMatrixExpiryColumnKey("EUSR - SHEA Gas"), true);
});

const matrixRow = (columnValues = {}, extra = {}) => ({
  id: "1",
  candidateName: "Aiden Mercer",
  columnValues,
  eusrCategoryRows: [],
  ...extra,
});
const categories = (rows) => rows.map((row) => `${row.category}|${row.source}|${row.expiryDate ?? ""}`);

test("no EUSR course: profile shows the EUSR card date, not the retired column", () => {
  const rows = buildCandidateCategoryRows({
    matrixRow: matrixRow({ "EUSR Expiry": "2028-07-27", "CSCS Expiry": "2027-01-27" }),
    eusrCardExpiry: "2028-07-27T07:00:00Z",
  });
  const eusr = rows.filter((row) => /eusr/i.test(row.category));
  assert.deepEqual(categories(eusr), ["EUSR card|EUSR|2028-07-27T07:00:00Z"]);
});

test("EUSR register course present: shown per course, no card fallback", () => {
  const rows = buildCandidateCategoryRows({
    eusrRecords: [
      { id: "9", candidateName: "Aiden Mercer", eusrCategory: "SHEA Gas", trainingDate: "2025-07-01", expiry: "2028-07-01", outcome: "Pass" },
    ],
    matrixRow: matrixRow({ "EUSR Expiry": "2028-07-27" }),
    eusrCardExpiry: "2028-07-27",
  });
  const eusr = rows.filter((row) => /eusr|shea/i.test(row.category));
  assert.equal(eusr.length, 1);
  assert.equal(eusr[0].category, "SHEA Gas");
  assert.equal(eusr[0].expiryDate, "2028-07-01");
});

test("per-course Matrix column present: shown, no card fallback", () => {
  const rows = buildCandidateCategoryRows({
    matrixRow: matrixRow({ "EUSR - SHEA Power": "2029-01-01" }),
    eusrCardExpiry: "2028-07-27",
  });
  const eusr = rows.filter((row) => /eusr/i.test(row.category));
  assert.deepEqual(categories(eusr).length, 1);
  assert.match(eusr[0].category, /SHEA Power/);
  assert.ok(!eusr.some((row) => row.category === "EUSR card"));
});

test("no EUSR at all: nothing is invented", () => {
  const rows = buildCandidateCategoryRows({ matrixRow: matrixRow({}), eusrCardExpiry: null });
  assert.ok(!rows.some((row) => /eusr/i.test(row.category)));
});
