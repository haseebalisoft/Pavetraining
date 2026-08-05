/**
 * Unit tests for the pure Workforce → Training Matrix sync rules.
 *
 * Runs with the built-in Node test runner + native TypeScript type-stripping
 * (Node >= 22); workforceMatrixSync.ts is dependency-free so it loads with no
 * Graph and no "@/" resolver hook:
 *
 *   node --test scripts/test-workforce-matrix-sync.mjs
 *
 * Confirmed issue being fixed: bulk workforce import used to skip the matrix
 * seed, leaving synthetic `workforce-only:<id>` rows that later matrix
 * spreadsheet imports tried to PATCH as real ids (and failed). These tests
 * cover the logic that makes the sync create/reuse REAL rows safely.
 */
import test from "node:test";
import assert from "node:assert/strict";

const BASE = new URL("../src/lib/", import.meta.url).pathname;
const {
  isSyntheticMatrixId,
  realMatrixItemId,
  findMatrixRowForCandidate,
  buildWorkforceMatrixSource,
} = await import(BASE + "services/bulkUpload/workforceMatrixSync.ts");

// --- Fake-id guard (requirement 2) ---------------------------------------

test("synthetic workforce-only ids are recognised and never used as update ids", () => {
  assert.equal(isSyntheticMatrixId("workforce-only:42"), true);
  assert.equal(isSyntheticMatrixId("example:42"), false);
  assert.equal(isSyntheticMatrixId("42"), false);

  // realMatrixItemId → null for blanks and synthetic ids (so caller creates a
  // real row instead of PATCHing a fake id), strips the example: prefix.
  assert.equal(realMatrixItemId("workforce-only:42"), null);
  assert.equal(realMatrixItemId(""), null);
  assert.equal(realMatrixItemId(null), null);
  assert.equal(realMatrixItemId("example:42"), "42");
  assert.equal(realMatrixItemId("42"), "42");
});

// --- Reuse existing row, never duplicate (requirements 3 & 4) -------------

const ROWS = [
  { id: "100", candidateName: "Jane Smith", dateOfBirth: "1990-01-01" },
  { id: "101", candidateName: "John Doe", dateOfBirth: "1985-05-05" },
  // Two rows share a name — only DOB can disambiguate.
  { id: "102", candidateName: "Sam Twin", dateOfBirth: "1970-02-02" },
  { id: "103", candidateName: "Sam Twin", dateOfBirth: "1999-09-09" },
];

test("existing row is reused (matched by name) — not duplicated", () => {
  const hit = findMatrixRowForCandidate(ROWS, {
    candidateName: "  jane   smith ",
    dateOfBirth: "1990-01-01",
  });
  assert.ok(hit);
  assert.equal(hit.id, "100");
});

test("no match returns null so the caller creates a new real row", () => {
  const hit = findMatrixRowForCandidate(ROWS, {
    candidateName: "Brand New Person",
    dateOfBirth: null,
  });
  assert.equal(hit, null);
});

test("duplicate names are disambiguated by DOB", () => {
  const younger = findMatrixRowForCandidate(ROWS, {
    candidateName: "Sam Twin",
    dateOfBirth: "1999-09-09",
  });
  assert.equal(younger.id, "103");
  const older = findMatrixRowForCandidate(ROWS, {
    candidateName: "Sam Twin",
    dateOfBirth: "1970-02-02",
  });
  assert.equal(older.id, "102");
});

test("ambiguous duplicate names with no DOB reuse the first (stable, no dupe)", () => {
  const hit = findMatrixRowForCandidate(ROWS, {
    candidateName: "Sam Twin",
    dateOfBirth: null,
  });
  assert.ok(hit);
  assert.equal(hit.id, "102");
});

// --- Never overwrite expiry with blanks (requirement 5) ------------------

test("blank expiries are omitted from source so existing columns are untouched", () => {
  const { source } = buildWorkforceMatrixSource({
    candidateName: "Jane Smith",
    dateOfBirth: null,
    cscsExpiry: null,
    eusrExpiry: "",
    swqrExpiry: "   ",
  });
  // Only Name is present — no DOB/expiry keys, so the upsert leaves those
  // columns (and any existing expiry data) untouched.
  assert.deepEqual(Object.keys(source), ["Name"]);
  assert.equal(source.Name, "Jane Smith");
  assert.equal("CSCS Expiry" in source, false);
  assert.equal("EUSR Expiry" in source, false);
  assert.equal("NRSWA Expiry" in source, false);
  assert.equal("DOB" in source, false);
});

test("non-blank expiries and DOB are included in source", () => {
  const { source } = buildWorkforceMatrixSource({
    candidateName: "Jane Smith",
    dateOfBirth: "1990-01-01",
    cscsExpiry: "2027-03-01",
    eusrExpiry: "2028-06-30",
    swqrExpiry: "2026-12-31",
  });
  assert.equal(source.DOB, "1990-01-01");
  assert.equal(source["CSCS Expiry"], "2027-03-01");
  assert.equal(source["EUSR Expiry"], "2028-06-30");
  assert.equal(source["NRSWA Expiry"], "2026-12-31");
});

// --- Profile detail fields (requirement 6) -------------------------------

test("non-blank profile fields are offered for write; blanks are omitted", () => {
  const { profileFields } = buildWorkforceMatrixSource({
    candidateName: "Jane Smith",
    companyName: "Acme Ltd",
    workforceNumber: "W00042",
    department: "Rail",
    trainingManager: "Pat Lead",
    supervisor: "  ",
    nporsNumbers: "NP-1",
    eusrNumber: null,
    swqrNumber: "SW-9",
    cscsNumber: "CS-7",
  });
  assert.equal(profileFields["Company"], "Acme Ltd");
  assert.equal(profileFields["Workforce Number"], "W00042");
  assert.equal(profileFields["Department"], "Rail");
  assert.equal(profileFields["Training Manager"], "Pat Lead");
  assert.equal(profileFields["NPORS Number"], "NP-1");
  assert.equal(profileFields["SWQR Number"], "SW-9");
  assert.equal(profileFields["CSCS Number"], "CS-7");
  // Blank supervisor and null EUSR number are not offered (no blank writes).
  assert.equal("Supervisor" in profileFields, false);
  assert.equal("EUSR Number" in profileFields, false);
});

// --- Name is always present (create needs a Title) -----------------------

test("source always carries the candidate name (Title on create)", () => {
  const { source } = buildWorkforceMatrixSource({ candidateName: "  Solo Name " });
  assert.equal(source.Name, "Solo Name");
});
