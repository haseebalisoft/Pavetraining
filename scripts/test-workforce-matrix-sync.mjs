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
  findMatrixRowByWorkforce,
  deriveMatrixLinkStatus,
  buildWorkforceMatrixSource,
  candidateNameKey,
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

// --- Parallel matrix sync: grouping key + no-duplicate invariant ----------
// Bulk import now groups imported candidates by candidateNameKey and runs the
// groups in parallel. These lock the two properties that make that safe.

test("candidateNameKey trims, lowercases, and collapses inner whitespace", () => {
  assert.equal(candidateNameKey("  John   Doe "), "john doe");
  assert.equal(candidateNameKey("JOHN DOE"), "john doe");
  // Same key ⇒ same group ⇒ never synced in parallel against each other.
  assert.equal(
    candidateNameKey("Jane  Smith"),
    candidateNameKey("jane smith"),
  );
  // Different names ⇒ different keys ⇒ safe to run concurrently.
  assert.notEqual(candidateNameKey("John Doe"), candidateNameKey("Jane Doe"));
  assert.equal(candidateNameKey(null), "");
});

/**
 * Simulate the importer's per-group sequential sync using the REAL matching
 * function, to prove that grouping by candidateNameKey reproduces the old
 * sequential result: within a name group the first candidate creates the row
 * and the rest update it — never a second create (no duplicate row).
 */
function simulateGroupSync(groupProfiles, snapshotRows) {
  const groupKey = candidateNameKey(groupProfiles[0].candidateName);
  const localRows = snapshotRows.filter(
    (row) => candidateNameKey(row.candidateName) === groupKey,
  );
  const actions = [];
  let nextId = 900;
  for (const profile of groupProfiles) {
    const existing = findMatrixRowForCandidate(localRows, profile);
    if (existing) {
      actions.push({ action: "update", id: existing.id });
    } else {
      const created = {
        id: String(nextId++),
        candidateName: profile.candidateName.trim(),
        dateOfBirth: profile.dateOfBirth ?? null,
      };
      localRows.push(created); // mirrors `if (sync.created) localRows.push(sync.row)`
      actions.push({ action: "create", id: created.id });
    }
  }
  return actions;
}

test("same-name group with no existing row: first creates, rest update (no dupe)", () => {
  const actions = simulateGroupSync(
    [
      { candidateName: "New Guy", dateOfBirth: null },
      { candidateName: "new  guy", dateOfBirth: null },
      { candidateName: "NEW GUY", dateOfBirth: null },
    ],
    [], // empty snapshot — nobody exists yet
  );
  assert.deepEqual(
    actions.map((a) => a.action),
    ["create", "update", "update"],
  );
  // All three resolve to the SAME row id — exactly one row was created.
  assert.equal(actions[1].id, actions[0].id);
  assert.equal(actions[2].id, actions[0].id);
});

test("same-name group with an existing snapshot row: all update, none create", () => {
  const actions = simulateGroupSync(
    [
      { candidateName: "Jane Smith", dateOfBirth: "1990-01-01" },
      { candidateName: "jane smith", dateOfBirth: "1990-01-01" },
    ],
    ROWS,
  );
  assert.deepEqual(
    actions.map((a) => a.action),
    ["update", "update"],
  );
  assert.equal(actions[0].id, "100");
  assert.equal(actions[1].id, "100");
});

// --- Strong link fields on the source payload (Phase 4) -------------------
// buildWorkforceMatrixSource now also emits a typed linkFields bucket so the
// upsert can write Number (ids) and text/choice (numbers/name/status) columns.

test("linkFields emits numeric ids and text link columns + MatrixLinkStatus=Linked", () => {
  const { linkFields } = buildWorkforceMatrixSource({
    id: "12",
    candidateName: "Jane Smith",
    workforceNumber: "W00042",
    companyItemId: 7,
    companyNumber: "C00002",
  });
  // Numbers → Number columns.
  assert.equal(linkFields.numbers.WorkforceItemId, 12);
  assert.equal(linkFields.numbers.CompanyItemId, 7);
  // Text/choice → text columns; status is always Linked on a live sync.
  assert.equal(linkFields.text.WorkforceNumber, "W00042");
  assert.equal(linkFields.text.CompanyNumber, "C00002");
  assert.equal(linkFields.text.CandidateName, "Jane Smith");
  assert.equal(linkFields.text.MatrixLinkStatus, "Linked");
});

test("WorkforceItemId falls back to `id`; explicit workforceItemId wins", () => {
  const fallback = buildWorkforceMatrixSource({
    id: "99",
    candidateName: "A",
  }).linkFields;
  assert.equal(fallback.numbers.WorkforceItemId, 99);

  const explicit = buildWorkforceMatrixSource({
    id: "99",
    workforceItemId: 5,
    candidateName: "A",
  }).linkFields;
  assert.equal(explicit.numbers.WorkforceItemId, 5);
});

test("non-numeric / missing ids are omitted from linkFields.numbers (no bad writes)", () => {
  const { linkFields } = buildWorkforceMatrixSource({
    id: "not-a-number",
    candidateName: "A",
    // no companyItemId
  });
  assert.equal("WorkforceItemId" in linkFields.numbers, false);
  assert.equal("CompanyItemId" in linkFields.numbers, false);
  // Blank numbers/name still omit their text keys, but status is always set.
  assert.equal("WorkforceNumber" in linkFields.text, false);
  assert.equal("CompanyNumber" in linkFields.text, false);
  assert.equal(linkFields.text.MatrixLinkStatus, "Linked");
});

// --- findMatrixRowByWorkforce: id > legacy > unambiguous name --------------

const LINK_ROWS = [
  // Row linked by id to workforce 500.
  {
    id: "200",
    candidateName: "Linked Larry",
    workforceItemId: 500,
    workforceNumber: "W00500",
    companyItemId: 9,
  },
  // Legacy row: no id, but WorkforceNumber + CompanyItemId identify it.
  {
    id: "201",
    candidateName: "Legacy Lou",
    workforceNumber: "W00600",
    companyItemId: 3,
  },
  // A single UNLINKED row for a unique name.
  { id: "202", candidateName: "Fresh Fiona" },
  // Two rows share a name — ambiguous, must never be adopted by name.
  { id: "203", candidateName: "Ambi Twin" },
  { id: "204", candidateName: "Ambi Twin" },
  // A lone same-name row already linked to ANOTHER workforce (999).
  { id: "205", candidateName: "Taken Tom", workforceItemId: 999 },
];

test("matches by WorkforceItemId first (strongest key)", () => {
  const { row, matchType } = findMatrixRowByWorkforce(LINK_ROWS, {
    candidateName: "Different Name Now", // renamed — name must not matter
    workforceItemId: 500,
  });
  assert.equal(matchType, "id");
  assert.equal(row.id, "200");
});

test("falls back to WorkforceNumber + CompanyItemId for legacy rows", () => {
  const { row, matchType } = findMatrixRowByWorkforce(LINK_ROWS, {
    candidateName: "Legacy Lou",
    workforceNumber: "W00600",
    companyItemId: 3,
  });
  assert.equal(matchType, "legacy");
  assert.equal(row.id, "201");
});

test("adopts an unambiguous, still-unlinked same-name row", () => {
  const { row, matchType } = findMatrixRowByWorkforce(LINK_ROWS, {
    candidateName: "  fresh   fiona ",
    workforceItemId: 700,
  });
  assert.equal(matchType, "name");
  assert.equal(row.id, "202");
});

test("never adopts an ambiguous same-name row (caller creates a fresh row)", () => {
  const { row, matchType } = findMatrixRowByWorkforce(LINK_ROWS, {
    candidateName: "Ambi Twin",
    workforceItemId: 800,
  });
  assert.equal(matchType, "none");
  assert.equal(row, null);
});

test("never hijacks a lone same-name row already linked to another workforce", () => {
  const { row, matchType } = findMatrixRowByWorkforce(LINK_ROWS, {
    candidateName: "Taken Tom",
    workforceItemId: 801,
  });
  assert.equal(matchType, "none");
  assert.equal(row, null);
});

test("legacy fallback requires BOTH number and company id", () => {
  // WorkforceNumber matches but company id differs → no legacy match, and the
  // name is unique+unlinked so it adopts by name instead of mis-linking.
  const { matchType } = findMatrixRowByWorkforce(LINK_ROWS, {
    candidateName: "Legacy Lou",
    workforceNumber: "W00600",
    companyItemId: 999, // wrong company
  });
  assert.equal(matchType, "name");
});

// --- deriveMatrixLinkStatus (display badges + hide-orphans filter) ---------

test("deriveMatrixLinkStatus classifies every link state", () => {
  // Resolved id → Linked.
  assert.equal(
    deriveMatrixLinkStatus({
      hasWorkforceItemId: true,
      workforceResolved: true,
      nameMatchCount: 1,
    }),
    "Linked",
  );
  // Stored id no longer resolves, name still matches → Needs Review.
  assert.equal(
    deriveMatrixLinkStatus({
      hasWorkforceItemId: true,
      workforceResolved: false,
      nameMatchCount: 1,
    }),
    "Needs Review",
  );
  // Stored id gone and no name match → Orphan.
  assert.equal(
    deriveMatrixLinkStatus({
      hasWorkforceItemId: true,
      workforceResolved: false,
      nameMatchCount: 0,
    }),
    "Orphan",
  );
  // Legacy (no id): unique name → Linked (keeps pre-migration matrix working).
  assert.equal(
    deriveMatrixLinkStatus({
      hasWorkforceItemId: false,
      workforceResolved: false,
      nameMatchCount: 1,
    }),
    "Linked",
  );
  // Legacy: ambiguous same-name → Needs Review.
  assert.equal(
    deriveMatrixLinkStatus({
      hasWorkforceItemId: false,
      workforceResolved: false,
      nameMatchCount: 2,
    }),
    "Needs Review",
  );
  // Legacy: no workforce at all → Orphan (hidden by default in the admin UI).
  assert.equal(
    deriveMatrixLinkStatus({
      hasWorkforceItemId: false,
      workforceResolved: false,
      nameMatchCount: 0,
    }),
    "Orphan",
  );
});
