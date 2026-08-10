/**
 * Task G scenarios for two-way Workforce <-> Training Matrix linking, driven by
 * the REAL client spreadsheets (Workforce list.xlsx + Training matrix example.xlsx).
 *
 * Run: node --test scripts/test-matrix-link-scenarios.mjs
 *
 * The pure match core is exercised directly (no Graph, no server-only imports):
 * the simulator below mirrors exactly what syncWorkforceToTrainingMatrix and
 * matrixImporter do, so a divergence here is a real divergence in production.
 */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath, pathToFileURL } from "node:url";
import { dirname, join } from "node:path";
import test from "node:test";
import * as XLSX from "xlsx";

const here = dirname(fileURLToPath(import.meta.url));
const repo = join(here, "..");

const {
  buildUnlinkedMatrixSource,
  buildWorkforceMatrixSource,
  candidateNameKey,
  findMatrixRowByWorkforce,
  isoDateKey,
  mergeUploadedCells,
  AMBIGUOUS_MATRIX_MATCH_WARNING,
} = await import(
  pathToFileURL(
    join(repo, "src/lib/services/bulkUpload/workforceMatrixSync.ts"),
  ).href
);

// ---------------------------------------------------------------------------
// Spreadsheet fixtures
// ---------------------------------------------------------------------------

function readSheet(fileName) {
  const wb = XLSX.read(readFileSync(join(repo, fileName)), { cellDates: false });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
}

function excelSerialToIso(value) {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  const ms = Math.round((value - 25569) * 86400 * 1000);
  const date = new Date(ms);
  if (Number.isNaN(date.getTime())) return null;
  return date.toISOString().slice(0, 10);
}

function cellToIso(value) {
  if (value == null) return null;
  if (typeof value === "number") return excelSerialToIso(value);
  const text = String(value).trim();
  if (!text) return null;
  return isoDateKey(text) || null;
}

const workforceSheet = readSheet("Workforce list.xlsx");
const matrixSheet = readSheet("Training matrix example.xlsx");

function pick(row, names) {
  for (const name of names) {
    if (row[name] != null && String(row[name]).trim() !== "") return row[name];
  }
  return null;
}

/** Workforce records shaped like AdminWorkforceRecord (ids assigned like SharePoint). */
const WORKFORCE = workforceSheet
  .map((row, index) => ({
    id: String(index + 1),
    candidateName: String(
      pick(row, ["Candidate Name", "CandidateName", "Name"]) ?? "",
    ).trim(),
    dateOfBirth: cellToIso(pick(row, ["DOB", "Date of birth", "DateOfBirth"])),
    companyName: String(
      pick(row, ["Company", "Company Name", "CompanyName"]) ?? "",
    ).trim(),
    companyNumber: String(pick(row, ["Company Number", "CompanyNumber"]) ?? "")
      .trim() || null,
    workforceNumber:
      String(pick(row, ["Workforce Number", "WorkforceNumber"]) ?? "").trim() ||
      null,
    companyId: null,
    cscsExpiry: null,
    eusrExpiry: null,
    swqrExpiry: null,
  }))
  .filter((row) => row.candidateName);

/** Matrix spreadsheet rows: Name + DOB + N-code expiry columns. */
const MATRIX_ROWS = matrixSheet
  .map((row) => {
    const cells = {};
    for (const [header, value] of Object.entries(row)) {
      if (value == null) continue;
      const iso = typeof value === "number" ? excelSerialToIso(value) : null;
      cells[header] = iso ?? String(value);
    }
    return {
      candidateName: String(
        pick(row, ["Candidate Name", "CandidateName", "Name"]) ?? "",
      ).trim(),
      dateOfBirth: cellToIso(pick(row, ["DOB", "Date of birth", "DateOfBirth"])),
      cells,
    };
  })
  .filter((row) => row.candidateName);

// ---------------------------------------------------------------------------
// Simulator — mirrors syncWorkforceToTrainingMatrix + matrixImporter
// ---------------------------------------------------------------------------

function makeStore() {
  let nextId = 1;
  return {
    rows: [],
    /** Mirrors upsertTrainingMatrixExampleRow: writes only own-properties. */
    upsert({ existingRow, candidateName, source, linkFields }) {
      const target =
        existingRow ??
        (() => {
          const row = {
            id: String(nextId++),
            candidateName,
            dateOfBirth: null,
            columnValues: {},
            manualOverrides: [],
            workforceItemId: null,
            workforceNumber: null,
            companyItemId: null,
            companyNumber: null,
            matrixLinkStatus: null,
          };
          this.rows.push(row);
          return row;
        })();
      target.candidateName = candidateName;
      Object.assign(target.columnValues, source);
      if (source.DOB) target.dateOfBirth = source.DOB;
      const numbers = linkFields?.numbers ?? {};
      const text = linkFields?.text ?? {};
      if (numbers.WorkforceItemId != null) {
        target.workforceItemId = String(numbers.WorkforceItemId);
      }
      if (numbers.CompanyItemId != null) {
        target.companyItemId = String(numbers.CompanyItemId);
      }
      if (text.WorkforceNumber) target.workforceNumber = text.WorkforceNumber;
      if (text.CompanyNumber) target.companyNumber = text.CompanyNumber;
      if (text.MatrixLinkStatus) target.matrixLinkStatus = text.MatrixLinkStatus;
      return target;
    },
  };
}

function peersFor(workforce, candidate) {
  const key = candidateNameKey(candidate.candidateName);
  const dob = isoDateKey(candidate.dateOfBirth);
  return {
    workforceNamePeers: workforce.filter(
      (row) => candidateNameKey(row.candidateName) === key,
    ).length,
    workforceNameDobPeers: dob
      ? workforce.filter(
          (row) =>
            candidateNameKey(row.candidateName) === key &&
            isoDateKey(row.dateOfBirth) === dob,
        ).length
      : 1,
  };
}

/** Mirrors syncWorkforceToTrainingMatrix. */
function syncWorkforce(store, workforce, candidate, uploadedCells) {
  const profile = {
    ...candidate,
    workforceItemId: candidate.id,
    companyItemId: candidate.companyId,
  };
  const match = findMatrixRowByWorkforce(
    store.rows,
    profile,
    peersFor(workforce, candidate),
  );
  if (match.ambiguous) {
    return {
      skipped: true,
      created: false,
      row: null,
      matchType: "none",
      warnings: [AMBIGUOUS_MATRIX_MATCH_WARNING],
    };
  }
  const built = buildWorkforceMatrixSource(profile);
  const source = mergeUploadedCells(built.source, uploadedCells);
  const before = store.rows.length;
  const row = store.upsert({
    existingRow: match.row,
    candidateName: candidate.candidateName,
    source,
    linkFields: built.linkFields,
  });
  return {
    skipped: false,
    created: store.rows.length > before,
    row,
    matchType: match.row ? match.matchType : "none",
    warnings: [],
  };
}

/** Mirrors matrixImporter.findWorkforceForMatrix. */
function findWorkforceForMatrix(workforce, row) {
  const key = candidateNameKey(row.candidateName);
  if (!key) return { candidate: null, ambiguous: false };
  const dob = isoDateKey(row.dateOfBirth);
  if (!dob) return { candidate: null, ambiguous: false };
  const matches = workforce.filter(
    (wf) =>
      candidateNameKey(wf.candidateName) === key &&
      isoDateKey(wf.dateOfBirth) === dob,
  );
  if (!matches.length) return { candidate: null, ambiguous: false };
  if (row.companyName) {
    const scoped = matches.filter(
      (wf) =>
        candidateNameKey(wf.companyName) === candidateNameKey(row.companyName),
    );
    if (scoped.length === 1) return { candidate: scoped[0], ambiguous: false };
    if (scoped.length > 1) return { candidate: null, ambiguous: true };
  }
  if (matches.length === 1) return { candidate: matches[0], ambiguous: false };
  return { candidate: null, ambiguous: true };
}

/** Mirrors commitMatrixImport for one row. */
function importMatrixRow(store, workforce, row) {
  const match = findWorkforceForMatrix(workforce, row);
  if (!match.candidate) {
    const { source, linkFields } = buildUnlinkedMatrixSource({
      candidateName: row.candidateName,
      dateOfBirth: row.dateOfBirth,
      uploadedCells: row.cells,
    });
    const written = store.upsert({
      existingRow: null,
      candidateName: row.candidateName,
      source,
      linkFields,
    });
    return {
      linked: false,
      ambiguous: match.ambiguous,
      row: written,
    };
  }
  const sync = syncWorkforce(store, workforce, match.candidate, row.cells);
  return { linked: !sync.skipped, ambiguous: sync.skipped, row: sync.row };
}

// ---------------------------------------------------------------------------
// Fixture sanity — the assumptions every scenario below relies on
// ---------------------------------------------------------------------------

test("fixtures: both spreadsheets parsed with rows", () => {
  assert.ok(WORKFORCE.length >= 10, `workforce rows: ${WORKFORCE.length}`);
  assert.ok(MATRIX_ROWS.length >= 10, `matrix rows: ${MATRIX_ROWS.length}`);
  assert.ok(WORKFORCE.every((row) => row.dateOfBirth));
  assert.ok(MATRIX_ROWS.every((row) => row.dateOfBirth));
});

test("fixtures: real data contains same-name candidates with different DOBs", () => {
  const byName = new Map();
  for (const row of WORKFORCE) {
    const key = candidateNameKey(row.candidateName);
    byName.set(key, (byName.get(key) ?? 0) + 1);
  }
  const shared = [...byName.values()].filter((count) => count > 1);
  assert.ok(shared.length > 0, "expected same-name pairs in the real data");

  // No same name AND same DOB, so Name+DOB is a safe unique key here.
  const byNameDob = new Set();
  for (const row of WORKFORCE) {
    const key = `${candidateNameKey(row.candidateName)}|${isoDateKey(row.dateOfBirth)}`;
    assert.equal(byNameDob.has(key), false, `duplicate name+DOB: ${key}`);
    byNameDob.add(key);
  }
});

// ---------------------------------------------------------------------------
// Scenario 1 + 5: Matrix first (Needs Review) -> Workforce later links the SAME row
// ---------------------------------------------------------------------------

test("scenario 1+5: matrix uploaded first becomes Needs Review, then Workforce links the same row (no duplicate)", () => {
  const store = makeStore();

  // Matrix upload with an EMPTY workforce list.
  for (const row of MATRIX_ROWS) {
    const result = importMatrixRow(store, [], row);
    assert.equal(result.linked, false);
    assert.equal(result.row.matrixLinkStatus, "Needs Review");
    assert.equal(result.row.workforceItemId, null);
  }
  assert.equal(store.rows.length, MATRIX_ROWS.length);
  const afterMatrix = store.rows.length;

  // Now the Workforce import runs. Every row must be ADOPTED, not duplicated.
  let created = 0;
  let skipped = 0;
  for (const candidate of WORKFORCE) {
    const sync = syncWorkforce(store, WORKFORCE, candidate);
    if (sync.skipped) {
      skipped += 1;
      continue;
    }
    if (sync.created) created += 1;
  }
  assert.equal(skipped, 0, "no candidate should be ambiguous in the real data");

  // Only candidates with NO matrix row may create one.
  const matrixKeys = new Set(
    MATRIX_ROWS.map(
      (row) => `${candidateNameKey(row.candidateName)}|${isoDateKey(row.dateOfBirth)}`,
    ),
  );
  const expectedCreates = WORKFORCE.filter(
    (wf) =>
      !matrixKeys.has(
        `${candidateNameKey(wf.candidateName)}|${isoDateKey(wf.dateOfBirth)}`,
      ),
  ).length;
  assert.equal(created, expectedCreates);
  assert.equal(store.rows.length, afterMatrix + expectedCreates);

  // Every adopted row is now Linked and owned by exactly one workforce record.
  const owners = new Set();
  for (const row of store.rows) {
    if (!row.workforceItemId) continue;
    assert.equal(row.matrixLinkStatus, "Linked");
    assert.equal(owners.has(row.workforceItemId), false, "row owner reused");
    owners.add(row.workforceItemId);
  }
});

// ---------------------------------------------------------------------------
// Scenario 2 + 3: Workforce first, then Matrix
// ---------------------------------------------------------------------------

test("scenario 2+3: Workforce first creates linked rows, then Matrix updates them in place", () => {
  const store = makeStore();

  for (const candidate of WORKFORCE) {
    const sync = syncWorkforce(store, WORKFORCE, candidate);
    assert.equal(sync.skipped, false);
    assert.equal(sync.created, true, `expected a new row for ${candidate.candidateName}`);
    assert.equal(sync.row.matrixLinkStatus, "Linked");
    assert.equal(sync.row.workforceItemId, candidate.id);
  }
  assert.equal(store.rows.length, WORKFORCE.length);
  const afterWorkforce = store.rows.length;

  let linked = 0;
  for (const row of MATRIX_ROWS) {
    const result = importMatrixRow(store, WORKFORCE, row);
    assert.equal(result.ambiguous, false, `ambiguous: ${row.candidateName}`);
    assert.equal(result.linked, true, `unmatched: ${row.candidateName}`);
    linked += 1;
  }
  assert.equal(linked, MATRIX_ROWS.length);
  assert.equal(
    store.rows.length,
    afterWorkforce,
    "matrix import must not create any new row",
  );
});

test("scenario 3: Workforce with no matrix row gets exactly one new linked row", () => {
  const store = makeStore();
  const candidate = WORKFORCE[0];
  const first = syncWorkforce(store, WORKFORCE, candidate);
  assert.equal(first.created, true);
  const second = syncWorkforce(store, WORKFORCE, candidate);
  assert.equal(second.created, false);
  assert.equal(second.row.id, first.row.id);
  assert.equal(store.rows.length, 1);
});

// ---------------------------------------------------------------------------
// Scenario 4: Matrix row with no Workforce match
// ---------------------------------------------------------------------------

test("scenario 4: matrix row with no Workforce match imports as Needs Review with its expiry data", () => {
  const store = makeStore();
  const row = MATRIX_ROWS[0];
  const result = importMatrixRow(store, WORKFORCE.slice(1), row);

  // WORKFORCE[0] removed, so row 0 has no owner (real data has no same name+DOB twin).
  const stillMatches = WORKFORCE.slice(1).some(
    (wf) =>
      candidateNameKey(wf.candidateName) === candidateNameKey(row.candidateName) &&
      isoDateKey(wf.dateOfBirth) === isoDateKey(row.dateOfBirth),
  );
  assert.equal(stillMatches, false, "fixture assumption: no remaining twin");

  assert.equal(result.linked, false);
  assert.equal(result.row.matrixLinkStatus, "Needs Review");
  assert.equal(result.row.workforceItemId, null);
  assert.equal(result.row.candidateName, row.candidateName);
  assert.equal(result.row.dateOfBirth, row.dateOfBirth);

  // The uploaded training data survived.
  const uploadedExpiries = Object.entries(row.cells).filter(
    ([header, value]) => header !== "Name" && header !== "DOB" && value,
  );
  for (const [header, value] of uploadedExpiries) {
    assert.equal(result.row.columnValues[header], value, `lost ${header}`);
  }
});

// ---------------------------------------------------------------------------
// Scenario 6: same name, different DOB
// ---------------------------------------------------------------------------

test("scenario 6: same name with different DOBs never collide", () => {
  const byName = new Map();
  for (const row of WORKFORCE) {
    const key = candidateNameKey(row.candidateName);
    const bucket = byName.get(key) ?? [];
    bucket.push(row);
    byName.set(key, bucket);
  }
  const pairs = [...byName.values()].filter((bucket) => bucket.length > 1);
  assert.ok(pairs.length > 0);

  for (const pair of pairs) {
    const store = makeStore();
    // Matrix rows for BOTH twins land first as Needs Review.
    for (const twin of pair) {
      const matrixRow = MATRIX_ROWS.find(
        (row) =>
          candidateNameKey(row.candidateName) === candidateNameKey(twin.candidateName) &&
          isoDateKey(row.dateOfBirth) === isoDateKey(twin.dateOfBirth),
      );
      assert.ok(matrixRow, `no matrix row for ${twin.candidateName} ${twin.dateOfBirth}`);
      importMatrixRow(store, [], matrixRow);
    }
    assert.equal(store.rows.length, pair.length);

    // Each twin adopts its OWN row by DOB — no create, no cross-claim.
    for (const twin of pair) {
      const sync = syncWorkforce(store, WORKFORCE, twin);
      assert.equal(sync.skipped, false, `${twin.candidateName} went ambiguous`);
      assert.equal(sync.created, false, `${twin.candidateName} duplicated a row`);
      assert.equal(isoDateKey(sync.row.dateOfBirth), isoDateKey(twin.dateOfBirth));
    }
    assert.equal(store.rows.length, pair.length);
    const owners = new Set(store.rows.map((row) => row.workforceItemId));
    assert.equal(owners.size, pair.length, "twins share a matrix row");
  }
});

// ---------------------------------------------------------------------------
// Scenario 7: same Name + same DOB in two companies
// ---------------------------------------------------------------------------

test("scenario 7: same Name + DOB in two companies does not auto-link without Company", () => {
  const store = makeStore();
  const base = WORKFORCE[0];
  const twinA = { ...base, id: "9001", companyName: "Alpha Ltd", companyNumber: "COMP-A", companyId: "501" };
  const twinB = { ...base, id: "9002", companyName: "Beta Ltd", companyNumber: "COMP-B", companyId: "502" };
  const workforce = [twinA, twinB];

  // Two unlinked matrix rows with the same Name + DOB and no Company.
  for (let i = 0; i < 2; i += 1) {
    const { source, linkFields } = buildUnlinkedMatrixSource({
      candidateName: base.candidateName,
      dateOfBirth: base.dateOfBirth,
    });
    store.upsert({
      existingRow: null,
      candidateName: base.candidateName,
      source,
      linkFields,
    });
  }

  const syncA = syncWorkforce(store, workforce, twinA);
  assert.equal(syncA.skipped, true, "must not guess between two identical rows");
  assert.deepEqual(syncA.warnings, [AMBIGUOUS_MATRIX_MATCH_WARNING]);
  assert.equal(store.rows.length, 2, "no row written while ambiguous");

  // With Company on the matrix row the tie resolves.
  store.rows[0].companyNumber = "COMP-A";
  const resolved = syncWorkforce(store, workforce, twinA);
  assert.equal(resolved.skipped, false);
  assert.equal(resolved.created, false);
  assert.equal(resolved.row.id, "1");
  assert.equal(resolved.matchType, "companyNameDob");

  // Twin B: row 1 is owned by twin A, and row 2 carries NO Company — so Name+DOB
  // alone would be a guess between two real people. Still refuses to link.
  const syncB = syncWorkforce(store, workforce, twinB);
  assert.equal(syncB.skipped, true, "Company is required to link the second twin");
  assert.deepEqual(syncB.warnings, [AMBIGUOUS_MATRIX_MATCH_WARNING]);
  assert.equal(store.rows.length, 2, "no row written while ambiguous");

  // Adding Company to the remaining row resolves it to twin B.
  store.rows[1].companyNumber = "COMP-B";
  const resolvedB = syncWorkforce(store, workforce, twinB);
  assert.equal(resolvedB.skipped, false);
  assert.equal(resolvedB.created, false);
  assert.equal(resolvedB.row.id, "2");
  assert.equal(resolvedB.matchType, "companyNameDob");
  assert.equal(store.rows.length, 2);
  assert.equal(store.rows[0].workforceItemId, "9001");
  assert.equal(store.rows[1].workforceItemId, "9002");
});

test("scenario 7b: matrix importer leaves same Name + DOB across companies unlinked", () => {
  const base = WORKFORCE[0];
  const workforce = [
    { ...base, id: "9001", companyName: "Alpha Ltd" },
    { ...base, id: "9002", companyName: "Beta Ltd" },
  ];
  const store = makeStore();
  const result = importMatrixRow(store, workforce, {
    candidateName: base.candidateName,
    dateOfBirth: base.dateOfBirth,
    cells: { "N001 - Ind FLT": "2030-01-01" },
  });
  assert.equal(result.linked, false);
  assert.equal(result.ambiguous, true);
  assert.equal(result.row.matrixLinkStatus, "Needs Review");
  assert.equal(result.row.columnValues["N001 - Ind FLT"], "2030-01-01");
});

// ---------------------------------------------------------------------------
// Scenario 8: blank cells never erase live values
// ---------------------------------------------------------------------------

test("scenario 8: blank matrix cells do not erase existing expiry values", () => {
  const store = makeStore();
  const candidate = WORKFORCE[0];

  // Seed a real expiry via a matrix upload.
  syncWorkforce(store, WORKFORCE, candidate, {
    "N001 - Ind FLT": "2030-06-01",
    "N003 - Reach Lift Truck": "2029-01-15",
  });
  const rowId = store.rows[0].id;
  assert.equal(store.rows[0].columnValues["N001 - Ind FLT"], "2030-06-01");

  // Re-upload with blanks in every shape the client's sheets produce.
  const sync = syncWorkforce(store, WORKFORCE, candidate, {
    "N001 - Ind FLT": "",
    "N003 - Reach Lift Truck": null,
    "N004 - Lorry Mounted Lift Truck": "   ",
    "N010 - Telescopic Handler": "—",
    "N020 - Tiltrotator System": "N/A",
    "N021 - Suction Excavator": "2031-02-02",
  });

  assert.equal(sync.created, false);
  assert.equal(sync.row.id, rowId);
  assert.equal(store.rows[0].columnValues["N001 - Ind FLT"], "2030-06-01");
  assert.equal(store.rows[0].columnValues["N003 - Reach Lift Truck"], "2029-01-15");
  assert.equal(
    "N004 - Lorry Mounted Lift Truck" in store.rows[0].columnValues,
    false,
  );
  assert.equal("N010 - Telescopic Handler" in store.rows[0].columnValues, false);
  assert.equal("N020 - Tiltrotator System" in store.rows[0].columnValues, false);
  // A non-blank value still writes through.
  assert.equal(store.rows[0].columnValues["N021 - Suction Excavator"], "2031-02-02");
});

test("scenario 8b: full matrix re-upload of the real sheet preserves every value", () => {
  const store = makeStore();
  for (const candidate of WORKFORCE) syncWorkforce(store, WORKFORCE, candidate);
  for (const row of MATRIX_ROWS) importMatrixRow(store, WORKFORCE, row);
  const snapshot = store.rows.map((row) => ({ ...row.columnValues }));

  // Second identical upload: same values, still no new rows.
  const before = store.rows.length;
  for (const row of MATRIX_ROWS) importMatrixRow(store, WORKFORCE, row);
  assert.equal(store.rows.length, before);
  store.rows.forEach((row, index) => {
    assert.deepEqual(row.columnValues, snapshot[index]);
  });

  // Third upload with every expiry cell blanked: nothing is erased.
  for (const row of MATRIX_ROWS) {
    const blanked = { ...row, cells: { Name: row.candidateName, DOB: row.dateOfBirth } };
    for (const header of Object.keys(row.cells)) {
      if (header !== "Name" && header !== "DOB") blanked.cells[header] = "";
    }
    importMatrixRow(store, WORKFORCE, blanked);
  }
  assert.equal(store.rows.length, before);
  store.rows.forEach((row, index) => {
    assert.deepEqual(row.columnValues, snapshot[index]);
  });
});

// ---------------------------------------------------------------------------
// Scenario 9: customer portal hides Needs Review rows
// ---------------------------------------------------------------------------

/** Mirrors the customer join in customerDashboardService.getCustomerMatrixRecords. */
function customerVisibleRows(rows, allowedWorkforce) {
  const allowedIds = new Set(allowedWorkforce.map((row) => String(row.id)));
  const byId = new Map();
  const legacyByName = new Map();
  for (const row of rows) {
    const status = row.matrixLinkStatus?.trim();
    if (status === "Orphan" || status === "Needs Review") continue;
    const owner = String(row.workforceItemId ?? "").trim();
    if (owner) {
      if (!allowedIds.has(owner) || byId.has(owner)) continue;
      byId.set(owner, row);
      continue;
    }
    const key = candidateNameKey(row.candidateName);
    if (!key) continue;
    legacyByName.set(key, legacyByName.has(key) ? null : row);
  }
  const nameCounts = new Map();
  for (const wf of allowedWorkforce) {
    const key = candidateNameKey(wf.candidateName);
    nameCounts.set(key, (nameCounts.get(key) ?? 0) + 1);
  }
  return allowedWorkforce.map((wf) => {
    const key = candidateNameKey(wf.candidateName);
    return {
      candidate: wf,
      matrix:
        byId.get(String(wf.id)) ??
        (nameCounts.get(key) === 1 ? (legacyByName.get(key) ?? null) : null),
    };
  });
}

test("scenario 9: customer portal hides Needs Review rows and shows them once linked", () => {
  const store = makeStore();
  // Matrix-first upload with no workforce: every row is Needs Review.
  for (const row of MATRIX_ROWS) importMatrixRow(store, [], row);

  const hidden = customerVisibleRows(store.rows, WORKFORCE);
  assert.equal(
    hidden.filter((entry) => entry.matrix).length,
    0,
    "Needs Review rows must not reach the customer portal",
  );

  // Workforce import links them.
  for (const candidate of WORKFORCE) syncWorkforce(store, WORKFORCE, candidate);
  const visible = customerVisibleRows(store.rows, WORKFORCE);
  assert.equal(visible.length, WORKFORCE.length);
  assert.equal(
    visible.filter((entry) => entry.matrix).length,
    WORKFORCE.length,
    "every linked candidate should now see its matrix row",
  );
  for (const entry of visible) {
    assert.equal(entry.matrix.workforceItemId, entry.candidate.id);
  }
});

test("scenario 9b: same-named candidates in different companies never see each other's rows", () => {
  const store = makeStore();
  const base = WORKFORCE[0];
  const twinA = { ...base, id: "9001", companyName: "Alpha Ltd", companyNumber: "COMP-A", companyId: "501" };
  const twinB = { ...base, id: "9002", companyName: "Beta Ltd", companyNumber: "COMP-B", companyId: "502" };

  syncWorkforce(store, [twinA], twinA, { "N001 - Ind FLT": "2030-01-01" });
  syncWorkforce(store, [twinB], twinB, { "N001 - Ind FLT": "2040-01-01" });
  assert.equal(store.rows.length, 2, "each company gets its own row");

  const alphaView = customerVisibleRows(store.rows, [twinA]);
  assert.equal(alphaView.length, 1);
  assert.equal(alphaView[0].matrix.workforceItemId, "9001");
  assert.equal(alphaView[0].matrix.columnValues["N001 - Ind FLT"], "2030-01-01");

  const betaView = customerVisibleRows(store.rows, [twinB]);
  assert.equal(betaView[0].matrix.workforceItemId, "9002");
  assert.equal(betaView[0].matrix.columnValues["N001 - Ind FLT"], "2040-01-01");
});

// ---------------------------------------------------------------------------
// Task D: duplicate prevention under repeated / interleaved uploads
// ---------------------------------------------------------------------------

test("task D: matrix then workforce then matrix again creates no duplicates", () => {
  const store = makeStore();
  for (const row of MATRIX_ROWS) importMatrixRow(store, [], row);
  const afterFirst = store.rows.length;

  for (const candidate of WORKFORCE) syncWorkforce(store, WORKFORCE, candidate);
  for (const row of MATRIX_ROWS) importMatrixRow(store, WORKFORCE, row);
  for (const candidate of WORKFORCE) syncWorkforce(store, WORKFORCE, candidate);
  for (const row of MATRIX_ROWS) importMatrixRow(store, WORKFORCE, row);

  const matrixKeys = new Set(
    MATRIX_ROWS.map(
      (row) => `${candidateNameKey(row.candidateName)}|${isoDateKey(row.dateOfBirth)}`,
    ),
  );
  const workforceOnly = WORKFORCE.filter(
    (wf) =>
      !matrixKeys.has(
        `${candidateNameKey(wf.candidateName)}|${isoDateKey(wf.dateOfBirth)}`,
      ),
  ).length;
  assert.equal(store.rows.length, afterFirst + workforceOnly);

  // One row per workforce owner, and every name+DOB appears exactly once.
  const owners = store.rows
    .map((row) => row.workforceItemId)
    .filter(Boolean);
  assert.equal(new Set(owners).size, owners.length);
  const keys = store.rows.map(
    (row) => `${candidateNameKey(row.candidateName)}|${isoDateKey(row.dateOfBirth)}`,
  );
  assert.equal(new Set(keys).size, keys.length, "duplicate name+DOB rows exist");
});

test("task D: name-only matrix row (no DOB) never auto-links", () => {
  const store = makeStore();
  const candidate = WORKFORCE[0];
  const result = importMatrixRow(store, WORKFORCE, {
    candidateName: candidate.candidateName,
    dateOfBirth: null,
    cells: { "N001 - Ind FLT": "2030-01-01" },
  });
  assert.equal(result.linked, false);
  assert.equal(result.row.matrixLinkStatus, "Needs Review");
  assert.equal(result.row.workforceItemId, null);
});
