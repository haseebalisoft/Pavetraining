/**
 * Unit tests for the Company Number integrity rules used by company bulk import.
 *
 * Runs with the built-in Node test runner + native TypeScript type-stripping
 * (Node >= 22), so no build step or Graph access is needed:
 *
 *   node --test scripts/test-company-number-rules.mjs
 *
 * They cover the client rule: a company's number must never be lost,
 * overwritten, or duplicated on bulk import. They exercise the REAL rule module
 * (resolveCompanyImport / resolveCompanyNumberOnUpdate) plus the real matching
 * and allocation helpers, wired the same way the importer wires them, without
 * mocking Graph.
 */
import test from "node:test";
import assert from "node:assert/strict";

const BASE = new URL("../src/lib/", import.meta.url).pathname;
const {
  resolveCompanyNumberOnUpdate,
  resolveCompanyImport,
  COMPANY_NUMBER_CHANGE_ERROR,
  companyNumberDuplicateInFileError,
  companyNumberTakenError,
} = await import(BASE + "services/bulkUpload/companyNumberRules.ts");
const { findCompanyByName, nameKey } = await import(
  BASE + "services/bulkUpload/matching.ts"
);
const { allocateNextCompanyNumber } = await import(BASE + "companyNumber.ts");

/** Mirror of companyImporter.matchCompanyByNumber. */
function matchCompanyByNumber(companies, number) {
  const key = nameKey(number);
  if (!key) return null;
  return companies.find((c) => nameKey(c.companyNumber) === key) ?? null;
}

/**
 * Mirror of companyDocumentsFolderName in
 * src/lib/services/customerDocumentsFolderService.ts (that file imports
 * "server-only" and cannot load here). Requirement 4: folder name is
 * `${CompanyNumber} - ${CompanyName}` and must stay stable after import.
 */
function companyFolderName(companyNumber, companyName) {
  const number = (companyNumber ?? "").trim();
  return number ? `${number} - ${companyName}` : companyName;
}

/**
 * Runs one uploaded row through the SAME pipeline commitCompanyImport uses:
 * real matching -> real resolveCompanyImport -> real allocator, plus the
 * in-file duplicate-number guard. Returns what the importer would do/write.
 */
function decide({ companies, row, duplicateMode, seenNumbers, allocatedInBatch = [] }) {
  const companyName = row.companyName?.trim() ?? "";
  const incoming = row.companyNumber?.trim() ?? "";
  if (!companyName) return { action: "error", message: "Company Name is required." };

  const matchByNumber = matchCompanyByNumber(companies, incoming);
  const matchByName = findCompanyByName(companies, companyName);
  const target = matchByName ?? matchByNumber;
  const toMatch = (r) => (r ? { id: r.id, companyNumber: r.companyNumber } : null);

  const d = resolveCompanyImport({
    incoming,
    matchByNumber: toMatch(matchByNumber),
    matchByName: toMatch(matchByName),
    duplicateMode,
  });

  if (d.action === "reject") return { action: "reject", targetId: target?.id ?? null, message: d.message };
  if (d.action === "skip") return { action: "skip", targetId: d.targetId };

  let finalNumber;
  if (d.action === "update") {
    finalNumber =
      d.companyNumber ??
      allocateNextCompanyNumber(companies, [...(seenNumbers ?? []), ...allocatedInBatch]);
  } else {
    finalNumber =
      d.companyNumber ??
      allocateNextCompanyNumber(companies, [...(seenNumbers ?? []), ...allocatedInBatch]);
  }

  if (seenNumbers?.has(nameKey(finalNumber))) {
    return { action: "reject", message: companyNumberDuplicateInFileError(finalNumber) };
  }
  seenNumbers?.add(nameKey(finalNumber));
  if (d.action === "create") allocatedInBatch.push(finalNumber);

  return {
    action: d.action, // "update" | "create"
    targetId: d.action === "update" ? d.targetId : null,
    companyNumber: finalNumber,
  };
}

function fresh() {
  return [
    { id: "1", companyName: "ABC Ltd", companyNumber: "C00007" },
    { id: "2", companyName: "XYZ Civils", companyNumber: "C00041" },
  ];
}

// ---- resolveCompanyNumberOnUpdate: the core anti-overwrite rule ----

test("Req1: blank incoming number preserves the existing company number", () => {
  const r = resolveCompanyNumberOnUpdate({ incoming: "", existing: "C00007" });
  assert.equal(r.kind, "preserve");
  assert.equal(r.companyNumber, "C00007");
});

test("Req1: identical incoming number matches, keeps existing", () => {
  const r = resolveCompanyNumberOnUpdate({ incoming: " c00007 ", existing: "C00007" });
  assert.equal(r.kind, "match");
  assert.equal(r.companyNumber, "C00007");
});

test("Req1: different incoming number is rejected with the required message", () => {
  const r = resolveCompanyNumberOnUpdate({ incoming: "C00042", existing: "C00007" });
  assert.equal(r.kind, "reject");
  assert.equal(r.message, COMPANY_NUMBER_CHANGE_ERROR);
  assert.equal(
    COMPANY_NUMBER_CHANGE_ERROR,
    "Company Number cannot be changed for an existing company.",
  );
});

// ---- Test 1: update with blank number preserves the old number ----

test("update existing company with blank Company Number preserves old number (the reported bug)", () => {
  const seenNumbers = new Set();
  const d = decide({
    companies: fresh(),
    row: { companyName: "ABC Ltd", companyNumber: "" },
    duplicateMode: "update",
    seenNumbers,
  });
  assert.equal(d.action, "update");
  assert.equal(d.targetId, "1");
  assert.equal(d.companyNumber, "C00007"); // NOT freshly allocated
});

// ---- Test 2: update with a different number rejects the row ----

test("update existing company with a different Company Number rejects the row", () => {
  const d = decide({
    companies: fresh(),
    row: { companyName: "ABC Ltd", companyNumber: "C00099" },
    duplicateMode: "update",
    seenNumbers: new Set(),
  });
  assert.equal(d.action, "reject");
  assert.equal(d.targetId, "1");
  assert.equal(d.message, COMPANY_NUMBER_CHANGE_ERROR);
});

// ---- Test 3: create new company with blank number generates only if allowed ----

test("create new company with blank Company Number generates the next free C-number", () => {
  const d = decide({
    companies: fresh(),
    row: { companyName: "Brand New Co", companyNumber: "" },
    duplicateMode: "update",
    seenNumbers: new Set(),
  });
  assert.equal(d.action, "create");
  assert.equal(d.companyNumber, "C00042"); // max existing C00041 + 1
});

test("create new company with a supplied unique number keeps that number", () => {
  const d = decide({
    companies: fresh(),
    row: { companyName: "Brand New Co", companyNumber: "C09000" },
    duplicateMode: "update",
    seenNumbers: new Set(),
  });
  assert.equal(d.action, "create");
  assert.equal(d.companyNumber, "C09000");
});

// ---- Test 4: duplicate Company Number inside the uploaded file is rejected ----

test("duplicate Company Number within the same file is rejected on the second row", () => {
  const companies = fresh();
  const seenNumbers = new Set();
  const allocatedInBatch = [];
  const first = decide({
    companies,
    row: { companyName: "New A", companyNumber: "C05000" },
    duplicateMode: "update",
    seenNumbers,
    allocatedInBatch,
  });
  const second = decide({
    companies, // not yet persisted between rows in this simulation
    row: { companyName: "New B", companyNumber: "C05000" },
    duplicateMode: "update",
    seenNumbers,
    allocatedInBatch,
  });
  assert.equal(first.action, "create");
  assert.equal(first.companyNumber, "C05000");
  assert.equal(second.action, "reject");
  assert.equal(second.message, companyNumberDuplicateInFileError("C05000"));
});

test("two blank-number new rows get DISTINCT auto-allocated numbers (no in-file dup)", () => {
  const companies = fresh();
  const seenNumbers = new Set();
  const allocatedInBatch = [];
  const a = decide({ companies, row: { companyName: "New A", companyNumber: "" }, duplicateMode: "update", seenNumbers, allocatedInBatch });
  const b = decide({ companies, row: { companyName: "New B", companyNumber: "" }, duplicateMode: "update", seenNumbers, allocatedInBatch });
  assert.equal(a.companyNumber, "C00042");
  assert.equal(b.companyNumber, "C00043");
  assert.notEqual(a.companyNumber, b.companyNumber);
});

// ---- Test 5: duplicate Company Number already existing in SharePoint is rejected ----

test("supplying an existing SharePoint Company Number under a NEW company name is rejected", () => {
  const d = decide({
    companies: fresh(),
    row: { companyName: "Totally Different Co", companyNumber: "C00007" }, // C00007 belongs to ABC Ltd
    duplicateMode: "create",
    seenNumbers: new Set(),
  });
  assert.equal(d.action, "reject");
  assert.equal(d.message, companyNumberTakenError("C00007", "1"));
});

test("create-despite-duplicate never reuses the matched company's number", () => {
  const d = decide({
    companies: fresh(),
    row: { companyName: "ABC Ltd", companyNumber: "" },
    duplicateMode: "create",
    seenNumbers: new Set(),
  });
  assert.equal(d.action, "create");
  assert.notEqual(d.companyNumber, "C00007");
  assert.equal(d.companyNumber, "C00042");
});

// ---- Test 6: folder name expectation remains stable ----

test("folder name stays stable when a blank-number update preserves the old number", () => {
  const companies = fresh();
  const before = companyFolderName("C00007", "ABC Ltd");
  const d = decide({
    companies,
    row: { companyName: "ABC Ltd", companyNumber: "" },
    duplicateMode: "update",
    seenNumbers: new Set(),
  });
  const after = companyFolderName(d.companyNumber, "ABC Ltd");
  assert.equal(before, "C00007 - ABC Ltd");
  assert.equal(after, before); // folder path unchanged -> documents not orphaned
});

test("folder name for a new company uses the generated number", () => {
  const d = decide({
    companies: fresh(),
    row: { companyName: "Brand New Co", companyNumber: "" },
    duplicateMode: "update",
    seenNumbers: new Set(),
  });
  assert.equal(companyFolderName(d.companyNumber, "Brand New Co"), "C00042 - Brand New Co");
});
