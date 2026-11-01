/**
 * Name + DOB matching guards (the client-approved rule): link the right
 * Workforce ↔ Matrix records, never duplicate, flag ambiguity instead of
 * guessing.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-name-dob-matching.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const { findMatrixRowByWorkforce, deriveMatrixLinkStatus } = await import(
  "@/lib/services/bulkUpload/workforceMatrixSync"
);

const row = (id, candidateName, dateOfBirth, extra = {}) => ({
  id,
  candidateName,
  dateOfBirth,
  workforceItemId: null,
  workforceNumber: null,
  companyItemId: null,
  companyNumber: null,
  ...extra,
});
const person = (id, candidateName, dateOfBirth, extra = {}) => ({
  id,
  workforceItemId: id,
  candidateName,
  dateOfBirth,
  workforceNumber: null,
  companyItemId: "424",
  companyNumber: "C00002",
  ...extra,
});

// Live pair at Derek Lewis Surfacing: two "Aiden Mercer"s, different DOBs.
const aidenA = person("1590", "Aiden Mercer", "1981-07-13");
const aidenB = person("1632", "Aiden Mercer", "1983-08-14");

test("same-name people each link to their own row by Name + DOB", () => {
  const rows = [row("1078", "Aiden Mercer", "1981-07-13"), row("1089", "Aiden Mercer", "1983-08-14")];
  const peers = { workforceNamePeers: 2, workforceNameDobPeers: 1 };
  assert.equal(findMatrixRowByWorkforce(rows, aidenA, peers).row?.id, "1078");
  assert.equal(findMatrixRowByWorkforce(rows, aidenB, peers).row?.id, "1089");
});

test("a row owned by one person is never taken by the same-named other", () => {
  const rows = [row("1078", "Aiden Mercer", "1981-07-13", { workforceItemId: "1590" })];
  const result = findMatrixRowByWorkforce(rows, aidenB, { workforceNamePeers: 2, workforceNameDobPeers: 1 });
  assert.equal(result.row, null);
  assert.equal(result.ambiguous, false); // caller creates B's own row — no duplicate of A's
});

test("two people with the same Name AND DOB: flagged, not guessed", () => {
  const twin = person("1700", "Aiden Mercer", "1981-07-13", { companyItemId: "425", companyNumber: "C00003" });
  const rows = [row("2000", "Aiden Mercer", "1981-07-13")];
  const result = findMatrixRowByWorkforce(rows, { ...twin, companyItemId: null, companyNumber: null }, {
    workforceNamePeers: 2,
    workforceNameDobPeers: 2,
  });
  assert.equal(result.row, null);
  assert.equal(result.ambiguous, true);
});

test("no-DOB row at the company is not claimed when someone else shares the name", () => {
  // Before: Company + Name accepted an unknown DOB, so whichever Aiden synced
  // first took the row.
  const rows = [row("2001", "Aiden Mercer", null, { companyItemId: "424" })];
  const result = findMatrixRowByWorkforce(rows, aidenA, { workforceNamePeers: 2, workforceNameDobPeers: 1 });
  assert.equal(result.row, null);
});

test("unchanged: unique name + company still links a no-DOB row", () => {
  const solo = person("1800", "Priya Shah", "1990-02-02");
  const rows = [row("2002", "Priya Shah", null, { companyItemId: "424" })];
  const result = findMatrixRowByWorkforce(rows, solo, { workforceNamePeers: 1, workforceNameDobPeers: 1 });
  assert.equal(result.row?.id, "2002");
  assert.equal(result.matchType, "companyNameDob");
});

test("unchanged: a stored WorkforceItemId wins even after a rename", () => {
  const renamed = person("1590", "Aiden J Mercer", "1981-07-13");
  const rows = [row("1078", "Aiden Mercer", "1981-07-13", { workforceItemId: "1590" })];
  assert.equal(findMatrixRowByWorkforce(rows, renamed, { workforceNamePeers: 1 }).row?.id, "1078");
});

test("unchanged: DOB conflict never links", () => {
  const rows = [row("2003", "Zaid", "1975-01-01")];
  const zaid = person("1500", "Zaid", "1972-01-01");
  assert.equal(findMatrixRowByWorkforce(rows, zaid, { workforceNamePeers: 1 }).row, null);
});

test("Admin status: a no-id row needs Name + DOB, not just a unique name", () => {
  const base = { hasWorkforceItemId: false, workforceResolved: false, nameMatchCount: 1 };
  assert.equal(deriveMatrixLinkStatus({ ...base, nameDobConfirmed: true }), "Linked");
  assert.equal(deriveMatrixLinkStatus({ ...base, nameDobConfirmed: false }), "Needs Review");
  // Callers that do not supply the DOB check keep the original behaviour.
  assert.equal(deriveMatrixLinkStatus(base), "Linked");
  assert.equal(deriveMatrixLinkStatus({ ...base, nameMatchCount: 2, nameDobConfirmed: true }), "Needs Review");
  assert.equal(deriveMatrixLinkStatus({ ...base, nameMatchCount: 0 }), "Orphan");
});
