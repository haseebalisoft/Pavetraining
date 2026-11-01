/**
 * Training Matrix "Needs Review" reasons: every unmatched or ambiguous row must
 * say exactly why (DOB mismatch, candidate not found, multiple matches,
 * company mismatch, …), in the upload and in Admin Matrix.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-matrix-review-reasons.mjs
 *
 * (Node >= 22 for TypeScript type-stripping; the stub resolves "@/".)
 */
import test from "node:test";
import assert from "node:assert/strict";

const { matchWorkforceForMatrixRow, explainMatrixLinkStatus } = await import(
  "@/lib/services/bulkUpload/matrixReviewReason"
);
const { bulkCommitOutcome, reviewReasonSummary } = await import(
  "@/lib/jobs/jobOutcome"
);

const workforce = [
  { id: "1", candidateName: "Zaid Khan", dateOfBirth: "1980-02-02", companyName: "Acme Ltd", companyNumber: "C1", workforceNumber: "W1" },
  { id: "2", candidateName: "John Smith", dateOfBirth: "05/05/1990", companyName: "Acme Ltd", companyNumber: "C1", workforceNumber: "W2" },
  { id: "3", candidateName: "John Smith", dateOfBirth: "1990-05-05", companyName: "Beta", companyNumber: "C2", workforceNumber: "W3" },
  { id: "4", candidateName: "Jane Doe", dateOfBirth: "1985-03-03", companyName: "Acme Ltd", companyNumber: "C1", workforceNumber: "W4" },
  { id: "5", candidateName: "Dup One", dateOfBirth: "1970-01-01", companyName: "Acme Ltd", companyNumber: "C1", workforceNumber: "W9" },
  { id: "6", candidateName: "Dup Two", dateOfBirth: "1971-01-01", companyName: "Beta", companyNumber: "C2", workforceNumber: "W9" },
];

function match(fields) {
  return matchWorkforceForMatrixRow(workforce, {
    candidateName: null,
    dateOfBirth: null,
    company: null,
    companyNumber: null,
    workforceNumber: null,
    ...fields,
  });
}

test("candidate not found names the candidate", () => {
  const result = match({ candidateName: "Nobody Here", dateOfBirth: "2000-01-01" });
  assert.equal(result.candidate, null);
  assert.equal(result.ambiguous, false);
  assert.equal(result.reason.code, "candidateNotFound");
  assert.equal(result.reason.label, "Candidate not found");
  assert.match(result.reason.detail, /No Workforce record is named "Nobody Here"/);
});

test("candidate not found suggests a same-DOB spelling difference", () => {
  const result = match({
    candidateName: "Jayne Doe",
    dateOfBirth: "03/03/1985",
    company: "Acme Ltd",
  });
  assert.equal(result.reason.code, "candidateNotFound");
  assert.match(result.reason.detail, /Possible spelling difference: Jane Doe \(DOB 03\/03\/1985/);
});

test("DOB mismatch quotes both dates", () => {
  const result = match({ candidateName: "Zaid Khan", dateOfBirth: "1975-01-01" });
  assert.equal(result.candidate, null);
  assert.equal(result.reason.code, "dobMismatch");
  assert.match(result.reason.detail, /DOB \(01\/01\/1975\)/);
  assert.match(result.reason.detail, /Zaid Khan \(DOB 02\/02\/1980, Acme Ltd, Workforce # W1\)/);
});

test("DOB missing never links on name alone", () => {
  const result = match({ candidateName: "Zaid Khan" });
  assert.equal(result.candidate, null);
  assert.equal(result.reason.code, "dobMissing");
  assert.match(result.reason.detail, /no DOB/);
});

test("multiple matches without company stay ambiguous", () => {
  const result = match({ candidateName: "John Smith", dateOfBirth: "1990-05-05" });
  assert.equal(result.candidate, null);
  assert.equal(result.ambiguous, true);
  assert.equal(result.reason.code, "multipleMatches");
  assert.match(result.reason.detail, /2 Workforce records match Name \+ DOB/);
});

test("company mismatch when the sheet company matches none of the ties", () => {
  const result = match({
    candidateName: "John Smith",
    dateOfBirth: "1990-05-05",
    company: "Gamma Ltd",
  });
  assert.equal(result.ambiguous, true);
  assert.equal(result.reason.code, "companyMismatch");
  assert.match(result.reason.detail, /none is at "Gamma Ltd"/);
});

test("company narrows a tie to one record (no reason)", () => {
  const result = match({
    candidateName: "John Smith",
    dateOfBirth: "1990-05-05",
    company: "Beta",
  });
  assert.equal(result.candidate.id, "3");
  assert.equal(result.matchType, "companyNameDob");
  assert.equal(result.reason, null);
});

test("unique Name + DOB links; a stale Workforce Number still falls back", () => {
  const result = match({
    candidateName: "Zaid Khan",
    dateOfBirth: "02/02/1980",
    workforceNumber: "W404",
  });
  assert.equal(result.candidate.id, "1");
  assert.equal(result.matchType, "nameDob");
});

test("duplicate Workforce Number is a multiple match", () => {
  const result = match({ candidateName: "Anyone", workforceNumber: "W9" });
  assert.equal(result.ambiguous, true);
  assert.equal(result.reason.code, "multipleMatches");
  assert.match(result.reason.detail, /Workforce Number "W9" is on 2 Workforce records/);
});

test("unknown Workforce Number is reported alongside the name reason", () => {
  const result = match({
    candidateName: "Nobody Here",
    dateOfBirth: "2000-01-01",
    workforceNumber: "W404",
  });
  assert.match(result.reason.detail, /^Workforce Number "W404" is not on any Workforce record\. No Workforce record is named/);
});

const adminBase = {
  candidateName: "Zaid Khan",
  dateOfBirth: "1975-01-01",
  companyName: null,
  staleWorkforceItemId: null,
  sameName: [],
  duplicateOf: null,
  scopedCompany: null,
};

test("Admin Matrix: Linked rows carry no reason", () => {
  assert.equal(explainMatrixLinkStatus({ ...adminBase, status: "Linked" }), null);
});

test("Admin Matrix: orphan is candidate not found", () => {
  const reason = explainMatrixLinkStatus({ ...adminBase, status: "Orphan" });
  assert.equal(reason.code, "candidateNotFound");
});

test("Admin Matrix: deleted link target with a DOB mismatch", () => {
  const reason = explainMatrixLinkStatus({
    ...adminBase,
    status: "Needs Review",
    staleWorkforceItemId: "77",
    sameName: [workforce[0]],
  });
  assert.equal(reason.code, "linkedRecordMissing");
  assert.match(reason.detail, /\(#77\) no longer exists/);
  assert.match(reason.detail, /doesn't match the Workforce record named "Zaid Khan"/);
});

test("Admin Matrix: same-name pair where only one shares the DOB", () => {
  const reason = explainMatrixLinkStatus({
    ...adminBase,
    status: "Needs Review",
    candidateName: "John Smith",
    dateOfBirth: "1990-05-05",
    sameName: [workforce[1], { ...workforce[2], dateOfBirth: "1991-01-01" }],
  });
  assert.equal(reason.code, "multipleMatches");
  assert.match(reason.detail, /only John Smith \(DOB 05\/05\/1990, Acme Ltd/);
});

test("Admin Matrix: second copy of an already-linked person", () => {
  const reason = explainMatrixLinkStatus({
    ...adminBase,
    status: "Needs Review",
    duplicateOf: { person: workforce[0], matrixRowId: "12" },
  });
  assert.equal(reason.code, "duplicateMatrixRow");
  assert.match(reason.detail, /another Training Matrix row \(#12\)/);
});

test("upload results summarise reasons in the outcome line", () => {
  const dob = { code: "dobMismatch", label: "DOB mismatch", detail: "x" };
  const missing = { code: "candidateNotFound", label: "Candidate not found", detail: "y" };
  const rows = [
    { rowNumber: 2, status: "Imported", messages: [], fields: {} },
    { rowNumber: 3, status: "Warning", messages: [], fields: {}, reviewReason: dob },
    { rowNumber: 4, status: "Warning", messages: [], fields: {}, reviewReason: dob },
    { rowNumber: 5, status: "Warning", messages: [], fields: {}, reviewReason: missing },
  ];
  assert.equal(reviewReasonSummary(rows), "2 DOB mismatch, 1 Candidate not found");
  const { outcome, reason } = bulkCommitOutcome({
    importType: "trainingMatrix",
    rows,
    summary: {
      totalRows: 4, readyRows: 0, warningRows: 3, duplicateRows: 0,
      errorRows: 0, skippedRows: 0, importedRows: 1,
    },
  });
  assert.equal(outcome, "completed_with_warnings");
  assert.equal(
    reason,
    "4 Matrix rows uploaded, 1 linked, 3 need review (2 DOB mismatch, 1 Candidate not found).",
  );
});
