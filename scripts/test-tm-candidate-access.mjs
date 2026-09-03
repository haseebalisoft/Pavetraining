/**
 * Training Manager candidate visibility — offline unit tests.
 *
 *   node --import ./scripts/_register-e2e-hook.mjs --test scripts/test-tm-candidate-access.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const { candidateRecordAllowed, filterCandidatesByAccess } = await import(
  new URL("../src/lib/services/customerAccessService.ts", import.meta.url)
);
const { normalizeAccessScopeValue, resolveCustomerRole } = await import(
  new URL("../src/lib/services/permissionService.ts", import.meta.url)
);

function tmContext(overrides = {}) {
  return {
    loggedInEmail: "tm@dbs.org",
    role: "Customer",
    customerRole: "TrainingManager",
    roleLabel: "Manager",
    companyId: "1",
    companyName: "DBS",
    canView: true,
    canDownload: true,
    canEdit: false,
    accessScope: "Full Company",
    normalizedAccessScope: "Company",
    departmentScopes: [],
    candidateScopeName: "Pat Lead",
    permissionStatus: "Active",
    ...overrides,
  };
}

function candidate(overrides = {}) {
  return {
    id: "10",
    candidateName: "Wayne Wonder",
    companyName: "DBS",
    department: "Plant Operations",
    trainingManager: "Pat Lead",
    supervisor: "Sam Super",
    email: null,
    ...overrides,
  };
}

test("Training Manager Full Company sees every company candidate", () => {
  const ctx = tmContext({ normalizedAccessScope: "Company" });
  const rows = [
    candidate({ id: "1", trainingManager: "Someone Else" }),
    candidate({ id: "2", trainingManager: null, department: "Other" }),
  ];
  assert.equal(filterCandidatesByAccess(rows, ctx).length, 2);
});

test("Training Manager Department Only sees matching department", () => {
  const ctx = tmContext({
    accessScope: "Department Only",
    normalizedAccessScope: "Department",
    departmentScopes: ["Plant Operations"],
  });
  assert.equal(candidateRecordAllowed(candidate(), ctx), true);
  assert.equal(
    candidateRecordAllowed(
      candidate({ department: "Accounts", trainingManager: "Other" }),
      ctx,
    ),
    false,
  );
});

test("Training Manager sees candidates assigned to them even without department match", () => {
  const ctx = tmContext({
    accessScope: "Department Only",
    normalizedAccessScope: "Department",
    departmentScopes: ["Accounts"],
    candidateScopeName: "Pat Lead",
  });
  assert.equal(
    candidateRecordAllowed(
      candidate({ department: "Plant Operations", trainingManager: "Pat Lead" }),
      ctx,
    ),
    true,
  );
  assert.equal(
    candidateRecordAllowed(
      candidate({
        department: "Plant Operations",
        trainingManager: "Other TM",
        supervisor: "Other",
      }),
      ctx,
    ),
    false,
  );
});

test("AssignedCandidates scope includes Training Manager assignments", () => {
  const ctx = tmContext({
    accessScope: "Assigned Candidates",
    normalizedAccessScope: "AssignedCandidates",
    departmentScopes: [],
    candidateScopeName: "Pat Lead",
  });
  assert.equal(
    candidateRecordAllowed(candidate({ trainingManager: "Pat Lead" }), ctx),
    true,
  );
  assert.equal(
    candidateRecordAllowed(candidate({ trainingManager: "Other" }), ctx),
    false,
  );
});

test("normalizeAccessScopeValue: Training Manager defaults to Company", () => {
  assert.equal(
    normalizeAccessScopeValue("Full Company", "TrainingManager", false),
    "Company",
  );
  assert.equal(
    normalizeAccessScopeValue("", "TrainingManager", false),
    "Company",
  );
  assert.equal(
    normalizeAccessScopeValue("Department Only", "TrainingManager", false),
    "Department",
  );
});

test("resolveCustomerRole treats Manager as TrainingManager", () => {
  assert.equal(resolveCustomerRole("Training Manager", "Full Company"), "TrainingManager");
  assert.equal(resolveCustomerRole("Manager", "Full Company"), "TrainingManager");
  assert.equal(resolveCustomerRole("manager", "Full Company"), "TrainingManager");
});
