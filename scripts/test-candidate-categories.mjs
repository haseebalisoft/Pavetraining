/**
 * Unit tests for the candidate-profile category aggregation.
 *
 *   node --import ./scripts/_register-e2e-hook.mjs --test scripts/test-candidate-categories.mjs
 *
 * candidateCategories.ts imports its sibling expiryFilters.ts via the "@/"
 * alias (same as the rest of the app), so it needs the shared resolver hook
 * rather than running as a fully dependency-free module.
 */
import test from "node:test";
import assert from "node:assert/strict";

const { buildCandidateCategoryRows } = await import(
  new URL("../src/lib/training/candidateCategories.ts", import.meta.url)
);

function isoDaysFromNow(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

test("sorts expired before expiring-soon before active before missing", () => {
  const rows = buildCandidateCategoryRows({
    nporsRecords: [
      {
        id: "n1",
        candidateName: "A",
        workforceId: "1",
        nporsNumber: "1",
        trainingDate: null,
        trainingAddress: null,
        noviceOrEwt: null,
        nporsCategory: "Active Plant",
        outcome: "Pass",
        expiry: isoDaysFromNow(400),
      },
    ],
    eusrRecords: [
      {
        id: "e1",
        candidateName: "A",
        workforceId: "1",
        eusrNumber: "1",
        eusrCategory: "Expired Cat",
        cardType: null,
        trainingDate: null,
        trainingAddress: null,
        outcome: "Pass",
        expiry: isoDaysFromNow(-10),
      },
    ],
    streetworksRecords: [
      {
        id: "s1",
        candidateName: "A",
        workforceId: "1",
        swqrNumber: "1",
        course: "Streetworks Cat",
        streetworksCategory: null,
        trainingDate: null,
        trainingDateEnd: null,
        trainingAddress: null,
        outcome: "Pass",
        expiry: isoDaysFromNow(30),
      },
    ],
    nvqRecords: [
      {
        id: "v1",
        candidateName: "A",
        nvqTitle: "NVQ Level 2",
        boltOn: null,
        dateRegistered: "2024-01-01",
        inductionDate: null,
        stageOfNvq: null,
        notes: null,
        completedDate: null,
        status: "Active",
      },
    ],
  });

  assert.equal(rows.length, 4);
  assert.equal(rows[0].category, "Expired Cat");
  assert.equal(rows[0].expiryStatus.status, "expired");
  assert.equal(rows[1].category, "Streetworks Cat");
  assert.equal(rows[1].expiryStatus.status, "urgent");
  assert.equal(rows[2].category, "Active Plant");
  assert.equal(rows[2].expiryStatus.status, "valid");
  assert.equal(rows[3].category, "NVQ Level 2");
  assert.equal(rows[3].expiryStatus.status, "missing");
  assert.equal(rows[3].outcomeLabel, "Active");
});

test("matrix-only categories appear only when a date is present, never duplicating a register", () => {
  const rows = buildCandidateCategoryRows({
    matrixRow: {
      id: "m1",
      candidateId: "1",
      candidateName: "A",
      companyName: "Co",
      dateOfBirth: null,
      department: null,
      trainingManager: null,
      supervisor: null,
      overallStatus: null,
      needsReview: false,
      nextExpiryDate: null,
      nporsCategories: null,
      nporsExpiry: null,
      nporsNumber: null,
      cscsNumber: null,
      cscsExpiry: isoDaysFromNow(200),
      swqrNumber: null,
      swqrExpiry: null,
      eusrNumber: null,
      eusrExpiry: null,
      inHouseCourse: null,
      inHouseExpiry: null,
      n001Expiry: isoDaysFromNow(50),
      n003Expiry: null,
      n004Expiry: null,
      n010Expiry: null,
      n020Expiry: null,
      n021Expiry: null,
      n027Expiry: null,
      n100Expiry: null,
    },
  });

  assert.equal(rows.length, 2);
  const categories = rows.map((row) => row.category).sort();
  assert.deepEqual(categories, ["CSCS", "N001 - Ind FLT"]);
});

test("no matrix row and no records produces an empty list", () => {
  const rows = buildCandidateCategoryRows({});
  assert.deepEqual(rows, []);
});
