/**
 * Regression: editing registers / Training Matrix must change what the
 * candidate profile "All Training Categories" section shows.
 *
 *   node --import ./scripts/_register-e2e-hook.mjs --test scripts/test-profile-updates-from-edits.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const { buildCandidateCategoryRows } = await import(
  new URL("../src/lib/training/candidateCategories.ts", import.meta.url)
);

function isoDaysFromNow(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

function blankMatrix(overrides = {}) {
  return {
    id: "m1",
    candidateId: "1",
    candidateName: "Wayne Wonder",
    companyName: "DBS",
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
    cscsExpiry: null,
    swqrNumber: null,
    swqrExpiry: null,
    eusrNumber: null,
    eusrExpiry: null,
    eusrCategoryRows: [],
    inHouseCourse: null,
    inHouseExpiry: null,
    n001Expiry: null,
    n003Expiry: null,
    n004Expiry: null,
    n010Expiry: null,
    n020Expiry: null,
    n021Expiry: null,
    n027Expiry: null,
    n100Expiry: null,
    columnValues: {},
    categoryTrainingDates: {},
    ...overrides,
  };
}

test("NPORS register edit appears on profile categories (including expired)", () => {
  const before = buildCandidateCategoryRows({ nporsRecords: [] });
  assert.equal(before.length, 0);

  const afterCreate = buildCandidateCategoryRows({
    nporsRecords: [
      {
        id: "n1",
        candidateName: "Wayne Wonder",
        workforceId: "1",
        nporsNumber: "2365632",
        trainingDate: "2024-01-10",
        trainingAddress: null,
        noviceOrEwt: null,
        nporsCategory: "N001 - Ind FLT",
        outcome: "Pass",
        expiry: isoDaysFromNow(200),
      },
    ],
  });
  assert.equal(afterCreate.length, 1);
  assert.equal(afterCreate[0].category, "N001 - Ind FLT");
  assert.equal(afterCreate[0].source, "NPORS");
  assert.equal(afterCreate[0].expiryStatus.status, "valid");

  const afterEdit = buildCandidateCategoryRows({
    nporsRecords: [
      {
        id: "n1",
        candidateName: "Wayne Wonder",
        workforceId: "1",
        nporsNumber: "2365632",
        trainingDate: "2024-01-10",
        trainingAddress: null,
        noviceOrEwt: null,
        nporsCategory: "N001 - Ind FLT",
        outcome: "Pass",
        expiry: isoDaysFromNow(-40),
      },
    ],
  });
  assert.equal(afterEdit.length, 1);
  assert.equal(afterEdit[0].expiryStatus.status, "expired");
  assert.equal(afterEdit[0].expiryDate, isoDaysFromNow(-40));
});

test("EUSR / Streetworks / In-House / NVQ register edits feed profile categories", () => {
  const rows = buildCandidateCategoryRows({
    eusrRecords: [
      {
        id: "e1",
        candidateName: "Wayne Wonder",
        workforceId: "1",
        eusrNumber: "1",
        eusrCategory: "SHEA Gas",
        cardType: null,
        trainingDate: "2025-06-01",
        trainingAddress: null,
        outcome: "Pass",
        expiry: isoDaysFromNow(120),
      },
    ],
    streetworksRecords: [
      {
        id: "s1",
        candidateName: "Wayne Wonder",
        workforceId: "1",
        swqrNumber: "1",
        course: "Operative",
        streetworksCategory: "LA + O1",
        trainingDate: "2023-01-01",
        trainingDateEnd: null,
        trainingAddress: null,
        outcome: "Pass",
        expiry: isoDaysFromNow(-5),
      },
    ],
    inHouseRecords: [
      {
        id: "i1",
        candidateName: "Wayne Wonder",
        workforceId: "1",
        course: "Asbestos Awareness",
        certificationNumber: null,
        trainingDate: "2024-02-02",
        trainingAddress: null,
        outcome: "Pass",
        expiry: isoDaysFromNow(50),
      },
    ],
    nvqRecords: [
      {
        id: "v1",
        candidateName: "Wayne Wonder",
        nvqTitle: "NVQ Level 2 Plant",
        boltOn: null,
        dateRegistered: "2024-03-03",
        inductionDate: null,
        stageOfNvq: null,
        notes: null,
        completedDate: null,
        status: "Active",
      },
    ],
  });

  const byCategory = Object.fromEntries(rows.map((row) => [row.category, row]));
  assert.equal(byCategory["SHEA Gas"].source, "EUSR");
  assert.equal(byCategory["LA + O1"].expiryStatus.status, "expired");
  assert.equal(byCategory["Asbestos Awareness"].source, "In-House");
  assert.equal(byCategory["NVQ Level 2 Plant"].source, "NVQ");
  assert.equal(rows.length, 4);
});

test("Training Matrix edit updates profile categories when no register covers them", () => {
  const before = buildCandidateCategoryRows({
    matrixRow: blankMatrix(),
  });
  assert.equal(before.length, 0);

  const afterExpiryEdit = buildCandidateCategoryRows({
    matrixRow: blankMatrix({
      cscsExpiry: isoDaysFromNow(90),
      columnValues: {
        "CSCS Expiry": isoDaysFromNow(90),
        "SSSTS Expiry": isoDaysFromNow(-15),
      },
      categoryTrainingDates: {
        "N009 - Rough Terrain Lift Truck": "2024-08-01",
      },
    }),
  });

  const byCategory = Object.fromEntries(
    afterExpiryEdit.map((row) => [row.category, row]),
  );
  assert.equal(byCategory.CSCS.expiryDate, isoDaysFromNow(90));
  assert.equal(byCategory.SSSTS.expiryStatus.status, "expired");
  assert.equal(
    byCategory["N009 - Rough Terrain Lift Truck"].trainingDate,
    "2024-08-01",
  );
});

test("Matrix edit is superseded by dated register for same NPORS category", () => {
  const rows = buildCandidateCategoryRows({
    nporsRecords: [
      {
        id: "n1",
        candidateName: "Wayne Wonder",
        workforceId: "1",
        nporsNumber: "1",
        trainingDate: "2025-01-01",
        trainingAddress: null,
        noviceOrEwt: null,
        nporsCategory: "N001 - Ind FLT",
        outcome: "Pass",
        expiry: isoDaysFromNow(300),
      },
    ],
    matrixRow: blankMatrix({
      n001Expiry: isoDaysFromNow(-100),
      columnValues: { "N001 - Ind FLT": isoDaysFromNow(-100) },
    }),
  });

  assert.equal(rows.length, 1);
  assert.equal(rows[0].source, "NPORS");
  assert.equal(rows[0].expiryStatus.status, "valid");
});

test("register API routes wire matrix sync after create/update/delete", () => {
  const registers = ["npors", "eusr", "streetworks", "in-house"];
  for (const name of registers) {
    const collection = readFileSync(
      resolve(ROOT, `src/app/api/admin/training-records/${name}/route.ts`),
      "utf8",
    );
    const item = readFileSync(
      resolve(ROOT, `src/app/api/admin/training-records/${name}/[id]/route.ts`),
      "utf8",
    );
    assert.match(
      collection,
      /triggerMatrixSyncAfterRegister/,
      `${name} POST must sync matrix`,
    );
    assert.match(
      item,
      /triggerMatrixSyncAfterRegister/,
      `${name} PATCH must sync matrix`,
    );
    assert.match(
      item,
      /triggerMatrixSyncAfterRegisterDelete/,
      `${name} DELETE must recompute matrix`,
    );
  }

  const nvqCreate = readFileSync(
    resolve(ROOT, "src/app/api/admin/nvq/route.ts"),
    "utf8",
  );
  const nvqUpdate = readFileSync(
    resolve(ROOT, "src/app/api/admin/nvq/[id]/route.ts"),
    "utf8",
  );
  assert.match(nvqCreate, /triggerMatrixSyncAfterNvq/);
  assert.match(nvqUpdate, /triggerMatrixSyncAfterNvq/);

  const matrixUpdate = readFileSync(
    resolve(ROOT, "src/app/api/admin/training-matrix/[id]/route.ts"),
    "utf8",
  );
  assert.match(matrixUpdate, /updateAdminMatrix/);
});

test("revalidateTrainingSurfaces covers workforce + customer candidate layouts", async () => {
  const source = readFileSync(
    resolve(ROOT, "src/lib/cache/revalidateTrainingSurfaces.ts"),
    "utf8",
  );
  assert.match(source, /\/admin\/workforce/);
  assert.match(source, /\/customer\/candidates/);
  assert.match(source, /\/admin\/training-matrix/);
  assert.match(source, /\/admin\/training-records/);
  assert.match(source, /nporsRegister/);
  assert.match(source, /trainingMatrixExample/);
  assert.match(source, /revalidatePath/);
});

test("candidate profile mounts LiveTrainingRefresh for tab return updates", () => {
  const profile = readFileSync(
    resolve(ROOT, "src/components/customer/CandidateProfileView.tsx"),
    "utf8",
  );
  assert.match(profile, /LiveTrainingRefresh/);
  assert.match(profile, /buildCandidateCategoryRows/);
});
