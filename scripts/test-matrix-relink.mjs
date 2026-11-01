/**
 * Matrix re-linking after an admin corrects a candidate's name or DOB, and the
 * bulk-upload job store's retry guards.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-matrix-relink.mjs
 *
 * (Node >= 22 for TypeScript type-stripping; the stub resolves "@/".)
 */
import test from "node:test";
import assert from "node:assert/strict";

const { planMatrixRelink, mergeMatrixCells } = await import(
  "@/lib/services/bulkUpload/matrixRelink"
);
const job = await import("@/components/admin/bulkUpload/bulkCommitJob");

const workforce = [
  { id: "10", candidateName: "Zaid Khan", dateOfBirth: "1972-01-01", companyName: "Acme Ltd", companyNumber: "C1", workforceNumber: "W10" },
  { id: "11", candidateName: "Amy Lee", dateOfBirth: "1990-02-02", companyName: "Acme Ltd", companyNumber: "C1", workforceNumber: "W11" },
];

function matrixRow(overrides) {
  return {
    id: "1",
    candidateName: "Zaid Khan",
    dateOfBirth: "1972-01-01",
    workforceItemId: null,
    companyNumber: null,
    storedCompanyName: null,
    columnValues: {},
    manualOverrides: [],
    categoryTrainingDates: {},
    ...overrides,
  };
}

test("corrected DOB now matching one free candidate links in place", () => {
  const row = matrixRow({ id: "5" });
  const plan = planMatrixRelink(row, workforce, [row]);
  assert.equal(plan.action, "link");
  assert.equal(plan.candidate.id, "10");
});

test("a row linked to a live record never moves", () => {
  const row = matrixRow({ id: "5", candidateName: "Amy Lee", workforceItemId: "10" });
  const plan = planMatrixRelink(row, workforce, [row]);
  assert.equal(plan.action, "alreadyLinked");
  assert.equal(plan.candidate.id, "10");
});

test("still-wrong DOB stays unmatched with the exact reason", () => {
  const row = matrixRow({ id: "5", dateOfBirth: "1975-01-01" });
  const plan = planMatrixRelink(row, workforce, [row]);
  assert.equal(plan.action, "unmatched");
  assert.equal(plan.reason.code, "dobMismatch");
  assert.match(plan.reason.detail, /01\/01\/1975/);
});

test("a stale Workforce Number is never used to re-link", () => {
  // Record W99 was deleted; W11 now belongs to Amy. The row must not jump to her.
  const row = matrixRow({ id: "5", candidateName: "Ghost Person", workforceItemId: "99" });
  const plan = planMatrixRelink(row, workforce, [row]);
  assert.equal(plan.action, "unmatched");
  assert.equal(plan.reason.code, "candidateNotFound");
});

test("candidate who already owns a linked row → merge, never a second link", () => {
  const owned = matrixRow({ id: "2", workforceItemId: "10" });
  const stray = matrixRow({ id: "7" });
  const plan = planMatrixRelink(stray, workforce, [owned, stray]);
  assert.equal(plan.action, "merge");
  assert.equal(plan.target.id, "2");
});

test("merge keeps every expiry: blanks filled, later wins, Manual untouched", () => {
  const target = matrixRow({
    id: "2",
    columnValues: {
      "CSCS Expiry": "2027-01-05",
      "SSSTS Expiry": "2026-01-01",
      "SMSTS Expiry": "2028-01-01",
      "EUSR Expiry": null,
    },
    manualOverrides: ["SMSTS Expiry"],
    categoryTrainingDates: { "N001 - Ind FLT": "2024-01-01" },
  });
  const duplicate = matrixRow({
    id: "7",
    columnValues: {
      Name: "Zaid",
      DOB: "1972-01-01",
      "CSCS Expiry": "2026-06-01", // older → target keeps 2027
      "SSSTS Expiry": "2026-10-27", // newer → wins
      "SMSTS Expiry": "2030-01-01", // target is Manual → untouched
      "EUSR Expiry": "05/01/2027", // target blank → filled
      Company: "Acme Ltd", // not a date → ignored
    },
    manualOverrides: ["EUSR Expiry"],
    categoryTrainingDates: { "N001 - Ind FLT": "2020-01-01", "N003 - Reach Lift Truck": "2023-03-03" },
  });
  const headers = ["Name", "DOB", "CSCS Expiry", "SSSTS Expiry", "SMSTS Expiry", "EUSR Expiry", "Company"];
  const merged = mergeMatrixCells(target, duplicate, headers);
  assert.deepEqual(merged.cells, {
    "SSSTS Expiry": "2026-10-27",
    "EUSR Expiry": "2027-01-05",
  });
  assert.deepEqual(merged.mergedHeaders.sort(), ["EUSR Expiry", "SSSTS Expiry"]);
  assert.deepEqual(merged.manualOverrides.sort(), ["EUSR Expiry", "SMSTS Expiry"]);
  assert.deepEqual(merged.categoryTrainingDates, {
    "N001 - Ind FLT": "2024-01-01",
    "N003 - Reach Lift Truck": "2023-03-03",
  });
});

// ---- bulk-upload job store -------------------------------------------------

const preview = {
  implemented: true,
  importType: "trainingMatrix",
  fileName: "matrix.xlsx",
  headers: [],
  rows: [{ rowNumber: 2, status: "Ready", messages: [], fields: { candidateName: "Zaid" } }],
  summary: { totalRows: 1, readyRows: 1, warningRows: 0, duplicateRows: 0, errorRows: 0, skippedRows: 0, importedRows: 0 },
};

function ndjsonResponse(lines) {
  const body = new ReadableStream({
    start(controller) {
      for (const line of lines) {
        controller.enqueue(new TextEncoder().encode(`${JSON.stringify(line)}\n`));
      }
      controller.close();
    },
  });
  return new Response(body, { status: 200 });
}

function waitForDone() {
  return new Promise((resolve) => {
    const check = () => {
      const current = job.getBulkCommitJob();
      if (current && current.phase === "done") return resolve(current);
      setTimeout(check, 5);
    };
    check();
  });
}

function start(notes) {
  return job.startBulkCommit({
    preview,
    file: null,
    duplicateMode: "skip",
    suppressNotifications: true,
    logoFiles: [],
    notify: (message, tone) => notes.push({ message, tone }),
  });
}

test("job store: result ends the job, and only one job runs at a time", async () => {
  const notes = [];
  let release;
  globalThis.fetch = () =>
    new Promise((resolve) => {
      release = () =>
        resolve(
          ndjsonResponse([
            { kind: "progress", event: { type: "phase-start", label: "rows" } },
            {
              kind: "result",
              result: {
                importType: "trainingMatrix",
                rows: [{ rowNumber: 2, status: "Imported", messages: [], fields: {} }],
                summary: { ...preview.summary, readyRows: 0, importedRows: 1 },
              },
            },
          ]),
        );
    });
  job.clearBulkCommitJob();
  assert.equal(start(notes), true);
  assert.equal(job.bulkCommitJobActive(), true);
  assert.equal(start(notes), false, "second commit refused while one runs");
  assert.match(notes[0].message, /already running/);
  job.clearBulkCommitJob();
  assert.ok(job.getBulkCommitJob(), "a running job cannot be cleared");
  release();
  const done = await waitForDone();
  assert.equal(done.outcome.outcome, "completed");
  assert.equal(done.needsRepreview, false);
  assert.equal(job.bulkCommitJobActive(), false);
});

test("job store: connection lost mid-import requires a fresh preview", async () => {
  globalThis.fetch = async () =>
    ndjsonResponse([{ kind: "progress", event: { type: "phase-start", label: "rows" } }]);
  job.clearBulkCommitJob();
  start([]);
  const done = await waitForDone();
  assert.equal(done.outcome.outcome, "failed");
  assert.equal(done.needsRepreview, true);
  assert.match(done.outcome.reason, /run Preview again before retrying/);
});

test("job store: a refused request saved nothing, so retry stays allowed", async () => {
  globalThis.fetch = async () =>
    new Response(JSON.stringify({ error: "This file is already being imported." }), {
      status: 409,
      headers: { "Content-Type": "application/json" },
    });
  job.clearBulkCommitJob();
  start([]);
  const done = await waitForDone();
  assert.equal(done.outcome.outcome, "failed");
  assert.equal(done.needsRepreview, false);
});

test("job store: an import with an image step finishes 'done' with its outcome", async () => {
  const phases = [];
  const unsubscribe = job.subscribeBulkCommitJob(() => {
    const current = job.getBulkCommitJob();
    if (current && phases.at(-1) !== current.phase) phases.push(current.phase);
  });
  globalThis.fetch = async () =>
    ndjsonResponse([
      {
        kind: "result",
        result: {
          importType: "workforce",
          rows: [{ rowNumber: 2, status: "Imported", savedEntityId: "1601", messages: [], fields: { workforceNumber: "W1" } }],
          summary: { ...preview.summary, readyRows: 0, importedRows: 1 },
        },
      },
    ]);
  job.clearBulkCommitJob();
  job.startBulkCommit({
    preview: { ...preview, importType: "workforce" },
    file: null,
    duplicateMode: "skip",
    suppressNotifications: true,
    logoFiles: [],
    photoFiles: [],
    notify: () => {},
  });
  const done = await waitForDone();
  unsubscribe();
  assert.deepEqual(phases, ["running", "finishing", "done"]);
  assert.equal(done.outcome.outcome, "completed");
  assert.equal(done.result.rows[0].savedEntityId, "1601");
  assert.equal(job.bulkCommitJobActive(), false, "navigation/unload is free again");
});
