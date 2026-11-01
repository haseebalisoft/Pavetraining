/**
 * Job lifecycle tests: every Training Matrix upload / refresh must end in
 * Completed, Completed with warnings, or Failed — with a reason — and never
 * hang.
 *
 * Covers the pure pieces (outcome classification, time-budget stop) plus the
 * Graph timeout middleware against a local HTTP server that never answers.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-job-lifecycle.mjs
 *
 * (Node >= 22 for TypeScript type-stripping. The stub resolves "@/" and
 * replaces "server-only" so graphClient.ts can load outside Next.)
 */
import test from "node:test";
import assert from "node:assert/strict";
import http from "node:http";

const {
  bulkCommitOutcome,
  matrixSyncOutcome,
  createJobDeadline,
  describeJobTransportFailure,
  JobWatchdogError,
  NOT_PROCESSED_ROW_MESSAGE,
} = await import("@/lib/jobs/jobOutcome");
const { stopIfPastDeadline, buildNotProcessedRows, countNotProcessed } =
  await import("@/lib/services/bulkUpload/commitDeadline");

function summary(overrides = {}) {
  return {
    totalRows: 0,
    readyRows: 0,
    warningRows: 0,
    duplicateRows: 0,
    errorRows: 0,
    skippedRows: 0,
    importedRows: 0,
    ...overrides,
  };
}

function commitResult({ importType = "trainingMatrix", rows = [], s = {}, stoppedEarly = null } = {}) {
  return {
    importType,
    fileName: "matrix.xlsx",
    duplicateMode: "update",
    suppressNotifications: true,
    rows,
    summary: summary(s),
    stoppedEarly,
  };
}

function syncResult({ s = {}, items = [], incomplete = null, dryRun = false } = {}) {
  return {
    dryRun,
    scope: "all",
    items,
    summary: {
      processed: 0,
      updated: 0,
      created: 0,
      skipped: 0,
      errors: 0,
      warnings: 0,
      ...s,
    },
    incomplete,
  };
}

// ---------------------------------------------------------------- deadline

test("createJobDeadline expires on the injected clock", () => {
  const clock = { t: 1_000 };
  const deadline = createJobDeadline(500, () => clock.t);
  assert.equal(deadline.expired(), false);
  assert.equal(deadline.remainingMs(), 500);
  clock.t = 1_499;
  assert.equal(deadline.expired(), false);
  clock.t = 1_500;
  assert.equal(deadline.expired(), true);
  assert.equal(deadline.remainingMs(), 0);
});

test("stopIfPastDeadline returns null while time remains", () => {
  const deadline = createJobDeadline(10_000);
  const rows = [{ rowNumber: 2, fields: {} }];
  assert.equal(stopIfPastDeadline(deadline, rows, 0), null);
  assert.equal(stopIfPastDeadline(undefined, rows, 0), null);
});

test("stopIfPastDeadline marks every unreached row Not processed", () => {
  const clock = { t: 0 };
  const deadline = createJobDeadline(100, () => clock.t);
  clock.t = 100;
  const rows = [2, 3, 4, 5].map((rowNumber) => ({
    rowNumber,
    fields: { candidateName: `C${rowNumber}` },
    source: { Name: `C${rowNumber}` },
  }));
  const warned = [];
  const log = { warn: (event, details) => warned.push({ event, details }) };
  const stopped = stopIfPastDeadline(deadline, rows, 1, log);
  assert.deepEqual(
    stopped.map((row) => row.rowNumber),
    [3, 4, 5],
  );
  for (const row of stopped) {
    assert.equal(row.status, "Skipped");
    assert.equal(row.notProcessed, true);
    assert.deepEqual(row.messages, [NOT_PROCESSED_ROW_MESSAGE]);
  }
  assert.equal(countNotProcessed(stopped), 3);
  assert.deepEqual(warned, [
    {
      event: "stopped-early",
      details: { reason: "time-budget", processed: 1, notProcessed: 3 },
    },
  ]);
});

test("buildNotProcessedRows does not share field objects with the input", () => {
  const input = [{ rowNumber: 2, fields: { a: "1" } }];
  const [row] = buildNotProcessedRows(input);
  row.fields.a = "changed";
  assert.equal(input[0].fields.a, "1");
});

// ------------------------------------------------------ bulk commit outcome

function row(rowNumber, status, extra = {}) {
  return { rowNumber, status, messages: [], fields: {}, ...extra };
}

test("matrix upload with only linked rows is Completed", () => {
  const status = bulkCommitOutcome(
    commitResult({
      rows: [2, 3, 4].map((rowNumber) => row(rowNumber, "Imported")),
      s: { totalRows: 3, importedRows: 3 },
    }),
  );
  assert.equal(status.outcome, "completed");
  assert.equal(status.reason, "3 Matrix rows uploaded, 3 linked.");
});

test("matrix Needs Review rows count as saved → Completed with warnings", () => {
  const status = bulkCommitOutcome(
    commitResult({
      rows: [
        row(2, "Warning", { linkOutcome: "needsReviewNoMatch" }),
        row(3, "Warning", { linkOutcome: "needsReviewNoMatch" }),
        row(4, "Warning", { linkOutcome: "needsReviewMultipleMatches" }),
        row(5, "Skipped"),
      ],
      s: { totalRows: 4, importedRows: 0, warningRows: 3, skippedRows: 1 },
    }),
  );
  assert.equal(status.outcome, "completed_with_warnings");
  assert.equal(
    status.reason,
    "3 Matrix rows uploaded, 0 linked, 3 need review, 1 empty row skipped.",
  );
});

test("matrix counts are exact: not processed ≠ skipped, ambiguous rows are not 'uploaded'", () => {
  // Shadow run of the client's sheet: 3 linked, then the time limit hit.
  // It used to say "64 empty rows skipped" — 47 of those were never reached.
  const rows = [
    row(2, "Imported"),
    row(3, "Imported"),
    row(4, "Imported"),
    row(5, "Warning", { linkOutcome: "skippedAmbiguous" }),
    row(6, "Error"),
    ...Array.from({ length: 17 }, (_, i) => row(100 + i, "Skipped")),
    ...Array.from({ length: 47 }, (_, i) => row(200 + i, "Skipped", { notProcessed: true })),
  ];
  const status = bulkCommitOutcome(
    commitResult({
      rows,
      s: { totalRows: rows.length, importedRows: 3, warningRows: 1, errorRows: 1, skippedRows: 17, notProcessedRows: 47 },
      stoppedEarly: { reason: "the import reached the server time limit.", notProcessedRows: 47 },
    }),
  );
  assert.equal(status.outcome, "completed_with_warnings");
  assert.equal(
    status.reason,
    "3 Matrix rows uploaded, 3 linked, 17 empty rows skipped, 1 row not saved (several Matrix rows match — resolve in Admin Matrix), 1 failed. Stopped early: the import reached the server time limit. 47 rows not processed — re-upload the file to finish.",
  );
});

test("every row failing is Failed with the row's own reason", () => {
  const status = bulkCommitOutcome(
    commitResult({
      rows: [
        {
          rowNumber: 2,
          status: "Error",
          messages: ["Row 2", "Candidate not found in Workforce."],
          fields: {},
        },
      ],
      s: { totalRows: 1, errorRows: 1 },
    }),
  );
  assert.equal(status.outcome, "failed");
  assert.equal(status.reason, "Candidate not found in Workforce.");
});

test("time-budget stop with saved rows is Completed with warnings and says so", () => {
  const status = bulkCommitOutcome(
    commitResult({
      rows: [
        ...Array.from({ length: 6 }, (_, i) => row(2 + i, "Imported")),
        ...Array.from({ length: 4 }, (_, i) => row(20 + i, "Skipped", { notProcessed: true })),
      ],
      s: { totalRows: 10, importedRows: 6, notProcessedRows: 4 },
      stoppedEarly: {
        reason: "the import reached the server time limit.",
        notProcessedRows: 4,
      },
    }),
  );
  assert.equal(status.outcome, "completed_with_warnings");
  assert.match(status.reason, /Stopped early: the import reached the server time limit\. 4 rows not processed — re-upload the file to finish\./);
});

test("time-budget stop before any row saved is Failed", () => {
  const status = bulkCommitOutcome(
    commitResult({
      s: { totalRows: 5, skippedRows: 5 },
      stoppedEarly: {
        reason: "the import reached the server time limit.",
        notProcessedRows: 5,
      },
    }),
  );
  assert.equal(status.outcome, "failed");
  assert.match(status.reason, /^No rows were saved\. Stopped early/);
});

test("register import keeps its Imported/Skipped phrasing", () => {
  const ok = bulkCommitOutcome(
    commitResult({ importType: "npors", s: { importedRows: 2, skippedRows: 1 } }),
  );
  assert.equal(ok.outcome, "completed");
  assert.equal(ok.reason, "Import finished: 2 imported, 1 skipped.");

  const mixed = bulkCommitOutcome(
    commitResult({
      importType: "npors",
      s: { importedRows: 2, errorRows: 1 },
    }),
  );
  assert.equal(mixed.outcome, "completed_with_warnings");
  assert.equal(
    mixed.reason,
    "Import finished with errors: 2 imported, 0 skipped, 1 failed. Review row Messages.",
  );
});

// -------------------------------------------------------- matrix sync outcome

test("clean refresh is Completed", () => {
  const status = matrixSyncOutcome(
    syncResult({ s: { processed: 5, updated: 3, created: 1 } }),
    "Refresh from registers",
  );
  assert.equal(status.outcome, "completed");
  assert.equal(
    status.reason,
    "Refresh from registers: updated 3, created 1 (5 candidates checked).",
  );
});

test("refresh where every candidate errored is Failed with the first error", () => {
  const status = matrixSyncOutcome(
    syncResult({
      s: { processed: 2, errors: 2 },
      items: [{ errors: ["Matrix write failed: 423 Locked"] }, { errors: ["x"] }],
    }),
  );
  assert.equal(status.outcome, "failed");
  assert.equal(status.reason, "Refresh failed for 2 candidates: Matrix write failed: 423 Locked");
});

test("partial refresh at the time budget is Completed with warnings", () => {
  const status = matrixSyncOutcome(
    syncResult({
      s: { processed: 40, updated: 12 },
      incomplete: {
        reason: "the refresh reached the server time limit.",
        remainingCandidates: 25,
      },
    }),
  );
  assert.equal(status.outcome, "completed_with_warnings");
  assert.match(status.reason, /25 candidates not processed — run it again to finish\.$/);
});

test("refresh that reached no candidates before the budget is Failed", () => {
  const status = matrixSyncOutcome(
    syncResult({
      incomplete: {
        reason: "the refresh reached the server time limit.",
        remainingCandidates: 80,
      },
    }),
  );
  assert.equal(status.outcome, "failed");
});

// ------------------------------------------------------- transport failures

test("transport failures get an actionable reason", () => {
  const timeout = new DOMException("signal timed out", "TimeoutError");
  assert.match(
    describeJobTransportFailure(timeout, "refresh"),
    /^The refresh took too long and was stopped\. Some changes may already be saved/,
  );
  assert.match(
    describeJobTransportFailure(new TypeError("Failed to fetch"), "import"),
    /^Lost connection to the server during the import\./,
  );
  assert.match(
    describeJobTransportFailure(
      new JobWatchdogError("No response from the server for 45 seconds, so the import was stopped."),
      "import",
    ),
    /^No response from the server for 45 seconds, so the import was stopped\. Some rows may already be saved — run Preview again before retrying/,
  );
  assert.equal(
    describeJobTransportFailure(new DOMException("x", "AbortError"), "preview"),
    "The preview took too long and was stopped. Nothing was saved. Try again.",
  );
});

// ------------------------------------------------- Graph timeout middleware

test("Graph timeout middleware ends a request the server never answers", async (t) => {
  const { GraphTimeoutHandler, isGraphTimeoutError, isTransientGraphError } =
    await import("@/lib/graph/graphRequestGuards");
  const { HTTPMessageHandler } = await import(
    "@microsoft/microsoft-graph-client"
  );

  // Accepts the connection, never responds.
  const sockets = new Set();
  const server = http.createServer(() => {});
  server.on("connection", (socket) => sockets.add(socket));
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => {
    for (const socket of sockets) socket.destroy();
    server.close();
  });
  const { port } = server.address();

  const handler = new GraphTimeoutHandler(150);
  handler.setNext(new HTTPMessageHandler());
  const started = Date.now();
  const error = await handler
    .execute({
      request: `http://127.0.0.1:${port}/v1.0/sites/x/lists/y/items`,
      options: { method: "POST" },
    })
    .then(() => null, (e) => e);
  const elapsed = Date.now() - started;

  assert.ok(error, "request should have been aborted");
  assert.equal(error.name, "GraphTimeoutError");
  assert.match(error.message, /did not respond within 150ms \(POST \/sites\/x\/lists\/y\/items\)/);
  assert.ok(elapsed < 2_000, `aborted after ${elapsed}ms`);
  assert.equal(isGraphTimeoutError(error), true);
  // Never replayed by withGraphMutationRetry — even once the SDK re-wraps it
  // as GraphError(statusCode -1, code "GraphTimeoutError").
  assert.equal(isTransientGraphError(error), false);
  assert.equal(
    isTransientGraphError({ statusCode: -1, code: "GraphTimeoutError", message: error.message }),
    false,
  );
  // Plain -1 network drops are still retried as before.
  assert.equal(isTransientGraphError({ statusCode: -1, message: "fetch failed" }), true);
});

test("Graph timeout middleware leaves fast responses and caller signals alone", async (t) => {
  const { GraphTimeoutHandler } = await import("@/lib/graph/graphRequestGuards");
  const { HTTPMessageHandler } = await import("@microsoft/microsoft-graph-client");
  const server = http.createServer((_req, res) => res.end("{}"));
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => server.close());
  const { port } = server.address();

  const handler = new GraphTimeoutHandler(1_000);
  handler.setNext(new HTTPMessageHandler());
  const context = { request: `http://127.0.0.1:${port}/v1.0/ok`, options: {} };
  await handler.execute(context);
  assert.equal(context.response.status, 200);

  const own = new AbortController();
  const withSignal = {
    request: `http://127.0.0.1:${port}/v1.0/ok`,
    options: { signal: own.signal },
  };
  await handler.execute(withSignal);
  assert.equal(withSignal.options.signal, own.signal);
});

test("matrix rows saved but category backups pending are reported, not hidden", () => {
  const status = bulkCommitOutcome(
    commitResult({
      rows: [
        row(2, "Imported"),
        row(3, "Imported", { backupPending: true }),
        row(4, "Imported", { backupPending: true }),
      ],
      s: { totalRows: 3, importedRows: 3 },
    }),
  );
  assert.equal(status.outcome, "completed_with_warnings");
  assert.equal(
    status.reason,
    "3 Matrix rows uploaded, 3 linked. Category backup records still to write for 2 rows (their Matrix dates are saved) — re-upload the file to finish.",
  );
});
