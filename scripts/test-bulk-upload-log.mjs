/**
 * Unit tests for the bulk-upload terminal logger.
 *
 * bulkUploadLog.ts is dependency-free (no "server-only", no Graph, no "@/"),
 * so it loads directly with the built-in Node test runner + TypeScript
 * type-stripping (Node >= 22):
 *
 *   node --test scripts/test-bulk-upload-log.mjs
 *
 * A fake clock + capturing sink make the elapsed timings and formatting fully
 * deterministic.
 */
import test from "node:test";
import assert from "node:assert/strict";

const BASE = new URL("../src/lib/", import.meta.url).pathname;
const { createBulkLogger } = await import(
  BASE + "services/bulkUpload/bulkUploadLog.ts"
);

/** Mutable clock + a sink that records every line by level. */
function harness(startAt = 0) {
  const clock = { t: startAt };
  const lines = { info: [], warn: [], error: [] };
  const sink = {
    info: (m) => lines.info.push(m),
    warn: (m) => lines.warn.push(m),
    error: (m) => lines.error.push(m),
  };
  return { clock, lines, sink, now: () => clock.t };
}

test("info line: scope, elapsed, event, and key=value details", () => {
  const h = harness(1000);
  const log = createBulkLogger("commit:workforce", {
    sink: h.sink,
    now: h.now,
    enabled: true,
  });
  h.clock.t = 1000; // +0ms
  log.info("start", { rows: 50, duplicateMode: "skip" });
  assert.equal(
    h.lines.info[0],
    "[bulk:commit:workforce] +0ms start rows=50 duplicateMode=skip",
  );
});

test("elapsed is measured from creation and rounded to whole ms", () => {
  const h = harness(500);
  const log = createBulkLogger("t", { sink: h.sink, now: h.now, enabled: true });
  h.clock.t = 500 + 1234.6;
  assert.equal(log.elapsed(), 1235);
  log.info("tick");
  assert.equal(h.lines.info[0], "[bulk:t] +1235ms tick");
});

test("phase logs start + done and returns the duration", () => {
  const h = harness(0);
  const log = createBulkLogger("t", { sink: h.sink, now: h.now, enabled: true });
  const phase = log.phase("phase3b:creates");
  assert.equal(h.lines.info[0], "[bulk:t] +0ms phase phase3b:creates start");
  h.clock.t = 2500;
  const ms = phase.end({ created: 48, errors: 2 });
  assert.equal(ms, 2500);
  assert.equal(
    h.lines.info[1],
    "[bulk:t] +2500ms phase phase3b:creates done ms=2500 created=48 errors=2",
  );
});

test("detail formatting: strings with spaces quoted, null kept, undefined dropped", () => {
  const h = harness(0);
  const log = createBulkLogger("t", { sink: h.sink, now: h.now, enabled: true });
  log.info("done", {
    fileName: "my file.csv",
    plain: "skip",
    missing: null,
    omit: undefined,
    ok: true,
  });
  assert.equal(
    h.lines.info[0],
    '[bulk:t] +0ms done fileName="my file.csv" plain=skip missing=null ok=true',
  );
});

test("debug is suppressed by default and emitted only in verbose mode", () => {
  const quiet = harness(0);
  const q = createBulkLogger("t", {
    sink: quiet.sink,
    now: quiet.now,
    enabled: true,
    verbose: false,
  });
  q.debug("row created", { row: 3 });
  assert.equal(quiet.lines.info.length, 0);

  const loud = harness(0);
  const l = createBulkLogger("t", {
    sink: loud.sink,
    now: loud.now,
    enabled: true,
    verbose: true,
  });
  l.debug("row created", { row: 3 });
  assert.equal(loud.lines.info[0], "[bulk:t] +0ms row created row=3");
});

test("enabled:false silences info/warn/error/phase", () => {
  const h = harness(0);
  const log = createBulkLogger("t", {
    sink: h.sink,
    now: h.now,
    enabled: false,
  });
  log.info("a");
  log.warn("b");
  log.error("c");
  log.phase("p").end();
  assert.equal(h.lines.info.length, 0);
  assert.equal(h.lines.warn.length, 0);
  assert.equal(h.lines.error.length, 0);
});

test("warn and error route to the matching sink channel", () => {
  const h = harness(0);
  const log = createBulkLogger("t", { sink: h.sink, now: h.now, enabled: true });
  log.warn("create failed", { row: 7, error: "boom" });
  log.error("fatal");
  assert.equal(h.lines.warn[0], "[bulk:t] +0ms create failed row=7 error=boom");
  assert.equal(h.lines.error[0], "[bulk:t] +0ms fatal");
  assert.equal(h.lines.info.length, 0);
});

test("onEvent receives structured log + phase events (for streaming to the UI)", () => {
  const h = harness(0);
  const events = [];
  const log = createBulkLogger("commit:workforce", {
    sink: h.sink,
    now: h.now,
    enabled: true,
    onEvent: (event) => events.push(event),
  });

  log.info("start", { rows: 50 });
  const phase = log.phase("phase3b:creates");
  h.clock.t = 1200;
  log.info("progress", { done: 10, total: 48 });
  h.clock.t = 3000;
  phase.end({ created: 48 });

  assert.deepEqual(events[0], {
    type: "log",
    level: "info",
    scope: "commit:workforce",
    elapsedMs: 0,
    event: "start",
    details: { rows: 50 },
  });
  assert.deepEqual(events[1], {
    type: "phase-start",
    scope: "commit:workforce",
    elapsedMs: 0,
    label: "phase3b:creates",
  });
  assert.deepEqual(events[2], {
    type: "log",
    level: "info",
    scope: "commit:workforce",
    elapsedMs: 1200,
    event: "progress",
    details: { done: 10, total: 48 },
  });
  assert.deepEqual(events[3], {
    type: "phase-end",
    scope: "commit:workforce",
    elapsedMs: 3000,
    label: "phase3b:creates",
    ms: 3000,
    details: { created: 48 },
  });
});

test("onEvent fires even when console output is disabled; debug never streams", () => {
  const h = harness(0);
  const events = [];
  const log = createBulkLogger("t", {
    sink: h.sink,
    now: h.now,
    enabled: false, // console silenced
    verbose: true,
    onEvent: (event) => events.push(event),
  });
  log.info("start", { rows: 3 });
  log.debug("row created", { row: 1 }); // must NOT stream
  assert.equal(h.lines.info.length, 0, "console stays silent");
  assert.equal(events.length, 1, "only the info event streamed");
  assert.equal(events[0].event, "start");
});

test("a throwing onEvent consumer never breaks logging", () => {
  const h = harness(0);
  const log = createBulkLogger("t", {
    sink: h.sink,
    now: h.now,
    enabled: true,
    onEvent: () => {
      throw new Error("stream closed");
    },
  });
  assert.doesNotThrow(() => log.info("start", { rows: 1 }));
  // Console line still written despite the consumer throwing.
  assert.equal(h.lines.info[0], "[bulk:t] +0ms start rows=1");
});

test("prefix option changes the [prefix:scope] tag; defaults to bulk", () => {
  const docs = harness(0);
  createBulkLogger("upload", {
    sink: docs.sink,
    now: docs.now,
    enabled: true,
    prefix: "docs",
  }).info("start", { bytes: 10 });
  assert.equal(docs.lines.info[0], "[docs:upload] +0ms start bytes=10");

  const bulk = harness(0);
  createBulkLogger("commit", {
    sink: bulk.sink,
    now: bulk.now,
    enabled: true,
  }).info("x");
  assert.equal(bulk.lines.info[0], "[bulk:commit] +0ms x");
});

test("envVar option lets a separate flag toggle output", () => {
  const prev = process.env.DOCUMENT_UPLOAD_LOGS;
  try {
    process.env.DOCUMENT_UPLOAD_LOGS = "off";
    const off = harness(0);
    createBulkLogger("upload", {
      sink: off.sink,
      now: off.now,
      prefix: "docs",
      envVar: "DOCUMENT_UPLOAD_LOGS",
    }).info("x");
    assert.equal(off.lines.info.length, 0, "custom env=off should suppress");

    process.env.DOCUMENT_UPLOAD_LOGS = "verbose";
    const verbose = harness(0);
    createBulkLogger("upload", {
      sink: verbose.sink,
      now: verbose.now,
      prefix: "docs",
      envVar: "DOCUMENT_UPLOAD_LOGS",
    }).debug("d");
    assert.equal(
      verbose.lines.info.length,
      1,
      "custom env=verbose should emit debug",
    );
  } finally {
    if (prev === undefined) delete process.env.DOCUMENT_UPLOAD_LOGS;
    else process.env.DOCUMENT_UPLOAD_LOGS = prev;
  }
});

test("BULK_UPLOAD_LOGS env drives the defaults (off / verbose / unset)", () => {
  const prev = process.env.BULK_UPLOAD_LOGS;
  try {
    process.env.BULK_UPLOAD_LOGS = "off";
    const off = harness(0);
    createBulkLogger("t", { sink: off.sink, now: off.now }).info("x");
    assert.equal(off.lines.info.length, 0, "off should suppress");

    process.env.BULK_UPLOAD_LOGS = "verbose";
    const verbose = harness(0);
    createBulkLogger("t", { sink: verbose.sink, now: verbose.now }).debug("x");
    assert.equal(verbose.lines.info.length, 1, "verbose should emit debug");

    delete process.env.BULK_UPLOAD_LOGS;
    const on = harness(0);
    const log = createBulkLogger("t", { sink: on.sink, now: on.now });
    log.info("x");
    log.debug("y");
    assert.equal(on.lines.info.length, 1, "unset defaults to enabled, no debug");
  } finally {
    if (prev === undefined) delete process.env.BULK_UPLOAD_LOGS;
    else process.env.BULK_UPLOAD_LOGS = prev;
  }
});
