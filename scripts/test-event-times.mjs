/**
 * Event start/end times: 15-minute choices, existing times preserved, end
 * strictly after start.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-event-times.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const {
  QUARTER_HOUR_TIMES,
  quarterHourOptions,
  splitDateTimeLocal,
  joinDateTimeLocal,
  roundUpToQuarterHour,
  endAfterStartError,
} = await import("@/lib/utils/quarterHour");

test("time choices are every 15 minutes across the day", () => {
  assert.equal(QUARTER_HOUR_TIMES.length, 96);
  assert.deepEqual(QUARTER_HOUR_TIMES.slice(36, 40), ["09:00", "09:15", "09:30", "09:45"]);
  assert.equal(QUARTER_HOUR_TIMES[0], "00:00");
  assert.equal(QUARTER_HOUR_TIMES.at(-1), "23:45");
});

test("an existing off-grid time stays selectable (data preserved)", () => {
  // Live events hold times like 08:39 and 17:06.
  const options = quarterHourOptions("08:39");
  assert.ok(options.includes("08:39"));
  assert.equal(options.length, 97);
  assert.equal(options.indexOf("08:39"), options.indexOf("08:30") + 1, "sorted in place");
  assert.equal(quarterHourOptions("09:15").length, 96);
  assert.equal(quarterHourOptions("").length, 96);
});

test("value format is unchanged (YYYY-MM-DDTHH:mm round-trips)", () => {
  assert.deepEqual(splitDateTimeLocal("2026-10-28T08:39"), { date: "2026-10-28", time: "08:39" });
  assert.equal(joinDateTimeLocal("2026-10-28", "08:39"), "2026-10-28T08:39");
  assert.equal(joinDateTimeLocal("2026-10-28", ""), "2026-10-28T09:00");
  assert.equal(joinDateTimeLocal("", "09:00"), "");
});

test("new events start on the next quarter hour", () => {
  const at = (h, m) => new Date(2026, 9, 31, h, m, 27);
  const fmt = (d) => `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
  assert.equal(fmt(roundUpToQuarterHour(at(9, 0))), "09:00");
  assert.equal(fmt(roundUpToQuarterHour(at(9, 1))), "09:15");
  assert.equal(fmt(roundUpToQuarterHour(at(9, 44))), "09:45");
  assert.equal(fmt(roundUpToQuarterHour(at(23, 50))), "00:00");
});

test("end must be strictly after start", () => {
  assert.equal(endAfterStartError("2026-10-31T09:00", "2026-10-31T09:15"), null);
  assert.match(endAfterStartError("2026-10-31T09:00", "2026-10-31T09:00"), /after the start/);
  assert.match(endAfterStartError("2026-10-31T10:00", "2026-10-31T09:45"), /after the start/);
  assert.equal(endAfterStartError("2026-10-31T09:00", ""), null, "blank end = nothing to compare");
});
