/**
 * Unit tests for the pure register → Training Matrix field sync rules.
 *
 * Runs with the built-in Node test runner + native TypeScript type-stripping
 * (Node >= 22); registerMatrixFieldSync.ts is dependency-free so it loads with
 * no Graph and no "@/" resolver hook:
 *
 *   node --test scripts/test-register-matrix-field-sync.mjs
 *
 * These cover the forward-apply rule (shouldApplyPassExpiry / latestPassExpiry)
 * and the delete/downgrade recompute engine (computeMatrixFieldAfterRemoval)
 * behind locked scenarios 1-3 of the register-save-sync fix:
 *   1. Create a Pass record with expiry → matrix should adopt it.
 *   2. Delete that record → matrix recomputes from any remaining source, or
 *      clears + Needs Review if none remain.
 *   3. A manual-override value survives a source delete and is flagged
 *      "Manual Override / Source Deleted" rather than blanked or recomputed.
 */
import test from "node:test";
import assert from "node:assert/strict";

const BASE = new URL("../src/lib/", import.meta.url).pathname;
const { shouldApplyPassExpiry, latestPassExpiry, computeMatrixFieldAfterRemoval } =
  await import(BASE + "services/bulkUpload/registerMatrixFieldSync.ts");

// --- shouldApplyPassExpiry: forward-apply guard (requirement: Fail never extends) ---

test("Fail outcome never applies, regardless of dates", () => {
  assert.equal(shouldApplyPassExpiry(null, "2027-01-01", "Fail"), false);
  assert.equal(shouldApplyPassExpiry("2020-01-01", "2027-01-01", "Fail"), false);
});

test("null/missing outcome never applies", () => {
  assert.equal(shouldApplyPassExpiry(null, "2027-01-01", null), false);
});

test("blank or whitespace-only incoming date never applies, even on Pass", () => {
  assert.equal(shouldApplyPassExpiry(null, "", "Pass"), false);
  assert.equal(shouldApplyPassExpiry(null, "   ", "Pass"), false);
  assert.equal(shouldApplyPassExpiry(null, null, "Pass"), false);
  assert.equal(shouldApplyPassExpiry(null, undefined, "Pass"), false);
});

test("unparsable incoming date never applies, even on Pass", () => {
  assert.equal(shouldApplyPassExpiry(null, "not-a-date", "Pass"), false);
});

test("Pass + no existing value + valid incoming → applies", () => {
  assert.equal(shouldApplyPassExpiry(null, "2027-01-01", "Pass"), true);
  assert.equal(shouldApplyPassExpiry("", "2027-01-01", "Pass"), true);
});

test("Pass + unparsable existing value is treated as empty → applies", () => {
  assert.equal(shouldApplyPassExpiry("garbage", "2027-01-01", "Pass"), true);
});

test("Pass + incoming strictly newer than existing → applies", () => {
  assert.equal(shouldApplyPassExpiry("2025-01-01", "2027-01-01", "Pass"), true);
});

test("Pass + incoming equal to existing → applies (boundary, >=)", () => {
  assert.equal(shouldApplyPassExpiry("2027-01-01", "2027-01-01", "Pass"), true);
});

test("Pass + incoming older than existing → never regresses forward-apply", () => {
  assert.equal(shouldApplyPassExpiry("2027-01-01", "2025-01-01", "Pass"), false);
});

// --- latestPassExpiry: picks the max Pass-backed date, ignores the rest ---

test("empty record list → null", () => {
  assert.equal(latestPassExpiry([]), null);
});

test("all-Fail records → null (Fail never contributes an expiry)", () => {
  assert.equal(
    latestPassExpiry([
      { trainingOutcome: "Fail", expiry: "2099-01-01" },
      { trainingOutcome: "Fail", expiry: "2098-01-01" },
    ]),
    null,
  );
});

test("all-blank expiries → null even when outcome is Pass", () => {
  assert.equal(
    latestPassExpiry([
      { trainingOutcome: "Pass", expiry: null },
      { trainingOutcome: "Pass", expiry: "" },
    ]),
    null,
  );
});

test("mixed Pass/Fail → only Pass records are considered", () => {
  const value = latestPassExpiry([
    { trainingOutcome: "Fail", expiry: "2099-01-01" },
    { trainingOutcome: "Pass", expiry: "2026-06-30" },
  ]);
  assert.equal(value, "2026-06-30");
});

test("multiple Pass records → returns the latest (max) date", () => {
  const value = latestPassExpiry([
    { trainingOutcome: "Pass", expiry: "2025-01-01" },
    { trainingOutcome: "Pass", expiry: "2028-03-15" },
    { trainingOutcome: "Pass", expiry: "2026-12-31" },
  ]);
  assert.equal(value, "2028-03-15");
});

test("unparsable Pass expiries are ignored; a valid one among them still wins", () => {
  const value = latestPassExpiry([
    { trainingOutcome: "Pass", expiry: "not-a-date" },
    { trainingOutcome: "Pass", expiry: "2026-12-31" },
  ]);
  assert.equal(value, "2026-12-31");
});

test("exact-tie dates → stable (first-seen) value returned", () => {
  const value = latestPassExpiry([
    { trainingOutcome: "Pass", expiry: "2026-12-31" },
    { trainingOutcome: "Pass", expiry: "2026-12-31" },
  ]);
  assert.equal(value, "2026-12-31");
});

// --- computeMatrixFieldAfterRemoval: the delete/downgrade recompute engine ---

// Scenario 2: delete the only supporting record → clear + Needs Review.
test("scenario 2: no remaining source and not a manual override → cleared + Needs Review", () => {
  const result = computeMatrixFieldAfterRemoval({
    header: "N001 - Plant Operations",
    currentValue: "2025-01-01",
    isManualOverride: false,
    remainingRecords: [],
  });
  assert.equal(result.action, "cleared");
  assert.equal(result.nextValue, null);
  assert.equal(result.forceNeedsReview, true);
  assert.match(result.note, /cleared/i);
  assert.match(result.note, /Needs Review/i);
});

// Scenario 2 (recompute branch): another valid source exists — even an OLDER
// one — recompute is authoritative, not forward-only.
test("scenario 2: another valid (older) source remains → recomputes to that older date", () => {
  const result = computeMatrixFieldAfterRemoval({
    header: "N001 - Plant Operations",
    currentValue: "2025-06-01", // the just-deleted record's expiry
    isManualOverride: false,
    remainingRecords: [{ trainingOutcome: "Pass", expiry: "2024-01-01" }], // older
  });
  assert.equal(result.action, "recomputed");
  assert.equal(result.nextValue, "2024-01-01");
  assert.equal(result.forceNeedsReview, false);
  assert.match(result.note, /recomputed/i);
});

test("recompute that lands back on the same value is reported as unchanged", () => {
  const result = computeMatrixFieldAfterRemoval({
    header: "N001 - Plant Operations",
    currentValue: "2026-12-31",
    isManualOverride: false,
    remainingRecords: [{ trainingOutcome: "Pass", expiry: "2026-12-31" }],
  });
  assert.equal(result.action, "unchanged");
  assert.equal(result.nextValue, "2026-12-31");
  assert.equal(result.forceNeedsReview, false);
});

test("Fail-only remaining records behave as if nothing remains (clears)", () => {
  const result = computeMatrixFieldAfterRemoval({
    header: "EUSR Expiry",
    currentValue: "2025-01-01",
    isManualOverride: false,
    remainingRecords: [{ trainingOutcome: "Fail", expiry: "2099-01-01" }],
  });
  assert.equal(result.action, "cleared");
  assert.equal(result.nextValue, null);
  assert.equal(result.forceNeedsReview, true);
});

test("nothing remains and current value already blank → unchanged, no-op", () => {
  const result = computeMatrixFieldAfterRemoval({
    header: "EUSR Expiry",
    currentValue: null,
    isManualOverride: false,
    remainingRecords: [],
  });
  assert.equal(result.action, "unchanged");
  assert.equal(result.nextValue, null);
  assert.equal(result.forceNeedsReview, false);
});

// Scenario 3: manual override always survives a source delete.
test("scenario 3: manual override with a value is preserved, never blanked", () => {
  const result = computeMatrixFieldAfterRemoval({
    header: "NRSWA Expiry",
    currentValue: "2030-01-01",
    isManualOverride: true,
    remainingRecords: [],
  });
  assert.equal(result.action, "preserved_manual_override");
  assert.equal(result.nextValue, "2030-01-01");
  assert.equal(result.forceNeedsReview, false);
  assert.match(result.note, /Manual Override \/ Source Deleted/);
});

test("manual override short-circuits even when a remaining source could recompute a different date", () => {
  // Proves overrides are never silently recomputed either — the admin's
  // manually-entered value wins regardless of what remains.
  const result = computeMatrixFieldAfterRemoval({
    header: "NRSWA Expiry",
    currentValue: "2030-01-01",
    isManualOverride: true,
    remainingRecords: [{ trainingOutcome: "Pass", expiry: "2099-12-31" }],
  });
  assert.equal(result.action, "preserved_manual_override");
  assert.equal(result.nextValue, "2030-01-01");
});

test("manual override flag with no existing value → unchanged (nothing to preserve)", () => {
  const result = computeMatrixFieldAfterRemoval({
    header: "NRSWA Expiry",
    currentValue: null,
    isManualOverride: true,
    remainingRecords: [],
  });
  assert.equal(result.action, "unchanged");
  assert.equal(result.nextValue, null);
  assert.equal(result.forceNeedsReview, false);
});
