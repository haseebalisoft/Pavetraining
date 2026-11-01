/**
 * Training Matrix editor: categories + expiry dates only, and a save only
 * marks cells Manual when their date really changed.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-matrix-editor.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const { MATRIX_EDITOR_DATE_HEADERS, matrixEditorGroup, matrixEditorLabel } =
  await import("@/lib/training/matrixEditorFields");
const { changedMatrixHeaders, mergeManualOverrides } = await import(
  "@/lib/training/matrixManualOverrides"
);

test("editor lists every training category exactly once, no Name/DOB", () => {
  assert.ok(!MATRIX_EDITOR_DATE_HEADERS.includes("Name"));
  assert.ok(!MATRIX_EDITOR_DATE_HEADERS.includes("DOB"));
  assert.equal(new Set(MATRIX_EDITOR_DATE_HEADERS).size, MATRIX_EDITOR_DATE_HEADERS.length);
  assert.ok(MATRIX_EDITOR_DATE_HEADERS.includes("N217 - Road Sweeper"));
});

test("categories are grouped and labelled for the editor", () => {
  assert.equal(matrixEditorGroup("CSCS Expiry"), "Cards & site safety");
  assert.equal(matrixEditorLabel("CSCS Expiry"), "CSCS expiry");
  assert.equal(matrixEditorLabel("Face ift"), "Face Fit expiry");
  assert.equal(matrixEditorGroup("EUSR - SHEA Gas"), "EUSR categories");
  assert.equal(matrixEditorGroup("N001 - Ind FLT"), "NPORS categories");
  assert.equal(matrixEditorLabel("N001 - Ind FLT"), "N001 - Ind FLT");
  const groups = new Set(MATRIX_EDITOR_DATE_HEADERS.map(matrixEditorGroup));
  assert.deepEqual([...groups].sort(), ["Cards & site safety", "EUSR categories", "NPORS categories"]);
});

test("saving unchanged dates marks nothing Manual (register sync keeps working)", () => {
  const before = { "CSCS Expiry": "2027-01-27", "N001 - Ind FLT": "2026-10-27T00:00:00Z", "N003 - Reach Lift Truck": null };
  const after = { "CSCS Expiry": "2027-01-27", "N001 - Ind FLT": "2026-10-27", "N003 - Reach Lift Truck": null };
  assert.deepEqual(changedMatrixHeaders(Object.keys(after), before, after), []);
});

test("only the dates the admin changed become Manual", () => {
  const before = { "CSCS Expiry": "2027-01-27", "N001 - Ind FLT": "2026-10-27", "N003 - Reach Lift Truck": null };
  const after = { "CSCS Expiry": "2028-01-27", "N001 - Ind FLT": "2026-10-27", "N003 - Reach Lift Truck": "2029-03-01" };
  const changed = changedMatrixHeaders(Object.keys(after), before, after);
  assert.deepEqual(changed, ["CSCS Expiry", "N003 - Reach Lift Truck"]);
  // Existing Manual cells are kept, new ones added once.
  assert.deepEqual(
    mergeManualOverrides(["SSSTS Expiry"], changed).sort(),
    ["CSCS Expiry", "N003 - Reach Lift Truck", "SSSTS Expiry"],
  );
});

test("clearing a date is a change", () => {
  const changed = changedMatrixHeaders(["CSCS Expiry"], { "CSCS Expiry": "2027-01-27" }, { "CSCS Expiry": null });
  assert.deepEqual(changed, ["CSCS Expiry"]);
});
