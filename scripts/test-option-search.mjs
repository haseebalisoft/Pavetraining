/**
 * Searchable NPORS category selector: name / code matching and ranking.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-option-search.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const { searchOptions } = await import("@/lib/utils/optionSearch");
const { getNporsCategoryOptions } = await import("@/lib/training/nporsCategoryOptions");

const options = [
  { value: "N001", label: "N001 - Ind FLT" },
  { value: "N017", label: "N017 - Abrasive Wheels – Cut-off Saw" },
  { value: "N201", label: "N201 - Excavator 180°" },
  { value: "N202", label: "N202 - Excavator 360°" },
  { value: "N214", label: "N214 - Road Roller" },
  { value: "N220", label: "N220 - Paver" },
  { value: "N702A", label: "N702A - Confined Spaces – Low Risk" },
  { value: "N703", label: "N703 - Fire Warden" },
];
const codes = (query) => searchOptions(options, query).map((o) => o.value);

test("empty search shows every option in its original order", () => {
  assert.deepEqual(codes(""), options.map((o) => o.value));
  assert.deepEqual(codes("   "), options.map((o) => o.value));
});

test("codes match however they are typed", () => {
  for (const q of ["N220", "n220", "220", "N 220"]) assert.equal(codes(q)[0], "N220", q);
  assert.equal(codes("17")[0], "N017");
  assert.equal(codes("n1")[0], "N001");
  assert.equal(codes("702a")[0], "N702A");
});

test("partial codes list the family, in code order", () => {
  assert.deepEqual(codes("N70"), ["N702A", "N703"]);
  assert.deepEqual(codes("N20"), ["N201", "N202"]);
});

test("names match case-insensitively, any word order", () => {
  assert.deepEqual(codes("paver"), ["N220"]);
  assert.deepEqual(codes("excavator"), ["N201", "N202"]);
  assert.deepEqual(codes("360 excav"), ["N202"]);
  assert.deepEqual(codes("ROLLER"), ["N214"]);
});

test("no match returns nothing (never a guess)", () => {
  assert.deepEqual(codes("helicopter"), []);
});

test("does not change option values (stored data unchanged)", () => {
  const before = JSON.stringify(options);
  searchOptions(options, "paver");
  assert.equal(JSON.stringify(options), before);
  assert.equal(searchOptions(options, "paver")[0], options[5]);
});

test("works on the real built-in NPORS list", () => {
  const all = getNporsCategoryOptions();
  assert.ok(all.length > 50);
  const [top] = searchOptions(all, "220");
  assert.equal(top.value, "N220");
  assert.ok(searchOptions(all, "paver").some((o) => o.value === "N220"));
});
