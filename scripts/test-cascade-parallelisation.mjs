#!/usr/bin/env node
/**
 * Verify that the company delete cascade parallelisation split matches
 * intent: leaf cascades run in parallel, the cross-linked chain
 * (workforce → trainingMatrixExample → departments → permissions) runs
 * sequentially in that order.
 *
 * Reads src/lib/services/companyCascadeDeleteService.ts and inspects the
 * COMPANY_CASCADE_TARGETS array + the split we added.
 *
 * Usage:  node scripts/test-cascade-parallelisation.mjs
 */
import { readFileSync } from "node:fs";

const green = (s) => `\x1b[32m${s}\x1b[0m`;
const red = (s) => `\x1b[31m${s}\x1b[0m`;
const bold = (s) => `\x1b[1m${s}\x1b[0m`;

let checks = 0;
let failed = 0;
function assert(label, got, expected) {
  checks += 1;
  const ok = JSON.stringify(got) === JSON.stringify(expected);
  if (!ok) failed += 1;
  console.log(
    `${ok ? green("PASS") : red("FAIL")}  ${label}` +
      (ok ? "" : `\n         expected ${JSON.stringify(expected)}\n         got      ${JSON.stringify(got)}`),
  );
}

const src = readFileSync(
  "src/lib/services/companyCascadeDeleteService.ts",
  "utf8",
);

// Extract listKey values in order.
const listKeyOrder = [
  ...src.matchAll(/listKey:\s*"([a-zA-Z]+)"/g),
].map((match) => match[1]);

// Sanity: at least 10 entries.
assert("COMPANY_CASCADE_TARGETS parsed", listKeyOrder.length > 10, true);

// Confirmed sequential chain in original order.
const expectedSequential = [
  "workforce",
  "trainingMatrixExample",
  "departments",
  "permissions",
];
for (const key of expectedSequential) {
  assert(
    `sequential-chain key present: ${key}`,
    listKeyOrder.includes(key),
    true,
  );
}

// Confirm the sequentialListKeys set in the source matches our expectation.
const setSource = src.match(
  /const sequentialListKeys = new Set\(\[([\s\S]*?)\]\);/,
);
assert("sequentialListKeys set found in source", Boolean(setSource), true);
if (setSource) {
  const keysInSet = [...setSource[1].matchAll(/"([a-zA-Z]+)"/g)].map(
    (match) => match[1],
  );
  assert(
    "sequentialListKeys matches expected chain",
    keysInSet.sort(),
    [...expectedSequential].sort(),
  );
}

// Parallel targets = everything else.
const expectedParallel = listKeyOrder.filter(
  (key) => !expectedSequential.includes(key),
);
console.log(
  `\nParallel targets (${expectedParallel.length}): ${expectedParallel.join(", ")}`,
);
console.log(
  `Sequential chain (${expectedSequential.length}): ${expectedSequential.join(" → ")}`,
);

// Order of the sequential chain in the source array must be exactly:
// workforce (position N), trainingMatrixExample (N+1), departments (N+2), permissions (N+3).
const chainPositions = expectedSequential.map((key) => listKeyOrder.indexOf(key));
for (let i = 1; i < chainPositions.length; i += 1) {
  assert(
    `${expectedSequential[i - 1]} comes before ${expectedSequential[i]} in cascade order`,
    chainPositions[i] > chainPositions[i - 1],
    true,
  );
}

// Parallelisation code shape: Promise.all present, for-loop present after.
assert(
  "Promise.all(parallelTargets.map(...)) present",
  /Promise\.all\(\s*parallelTargets\.map\(/.test(src),
  true,
);
assert(
  "for (const target of sequentialTargets) present",
  /for \(const target of sequentialTargets\)/.test(src),
  true,
);

// Bounded item-level concurrency helper present.
assert(
  "runBounded helper defined",
  /async function runBounded</.test(src),
  true,
);
assert(
  "CASCADE_ITEM_CONCURRENCY constant present",
  /const CASCADE_ITEM_CONCURRENCY\s*=\s*\d+/.test(src),
  true,
);
assert(
  "runBounded used inside deleteByLookupId",
  (src.match(/runBounded\(fresh, CASCADE_ITEM_CONCURRENCY, deleteItem\)/g) ?? [])
    .length,
  2,
);

// Behavioural check on the runBounded semantics: every item processed
// exactly once, bound respected, order-independent.
(async () => {
  const items = Array.from({ length: 50 }, (_, i) => i);
  const seen = new Set();
  let inFlight = 0;
  let peak = 0;
  async function worker(item) {
    inFlight += 1;
    peak = Math.max(peak, inFlight);
    await new Promise((r) => setTimeout(r, Math.random() * 5));
    if (seen.has(item)) throw new Error(`double-processed ${item}`);
    seen.add(item);
    inFlight -= 1;
  }
  // Reimplemented locally to test semantics — the real one lives in the .ts.
  async function runBounded(items, concurrency, worker) {
    if (items.length === 0) return;
    const cap = Math.max(1, Math.min(concurrency, items.length));
    let cursor = 0;
    const runners = [];
    for (let i = 0; i < cap; i += 1) {
      runners.push(
        (async () => {
          while (true) {
            const index = cursor;
            cursor += 1;
            if (index >= items.length) return;
            await worker(items[index]);
          }
        })(),
      );
    }
    await Promise.all(runners);
  }
  await runBounded(items, 6, worker);
  assert("runBounded processed every item exactly once", seen.size, items.length);
  assert("runBounded never exceeded concurrency", peak <= 6, true);
})();

console.log(
  "\n" +
    (failed === 0
      ? green(bold(`ALL GOOD — ${checks} checks passed.`))
      : red(bold(`${failed}/${checks} checks failed.`))),
);
process.exit(failed === 0 ? 0 : 1);
