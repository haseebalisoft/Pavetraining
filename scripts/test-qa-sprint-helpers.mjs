/**
 * Unit checks for email + UK date helpers (no SharePoint).
 *
 *   node scripts/test-qa-sprint-helpers.mjs
 */
import assert from "node:assert/strict";

import {
  isValidEmail,
  normalizeEmail,
  optionalEmail,
} from "../src/lib/utils/email.ts";
import { formatEventDuration } from "../src/lib/utils/eventDuration.ts";
import { formatDate } from "../src/lib/utils/formatDate.ts";
import { normalizeDateValue } from "../src/lib/utils/ukDate.ts";

function ok(label) {
  console.log(`  ✓ ${label}`);
}

assert.equal(isValidEmail("test@example.org"), true);
ok("test@example.org accepted");
assert.equal(isValidEmail("test@example.co.uk"), true);
ok("test@example.co.uk accepted");
assert.equal(isValidEmail("test@subdomain.company.training"), true);
ok("test@subdomain.company.training accepted");
assert.equal(optionalEmail(""), null);
assert.equal(optionalEmail("   "), null);
ok("blank optional email accepted as null");
assert.equal(isValidEmail("not-an-email"), false);
assert.equal(isValidEmail("missing-domain@"), false);
ok("invalid email rejected");
assert.equal(normalizeEmail("  Ada@Org.Example  "), "ada@org.example");
ok("normalize trims + lowercases");

assert.equal(normalizeDateValue("06/08/2026"), "2026-08-06");
ok("06/08/2026 → 6 August 2026");
assert.equal(normalizeDateValue("13/08/2026"), "2026-08-13");
ok("13/08/2026 valid UK");
assert.equal(normalizeDateValue("08/13/2026"), null);
ok("08/13/2026 rejected as invalid UK month");
assert.equal(normalizeDateValue("6/8/2026"), "2026-08-06");
ok("d/M/yyyy accepted");
assert.equal(normalizeDateValue("44942"), "2023-01-16");
ok("Excel serial imports");
assert.equal(formatDate("2026-08-06"), "06/08/2026");
ok("display uses dd/MM/yyyy");

assert.equal(
  formatEventDuration("2026-08-06T09:00:00", "2026-08-06T16:00:00"),
  "7 hours",
);
ok("same-day 09:00–16:00 = 7 hours");
assert.equal(
  formatEventDuration("2026-08-06T09:00:00", "2026-08-07T16:00:00"),
  "2 days",
);
ok("multi-day = 2 days");
assert.equal(formatEventDuration("2026-08-06T09:00:00", null), null);
ok("missing end → null");

console.log("\nAll QA helper checks passed.");
