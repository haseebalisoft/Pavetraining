/**
 * NPORS blank expiry defaults to training date + 3 years.
 *
 *   node --import ./scripts/_register-e2e-hook.mjs --test scripts/test-npors-default-expiry.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const { addCalendarYearsIso } = await import(
  new URL("../src/lib/utils/formatDate.ts", import.meta.url)
);

test("NPORS blank expiry is training date + 3 calendar years", () => {
  assert.equal(addCalendarYearsIso("2026-03-15", 3), "2029-03-15");
  assert.equal(addCalendarYearsIso("2024-02-29", 3), "2027-03-01");
});

test("Admin NPORS form declares 3-year blank expiry default", () => {
  const source = readFileSync(
    resolve(ROOT, "src/components/admin/pages/AdminRegisterClient.tsx"),
    "utf8",
  );
  assert.match(
    source,
    /Expiry \(3 years from training date if left blank\)/,
  );
  assert.match(source, /defaultExpiryYears:\s*kind === "npors".*\? 3/);
});

test("Server fills blank NPORS expiry with +3 years on create and update", () => {
  const source = readFileSync(
    resolve(ROOT, "src/lib/services/adminCrudService.ts"),
    "utf8",
  );
  assert.match(source, /key === "nporsRegister" \|\| key === "eusrRegister"/);
  assert.match(source, /addCalendarYearsIso\(\s*optionalText\(input\.trainingDate\),\s*3,/);
  assert.match(source, /blankExpirySubmitted/);
});

test("UI auto-fills blank expiry when training date is present", () => {
  const source = readFileSync(
    resolve(ROOT, "src/components/admin/AdminCrudPage.tsx"),
    "utf8",
  );
  assert.match(source, /withAutoFilledBlankExpiry/);
  assert.match(source, /defaultExpiryYears/);
  assert.match(source, /field\.name === "trainingDate"/);
});
