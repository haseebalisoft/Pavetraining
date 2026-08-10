/**
 * Unit tests for the shared email validation (src/lib/validation/email.ts).
 *
 * Runs with the built-in Node test runner + native TypeScript type-stripping
 * (Node >= 22); the module is dependency-free so it loads with no "@/"
 * resolver and no "server-only":
 *
 *   node --test scripts/test-email-validation.mjs
 *
 * Covers the requirements for "fix email validation globally": accept all
 * valid domains/TLDs (.org, .co.uk, .training, long TLDs, subdomains), trim +
 * lowercase, blank optional emails never block save, and clear rejection of
 * malformed values.
 */
import test from "node:test";
import assert from "node:assert/strict";

const BASE = new URL("../src/lib/", import.meta.url).pathname;
const { normalizeEmail, isValidEmail, validateEmailField } = await import(
  BASE + "validation/email.ts"
);

// --- Required cases from the brief -----------------------------------------

test("accepts test@example.org", () => {
  assert.equal(isValidEmail("test@example.org"), true);
  assert.deepEqual(validateEmailField("test@example.org"), {
    ok: true,
    email: "test@example.org",
  });
});

test("accepts test@example.co.uk", () => {
  assert.equal(isValidEmail("test@example.co.uk"), true);
  assert.deepEqual(validateEmailField("test@example.co.uk"), {
    ok: true,
    email: "test@example.co.uk",
  });
});

test("accepts test@subdomain.company.training", () => {
  assert.equal(isValidEmail("test@subdomain.company.training"), true);
  assert.deepEqual(validateEmailField("test@subdomain.company.training"), {
    ok: true,
    email: "test@subdomain.company.training",
  });
});

test("blank optional email is accepted (never blocks save) → null", () => {
  for (const blank of ["", "   ", null, undefined]) {
    assert.deepEqual(
      validateEmailField(blank),
      { ok: true, email: null },
      `blank ${JSON.stringify(blank)} should pass as null`,
    );
    assert.deepEqual(validateEmailField(blank, { required: false }), {
      ok: true,
      email: null,
    });
  }
});

test("invalid email is rejected clearly", () => {
  for (const bad of [
    "not-an-email",
    "missing@domain",
    "@no-local.com",
    "no-at-sign.com",
    "spaces in@email.com",
    "two@@at.com",
  ]) {
    const result = validateEmailField(bad, { label: "Email" });
    assert.equal(result.ok, false, `${bad} should be rejected`);
    assert.equal(result.error, "Email must be a valid email address.");
    assert.equal(isValidEmail(bad), false);
  }
});

// --- Do not restrict to .com / known domains -------------------------------

test("accepts arbitrary and long TLDs, not just .com", () => {
  for (const email of [
    "a@b.io",
    "person@charity.org",
    "team@startup.technology",
    "hello@agency.international",
    "user@host.museum",
    "name@sub.deep.example.co.uk",
  ]) {
    assert.equal(isValidEmail(email), true, `${email} should be valid`);
  }
});

// --- Trim + lowercase ------------------------------------------------------

test("normalizeEmail trims surrounding whitespace and lowercases", () => {
  assert.equal(normalizeEmail("  Test@Example.ORG  "), "test@example.org");
  assert.equal(normalizeEmail("USER@DOMAIN.CO.UK"), "user@domain.co.uk");
  assert.equal(normalizeEmail(null), "");
  assert.equal(normalizeEmail(undefined), "");
});

test("validateEmailField returns the trimmed + lowercased value", () => {
  assert.deepEqual(validateEmailField("  Admin@Company.Training "), {
    ok: true,
    email: "admin@company.training",
  });
});

// --- Required vs optional --------------------------------------------------

test("required + blank is rejected with a clear message", () => {
  const result = validateEmailField("", {
    required: true,
    label: "User email",
  });
  assert.equal(result.ok, false);
  assert.equal(result.error, "User email is required.");
});

test("required + valid returns the normalized email", () => {
  assert.deepEqual(
    validateEmailField(" Owner@Firm.CO.UK ", { required: true }),
    { ok: true, email: "owner@firm.co.uk" },
  );
});

test("optional but non-blank still validates format", () => {
  const result = validateEmailField("garbage", { required: false });
  assert.equal(result.ok, false);
  assert.equal(result.error, "Email must be a valid email address.");
});

test("custom label is used in error messages", () => {
  const result = validateEmailField("nope", { label: "Accounts email" });
  assert.equal(result.ok, false);
  assert.equal(result.error, "Accounts email must be a valid email address.");
});
