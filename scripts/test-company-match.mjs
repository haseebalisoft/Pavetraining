/**
 * Unit tests for WORKFORCE bulk-import company matching.
 *
 * Runs with the built-in Node test runner + native TypeScript type-stripping
 * (Node >= 22). companyMatch.ts uses the project's "@/" alias for the shared
 * normalization helpers, so we register a tiny resolver hook first:
 *
 *   node --import ./scripts/_register-alias.mjs --test scripts/test-company-match.mjs
 *
 * The client rule under test: match a workforce row to its company by Company
 * Number FIRST, fall back to Company Name only when it uniquely identifies one
 * company, never assign by fuzzy name when a number is present, and never
 * silently create a duplicate company from a spelling variation. The resolver
 * is exercised exactly as candidateImporter wires it (same helper, same shape).
 */
import test from "node:test";
import assert from "node:assert/strict";

const BASE = new URL("../src/lib/", import.meta.url).pathname;
const { resolveWorkforceCompanyMatch, AMBIGUOUS_COMPANY_NAME_ERROR } =
  await import(BASE + "services/bulkUpload/companyMatch.ts");

/** Company List fixture (id === SharePoint Company lookup id). */
const COMPANIES = [
  { id: "10", companyName: "Acme Construction Ltd", companyNumber: "C00001" },
  { id: "20", companyName: "Beta Rail Limited", companyNumber: "C00002" },
  // Two DIFFERENT companies that share the same display name — only a Company
  // Number can disambiguate them.
  { id: "30", companyName: "Metro Services", companyNumber: "C00003" },
  { id: "31", companyName: "Metro Services", companyNumber: "C00004" },
  // A company with no number yet (never used to override a supplied number).
  { id: "40", companyName: "Gamma Utilities", companyNumber: null },
];

const NO_AUTO = { autoCreateMissing: false };
const AUTO = { autoCreateMissing: true };

// 1. Company Number matches the correct company.
test("Company Number matches the correct company (by number, correct lookup id)", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: "C00002", companyName: "Beta Rail Limited" },
    NO_AUTO,
  );
  assert.equal(result.kind, "matched");
  assert.equal(result.matchedBy, "companyNumber");
  assert.equal(result.company.id, "20");
  assert.equal(result.report.matchedCompanyId, "20");
  assert.equal(result.report.matchedCompanyNumber, "C00002");
  assert.equal(result.report.matchedCompanyName, "Beta Rail Limited");
  assert.equal(result.report.error, null);
});

// 2. Wrong / unknown Company Number errors (never falls back to name).
test("Unknown Company Number errors and does NOT fall back to name", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    // The name would match id 10, but a wrong number must not silently match it.
    { companyNumber: "C99999", companyName: "Acme Construction Ltd" },
    NO_AUTO,
  );
  assert.equal(result.kind, "error");
  assert.equal(result.company, undefined);
  assert.equal(result.report.matchedCompanyId, null);
  assert.match(result.report.error, /C99999/);
});

// 3. Same company name across multiple records errors without a Company Number.
test("Ambiguous name (shared across companies) errors without a Company Number", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: null, companyName: "Metro Services" },
    NO_AUTO,
  );
  assert.equal(result.kind, "error");
  assert.equal(result.report.matchedCompanyId, null);
  assert.equal(result.report.error, AMBIGUOUS_COMPANY_NAME_ERROR);
});

// 3b. ...but the SAME ambiguous name resolves cleanly when a number is given.
test("Ambiguous name is resolved when the Company Number is provided", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: "C00004", companyName: "Metro Services" },
    NO_AUTO,
  );
  assert.equal(result.kind, "matched");
  assert.equal(result.matchedBy, "companyNumber");
  assert.equal(result.company.id, "31");
});

// 4. Company Name fallback works only when it uniquely identifies one company.
test("Unique Company Name matches (fallback, correct lookup id)", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: null, companyName: "Beta Rail Limited" },
    NO_AUTO,
  );
  assert.equal(result.kind, "matched");
  assert.equal(result.matchedBy, "companyName");
  assert.equal(result.company.id, "20");
  assert.equal(result.report.matchedCompanyId, "20");
});

// 4b. Name fallback tolerates normalized variations (Ltd/Limited/punctuation).
test("Unique Company Name matches via normalization (Ltd vs Limited)", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: null, companyName: "acme construction limited" },
    NO_AUTO,
  );
  assert.equal(result.kind, "matched");
  assert.equal(result.matchedBy, "companyName");
  assert.equal(result.company.id, "10");
});

// 5. No duplicate company is created accidentally.
test("Unknown name does NOT create a company when auto-create is OFF", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: null, companyName: "Totally New Co" },
    NO_AUTO,
  );
  assert.equal(result.kind, "error");
  assert.notEqual(result.kind, "create");
  assert.match(result.report.error, /not been found|not found|not been|not found/i);
});

test("A resolvable name/number never yields a create action (no accidental dupes)", () => {
  // Both a matching number and a matching unique name must resolve to "matched",
  // so the committer reuses the existing company rather than creating one.
  for (const input of [
    { companyNumber: "C00001", companyName: "Acme Construction Ltd" },
    { companyNumber: null, companyName: "Gamma Utilities" },
  ]) {
    const result = resolveWorkforceCompanyMatch(COMPANIES, input, AUTO);
    assert.equal(result.kind, "matched", JSON.stringify(input));
  }
});

// 6. Workforce row links to the correct Company lookup id (spelling variation
//    of an ambiguous set: the NUMBER wins, name is ignored for the link).
test("Number wins over a differing name and links the correct lookup id", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    // Number points at Acme (id 10); the row's name is a variant/typo.
    { companyNumber: "C00001", companyName: "ACME (misspelled) Constructon" },
    NO_AUTO,
  );
  assert.equal(result.kind, "matched");
  assert.equal(result.matchedBy, "companyNumber");
  assert.equal(result.company.id, "10");
  // The mismatch is surfaced as a warning, not a silent re-link.
  assert.ok(result.report.warning);
  assert.match(result.report.warning, /Company Number/i);
});

// 7. Auto-create: unknown number → create (caller will create the company).
test("Auto-create ON: unknown Company Number yields a create action", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: "C40000", companyName: "New Ventures Ltd" },
    AUTO,
  );
  assert.equal(result.kind, "create");
  assert.equal(result.report.matchedCompanyNumber, "C40000");
  assert.equal(result.report.matchedCompanyName, "New Ventures Ltd");
  assert.equal(result.report.warning, null);
  assert.equal(result.report.error, null);
});

// 7b. Auto-create: unknown name → create.
test("Auto-create ON: unknown Company Name yields a create action", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: null, companyName: "Brand New Co" },
    AUTO,
  );
  assert.equal(result.kind, "create");
  assert.equal(result.report.matchedCompanyName, "Brand New Co");
});

// 7c. Auto-create still refuses an ambiguous name (data-integrity guard).
test("Auto-create ON: ambiguous name still errors (never creates a dupe)", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: null, companyName: "Metro Services" },
    AUTO,
  );
  assert.equal(result.kind, "error");
  assert.equal(result.report.error, AMBIGUOUS_COMPANY_NAME_ERROR);
});

// 8. Neither field provided → error.
test("Missing both Company Number and Company Name errors", () => {
  const result = resolveWorkforceCompanyMatch(
    COMPANIES,
    { companyNumber: "  ", companyName: "" },
    AUTO,
  );
  assert.equal(result.kind, "error");
  assert.match(result.report.error, /required/i);
});

// 8b. A number that is duplicated across companies (dirty data) is reported,
//     not silently matched to the first.
test("Company Number that matches multiple companies errors (dirty data)", () => {
  const dirty = [
    { id: "50", companyName: "Dup A", companyNumber: "C05000" },
    { id: "51", companyName: "Dup B", companyNumber: "C05000" },
  ];
  const result = resolveWorkforceCompanyMatch(
    dirty,
    { companyNumber: "C05000", companyName: "Dup A" },
    NO_AUTO,
  );
  assert.equal(result.kind, "error");
  assert.match(result.report.error, /more than one/i);
});
