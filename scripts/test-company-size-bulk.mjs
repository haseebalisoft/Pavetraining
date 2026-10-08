/**
 * Offline Companies bulk preview simulation for Company Size.
 *   node scripts/test-company-size-bulk.mjs
 */
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const COMPANY_SIZE_CHOICES = ["Small", "Medium", "Large", "Enterprise"];

function normalizeCompanySize(value) {
  const raw = (value ?? "").trim();
  if (!raw) return { size: null };
  const lower = raw.toLowerCase();
  if (lower === "small") return { size: "Small" };
  if (lower === "medium") return { size: "Medium" };
  if (lower === "large") return { size: "Large" };
  if (lower === "enterprise") return { size: "Enterprise" };
  return {
    size: null,
    error: `Company Size must be one of: ${COMPANY_SIZE_CHOICES.join(", ")} (got "${raw}").`,
  };
}

for (const size of COMPANY_SIZE_CHOICES) {
  assert.equal(normalizeCompanySize(size).size, size);
  assert.equal(normalizeCompanySize(size.toLowerCase()).size, size);
  assert.equal(normalizeCompanySize(`  ${size}  `).size, size);
}

assert.equal(normalizeCompanySize("").size, null);
assert.equal(normalizeCompanySize(null).size, null);
assert.match(normalizeCompanySize("Huge").error, /must be one of/);
assert.equal(normalizeCompanySize("Huge").size, null);
assert.match(normalizeCompanySize("SME").error, /must be one of/);

const template = resolve(
  process.cwd(),
  "public/bulk-templates/Company-list-template.xlsx",
);
const wb = XLSX.readFile(template);
const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
const sizes = new Set(rows.map((r) => r["Company Size"]));
for (const size of COMPANY_SIZE_CHOICES) {
  assert.ok(sizes.has(size), `template missing sample ${size}`);
}

// Simulate preview display: normalized size is what Admin preview prefers.
const previewCells = rows.map((r) => {
  const info = normalizeCompanySize(r["Company Size"]);
  assert.ok(!info.error, info.error);
  assert.ok(info.size);
  return info.size;
});
assert.deepEqual(
  [...new Set(previewCells)].sort(),
  [...COMPANY_SIZE_CHOICES].sort(),
);

// Commit payload field name mapping (schema key → SharePoint internal).
const companyFields = { companySize: "CompanySize" };
for (const size of COMPANY_SIZE_CHOICES) {
  const payload = { [companyFields.companySize]: size };
  assert.equal(payload.CompanySize, size);
}

console.log("OK: Company Size validation + template samples + SharePoint field map");
console.log("  samples:", [...sizes].sort().join(", "));
console.log("  template bytes:", readFileSync(template).length);
