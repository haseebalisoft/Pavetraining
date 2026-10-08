/**
 * Offline checks for Company Logo bulk-import planning (no Graph).
 *   node scripts/test-company-logo-bulk.mjs
 */
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { resolve } from "node:path";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const PLACEHOLDER_LOGO =
  /^(logo(\s+placeholder)?|placeholder|n\/?a|none|null|-|—|–)$/i;

function planCompanyLogoImport(value) {
  const raw = (value ?? "").trim();
  if (!raw || PLACEHOLDER_LOGO.test(raw)) {
    return {
      kind: "missing",
      warning:
        "Company Logo missing — company will import without a logo. Put an https image URL in Company Logo to upload one.",
    };
  }
  let parsed;
  try {
    parsed = new URL(raw);
  } catch {
    return {
      kind: "invalid",
      warning: `Company Logo "${raw}" is not a valid image URL — company will import without a logo.`,
    };
  }
  if (parsed.protocol !== "https:" && parsed.protocol !== "http:") {
    return {
      kind: "invalid",
      warning: `Company Logo must be an http(s) image URL (got ${parsed.protocol}) — company will import without a logo.`,
    };
  }
  const host = parsed.hostname.toLowerCase();
  if (
    host === "localhost" ||
    host === "127.0.0.1" ||
    host === "::1" ||
    host.endsWith(".local") ||
    host.endsWith(".internal")
  ) {
    return {
      kind: "invalid",
      warning:
        "Company Logo URL points to a local/internal host — company will import without a logo.",
    };
  }
  return { kind: "url", url: parsed.toString(), fileName: "logo.png" };
}

assert.equal(planCompanyLogoImport("").kind, "missing");
assert.equal(planCompanyLogoImport("Logo placeholder").kind, "missing");
assert.equal(planCompanyLogoImport("n/a").kind, "missing");
assert.equal(planCompanyLogoImport("not a url").kind, "invalid");
assert.equal(planCompanyLogoImport("ftp://x/a.png").kind, "invalid");
assert.equal(planCompanyLogoImport("http://localhost/a.png").kind, "invalid");
assert.equal(
  planCompanyLogoImport("https://cdn.example.com/brand/logo.png").kind,
  "url",
);

const template = resolve(
  process.cwd(),
  "public/bulk-templates/Company-list-template.xlsx",
);
const wb = XLSX.readFile(template);
const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
  defval: "",
});
assert.ok(Object.keys(rows[0] ?? {}).includes("Company Logo"));
const plans = rows.map((r) => planCompanyLogoImport(r["Company Logo"]));
const urls = plans.filter((p) => p.kind === "url").length;
const missing = plans.filter((p) => p.kind === "missing").length;
assert.ok(urls >= 3, `expected >=3 logo URL samples, got ${urls}`);
assert.ok(missing >= 1, `expected >=1 missing logo sample, got ${missing}`);

console.log("OK: Company Logo bulk plan + template samples");
console.log(`  template URL samples: ${urls}, missing: ${missing}`);
