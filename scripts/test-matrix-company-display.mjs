/**
 * Offline unit check: Admin Matrix company display resolution order.
 * Linked rows follow Workforce; unlinked use Company List keys then stored text.
 *   node scripts/test-matrix-company-display.mjs
 */
import assert from "node:assert/strict";

function resolveMatrixCompany(example, wf, companyById, companyByNumber, filterName) {
  if (wf) {
    const wfCompanyId = String(wf.companyId ?? "").trim();
    if (wfCompanyId) {
      const byId = companyById.get(wfCompanyId);
      if (byId) {
        return {
          companyName: byId.companyName,
          companyNumber: byId.companyNumber ?? wf.companyNumber ?? null,
        };
      }
    }
    const wfNumberKey = String(wf.companyNumber ?? "").trim().toLowerCase();
    if (wfNumberKey) {
      const byNumber = companyByNumber.get(wfNumberKey);
      if (byNumber) {
        return {
          companyName: byNumber.companyName,
          companyNumber: byNumber.companyNumber ?? wf.companyNumber,
        };
      }
    }
    return {
      companyName: wf.companyName?.trim() || null,
      companyNumber: wf.companyNumber?.trim() || null,
    };
  }

  const itemId = String(example.companyItemId ?? "").trim();
  if (itemId) {
    const byId = companyById.get(itemId);
    if (byId) {
      return {
        companyName: byId.companyName,
        companyNumber: byId.companyNumber ?? example.companyNumber ?? null,
      };
    }
  }
  const numberKey = String(example.companyNumber ?? "").trim().toLowerCase();
  if (numberKey) {
    const byNumber = companyByNumber.get(numberKey);
    if (byNumber) {
      return {
        companyName: byNumber.companyName,
        companyNumber: byNumber.companyNumber ?? example.companyNumber,
      };
    }
  }
  const stored = example.storedCompanyName?.trim() || null;
  if (stored) {
    return {
      companyName: stored,
      companyNumber: example.companyNumber?.trim() || null,
    };
  }
  return {
    companyName: filterName?.trim() || null,
    companyNumber: example.companyNumber?.trim() || null,
  };
}

const companies = [
  { id: "10", companyName: "Alpha Civils", companyNumber: "C00001" },
  { id: "20", companyName: "Beta Rail", companyNumber: "C00002" },
  { id: "30", companyName: "Gamma Gas", companyNumber: "C00003" },
];
const companyById = new Map(companies.map((c) => [c.id, c]));
const companyByNumber = new Map(
  companies.map((c) => [c.companyNumber.toLowerCase(), c]),
);

// Linked Workforce wins over matrix CompanyItemId / stored text.
const a = resolveMatrixCompany(
  { companyItemId: "10", companyNumber: "C00001", storedCompanyName: "Stale" },
  {
    companyId: "20",
    companyName: "Beta Rail",
    companyNumber: "C00002",
  },
  companyById,
  companyByNumber,
  null,
);
assert.equal(a.companyName, "Beta Rail");
assert.equal(a.companyNumber, "C00002");

// Unlinked: CompanyItemId
const b = resolveMatrixCompany(
  { companyItemId: "10", companyNumber: "C00001", storedCompanyName: null },
  null,
  companyById,
  companyByNumber,
  null,
);
assert.equal(b.companyName, "Alpha Civils");
assert.equal(b.companyNumber, "C00001");

// Unlinked: CompanyNumber
const c = resolveMatrixCompany(
  {
    companyItemId: null,
    companyNumber: "C00002",
    storedCompanyName: "Old Name",
  },
  null,
  companyById,
  companyByNumber,
  null,
);
assert.equal(c.companyName, "Beta Rail");
assert.equal(c.companyNumber, "C00002");

// Unlinked: stored name
const d = resolveMatrixCompany(
  {
    companyItemId: null,
    companyNumber: null,
    storedCompanyName: "Stored Unlinked Co",
  },
  null,
  companyById,
  companyByNumber,
  null,
);
assert.equal(d.companyName, "Stored Unlinked Co");

console.log(
  "OK: Matrix company display prefers Workforce when linked, else ItemId → Number → stored",
);
