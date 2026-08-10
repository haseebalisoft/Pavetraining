/**
 * Thorough local checks for Training matrix example.xlsx (no SharePoint, no .ts imports).
 * Works on Node 20+.
 *
 *   node scripts/test-matrix-example-xlsx.mjs
 */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { resolve } from "node:path";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

/** Same UK-first rules as src/lib/utils/ukDate.ts (inlined for Node 20). */
function normalizeDateValue(value) {
  if (value == null || !String(value).trim()) return null;
  const text = String(value).trim();
  if (/^(—|–|-|n\/?a|null|none)$/i.test(text)) return null;

  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);

  const uk = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (uk) {
    const day = Number(uk[1]);
    const month = Number(uk[2]);
    const year = Number(uk[3]);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    const probe = new Date(Date.UTC(year, month - 1, day));
    if (
      probe.getUTCFullYear() !== year ||
      probe.getUTCMonth() !== month - 1 ||
      probe.getUTCDate() !== day
    ) {
      return null;
    }
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 20000 && serial < 60000) {
      const ms = Date.UTC(1899, 11, 30) + Math.round(serial) * 86_400_000;
      return new Date(ms).toISOString().slice(0, 10);
    }
  }

  return null;
}

function formatDateUk(iso) {
  if (!iso || !/^\d{4}-\d{2}-\d{2}/.test(iso)) return "—";
  const [y, m, d] = iso.slice(0, 10).split("-");
  return `${d}/${m}/${y}`;
}

function cellToString(value) {
  if (value == null) return null;
  if (typeof value === "number") {
    if (value > 20000 && value < 60000) {
      const ms = Date.UTC(1899, 11, 30) + Math.round(value) * 86_400_000;
      return new Date(ms).toISOString().slice(0, 10);
    }
    return String(value);
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  const text = String(value).trim();
  if (!text || /^(—|–|-|n\/?a|null|none)$/i.test(text)) return null;
  return text;
}

const filePath = resolve(process.cwd(), "Training matrix example.xlsx");
const wb = XLSX.read(readFileSync(filePath), {
  type: "buffer",
  cellDates: false,
  raw: true,
});
const sheet = wb.Sheets[wb.SheetNames[0]];
const matrix = XLSX.utils.sheet_to_json(sheet, {
  header: 1,
  defval: null,
  raw: true,
  blankrows: false,
});

const headers = (matrix[0] ?? [])
  .map((c) => (c == null ? "" : String(c).trim()))
  .filter(Boolean);

const rows = [];
for (let i = 1; i < matrix.length; i += 1) {
  const raw = matrix[i] ?? [];
  const row = {};
  let hasAny = false;
  for (let c = 0; c < headers.length; c += 1) {
    const value = cellToString(raw[c]);
    row[headers[c]] = value;
    if (value) hasAny = true;
  }
  if (hasAny) rows.push(row);
}

console.log(`File: ${filePath}`);
console.log(`Node: ${process.version}`);
console.log(`Rows: ${rows.length}`);

let named = 0;
let dobOk = 0;
const samples = [];

for (const row of rows) {
  const name = row.Name?.trim();
  if (name) named += 1;
  const dobIso = normalizeDateValue(row.DOB);
  if (dobIso && /^\d{4}-\d{2}-\d{2}$/.test(dobIso)) {
    dobOk += 1;
    if (samples.length < 5) {
      samples.push({
        name,
        dobIso,
        display: formatDateUk(dobIso),
        cscs: normalizeDateValue(row["CSCS Expiry"]),
      });
    }
  }
}

console.log("\nSamples:");
for (const s of samples) console.log(" ", s);
console.log("\nCounts:", {
  named,
  dobOk,
  emptyName: rows.length - named,
  total: rows.length,
});

assert.ok(named >= 50, "expected at least 50 named candidate rows");

const zeeshan = rows.find((r) => String(r.Name).toLowerCase() === "zeeshan");
assert.ok(zeeshan, "Zeeshan row present");
assert.equal(
  normalizeDateValue(zeeshan.DOB),
  "1979-06-12",
  "Zeeshan DOB from Excel serial is 12 June 1979",
);

const zaid = rows.find((r) => String(r.Name).toLowerCase() === "zaid");
assert.ok(zaid, "Zaid row present");
assert.equal(normalizeDateValue(zaid.DOB), "1972-01-01", "Zaid DOB is 1 Jan 1972");

console.log("\nZeeshan UK:", formatDateUk(normalizeDateValue(zeeshan.DOB)));
console.log("Zaid UK:", formatDateUk(normalizeDateValue(zaid.DOB)));
console.log("\nOK — Training matrix example.xlsx dates + names are importable.");
console.log(
  "Upsert rule: create if missing, OVERRIDE if matrix row already exists.",
);
console.log("Empty Name rows will show as Errors (padding rows in this file).");
