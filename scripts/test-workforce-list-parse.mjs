/**
 * Dry-run: parse Workforce list.xlsx the same way the portal bulk uploader does.
 * Does not call SharePoint.
 *
 *   node scripts/test-workforce-list-parse.mjs
 *   node scripts/test-workforce-list-parse.mjs "path/to/file.xlsx"
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import * as XLSX from "xlsx";

const EXCEL_PATH = resolve(
  process.cwd(),
  process.argv[2] || "Workforce list.xlsx",
);

function cellToString(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof value === "number") return String(value);
  const text = String(value).trim();
  if (!text) return null;
  if (/^(—|–|-|n\/?a|null|none)$/i.test(text)) return null;
  return text;
}

function normalizeHeader(header) {
  return header.trim().toLowerCase().replace(/\s+/g, " ");
}

function pickField(row, aliases) {
  const map = new Map();
  for (const [key, value] of Object.entries(row)) {
    map.set(normalizeHeader(key), value);
  }
  for (const alias of aliases) {
    const hit = map.get(normalizeHeader(alias));
    if (hit) return hit;
  }
  return null;
}

function normalizeDateValue(value) {
  if (!value?.trim()) return null;
  const text = value.trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  const uk = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (uk) {
    return `${uk[3]}-${uk[2].padStart(2, "0")}-${uk[1].padStart(2, "0")}`;
  }
  const short = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2})$/);
  if (short) {
    const first = Number(short[1]);
    const second = Number(short[2]);
    let year = Number(short[3]);
    year += year >= 70 ? 1900 : 2000;
    let month;
    let day;
    if (first > 12 && second <= 12) {
      day = first;
      month = second;
    } else {
      month = first;
      day = second;
    }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }
  if (/^\d+(\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 20000 && serial < 60000) {
      const parsed = XLSX.SSF.parse_date_code(serial);
      if (parsed) {
        return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
      }
    }
  }
  return text;
}

const bytes = readFileSync(EXCEL_PATH);
const workbook = XLSX.read(bytes, { type: "buffer", cellDates: true, raw: false });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const matrix = XLSX.utils.sheet_to_json(sheet, {
  header: 1,
  defval: null,
  raw: false,
  blankrows: false,
});

const headers = (matrix[0] ?? [])
  .map((cell) => (cell == null ? "" : String(cell).trim()))
  .filter((h) => h.length > 0);

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
  if (hasAny) rows.push({ rowNumber: i + 1, raw: row });
}

let ready = 0;
let errors = 0;
const errorSamples = [];
const companies = new Set();

for (const { rowNumber, raw } of rows) {
  const candidateName = pickField(raw, ["Candidate Name", "Name"]);
  const company = pickField(raw, ["Company Name", "Company"]);
  const email = pickField(raw, ["Email", "E-mail"]);
  const department = pickField(raw, ["Department", " Department"]);
  const dob = normalizeDateValue(pickField(raw, ["Date of birth", "DOB"]));
  const messages = [];
  if (!candidateName) messages.push("Candidate Name required");
  if (!company) messages.push("Company Name required");
  if (!email) messages.push("Email required");
  companies.add(company || "(missing)");
  if (messages.length) {
    errors += 1;
    if (errorSamples.length < 10) {
      errorSamples.push({ rowNumber, candidateName, messages, dob, department });
    }
  } else {
    ready += 1;
  }
}

console.log(`File: ${EXCEL_PATH}`);
console.log(`Headers (${headers.length}): ${headers.join(" | ")}`);
console.log(`Data rows: ${rows.length}`);
console.log(`Unique companies: ${companies.size}`);
console.log(`Would pass required-field checks: ${ready}`);
console.log(`Required-field errors: ${errors}`);
if (errorSamples.length) {
  console.log("Error samples:", JSON.stringify(errorSamples, null, 2));
}

const sample = rows[0];
if (sample) {
  console.log("\nSample mapped row 1:");
  console.log(
    JSON.stringify(
      {
        candidateName: pickField(sample.raw, ["Candidate Name"]),
        company: pickField(sample.raw, ["Company Name"]),
        email: pickField(sample.raw, ["Email"]),
        department: pickField(sample.raw, ["Department", " Department"]),
        dateOfBirth: normalizeDateValue(
          pickField(sample.raw, ["Date of birth", "DOB"]),
        ),
        status: pickField(sample.raw, ["Status"]) || "Active",
        workforceNumber: pickField(sample.raw, ["Workforce Number"]),
      },
      null,
      2,
    ),
  );
}

if (errors > 0) process.exitCode = 1;
else console.log("\nRESULT: PASS — spreadsheet parses cleanly for portal bulk upload");
