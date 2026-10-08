/**
 * Patch Company-list-template.xlsx:
 * - Set one sample row Company Size to Enterprise (so all 4 choices appear)
 * - Add Excel data-validation dropdown on Company Size (column C)
 *
 *   node scripts/patch-company-size-template.mjs
 */
import { createRequire } from "node:module";
import { readFileSync, writeFileSync } from "node:fs";
import { deflateRawSync, inflateRawSync } from "node:zlib";
import { resolve } from "node:path";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const TEMPLATE = resolve(
  process.cwd(),
  "public/bulk-templates/Company-list-template.xlsx",
);

const COMPANY_SIZES = ["Small", "Medium", "Large", "Enterprise"];

function parseZip(buf) {
  const entries = [];
  let i = 0;
  while (i < buf.length - 4) {
    const sig = buf.readUInt32LE(i);
    if (sig === 0x04034b50) {
      const compMethod = buf.readUInt16LE(i + 8);
      const compSize = buf.readUInt32LE(i + 18);
      const nameLen = buf.readUInt16LE(i + 26);
      const extraLen = buf.readUInt16LE(i + 28);
      const name = buf.slice(i + 30, i + 30 + nameLen).toString("utf8");
      const dataStart = i + 30 + nameLen + extraLen;
      const data = buf.slice(dataStart, dataStart + compSize);
      let content;
      if (compMethod === 0) content = Buffer.from(data);
      else if (compMethod === 8) content = inflateRawSync(data);
      else throw new Error(`Unsupported compression ${compMethod} for ${name}`);
      entries.push({ name, content });
      i = dataStart + compSize;
    } else if (sig === 0x02014b50 || sig === 0x06054b50) {
      break;
    } else {
      i += 1;
    }
  }
  return entries;
}

function crc32(buf) {
  let c = ~0;
  for (let i = 0; i < buf.length; i += 1) {
    c ^= buf[i];
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
    }
  }
  return ~c >>> 0;
}

function buildZip(entries) {
  const parts = [];
  const central = [];
  let offset = 0;
  for (const entry of entries) {
    const nameBuf = Buffer.from(entry.name, "utf8");
    const raw = entry.content;
    const compressed = deflateRawSync(raw);
    const useStore = compressed.length >= raw.length;
    const payload = useStore ? raw : compressed;
    const method = useStore ? 0 : 8;
    const crc = crc32(raw);
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(payload.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    nameBuf.copy(local, 30);
    parts.push(local, payload);

    const cen = Buffer.alloc(46 + nameBuf.length);
    cen.writeUInt32LE(0x02014b50, 0);
    cen.writeUInt16LE(20, 4);
    cen.writeUInt16LE(20, 6);
    cen.writeUInt16LE(0, 8);
    cen.writeUInt16LE(method, 10);
    cen.writeUInt16LE(0, 12);
    cen.writeUInt16LE(0, 14);
    cen.writeUInt32LE(crc, 16);
    cen.writeUInt32LE(payload.length, 20);
    cen.writeUInt32LE(raw.length, 24);
    cen.writeUInt16LE(nameBuf.length, 28);
    cen.writeUInt16LE(0, 30);
    cen.writeUInt16LE(0, 32);
    cen.writeUInt16LE(0, 34);
    cen.writeUInt16LE(0, 36);
    cen.writeUInt32LE(0, 38);
    cen.writeUInt32LE(offset, 42);
    nameBuf.copy(cen, 46);
    central.push(cen);
    offset += local.length + payload.length;
  }
  const centralBuf = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...parts, centralBuf, end]);
}

// 1) Ensure sample data includes all four sizes via SheetJS rewrite.
{
  const wb = XLSX.readFile(TEMPLATE);
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  const largeIdx = rows.findIndex((r) => r["Company Size"] === "Large");
  if (largeIdx >= 0) {
    rows[largeIdx]["Company Size"] = "Enterprise";
  }
  const present = new Set(rows.map((r) => r["Company Size"]));
  for (const size of COMPANY_SIZES) {
    if (!present.has(size)) {
      throw new Error(`Template still missing Company Size "${size}"`);
    }
  }
  const headers = Object.keys(rows[0] ?? {});
  const aoa = [headers, ...rows.map((r) => headers.map((h) => r[h] ?? ""))];
  const next = XLSX.utils.aoa_to_sheet(aoa);
  wb.Sheets[sheetName] = next;
  XLSX.writeFile(wb, TEMPLATE);
}

// 2) Add data validation on Company Size column (C2:C1000).
{
  const buf = readFileSync(TEMPLATE);
  const entries = parseZip(buf);
  const sheetEntry = entries.find((e) => e.name === "xl/worksheets/sheet1.xml");
  if (!sheetEntry) throw new Error("sheet1.xml not found");
  let xml = sheetEntry.content.toString("utf8");
  xml = xml.replace(/<dataValidations[\s\S]*?<\/dataValidations>/g, "");
  const list = COMPANY_SIZES.map((s) => s.replace(/"/g, "")).join(",");
  const validation = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showDropDown="0" showErrorMessage="1" errorTitle="Company Size" error="Choose Small, Medium, Large, or Enterprise." sqref="C2:C1000"><formula1>"${list}"</formula1></dataValidation></dataValidations>`;
  if (xml.includes("</worksheet>")) {
    xml = xml.replace("</worksheet>", `${validation}</worksheet>`);
  } else {
    throw new Error("Could not find </worksheet>");
  }
  sheetEntry.content = Buffer.from(xml, "utf8");
  writeFileSync(TEMPLATE, buildZip(entries));
}

// Verify
{
  const wb = XLSX.readFile(TEMPLATE);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const sizes = [
    ...new Set(
      XLSX.utils.sheet_to_json(ws).map((r) => r["Company Size"]),
    ),
  ].sort();
  const xml = parseZip(readFileSync(TEMPLATE))
    .find((e) => e.name === "xl/worksheets/sheet1.xml")
    .content.toString("utf8");
  console.log("sizes in template:", sizes.join(", "));
  console.log(
    "dataValidation present:",
    xml.includes("dataValidations") && xml.includes("Enterprise"),
  );
}
