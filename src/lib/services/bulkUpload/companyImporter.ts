import "server-only";

import {
  createAdminCompany,
  listAdminCompanies,
  updateAdminCompany,
} from "@/lib/services/adminCrudService";
import {
  findCompanyByName,
  nameKey,
  normalizeCompanyKey,
} from "@/lib/services/bulkUpload/matching";
import {
  pickField,
  type ParsedSpreadsheet,
} from "@/lib/services/bulkUpload/parseSpreadsheet";
import { summarizeBulkRows } from "@/lib/services/bulkUpload/candidateImporter";
import type {
  BulkCommitRowInput,
  BulkDuplicateMode,
  BulkPreviewRow,
} from "@/types/bulkUpload";
import type { Company } from "@/types/models";

const COMPANY_NUMBER_ALIASES = [
  "Company Number",
  "CompanyNumber",
  "Company No",
  "Company No.",
];
const COMPANY_NAME_ALIASES = [
  "Company Name",
  "CompanyName",
  "Company",
  "Name",
];
const COMPANY_SIZE_ALIASES = ["Company Size", "CompanySize", "Size"];
const REGISTERED_ADDRESS_ALIASES = [
  "Registered Address",
  "RegisteredAddress",
  "Address",
];
const COMPANY_REG_ALIASES = [
  "Company Reg Number",
  "CompanyRegNumber",
  "Company Registration Number",
  "Reg Number",
];
const VAT_ALIASES = ["VAT No", "VATNo", "VAT Number", "Vat No"];
const TEL_ALIASES = ["Tel No", "TelNo", "Telephone", "Phone", "Tel"];
const EMAIL_ALIASES = ["Email", "E-mail", "Company Email"];
const MAIN_CONTACT_ALIASES = [
  "Main Contact",
  "MainContact",
  "Main Contact Name",
];
const ACCOUNTS_CONTACT_ALIASES = [
  "Accounts Contact Name",
  "AccountsContactName",
  "Accounts Contact",
];
const ACCOUNTS_ADDRESS_ALIASES = [
  "Accounts address",
  "Accounts Address",
  "Accountsaddress",
  "AccountsAddress",
];
const ACCOUNTS_NUMBER_ALIASES = [
  "Accounts Contact number",
  "Accounts Contact Number",
  "AccountsContactnumber",
  "AccountsContactNumber",
];
const ACCOUNTS_EMAIL_ALIASES = [
  "Accounts email",
  "Accounts Email",
  "Accountsemail",
  "AccountsEmail",
];
const NOTES_ALIASES = [
  "Notes prices agreed",
  "Notes Prices Agreed",
  "Notespricesagreed",
  "Notes",
];
const LOGO_ALIASES = ["Company Logo", "CompanyLogo", "Logo"];
const STATUS_ALIASES = ["Status", "Company Status"];

function mapCompanyFields(
  raw: Record<string, string | null>,
): Record<string, string | null> {
  return {
    companyNumber: pickField(raw, COMPANY_NUMBER_ALIASES),
    companyName: pickField(raw, COMPANY_NAME_ALIASES),
    companySize: pickField(raw, COMPANY_SIZE_ALIASES),
    registeredAddress: pickField(raw, REGISTERED_ADDRESS_ALIASES),
    companyRegNumber: pickField(raw, COMPANY_REG_ALIASES),
    vatNo: pickField(raw, VAT_ALIASES),
    telNo: pickField(raw, TEL_ALIASES),
    email: pickField(raw, EMAIL_ALIASES),
    mainContact: pickField(raw, MAIN_CONTACT_ALIASES),
    accountsContactName: pickField(raw, ACCOUNTS_CONTACT_ALIASES),
    accountsAddress: pickField(raw, ACCOUNTS_ADDRESS_ALIASES),
    accountsContactNumber: pickField(raw, ACCOUNTS_NUMBER_ALIASES),
    accountsEmail: pickField(raw, ACCOUNTS_EMAIL_ALIASES),
    notesPricesAgreed: pickField(raw, NOTES_ALIASES),
    companyLogo: pickField(raw, LOGO_ALIASES),
    status: pickField(raw, STATUS_ALIASES),
  };
}

function normalizeStatus(value: string | null | undefined): {
  status: string;
  warning?: string;
} {
  const raw = (value ?? "").trim();
  if (!raw) return { status: "Active" };
  const lower = raw.toLowerCase();
  if (lower === "active") return { status: "Active" };
  if (lower === "inactive") return { status: "Inactive" };
  if (lower === "on hold" || lower === "onhold" || lower === "hold") {
    return {
      status: "Inactive",
      warning: `Status "${raw}" mapped to Inactive (SharePoint allows Active | Inactive).`,
    };
  }
  return {
    status: "Active",
    warning: `Unknown status "${raw}" — defaulting to Active.`,
  };
}

function normalizeCompanySize(value: string | null | undefined): {
  size: string | null;
  warning?: string;
} {
  const raw = value?.trim() ?? "";
  if (!raw) return { size: null };
  const lower = raw.toLowerCase();
  if (lower === "small") return { size: "Small" };
  if (lower === "medium") return { size: "Medium" };
  if (lower === "large") return { size: "Large" };
  if (lower === "enterprise") return { size: "Enterprise" };
  return {
    size: raw,
    warning: `Company Size "${raw}" is unusual — expected Small, Medium, Large, or Enterprise.`,
  };
}

function normalizeLogo(value: string | null | undefined): string | null {
  // Thumbnail column — spreadsheet text/placeholders are never written to SharePoint.
  void value;
  return null;
}

function companyWritePayload(
  fields: Record<string, string | null>,
): Record<string, unknown> {
  const statusInfo = normalizeStatus(fields.status);
  const sizeInfo = normalizeCompanySize(fields.companySize);
  const payload: Record<string, unknown> = {
    companyName: fields.companyName?.trim(),
    companyNumber: fields.companyNumber?.trim(),
    status: statusInfo.status,
  };
  const assign = (key: string, value: string | null | undefined) => {
    if (value?.trim()) payload[key] = value.trim();
  };
  if (sizeInfo.size) payload.companySize = sizeInfo.size;
  assign("registeredAddress", fields.registeredAddress);
  assign("companyRegNumber", fields.companyRegNumber);
  assign("vatNo", fields.vatNo);
  assign("telNo", fields.telNo);
  assign("email", fields.email);
  assign("mainContact", fields.mainContact);
  assign("accountsContactName", fields.accountsContactName);
  assign("accountsAddress", fields.accountsAddress);
  assign("accountsContactNumber", fields.accountsContactNumber);
  assign("accountsEmail", fields.accountsEmail);
  assign("notesPricesAgreed", fields.notesPricesAgreed);
  const logo = normalizeLogo(fields.companyLogo);
  if (logo) payload.companyLogo = logo;
  return payload;
}

function findCompanyDuplicate(
  companies: Company[],
  fields: Record<string, string | null>,
): { kind: "companyNumber" | "companyName"; record: Company } | null {
  const number = fields.companyNumber?.trim();
  if (number) {
    const byNumber = companies.find(
      (c) => nameKey(c.companyNumber) === nameKey(number),
    );
    if (byNumber) return { kind: "companyNumber", record: byNumber };
  }

  const byName = findCompanyByName(companies, fields.companyName);
  if (byName) return { kind: "companyName", record: byName };
  return null;
}

function validateCompanyRow(
  rowNumber: number,
  fields: Record<string, string | null>,
  companies: Company[],
  seenInFile: { numbers: Set<string>; names: Set<string> },
): BulkPreviewRow {
  const messages: string[] = [];
  const companyName = fields.companyName?.trim() ?? "";
  const companyNumber = fields.companyNumber?.trim() ?? "";

  if (!companyName) messages.push("Company Name is required.");
  if (!companyNumber) messages.push("Company Number is required.");
  if (!fields.email?.trim()) {
    messages.push(
      "Email is required in SharePoint Company List — row will fail without it.",
    );
  }
  if (!fields.registeredAddress?.trim()) {
    messages.push(
      "Registered Address is required in SharePoint Company List — row will fail without it.",
    );
  }
  if (!fields.companySize?.trim()) {
    messages.push(
      "Company Size is required in SharePoint Company List — row will fail without it.",
    );
  }

  if (companyNumber) {
    const key = nameKey(companyNumber);
    if (seenInFile.numbers.has(key)) {
      messages.push(
        `Duplicate Company Number "${companyNumber}" appears earlier in this file.`,
      );
    } else {
      seenInFile.numbers.add(key);
    }
  }

  if (companyName) {
    const key = normalizeCompanyKey(companyName);
    if (seenInFile.names.has(key)) {
      messages.push(
        `Duplicate Company Name "${companyName}" appears earlier in this file.`,
      );
    } else {
      seenInFile.names.add(key);
    }
  }

  const sizeInfo = normalizeCompanySize(fields.companySize);
  if (sizeInfo.warning) messages.push(sizeInfo.warning);

  const statusInfo = normalizeStatus(fields.status);
  if (statusInfo.warning) messages.push(statusInfo.warning);

  const normalizedFields = {
    ...fields,
    status: statusInfo.status,
    companySize: sizeInfo.size,
  };

  if (
    !companyName ||
    !companyNumber ||
    !fields.email?.trim() ||
    !fields.registeredAddress?.trim() ||
    !fields.companySize?.trim()
  ) {
    return {
      rowNumber,
      status: "Error",
      messages,
      fields: normalizedFields,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
    };
  }

  const duplicate = findCompanyDuplicate(companies, fields);
  if (duplicate) {
    return {
      rowNumber,
      status: "Duplicate",
      messages: [
        ...messages,
        duplicate.kind === "companyNumber"
          ? `Matches existing company by Company Number (#${duplicate.record.id}).`
          : `Matches existing company by Company Name (#${duplicate.record.id}).`,
      ],
      fields: normalizedFields,
      matchedEntityId: duplicate.record.id,
      matchedEntityName: duplicate.record.companyName,
      duplicateMatch: null,
    };
  }

  const hasWarnings = messages.length > 0;
  return {
    rowNumber,
    status: hasWarnings ? "Warning" : "Ready",
    messages,
    fields: normalizedFields,
    matchedEntityId: null,
    matchedEntityName: null,
    duplicateMatch: null,
  };
}

export async function previewCompanyImport(
  spreadsheet: ParsedSpreadsheet,
): Promise<BulkPreviewRow[]> {
  const companies = await listAdminCompanies();
  const seenInFile = { numbers: new Set<string>(), names: new Set<string>() };
  const rows: BulkPreviewRow[] = [];

  for (let i = 0; i < spreadsheet.rows.length; i++) {
    const raw = spreadsheet.rows[i] ?? {};
    const fields = mapCompanyFields(raw);
    const preview = validateCompanyRow(
      i + 2,
      fields,
      companies,
      seenInFile,
    );
    preview.source = raw;
    rows.push(preview);
  }

  return rows;
}

export async function commitCompanyImport(input: {
  rows: BulkCommitRowInput[];
  duplicateMode: BulkDuplicateMode;
}): Promise<BulkPreviewRow[]> {
  let companies = await listAdminCompanies();
  const results: BulkPreviewRow[] = [];

  for (const row of input.rows) {
    const fields = { ...row.fields };
    const companyName = fields.companyName?.trim() ?? "";
    const companyNumber = fields.companyNumber?.trim() ?? "";

    if (!companyName || !companyNumber) {
      results.push({
        rowNumber: row.rowNumber,
        status: "Error",
        messages: ["Company Name and Company Number are required."],
        fields,
        matchedEntityId: null,
        matchedEntityName: null,
        duplicateMatch: null,
      });
      continue;
    }

    const duplicate = findCompanyDuplicate(companies, fields);
    const payload = companyWritePayload(fields);

    try {
      if (duplicate) {
        if (input.duplicateMode === "skip") {
          results.push({
            rowNumber: row.rowNumber,
            status: "Skipped",
            messages: [
              `Skipped duplicate company "${duplicate.record.companyName}".`,
            ],
            fields,
            matchedEntityId: duplicate.record.id,
            matchedEntityName: duplicate.record.companyName,
            duplicateMatch: null,
          });
          continue;
        }

        if (input.duplicateMode === "update") {
          await updateAdminCompany(duplicate.record.id, payload);
          results.push({
            rowNumber: row.rowNumber,
            status: "Imported",
            messages: [
              `Updated existing company #${duplicate.record.id} (${duplicate.record.companyName}).`,
            ],
            fields,
            matchedEntityId: duplicate.record.id,
            matchedEntityName: duplicate.record.companyName,
            duplicateMatch: null,
          });
          companies = await listAdminCompanies();
          continue;
        }

        // create despite duplicate — force unique number suffix
        const forcedNumber = `${companyNumber}-DUP-${row.rowNumber}`;
        const created = await createAdminCompany({
          ...payload,
          companyNumber: forcedNumber,
        });
        results.push({
          rowNumber: row.rowNumber,
          status: "Imported",
          messages: [
            `Created new company despite duplicate (number set to ${forcedNumber}).`,
          ],
          fields: { ...fields, companyNumber: forcedNumber },
          matchedEntityId: created.id,
          matchedEntityName: created.companyName,
          duplicateMatch: null,
        });
        companies = await listAdminCompanies();
        continue;
      }

      const created = await createAdminCompany(payload);
      results.push({
        rowNumber: row.rowNumber,
        status: "Imported",
        messages: [`Created company #${created.id}.`],
        fields,
        matchedEntityId: created.id,
        matchedEntityName: created.companyName,
        duplicateMatch: null,
      });
      companies = await listAdminCompanies();
    } catch (error) {
      results.push({
        rowNumber: row.rowNumber,
        status: "Error",
        messages: [
          error instanceof Error ? error.message : "Company import failed.",
        ],
        fields,
        matchedEntityId: null,
        matchedEntityName: null,
        duplicateMatch: null,
      });
    }
  }

  return results;
}

export { summarizeBulkRows };
