import "server-only";

import { allocateNextCompanyNumber } from "@/lib/companyNumber";
import {
  companyNumberDuplicateInFileError,
  resolveCompanyImport,
  type CompanyImportMatch,
} from "@/lib/services/bulkUpload/companyNumberRules";
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
import {
  createBulkLogger,
  type BulkLogger,
} from "@/lib/services/bulkUpload/bulkUploadLog";
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

/** Existing company whose Company Number equals `number` (case/space-insensitive). */
function matchCompanyByNumber(
  companies: Company[],
  number: string | null | undefined,
): Company | null {
  const key = nameKey(number);
  if (!key) return null;
  return companies.find((c) => nameKey(c.companyNumber) === key) ?? null;
}

function toImportMatch(record: Company | null): CompanyImportMatch {
  return record ? { id: record.id, companyNumber: record.companyNumber } : null;
}

function validateCompanyRow(
  rowNumber: number,
  fields: Record<string, string | null>,
  companies: Company[],
  seenInFile: { numbers: Set<string>; names: Set<string> },
  allocatedNumbers: string[],
): BulkPreviewRow {
  const messages: string[] = [];
  const companyName = fields.companyName?.trim() ?? "";
  const incoming = fields.companyNumber?.trim() ?? "";
  let companyNumber = incoming;

  if (!companyName) messages.push("Company Name is required.");
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

  // Match against existing companies BEFORE any allocation so an existing
  // company's number is never replaced by a freshly generated one.
  const matchByNumber = matchCompanyByNumber(companies, incoming);
  const matchByName = companyName
    ? findCompanyByName(companies, companyName)
    : null;
  const decision = resolveCompanyImport({
    incoming,
    matchByNumber: toImportMatch(matchByNumber),
    matchByName: toImportMatch(matchByName),
    duplicateMode: "update", // preview classifies matched rows; commit applies the chosen mode
  });

  let matchedRecord: Company | null = null;
  let hardError = false;

  if (decision.action === "reject") {
    hardError = true;
    messages.push(decision.message);
    matchedRecord = matchByName ?? matchByNumber;
  } else if (decision.action === "update") {
    matchedRecord = matchByName ?? matchByNumber;
    if (decision.companyNumber) {
      companyNumber = decision.companyNumber;
      if (!incoming) {
        messages.push(
          `Company Number preserved from existing company: ${companyNumber}.`,
        );
      }
    } else {
      // Matched company has no number yet — assign one (not an overwrite).
      companyNumber = allocateNextCompanyNumber(companies, [
        ...seenInFile.numbers,
        ...allocatedNumbers,
      ]);
      messages.push(
        `Existing company has no Company Number — ${companyNumber} will be assigned.`,
      );
    }
  } else if (!incoming) {
    // create with blank number — auto-assign for the new company only.
    companyNumber = allocateNextCompanyNumber(companies, [
      ...seenInFile.numbers,
      ...allocatedNumbers,
    ]);
    messages.push(`Company Number will be auto-assigned: ${companyNumber}.`);
  }

  // Reject the same Company Number appearing twice within one uploaded file.
  let duplicateInFile = false;
  if (!hardError && companyNumber) {
    const key = nameKey(companyNumber);
    if (seenInFile.numbers.has(key)) {
      duplicateInFile = true;
      messages.push(companyNumberDuplicateInFileError(companyNumber));
    } else {
      seenInFile.numbers.add(key);
      allocatedNumbers.push(companyNumber);
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
    companyNumber,
    status: statusInfo.status,
    companySize: sizeInfo.size,
  };

  if (
    hardError ||
    duplicateInFile ||
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
      matchedEntityId: matchedRecord?.id ?? null,
      matchedEntityName: matchedRecord?.companyName ?? null,
      duplicateMatch: null,
    };
  }

  if (matchedRecord) {
    return {
      rowNumber,
      status: "Duplicate",
      messages: [
        ...messages,
        `Matches existing company #${matchedRecord.id} (${matchedRecord.companyName}).`,
      ],
      fields: normalizedFields,
      matchedEntityId: matchedRecord.id,
      matchedEntityName: matchedRecord.companyName,
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
  const allocatedNumbers: string[] = [];
  const rows: BulkPreviewRow[] = [];

  for (let i = 0; i < spreadsheet.rows.length; i++) {
    const raw = spreadsheet.rows[i] ?? {};
    const fields = mapCompanyFields(raw);
    const preview = validateCompanyRow(
      i + 2,
      fields,
      companies,
      seenInFile,
      allocatedNumbers,
    );
    preview.source = raw;
    rows.push(preview);
  }

  return rows;
}

export async function commitCompanyImport(input: {
  rows: BulkCommitRowInput[];
  duplicateMode: BulkDuplicateMode;
  log?: BulkLogger;
}): Promise<BulkPreviewRow[]> {
  const log = input.log ?? createBulkLogger("commit:company");
  const loadPhase = log.phase("load:companies");
  let companies = await listAdminCompanies();
  loadPhase.end({ companies: companies.length });
  const results: BulkPreviewRow[] = [];
  const allocatedInBatch: string[] = [];
  const seenNumbers = new Set<string>(); // reject the same number twice in one file

  const rowsPhase = log.phase("commit:rows");
  let created = 0;
  let updated = 0;
  for (const row of input.rows) {
    const fields = { ...row.fields };
    const companyName = fields.companyName?.trim() ?? "";
    const incoming = fields.companyNumber?.trim() ?? "";

    if (!companyName) {
      results.push({
        rowNumber: row.rowNumber,
        status: "Error",
        messages: ["Company Name is required."],
        fields,
        matchedEntityId: null,
        matchedEntityName: null,
        duplicateMatch: null,
      });
      continue;
    }

    // Match BEFORE any allocation so an existing company's Company Number is
    // never overwritten by a freshly generated one.
    const matchByNumber = matchCompanyByNumber(companies, incoming);
    const matchByName = findCompanyByName(companies, companyName);
    const target = matchByName ?? matchByNumber;
    const decision = resolveCompanyImport({
      incoming,
      matchByNumber: toImportMatch(matchByNumber),
      matchByName: toImportMatch(matchByName),
      duplicateMode: input.duplicateMode,
    });

    const isDuplicateInFile = (value: string | null): boolean =>
      !!value && seenNumbers.has(nameKey(value));

    try {
      if (decision.action === "reject") {
        results.push({
          rowNumber: row.rowNumber,
          status: "Error",
          messages: [decision.message],
          fields,
          matchedEntityId: target?.id ?? null,
          matchedEntityName: target?.companyName ?? null,
          duplicateMatch: null,
        });
        continue;
      }

      if (decision.action === "skip") {
        results.push({
          rowNumber: row.rowNumber,
          status: "Skipped",
          messages: [
            `Skipped duplicate company "${target?.companyName ?? companyName}".`,
          ],
          fields,
          matchedEntityId: decision.targetId,
          matchedEntityName: target?.companyName ?? null,
          duplicateMatch: null,
        });
        continue;
      }

      if (decision.action === "update") {
        let finalNumber = decision.companyNumber;
        let payload: Record<string, unknown>;
        if (finalNumber) {
          // Preserve the existing number: omit it from the payload entirely so
          // SharePoint never rewrites it.
          if (isDuplicateInFile(finalNumber)) {
            results.push({
              rowNumber: row.rowNumber,
              status: "Error",
              messages: [companyNumberDuplicateInFileError(finalNumber)],
              fields: { ...fields, companyNumber: finalNumber },
              matchedEntityId: decision.targetId,
              matchedEntityName: target?.companyName ?? null,
              duplicateMatch: null,
            });
            continue;
          }
          const { companyNumber: _omitNumber, ...rest } = companyWritePayload({
            ...fields,
            companyNumber: finalNumber,
          });
          void _omitNumber;
          payload = rest;
        } else {
          // Matched company has no number yet — assign one (not an overwrite).
          finalNumber = allocateNextCompanyNumber(companies, [
            ...seenNumbers,
            ...allocatedInBatch,
          ]);
          payload = companyWritePayload({ ...fields, companyNumber: finalNumber });
        }

        seenNumbers.add(nameKey(finalNumber));
        fields.companyNumber = finalNumber;
        await updateAdminCompany(decision.targetId, payload);
        // Keep the in-memory list current instead of re-reading the whole
        // Companies list from Graph after every row (was O(n²) reads + a
        // stale-read risk right after a write). The matched company is already
        // in `companies`; only its number may have just been assigned.
        companies = companies.map((company) =>
          company.id === decision.targetId
            ? { ...company, companyNumber: finalNumber }
            : company,
        );
        updated += 1;
        results.push({
          rowNumber: row.rowNumber,
          status: "Imported",
          messages: [
            `Updated existing company #${decision.targetId} (${target?.companyName ?? companyName}); Company Number ${finalNumber} preserved.`,
          ],
          fields,
          matchedEntityId: decision.targetId,
          matchedEntityName: target?.companyName ?? null,
          duplicateMatch: null,
        });
        log.debug("row updated", {
          row: row.rowNumber,
          number: finalNumber,
          id: decision.targetId,
        });
        if ((created + updated) % 10 === 0) {
          log.info("progress", {
            done: created + updated,
            total: input.rows.length,
          });
        }
        continue;
      }

      // action === "create" — new company. A blank number is allocated only
      // here; a supplied number was already proven unique by resolveCompanyImport.
      let finalNumber =
        decision.companyNumber ??
        allocateNextCompanyNumber(companies, [
          ...seenNumbers,
          ...allocatedInBatch,
        ]);
      if (isDuplicateInFile(finalNumber)) {
        results.push({
          rowNumber: row.rowNumber,
          status: "Error",
          messages: [companyNumberDuplicateInFileError(finalNumber)],
          fields: { ...fields, companyNumber: finalNumber },
          matchedEntityId: null,
          matchedEntityName: null,
          duplicateMatch: null,
        });
        continue;
      }
      seenNumbers.add(nameKey(finalNumber));
      allocatedInBatch.push(finalNumber);
      fields.companyNumber = finalNumber;
      const createdCompany = await createAdminCompany(
        // Pass the live in-memory list so createAdminCompany does not re-read
        // the whole Companies list from Graph for its duplicate check (that
        // internal read + the old per-row re-fetch below were both O(n²)).
        {
          ...companyWritePayload({ ...fields, companyNumber: finalNumber }),
          existingCompanies: companies,
        },
      );
      finalNumber = createdCompany.companyNumber ?? finalNumber;
      // Append the new company to the in-memory list so later rows in the same
      // file match it by name/number — no Graph re-fetch, and no read-after-
      // write staleness.
      companies = [...companies, createdCompany];
      created += 1;
      results.push({
        rowNumber: row.rowNumber,
        status: "Imported",
        messages: [
          target
            ? `Created new company despite duplicate (number set to ${finalNumber}).`
            : `Created company #${createdCompany.id} (${finalNumber}).`,
        ],
        fields: { ...fields, companyNumber: finalNumber },
        matchedEntityId: createdCompany.id,
        matchedEntityName: createdCompany.companyName,
        duplicateMatch: null,
      });
      log.debug("row created", {
        row: row.rowNumber,
        number: finalNumber,
        id: createdCompany.id,
      });
      if ((created + updated) % 10 === 0) {
        log.info("progress", {
          done: created + updated,
          total: input.rows.length,
        });
      }
    } catch (error) {
      log.warn("row failed", {
        row: row.rowNumber,
        company: companyName,
        error: error instanceof Error ? error.message : String(error),
      });
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

  rowsPhase.end({
    created,
    updated,
    skipped: results.filter((r) => r.status === "Skipped").length,
    errors: results.filter((r) => r.status === "Error").length,
  });

  return results;
}

export { summarizeBulkRows };
