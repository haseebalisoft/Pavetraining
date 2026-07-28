import "server-only";

import {
  createAdminCompany,
  createAdminWorkforce,
  findPermissionPerson,
  listAdminCompanies,
  listAdminWorkforce,
  loadPermissionPeople,
  updateAdminWorkforce,
  type AdminWorkforceRecord,
  type PermissionPersonRef,
} from "@/lib/services/adminCrudService";
import {
  findCandidateDuplicate,
  findCompanyByName,
} from "@/lib/services/bulkUpload/matching";
import {
  normalizeDateValue,
  pickField,
  type ParsedSpreadsheet,
} from "@/lib/services/bulkUpload/parseSpreadsheet";
import type {
  BulkCommitRowInput,
  BulkDuplicateMode,
  BulkImportSummary,
  BulkPreviewRow,
} from "@/types/bulkUpload";
import type { Company } from "@/types/models";

const NAME_ALIASES = [
  "Candidate Name",
  "CandidateName",
  "Name",
  "Full Name",
];
const COMPANY_ALIASES = ["Company", "Company Name", "CompanyName"];
const COMPANY_NUMBER_ALIASES = [
  "Company Number",
  "CompanyNumber",
  "Company No",
];
const DEPT_ALIASES = ["Department", "Dept", " Department"];
const DOB_ALIASES = [
  "DOB",
  "Date of Birth",
  "Date of birth",
  "DateOfBirth",
  "Birth Date",
];
const WF_ALIASES = [
  "Workforce Number",
  "WorkforceNumber",
  "Workforce No",
  "Workforce No.",
];
const STATUS_ALIASES = ["Status"];
const TM_ALIASES = ["Training Manager", "Training manager", "TrainingManager"];
const SUPERVISOR_ALIASES = ["Supervisor"];

function emptySummary(): BulkImportSummary {
  return {
    totalRows: 0,
    readyRows: 0,
    warningRows: 0,
    duplicateRows: 0,
    errorRows: 0,
    skippedRows: 0,
    importedRows: 0,
  };
}

export function summarizeBulkRows(rows: BulkPreviewRow[]): BulkImportSummary {
  const summary = emptySummary();
  summary.totalRows = rows.length;
  for (const row of rows) {
    switch (row.status) {
      case "Ready":
        summary.readyRows += 1;
        break;
      case "Warning":
        summary.warningRows += 1;
        break;
      case "Duplicate":
        summary.duplicateRows += 1;
        break;
      case "Error":
        summary.errorRows += 1;
        break;
      case "Skipped":
        summary.skippedRows += 1;
        break;
      case "Imported":
        summary.importedRows += 1;
        break;
      default:
        break;
    }
  }
  return summary;
}

function mapCandidateFields(
  raw: Record<string, string | null>,
): Record<string, string | null> {
  return {
    candidateName: pickField(raw, NAME_ALIASES),
    company: pickField(raw, COMPANY_ALIASES),
    companyNumber: pickField(raw, COMPANY_NUMBER_ALIASES),
    department: pickField(raw, DEPT_ALIASES),
    dateOfBirth: normalizeDateValue(pickField(raw, DOB_ALIASES)),
    workforceNumber: pickField(raw, WF_ALIASES),
    status: pickField(raw, STATUS_ALIASES),
    trainingManager: pickField(raw, TM_ALIASES),
    supervisor: pickField(raw, SUPERVISOR_ALIASES),
    candidateAddress: pickField(raw, [
      "Candidate Address",
      "CandidateAddress",
      "Address",
    ]),
    email: pickField(raw, ["Email", "E-mail"]),
    contactNumber: pickField(raw, [
      "Contact number",
      "Contact Number",
      "Phone",
      "Mobile",
    ]),
    niNumber: pickField(raw, ["Ni Number", "NI Number", "NiNumber", "NI"]),
    nporsNumbers: pickField(raw, [
      "NPORS Number",
      "NPORS Numbers",
      "NPORSNumbers",
    ]),
    cscsNumber: pickField(raw, ["CSCS Number", "CSCSNumber"]),
    cscsExpiry: normalizeDateValue(
      pickField(raw, ["Cscs Expiry", "CSCS Expiry", "CscsExpiry"]),
    ),
    swqrNumber: pickField(raw, ["SWQR Number", "SWQRNumber"]),
    swqrExpiry: normalizeDateValue(
      pickField(raw, ["Swqr Expiry", "SWQR Expiry", "SwqrExpiry"]),
    ),
    eusrNumber: pickField(raw, ["EUSR Number", "EUSRNumber"]),
    eusrExpiry: normalizeDateValue(
      pickField(raw, ["Eusr Expiry", "EUSR Expiry", "EusrExpiry"]),
    ),
    inHouseCertificationNumber: pickField(raw, [
      "In House Certification Number",
      "In-House Certification Number",
      "InHouseCertificationNumber",
    ]),
    notes: pickField(raw, ["Notes"]),
  };
}

function workforceWritePayload(
  fields: Record<string, string | null>,
  resolvedCompanyName: string,
): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    candidateName: fields.candidateName,
    companyName: resolvedCompanyName,
  };
  const assign = (key: string, value: string | null | undefined) => {
    if (value?.trim()) payload[key] = value.trim();
  };
  assign("workforceNumber", fields.workforceNumber);
  assign("dateOfBirth", fields.dateOfBirth);
  // Department is a Lookup → Departments list (resolved in create/update).
  assign("department", fields.department);
  assign("departmentText", fields.department);
  assign("status", fields.status ?? "Active");
  assign("trainingManager", fields.trainingManager);
  assign("supervisor", fields.supervisor);
  assign("candidateAddress", fields.candidateAddress);
  assign("email", fields.email);
  assign("contactNumber", fields.contactNumber);
  assign("niNumber", fields.niNumber);
  assign("nporsNumbers", fields.nporsNumbers);
  assign("cscsNumber", fields.cscsNumber);
  assign("cscsExpiry", fields.cscsExpiry);
  assign("swqrNumber", fields.swqrNumber);
  assign("swqrExpiry", fields.swqrExpiry);
  assign("eusrNumber", fields.eusrNumber);
  assign("eusrExpiry", fields.eusrExpiry);
  assign("inHouseCertificationNumber", fields.inHouseCertificationNumber);
  assign("notes", fields.notes);
  return payload;
}

function buildNonBlankUpdate(
  fields: Record<string, string | null>,
  resolvedCompanyName: string,
): Record<string, unknown> {
  const full = workforceWritePayload(fields, resolvedCompanyName);
  const payload: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(full)) {
    if (value === null || value === undefined) continue;
    if (typeof value === "string" && !value.trim()) continue;
    payload[key] = value;
  }
  return payload;
}

function validateCandidateRow(
  rowNumber: number,
  fields: Record<string, string | null>,
  companies: Company[],
  workforce: AdminWorkforceRecord[],
  people: PermissionPersonRef[],
): BulkPreviewRow {
  const messages: string[] = [];
  const candidateName = fields.candidateName?.trim() ?? "";
  const companyInput = fields.company?.trim() ?? "";

  if (!candidateName) {
    messages.push("Candidate Name is required.");
  }
  if (!companyInput) {
    messages.push("Company Name is required.");
  }
  if (!fields.email?.trim()) {
    messages.push(
      "Email is required in SharePoint Workforce List — row will fail without it.",
    );
  }

  const company = findCompanyByName(companies, companyInput);
  if (companyInput && !company) {
    messages.push(
      `Company "${companyInput}" was not found — it will be created on import.`,
    );
  }

  const tm = fields.trainingManager?.trim();
  if (tm && !findPermissionPerson(people, tm)) {
    messages.push(
      `Training manager "${tm}" not in Permissions — will be created on import.`,
    );
  }
  const supervisor = fields.supervisor?.trim();
  if (supervisor && !findPermissionPerson(people, supervisor)) {
    messages.push(
      `Supervisor "${supervisor}" not in Permissions — will be created on import.`,
    );
  }

  if (!candidateName || !companyInput) {
    return {
      rowNumber,
      status: "Error",
      messages,
      fields,
      resolvedCompanyName: company?.companyName ?? null,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
    };
  }

  if (!fields.email?.trim()) {
    return {
      rowNumber,
      status: "Error",
      messages,
      fields,
      resolvedCompanyName: company?.companyName ?? null,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
    };
  }

  const resolvedCompanyName = company?.companyName ?? companyInput;
  const duplicate = findCandidateDuplicate(workforce, {
    candidateName,
    companyName: resolvedCompanyName,
    workforceNumber: fields.workforceNumber,
    dateOfBirth: fields.dateOfBirth,
  });

  if (duplicate?.kind === "workforceNumber") {
    return {
      rowNumber,
      status: "Duplicate",
      messages: [
        `Duplicate: workforce number matches existing candidate "${duplicate.record.candidateName}".`,
      ],
      fields: { ...fields, company: resolvedCompanyName },
      resolvedCompanyName,
      matchedEntityId: duplicate.record.id,
      matchedEntityName: duplicate.record.candidateName,
      duplicateMatch: "workforceNumber",
    };
  }

  if (duplicate?.kind === "nameDobCompany") {
    return {
      rowNumber,
      status: "Duplicate",
      messages: [
        `Duplicate: name + DOB + company matches existing candidate "${duplicate.record.candidateName}".`,
      ],
      fields: { ...fields, company: resolvedCompanyName },
      resolvedCompanyName,
      matchedEntityId: duplicate.record.id,
      matchedEntityName: duplicate.record.candidateName,
      duplicateMatch: "nameDobCompany",
    };
  }

  if (duplicate?.kind === "nameCompany") {
    return {
      rowNumber,
      status: "Warning",
      messages: [
        `Possible match: name + company matches existing candidate "${duplicate.record.candidateName}". Import will create a new record unless you choose Update existing.`,
      ],
      fields: { ...fields, company: resolvedCompanyName },
      resolvedCompanyName,
      matchedEntityId: duplicate.record.id,
      matchedEntityName: duplicate.record.candidateName,
      duplicateMatch: "nameCompany",
    };
  }

  if (!company) {
    // keep as Warning — will auto-create company on commit
    return {
      rowNumber,
      status: "Warning",
      messages,
      fields: { ...fields, company: resolvedCompanyName },
      resolvedCompanyName,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
    };
  }

  if (!fields.workforceNumber) {
    messages.push("Workforce Number is missing (optional).");
  }
  if (!fields.dateOfBirth) {
    messages.push("DOB is missing (optional).");
  }

  return {
    rowNumber,
    status: messages.length ? "Warning" : "Ready",
    messages,
    fields: { ...fields, company: resolvedCompanyName },
    resolvedCompanyName,
    matchedEntityId: null,
    matchedEntityName: null,
    duplicateMatch: null,
  };
}

async function ensureCompany(
  companies: Company[],
  companyName: string,
  companyNumber: string | null,
): Promise<{ companies: Company[]; company: Company }> {
  const existing = findCompanyByName(companies, companyName);
  if (existing) {
    return { companies, company: existing };
  }
  const created = await createAdminCompany({
    companyName,
    companyNumber: companyNumber?.trim() || `AUTO-${Date.now()}`,
    status: "Active",
  });
  return { companies: [...companies, created], company: created };
}

export async function previewCandidateImport(
  spreadsheet: ParsedSpreadsheet,
): Promise<BulkPreviewRow[]> {
  const [companies, workforce, people] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
    loadPermissionPeople(),
  ]);

  return spreadsheet.rows.map((raw, index) => {
    const fields = mapCandidateFields(raw);
    const validated = validateCandidateRow(
      index + 2,
      fields,
      companies,
      workforce,
      people,
    );
    return {
      ...validated,
      source: raw,
    };
  });
}

export async function commitCandidateImport(input: {
  rows: BulkCommitRowInput[];
  duplicateMode: BulkDuplicateMode;
}): Promise<BulkPreviewRow[]> {
  let companies = await listAdminCompanies();
  const [workforce, initialPeople] = await Promise.all([
    listAdminWorkforce(),
    loadPermissionPeople(),
  ]);
  let people = initialPeople;
  const liveWorkforce = [...workforce];
  const results: BulkPreviewRow[] = [];

  for (const row of input.rows) {
    const fields = mapCandidateFields(row.fields);
    // Prefer already-normalized preview fields when present.
    for (const [key, value] of Object.entries(row.fields)) {
      if (value != null && value !== "") {
        fields[key] = value;
      }
    }
    fields.dateOfBirth = normalizeDateValue(fields.dateOfBirth);
    fields.cscsExpiry = normalizeDateValue(fields.cscsExpiry);
    fields.swqrExpiry = normalizeDateValue(fields.swqrExpiry);
    fields.eusrExpiry = normalizeDateValue(fields.eusrExpiry);

    const validated = validateCandidateRow(
      row.rowNumber,
      fields,
      companies,
      liveWorkforce,
      people,
    );

    if (validated.status === "Error") {
      results.push(validated);
      continue;
    }

    try {
      const ensured = await ensureCompany(
        companies,
        validated.resolvedCompanyName ?? validated.fields.company ?? "",
        validated.fields.companyNumber,
      );
      companies = ensured.companies;
      const companyName = ensured.company.companyName;

      const runCreate = async (extraMessages: string[]) => {
        const created = await createAdminWorkforce({
          ...workforceWritePayload(validated.fields, companyName),
          createMissingPermissionPeople: true,
        });
        liveWorkforce.push(created);
        people = await loadPermissionPeople();
        results.push({
          ...validated,
          status: "Imported",
          matchedEntityId: created.id,
          matchedEntityName: created.candidateName,
          resolvedCompanyName: companyName,
          messages: [...validated.messages, ...extraMessages],
        });
      };

      const runUpdate = async (id: string, extraMessages: string[]) => {
        const updated = await updateAdminWorkforce(id, {
          ...buildNonBlankUpdate(validated.fields, companyName),
          createMissingPermissionPeople: true,
        });
        const idx = liveWorkforce.findIndex((w) => w.id === updated.id);
        if (idx >= 0) liveWorkforce[idx] = updated;
        people = await loadPermissionPeople();
        results.push({
          ...validated,
          status: "Imported",
          resolvedCompanyName: companyName,
          messages: [...validated.messages, ...extraMessages],
        });
      };

      if (
        validated.status === "Warning" &&
        validated.duplicateMatch === "nameCompany"
      ) {
        if (input.duplicateMode === "update" && validated.matchedEntityId) {
          await runUpdate(
            validated.matchedEntityId,
            ["Updated existing candidate (name + company match)."],
          );
          continue;
        }
        await runCreate([
          "Created new candidate (soft name match was a warning only).",
        ]);
        continue;
      }

      if (validated.status === "Duplicate") {
        if (input.duplicateMode === "skip") {
          results.push({
            ...validated,
            status: "Skipped",
            messages: [...validated.messages, "Skipped duplicate (default)."],
          });
          continue;
        }
        if (input.duplicateMode === "update" && validated.matchedEntityId) {
          await runUpdate(validated.matchedEntityId, [
            "Updated existing candidate.",
          ]);
          continue;
        }
        if (input.duplicateMode === "create") {
          await runCreate([
            "Created new candidate (admin confirmed create despite duplicate).",
          ]);
          continue;
        }
        results.push({
          ...validated,
          status: "Skipped",
          messages: [...validated.messages, "Skipped duplicate."],
        });
        continue;
      }

      await runCreate(["Imported successfully."]);
    } catch (error) {
      results.push({
        ...validated,
        status: "Error",
        messages: [
          ...validated.messages,
          error instanceof Error
            ? error.message
            : "Failed to import candidate.",
        ],
      });
    }
  }

  return results;
}

export function buildValidationReportCsv(rows: BulkPreviewRow[]): string {
  const header = [
    "Row",
    "Status",
    "Candidate Name",
    "Company",
    "Workforce Number",
    "DOB",
    "Department",
    "Matched Id",
    "Messages",
  ];
  const lines = [header.map(escapeCsv).join(",")];
  for (const row of rows) {
    lines.push(
      [
        String(row.rowNumber),
        row.status,
        row.fields.candidateName ?? "",
        row.fields.company ?? "",
        row.fields.workforceNumber ?? "",
        row.fields.dateOfBirth ?? "",
        row.fields.department ?? "",
        row.matchedEntityId ?? "",
        row.messages.join("; "),
      ]
        .map(escapeCsv)
        .join(","),
    );
  }
  return `${lines.join("\r\n")}\r\n`;
}

function escapeCsv(value: string): string {
  if (/[",\r\n]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}
