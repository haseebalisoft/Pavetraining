import "server-only";

import {
  createAdminWorkforce,
  listAdminCompanies,
  listAdminWorkforce,
  updateAdminWorkforce,
  type AdminWorkforceRecord,
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

const NAME_ALIASES = [
  "Candidate Name",
  "CandidateName",
  "Name",
  "Full Name",
];
const COMPANY_ALIASES = ["Company", "Company Name", "CompanyName"];
const DEPT_ALIASES = ["Department", "Dept"];
const DOB_ALIASES = ["DOB", "Date of Birth", "DateOfBirth", "Birth Date"];
const WF_ALIASES = [
  "Workforce Number",
  "WorkforceNumber",
  "Workforce No",
  "Workforce No.",
];
const STATUS_ALIASES = ["Status"];
const TM_ALIASES = ["Training Manager", "TrainingManager"];
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
    department: pickField(raw, DEPT_ALIASES),
    dateOfBirth: normalizeDateValue(pickField(raw, DOB_ALIASES)),
    workforceNumber: pickField(raw, WF_ALIASES),
    status: pickField(raw, STATUS_ALIASES),
    trainingManager: pickField(raw, TM_ALIASES),
    supervisor: pickField(raw, SUPERVISOR_ALIASES),
  };
}

function validateCandidateRow(
  rowNumber: number,
  fields: Record<string, string | null>,
  companies: Awaited<ReturnType<typeof listAdminCompanies>>,
  workforce: AdminWorkforceRecord[],
): BulkPreviewRow {
  const messages: string[] = [];
  const candidateName = fields.candidateName?.trim() ?? "";
  const companyInput = fields.company?.trim() ?? "";

  if (!candidateName) {
    messages.push("Candidate Name is required.");
  }
  if (!companyInput) {
    messages.push("Company is required.");
  }

  const company = findCompanyByName(companies, companyInput);
  if (companyInput && !company) {
    messages.push(`Company "${companyInput}" was not found.`);
  }

  if (messages.length) {
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

  const resolvedCompanyName = company!.companyName;
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

  if (!fields.workforceNumber) {
    messages.push("Workforce Number is missing (optional).");
  }
  if (!fields.dateOfBirth) {
    messages.push("DOB is missing (optional).");
  }
  if (!fields.department) {
    messages.push("Department is missing (optional).");
  }

  const exactCompany =
    company &&
    company.companyName.trim().toLowerCase() === companyInput.toLowerCase();
  if (!exactCompany) {
    messages.push(
      `Company matched as "${resolvedCompanyName}" (normalized name match).`,
    );
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

export async function previewCandidateImport(
  spreadsheet: ParsedSpreadsheet,
): Promise<BulkPreviewRow[]> {
  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);

  return spreadsheet.rows.map((raw, index) => {
    const fields = mapCandidateFields(raw);
    return validateCandidateRow(index + 2, fields, companies, workforce);
  });
}

/**
 * Build update payload that never overwrites existing fields with blanks.
 */
function buildNonBlankUpdate(
  fields: Record<string, string | null>,
  resolvedCompanyName: string,
): Record<string, unknown> {
  const payload: Record<string, unknown> = {};
  if (fields.candidateName?.trim()) {
    payload.candidateName = fields.candidateName.trim();
  }
  if (resolvedCompanyName.trim()) {
    payload.companyName = resolvedCompanyName.trim();
  }
  if (fields.workforceNumber?.trim()) {
    payload.workforceNumber = fields.workforceNumber.trim();
  }
  if (fields.dateOfBirth?.trim()) {
    payload.dateOfBirth = fields.dateOfBirth.trim();
  }
  if (fields.department?.trim()) {
    payload.department = fields.department.trim();
  }
  if (fields.status?.trim()) {
    payload.status = fields.status.trim();
  }
  if (fields.trainingManager?.trim()) {
    payload.trainingManager = fields.trainingManager.trim();
  }
  if (fields.supervisor?.trim()) {
    payload.supervisor = fields.supervisor.trim();
  }
  return payload;
}

export async function commitCandidateImport(input: {
  rows: BulkCommitRowInput[];
  duplicateMode: BulkDuplicateMode;
}): Promise<BulkPreviewRow[]> {
  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);

  // Live copy we update as we create so later rows see earlier imports.
  const liveWorkforce = [...workforce];
  const results: BulkPreviewRow[] = [];

  for (const row of input.rows) {
    const fields = {
      candidateName: row.fields.candidateName ?? null,
      company: row.fields.company ?? null,
      department: row.fields.department ?? null,
      dateOfBirth: normalizeDateValue(row.fields.dateOfBirth ?? null),
      workforceNumber: row.fields.workforceNumber ?? null,
      status: row.fields.status ?? null,
      trainingManager: row.fields.trainingManager ?? null,
      supervisor: row.fields.supervisor ?? null,
    };

    const validated = validateCandidateRow(
      row.rowNumber,
      fields,
      companies,
      liveWorkforce,
    );

    if (validated.status === "Error") {
      results.push({ ...validated, status: "Error" });
      continue;
    }

    // Soft name+company warning: treat as Ready for create unless update mode
    // and matched id present.
    if (validated.status === "Warning" && validated.duplicateMatch === "nameCompany") {
      if (input.duplicateMode === "update" && validated.matchedEntityId) {
        try {
          const payload = buildNonBlankUpdate(
            validated.fields,
            validated.resolvedCompanyName ?? validated.fields.company ?? "",
          );
          const updated = await updateAdminWorkforce(
            validated.matchedEntityId,
            payload,
          );
          const idx = liveWorkforce.findIndex((w) => w.id === updated.id);
          if (idx >= 0) liveWorkforce[idx] = updated;
          results.push({
            ...validated,
            status: "Imported",
            messages: [
              ...validated.messages,
              "Updated existing candidate (name + company match).",
            ],
          });
        } catch (error) {
          results.push({
            ...validated,
            status: "Error",
            messages: [
              ...validated.messages,
              error instanceof Error
                ? error.message
                : "Failed to update candidate.",
            ],
          });
        }
        continue;
      }
      // Default / create: create new despite soft warning
      try {
        const created = await createAdminWorkforce({
          candidateName: validated.fields.candidateName,
          companyName:
            validated.resolvedCompanyName ?? validated.fields.company,
          workforceNumber: validated.fields.workforceNumber,
          dateOfBirth: validated.fields.dateOfBirth,
          department: validated.fields.department,
          status: validated.fields.status ?? "Active",
          trainingManager: validated.fields.trainingManager,
          supervisor: validated.fields.supervisor,
        });
        liveWorkforce.push(created);
        results.push({
          ...validated,
          status: "Imported",
          matchedEntityId: created.id,
          matchedEntityName: created.candidateName,
          messages: [
            ...validated.messages,
            "Created new candidate (soft name match was a warning only).",
          ],
        });
      } catch (error) {
        results.push({
          ...validated,
          status: "Error",
          messages: [
            ...validated.messages,
            error instanceof Error
              ? error.message
              : "Failed to create candidate.",
          ],
        });
      }
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
        try {
          const payload = buildNonBlankUpdate(
            validated.fields,
            validated.resolvedCompanyName ?? validated.fields.company ?? "",
          );
          const updated = await updateAdminWorkforce(
            validated.matchedEntityId,
            payload,
          );
          const idx = liveWorkforce.findIndex((w) => w.id === updated.id);
          if (idx >= 0) liveWorkforce[idx] = updated;
          results.push({
            ...validated,
            status: "Imported",
            messages: [...validated.messages, "Updated existing candidate."],
          });
        } catch (error) {
          results.push({
            ...validated,
            status: "Error",
            messages: [
              ...validated.messages,
              error instanceof Error
                ? error.message
                : "Failed to update candidate.",
            ],
          });
        }
        continue;
      }

      if (input.duplicateMode === "create") {
        try {
          const created = await createAdminWorkforce({
            candidateName: validated.fields.candidateName,
            companyName:
              validated.resolvedCompanyName ?? validated.fields.company,
            workforceNumber: validated.fields.workforceNumber,
            dateOfBirth: validated.fields.dateOfBirth,
            department: validated.fields.department,
            status: validated.fields.status ?? "Active",
            trainingManager: validated.fields.trainingManager,
            supervisor: validated.fields.supervisor,
          });
          liveWorkforce.push(created);
          results.push({
            ...validated,
            status: "Imported",
            matchedEntityId: created.id,
            matchedEntityName: created.candidateName,
            messages: [
              ...validated.messages,
              "Created new candidate (admin confirmed create despite duplicate).",
            ],
          });
        } catch (error) {
          results.push({
            ...validated,
            status: "Error",
            messages: [
              ...validated.messages,
              error instanceof Error
                ? error.message
                : "Failed to create candidate.",
            ],
          });
        }
        continue;
      }

      results.push({
        ...validated,
        status: "Skipped",
        messages: [...validated.messages, "Skipped duplicate."],
      });
      continue;
    }

    // Ready or informational Warning (missing optional fields)
    try {
      const created = await createAdminWorkforce({
        candidateName: validated.fields.candidateName,
        companyName: validated.resolvedCompanyName ?? validated.fields.company,
        workforceNumber: validated.fields.workforceNumber,
        dateOfBirth: validated.fields.dateOfBirth,
        department: validated.fields.department,
        status: validated.fields.status ?? "Active",
        trainingManager: validated.fields.trainingManager,
        supervisor: validated.fields.supervisor,
      });
      liveWorkforce.push(created);
      results.push({
        ...validated,
        status: "Imported",
        matchedEntityId: created.id,
        matchedEntityName: created.candidateName,
        messages: [...validated.messages, "Imported successfully."],
      });
    } catch (error) {
      results.push({
        ...validated,
        status: "Error",
        messages: [
          ...validated.messages,
          error instanceof Error
            ? error.message
            : "Failed to create candidate.",
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
