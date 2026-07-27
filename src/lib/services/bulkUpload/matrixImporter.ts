import "server-only";

import {
  createAdminMatrix,
  listAdminMatrix,
  listAdminWorkforce,
  updateAdminMatrix,
  type AdminMatrixRecord,
  type AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
import { CLIENT_MATRIX_CATEGORY_COLUMNS } from "@/lib/services/bulkUpload/clientTemplateHeaders";
import {
  extractCategoryWritesFromRow,
  upsertMatrixCategoryRecords,
} from "@/lib/services/bulkUpload/matrixCategoryService";
import { nameKey, normalizeCompanyKey } from "@/lib/services/bulkUpload/matching";
import {
  normalizeDateValue,
  pickField,
  type ParsedSpreadsheet,
} from "@/lib/services/bulkUpload/parseSpreadsheet";
import type {
  BulkCommitRowInput,
  BulkDuplicateMode,
  BulkPreviewRow,
} from "@/types/bulkUpload";

function mapMatrixFields(
  raw: Record<string, string | null>,
): Record<string, string | null> {
  const fields: Record<string, string | null> = {
    candidateName: pickField(raw, [
      "Candidate Name",
      "CandidateName",
      "Name",
    ]),
    workforceNumber: pickField(raw, [
      "Workforce Number",
      "WorkforceNumber",
      "Workforce No",
    ]),
    company: pickField(raw, ["Company", "Company Name", "CompanyName"]),
    department: pickField(raw, ["Department", "Dept", " Department"]),
    dateOfBirth: normalizeDateValue(
      pickField(raw, ["DOB", "Date of birth", "DateOfBirth"]),
    ),
    overallStatus: pickField(raw, ["Overall Status", "OverallStatus", "Status"]),
    needsReview: pickField(raw, ["Needs Review", "NeedsReview"]),
    matrixNotes: pickField(raw, ["Matrix Notes", "MatrixNotes", "Notes"]),
    nextExpiryDate: normalizeDateValue(
      pickField(raw, ["Next Expiry Date", "NextExpiryDate"]),
    ),
    cscsExpiry: normalizeDateValue(
      pickField(raw, ["CSCS Expiry", "Cscs Expiry"]),
    ),
    ssstsExpiry: normalizeDateValue(pickField(raw, ["SSSTS Expiry"])),
    smstsExpiry: normalizeDateValue(pickField(raw, ["SMSTS Expiry"])),
    nrswaExpiry: normalizeDateValue(pickField(raw, ["NRSWA Expiry"])),
    eusrExpiry: normalizeDateValue(
      pickField(raw, ["EUSR Expiry", "Eusr Expiry"]),
    ),
    faceFitExpiry: normalizeDateValue(
      pickField(raw, ["Face ift", "Face Fit", "FaceFit"]),
    ),
  };

  // Map known Training Matrix list expiry columns from client N-code headers.
  for (const column of CLIENT_MATRIX_CATEGORY_COLUMNS) {
    if (!column.matrixField) continue;
    const value = normalizeDateValue(
      pickField(raw, [column.header, `${column.code} Expiry`, column.code]),
    );
    if (value) fields[column.matrixField] = value;
  }

  // Count filled category cells for preview messaging.
  const categoryWrites = extractCategoryWritesFromRow(raw);
  fields._categoryCount = String(categoryWrites.length);

  return fields;
}

function findWorkforceForMatrix(
  workforce: AdminWorkforceRecord[],
  fields: Record<string, string | null>,
): AdminWorkforceRecord | null {
  if (fields.workforceNumber?.trim()) {
    const numberKey = nameKey(fields.workforceNumber);
    const byNumber = workforce.find(
      (row) => nameKey(row.workforceNumber) === numberKey,
    );
    if (byNumber) return byNumber;
  }

  if (!fields.candidateName?.trim()) return null;
  const cName = nameKey(fields.candidateName);
  const companyHint = fields.company?.trim()
    ? nameKey(fields.company)
    : null;
  const companyNorm = fields.company?.trim()
    ? normalizeCompanyKey(fields.company)
    : null;

  if (companyHint) {
    const byNameCompany = workforce.find(
      (row) =>
        nameKey(row.candidateName) === cName &&
        (nameKey(row.companyName) === companyHint ||
          normalizeCompanyKey(row.companyName) === companyNorm),
    );
    if (byNameCompany) return byNameCompany;
  }

  // Client matrix template often has Name + DOB only.
  if (fields.dateOfBirth?.trim()) {
    const byNameDob = workforce.filter(
      (row) =>
        nameKey(row.candidateName) === cName &&
        nameKey(row.dateOfBirth).slice(0, 10) ===
          nameKey(fields.dateOfBirth).slice(0, 10),
    );
    if (byNameDob.length === 1) return byNameDob[0]!;
  }

  const matches = workforce.filter((row) => nameKey(row.candidateName) === cName);
  return matches.length === 1 ? matches[0]! : null;
}

function findMatrixDuplicate(
  matrix: AdminMatrixRecord[],
  candidate: AdminWorkforceRecord,
): AdminMatrixRecord | null {
  const cName = nameKey(candidate.candidateName);
  const coName = nameKey(candidate.companyName);
  const coNorm = normalizeCompanyKey(candidate.companyName);
  return (
    matrix.find(
      (row) =>
        nameKey(row.candidateName) === cName &&
        (nameKey(row.companyName) === coName ||
          normalizeCompanyKey(row.companyName) === coNorm),
    ) ?? null
  );
}

function earliestExpiry(fields: Record<string, string | null>): string | null {
  const dates: number[] = [];
  for (const [key, value] of Object.entries(fields)) {
    if (!value?.trim()) continue;
    if (
      !/expiry$/i.test(key) &&
      !["cscsExpiry", "ssstsExpiry", "smstsExpiry", "nrswaExpiry", "eusrExpiry", "faceFitExpiry", "nextExpiryDate"].includes(
        key,
      )
    ) {
      continue;
    }
    const t = new Date(value).getTime();
    if (!Number.isNaN(t)) dates.push(t);
  }
  if (!dates.length) return null;
  return new Date(Math.min(...dates)).toISOString().slice(0, 10);
}

function validateMatrixRow(
  rowNumber: number,
  fields: Record<string, string | null>,
  workforce: AdminWorkforceRecord[],
  matrix: AdminMatrixRecord[],
): BulkPreviewRow {
  const messages: string[] = [];
  if (!fields.candidateName?.trim() && !fields.workforceNumber?.trim()) {
    messages.push("Name (Candidate) or Workforce Number is required.");
  }

  const candidate = findWorkforceForMatrix(workforce, fields);
  if (!candidate) {
    messages.push(
      "Candidate was not found in Workforce. Import Workforce first (same Name / DOB), then matrix rows.",
    );
    return {
      rowNumber,
      status: "Error",
      messages,
      fields,
      resolvedCompanyName: fields.company,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
    };
  }

  const resolvedCompanyName = candidate.companyName;
  const enrichedFields = {
    ...fields,
    candidateName: candidate.candidateName,
    company: resolvedCompanyName,
    workforceNumber: candidate.workforceNumber,
    department: fields.department ?? candidate.department,
    nextExpiryDate:
      fields.nextExpiryDate ?? earliestExpiry(fields),
  };

  const categoryCount = Number(fields._categoryCount ?? "0");
  if (categoryCount > 0) {
    messages.push(
      `${categoryCount} category expiry value(s) will be written to Training Matrix Category Records.`,
    );
  }

  const existing = findMatrixDuplicate(matrix, candidate);
  if (existing) {
    return {
      rowNumber,
      status: "Duplicate",
      messages: [
        `Duplicate matrix row for "${candidate.candidateName}" at ${resolvedCompanyName}.`,
        ...messages,
      ],
      fields: enrichedFields,
      resolvedCompanyName,
      matchedEntityId: existing.id,
      matchedEntityName: existing.candidateName,
      duplicateMatch: "nameCompany",
    };
  }

  return {
    rowNumber,
    status: messages.length ? "Warning" : "Ready",
    messages,
    fields: enrichedFields,
    resolvedCompanyName,
    matchedEntityId: candidate.id,
    matchedEntityName: candidate.candidateName,
    duplicateMatch: null,
  };
}

function buildMatrixWritePayload(fields: Record<string, string | null>) {
  const payload: Record<string, unknown> = {
    candidateName: fields.candidateName,
    companyName: fields.company,
  };
  if (fields.department?.trim()) payload.department = fields.department.trim();
  if (fields.overallStatus?.trim()) {
    payload.overallStatus = fields.overallStatus.trim();
  }
  if (fields.matrixNotes?.trim()) payload.matrixNotes = fields.matrixNotes.trim();
  const next = fields.nextExpiryDate ?? earliestExpiry(fields);
  if (next) payload.nextExpiryDate = next;
  for (const key of [
    "n001Expiry",
    "n003Expiry",
    "n004Expiry",
    "n010Expiry",
    "n020Expiry",
    "n021Expiry",
    "n027Expiry",
    "n100Expiry",
  ] as const) {
    if (fields[key]?.trim()) payload[key] = fields[key]!.trim();
  }
  return payload;
}

async function writeCategoriesForRow(
  fields: Record<string, string | null>,
  raw: Record<string, string | null>,
): Promise<string[]> {
  const categories = extractCategoryWritesFromRow({ ...raw, ...fields });
  if (!categories.length) return [];
  const result = await upsertMatrixCategoryRecords({
    candidateName: fields.candidateName ?? "",
    companyName: fields.company ?? "",
    categories,
  });
  const messages = [
    `Category records: ${result.written} written` +
      (result.failed ? `, ${result.failed} failed` : ""),
  ];
  if (result.errors.length) {
    messages.push(...result.errors.slice(0, 5));
  }
  return messages;
}

export async function previewMatrixImport(
  spreadsheet: ParsedSpreadsheet,
): Promise<BulkPreviewRow[]> {
  const [workforce, matrix] = await Promise.all([
    listAdminWorkforce(),
    listAdminMatrix(),
  ]);

  return spreadsheet.rows.map((raw, index) => {
    const fields = mapMatrixFields(raw);
    const validated = validateMatrixRow(index + 2, fields, workforce, matrix);
    return {
      ...validated,
      source: raw,
    };
  });
}

export async function commitMatrixImport(input: {
  rows: BulkCommitRowInput[];
  duplicateMode: BulkDuplicateMode;
}): Promise<BulkPreviewRow[]> {
  const [workforce, matrix] = await Promise.all([
    listAdminWorkforce(),
    listAdminMatrix(),
  ]);
  const liveMatrix = [...matrix];
  const results: BulkPreviewRow[] = [];

  for (const row of input.rows) {
    const fields = mapMatrixFields(row.fields);
    for (const [key, value] of Object.entries(row.fields)) {
      if (value && !fields[key]) fields[key] = value;
      if (
        value &&
        ["candidateName", "company", "workforceNumber", "dateOfBirth"].includes(
          key,
        )
      ) {
        fields[key] = value;
      }
    }
    fields.dateOfBirth = normalizeDateValue(fields.dateOfBirth);
    for (const key of Object.keys(fields)) {
      if (/expiry$/i.test(key) || key === "nextExpiryDate") {
        fields[key] = normalizeDateValue(fields[key]);
      }
    }

    const validated = validateMatrixRow(
      row.rowNumber,
      fields,
      workforce,
      liveMatrix,
    );

    if (validated.status === "Error") {
      results.push(validated);
      continue;
    }

    try {
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
          const updated = await updateAdminMatrix(
            validated.matchedEntityId,
            buildMatrixWritePayload(validated.fields),
          );
          const idx = liveMatrix.findIndex((m) => m.id === updated.id);
          if (idx >= 0) liveMatrix[idx] = updated;
          const catMessages = await writeCategoriesForRow(
            validated.fields,
            row.fields,
          );
          results.push({
            ...validated,
            status: "Imported",
            messages: [
              ...validated.messages,
              "Updated existing matrix row.",
              ...catMessages,
            ],
          });
          continue;
        }

        if (input.duplicateMode === "create") {
          const created = await createAdminMatrix(
            buildMatrixWritePayload(validated.fields),
          );
          liveMatrix.push(created);
          const catMessages = await writeCategoriesForRow(
            validated.fields,
            row.fields,
          );
          results.push({
            ...validated,
            status: "Imported",
            matchedEntityId: created.id,
            messages: [
              ...validated.messages,
              "Created new matrix row (admin confirmed create despite duplicate).",
              ...catMessages,
            ],
          });
          continue;
        }

        results.push({
          ...validated,
          status: "Skipped",
          messages: [...validated.messages, "Skipped duplicate."],
        });
        continue;
      }

      const created = await createAdminMatrix(
        buildMatrixWritePayload(validated.fields),
      );
      liveMatrix.push(created);
      const catMessages = await writeCategoriesForRow(
        validated.fields,
        row.fields,
      );
      results.push({
        ...validated,
        status: "Imported",
        matchedEntityId: created.id,
        messages: [
          ...validated.messages,
          "Imported successfully.",
          ...catMessages,
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
            : "Failed to import matrix row.",
        ],
      });
    }
  }

  return results;
}
