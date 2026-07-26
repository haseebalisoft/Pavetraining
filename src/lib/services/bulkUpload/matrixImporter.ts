import "server-only";

import {
  createAdminMatrix,
  listAdminMatrix,
  listAdminWorkforce,
  updateAdminMatrix,
  type AdminMatrixRecord,
  type AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
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

function asBoolField(value: string | null): boolean | null {
  if (!value?.trim()) return null;
  const normalized = value.trim().toLowerCase();
  if (["true", "yes", "1"].includes(normalized)) return true;
  if (["false", "no", "0"].includes(normalized)) return false;
  return null;
}

function mapMatrixFields(
  raw: Record<string, string | null>,
): Record<string, string | null> {
  return {
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
    department: pickField(raw, ["Department", "Dept"]),
    dateOfBirth: normalizeDateValue(
      pickField(raw, ["DOB", "Date of birth", "DateOfBirth"]),
    ),
    overallStatus: pickField(raw, ["Overall Status", "OverallStatus", "Status"]),
    needsReview: pickField(raw, ["Needs Review", "NeedsReview"]),
    matrixNotes: pickField(raw, ["Matrix Notes", "MatrixNotes", "Notes"]),
    nextExpiryDate: normalizeDateValue(
      pickField(raw, ["Next Expiry Date", "NextExpiryDate"]),
    ),
    n001Expiry: normalizeDateValue(pickField(raw, ["N001 Expiry", "N001Expiry"])),
    n003Expiry: normalizeDateValue(pickField(raw, ["N003 Expiry", "N003Expiry"])),
    n004Expiry: normalizeDateValue(pickField(raw, ["N004 Expiry", "N004Expiry"])),
    n010Expiry: normalizeDateValue(pickField(raw, ["N010 Expiry", "N010Expiry"])),
    n020Expiry: normalizeDateValue(pickField(raw, ["N020 Expiry", "N020Expiry"])),
    n021Expiry: normalizeDateValue(pickField(raw, ["N021 Expiry", "N021Expiry"])),
    n027Expiry: normalizeDateValue(pickField(raw, ["N027 Expiry", "N027Expiry"])),
    n100Expiry: normalizeDateValue(pickField(raw, ["N100 Expiry", "N100Expiry"])),
  };
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

function validateMatrixRow(
  rowNumber: number,
  fields: Record<string, string | null>,
  workforce: AdminWorkforceRecord[],
  matrix: AdminMatrixRecord[],
): BulkPreviewRow {
  const messages: string[] = [];
  if (!fields.candidateName?.trim() && !fields.workforceNumber?.trim()) {
    messages.push("Candidate Name or Workforce Number is required.");
  }

  const candidate = findWorkforceForMatrix(workforce, fields);
  if (!candidate) {
    messages.push(
      "Candidate was not found in Workforce. Import Workforce first, then matrix rows.",
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
  };

  const existing = findMatrixDuplicate(matrix, candidate);
  if (existing) {
    return {
      rowNumber,
      status: "Duplicate",
      messages: [
        `Duplicate matrix row for "${candidate.candidateName}" at ${resolvedCompanyName}.`,
      ],
      fields: enrichedFields,
      resolvedCompanyName,
      matchedEntityId: existing.id,
      matchedEntityName: existing.candidateName,
      duplicateMatch: "nameCompany",
    };
  }

  if (!fields.overallStatus) {
    messages.push("Overall Status is missing (optional).");
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
  const needsReview = asBoolField(fields.needsReview);
  const payload: Record<string, unknown> = {
    candidateName: fields.candidateName,
    companyName: fields.company,
  };
  if (fields.department?.trim()) payload.department = fields.department.trim();
  if (fields.overallStatus?.trim()) {
    payload.overallStatus = fields.overallStatus.trim();
  }
  if (needsReview !== null) payload.needsReview = needsReview;
  if (fields.matrixNotes?.trim()) payload.matrixNotes = fields.matrixNotes.trim();
  if (fields.nextExpiryDate?.trim()) {
    payload.nextExpiryDate = fields.nextExpiryDate.trim();
  }
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

export async function previewMatrixImport(
  spreadsheet: ParsedSpreadsheet,
): Promise<BulkPreviewRow[]> {
  const [workforce, matrix] = await Promise.all([
    listAdminWorkforce(),
    listAdminMatrix(),
  ]);

  return spreadsheet.rows.map((raw, index) => {
    const fields = mapMatrixFields(raw);
    return validateMatrixRow(index + 2, fields, workforce, matrix);
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
    // Prefer already-normalized fields from preview when present.
    for (const [key, value] of Object.entries(row.fields)) {
      if (value && !fields[key]) fields[key] = value;
      if (value && ["candidateName", "company", "workforceNumber"].includes(key)) {
        fields[key] = value;
      }
    }
    fields.dateOfBirth = normalizeDateValue(fields.dateOfBirth);
    for (const key of [
      "nextExpiryDate",
      "n001Expiry",
      "n003Expiry",
      "n004Expiry",
      "n010Expiry",
      "n020Expiry",
      "n021Expiry",
      "n027Expiry",
      "n100Expiry",
    ] as const) {
      fields[key] = normalizeDateValue(fields[key]);
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
          const updated = await updateAdminMatrix(
            validated.matchedEntityId,
            buildMatrixWritePayload(validated.fields),
          );
          const idx = liveMatrix.findIndex((m) => m.id === updated.id);
          if (idx >= 0) liveMatrix[idx] = updated;
          results.push({
            ...validated,
            status: "Imported",
            messages: [...validated.messages, "Updated existing matrix row."],
          });
        } catch (error) {
          results.push({
            ...validated,
            status: "Error",
            messages: [
              ...validated.messages,
              error instanceof Error
                ? error.message
                : "Failed to update matrix row.",
            ],
          });
        }
        continue;
      }

      if (input.duplicateMode === "create") {
        try {
          const created = await createAdminMatrix(
            buildMatrixWritePayload(validated.fields),
          );
          liveMatrix.push(created);
          results.push({
            ...validated,
            status: "Imported",
            matchedEntityId: created.id,
            messages: [
              ...validated.messages,
              "Created new matrix row (admin confirmed create despite duplicate).",
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
                : "Failed to create matrix row.",
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

    try {
      const created = await createAdminMatrix(
        buildMatrixWritePayload(validated.fields),
      );
      liveMatrix.push(created);
      results.push({
        ...validated,
        status: "Imported",
        matchedEntityId: created.id,
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
            : "Failed to create matrix row.",
        ],
      });
    }
  }

  return results;
}
