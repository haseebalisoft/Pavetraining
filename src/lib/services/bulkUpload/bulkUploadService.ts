import "server-only";

import {
  buildValidationReportCsv,
  commitCandidateImport,
  previewCandidateImport,
  summarizeBulkRows,
} from "@/lib/services/bulkUpload/candidateImporter";
import {
  commitCompanyImport,
  previewCompanyImport,
} from "@/lib/services/bulkUpload/companyImporter";
import {
  commitMatrixImport,
  previewMatrixImport,
} from "@/lib/services/bulkUpload/matrixImporter";
import {
  commitRegisterImport,
  isRegisterImportType,
  previewRegisterImport,
} from "@/lib/services/bulkUpload/registerImporter";
import { parseSpreadsheetBuffer } from "@/lib/services/bulkUpload/parseSpreadsheet";
import {
  BULK_IMPORT_TEMPLATES,
  buildTemplateCsv,
  getBulkImportTemplate,
} from "@/lib/services/bulkUpload/templates";
import { ValidationError } from "@/lib/services/validationService";
import type {
  BulkCommitResult,
  BulkCommitRowInput,
  BulkDuplicateMode,
  BulkImportType,
  BulkPreviewResult,
} from "@/types/bulkUpload";

const MAX_BULK_UPLOAD_BYTES = 15 * 1024 * 1024;

export { BULK_IMPORT_TEMPLATES, buildTemplateCsv, buildValidationReportCsv };

function parseImportType(value: string | null | undefined): BulkImportType {
  const allowed: BulkImportType[] = [
    "company",
    "workforce",
    "trainingMatrix",
    "npors",
    "eusr",
    "streetworks",
    "inHouse",
    "nvq",
  ];
  const normalized = (value ?? "").trim() as BulkImportType;
  if (!allowed.includes(normalized)) {
    throw new ValidationError("Unsupported import type.");
  }
  return normalized;
}

function parseDuplicateMode(
  value: string | null | undefined,
): BulkDuplicateMode {
  const normalized = (value ?? "skip").trim().toLowerCase();
  if (normalized === "update" || normalized === "create" || normalized === "skip") {
    return normalized;
  }
  throw new ValidationError(
    "Duplicate mode must be skip, update, or create.",
  );
}

function asBool(value: unknown, fallback: boolean): boolean {
  if (value === null || value === undefined || value === "") return fallback;
  if (typeof value === "boolean") return value;
  const text = String(value).trim().toLowerCase();
  if (["true", "1", "yes"].includes(text)) return true;
  if (["false", "0", "no"].includes(text)) return false;
  return fallback;
}

async function readUploadFile(file: File): Promise<{
  fileName: string;
  bytes: Uint8Array;
}> {
  if (!(file instanceof File)) {
    throw new ValidationError("File is required.");
  }
  if (file.size <= 0) {
    throw new ValidationError("Uploaded file is empty.");
  }
  if (file.size > MAX_BULK_UPLOAD_BYTES) {
    throw new ValidationError(
      `File exceeds the ${Math.floor(MAX_BULK_UPLOAD_BYTES / (1024 * 1024))} MB limit.`,
    );
  }
  const fileName = file.name?.trim() || "upload.csv";
  const lower = fileName.toLowerCase();
  if (!lower.endsWith(".xlsx") && !lower.endsWith(".csv")) {
    throw new ValidationError("Only .xlsx and .csv files are supported.");
  }
  return {
    fileName,
    bytes: new Uint8Array(await file.arrayBuffer()),
  };
}

export async function previewBulkUpload(input: {
  importType: string;
  file: File;
  suppressNotifications?: boolean;
}): Promise<BulkPreviewResult> {
  const importType = parseImportType(input.importType);
  const template = getBulkImportTemplate(importType);
  const { fileName, bytes } = await readUploadFile(input.file);
  const suppressNotifications = input.suppressNotifications ?? true;

  if (!template?.implemented) {
    return {
      importType,
      fileName,
      headers: template?.columns.map((c) => c.label) ?? [],
      rows: [],
      summary: {
        totalRows: 0,
        readyRows: 0,
        warningRows: 0,
        duplicateRows: 0,
        errorRows: 0,
        skippedRows: 0,
        importedRows: 0,
      },
      suppressNotifications,
      implemented: false,
      message: `${template?.label ?? importType} import is not available yet. Download the template for the planned columns.`,
    };
  }

  const spreadsheet = parseSpreadsheetBuffer(bytes, fileName);

  if (importType === "company") {
    const rows = await previewCompanyImport(spreadsheet);
    return {
      importType,
      fileName,
      headers: spreadsheet.headers,
      rows,
      summary: summarizeBulkRows(rows),
      suppressNotifications,
      implemented: true,
      message: null,
    };
  }

  if (importType === "workforce") {
    const rows = await previewCandidateImport(spreadsheet);
    return {
      importType,
      fileName,
      headers: spreadsheet.headers,
      rows,
      summary: summarizeBulkRows(rows),
      suppressNotifications,
      implemented: true,
      message: null,
    };
  }

  if (importType === "trainingMatrix") {
    const rows = await previewMatrixImport(spreadsheet);
    return {
      importType,
      fileName,
      headers: spreadsheet.headers,
      rows,
      summary: summarizeBulkRows(rows),
      suppressNotifications,
      implemented: true,
      message: null,
    };
  }

  if (isRegisterImportType(importType)) {
    const rows = await previewRegisterImport(importType, spreadsheet);
    return {
      importType,
      fileName,
      headers: spreadsheet.headers,
      rows,
      summary: summarizeBulkRows(rows),
      suppressNotifications,
      implemented: true,
      message:
        importType === "inHouse"
          ? "In-House: use Course = Asbestos Awareness (or N031) with Outcome Pass + Expiry to update Training Matrix N031. Other courses stay on the In-House register only."
          : importType === "nvq"
            ? "NVQ import is standalone and will not update the Training Matrix."
            : "NPORS / EUSR / Streetworks imports sync Pass expiry dates into the Training Matrix after each row.",
    };
  }

  throw new ValidationError("Import type is not implemented.");
}

export async function commitBulkUpload(input: {
  importType: string;
  fileName?: string | null;
  duplicateMode?: string | null;
  suppressNotifications?: boolean;
  rows: BulkCommitRowInput[];
}): Promise<BulkCommitResult> {
  const importType = parseImportType(input.importType);
  const template = getBulkImportTemplate(importType);
  const duplicateMode = parseDuplicateMode(input.duplicateMode);
  const suppressNotifications = asBool(input.suppressNotifications, true);
  const fileName = input.fileName?.trim() || "upload.csv";

  if (!template?.implemented) {
    throw new ValidationError(
      `${template?.label ?? importType} import is not available yet.`,
    );
  }

  if (!Array.isArray(input.rows) || input.rows.length === 0) {
    throw new ValidationError("No rows provided for import.");
  }

  if (importType === "company") {
    const rows = await commitCompanyImport({
      rows: input.rows,
      duplicateMode,
    });
    const summary = summarizeBulkRows(rows);
    return {
      importType,
      fileName,
      duplicateMode,
      suppressNotifications,
      rows,
      summary,
      message: "Company import completed.",
    };
  }

  if (importType === "workforce") {
    const rows = await commitCandidateImport({
      rows: input.rows,
      duplicateMode,
    });
    const summary = summarizeBulkRows(rows);
    return {
      importType,
      fileName,
      duplicateMode,
      suppressNotifications,
      rows,
      summary,
      message: suppressNotifications
        ? "Import completed. Customer notifications were suppressed."
        : "Import completed.",
    };
  }

  if (importType === "trainingMatrix") {
    const rows = await commitMatrixImport({
      rows: input.rows,
      duplicateMode,
    });
    const summary = summarizeBulkRows(rows);
    return {
      importType,
      fileName,
      duplicateMode,
      suppressNotifications,
      rows,
      summary,
      message: suppressNotifications
        ? "Import completed. Customer notifications were suppressed. Import Workforce before matrix if candidates are missing."
        : "Import completed.",
    };
  }

  if (isRegisterImportType(importType)) {
    const rows = await commitRegisterImport({
      importType,
      rows: input.rows,
      duplicateMode,
    });
    const summary = summarizeBulkRows(rows);
    return {
      importType,
      fileName,
      duplicateMode,
      suppressNotifications,
      rows,
      summary,
      message:
        importType === "inHouse" || importType === "nvq"
          ? `${template.label} import completed (standalone — Training Matrix not updated).`
          : `${template.label} import completed. Matching Training Matrix rows were synced for Pass outcomes.`,
    };
  }

  throw new ValidationError("Import type is not implemented.");
}

export function listBulkUploadTemplates() {
  return BULK_IMPORT_TEMPLATES.map((row) => ({
    importType: row.importType,
    label: row.label,
    description: row.description,
    fileName: row.fileName,
    columns: row.columns,
    implemented: row.implemented,
  }));
}
