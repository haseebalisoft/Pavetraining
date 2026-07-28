export type BulkImportType =
  | "company"
  | "workforce"
  | "trainingMatrix"
  | "npors"
  | "eusr"
  | "streetworks"
  | "inHouse"
  | "nvq";

export type BulkRowStatus =
  | "Ready"
  | "Warning"
  | "Duplicate"
  | "Error"
  | "Skipped"
  | "Imported";

/** Default: skip duplicates. Update/create require admin confirmation. */
export type BulkDuplicateMode = "skip" | "update" | "create";

export type BulkImportColumn = {
  key: string;
  label: string;
  required?: boolean;
};

export type BulkImportTemplate = {
  importType: BulkImportType;
  label: string;
  description: string;
  fileName: string;
  columns: BulkImportColumn[];
  implemented: boolean;
};

export type BulkCandidateFields = {
  candidateName: string;
  company: string;
  department: string | null;
  dateOfBirth: string | null;
  workforceNumber: string | null;
  status: string | null;
  trainingManager: string | null;
  supervisor: string | null;
};

export type BulkPreviewRow = {
  rowNumber: number;
  status: BulkRowStatus;
  messages: string[];
  fields: Record<string, string | null>;
  /** Original spreadsheet cells keyed by Excel header (exact template columns). */
  source?: Record<string, string | null>;
  /** Resolved SharePoint company display name when known. */
  resolvedCompanyName?: string | null;
  /** Matched existing entity id when duplicate detected. */
  matchedEntityId?: string | null;
  matchedEntityName?: string | null;
  duplicateMatch?: "workforceNumber" | "nameDobCompany" | "nameCompany" | null;
};

export type BulkImportSummary = {
  totalRows: number;
  readyRows: number;
  warningRows: number;
  duplicateRows: number;
  errorRows: number;
  skippedRows: number;
  importedRows: number;
};

export type BulkPreviewResult = {
  importType: BulkImportType;
  fileName: string;
  /** Exact header row from the uploaded spreadsheet. */
  headers: string[];
  rows: BulkPreviewRow[];
  summary: BulkImportSummary;
  suppressNotifications: boolean;
  implemented: boolean;
  message?: string | null;
};

export type BulkCommitResult = {
  importType: BulkImportType;
  fileName: string;
  duplicateMode: BulkDuplicateMode;
  suppressNotifications: boolean;
  rows: BulkPreviewRow[];
  summary: BulkImportSummary;
  message?: string | null;
};

export type BulkCommitRowInput = {
  rowNumber: number;
  fields: Record<string, string | null>;
  /** Original spreadsheet cells — required for Training Matrix N-code columns. */
  source?: Record<string, string | null>;
};
