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
  /**
   * Optional email of the Training Manager. When present in the spreadsheet,
   * the importer will (a) try to match this email against existing Permissions
   * rows first, and (b) use it as the real UserEmail when auto-creating a new
   * Permissions row. When omitted, a pending.{tm|sp}.{name}@pave.local
   * placeholder is used so the person still appears on the Permissions list.
   */
  trainingManagerEmail: string | null;
  supervisor: string | null;
  /** Same as trainingManagerEmail, for the Supervisor column. */
  supervisorEmail: string | null;
};

/**
 * How a row resolved against the Workforce ↔ Training Matrix link. Reported per
 * row in the preview and the commit result so an admin can see, before and
 * after importing, exactly which matrix row each candidate attached to.
 */
export type BulkLinkOutcome =
  | "linkedCompanyNameDob"
  | "linkedNameDob"
  | "linkedExistingNeedsReview"
  | "createdLinked"
  | "needsReviewNoMatch"
  | "needsReviewMultipleMatches"
  | "skippedAmbiguous";

export const BULK_LINK_OUTCOME_LABELS: Record<BulkLinkOutcome, string> = {
  linkedCompanyNameDob: "Linked by Company + Name + DOB",
  linkedNameDob: "Linked by Name + DOB",
  linkedExistingNeedsReview: "Existing Needs Review row will be linked",
  createdLinked: "Matrix created new linked row",
  needsReviewNoMatch: "Needs Review - no Workforce match",
  needsReviewMultipleMatches: "Needs Review - multiple matches",
  skippedAmbiguous: "Matrix skipped due to ambiguous match",
};

/**
 * Document-folder outcome for a bulk-imported Workforce row. The app is the
 * sole folder-creation owner (no Power Automate step in this flow), so every
 * row resolves synchronously to one of these — there is no "Pending" state.
 */
export type BulkFolderOutcome = "folderCreated" | "folderExisted" | "folderFailed";

export const BULK_FOLDER_OUTCOME_LABELS: Record<BulkFolderOutcome, string> = {
  folderCreated: "Folder created",
  folderExisted: "Folder already existed",
  folderFailed: "Folder creation failed",
};

export type BulkPreviewRow = {
  rowNumber: number;
  status: BulkRowStatus;
  messages: string[];
  /** Workforce ↔ Training Matrix link result for this row. */
  linkOutcome?: BulkLinkOutcome | null;
  /** Matrix row this candidate linked to (`example:<id>`), once known. */
  matrixRowId?: string | null;
  /** Document-folder result for this candidate (Workforce bulk import only). */
  folderOutcome?: BulkFolderOutcome | null;
  fields: Record<string, string | null>;
  /** Original spreadsheet cells keyed by Excel header (exact template columns). */
  source?: Record<string, string | null>;
  /** Resolved SharePoint company display name when known. */
  resolvedCompanyName?: string | null;
  /** Matched existing entity id when duplicate detected. */
  matchedEntityId?: string | null;
  matchedEntityName?: string | null;
  duplicateMatch?: "workforceNumber" | "nameDobCompany" | "nameCompany" | null;
  /** Company-match report (workforce import): which Company List item the row resolved to. */
  matchedCompanyId?: string | null;
  matchedCompanyNumber?: string | null;
  matchedCompanyName?: string | null;
  /** How the company was matched, or "create" when a new company will be made. */
  companyMatchedBy?: "companyNumber" | "companyName" | "create" | null;
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
