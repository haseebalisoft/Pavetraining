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
 * Why a Training Matrix row could not be linked to exactly one Workforce
 * record. Shown on Needs Review / Orphan rows in the upload preview, the
 * upload results, and Admin Matrix, so an admin knows what to fix.
 */
export type MatrixReviewReasonCode =
  | "candidateNotFound"
  | "dobMissing"
  | "dobMismatch"
  | "companyMismatch"
  | "multipleMatches"
  | "multipleMatrixRows"
  | "linkedRecordMissing"
  | "duplicateMatrixRow";

export const MATRIX_REVIEW_REASON_LABELS: Record<MatrixReviewReasonCode, string> = {
  candidateNotFound: "Candidate not found",
  dobMissing: "DOB missing",
  dobMismatch: "DOB mismatch",
  companyMismatch: "Company mismatch",
  multipleMatches: "Multiple matches",
  multipleMatrixRows: "Multiple matrix rows",
  linkedRecordMissing: "Linked record deleted",
  duplicateMatrixRow: "Duplicate matrix row",
};

export type MatrixReviewReason = {
  code: MatrixReviewReasonCode;
  /** Short label, e.g. "DOB mismatch". */
  label: string;
  /** Full explanation naming the records compared and what to do next. */
  detail: string;
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
  /** Why a Training Matrix row was left Needs Review (matrix imports only). */
  reviewReason?: MatrixReviewReason | null;
  /** Matrix row this candidate linked to (`example:<id>`), once known. */
  matrixRowId?: string | null;
  /** Document-folder result for this candidate (Workforce bulk import only). */
  folderOutcome?: BulkFolderOutcome | null;
  fields: Record<string, string | null>;
  /** Original spreadsheet cells keyed by Excel header (exact template columns). */
  source?: Record<string, string | null>;
  /** Resolved SharePoint company display name when known. */
  resolvedCompanyName?: string | null;
  /** True when the commit stopped (time budget) before reaching this row. */
  notProcessed?: boolean;
  /**
   * Matrix row saved, but its category backup records were not written
   * before the time budget ran out (re-upload finishes them).
   */
  backupPending?: boolean;
  /** Matched existing entity id when duplicate detected. */
  matchedEntityId?: string | null;
  /** SharePoint item id the commit created/updated (company imports). */
  savedEntityId?: string | null;
  /**
   * Company imports: logo file name from the Company Logo column that the
   * browser must upload from the attached files once the row exists.
   */
  pendingLogoFile?: string | null;
  /**
   * Workforce imports: photo file name from the Candidate Photo column that
   * the browser must upload from the attached files once the candidate exists.
   */
  pendingPhotoFile?: string | null;
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
  /** Deliberately skipped (empty row, duplicate in Skip mode, …). */
  skippedRows: number;
  importedRows: number;
  /**
   * Never reached because the import stopped at the server time limit. Kept
   * apart from skippedRows: "64 empty rows skipped" was reported when 47 of
   * them were really unprocessed.
   */
  notProcessedRows?: number;
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
  /**
   * Set when the commit hit its server time budget and returned before
   * processing every row. Those rows come back as Skipped with
   * `notProcessed: true`.
   */
  stoppedEarly?: { reason: string; notProcessedRows: number } | null;
};

export type BulkCommitRowInput = {
  rowNumber: number;
  fields: Record<string, string | null>;
  /** Original spreadsheet cells — required for Training Matrix N-code columns. */
  source?: Record<string, string | null>;
};
