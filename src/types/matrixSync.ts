import type { MatrixReviewReason } from "@/types/bulkUpload";

export type RegisterSource = "NPORS" | "EUSR" | "NRSWA" | "In-House";

export interface MatrixSyncResultItem {
  candidate: string;
  company: string;
  candidateId: string | null;
  companyId: string | null;
  registerSources: RegisterSource[];
  matrixRowId: string | null;
  matrixRowFound: boolean;
  matrixRowCreated: boolean;
  fieldsUpdated: string[];
  warnings: string[];
  errors: string[];
  skipped: boolean;
  skipReason?: string;
}

export interface MatrixSyncResult {
  dryRun: boolean;
  scope: "all" | "company" | "candidate" | "register" | "register-delete";
  items: MatrixSyncResultItem[];
  summary: {
    processed: number;
    updated: number;
    created: number;
    skipped: number;
    errors: number;
    warnings: number;
  };
  /**
   * Set when the sync hit its server time budget before reaching every
   * candidate. Already-synced candidates are kept; running it again finishes.
   */
  incomplete?: { reason: string; remainingCandidates: number } | null;
}

/** What re-running Workforce matching did after a name / DOB correction. */
export type MatrixRelinkOutcome =
  | "linked"
  | "merged"
  | "alreadyLinked"
  | "unmatched";

export interface MatrixRelinkResult {
  outcome: MatrixRelinkOutcome;
  /** Row that now holds the candidate's training (`example:<id>`). */
  matrixRowId: string | null;
  workforceId: string | null;
  candidateName: string;
  /** One sentence for the admin: what happened and, if unmatched, why. */
  message: string;
  /** Set when `unmatched`: exactly why no Workforce record was chosen. */
  reason: MatrixReviewReason | null;
  /** `merged`: expiry columns copied from the removed duplicate row. */
  mergedHeaders: string[];
}
