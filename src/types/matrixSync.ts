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
  scope: "all" | "company" | "candidate" | "register";
  items: MatrixSyncResultItem[];
  summary: {
    processed: number;
    updated: number;
    created: number;
    skipped: number;
    errors: number;
    warnings: number;
  };
}
