/**
 * Domain models for SPFx PAVE portals (aligned with Next.js src/types/models.ts).
 */

/** Portal route bucket — Admin web part vs Customer web part. */
export type RoleType = "Admin" | "Customer";

/** Customer-side role from Permissions List (Wayne brief). */
export type CustomerRoleType =
  | "TrainingManager"
  | "Supervisor"
  | "Candidate";

/** Normalized AccessScope for filtering. */
export type NormalizedAccessScope =
  | "All"
  | "Company"
  | "Department"
  | "AssignedCandidates"
  | "CandidateOnly";

export interface PermissionProfile {
  id: string;
  userEmail: string;
  /** Admin | Customer — used for portal routing compatibility. */
  roleType: RoleType;
  /** Raw SharePoint RoleType choice text. */
  sharePointRoleType: string;
  /** Customer-side role when user may use the customer portal. */
  customerRole: CustomerRoleType | null;
  /** UI label e.g. "Training Manager". */
  roleLabel: string;
  status: string;
  companyId: string;
  companyDisplayName?: string;
  /** Raw SharePoint AccessScope choice. */
  accessScope: string;
  /** Normalized scope used by customer filters. */
  normalizedAccessScope: NormalizedAccessScope;
  /** From Permissions.Departments (multi-choice) / DepartmentsAllowed. */
  departmentScopes: string[];
  /** Display name used for CandidateOnly / Name field. */
  candidateScopeName: string | null;
  /** Permissions List item id — used when Supervisor lookup points here. */
  permissionItemId: string;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
  /** PAVE Admin portal (includes legacy Training Manager). */
  canAccessAdmin: boolean;
  /** Customer portal (TM / Supervisor / Candidate). */
  canAccessCustomer: boolean;
}

export interface ListRow {
  id: string;
  fields: Record<string, unknown>;
}

export type PortalViewId =
  | "dashboard"
  | "companies"
  | "workforce"
  | "training-matrix"
  | "training-records"
  | "nvq"
  | "documents"
  | "events"
  | "offers"
  | "permissions"
  | "automation"
  | "logs"
  | "candidates"
  | "nvq-progress"
  | "support";
