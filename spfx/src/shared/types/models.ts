/**
 * Domain models for SPFx PAVE portals (aligned with Next.js src/types/models.ts).
 */

export type RoleType = "Admin" | "Customer";

export interface PermissionProfile {
  id: string;
  userEmail: string;
  roleType: RoleType;
  status: string;
  companyId: string;
  companyDisplayName?: string;
  accessScope: string;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
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
