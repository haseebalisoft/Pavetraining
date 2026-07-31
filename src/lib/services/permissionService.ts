import "server-only";

import { cache } from "react";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asString,
  buildSchemaFieldEqualsFilter,
  getListItemsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import type {
  CustomerRoleType,
  NormalizedAccessScope,
  PermissionProfile,
  RoleType,
} from "@/types/models";

const permissionFields = getSharePointFields("permissions");

/**
 * SharePoint RoleType values used by the portal:
 * Training Manager | Supervisor | Admin | Candidate
 *
 * Portal routing bucket (`RoleType`):
 * - Admin: Admin + Training Manager (PAVE staff → /admin; TM also sees customer APIs)
 * - Customer: Supervisor + Candidate → /customer (Training Matrix, docs, etc.)
 */
export type PermissionFormRole = "Admin" | "Customer" | "Candidate";

export function normalizePermissionRoleType(
  value: unknown,
): RoleType | null {
  const role = asString(value);
  if (!role) return null;
  const normalized = role.toLowerCase().trim();
  if (
    normalized === "admin" ||
    normalized === "training manager" ||
    normalized === "trainingmanager"
  ) {
    return "Admin";
  }
  if (
    normalized === "customer" ||
    normalized === "supervisor" ||
    normalized === "candidate"
  ) {
    return "Customer";
  }
  return null;
}

/** Form / admin UI role including first-class Candidate. */
export function normalizePermissionFormRole(
  value: unknown,
): PermissionFormRole | null {
  const role = asString(value);
  if (!role) return null;
  const normalized = role.toLowerCase().trim();
  if (
    normalized === "admin" ||
    normalized === "training manager" ||
    normalized === "trainingmanager"
  ) {
    return "Admin";
  }
  if (normalized === "candidate") return "Candidate";
  if (normalized === "customer" || normalized === "supervisor") {
    return "Customer";
  }
  return null;
}

/**
 * Values written back to the SharePoint RoleType column.
 * Candidate is first-class (accepted by the live list even if not in the choice UI).
 */
export function toSharePointRoleType(
  role: RoleType | PermissionFormRole,
): string {
  if (role === "Admin") return "Training Manager";
  if (role === "Candidate") return "Candidate";
  return "Supervisor";
}

/** Admin form value derived from SharePoint RoleType + AccessScope. */
export function permissionFormRoleFromSharePoint(
  sharePointRole: string,
  accessScope: string,
): PermissionFormRole {
  const customerRole = resolveCustomerRole(sharePointRole, accessScope);
  if (customerRole === "Candidate") return "Candidate";
  const routing = normalizePermissionRoleType(sharePointRole);
  return routing === "Admin" ? "Admin" : "Customer";
}

export function resolveCustomerRole(
  sharePointRole: string,
  accessScope: string,
): CustomerRoleType | null {
  const role = sharePointRole.toLowerCase().trim();
  const scope = accessScope.toLowerCase().trim();

  if (role === "admin") return null;

  if (role === "training manager" || role === "trainingmanager") {
    return "TrainingManager";
  }

  if (role === "candidate") return "Candidate";

  if (role === "supervisor" || role === "customer") {
    if (scope.includes("candidate")) return "Candidate";
    return "Supervisor";
  }

  return null;
}

export function roleLabelFor(
  sharePointRole: string,
  customerRole: CustomerRoleType | null,
): string {
  if (customerRole === "TrainingManager") return "Training Manager";
  if (customerRole === "Supervisor") return "Supervisor";
  if (customerRole === "Candidate") return "Candidate";
  return sharePointRole.trim() || "Admin";
}

export function normalizeAccessScopeValue(
  value: string,
  customerRole: CustomerRoleType | null,
  isAdminOnly: boolean,
): NormalizedAccessScope {
  if (isAdminOnly) return "All";
  const s = value.toLowerCase().trim();
  if (s === "all" || s.includes("all compan")) return "All";
  if (s.includes("candidate")) return "CandidateOnly";
  if (s.includes("assigned")) return "AssignedCandidates";
  if (s.includes("department")) return "Department";
  if (s.includes("full company") || s === "company") return "Company";
  if (customerRole === "Candidate") return "CandidateOnly";
  if (customerRole === "Supervisor") return "Department";
  if (customerRole === "TrainingManager") return "Company";
  return "Company";
}

function parseMultiChoice(value: unknown): string[] {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value
      .map((entry) => {
        if (typeof entry === "string") return entry.trim();
        if (entry && typeof entry === "object") {
          return (
            asString((entry as { LookupValue?: unknown }).LookupValue) ||
            asString((entry as { Title?: unknown }).Title) ||
            asString((entry as { Name?: unknown }).Name) ||
            ""
          );
        }
        return "";
      })
      .filter(Boolean);
  }
  if (typeof value === "string") {
    return value
      .split(/;|#/)
      .map((part) => part.trim())
      .filter((part) => part && !/^\d+$/.test(part));
  }
  return [];
}

function resolveCompanyId(fields: SharePointFields): string | null {
  const lookupId = asString(fields[permissionFields.companyLookupId]);
  if (lookupId) return lookupId;

  const companyValue = fields[permissionFields.company];
  if (
    companyValue &&
    typeof companyValue === "object" &&
    "LookupId" in companyValue
  ) {
    const nestedId = asString(
      (companyValue as { LookupId?: unknown }).LookupId,
    );
    if (nestedId) return nestedId;
  }

  return asString(companyValue) ?? null;
}

function resolveCompanyDisplayName(fields: SharePointFields): string | undefined {
  const companyValue = fields[permissionFields.company];

  if (typeof companyValue === "string") {
    if (/^\d+$/.test(companyValue.trim())) return undefined;
    return companyValue.trim() || undefined;
  }

  if (
    companyValue &&
    typeof companyValue === "object" &&
    "LookupValue" in companyValue
  ) {
    return asString((companyValue as { LookupValue?: unknown }).LookupValue);
  }

  return undefined;
}

function mapPermissionItem(
  id: string,
  fields: SharePointFields,
): PermissionProfile | null {
  const userEmail = asString(fields[permissionFields.userEmail])?.trim();
  const status = asString(fields[permissionFields.status])?.trim();
  const sharePointRoleType =
    asString(fields[permissionFields.roleType])?.trim() || "";
  const roleType = normalizePermissionRoleType(sharePointRoleType);
  const companyId = resolveCompanyId(fields);
  const accessScope = asString(fields[permissionFields.accessScope])?.trim();

  if (!userEmail || !status || !roleType || !companyId) {
    return null;
  }

  const resolvedAccessScope = accessScope || "Full Company";
  const customerRole = resolveCustomerRole(sharePointRoleType, resolvedAccessScope);
  const isAdminOnly = roleType === "Admin" && customerRole === null;
  const canAccessAdmin =
    roleType === "Admin" || customerRole === "TrainingManager";
  const canAccessCustomer = customerRole !== null;
  const departmentScopes = Array.from(
    new Set([
      ...parseMultiChoice(fields[permissionFields.departments]),
      ...parseMultiChoice(fields[permissionFields.departmentsAllowed]),
    ]),
  );
  const candidateScopeName =
    asString(fields[permissionFields.name])?.trim() || null;

  return {
    id,
    userEmail: userEmail.toLowerCase(),
    status,
    roleType,
    sharePointRoleType,
    customerRole,
    roleLabel: roleLabelFor(sharePointRoleType, customerRole),
    companyId,
    companyDisplayName: resolveCompanyDisplayName(fields),
    accessScope: resolvedAccessScope,
    normalizedAccessScope: normalizeAccessScopeValue(
      resolvedAccessScope,
      customerRole,
      isAdminOnly,
    ),
    departmentScopes,
    candidateScopeName,
    canView: asBoolean(fields[permissionFields.canView]),
    canDownload: asBoolean(fields[permissionFields.canDownload]),
    canEdit: asBoolean(fields[permissionFields.canEdit]),
    canAccessAdmin,
    canAccessCustomer,
    receiveDocumentNotifications: preferenceOrDefault(
      fields,
      "receiveDocumentNotifications",
      true,
    ),
    receiveExpiryNotifications: preferenceOrDefault(
      fields,
      "receiveExpiryNotifications",
      true,
    ),
    customerNotificationsEnabled: preferenceOrDefault(
      fields,
      "customerNotificationsEnabled",
      true,
    ),
  };
}

/**
 * Missing preference columns default to enabled so existing tenants keep notifying.
 */
function preferenceOrDefault(
  fields: SharePointFields,
  key: keyof typeof permissionFields,
  defaultValue: boolean,
): boolean {
  const internal = permissionFields[key];
  if (!(internal in fields)) {
    return defaultValue;
  }
  return asBoolean(fields[internal]);
}

/**
 * Reads the SharePoint Permissions List for the signed-in user.
 * Company is always taken from this list — never from client input.
 */
export const getActivePermissionByEmail = cache(
  async (email: string): Promise<PermissionProfile | null> => {
    const normalizedEmail = email.trim().toLowerCase();
    if (!normalizedEmail) {
      return null;
    }

    const exactFilter = [
      buildSchemaFieldEqualsFilter("permissions", "userEmail", normalizedEmail),
      buildSchemaFieldEqualsFilter("permissions", "status", "Active"),
    ].join(" and ");

    let items = await getListItemsByKey("permissions", {
      filter: exactFilter,
      top: 10,
    });

    if (items.length === 0) {
      items = await getListItemsByKey("permissions", {
        filter: buildSchemaFieldEqualsFilter("permissions", "status", "Active"),
        top: 500,
      });
    }

    const matches: PermissionProfile[] = [];
    for (const item of items) {
      const permission = mapPermissionItem(item.id, item.fields);
      if (
        permission &&
        permission.userEmail === normalizedEmail &&
        permission.status.toLowerCase() === "active"
      ) {
        matches.push(permission);
      }
    }

    if (matches.length === 0) {
      return null;
    }

    // Prefer Admin (Training Manager) when both roles exist for one email.
    return (
      matches.find((permission) => permission.roleType === "Admin") ??
      matches[0]
    );
  },
);

export function accessScopeBadgeLabel(permission: PermissionProfile): string {
  const scope = permission.normalizedAccessScope;
  if (scope === "Company" || scope === "All") return "Company-wide";
  if (scope === "Department") {
    if (permission.departmentScopes.length > 0) {
      return `${permission.departmentScopes.join(", ")} department`;
    }
    return "Department";
  }
  if (scope === "AssignedCandidates") return "Assigned candidates";
  if (scope === "CandidateOnly") return "Own records only";
  return permission.accessScope || "Company";
}
