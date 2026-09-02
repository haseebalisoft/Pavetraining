import "server-only";

import { cache } from "react";

import { isAlwaysAdminEmail } from "@/lib/auth/protectedAdmins";
export { isAlwaysAdminEmail } from "@/lib/auth/protectedAdmins";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { getSharePointListId } from "@/lib/config/sharepoint";
import {
  asBoolean,
  asString,
  buildSchemaFieldEqualsFilter,
  getListItems,
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
 * Auth decisions must not use the SharePoint list cache. A stale/empty cached
 * Permissions query caused Admin UI to load (RSC) while DELETE /api/admin/*
 * returned 403 “You do not have access to this portal.”
 */
async function fetchActivePermissionItems(filter: string, top: number) {
  return getListItems(getSharePointListId("permissions"), { filter, top });
}

/**
 * SharePoint RoleType values used by the portal:
 * Admin | Training Manager | Supervisor | Customer
 *
 * Customer (and any leftover Candidate rows) = own profile only.
 *
 * Portal routing bucket (`RoleType`):
 * - Admin: literal Admin → /admin; Training Manager also uses this bucket
 *   for legacy APIs but still lands on /customer
 * - Customer: Supervisor + Customer → /customer
 */
export type PermissionFormRole =
  | "Admin"
  | "Manager"
  | "Supervisor"
  | "Customer"
  | "Candidate";

export function normalizePermissionRoleType(
  value: unknown,
): RoleType | null {
  const role = asString(value);
  if (!role) return null;
  const normalized = role.toLowerCase().trim();
  if (
    normalized === "admin" ||
    normalized === "training manager" ||
    normalized === "trainingmanager" ||
    normalized === "manager"
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

/** Form / admin UI role. Admin, Manager, Supervisor, and Customer are distinct. */
export function normalizePermissionFormRole(
  value: unknown,
): PermissionFormRole | null {
  const role = asString(value);
  if (!role) return null;
  const normalized = role.toLowerCase().trim().replace(/\s+/g, " ");
  if (normalized === "admin") return "Admin";
  if (
    normalized === "manager" ||
    normalized === "training manager" ||
    normalized === "trainingmanager"
  ) {
    return "Manager";
  }
  if (normalized === "supervisor") return "Supervisor";
  // One Customer role: own profile. Legacy "Candidate" is the same thing.
  if (normalized === "customer" || normalized === "candidate") return "Customer";
  return null;
}

/**
 * Values written back to the SharePoint RoleType column from the admin form.
 */
export function toSharePointRoleType(role: PermissionFormRole): string {
  if (role === "Admin") return "Admin";
  if (role === "Manager") return "Training Manager";
  if (role === "Supervisor") return "Supervisor";
  return "Customer";
}

/** Read RoleType whether Graph returns a string, multi-choice array, or lookup. */
export function parseSharePointRoleType(value: unknown): string {
  if (typeof value === "string" && value.trim()) return value.trim();
  if (Array.isArray(value)) {
    const parts = value
      .map((entry) => String(entry ?? "").trim())
      .filter(Boolean);
    const preferred = parts.find((part) =>
      /^(admin|training manager|trainingmanager|manager|supervisor|customer|candidate)$/i.test(
        part,
      ),
    );
    return preferred ?? parts[0] ?? "";
  }
  if (value && typeof value === "object") {
    const lookup = asString(
      (value as { LookupValue?: unknown }).LookupValue,
    );
    if (lookup) return lookup;
  }
  return asString(value) ?? "";
}

/**
 * Workforce TM/Supervisor auto-create still uses the routing bucket
 * (`RoleType` Admin | Customer), not the admin form roles.
 */
export function routingRoleToSharePointRoleType(role: RoleType): string {
  return role === "Admin" ? "Training Manager" : "Supervisor";
}

/** Admin form value derived from SharePoint RoleType + AccessScope. */
export function permissionFormRoleFromSharePoint(
  sharePointRole: string,
  _accessScope: string,
): PermissionFormRole {
  const key = sharePointRole.toLowerCase().trim().replace(/\s+/g, " ");
  if (key === "admin") return "Admin";
  if (key === "training manager" || key === "trainingmanager" || key === "manager") {
    return "Manager";
  }
  if (key === "supervisor") return "Supervisor";
  if (key === "customer" || key === "candidate") return "Customer";
  const customerRole = resolveCustomerRole(sharePointRole, _accessScope);
  if (customerRole === "TrainingManager") return "Manager";
  if (customerRole === "Supervisor") return "Supervisor";
  if (customerRole === "Candidate") return "Customer";
  return "Manager";
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

  // Customer (and legacy Candidate) = own profile only, like before.
  if (role === "customer" || role === "candidate") return "Candidate";

  if (role === "supervisor") {
    if (scope.includes("candidate")) return "Candidate";
    return "Supervisor";
  }

  return null;
}

export function roleLabelFor(
  sharePointRole: string,
  customerRole: CustomerRoleType | null,
): string {
  const key = sharePointRole.toLowerCase().trim().replace(/\s+/g, " ");
  if (key === "admin") return "Admin";
  if (key === "training manager" || key === "trainingmanager" || key === "manager") {
    return "Manager";
  }
  if (key === "supervisor") return "Supervisor";
  if (key === "customer" || key === "candidate") return "Customer";
  if (customerRole === "TrainingManager") return "Manager";
  if (customerRole === "Supervisor") return "Supervisor";
  if (customerRole === "Candidate") return "Customer";
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
    parseSharePointRoleType(fields[permissionFields.roleType]);
  const roleType = normalizePermissionRoleType(sharePointRoleType);
  const companyId = resolveCompanyId(fields);
  const accessScope = asString(fields[permissionFields.accessScope])?.trim();

  if (!userEmail || !status || !roleType || !companyId) {
    return null;
  }

  const resolvedAccessScope = accessScope || "Full Company";
  const customerRole = resolveCustomerRole(sharePointRoleType, resolvedAccessScope);
  const isAdminOnly = roleType === "Admin" && customerRole === null;
  // Strict rule (client sign-off requirement): ONLY literal SharePoint
  // `RoleType = Admin` can enter /admin. Training Managers, Supervisors,
  // and Candidates route to /customer/*. `customerRole === null` reads
  // exactly the same intent — it's true only when the SP role is "Admin".
  const canAccessAdmin =
    sharePointRoleType.trim().toLowerCase() === "admin" &&
    customerRole === null;
  const canAccessCustomer = customerRole !== null;
  const departmentScopes = Array.from(
    new Set([
      ...parseMultiChoice(fields[permissionFields.departments]),
      ...parseMultiChoice(fields[permissionFields.departmentsAllowed]),
    ]),
  );
  const candidateScopeName =
    asString(fields[permissionFields.name])?.trim() || null;

  // Access Scope is authoritative for Training Managers:
  // - Full Company → whole company workforce / matrix
  // - Department Only → departmentScopes (empty scopes ⇒ nothing)
  // Do NOT downgrade Full Company just because Departments is filled —
  // that field is often populated as coverage metadata and was hiding
  // newly added workforce from the customer portal.
  let normalizedAccessScope = normalizeAccessScopeValue(
    resolvedAccessScope,
    customerRole,
    isAdminOnly,
  );
  if (
    customerRole === "TrainingManager" &&
    resolvedAccessScope.toLowerCase().includes("department") &&
    departmentScopes.length === 0
  ) {
    // Explicit Department Only with no coverage → empty set (not company-wide).
    normalizedAccessScope = "Department";
  }

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
    normalizedAccessScope,
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
 * True when the resolved profile is a full SharePoint Admin — i.e. literal
 * SharePoint RoleType = "Admin" (customerRole is null).
 *
 * The hardcoded protected-admin list intentionally does NOT override this:
 * that list only exists to (a) prevent the email being deleted / set to
 * Inactive from the app UI, and (b) act as a fallback synthetic profile
 * when no SharePoint row exists at all. An existing SharePoint row is
 * always honoured as-is; to promote a protected admin back to full Admin,
 * update their SharePoint Permissions row.
 */
export function isSharePointAdminForProfile(
  profile: Pick<PermissionProfile, "customerRole">,
): boolean {
  return profile.customerRole === null;
}

/** Synthetic Admin profile when SharePoint row is missing for a hardcoded email. */
function alwaysAdminPermissionProfile(email: string): PermissionProfile {
  return {
    id: "always-admin",
    userEmail: email,
    status: "Active",
    roleType: "Admin",
    sharePointRoleType: "Admin",
    customerRole: null,
    roleLabel: "Admin",
    companyId: "0",
    companyDisplayName: "All companies",
    accessScope: "All",
    normalizedAccessScope: "All",
    departmentScopes: [],
    candidateScopeName: null,
    canView: true,
    canDownload: true,
    canEdit: true,
    canAccessAdmin: true,
    canAccessCustomer: false,
    receiveDocumentNotifications: true,
    receiveExpiryNotifications: true,
    customerNotificationsEnabled: true,
  };
}

/**
 * Reads the SharePoint Permissions List for the signed-in user.
 * Company is always taken from this list — never from client input.
 * Hardcoded always-admin emails still reach /admin if their SP row is gone.
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

    let items = await fetchActivePermissionItems(exactFilter, 10);

    if (items.length === 0) {
      items = await fetchActivePermissionItems(
        buildSchemaFieldEqualsFilter("permissions", "status", "Active"),
        500,
      );
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

    // Prefer Admin routing (pure SP Admin OR TM legacy routing) when both
    // roles exist for one email. The strict access rule is applied later on
    // canAccessAdmin — this pick just decides which row wins when duplicates.
    const chosen =
      matches.find((permission) => permission.roleType === "Admin") ??
      matches[0] ??
      null;

    // Hardcoded protected admins only fall back to the synthetic Admin
    // profile when NO SharePoint row exists for them. If a row exists (even
    // with a non-admin RoleType), that row is authoritative — matching the
    // client's rule "existing SharePoint row is always honoured as-is".
    if (chosen) {
      return chosen;
    }
    if (isAlwaysAdminEmail(normalizedEmail)) {
      return alwaysAdminPermissionProfile(normalizedEmail);
    }
    return null;
  },
);

export type PermissionResolutionReason = "not_found" | "inactive";

/**
 * Only called after getActivePermissionByEmail returns null, to distinguish
 * "no Permissions row for this email at all" from "a row exists but isn't
 * Active" — so access-denied can say "Permission not configured" vs.
 * "Access pending" instead of one generic message for both.
 */
export const getPermissionResolutionReason = cache(
  async (email: string): Promise<PermissionResolutionReason> => {
    const normalizedEmail = email.trim().toLowerCase();
    if (!normalizedEmail) {
      return "not_found";
    }

    const filter = buildSchemaFieldEqualsFilter(
      "permissions",
      "userEmail",
      normalizedEmail,
    );
    const items = await fetchActivePermissionItems(filter, 10);
    for (const item of items) {
      const permission = mapPermissionItem(item.id, item.fields);
      if (permission && permission.userEmail === normalizedEmail) {
        return "inactive";
      }
    }
    return "not_found";
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
