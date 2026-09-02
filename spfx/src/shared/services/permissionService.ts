import { getSharePointFields } from "../schema/sharepointSchema";
import type {
  CustomerRoleType,
  NormalizedAccessScope,
  PermissionProfile,
  RoleType,
} from "../types/models";
import {
  asBoolean,
  asString,
  fieldEqualsFilter,
  getListItem,
  getListItems,
  normalizeSharePointUserEmail,
  type SpListClient,
} from "./sharePointListService";

const permissionFields = getSharePointFields("permissions");
const companyFields = getSharePointFields("company");

function emailsMatch(a: string, b: string): boolean {
  const left = normalizeSharePointUserEmail(a);
  const right = normalizeSharePointUserEmail(b);
  if (!left || !right) return false;
  return left === right;
}

/**
 * SharePoint RoleType choices today: Training Manager | Supervisor
 * (Admin / Candidate may be added later — already mapped).
 *
 * Portal routing:
 * - Admin bucket: Admin + Training Manager (legacy PAVE staff)
 * - Customer bucket: Training Manager + Supervisor + Candidate
 */
export function normalizePermissionRoleType(value: unknown): RoleType | null {
  const role = asString(value);
  if (!role) return null;
  const normalized = role.toLowerCase().trim();
  if (
    normalized === "admin" ||
    normalized === "training manager" ||
    normalized === "trainingmanager" ||
    normalized.indexOf("training manager") >= 0
  ) {
    return "Admin";
  }
  if (
    normalized === "customer" ||
    normalized === "supervisor" ||
    normalized.indexOf("supervisor") >= 0 ||
    normalized === "candidate" ||
    normalized.indexOf("candidate") >= 0
  ) {
    return "Customer";
  }
  return null;
}

export function resolveCustomerRole(
  sharePointRole: string,
  accessScope: string
): CustomerRoleType | null {
  const role = sharePointRole.toLowerCase().trim();
  const scope = accessScope.toLowerCase().trim();

  if (role === "admin") return null;

  if (
    role === "training manager" ||
    role === "trainingmanager" ||
    role.indexOf("training manager") >= 0
  ) {
    return "TrainingManager";
  }

  if (role === "candidate" || role.indexOf("candidate") >= 0) {
    return "Candidate";
  }

  if (role === "supervisor" || role.indexOf("supervisor") >= 0) {
    if (scope.indexOf("candidate") >= 0) return "Candidate";
    return "Supervisor";
  }

  if (role === "customer") {
    if (scope.indexOf("candidate") >= 0) return "Candidate";
    if (scope.indexOf("department") >= 0) return "Supervisor";
    return "Supervisor";
  }

  return null;
}

export function roleLabelFor(
  sharePointRole: string,
  customerRole: CustomerRoleType | null
): string {
  const key = sharePointRole.toLowerCase().trim().replace(/\s+/g, " ");
  if (key === "admin") return "Admin";
  if (key === "training manager" || key === "trainingmanager" || key === "manager") {
    return "Manager";
  }
  if (key === "supervisor") return "Supervisor";
  if (key === "customer") return "Customer";
  if (key === "candidate") return "Candidate";
  if (customerRole === "TrainingManager") return "Manager";
  if (customerRole === "Supervisor") return "Supervisor";
  if (customerRole === "Candidate") return "Candidate";
  const raw = sharePointRole.trim();
  if (raw) return raw;
  return "Admin";
}

export function normalizeAccessScopeValue(
  value: string,
  customerRole: CustomerRoleType | null,
  isAdminOnly: boolean
): NormalizedAccessScope {
  if (isAdminOnly) return "All";
  const s = value.toLowerCase().trim();
  if (s === "all" || s.indexOf("all compan") >= 0) return "All";
  if (s.indexOf("candidate") >= 0) return "CandidateOnly";
  if (s.indexOf("assigned") >= 0) return "AssignedCandidates";
  if (s.indexOf("department") >= 0) return "Department";
  if (s.indexOf("full company") >= 0 || s === "company") return "Company";
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

function resolveCompanyId(fields: Record<string, unknown>): string | null {
  const lookupId = asString(fields[permissionFields.companyLookupId]);
  if (lookupId) return lookupId;

  const companyIdRest = asString(fields.CompanyId);
  if (companyIdRest) return companyIdRest;

  const companyValue = fields[permissionFields.company];
  if (
    companyValue &&
    typeof companyValue === "object" &&
    "LookupId" in (companyValue as object)
  ) {
    const nestedId = asString(
      (companyValue as { LookupId?: unknown }).LookupId
    );
    if (nestedId) return nestedId;
  }

  if (typeof companyValue === "number") {
    return String(companyValue);
  }

  return asString(companyValue);
}

function resolveCompanyDisplayName(
  fields: Record<string, unknown>
): string | undefined {
  const companyValue = fields[permissionFields.company];
  if (typeof companyValue === "string") {
    if (/^\d+$/.test(companyValue.trim())) return undefined;
    return companyValue.trim() || undefined;
  }
  if (
    companyValue &&
    typeof companyValue === "object" &&
    "LookupValue" in (companyValue as object)
  ) {
    return (
      asString((companyValue as { LookupValue?: unknown }).LookupValue) ??
      undefined
    );
  }
  if (
    companyValue &&
    typeof companyValue === "object" &&
    "CompanyName" in (companyValue as object)
  ) {
    return (
      asString((companyValue as { CompanyName?: unknown }).CompanyName) ??
      undefined
    );
  }
  const raw = asString(fields.Company);
  if (raw && !/^\d+$/.test(raw)) return raw;
  return undefined;
}

async function lookupCompanyName(
  client: SpListClient,
  companyId: string
): Promise<string | undefined> {
  if (!companyId || companyId === "0") return undefined;
  try {
    const item = await getListItem(client, "company", companyId);
    if (!item) return undefined;
    return (
      asString(item.fields[companyFields.companyName]) ||
      asString(item.fields.Title) ||
      undefined
    );
  } catch {
    return undefined;
  }
}

function mapPermissionItem(
  id: string,
  fields: Record<string, unknown>
): PermissionProfile | null {
  const userEmail = asString(fields[permissionFields.userEmail])?.trim();
  const status = asString(fields[permissionFields.status])?.trim();
  const sharePointRoleType =
    asString(fields[permissionFields.roleType])?.trim() || "";
  const roleType = normalizePermissionRoleType(sharePointRoleType);
  const companyId = resolveCompanyId(fields) || "0";
  const accessScope =
    asString(fields[permissionFields.accessScope])?.trim() ?? "Full Company";

  if (!userEmail || !status || !roleType) {
    return null;
  }

  const customerRole = resolveCustomerRole(sharePointRoleType, accessScope);
  const isAdminOnly = roleType === "Admin" && customerRole === null;
  const canAccessAdmin =
    roleType === "Admin" || customerRole === "TrainingManager";
  const canAccessCustomer = customerRole !== null;
  const departmentScopes = Array.from(
    new Set([
      ...parseMultiChoice(fields[permissionFields.departments]),
      ...parseMultiChoice(fields[permissionFields.departmentsAllowed]),
    ])
  );
  const candidateScopeName =
    asString(fields[permissionFields.name])?.trim() || null;

  return {
    id,
    permissionItemId: id,
    userEmail: normalizeSharePointUserEmail(userEmail),
    status,
    roleType,
    sharePointRoleType,
    customerRole,
    roleLabel: roleLabelFor(sharePointRoleType, customerRole),
    companyId,
    companyDisplayName: resolveCompanyDisplayName(fields),
    accessScope,
    normalizedAccessScope: normalizeAccessScopeValue(
      accessScope,
      customerRole,
      isAdminOnly
    ),
    departmentScopes,
    candidateScopeName,
    canView: asBoolean(fields[permissionFields.canView]),
    canDownload: asBoolean(fields[permissionFields.canDownload]),
    canEdit: asBoolean(fields[permissionFields.canEdit]),
    canAccessAdmin,
    canAccessCustomer,
  };
}

/**
 * Active Permissions List row for signed-in email.
 * Resolves Company Name from Company List when REST only returns CompanyId.
 */
export async function getActivePermissionByEmail(
  client: SpListClient,
  email: string
): Promise<PermissionProfile | null> {
  const normalizedEmail = normalizeSharePointUserEmail(email);
  if (!normalizedEmail) return null;

  let items = await getListItems(client, "permissions", {
    filter: [
      fieldEqualsFilter(permissionFields.userEmail, normalizedEmail),
      fieldEqualsFilter(permissionFields.status, "Active"),
    ].join(" and "),
    top: 100,
    maxItems: 100,
  });

  if (items.length === 0) {
    items = await getListItems(client, "permissions", {
      filter: fieldEqualsFilter(permissionFields.status, "Active"),
      top: 5000,
      maxItems: 5000,
    });
  }

  const matches: PermissionProfile[] = [];
  for (let i = 0; i < items.length; i++) {
    const permission = mapPermissionItem(items[i].id, items[i].fields);
    if (
      permission &&
      emailsMatch(permission.userEmail, normalizedEmail) &&
      permission.status.toLowerCase() === "active"
    ) {
      matches.push(permission);
    }
  }

  if (matches.length === 0) return null;

  // Prefer Admin (Training Manager) when both roles exist for one email.
  const chosen =
    matches.find((p) => p.roleType === "Admin") ?? matches[0];

  if (
    !chosen.companyDisplayName &&
    chosen.companyId &&
    chosen.companyId !== "0"
  ) {
    const name = await lookupCompanyName(client, chosen.companyId);
    if (name) {
      return { ...chosen, companyDisplayName: name };
    }
  }

  return chosen;
}

export function siteAdminPermissionProfile(
  email: string
): PermissionProfile {
  return {
    id: "site-admin",
    permissionItemId: "site-admin",
    userEmail: normalizeSharePointUserEmail(email),
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
  };
}

/** Access badge text for customer header. */
export function accessScopeBadgeLabel(permission: PermissionProfile): string {
  const scope = permission.normalizedAccessScope;
  if (scope === "Company" || scope === "All") return "Company-wide";
  if (scope === "Department") {
    if (permission.departmentScopes.length > 0) {
      return permission.departmentScopes.join(", ") + " department";
    }
    return "Department";
  }
  if (scope === "AssignedCandidates") return "Assigned candidates";
  if (scope === "CandidateOnly") return "Own records only";
  return permission.accessScope || "Company";
}
