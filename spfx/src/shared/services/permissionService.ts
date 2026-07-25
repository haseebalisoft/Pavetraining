import { getSharePointFields } from "../schema/sharepointSchema";
import type { PermissionProfile, RoleType } from "../types/models";
import {
  asBoolean,
  asString,
  fieldEqualsFilter,
  getListItems,
  type SpListClient,
} from "./sharePointListService";

const permissionFields = getSharePointFields("permissions");

/**
 * SharePoint RoleType choices are Training Manager / Supervisor.
 * Portal routing only has Admin and Customer (same as Next.js).
 */
export function normalizePermissionRoleType(value: unknown): RoleType | null {
  const role = asString(value);
  if (!role) return null;
  const normalized = role.toLowerCase().trim();
  if (normalized === "admin" || normalized === "training manager") {
    return "Admin";
  }
  if (normalized === "customer" || normalized === "supervisor") {
    return "Customer";
  }
  return null;
}

function resolveCompanyId(fields: Record<string, unknown>): string | null {
  const lookupId = asString(fields[permissionFields.companyLookupId]);
  if (lookupId) return lookupId;

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

  const companyIdRest = asString(fields.CompanyId);
  if (companyIdRest) return companyIdRest;

  return asString(companyValue);
}

function resolveCompanyDisplayName(
  fields: Record<string, unknown>
): string | undefined {
  const companyValue = fields[permissionFields.company];
  if (typeof companyValue === "string") {
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
  return asString(fields.Company) ?? undefined;
}

function mapPermissionItem(
  id: string,
  fields: Record<string, unknown>
): PermissionProfile | null {
  const userEmail = asString(fields[permissionFields.userEmail])?.trim();
  const status = asString(fields[permissionFields.status])?.trim();
  const roleType = normalizePermissionRoleType(
    fields[permissionFields.roleType]
  );
  const companyId = resolveCompanyId(fields);
  const accessScope =
    asString(fields[permissionFields.accessScope])?.trim() ?? "Company";

  if (!userEmail || !status || !roleType || !companyId) {
    return null;
  }

  return {
    id,
    userEmail: userEmail.toLowerCase(),
    status,
    roleType,
    companyId,
    companyDisplayName: resolveCompanyDisplayName(fields),
    accessScope,
    canView: asBoolean(fields[permissionFields.canView]),
    canDownload: asBoolean(fields[permissionFields.canDownload]),
    canEdit: asBoolean(fields[permissionFields.canEdit]),
  };
}

/**
 * Same logic as Next.js permissionService: Active row for signed-in email.
 * Prefer Admin when both roles exist.
 */
export async function getActivePermissionByEmail(
  client: SpListClient,
  email: string
): Promise<PermissionProfile | null> {
  const normalizedEmail = email.trim().toLowerCase();
  if (!normalizedEmail) return null;

  let items = await getListItems(client, "permissions", {
    filter: [
      fieldEqualsFilter(permissionFields.userEmail, normalizedEmail),
      fieldEqualsFilter(permissionFields.status, "Active"),
    ].join(" and "),
    top: 20,
  });

  if (items.length === 0) {
    items = await getListItems(client, "permissions", {
      filter: fieldEqualsFilter(permissionFields.status, "Active"),
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

  if (matches.length === 0) return null;
  return matches.find((p) => p.roleType === "Admin") ?? matches[0];
}
