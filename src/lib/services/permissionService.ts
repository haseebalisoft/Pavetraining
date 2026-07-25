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
import type { PermissionProfile, RoleType } from "@/types/models";

const permissionFields = getSharePointFields("permissions");

/**
 * SharePoint RoleType choices are Training Manager / Supervisor.
 * Portal routing only has Admin (/admin) and Customer (/customer).
 */
export function normalizePermissionRoleType(
  value: unknown,
): RoleType | null {
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

/** Values written back to the SharePoint RoleType choice column. */
export function toSharePointRoleType(role: RoleType): string {
  return role === "Admin" ? "Training Manager" : "Supervisor";
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
  const roleType = normalizePermissionRoleType(
    fields[permissionFields.roleType],
  );
  const companyId = resolveCompanyId(fields);
  const accessScope = asString(fields[permissionFields.accessScope])?.trim();

  if (!userEmail || !status || !roleType || !companyId || !accessScope) {
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
 * Reads the SharePoint Permissions List for the signed-in user.
 * Company is always taken from this list — never from client input.
 * Deduped per request via React.cache; list reads use short SharePoint cache.
 */
export const getActivePermissionByEmail = cache(
  async (email: string): Promise<PermissionProfile | null> => {
    const normalizedEmail = email.trim().toLowerCase();
    if (!normalizedEmail) {
      return null;
    }

    // Prefer exact SharePoint filter, then fall back to Active scan so leading/
    // trailing spaces in UserEmail still match the signed-in account.
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
