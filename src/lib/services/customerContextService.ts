import "server-only";

import { cache } from "react";

import { getCompanyById } from "@/lib/services/companyService";
import {
  accessScopeBadgeLabel,
  getActivePermissionByEmail,
  isAlwaysAdminEmail,
  isSharePointAdminForProfile,
} from "@/lib/services/permissionService";
import {
  AccessDeniedError,
  UnauthorizedError,
} from "@/lib/services/errorHandler";
import type {
  AdminContext,
  CustomerContext,
  CustomerRoleType,
  MeResponse,
} from "@/types/models";

export { AccessDeniedError, UnauthorizedError } from "@/lib/services/errorHandler";
export { accessScopeBadgeLabel };

async function resolveCompanyName(
  companyId: string,
  fallbackDisplayName?: string,
): Promise<string> {
  const company = await getCompanyById(companyId);
  if (company?.companyName) {
    return company.companyName;
  }

  if (fallbackDisplayName?.trim()) {
    return fallbackDisplayName.trim();
  }

  throw new Error(
    `Company ${companyId} was referenced in Permissions but could not be resolved.`,
  );
}

/**
 * Builds the authenticated user's portal context from SharePoint Permissions.
 * Customer companyId is never accepted from the client — only from Permissions.
 */
export const getMeContext = cache(
  async (email: string): Promise<MeResponse | null> => {
    const permission = await getActivePermissionByEmail(email);
    if (!permission) {
      return null;
    }

    const companyName =
      permission.canAccessCustomer || permission.roleType === "Customer"
        ? await resolveCompanyName(
            permission.companyId,
            permission.companyDisplayName,
          )
        : (permission.companyDisplayName ?? null);

    // STRICT: only literal SP RoleType = Admin can reach /admin. Every
    // customer-side role (Training Manager / Supervisor / Candidate) lands
    // on the Customer Dashboard. Since `canAccessAdmin` is now equivalent
    // to `customerRole === null && sharePointRoleType === "Admin"`, either
    // predicate is sufficient — we keep `canAccessAdmin` for readability.
    const redirectTo =
      permission.customerRole != null
        ? "/customer/dashboard"
        : permission.canAccessAdmin
          ? "/admin"
          : "/customer/dashboard";

    return {
      loggedInEmail: permission.userEmail,
      role:
        permission.customerRole != null
          ? "Customer"
          : permission.canAccessAdmin
            ? "Admin"
            : permission.roleType,
      redirectTo,
      companyId: permission.companyId,
      companyName,
      canView: permission.canView,
      canDownload: permission.canDownload,
      canEdit: permission.canEdit,
      accessScope: permission.accessScope,
      customerRole: permission.customerRole,
      roleLabel: permission.roleLabel,
      normalizedAccessScope: permission.normalizedAccessScope,
    };
  },
);

export const getCustomerContext = cache(
  async (email: string): Promise<CustomerContext> => {
    const permission = await getActivePermissionByEmail(email);

    if (
      !permission ||
      !permission.canAccessCustomer ||
      !permission.customerRole
    ) {
      throw new AccessDeniedError(
        "No active Customer permission found for this account.",
      );
    }

    if (permission.status.toLowerCase() !== "active") {
      throw new AccessDeniedError(
        "No active Customer permission found for this account.",
      );
    }

    if (!permission.canView) {
      throw new AccessDeniedError(
        "Your account does not have view permission.",
      );
    }

    const companyName = await resolveCompanyName(
      permission.companyId,
      permission.companyDisplayName,
    );

    return {
      loggedInEmail: permission.userEmail,
      role: "Customer",
      customerRole: permission.customerRole as CustomerRoleType,
      roleLabel: permission.roleLabel,
      companyId: permission.companyId,
      companyName,
      canView: permission.canView,
      canDownload: permission.canDownload,
      canEdit: permission.canEdit,
      accessScope: permission.accessScope,
      normalizedAccessScope: permission.normalizedAccessScope,
      departmentScopes: permission.departmentScopes,
      candidateScopeName: permission.candidateScopeName,
      permissionStatus: "Active",
    };
  },
);

export const getAdminContext = cache(
  async (email: string): Promise<AdminContext> => {
    const permission = await getActivePermissionByEmail(email);

    if (!permission || !permission.canAccessAdmin) {
      throw new AccessDeniedError(
        "No active Admin permission found for this account.",
      );
    }

    if (permission.status.toLowerCase() !== "active") {
      throw new AccessDeniedError(
        "No active Admin permission found for this account.",
      );
    }

    return {
      loggedInEmail: permission.userEmail,
      role: "Admin",
      selectedCompanyId: null,
      canView: permission.canView,
      canDownload: permission.canDownload,
      canEdit: permission.canEdit,
      accessScope: permission.accessScope,
      sharePointRoleType: permission.sharePointRoleType,
      customerRole: permission.customerRole,
      roleLabel: permission.roleLabel,
      // Hardcoded protected admins ALWAYS resolve as SharePoint Admin even
      // when their SP row has been edited to Training Manager — otherwise
      // they silently lose Bulk Upload / Permissions in the nav.
      isSharePointAdmin: isSharePointAdminForProfile(permission),
      isAlwaysAdminEmail: isAlwaysAdminEmail(permission.userEmail),
    };
  },
);
