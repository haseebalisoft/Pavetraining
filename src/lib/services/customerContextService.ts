import "server-only";

import { cache } from "react";

import { getCompanyById } from "@/lib/services/companyService";
import { getActivePermissionByEmail } from "@/lib/services/permissionService";
import {
  AccessDeniedError,
  UnauthorizedError,
} from "@/lib/services/errorHandler";
import type {
  AdminContext,
  CustomerContext,
  MeResponse,
} from "@/types/models";

export { AccessDeniedError, UnauthorizedError } from "@/lib/services/errorHandler";

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
 * Deduped per request when layout + page both resolve context.
 */
export const getMeContext = cache(
  async (email: string): Promise<MeResponse | null> => {
    const permission = await getActivePermissionByEmail(email);
    if (!permission) {
      return null;
    }

    const companyName =
      permission.roleType === "Customer"
        ? await resolveCompanyName(
            permission.companyId,
            permission.companyDisplayName,
          )
        : (permission.companyDisplayName ?? null);

    return {
      loggedInEmail: permission.userEmail,
      role: permission.roleType,
      redirectTo: permission.roleType === "Admin" ? "/admin" : "/customer",
      companyId: permission.companyId,
      companyName,
      canView: permission.canView,
      canDownload: permission.canDownload,
      canEdit: permission.canEdit,
      accessScope: permission.accessScope,
    };
  },
);

export const getCustomerContext = cache(
  async (email: string): Promise<CustomerContext> => {
    const permission = await getActivePermissionByEmail(email);

    if (!permission || permission.roleType !== "Customer") {
      throw new AccessDeniedError(
        "No active Customer permission found for this account.",
      );
    }

    if (permission.status.toLowerCase() !== "active") {
      throw new AccessDeniedError(
        "No active Customer permission found for this account.",
      );
    }

    const companyName = await resolveCompanyName(
      permission.companyId,
      permission.companyDisplayName,
    );

    return {
      loggedInEmail: permission.userEmail,
      role: "Customer",
      companyId: permission.companyId,
      companyName,
      canView: permission.canView,
      canDownload: permission.canDownload,
      canEdit: permission.canEdit,
      accessScope: permission.accessScope,
      permissionStatus: "Active",
    };
  },
);

export const getAdminContext = cache(
  async (email: string): Promise<AdminContext> => {
    const permission = await getActivePermissionByEmail(email);

    if (!permission || permission.roleType !== "Admin") {
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
    };
  },
);
