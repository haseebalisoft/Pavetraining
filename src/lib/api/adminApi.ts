import { NextResponse } from "next/server";

import {
  withAdminApi as withAdminApiGuard,
  withSharePointAdminApi as withSharePointAdminApiGuard,
} from "@/lib/api/apiGuards";
import type { AdminContext } from "@/types/models";

export { ValidationError } from "@/lib/services/validationService";
export {
  AccessDeniedError,
  NotFoundError,
  UnauthorizedError,
} from "@/lib/services/errorHandler";

/**
 * Shared admin API guard: RoleType = Admin + Status Active only.
 * Mutating requests are audited to Training Manager Logs when configured.
 */
export async function withAdminApi<T>(
  routeLabel: string,
  handler: (context: AdminContext, request: Request) => Promise<T>,
  options?: {
    errorMessage?: string;
    audit?: boolean;
    entityName?: string;
  },
  request: Request = new Request("http://localhost"),
): Promise<NextResponse> {
  return withAdminApiGuard(routeLabel, handler, options, request);
}

/**
 * Stricter admin guard for pure-SharePoint-Admin-only routes (Bulk Upload,
 * Permissions writes). Rejects Training Managers with 403.
 */
export async function withSharePointAdminApi<T>(
  routeLabel: string,
  handler: (context: AdminContext, request: Request) => Promise<T>,
  options?: {
    errorMessage?: string;
    audit?: boolean;
    entityName?: string;
  },
  request: Request = new Request("http://localhost"),
): Promise<NextResponse> {
  return withSharePointAdminApiGuard(routeLabel, handler, options, request);
}
