import { NextResponse } from "next/server";

import {
  PUBLIC_ERROR_MESSAGES,
  toPublicErrorPayload,
} from "@/lib/services/errorHandler";
import { extractItemId, writeAuditLog } from "@/lib/services/auditLogService";
import {
  requireAdminAccess,
  requireCustomerAccess,
} from "@/lib/services/securityService";
import type { AdminContext, CustomerContext } from "@/types/models";

export {
  AccessDeniedError,
  NotFoundError,
  UnauthorizedError,
  ValidationError,
} from "@/lib/services/errorHandler";

function entityFromRouteLabel(routeLabel: string): string {
  return routeLabel.replace(/^(GET|POST|PATCH|PUT|DELETE)\s+/i, "").trim();
}

function actionFromRequest(request: Request): string {
  return request.method.toUpperCase();
}

/**
 * Customer API guard:
 * - authenticated user
 * - Active Customer permission
 * - company from Permissions only
 * - rejects client companyId / export attempts
 */
export async function withCustomerApi<T>(
  routeLabel: string,
  handler: (context: CustomerContext, request: Request) => Promise<T>,
  request: Request,
  options?: {
    audit?: boolean;
    entityName?: string;
  },
): Promise<NextResponse> {
  let email = "unknown";
  try {
    const context = await requireCustomerAccess(request);
    email = context.loggedInEmail;
    const data = await handler(context, request);

    if (options?.audit) {
      await writeAuditLog({
        userEmail: email,
        action: actionFromRequest(request),
        entityName: options.entityName ?? entityFromRouteLabel(routeLabel),
        itemId: extractItemId(data),
        success: true,
      });
    }

    return NextResponse.json(data);
  } catch (error) {
    if (options?.audit) {
      await writeAuditLog({
        userEmail: email,
        action: actionFromRequest(request),
        entityName: options.entityName ?? entityFromRouteLabel(routeLabel),
        success: false,
        errorMessage:
          error instanceof Error ? error.message : "Unknown error",
      });
    }

    return handleApiError(routeLabel, error);
  }
}

/**
 * Admin API guard:
 * - authenticated user
 * - Active Admin permission
 * - audits mutating requests by default
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
  let email = "unknown";
  const method = request.method.toUpperCase();
  const shouldAudit =
    options?.audit ?? (method === "POST" || method === "PATCH" || method === "DELETE");

  try {
    const context = await requireAdminAccess();
    email = context.loggedInEmail;
    const data = await handler(context, request);

    if (shouldAudit) {
      await writeAuditLog({
        userEmail: email,
        action: method,
        entityName: options?.entityName ?? entityFromRouteLabel(routeLabel),
        itemId: extractItemId(data),
        success: true,
      });
    }

    return NextResponse.json(data);
  } catch (error) {
    if (shouldAudit) {
      await writeAuditLog({
        userEmail: email,
        action: method,
        entityName: options?.entityName ?? entityFromRouteLabel(routeLabel),
        success: false,
        errorMessage:
          error instanceof Error ? error.message : "Unknown error",
      });
    }

    return handleApiError(routeLabel, error);
  }
}

export function handleApiError(
  routeLabel: string,
  error: unknown,
): NextResponse {
  const payload = toPublicErrorPayload(error);

  if (payload.status >= 500) {
    console.error(`[${routeLabel}]`, error);
  }

  return NextResponse.json({ error: payload.error }, { status: payload.status });
}

export function jsonPublicError(
  status: 401 | 403 | 404 | 500,
): NextResponse {
  const error =
    status === 404
      ? PUBLIC_ERROR_MESSAGES.notFound
      : status === 500
        ? PUBLIC_ERROR_MESSAGES.serverError
        : PUBLIC_ERROR_MESSAGES.accessDenied;

  return NextResponse.json({ error }, { status });
}
