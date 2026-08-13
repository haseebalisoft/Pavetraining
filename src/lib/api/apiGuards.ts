import { NextResponse } from "next/server";

import {
  PUBLIC_ERROR_MESSAGES,
  toPublicErrorPayload,
} from "@/lib/services/errorHandler";
import {
  extractItemId,
  extractItemName,
  logAdminCreate,
  logAdminDelete,
  logAdminUpdate,
  sanitizeAuditError,
  writeAuditLog,
} from "@/lib/services/auditLogService";
import {
  requireAdminAccess,
  requireCustomerAccess,
} from "@/lib/services/securityService";
import type { AdminContext, CustomerContext } from "@/types/models";

import {
  AccessDeniedError,
  NotFoundError,
  UnauthorizedError,
  ValidationError,
} from "@/lib/services/errorHandler";

export {
  AccessDeniedError,
  NotFoundError,
  UnauthorizedError,
  ValidationError,
};

function entityFromRouteLabel(routeLabel: string): string {
  return routeLabel.replace(/^(GET|POST|PATCH|PUT|DELETE)\s+/i, "").trim();
}

function actionFromRequest(request: Request): string {
  return request.method.toUpperCase();
}

function inferEntityType(routeLabel: string, fallback?: string): string {
  if (fallback) return fallback;
  const path = entityFromRouteLabel(routeLabel).toLowerCase();
  if (path.includes("compan")) return "Company List";
  if (path.includes("workforce") || path.includes("candidate")) {
    return "Workforce";
  }
  if (path.includes("document")) return "Customer Documents";
  if (path.includes("event")) return "Events";
  if (path.includes("permission")) return "Permissions";
  if (path.includes("training-matrix") || path.includes("matrix")) {
    return "Training Matrix";
  }
  if (path.includes("npors")) return "NPORS Register";
  if (path.includes("eusr")) return "EUSR Register";
  if (path.includes("streetworks") || path.includes("nrswa")) {
    return "NRSWA Register";
  }
  if (path.includes("in-house")) return "In-House Certificates";
  if (path.includes("nvq")) return "NVQ Register";
  if (path.includes("offer")) return "Offers";
  if (path.includes("settings")) return "Portal Settings";
  if (path.includes("notification")) return "Notifications";
  return entityFromRouteLabel(routeLabel);
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
  const method = request.method.toUpperCase();
  const shouldAudit =
    options?.audit ??
    (method === "POST" || method === "PATCH" || method === "PUT" || method === "DELETE");
  try {
    const context = await requireCustomerAccess(request);
    email = context.loggedInEmail;
    const data = await handler(context, request);

    if (shouldAudit) {
      await writeAuditLog({
        userEmail: email,
        action: actionFromRequest(request),
        entityName: options?.entityName ?? entityFromRouteLabel(routeLabel),
        entityType: inferEntityType(routeLabel, options?.entityName),
        itemId: extractItemId(data),
        success: true,
        roleType: context.roleLabel,
        company: context.companyName,
        request,
      });
    }

    return NextResponse.json(data);
  } catch (error) {
    if (shouldAudit) {
      await writeAuditLog({
        userEmail: email,
        action: actionFromRequest(request),
        entityName: options?.entityName ?? entityFromRouteLabel(routeLabel),
        entityType: inferEntityType(routeLabel, options?.entityName),
        success: false,
        errorMessage: sanitizeAuditError(error),
        roleType: undefined,
        request,
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
  const entityType = inferEntityType(routeLabel, options?.entityName);

  try {
    const context = await requireAdminAccess();
    email = context.loggedInEmail;
    const data = await handler(context, request);
    const itemId = extractItemId(data);
    const entityName = options?.entityName ?? extractItemName(data) ?? entityType;

    if (shouldAudit) {
      if (method === "POST") {
        await logAdminCreate({
          userEmail: email,
          entityType,
          entityId: itemId,
          entityName,
          request,
        });
      } else if (method === "PATCH" || method === "PUT") {
        await logAdminUpdate({
          userEmail: email,
          entityType,
          entityId: itemId,
          entityName,
          request,
        });
      } else if (method === "DELETE") {
        await logAdminDelete({
          userEmail: email,
          entityType,
          entityId: itemId,
          entityName,
          request,
        });
      } else {
        await writeAuditLog({
          userEmail: email,
          action: method,
          entityName: options?.entityName ?? entityFromRouteLabel(routeLabel),
          entityType,
          itemId,
          success: true,
          roleType: context.roleLabel || "Admin",
          request,
        });
      }
    }

    return NextResponse.json(data);
  } catch (error) {
    if (shouldAudit) {
      await writeAuditLog({
        userEmail: email,
        action:
          method === "POST"
            ? "ADMIN_CREATE"
            : method === "PATCH" || method === "PUT"
              ? "ADMIN_UPDATE"
              : method === "DELETE"
                ? "ADMIN_DELETE"
                : method,
        entityName: options?.entityName ?? entityFromRouteLabel(routeLabel),
        entityType,
        success: false,
        errorMessage: sanitizeAuditError(error),
        roleType: "Admin",
        request,
      });
    }

    return handleApiError(routeLabel, error);
  }
}

/**
 * Same as `withAdminApi` but rejects Training Managers with 403. Use for
 * routes that only pure SharePoint Admins (or hardcoded protected admins)
 * are allowed to hit — Bulk Upload commits, Permissions writes, etc.
 *
 * Non-SP admins hitting these routes get a normal AccessDeniedError JSON
 * response with `status: 403`; the nav already hides these pages for them,
 * so the API layer is defence-in-depth against direct HTTP requests.
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
  return withAdminApi(
    routeLabel,
    async (context, req) => {
      if (!context.isSharePointAdmin) {
        throw new AccessDeniedError(
          "This action requires full Admin access. Training Managers cannot perform it.",
        );
      }
      return handler(context, req);
    },
    options,
    request,
  );
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
