import "server-only";

import { requireAuthenticatedEmail } from "@/lib/auth/session";
import {
  AccessDeniedError,
  getAdminContext,
  getCustomerContext,
} from "@/lib/services/customerContextService";
import type { AdminContext, CustomerContext } from "@/types/models";

const COMPANY_ID_KEYS = [
  "companyId",
  "company_id",
  "CompanyId",
  "selectedCompanyId",
] as const;

/**
 * Customer portal security: active customer-side permission + assigned company only.
 * Supports Training Manager / Supervisor / Candidate via Permissions List.
 * Client-supplied companyId is rejected.
 */
export async function requireCustomerAccess(
  request?: Request,
): Promise<CustomerContext> {
  const email = await requireAuthenticatedEmail();

  if (request) {
    rejectClientCompanyOverride(request);
    rejectCustomerExportAttempt(request);
  }

  // getCustomerContext enforces Active + canAccessCustomer and resolves company/scope.
  return getCustomerContext(email);
}

/**
 * Admin portal security: active Admin permission only (literal Admin or Training Manager legacy).
 */
export async function requireAdminAccess(): Promise<AdminContext> {
  const email = await requireAuthenticatedEmail();
  // getAdminContext enforces Active + canAccessAdmin.
  return getAdminContext(email);
}

/**
 * Customers must never supply a company override in query or JSON body.
 */
export async function rejectClientCompanyOverride(
  request: Request,
): Promise<void> {
  const url = new URL(request.url);
  for (const key of COMPANY_ID_KEYS) {
    if (url.searchParams.has(key)) {
      throw new AccessDeniedError(
        "You do not have access to this portal.",
      );
    }
  }

  const method = request.method.toUpperCase();
  if (method === "GET" || method === "HEAD") {
    return;
  }

  try {
    const clone = request.clone();
    const contentType = clone.headers.get("content-type") ?? "";
    if (!contentType.includes("application/json")) {
      return;
    }
    const body = (await clone.json()) as Record<string, unknown>;
    for (const key of COMPANY_ID_KEYS) {
      if (key in body && body[key] !== undefined && body[key] !== null) {
        throw new AccessDeniedError(
          "You do not have access to this portal.",
        );
      }
    }
  } catch (error) {
    if (error instanceof AccessDeniedError) {
      throw error;
    }
    // Non-JSON bodies are ignored here.
  }
}

/**
 * Block customer Excel/CSV export style requests.
 */
export function rejectCustomerExportAttempt(request: Request): void {
  const url = new URL(request.url);
  const format = url.searchParams.get("format")?.toLowerCase();
  const exportFlag = url.searchParams.get("export")?.toLowerCase();

  if (
    format === "csv" ||
    format === "excel" ||
    format === "xlsx" ||
    exportFlag === "1" ||
    exportFlag === "true" ||
    exportFlag === "csv" ||
    exportFlag === "excel"
  ) {
    throw new AccessDeniedError();
  }

  const path = url.pathname.toLowerCase();
  if (
    path.includes("/export") ||
    path.endsWith(".csv") ||
    path.endsWith(".xlsx") ||
    path.endsWith(".xls")
  ) {
    throw new AccessDeniedError();
  }
}

export function assertCompanyMatch(
  resourceCompanyName: string | null | undefined,
  customerCompanyName: string,
): void {
  const left = (resourceCompanyName ?? "").trim().toLowerCase();
  const right = customerCompanyName.trim().toLowerCase();
  if (!left || left !== right) {
    throw new AccessDeniedError(
      "You do not have access to this portal.",
    );
  }
}

export function assertCustomerVisible(visible: boolean): void {
  if (!visible) {
    throw new AccessDeniedError(
      "You do not have access to this portal.",
    );
  }
}
