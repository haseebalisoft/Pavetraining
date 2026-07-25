import { NextResponse } from "next/server";

import { withCustomerApi as withCustomerApiGuard } from "@/lib/api/apiGuards";
import type { CustomerContext } from "@/types/models";

/**
 * Shared customer API guard: company always from Permissions List.
 * Rejects client companyId and export attempts.
 */
export async function withCustomerApi<T>(
  routeLabel: string,
  handler: (context: CustomerContext, request?: Request) => Promise<T>,
  options?: { errorMessage?: string; audit?: boolean; entityName?: string },
  request: Request = new Request("http://localhost"),
): Promise<NextResponse> {
  return withCustomerApiGuard(
    routeLabel,
    async (context, req) => handler(context, req),
    request,
    {
      audit: options?.audit,
      entityName: options?.entityName,
    },
  );
}

export {
  AccessDeniedError,
  NotFoundError,
  UnauthorizedError,
  ValidationError,
} from "@/lib/services/errorHandler";
