import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { linkAdminMatrixToWorkforce } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

/**
 * Repair action for the Admin Training Matrix page. Attaches an existing
 * Matrix Update row to a Workforce candidate so it flips from
 * Orphan / Needs Review to Linked and becomes visible in the customer portal.
 * Body: { workforceId: string }
 */
export async function POST(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "POST /api/admin/training-matrix/[id]/link",
    async (_ctx, req) => {
      const body = (await req.json().catch(() => null)) as
        | { workforceId?: unknown }
        | null;
      const workforceId = String(body?.workforceId ?? "").trim();
      if (!workforceId) {
        throw new ValidationError("workforceId is required.");
      }
      const record = await linkAdminMatrixToWorkforce(id, workforceId);
      return { record };
    },
    { errorMessage: "Failed to link matrix row" },
    request,
  );
}
