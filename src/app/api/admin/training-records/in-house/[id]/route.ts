import { withAdminApi } from "@/lib/api/adminApi";
import { updateAdminRegister } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/training-records/in-house/[id]",
    async (_adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      // In-House is standalone — no Training Matrix sync.
      const record = await updateAdminRegister(
        "inHouseCertificates",
        id,
        body,
      );
      return { record, matrixSync: null };
    },
    { errorMessage: "Failed to update In-House record" },
    request,
  );
}
