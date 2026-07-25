import { withAdminApi } from "@/lib/api/adminApi";
import { listAdminDocuments } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const params = new URL(request.url).searchParams;
  const companyName = params.get("companyName");
  const candidate = params.get("candidate");
  const documentType = params.get("documentType");
  const visibility = params.get("customerVisible");

  let customerVisible: boolean | null = null;
  if (visibility === "true" || visibility === "yes") {
    customerVisible = true;
  } else if (visibility === "false" || visibility === "no") {
    customerVisible = false;
  }

  return withAdminApi(
    "GET /api/admin/documents",
    async () => ({
      records: await listAdminDocuments({
        companyName,
        candidate,
        documentType,
        customerVisible,
      }),
    }),
    { errorMessage: "Failed to load documents" },
    request,
  );
}
