import { withAdminApi } from "@/lib/api/adminApi";
import { listAdminNporsCategoryOptions } from "@/lib/services/nporsCategoriesService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/npors-categories",
    async () => ({ options: await listAdminNporsCategoryOptions() }),
    { errorMessage: "Failed to load NPORS categories" },
    request,
  );
}
