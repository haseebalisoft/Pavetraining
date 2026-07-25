import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminOffer,
  listAdminOffers,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/offers",
    async () => ({ records: await listAdminOffers(companyName) }),
    { errorMessage: "Failed to load offers" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/offers",
    async (_context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminOffer(body);
      return { record };
    },
    { errorMessage: "Failed to create offer" },
    request,
  );
}
