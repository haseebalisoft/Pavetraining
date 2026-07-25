import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminEvent,
  listAdminEvents,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/events",
    async () => ({ records: await listAdminEvents(companyName) }),
    { errorMessage: "Failed to load events" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/events",
    async (_context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminEvent(body);
      return { record };
    },
    { errorMessage: "Failed to create event" },
    request,
  );
}
