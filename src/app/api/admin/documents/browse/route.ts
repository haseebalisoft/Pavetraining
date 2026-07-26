import { withAdminApi } from "@/lib/api/adminApi";
import { listAdminDocumentsAtPath } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

function parsePath(request: Request): string[] {
  const params = new URL(request.url).searchParams;
  const raw = params.get("path");
  if (!raw?.trim()) {
    return [];
  }
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (Array.isArray(parsed)) {
      return parsed.map((part) => String(part).trim()).filter(Boolean);
    }
  } catch {
    // Fall through to slash-separated path.
  }
  return raw
    .split("/")
    .map((part) => {
      try {
        return decodeURIComponent(part).trim();
      } catch {
        return part.trim();
      }
    })
    .filter(Boolean);
}

/**
 * Browse one level of the Customer Documents library folder tree.
 * GET /api/admin/documents/browse?path=[] 
 * GET /api/admin/documents/browse?path=["C00024 - Murphy Plant Ltd"]
 */
export async function GET(request: Request) {
  const path = parsePath(request);
  return withAdminApi(
    "GET /api/admin/documents/browse",
    async () => ({
      path,
      records: await listAdminDocumentsAtPath(path),
    }),
    { errorMessage: "Failed to browse Customer Documents" },
    request,
  );
}
