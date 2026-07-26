import { withAdminApi } from "@/lib/api/adminApi";
import { getNotificationSettings } from "@/lib/services/notificationConfig";
import { listNotificationLogs } from "@/lib/services/notificationLogService";
import { listAdminDocuments } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const url = new URL(request.url);
  const failedOnly = url.searchParams.get("failedOnly") === "true";
  const top = Number(url.searchParams.get("top") ?? "80");

  return withAdminApi(
    "GET /api/admin/notifications/logs",
    async () => {
      const [logs, settings, documents] = await Promise.all([
        listNotificationLogs({
          top: Number.isFinite(top) ? Math.min(top, 200) : 80,
          failedOnly,
        }),
        Promise.resolve().then(() => getNotificationSettings()),
        listAdminDocuments().catch(() => []),
      ]);

      const documentStatus = documents
        .filter((doc) => !doc.isFolder)
        .slice(0, 40)
        .map((doc) => ({
          id: doc.id,
          name: doc.name,
          company: doc.company,
          documentType: doc.documentType,
          customerVisible: doc.customerVisible,
          notifyCustomer: doc.notifyCustomer,
          notificationSent: doc.notificationSent,
        }));

      return {
        settings,
        logs,
        failedLogs: logs.filter((entry) => entry.status === "failed"),
        documentStatus,
      };
    },
    {
      errorMessage: "Failed to load notification logs",
      audit: false,
    },
    request,
  );
}
