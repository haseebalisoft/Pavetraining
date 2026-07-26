import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import {
  getDefaultSettings,
  getSettings,
  updateSettings,
} from "@/lib/services/settingsService";
import type { PortalSettings } from "@/types/portalSettings";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/settings",
    async () => getSettings({ bypassCache: true }),
    {
      errorMessage: "Failed to load portal settings",
      audit: false,
    },
    request,
  );
}

export async function PATCH(request: Request) {
  return withAdminApi(
    "PATCH /api/admin/settings",
    async (context, req) => {
      const body = (await req.json().catch(() => null)) as
        | { settings?: Partial<PortalSettings>; resetToDefaults?: boolean }
        | Partial<PortalSettings>
        | null;

      if (!body || typeof body !== "object") {
        throw new ValidationError("Settings payload is required.");
      }

      const resetToDefaults =
        "resetToDefaults" in body && Boolean(body.resetToDefaults);

      const partial = resetToDefaults
        ? getDefaultSettings()
        : "settings" in body && body.settings
          ? body.settings
          : (body as Partial<PortalSettings>);

      return updateSettings(partial, {
        actorEmail: context.loggedInEmail,
      });
    },
    {
      errorMessage: "Failed to update portal settings",
      entityName: "portal-settings",
    },
    request,
  );
}
