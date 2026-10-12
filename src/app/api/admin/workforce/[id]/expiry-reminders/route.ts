import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { getSettings, updateSettings } from "@/lib/services/settingsService";

export const dynamic = "force-dynamic";

/**
 * Switch one candidate's expiry reminder emails on or off. Stored as a list of
 * Workforce ids in Portal Settings (the Workforce list has no column for it).
 */
export async function PUT(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PUT /api/admin/workforce/[id]/expiry-reminders",
    async (ctx, req) => {
      const body = (await req.json().catch(() => null)) as {
        enabled?: unknown;
      } | null;
      if (typeof body?.enabled !== "boolean") {
        throw new ValidationError("enabled must be true or false.");
      }
      const workforceId = id.trim();
      // Read fresh so a concurrent toggle elsewhere isn't overwritten.
      const current = await getSettings({ bypassCache: true });
      const off = new Set(current.settings.expiryRemindersOffWorkforceIds);
      if (body.enabled) off.delete(workforceId);
      else off.add(workforceId);

      const saved = await updateSettings(
        { ...current.settings, expiryRemindersOffWorkforceIds: Array.from(off) },
        { actorEmail: ctx.loggedInEmail },
      );
      if (!saved.listConfigured) {
        throw new ValidationError(
          "Portal Settings list is not available, so the change could not be saved.",
        );
      }
      return {
        enabled: !saved.settings.expiryRemindersOffWorkforceIds.includes(
          workforceId,
        ),
      };
    },
    { errorMessage: "Failed to update expiry reminders" },
    request,
  );
}
