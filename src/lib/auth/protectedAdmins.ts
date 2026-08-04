import "server-only";

/**
 * Comma/semicolon-separated emails that cannot be deleted or deactivated
 * via the admin Permissions UI/API. Set in `.env.local` and Vercel:
 *
 *   PROTECTED_ADMIN_EMAILS=wayne@example.com,you@example.com
 *
 * Login is NOT hardcoded — these emails still need an Active SharePoint
 * Permissions row. This only prevents the app from removing that row.
 */
export function getProtectedAdminEmails(): Set<string> {
  const raw = process.env.PROTECTED_ADMIN_EMAILS?.trim() ?? "";
  if (!raw) return new Set();
  return new Set(
    raw
      .split(/[,:;]+/)
      .map((part) => part.trim().toLowerCase())
      .filter(Boolean),
  );
}

export function isProtectedAdminEmail(email: string | null | undefined): boolean {
  const normalized = String(email ?? "")
    .trim()
    .toLowerCase();
  if (!normalized) return false;
  return getProtectedAdminEmails().has(normalized);
}

export function assertNotProtectedAdmin(
  email: string | null | undefined,
  action: "delete" | "deactivate",
): void {
  if (!isProtectedAdminEmail(email)) return;
  const label = action === "delete" ? "deleted" : "set to Inactive";
  throw new Error(
    `This Permissions account is protected and cannot be ${label} from the app. Update SharePoint as a Site Owner only if you truly need to remove it.`,
  );
}
