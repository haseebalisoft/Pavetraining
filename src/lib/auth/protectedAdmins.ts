import "server-only";

/**
 * Hardcoded emails that always retain Admin portal access, even if their
 * SharePoint Permissions row is missing or deleted.
 */
const ALWAYS_ADMIN_EMAILS = [
  "wayne.curry@pavetraining.co.uk",
] as const;

/**
 * Extra protect list via env (cannot delete/deactivate from app).
 * Set in `.env.local` / Vercel:
 *   PROTECTED_ADMIN_EMAILS=you@example.com,other@example.com
 */
export function getProtectedAdminEmails(): Set<string> {
  const fromEnv = (process.env.PROTECTED_ADMIN_EMAILS?.trim() ?? "")
    .split(/[,:;]+/)
    .map((part) => part.trim().toLowerCase())
    .filter(Boolean);

  return new Set([...ALWAYS_ADMIN_EMAILS, ...fromEnv]);
}

export function isAlwaysAdminEmail(
  email: string | null | undefined,
): boolean {
  const normalized = String(email ?? "")
    .trim()
    .toLowerCase();
  if (!normalized) return false;
  return (ALWAYS_ADMIN_EMAILS as readonly string[]).includes(normalized);
}

export function isProtectedAdminEmail(
  email: string | null | undefined,
): boolean {
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
