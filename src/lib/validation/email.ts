/**
 * Shared email validation — the single source of truth for the whole portal.
 *
 * Deliberately dependency-free: NO "server-only" and NO imports, so it is safe
 * to use from server code (services, route handlers), from client components,
 * and from the pure `node --test` unit runner (which loads this .ts file
 * directly with Node's native type-stripping).
 *
 * Design rules (see scripts/test-email-validation.mjs):
 *  - Accept ANY valid domain / TLD. .org, .co.uk, .training, long TLDs and
 *    subdomains all pass. We do NOT restrict to .com or an allowlist.
 *  - Trim whitespace and lowercase — emails are matched case-insensitively
 *    everywhere in this app (permissions, company match, OTP login).
 *  - Optional fields never block save when blank; required fields report a
 *    clear message only when the value is missing or malformed.
 */

/**
 * Pragmatic email shape: exactly one "@", at least one "." in the domain, and
 * no whitespace. This is intentionally permissive about the TLD/domain so that
 * .org, .co.uk, .training, long TLDs and subdomains are all accepted.
 */
export const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/** Trim + lowercase. Coerces null/undefined to "". */
export function normalizeEmail(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase();
}

/** True when `value` (after trim + lowercase) is a valid email address. */
export function isValidEmail(value: unknown): boolean {
  return EMAIL_REGEX.test(normalizeEmail(value));
}

export type EmailValidation =
  | { ok: true; email: string | null }
  | { ok: false; error: string };

/**
 * Validate one email form/API field.
 *  - optional (default) + blank  → { ok: true,  email: null }   (never blocks save)
 *  - required + blank            → { ok: false, error }
 *  - non-blank + invalid format  → { ok: false, error }
 *  - non-blank + valid           → { ok: true,  email: <normalized> }
 */
export function validateEmailField(
  value: unknown,
  options: { required?: boolean; label?: string } = {},
): EmailValidation {
  const label = options.label ?? "Email";
  const email = normalizeEmail(value);
  if (!email) {
    return options.required
      ? { ok: false, error: `${label} is required.` }
      : { ok: true, email: null };
  }
  if (!EMAIL_REGEX.test(email)) {
    return { ok: false, error: `${label} must be a valid email address.` };
  }
  return { ok: true, email };
}
