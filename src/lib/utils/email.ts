/**
 * Shared email helpers for forms + API validation.
 * Accepts normal work domains (.com, .org, .co.uk, .training, subdomains, +tags).
 * Does not whitelist providers (Gmail vs corporate) and does not block uncommon TLDs.
 */

/** Practical check: local@domain with at least one dot in the domain, no spaces. */
const EMAIL_PATTERN =
  /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function normalizeEmail(value: string): string {
  return value.trim().toLowerCase();
}

/** True when value is a non-empty valid email. */
export function isValidEmail(value: string | null | undefined): boolean {
  if (!value?.trim()) return false;
  const email = normalizeEmail(value);
  if (email.length > 254) return false;
  if (email.includes("..")) return false;
  if (email.startsWith(".") || email.endsWith(".")) return false;
  if (email.includes("@.") || email.includes(".@")) return false;
  const at = email.indexOf("@");
  if (at <= 0 || at !== email.lastIndexOf("@")) return false;
  return EMAIL_PATTERN.test(email);
}

/**
 * Optional fields: blank → null; invalid → null when used as soft parse.
 */
export function optionalEmail(
  value: string | null | undefined,
): string | null {
  if (!value?.trim()) return null;
  const email = normalizeEmail(value);
  return isValidEmail(email) ? email : null;
}
