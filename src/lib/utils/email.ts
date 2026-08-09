/**
 * Shared email helpers for forms + API validation.
 * Accepts any RFC-like address with a real domain (including .org, .co.uk,
 * .training, subdomains). Does not restrict to known TLDs.
 */

const EMAIL_PATTERN =
  /^[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)+$/i;

export function normalizeEmail(value: string): string {
  return value.trim().toLowerCase();
}

/** True when value is a non-empty valid email. */
export function isValidEmail(value: string | null | undefined): boolean {
  if (!value?.trim()) return false;
  const email = normalizeEmail(value);
  if (email.length > 254) return false;
  if (email.includes("..")) return false;
  return EMAIL_PATTERN.test(email);
}

/**
 * Optional fields: blank → null; invalid → null when `strict` is false,
 * or callers throw via assert helpers.
 */
export function optionalEmail(
  value: string | null | undefined,
): string | null {
  if (!value?.trim()) return null;
  const email = normalizeEmail(value);
  return isValidEmail(email) ? email : null;
}
