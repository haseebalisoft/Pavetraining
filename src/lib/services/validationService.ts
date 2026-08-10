import "server-only";

import { ValidationError } from "@/lib/services/errorHandler";
import { validateEmailField } from "@/lib/validation/email";

/**
 * Input validation helpers for admin/customer APIs.
 */

export function requireNonEmptyString(
  value: unknown,
  label: string,
): string {
  if (typeof value !== "string" || !value.trim()) {
    throw new ValidationError(`${label} is required.`);
  }
  return value.trim();
}

export function optionalTrimmedString(value: unknown): string | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  if (typeof value === "string") {
    return value.trim() || null;
  }
  return String(value).trim() || null;
}

export function parseBooleanStrict(value: unknown): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (["true", "yes", "1"].includes(normalized)) return true;
    if (["false", "no", "0"].includes(normalized)) return false;
  }
  throw new ValidationError("Boolean fields must be true/false or Yes/No.");
}

export function assertPassFailOutcome(value: unknown): "Pass" | "Fail" | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const text = String(value).trim();
  if (/^pass/i.test(text)) return "Pass";
  if (/^fail/i.test(text)) return "Fail";
  throw new ValidationError("Training outcome must be Pass or Fail.");
}

/** Required email → normalized (trim + lowercase), or throws ValidationError. */
export function assertEmail(value: unknown, label = "Email"): string {
  const result = validateEmailField(value, { required: true, label });
  if (!result.ok) throw new ValidationError(result.error);
  // required + ok ⇒ email is a non-null normalized string.
  return result.email as string;
}

/**
 * Optional email → normalized (trim + lowercase) or null when blank. Throws
 * ValidationError only when a non-blank value is malformed, so blank optional
 * email fields never block a save.
 */
export function normalizeOptionalEmail(
  value: unknown,
  label = "Email",
): string | null {
  const result = validateEmailField(value, { required: false, label });
  if (!result.ok) throw new ValidationError(result.error);
  return result.email;
}

export { ValidationError };
