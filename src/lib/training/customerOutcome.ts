import type { CustomerOutcome } from "@/types/models";

/**
 * Maps SharePoint training outcomes to customer-safe Pass/Fail only.
 * Admin-only or ambiguous values are treated as null (hidden).
 */
export function toCustomerOutcome(
  raw: string | null | undefined,
): CustomerOutcome | null {
  if (!raw?.trim()) {
    return null;
  }

  const value = raw.trim().toLowerCase();

  if (
    value === "pass" ||
    value === "passed" ||
    value === "p" ||
    value.startsWith("pass")
  ) {
    return "Pass";
  }

  if (
    value === "fail" ||
    value === "failed" ||
    value === "f" ||
    value.startsWith("fail")
  ) {
    return "Fail";
  }

  return null;
}
