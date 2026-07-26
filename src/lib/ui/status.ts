/**
 * Shared premium UI primitives for PAVE portals.
 */

import {
  getExpiryStatus,
  type ExpiryStatusCode,
} from "@/lib/training/expiryFilters";

export type StatusTone =
  | "neutral"
  | "ok"
  | "warn"
  | "danger"
  | "info"
  | "missing";

export function toneForExpiryStatus(status: ExpiryStatusCode): StatusTone {
  switch (status) {
    case "missing":
      return "missing";
    case "expired":
    case "urgent":
      return "danger";
    case "upcoming":
      return "warn";
    case "valid":
      return "ok";
  }
}

export function toneForExpiry(expiry: string | null | undefined): StatusTone {
  return toneForExpiryStatus(getExpiryStatus(expiry).status);
}

export function toneForOutcome(
  outcome: string | null | undefined,
): StatusTone {
  if (!outcome) return "neutral";
  if (/pass/i.test(outcome)) return "ok";
  if (/fail/i.test(outcome)) return "danger";
  return "neutral";
}

export function toneForYesNo(value: boolean): StatusTone {
  return value ? "ok" : "neutral";
}
