/**
 * Shared premium UI primitives for PAVE portals.
 */

export type StatusTone =
  | "neutral"
  | "ok"
  | "warn"
  | "danger"
  | "info"
  | "missing";

export function toneForExpiry(expiry: string | null | undefined): StatusTone {
  if (!expiry?.trim()) return "missing";
  const days = Math.ceil(
    (new Date(expiry).setHours(0, 0, 0, 0) - new Date().setHours(0, 0, 0, 0)) /
      86_400_000,
  );
  if (Number.isNaN(days) || days < 0) return "danger";
  if (days <= 90) return "danger";
  if (days <= 180) return "warn";
  return "ok";
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
