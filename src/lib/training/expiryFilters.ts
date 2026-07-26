/**
 * Shared expiry colour / status model — single source of truth for all portals.
 *
 * Grey  = missing date → Records to Review
 * Red   = expired (< 0) or urgent (0–90 days / within 3 months)
 * Amber = upcoming (91–270 days / within 9 months)
 * Green = valid (271+ days / more than 9 months)
 *
 * Day-window filters (within-3m / 6m / 9m) are separate from colour buckets.
 */

export type ExpiryStatusCode =
  | "missing"
  | "expired"
  | "urgent"
  | "upcoming"
  | "valid";

export type ExpiryColour = "grey" | "red" | "amber" | "green";

export type ExpiryStatusLabel =
  | "Records to Review"
  | "Expired"
  | "Urgent"
  | "Upcoming"
  | "Valid";

export interface ExpiryStatus {
  status: ExpiryStatusCode;
  label: ExpiryStatusLabel;
  colour: ExpiryColour;
  daysUntilExpiry: number | null;
}

/** @deprecated Prefer ExpiryStatusCode via getExpiryStatus().status */
export type ExpiryTone =
  | "missing"
  | "expired"
  | "critical"
  | "warning"
  | "ok";

export type ExpiryFilter =
  | "all"
  | "missing"
  | "expired"
  | "urgent"
  | "upcoming"
  | "valid"
  /** Expires in 0–90 days (not yet expired). */
  | "within-3m"
  /** Expires in 0–180 days (not yet expired). */
  | "within-6m"
  /** Expires in 0–270 days (not yet expired). */
  | "within-9m"
  /** @deprecated Use "within-3m" or "urgent" */
  | "expiring-3m"
  /** @deprecated Use "within-6m" */
  | "expiring-6m"
  /** @deprecated Use "within-9m" */
  | "expiring-9m";

const MS_PER_DAY = 86_400_000;

/** Urgent / within 3 months (inclusive). */
export const EXPIRY_URGENT_DAYS = 90;
/** End of 6-month window filter. */
export const EXPIRY_WITHIN_6M_DAYS = 180;
/** Upcoming band ends / within 9 months (inclusive). */
export const EXPIRY_UPCOMING_DAYS = 270;

export function daysUntilExpiry(
  expiry: string | null | undefined,
  now = new Date(),
): number | null {
  if (!expiry?.trim()) {
    return null;
  }

  const expiryDate = new Date(expiry);
  if (Number.isNaN(expiryDate.getTime())) {
    return null;
  }

  const startOfToday = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
  );
  const startOfExpiry = new Date(
    expiryDate.getFullYear(),
    expiryDate.getMonth(),
    expiryDate.getDate(),
  );

  return Math.round(
    (startOfExpiry.getTime() - startOfToday.getTime()) / MS_PER_DAY,
  );
}

/**
 * Shared expiry classifier — use this everywhere dates are coloured or filtered.
 * Do not reimplement these thresholds in components.
 */
export function getExpiryStatus(
  date: string | null | undefined,
  now = new Date(),
): ExpiryStatus {
  const daysUntilExpiryValue = daysUntilExpiry(date, now);

  if (daysUntilExpiryValue === null) {
    return {
      status: "missing",
      label: "Records to Review",
      colour: "grey",
      daysUntilExpiry: null,
    };
  }

  if (daysUntilExpiryValue < 0) {
    return {
      status: "expired",
      label: "Expired",
      colour: "red",
      daysUntilExpiry: daysUntilExpiryValue,
    };
  }

  if (daysUntilExpiryValue <= EXPIRY_URGENT_DAYS) {
    return {
      status: "urgent",
      label: "Urgent",
      colour: "red",
      daysUntilExpiry: daysUntilExpiryValue,
    };
  }

  if (daysUntilExpiryValue <= EXPIRY_UPCOMING_DAYS) {
    return {
      status: "upcoming",
      label: "Upcoming",
      colour: "amber",
      daysUntilExpiry: daysUntilExpiryValue,
    };
  }

  return {
    status: "valid",
    label: "Valid",
    colour: "green",
    daysUntilExpiry: daysUntilExpiryValue,
  };
}

/** Maps status → legacy tone tokens used by some CSS modules. */
export function getExpiryTone(
  expiry: string | null | undefined,
  now = new Date(),
): ExpiryTone {
  switch (getExpiryStatus(expiry, now).status) {
    case "missing":
      return "missing";
    case "expired":
      return "expired";
    case "urgent":
      return "critical";
    case "upcoming":
      return "warning";
    case "valid":
      return "ok";
  }
}

function normalizeExpiryFilter(filter: ExpiryFilter): ExpiryFilter {
  if (filter === "expiring-3m") return "within-3m";
  if (filter === "expiring-6m") return "within-6m";
  if (filter === "expiring-9m") return "within-9m";
  return filter;
}

/**
 * Filter matcher for a single expiry date.
 * Window filters (within-*) include dates that expire inside the window and
 * exclude already-expired dates (use "expired" for those).
 */
export function matchesExpiryFilter(
  expiry: string | null | undefined,
  filter: ExpiryFilter,
  now = new Date(),
): boolean {
  const normalized = normalizeExpiryFilter(filter);
  if (normalized === "all") {
    return true;
  }

  const days = daysUntilExpiry(expiry, now);
  const { status } = getExpiryStatus(expiry, now);

  if (normalized === "within-3m") {
    return days !== null && days >= 0 && days <= EXPIRY_URGENT_DAYS;
  }
  if (normalized === "within-6m") {
    return days !== null && days >= 0 && days <= EXPIRY_WITHIN_6M_DAYS;
  }
  if (normalized === "within-9m") {
    return days !== null && days >= 0 && days <= EXPIRY_UPCOMING_DAYS;
  }

  return status === normalized;
}

/**
 * True when any of the provided dates match the filter (row-level matrix filters).
 */
export function matchesAnyExpiryFilter(
  dates: Array<string | null | undefined>,
  filter: ExpiryFilter,
  now = new Date(),
): boolean {
  const normalized = normalizeExpiryFilter(filter);
  if (normalized === "all") {
    return true;
  }
  return dates.some((date) => matchesExpiryFilter(date, normalized, now));
}

/** Earliest parseable date among values (null if none). */
export function earliestExpiryDate(
  dates: Array<string | null | undefined>,
): string | null {
  let best: { ms: number; value: string } | null = null;
  for (const value of dates) {
    if (!value?.trim()) continue;
    const ms = new Date(value).getTime();
    if (Number.isNaN(ms)) continue;
    if (!best || ms < best.ms) {
      best = { ms, value };
    }
  }
  return best?.value ?? null;
}

export function formatDisplayDate(value: string | null | undefined): string {
  if (!value?.trim()) {
    return "—";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(date);
}

/** Legend copy for UIs — keep in sync with getExpiryStatus. */
export const EXPIRY_STATUS_LEGEND: ReadonlyArray<{
  status: ExpiryStatusCode;
  label: ExpiryStatusLabel;
  colour: ExpiryColour;
  description: string;
}> = [
  {
    status: "expired",
    label: "Expired",
    colour: "red",
    description: "Past the expiry date",
  },
  {
    status: "urgent",
    label: "Urgent",
    colour: "red",
    description: "Expires within 3 months (0–90 days)",
  },
  {
    status: "upcoming",
    label: "Upcoming",
    colour: "amber",
    description: "Expires within 3–9 months (91–270 days)",
  },
  {
    status: "valid",
    label: "Valid",
    colour: "green",
    description: "More than 9 months remaining (271+ days)",
  },
  {
    status: "missing",
    label: "Records to Review",
    colour: "grey",
    description: "No expiry date recorded",
  },
];
