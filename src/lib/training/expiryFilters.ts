/**
 * Shared expiry colour / status model — single source of truth for all portals.
 *
 * Grey  = missing date → Records to Review
 * Red   = expired (< 0) or urgent (0–90 days / within 3 months)
 * Amber = upcoming (91–180 days / within 6 months)
 * Green = valid (181+ days / 6 months or more, open-ended)
 *
 * The 3- and 6-month filters are cumulative; 6m-plus is open-ended from day 181.
 */

export type ExpiryStatusCode =
  | "missing"
  | "expired"
  | "urgent"
  | "upcoming"
  | "valid";

export type ExpiryColour = "grey" | "red" | "amber" | "green";

export type ExpiryStatusLabel =
  | "Not applicable"
  | "Expired"
  | "Expiring soon"
  | "Compliant";

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
  /** Expires in 181+ days, with no upper bound. */
  | "6m-plus"
  /** @deprecated Use "6m-plus". */
  | "9m-plus"
  /** @deprecated Use "6m-plus". */
  | "within-9m"
  /** @deprecated Use "within-3m" or "urgent" */
  | "expiring-3m"
  /** @deprecated Use "within-6m" */
  | "expiring-6m"
  /** @deprecated Use "6m-plus" */
  | "expiring-9m";

const MS_PER_DAY = 86_400_000;

/** Urgent / within 3 months (inclusive) — red. */
export const EXPIRY_URGENT_DAYS = 90;
/** End of 6-month amber window / start of green (open-ended). */
export const EXPIRY_WITHIN_6M_DAYS = 180;
/**
 * Upcoming (amber) status ends here; valid/green starts the next day.
 * Kept as alias of EXPIRY_WITHIN_6M_DAYS so callers share one constant.
 */
export const EXPIRY_UPCOMING_DAYS = EXPIRY_WITHIN_6M_DAYS;

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
      label: "Not applicable",
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
      label: "Expiring soon",
      colour: "red",
      daysUntilExpiry: daysUntilExpiryValue,
    };
  }

  if (daysUntilExpiryValue <= EXPIRY_UPCOMING_DAYS) {
    return {
      status: "upcoming",
      label: "Expiring soon",
      colour: "amber",
      daysUntilExpiry: daysUntilExpiryValue,
    };
  }

  return {
    status: "valid",
    label: "Compliant",
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
  if (
    filter === "within-9m" ||
    filter === "expiring-9m" ||
    filter === "9m-plus"
  ) {
    return "6m-plus";
  }
  return filter;
}

/**
 * Filter matcher for a single expiry date.
 * within-3m / within-6m include non-expired dates inside their cumulative windows.
 * 6m-plus is open-ended from day 181.
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

  if (normalized === "missing") return status === "missing";
  if (normalized === "expired") return status === "expired";
  if (normalized === "urgent") return status === "urgent";
  if (normalized === "upcoming") return status === "upcoming";
  if (normalized === "valid") return status === "valid";

  if (normalized === "within-3m") {
    return days !== null && days >= 0 && days <= EXPIRY_URGENT_DAYS;
  }
  if (normalized === "within-6m") {
    return days !== null && days >= 0 && days <= EXPIRY_WITHIN_6M_DAYS;
  }
  if (normalized === "6m-plus") {
    return days !== null && days > EXPIRY_UPCOMING_DAYS;
  }

  return true;
}

/**
 * True when any date in the row matches the filter (matrix row-level matching).
 */
export function matchesAnyExpiryFilter(
  dates: Array<string | null | undefined>,
  filter: ExpiryFilter,
  now = new Date(),
): boolean {
  const normalized = normalizeExpiryFilter(filter);
  if (normalized === "all") return true;
  if (normalized === "missing") {
    return dates.length === 0 || dates.every((d) => !d?.trim());
  }
  if (dates.length === 0) return false;
  return dates.some((date) => matchesExpiryFilter(date, normalized, now));
}

/** Earliest non-null date among candidates (ISO-ish strings). */
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

/** Legend copy for UIs — keep in sync with getExpiryStatus. */
export const EXPIRY_STATUS_LEGEND: ReadonlyArray<{
  status: ExpiryStatusCode;
  label: ExpiryStatusLabel;
  colour: ExpiryColour;
  description: string;
}> = [
  {
    status: "valid",
    label: "Compliant",
    colour: "green",
    description: "6 months or more remaining (181+ days; no upper limit)",
  },
  {
    status: "upcoming",
    label: "Expiring soon",
    colour: "amber",
    description: "Expires in 91–180 days (within 6 months)",
  },
  {
    status: "urgent",
    label: "Expiring soon",
    colour: "red",
    description: "Expires within 3 months (0–90 days)",
  },
  {
    status: "expired",
    label: "Expired",
    colour: "red",
    description: "Past the expiry date",
  },
  {
    status: "missing",
    label: "Not applicable",
    colour: "grey",
    description: "No expiry date recorded",
  },
];
