import "server-only";

import {
  earliestExpiryDate,
  getExpiryStatus,
  type ExpiryStatus,
  type ExpiryStatusCode,
} from "@/lib/training/expiryFilters";

/**
 * SharePoint Training Matrix OverallStatus choice values
 * (Valid | Expiring Soon | Expired | Missing Data).
 */
export type SharePointOverallStatus =
  | "Valid"
  | "Expiring Soon"
  | "Expired"
  | "Missing Data";

export interface MatrixStatusComputation {
  nextExpiryDate: string | null;
  overallStatus: SharePointOverallStatus;
  needsReview: boolean;
  expiryStatus: ExpiryStatus;
}

/**
 * Maps shared expiry colour logic → SharePoint OverallStatus choices.
 * Upcoming (91–270) stays Valid; Urgent (0–90) is Expiring Soon.
 */
export function toSharePointOverallStatus(
  code: ExpiryStatusCode,
): SharePointOverallStatus {
  switch (code) {
    case "missing":
      return "Missing Data";
    case "expired":
      return "Expired";
    case "urgent":
      return "Expiring Soon";
    case "upcoming":
    case "valid":
      return "Valid";
  }
}

/** Records to Review when date is missing or already expired. */
export function shouldNeedsReview(code: ExpiryStatusCode): boolean {
  return code === "missing" || code === "expired";
}

/**
 * Computes NextExpiryDate, OverallStatus, and NeedsReview from date set.
 */
export function computeMatrixStatusFromDates(
  dates: Array<string | null | undefined>,
  now = new Date(),
): MatrixStatusComputation {
  const nextExpiryDate = earliestExpiryDate(dates);
  const expiryStatus = getExpiryStatus(nextExpiryDate, now);
  return {
    nextExpiryDate,
    overallStatus: toSharePointOverallStatus(expiryStatus.status),
    needsReview: shouldNeedsReview(expiryStatus.status),
    expiryStatus,
  };
}

export function describeExpiryForSync(date: string | null | undefined): string {
  const status = getExpiryStatus(date);
  if (status.daysUntilExpiry === null) {
    return status.label;
  }
  return `${status.label} (${status.daysUntilExpiry}d)`;
}
