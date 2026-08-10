/**
 * Pure logic for applying/removing training-register expiry data on a single
 * Training Matrix column.
 *
 * Dependency-free (no server-only / Graph imports) so scripts can exercise the
 * forward-apply and delete-recompute rules under Node. The Graph I/O wrapper
 * lives in trainingMatrixSyncService.ts, which composes these helpers. See
 * workforceMatrixSync.ts for the sibling row-matching / link-field helpers.
 */

export type NormalizedOutcome = "Pass" | "Fail" | null;

function parseDateMs(value: string | null | undefined): number | null {
  if (!value?.trim()) return null;
  const ms = new Date(value).getTime();
  return Number.isNaN(ms) ? null : ms;
}

/**
 * Apply register expiry onto an existing matrix/workforce date only when:
 * - outcome is Pass
 * - incoming date exists
 * - existing is empty OR incoming is clearly newer (later or equal)
 * Fail never extends. Missing outcome does not extend.
 */
export function shouldApplyPassExpiry(
  existing: string | null | undefined,
  incoming: string | null | undefined,
  outcome: NormalizedOutcome,
): boolean {
  if (outcome !== "Pass") return false;
  if (!incoming?.trim()) return false;
  const incomingMs = parseDateMs(incoming);
  if (incomingMs === null) return false;
  const existingMs = parseDateMs(existing);
  if (existingMs === null) return true;
  return incomingMs >= existingMs;
}

/** Latest Pass expiry among a set of source records for one matrix column. */
export function latestPassExpiry(
  records: ReadonlyArray<{
    trainingOutcome: NormalizedOutcome;
    expiry: string | null;
  }>,
): string | null {
  let best: { ms: number; value: string } | null = null;
  for (const record of records) {
    if (record.trainingOutcome !== "Pass") continue;
    if (!record.expiry?.trim()) continue;
    const ms = parseDateMs(record.expiry);
    if (ms === null) continue;
    if (!best || ms > best.ms) {
      best = { ms, value: record.expiry };
    }
  }
  return best?.value ?? null;
}

export type MatrixFieldRemovalAction =
  | "recomputed"
  | "cleared"
  | "preserved_manual_override"
  | "unchanged";

export interface MatrixFieldRemovalResult {
  action: MatrixFieldRemovalAction;
  /** Value to write; null means clear the column. Ignore when action is "unchanged". */
  nextValue: string | null;
  /** True when the caller should force the row into Needs Review. */
  forceNeedsReview: boolean;
  /** Audit-log / toast note explaining what happened and why. */
  note: string;
}

/**
 * Recompute one Training Matrix expiry column after a source register record
 * was deleted or downgraded to Fail.
 *  - Manual-override columns are NEVER modified — the existing value is kept
 *    and flagged so an admin knows the source record behind it is gone.
 *  - Otherwise, recompute strictly from the REMAINING valid (Pass) source
 *    records: if one still supports a date, use it — even if that date is
 *    OLDER than the value being replaced. Unlike incremental sync (which only
 *    ever extends forward), deletion recompute is authoritative: it reflects
 *    exactly what remains, not the newest value ever seen.
 *  - If nothing remains, clear a system-synced value rather than leave a
 *    stale date with no backing record, and force Needs Review.
 */
export function computeMatrixFieldAfterRemoval(params: {
  header: string;
  currentValue: string | null;
  isManualOverride: boolean;
  remainingRecords: ReadonlyArray<{
    trainingOutcome: NormalizedOutcome;
    expiry: string | null;
  }>;
}): MatrixFieldRemovalResult {
  const { header, currentValue, isManualOverride, remainingRecords } = params;

  if (isManualOverride) {
    if (!currentValue) {
      return {
        action: "unchanged",
        nextValue: null,
        forceNeedsReview: false,
        note: `${header}: manual override with no value — nothing to preserve.`,
      };
    }
    return {
      action: "preserved_manual_override",
      nextValue: currentValue,
      forceNeedsReview: false,
      note: `${header}: Manual Override / Source Deleted — kept ${currentValue}; the source record behind it was removed.`,
    };
  }

  const recomputedValue = latestPassExpiry(remainingRecords);

  if (recomputedValue) {
    if (recomputedValue === currentValue) {
      return {
        action: "unchanged",
        nextValue: currentValue,
        forceNeedsReview: false,
        note: `${header}: unchanged — a remaining source record still supports ${currentValue}.`,
      };
    }
    return {
      action: "recomputed",
      nextValue: recomputedValue,
      forceNeedsReview: false,
      note: `${header}: recomputed from remaining source record(s) → ${recomputedValue}.`,
    };
  }

  if (!currentValue) {
    return {
      action: "unchanged",
      nextValue: null,
      forceNeedsReview: false,
      note: `${header}: already blank.`,
    };
  }

  return {
    action: "cleared",
    nextValue: null,
    forceNeedsReview: true,
    note: `${header}: cleared — the system-synced value had no remaining source record. Flagged Needs Review.`,
  };
}
