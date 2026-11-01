/**
 * Time-budget helpers for serial bulk commits (Training Matrix + registers).
 *
 * Serverless routes are killed at `maxDuration` with no chance to report. A
 * commit that checks its deadline between rows can instead stop cleanly and
 * return every unreached row as "Not processed", so the admin gets a finished
 * result and a clear reason rather than a dead connection.
 *
 * Dependency-free so it can be unit-tested under Node without Graph.
 */

import {
  NOT_PROCESSED_ROW_MESSAGE,
  type JobDeadline,
} from "@/lib/jobs/jobOutcome";
import type { BulkCommitRowInput, BulkPreviewRow } from "@/types/bulkUpload";
import type { BulkLogger } from "@/lib/services/bulkUpload/bulkUploadLog";

/** Rows the commit never reached, as Skipped results with a clear message. */
export function buildNotProcessedRows(
  rows: BulkCommitRowInput[],
): BulkPreviewRow[] {
  return rows.map((row) => ({
    rowNumber: row.rowNumber,
    status: "Skipped",
    messages: [NOT_PROCESSED_ROW_MESSAGE],
    fields: { ...row.fields },
    source: row.source,
    matchedEntityId: null,
    matchedEntityName: null,
    duplicateMatch: null,
    notProcessed: true,
  }));
}

/**
 * Call at the top of each row iteration. When the deadline has passed, returns
 * the Not-processed results for row `index` onwards (and logs it); otherwise
 * null, meaning carry on.
 */
export function stopIfPastDeadline(
  deadline: JobDeadline | undefined,
  rows: BulkCommitRowInput[],
  index: number,
  log?: BulkLogger,
): BulkPreviewRow[] | null {
  if (!deadline?.expired()) return null;
  const remaining = rows.slice(index);
  log?.warn("stopped-early", {
    reason: "time-budget",
    processed: index,
    notProcessed: remaining.length,
  });
  return buildNotProcessedRows(remaining);
}

export function countNotProcessed(rows: BulkPreviewRow[]): number {
  return rows.filter((row) => row.notProcessed).length;
}
