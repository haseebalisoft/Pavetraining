/**
 * Terminal outcomes for long-running admin jobs (bulk upload commits and
 * Training Matrix syncs).
 *
 * Every job the admin starts must end in exactly one of three states, with a
 * human-readable reason — never an open-ended spinner:
 *
 *   completed                — everything succeeded
 *   completed_with_warnings  — finished, but some rows need attention or were
 *                              not processed in time
 *   failed                   — nothing was saved, or the job could not finish
 *
 * Dependency-free (types only) so client components and Node test scripts can
 * share it.
 */

import type { BulkCommitResult } from "@/types/bulkUpload";
import type { MatrixSyncResult } from "@/types/matrixSync";

export type JobOutcome = "completed" | "completed_with_warnings" | "failed";

export interface JobTerminalStatus {
  outcome: JobOutcome;
  reason: string;
}

export const JOB_OUTCOME_LABELS: Record<JobOutcome, string> = {
  completed: "Completed",
  completed_with_warnings: "Completed with warnings",
  failed: "Failed",
};

/** Wall-clock budget for a job, checked between units of work. */
export interface JobDeadline {
  readonly at: number;
  expired: () => boolean;
  remainingMs: () => number;
}

export function createJobDeadline(
  budgetMs: number,
  now: () => number = Date.now,
): JobDeadline {
  const at = now() + Math.max(0, budgetMs);
  return {
    at,
    expired: () => now() >= at,
    remainingMs: () => Math.max(0, at - now()),
  };
}

/** Message attached to rows a commit stopped before reaching. */
export const NOT_PROCESSED_ROW_MESSAGE =
  "Not processed — the import reached the server time limit before this row. Re-upload the file to finish: rows already saved are updated, not duplicated.";

function plural(count: number, one: string, many = `${one}s`): string {
  return `${count} ${count === 1 ? one : many}`;
}

/**
 * Training Matrix commit counts from the rows themselves, each row in exactly
 * one bucket: linked (Imported), Needs Review (saved unlinked), empty
 * (skipped), not saved because several Matrix rows matched, failed (Error),
 * not processed (time limit — reported by the stopped-early note).
 */
export function matrixRowCounts(result: Pick<BulkCommitResult, "rows">): {
  linked: number;
  needsReview: number;
  empty: number;
  ambiguous: number;
  failed: number;
  notProcessed: number;
} {
  const counts = {
    linked: 0,
    needsReview: 0,
    empty: 0,
    ambiguous: 0,
    failed: 0,
    notProcessed: 0,
  };
  for (const row of result.rows) {
    if (row.notProcessed) counts.notProcessed += 1;
    else if (row.status === "Imported") counts.linked += 1;
    else if (row.status === "Error") counts.failed += 1;
    else if (row.linkOutcome === "skippedAmbiguous") counts.ambiguous += 1;
    else if (row.status === "Warning") counts.needsReview += 1;
    else if (row.status === "Skipped") counts.empty += 1;
  }
  return counts;
}

/**
 * Classifies a finished bulk commit. The reason text is what the admin sees in
 * the toast and in the status banner.
 */
export function bulkCommitOutcome(result: BulkCommitResult): JobTerminalStatus {
  const { summary } = result;
  const warningN = summary.warningRows;
  const stopped = result.stoppedEarly ?? null;
  const isMatrix = result.importType === "trainingMatrix";
  const matrix = isMatrix ? matrixRowCounts(result) : null;
  // Matrix Needs Review rows ARE saved (unlinked) — reporting them as a
  // failure misleads the client. Rows skipped because several Matrix rows
  // matched are Warnings too, but nothing was written for them.
  const savedRows = matrix
    ? matrix.linked + matrix.needsReview
    : summary.importedRows;
  const backupPending = result.rows.filter((row) => row.backupPending).length;
  const stoppedNote =
    (stopped
      ? ` Stopped early: ${stopped.reason} ${plural(stopped.notProcessedRows, "row")} not processed — re-upload the file to finish.`
      : "") +
    (backupPending
      ? ` Category backup records still to write for ${plural(backupPending, "row")} (their Matrix dates are saved) — re-upload the file to finish.`
      : "");

  if (savedRows === 0 && (summary.errorRows > 0 || stopped)) {
    const firstError = result.rows
      .find((row) => row.status === "Error" && row.messages.length)
      ?.messages.slice(-1)[0];
    const base =
      summary.errorRows > 0
        ? firstError ??
          `Import failed for ${plural(summary.errorRows, "row")}. Check Messages in the table.`
        : "No rows were saved.";
    return { outcome: "failed", reason: `${base}${stoppedNote}`.trim() };
  }

  let reason: string;
  if (isMatrix) {
    // Client-approved phrasing: "N Matrix rows uploaded. X linked, Y need
    // review, Z empty rows skipped, W failed."
    const counts = matrix!;
    const parts = [
      `${savedRows} Matrix row${savedRows === 1 ? "" : "s"} uploaded`,
      `${counts.linked} linked`,
    ];
    if (counts.needsReview > 0) {
      const reasons = reviewReasonSummary(
        result.rows.filter((row) => row.linkOutcome !== "skippedAmbiguous"),
      );
      parts.push(
        `${counts.needsReview} need review${reasons ? ` (${reasons})` : ""}`,
      );
    }
    if (counts.empty > 0) {
      parts.push(plural(counts.empty, "empty row") + " skipped");
    }
    if (counts.ambiguous > 0) {
      parts.push(
        `${plural(counts.ambiguous, "row")} not saved (several Matrix rows match — resolve in Admin Matrix)`,
      );
    }
    if (counts.failed > 0) parts.push(`${counts.failed} failed`);
    reason = `${parts.join(", ")}.`;
  } else if (summary.errorRows > 0) {
    reason =
      `Import finished with errors: ${summary.importedRows} imported, ${summary.skippedRows} skipped, ${summary.errorRows} failed` +
      (warningN ? `, ${warningN} warning(s)` : "") +
      ". Review row Messages.";
  } else {
    reason =
      `Import finished: ${summary.importedRows} imported, ${summary.skippedRows} skipped` +
      (warningN
        ? `, ${warningN} warning(s) — see row messages for exact reasons`
        : "") +
      ".";
  }

  const hasWarnings =
    summary.errorRows > 0 ||
    warningN > 0 ||
    Boolean(stopped) ||
    backupPending > 0;
  return {
    outcome: hasWarnings ? "completed_with_warnings" : "completed",
    reason: `${reason}${stoppedNote}`,
  };
}

/** "2 DOB mismatch, 1 Candidate not found" across rows left Needs Review. */
export function reviewReasonSummary(
  rows: Array<{ reviewReason?: { label: string } | null }>,
): string {
  const counts = new Map<string, number>();
  for (const row of rows) {
    const label = row.reviewReason?.label;
    if (label) counts.set(label, (counts.get(label) ?? 0) + 1);
  }
  return [...counts].map(([label, count]) => `${count} ${label}`).join(", ");
}

/** Classifies a finished Training Matrix sync ("Refresh from registers"). */
export function matrixSyncOutcome(
  result: MatrixSyncResult,
  label = "Refresh",
): JobTerminalStatus {
  const { summary } = result;
  const incomplete = result.incomplete ?? null;
  const incompleteNote = incomplete
    ? ` Stopped early: ${incomplete.reason} ${plural(incomplete.remainingCandidates, "candidate")} not processed — run it again to finish.`
    : "";
  const changed = summary.updated + summary.created;
  const verb = result.dryRun ? "would update" : "updated";

  if (summary.errors > 0 && changed === 0 && summary.errors >= summary.processed) {
    const firstError = result.items.find((item) => item.errors.length)
      ?.errors[0];
    return {
      outcome: "failed",
      reason:
        `${label} failed for ${plural(summary.errors, "candidate")}` +
        (firstError ? `: ${firstError}` : ".") +
        incompleteNote,
    };
  }

  if (incomplete && summary.processed === 0) {
    return {
      outcome: "failed",
      reason: `${label} did not reach any candidates.${incompleteNote}`,
    };
  }

  const parts = [
    `${label}: ${verb} ${summary.updated}, created ${summary.created} (${plural(summary.processed, "candidate")} checked)`,
  ];
  if (summary.errors > 0) parts.push(`${summary.errors} with errors`);
  if (summary.warnings > 0) parts.push(plural(summary.warnings, "warning"));
  const reason = `${parts.join(", ")}.${incompleteNote}`;

  return {
    outcome:
      summary.errors > 0 || summary.warnings > 0 || incomplete
        ? "completed_with_warnings"
        : "completed",
    reason,
  };
}

/**
 * Turns a client-side transport failure (dropped connection, abort, gateway
 * timeout) into a reason that says what happened and what to do next.
 */
export function describeJobTransportFailure(
  error: unknown,
  what: "import" | "refresh" | "preview",
): string {
  const partial =
    what === "preview"
      ? "Nothing was saved. Try again."
      : what === "import"
        ? "Some rows may already be saved — run Preview again before retrying so nothing is imported twice."
        : "Some changes may already be saved — reload the page to check before running it again.";
  if (error instanceof JobWatchdogError) {
    return `${error.message} ${partial}`;
  }
  const name = error instanceof Error ? error.name : "";
  const message = error instanceof Error ? error.message : String(error ?? "");
  if (name === "AbortError" || name === "TimeoutError") {
    return `The ${what} took too long and was stopped. ${partial}`;
  }
  if (
    error instanceof TypeError ||
    /network|failed to fetch|load failed|terminated|connection/i.test(message)
  ) {
    return `Lost connection to the server during the ${what}. ${partial}`;
  }
  return message || `The ${what} failed. ${partial}`;
}

/** Raised by client watchdogs when a job stops reporting progress. */
export class JobWatchdogError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "JobWatchdogError";
  }
}
