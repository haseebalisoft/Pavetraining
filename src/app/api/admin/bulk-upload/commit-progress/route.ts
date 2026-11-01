import { createHash } from "node:crypto";

import { after } from "next/server";

import { ValidationError } from "@/lib/api/adminApi";
import { createJobDeadline } from "@/lib/jobs/jobOutcome";
import { logBulkUpload } from "@/lib/services/auditLogService";
import { commitBulkUpload } from "@/lib/services/bulkUpload/bulkUploadService";
import { normalizeSourceRow } from "@/lib/services/bulkUpload/parseSpreadsheet";
import { requireSharePointAdminAccess } from "@/lib/services/securityService";
import { AccessDeniedError, UnauthorizedError } from "@/lib/services/errorHandler";
import type { BulkCommitRowInput } from "@/types/bulkUpload";

export const dynamic = "force-dynamic";
/** Large Workforce imports can exceed the default serverless window. */
export const maxDuration = 300;

/**
 * Stop starting new rows 40s before maxDuration: leaves room for the row in
 * flight (each Graph attempt is capped by the client timeout) and for the
 * final result line, so the platform never kills the stream mid-import.
 */
const BULK_COMMIT_BUDGET_MS = (300 - 40) * 1000;
const BULK_COMMIT_HEARTBEAT_MS = 10_000;

type BulkAuditInput = Parameters<typeof logBulkUpload>[0];

/**
 * Imports running on this server instance, keyed by a fingerprint of the
 * import type + rows. A second Confirm of the same file while the first is
 * still writing (double-click, second tab) is refused instead of racing it
 * and creating the same records twice. Per-instance only — the browser job
 * store enforces one import per tab, and every importer re-checks live
 * SharePoint data for duplicates at commit time.
 */
const runningImports = new Map<string, number>();

function importFingerprint(importType: string, rows: unknown): string {
  return createHash("sha256")
    .update(importType)
    .update(JSON.stringify(rows))
    .digest("hex");
}

/** Audit once the response is done; failures are logged, never surfaced. */
function auditAfterResponse(input: BulkAuditInput): void {
  const write = () =>
    logBulkUpload(input).catch((error) => {
      console.error("[bulk-upload] commit audit write failed", error);
    });
  try {
    after(write);
  } catch {
    // Outside a request scope (e.g. scripts) — fire and forget.
    void write();
  }
}

function jsonError(message: string, status: number): Response {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

/**
 * Streaming sibling of /commit: same admin-only commit, but the response is a
 * newline-delimited JSON stream so the browser can show live progress instead
 * of staring at a spinner for minutes. Each line is one of:
 *   { "kind": "progress",  "event": <BulkLogEvent> }
 *   { "kind": "heartbeat", "elapsedMs": number }   (every 10s)
 *   { "kind": "result",    "result": <BulkCommitResult> }
 *   { "kind": "error",     "message": string }
 * The connection stays open for the whole import. Exactly one terminal line
 * ("result" or "error") is always sent: serial imports stop at the time budget
 * and return unreached rows as Not processed rather than being killed.
 */
export async function POST(request: Request): Promise<Response> {
  const startedAt = Date.now();
  // Budget runs from request start so auth + body parsing count against it.
  const deadline = createJobDeadline(BULK_COMMIT_BUDGET_MS);
  // Auth first (same guard as withAdminApi) — errors return normal JSON so the
  // client can surface them before any streaming starts.
  let loggedInEmail: string;
  try {
    const context = await requireSharePointAdminAccess();
    loggedInEmail = context.loggedInEmail;
  } catch (error) {
    if (error instanceof UnauthorizedError) {
      return jsonError("You must be signed in.", 401);
    }
    if (error instanceof AccessDeniedError) {
      return jsonError(
        error.message || "Full Admin access is required.",
        403,
      );
    }
    return jsonError("Full Admin access is required.", 403);
  }

  let body: Record<string, unknown>;
  try {
    body = (await request.json()) as Record<string, unknown>;
  } catch {
    return jsonError("Expected JSON body.", 400);
  }

  const importType = String(body.importType ?? "").trim();
  const fileName =
    body.fileName === null || body.fileName === undefined
      ? null
      : String(body.fileName);
  const duplicateMode =
    body.duplicateMode === null || body.duplicateMode === undefined
      ? "skip"
      : String(body.duplicateMode);
  const suppressNotifications =
    body.suppressNotifications === undefined
      ? true
      : Boolean(body.suppressNotifications);
  const autoCreateMissingCompanies =
    body.autoCreateMissingCompanies === undefined
      ? true
      : Boolean(body.autoCreateMissingCompanies);
  const autoCreateMissingDepartments =
    body.autoCreateMissingDepartments === undefined
      ? true
      : Boolean(body.autoCreateMissingDepartments);

  if (!Array.isArray(body.rows)) {
    return jsonError("rows must be an array.", 400);
  }

  let rows: BulkCommitRowInput[];
  try {
    rows = body.rows.map((raw, index) => {
      if (!raw || typeof raw !== "object") {
        throw new ValidationError(`Invalid row at index ${index}.`);
      }
      const row = raw as Record<string, unknown>;
      const rowNumber = Number(row.rowNumber);
      if (!Number.isFinite(rowNumber) || rowNumber < 1) {
        throw new ValidationError(`Invalid rowNumber at index ${index}.`);
      }
      const fieldsRaw = row.fields;
      if (!fieldsRaw || typeof fieldsRaw !== "object") {
        throw new ValidationError(`Missing fields at row ${rowNumber}.`);
      }
      const fields: Record<string, string | null> = {};
      for (const [key, value] of Object.entries(
        fieldsRaw as Record<string, unknown>,
      )) {
        if (value === null || value === undefined || value === "") {
          fields[key] = null;
        } else {
          fields[key] = String(value);
        }
      }
      return { rowNumber, fields, source: normalizeSourceRow(row.source) };
    });
  } catch (error) {
    return jsonError(
      error instanceof Error ? error.message : "Invalid rows.",
      400,
    );
  }

  const fingerprint = importFingerprint(importType, body.rows);
  if (runningImports.has(fingerprint)) {
    return jsonError(
      "This file is already being imported. Wait for that import to finish, then run Preview again to see what was saved.",
      409,
    );
  }
  runningImports.set(fingerprint, startedAt);

  // The import must finish even if the admin closes the tab: the stream's
  // `cancel()` only stops sending, and `after()` keeps the function alive
  // until the commit settles (bounded by maxDuration and the time budget).
  let settleJob!: () => void;
  const jobSettled = new Promise<void>((resolve) => {
    settleJob = resolve;
  });
  try {
    after(() => jobSettled);
  } catch {
    // Outside a request scope (scripts) — nothing to keep alive.
  }

  const encoder = new TextEncoder();
  let closed = false;
  const stream = new ReadableStream<Uint8Array>({
    async start(controller) {
      const send = (obj: unknown) => {
        if (closed) return;
        try {
          controller.enqueue(encoder.encode(`${JSON.stringify(obj)}\n`));
        } catch {
          closed = true;
        }
      };
      // Keeps the connection visibly alive through long phases (one slow
      // Graph call can take tens of seconds) so the browser watchdog can tell
      // "busy" from "dead", and idle proxies don't drop the stream.
      const heartbeat = setInterval(
        () => send({ kind: "heartbeat", elapsedMs: Date.now() - startedAt }),
        BULK_COMMIT_HEARTBEAT_MS,
      );

      try {
        const result = await commitBulkUpload({
          importType,
          fileName,
          duplicateMode,
          suppressNotifications,
          autoCreateMissingCompanies,
          autoCreateMissingDepartments,
          rows,
          onEvent: (event) => send({ kind: "progress", event }),
          deadline,
          actorEmail: loggedInEmail,
        });

        // Deliver the result first; the audit write is a SharePoint call and
        // must never hold the admin's job open.
        send({ kind: "result", result });
        auditAfterResponse({
          userEmail: loggedInEmail,
          phase: "commit",
          success: true,
          itemCount: result.summary.importedRows,
          request,
          metadata: {
            importType: result.importType,
            fileName: result.fileName,
            duplicateMode: result.duplicateMode,
            suppressNotifications: result.suppressNotifications,
            totalRows: result.summary.totalRows,
            importedRows: result.summary.importedRows,
            skippedRows: result.summary.skippedRows,
            duplicateRows: result.summary.duplicateRows,
            errorRows: result.summary.errorRows,
            warningRows: result.summary.warningRows,
            notProcessedRows: result.stoppedEarly?.notProcessedRows ?? 0,
            streamed: true,
          },
        });
      } catch (error) {
        const message =
          error instanceof Error && error.message
            ? error.message
            : "Commit failed for an unknown reason.";
        send({ kind: "error", message });
        auditAfterResponse({
          userEmail: loggedInEmail,
          phase: "commit",
          success: false,
          request,
          errorMessage: message,
          metadata: {
            importType: importType || null,
            fileName,
            duplicateMode,
            rowCount: rows.length,
            streamed: true,
          },
        });
      } finally {
        clearInterval(heartbeat);
        runningImports.delete(fingerprint);
        settleJob();
        closed = true;
        try {
          controller.close();
        } catch {
          // already closed
        }
      }
    },
    cancel() {
      // Browser went away; stop enqueueing. The commit itself runs on (kept
      // alive by `after`) so a half-written import is never abandoned.
      closed = true;
    },
  });

  return new Response(stream, {
    headers: {
      "Content-Type": "application/x-ndjson; charset=utf-8",
      // Discourage any intermediate proxy/CDN from buffering the stream so
      // progress lines arrive as they happen, not all at the end.
      "Cache-Control": "no-store, no-transform",
      "X-Accel-Buffering": "no",
    },
  });
}
