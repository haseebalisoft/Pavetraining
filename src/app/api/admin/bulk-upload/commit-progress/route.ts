import { ValidationError } from "@/lib/api/adminApi";
import { logBulkUpload } from "@/lib/services/auditLogService";
import { commitBulkUpload } from "@/lib/services/bulkUpload/bulkUploadService";
import { requireAdminAccess } from "@/lib/services/securityService";
import { AccessDeniedError, UnauthorizedError } from "@/lib/services/errorHandler";
import type { BulkCommitRowInput } from "@/types/bulkUpload";

export const dynamic = "force-dynamic";
/** Large Workforce imports can exceed the default serverless window. */
export const maxDuration = 300;

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
 *   { "kind": "progress", "event": <BulkLogEvent> }
 *   { "kind": "result",   "result": <BulkCommitResult> }
 *   { "kind": "error",    "message": string }
 * The connection stays open for the whole import (the model that already works
 * here), and the final result is the last "result" line.
 */
export async function POST(request: Request): Promise<Response> {
  // Auth first (same guard as withAdminApi) — errors return normal JSON so the
  // client can surface them before any streaming starts.
  let loggedInEmail: string;
  try {
    const context = await requireAdminAccess();
    loggedInEmail = context.loggedInEmail;
  } catch (error) {
    if (error instanceof UnauthorizedError) {
      return jsonError("You must be signed in.", 401);
    }
    if (error instanceof AccessDeniedError) {
      return jsonError("Admin access is required.", 403);
    }
    return jsonError("Admin access is required.", 403);
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
      ? false
      : Boolean(body.autoCreateMissingCompanies);

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
      return { rowNumber, fields };
    });
  } catch (error) {
    return jsonError(
      error instanceof Error ? error.message : "Invalid rows.",
      400,
    );
  }

  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    async start(controller) {
      let closed = false;
      const send = (obj: unknown) => {
        if (closed) return;
        try {
          controller.enqueue(encoder.encode(`${JSON.stringify(obj)}\n`));
        } catch {
          closed = true;
        }
      };

      try {
        const result = await commitBulkUpload({
          importType,
          fileName,
          duplicateMode,
          suppressNotifications,
          autoCreateMissingCompanies,
          rows,
          onEvent: (event) => send({ kind: "progress", event }),
        });

        await logBulkUpload({
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
            streamed: true,
          },
        }).catch(() => {});

        send({ kind: "result", result });
      } catch (error) {
        await logBulkUpload({
          userEmail: loggedInEmail,
          phase: "commit",
          success: false,
          request,
          errorMessage: error instanceof Error ? error.message : "Commit failed",
          metadata: {
            importType: importType || null,
            fileName,
            duplicateMode,
            rowCount: rows.length,
            streamed: true,
          },
        }).catch(() => {});

        send({
          kind: "error",
          message: error instanceof Error ? error.message : "Commit failed",
        });
      } finally {
        closed = true;
        try {
          controller.close();
        } catch {
          // already closed
        }
      }
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
