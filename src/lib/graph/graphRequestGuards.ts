/**
 * Graph request guards: per-request timeout middleware + error classification.
 *
 * Kept dependency-free (SDK *types* only, erased at build) so the Node test
 * runner can load it without Next or the SDK's bundler-style directory
 * imports — see scripts/test-job-lifecycle.mjs. graphClient.ts wires it in.
 */

import type { Context, Middleware } from "@microsoft/microsoft-graph-client";

export function graphErrorStatus(error: unknown): number | null {
  if (!error || typeof error !== "object") return null;
  const value = (error as { statusCode?: unknown; status?: unknown }).statusCode ??
    (error as { status?: unknown }).status;
  return typeof value === "number" ? value : null;
}

function errorText(error: unknown): string {
  if (error instanceof Error) {
    const cause = "cause" in error ? String(error.cause ?? "") : "";
    return `${error.message} ${cause}`.toLowerCase();
  }
  return String(error ?? "").toLowerCase();
}

export function isGraphTimeoutError(error: unknown): boolean {
  if (!error || typeof error !== "object") return false;
  const { name, code } = error as { name?: unknown; code?: unknown };
  return name === "GraphTimeoutError" || code === "GraphTimeoutError";
}

/** True for transient network / throttling failures Graph SDK surfaces poorly. */
export function isTransientGraphError(error: unknown): boolean {
  // The SDK re-wraps our timeout as GraphError(statusCode -1, code
  // "GraphTimeoutError"). Never replay it: a timed-out write may have landed.
  if (isGraphTimeoutError(error)) return false;
  const status = graphErrorStatus(error);
  if (status === -1 || status === 408 || status === 425 || status === 429) {
    return true;
  }
  if (status !== null && status >= 500) return true;

  const text = errorText(error);
  return [
    "fetch failed",
    "terminated",
    "econnreset",
    "econnrefused",
    "etimedout",
    "socket hang up",
    "und_err_",
    "network",
    "other side closed",
  ].some((marker) => text.includes(marker));
}

/**
 * A Graph call that never answers. Deliberately has no status code and none of
 * the transient markers, so `withGraphMutationRetry` does NOT replay it — a
 * timed-out POST may still have landed, and replaying it could duplicate rows.
 */
export class GraphTimeoutError extends Error {
  constructor(timeoutMs: number, target: string) {
    const seconds = timeoutMs >= 1000 ? `${Math.round(timeoutMs / 1000)}s` : `${timeoutMs}ms`;
    super(
      `Microsoft Graph did not respond within ${seconds} (${target}). SharePoint may be slow — try again shortly.`,
    );
    this.name = "GraphTimeoutError";
  }
}

function describeGraphTarget(context: Context): string {
  const method = context.options?.method ?? "GET";
  const url =
    typeof context.request === "string" ? context.request : context.request.url;
  try {
    const { pathname } = new URL(url);
    return `${method} ${pathname.replace(/^\/(v1\.0|beta)/, "")}`;
  } catch {
    return method;
  }
}

/**
 * Without this, a stalled SharePoint request holds its caller open forever:
 * fetch has no default timeout, and only the platform's maxDuration (absent
 * under `next dev` / `next start`) would ever end it. Callers that pass their
 * own AbortSignal keep it. The timer covers the attempt up to response
 * headers; insert it just before HTTPMessageHandler so each SDK retry is
 * capped individually.
 */
export class GraphTimeoutHandler implements Middleware {
  private next: Middleware | undefined;

  constructor(private readonly timeoutMs: number) {}

  setNext(next: Middleware): void {
    this.next = next;
  }

  async execute(context: Context): Promise<void> {
    if (!this.next) return;
    if (context.options?.signal) {
      await this.next.execute(context);
      return;
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    context.options = { ...(context.options ?? {}), signal: controller.signal };
    try {
      await this.next.execute(context);
    } catch (error) {
      if (controller.signal.aborted) {
        throw new GraphTimeoutError(this.timeoutMs, describeGraphTarget(context));
      }
      throw error;
    } finally {
      clearTimeout(timer);
    }
  }
}
