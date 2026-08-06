/**
 * Terminal / server logging for the bulk-upload commit flow.
 *
 * Dependency-free (no "server-only", no Graph) so it can be unit-tested with a
 * fake clock + capturing sink. Every line is prefixed with the import scope and
 * the elapsed time since the logger was created, so a whole commit reads as a
 * single timeline:
 *
 *   [bulk:commit:workforce] +0ms start rows=50 duplicateMode=skip
 *   [bulk:commit:workforce] +812ms phase load done ms=812 companies=37 workforce=1204
 *   [bulk:commit:workforce] +1903ms phase phase3b:creates done ms=too... created=48 errors=2
 *   [bulk:commit:workforce] +5210ms done imported=48 skipped=0 errors=2 ms=5210
 *
 * Levels:
 *   info  — always emitted (phase timing, milestones, final summary)
 *   warn  — always emitted (per-row failures; the things that "break the logic")
 *   error — always emitted
 *   debug — only when BULK_UPLOAD_LOGS=verbose (per-row detail; ~1 line per row)
 *
 * Disable entirely with BULK_UPLOAD_LOGS=off.
 */

export type BulkLogLevel = "info" | "warn" | "error" | "debug";

export interface BulkLogSink {
  info: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string) => void;
}

/**
 * Structured mirror of each log line, for consumers that need the data (not the
 * formatted string) — e.g. streaming live progress to the admin UI. Emitted
 * regardless of whether console output is enabled, and never includes the
 * verbose-only per-row `debug` lines.
 */
export type BulkLogEvent =
  | {
      type: "log";
      level: Exclude<BulkLogLevel, "debug">;
      scope: string;
      elapsedMs: number;
      event: string;
      details: Record<string, unknown>;
    }
  | { type: "phase-start"; scope: string; elapsedMs: number; label: string }
  | {
      type: "phase-end";
      scope: string;
      elapsedMs: number;
      label: string;
      ms: number;
      details: Record<string, unknown>;
    };

export interface BulkLoggerOptions {
  /** Where lines go. Defaults to console. */
  sink?: BulkLogSink;
  /** Monotonic clock in ms. Defaults to performance.now() / Date.now(). */
  now?: () => number;
  /** Force enable/disable. Defaults to env (on unless the env var = off). */
  enabled?: boolean;
  /** Emit debug (per-row) lines. Defaults to the env var = verbose. */
  verbose?: boolean;
  /**
   * Tag prefix shown as `[<prefix>:<scope>]`. Defaults to "bulk". Set e.g.
   * "docs" to reuse this logger for the customer-document upload flow.
   */
  prefix?: string;
  /**
   * Env var that toggles output (`off` silences, `verbose` adds debug lines).
   * Defaults to "BULK_UPLOAD_LOGS".
   */
  envVar?: string;
  /**
   * Structured event sink (fires even when console output is disabled). Used to
   * stream live progress to the browser. A throwing consumer never breaks the
   * import.
   */
  onEvent?: (event: BulkLogEvent) => void;
}

export interface BulkPhaseTimer {
  /** Logs "phase <label> done ms=<dur> …details" and returns the duration. */
  end: (details?: Record<string, unknown>) => number;
}

export interface BulkLogger {
  readonly scope: string;
  /** Whole-milliseconds since the logger was created. */
  elapsed: () => number;
  info: (event: string, details?: Record<string, unknown>) => void;
  warn: (event: string, details?: Record<string, unknown>) => void;
  error: (event: string, details?: Record<string, unknown>) => void;
  /** Per-row detail; only emitted in verbose mode. */
  debug: (event: string, details?: Record<string, unknown>) => void;
  /** Start a phase; call .end() to log its duration. */
  phase: (label: string) => BulkPhaseTimer;
}

const consoleSink: BulkLogSink = {
  info: (message) => console.log(message),
  warn: (message) => console.warn(message),
  error: (message) => console.error(message),
};

function defaultNow(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") {
    return performance.now();
  }
  return Date.now();
}

function readEnv(name: string): string | undefined {
  if (typeof process === "undefined" || !process.env) return undefined;
  const value = process.env[name];
  return typeof value === "string" ? value.trim().toLowerCase() : undefined;
}

function formatValue(value: unknown): string {
  if (value === null) return "null";
  if (value === undefined) return "undefined";
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : value.toFixed(1);
  }
  if (typeof value === "boolean") return String(value);
  if (typeof value === "string") {
    // Quote so spaces / special chars stay on one readable token.
    return /[\s"=]/.test(value) ? JSON.stringify(value) : value;
  }
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function formatDetails(details?: Record<string, unknown>): string {
  if (!details) return "";
  const parts: string[] = [];
  for (const [key, value] of Object.entries(details)) {
    if (value === undefined) continue;
    parts.push(`${key}=${formatValue(value)}`);
  }
  return parts.length ? ` ${parts.join(" ")}` : "";
}

/**
 * Create a scoped logger. `scope` is the tag shown in each line, e.g.
 * "commit:workforce".
 */
export function createBulkLogger(
  scope: string,
  options: BulkLoggerOptions = {},
): BulkLogger {
  const sink = options.sink ?? consoleSink;
  const now = options.now ?? defaultNow;
  const prefix = options.prefix ?? "bulk";
  const envFlag = readEnv(options.envVar ?? "BULK_UPLOAD_LOGS");
  const enabled = options.enabled ?? envFlag !== "off";
  const verbose = options.verbose ?? envFlag === "verbose";
  const onEvent = options.onEvent;
  const start = now();

  const elapsed = () => Math.round(now() - start);

  const toConsole = (
    level: Exclude<BulkLogLevel, "debug">,
    event: string,
    details?: Record<string, unknown>,
  ) => {
    if (!enabled) return;
    sink[level](`[${prefix}:${scope}] +${elapsed()}ms ${event}${formatDetails(details)}`);
  };

  const emit = (event: BulkLogEvent) => {
    if (!onEvent) return;
    try {
      onEvent(event);
    } catch {
      // A failing consumer (e.g. a browser that closed the stream) must never
      // break the import.
    }
  };

  const write = (
    level: Exclude<BulkLogLevel, "debug">,
    event: string,
    details?: Record<string, unknown>,
  ) => {
    toConsole(level, event, details);
    emit({
      type: "log",
      level,
      scope,
      elapsedMs: elapsed(),
      event,
      details: details ?? {},
    });
  };

  return {
    scope,
    elapsed,
    info: (event, details) => write("info", event, details),
    warn: (event, details) => write("warn", event, details),
    error: (event, details) => write("error", event, details),
    debug: (event, details) => {
      // Per-row detail: console only in verbose mode, and never streamed.
      if (verbose) toConsole("info", event, details);
    },
    phase: (label) => {
      const phaseStart = now();
      toConsole("info", `phase ${label} start`);
      emit({ type: "phase-start", scope, elapsedMs: elapsed(), label });
      return {
        end: (details) => {
          const ms = Math.round(now() - phaseStart);
          toConsole("info", `phase ${label} done`, { ms, ...(details ?? {}) });
          emit({
            type: "phase-end",
            scope,
            elapsedMs: elapsed(),
            label,
            ms,
            details: details ?? {},
          });
          return ms;
        },
      };
    },
  };
}
