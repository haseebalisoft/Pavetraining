/**
 * Client-side owner of the one bulk-upload commit that can run at a time.
 *
 * The commit used to live in AdminBulkUploadClient's component state, so
 * leaving the page dropped the stream reader and the admin never learned how
 * the import ended — the reason for "Please keep this window open". Here the
 * job lives in a module that survives in-app navigation: the page subscribes
 * on mount (useSyncExternalStore) and shows live progress or the final result,
 * however many times the admin leaves and comes back.
 *
 * The server keeps writing even if the tab closes (commit-progress registers
 * the work with `after()`), so leaving never abandons a half-written row. What
 * a closed tab loses is the report — hence the beforeunload prompt while a job
 * runs, and only then.
 *
 * Duplicate protection on retry:
 *  - one job at a time (a second Confirm while one runs is refused);
 *  - a job that ended without a confirmed result sets `needsRepreview`: some
 *    rows may be saved, so the page requires Preview again (saved rows then
 *    show as Duplicate / update) before another commit.
 */

import { readPublicApiError } from "@/lib/errors/publicMessages";
import {
  JobWatchdogError,
  bulkCommitOutcome,
  describeJobTransportFailure,
  type JobTerminalStatus,
} from "@/lib/jobs/jobOutcome";
import type {
  BulkCommitResult,
  BulkDuplicateMode,
  BulkImportType,
  BulkPreviewResult,
} from "@/types/bulkUpload";

/** One phase of a commit, with the display label and a rough weight (share of
 *  total time) so the gauge advances smoothly across phases of very different
 *  cost. `parallel` phases show the "running in parallel" hint. */
interface CommitPhaseModel {
  key: string;
  label: string;
  weight: number;
  parallel?: boolean;
}

const WORKFORCE_PHASE_MODEL: CommitPhaseModel[] = [
  { key: "load", label: "Loading data", weight: 2 },
  { key: "phase1:validate+companies", label: "Validating rows", weight: 3 },
  {
    key: "phase2:permissions+departments",
    label: "Setting up people & departments",
    weight: 5,
  },
  { key: "phase3a:updates", label: "Updating existing records", weight: 5 },
  {
    key: "phase3b:creates",
    label: "Creating workforce records",
    weight: 20,
    parallel: true,
  },
  {
    key: "phase3c:trainingMatrix",
    label: "Building Training Matrix",
    weight: 30,
    parallel: true,
  },
  {
    key: "phase3d:documentFolders",
    label: "Creating document folders",
    weight: 35,
    parallel: true,
  },
  { key: "phase3e:photos", label: "Storing candidate photos", weight: 5 },
];

const COMPANY_PHASE_MODEL: CommitPhaseModel[] = [
  { key: "load:companies", label: "Loading companies", weight: 5 },
  { key: "commit:rows", label: "Importing companies", weight: 95 },
];

/** Matrix commits: one list load, then serial per-row writes (the long part). */
const TRAINING_MATRIX_PHASE_MODEL: CommitPhaseModel[] = [
  { key: "load", label: "Loading Workforce and Training Matrix", weight: 1 },
  { key: "rows", label: "Writing Training Matrix rows", weight: 3 },
  {
    key: "categories",
    label: "Writing category backup records",
    weight: 6,
  },
];

function phaseModelFor(importType: BulkImportType): CommitPhaseModel[] {
  if (importType === "workforce") return WORKFORCE_PHASE_MODEL;
  if (importType === "company") return COMPANY_PHASE_MODEL;
  if (importType === "trainingMatrix") return TRAINING_MATRIX_PHASE_MODEL;
  // Other import types only report start/done — no weighted phases, so the
  // gauge runs indeterminate until the result arrives.
  return [];
}

/**
 * The server sends a heartbeat every 10s and always ends with a result or
 * error line before its 300s limit. Silence for 45s means the connection is
 * dead; the hard cap is the server limit plus slack. Either way the job is
 * marked Failed with a reason instead of spinning.
 */
const COMMIT_IDLE_TIMEOUT_MS = 45_000;
const COMMIT_HARD_TIMEOUT_MS = 330_000;

async function readWithWatchdog(
  reader: ReadableStreamDefaultReader<Uint8Array>,
  hardDeadlineAt: number,
): Promise<ReadableStreamReadResult<Uint8Array>> {
  const untilHard = hardDeadlineAt - Date.now();
  const hard = untilHard <= COMMIT_IDLE_TIMEOUT_MS;
  const waitMs = Math.max(0, Math.min(COMMIT_IDLE_TIMEOUT_MS, untilHard));
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      reader.read(),
      new Promise<never>((_, reject) => {
        timer = setTimeout(
          () =>
            reject(
              new JobWatchdogError(
                hard
                  ? `The import did not finish within ${Math.round(COMMIT_HARD_TIMEOUT_MS / 60_000 * 10) / 10} minutes, so it was stopped.`
                  : `No response from the server for ${COMMIT_IDLE_TIMEOUT_MS / 1000} seconds, so the import was stopped.`,
              ),
            ),
          waitMs,
        );
      }),
    ]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

export interface CommitProgressState {
  totalRows: number;
  phaseLabel: string;
  phaseKey: string | null;
  done: number;
  total: number;
  completed: string[];
  parallel: boolean;
  /** null = indeterminate (no weighted phase model / not started). */
  overallPct: number | null;
}

function computeOverallPct(
  model: CommitPhaseModel[],
  completed: Set<string>,
  phaseKey: string | null,
  done: number,
  total: number,
): number | null {
  if (model.length === 0) return null;
  const totalWeight = model.reduce((sum, phase) => sum + phase.weight, 0);
  let acc = 0;
  for (const phase of model) {
    if (completed.has(phase.key)) {
      acc += phase.weight;
    } else if (phase.key === phaseKey && total > 0) {
      acc += phase.weight * Math.min(1, done / total);
    }
  }
  // Cap below 100 until the final "done" event flips it to complete.
  return Math.min(99, Math.round((acc / totalWeight) * 100));
}

/** "Alder & Finch Ltd.PNG" → "alderfinchltd" — loose key for logo matching. */
function logoMatchKey(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/\.(png|jpe?g|gif|webp|bmp|svg)$/, "")
    .replace(/[^a-z0-9]+/g, "");
}

/** Same placeholder rule as the server (bulkUpload/importImage.ts). */
const IMAGE_PLACEHOLDER =
  /^((logo|photo|picture|image)(\s+placeholder)?|placeholder|n\/?a|none|null|-|—|–)$/i;

export interface AttachedImageUploadReport {
  /** Result rows with each image note replaced by its real outcome. */
  rows: BulkCommitResult["rows"];
  uploaded: number;
  failed: number;
  /** File names the sheet named but the admin did not attach. */
  notAttached: string[];
  /** Attached files no imported row used. */
  unused: string[];
}

/** @deprecated name kept for existing callers/tests. */
export type CompanyLogoUploadReport = AttachedImageUploadReport;

type ImageUpload = (entityId: string, file: File) => Promise<void>;

interface AttachedImageKind {
  /** Row cell holding the image ("companyLogo" / "candidatePhoto"). */
  cellField: string;
  /** Row property the server sets when it waits for an attached file. */
  pending: (row: BulkCommitResult["rows"][number]) => string | null | undefined;
  /**
   * When the cell is blank or a placeholder, a file named after this value
   * matches instead (company name / Workforce Number).
   */
  fallback: (row: BulkCommitResult["rows"][number]) => string | null | undefined;
  /** "Company Logo" / "Candidate Photo". */
  label: string;
  /** Prefix of the server's "nothing to upload" note to drop on a fallback match. */
  missingNotePrefix: string | null;
  /** Where the admin can upload one manually. */
  manualHint: string;
}

/**
 * Upload attached image files to the rows an import SAVED (never a skipped
 * duplicate or an error), through the same endpoints and storage as the
 * Admin upload buttons. One file may serve several rows. Each row's note is
 * replaced with what actually happened, so the report never says "uploaded"
 * for an image that was not.
 */
async function uploadAttachedImages(
  rows: BulkCommitResult["rows"],
  files: File[],
  kind: AttachedImageKind,
  upload: ImageUpload,
): Promise<AttachedImageUploadReport> {
  const byKey = new Map(files.map((file) => [logoMatchKey(file.name), file]));
  const used = new Set<File>();
  const notAttached: string[] = [];
  let uploaded = 0;
  let failed = 0;

  const out: BulkCommitResult["rows"] = [];
  for (const row of rows) {
    const entityId = row.savedEntityId;
    if (!entityId || row.status !== "Imported") {
      out.push(row);
      continue;
    }
    const cell = row.fields[kind.cellField]?.trim() ?? "";
    const named = kind.pending(row)?.trim() || null;
    const fallbackKey = kind.fallback(row)?.trim();
    const fallbackAllowed = !cell || IMAGE_PLACEHOLDER.test(cell);
    const file = named
      ? byKey.get(logoMatchKey(named))
      : fallbackAllowed && fallbackKey
        ? byKey.get(logoMatchKey(fallbackKey))
        : undefined;

    const withoutNotes = row.messages.filter(
      (message) =>
        !(named && message.startsWith(`${kind.label} "${named}" — waiting`)) &&
        !(
          file &&
          !named &&
          kind.missingNotePrefix &&
          message.startsWith(kind.missingNotePrefix)
        ),
    );

    if (!file) {
      if (named) {
        notAttached.push(named);
        out.push({
          ...row,
          messages: [
            ...withoutNotes,
            `${kind.label} "${named}" NOT uploaded — that file was not attached. Attach it and re-import, or ${kind.manualHint}.`,
          ],
        });
      } else {
        out.push(row);
      }
      continue;
    }

    used.add(file);
    try {
      await upload(entityId, file);
      uploaded += 1;
      out.push({
        ...row,
        messages: [...withoutNotes, `${kind.label} uploaded from "${file.name}".`],
      });
    } catch (error) {
      failed += 1;
      out.push({
        ...row,
        messages: [
          ...withoutNotes,
          `${kind.label} "${file.name}" upload FAILED: ${error instanceof Error ? error.message : "upload failed"}`,
        ],
      });
    }
  }

  const unused = files.filter((file) => !used.has(file)).map((file) => file.name);
  return { rows: out, uploaded, failed, notAttached, unused };
}

/**
 * Company import: a logo file matches by the file name in Company Logo, or —
 * when that cell is blank / a placeholder — by the company name
 * ("Alder & Finch Supplies Ltd.png").
 */
export function uploadAttachedCompanyLogos(
  rows: BulkCommitResult["rows"],
  files: File[],
  upload: ImageUpload = (id, file) =>
    postImage(`/api/admin/companies/${id}/logo`, file),
): Promise<AttachedImageUploadReport> {
  return uploadAttachedImages(
    rows,
    files,
    {
      cellField: "companyLogo",
      pending: (row) => row.pendingLogoFile,
      fallback: (row) => row.fields.companyName,
      label: "Company Logo",
      missingNotePrefix: "No logo",
      manualHint: "use Upload logo on the company page",
    },
    upload,
  );
}

/**
 * Workforce import: a photo file matches by the file name in Candidate Photo,
 * or — when that cell is blank / a placeholder — by the candidate's Workforce
 * Number ("PAVE-W0001.jpg"). Never by candidate name: two people can share
 * a name, and a photo on the wrong profile is worse than none.
 */
export function uploadAttachedCandidatePhotos(
  rows: BulkCommitResult["rows"],
  files: File[],
  upload: ImageUpload = (id, file) =>
    postImage(`/api/admin/workforce/${id}/photo`, file),
): Promise<AttachedImageUploadReport> {
  return uploadAttachedImages(
    rows,
    files,
    {
      cellField: "candidatePhoto",
      pending: (row) => row.pendingPhotoFile,
      fallback: (row) => row.fields.workforceNumber,
      label: "Candidate Photo",
      missingNotePrefix: null,
      manualHint: "use Upload photo on the candidate's profile",
    },
    upload,
  );
}

async function postImage(url: string, file: File): Promise<void> {
  const form = new FormData();
  form.set("file", file);
  const response = await fetch(url, {
    method: "POST",
    body: form,
    signal: AbortSignal.timeout(60_000),
  });
  if (!response.ok) throw new Error(await readPublicApiError(response));
}

/** Fold attached-image results into the import outcome (banner + toast agree). */
export function withImageOutcome(
  outcome: JobTerminalStatus,
  report: AttachedImageUploadReport,
  noun: "logo(s)" | "photo(s)",
  owner: "company" | "candidate",
): JobTerminalStatus {
  const notes: string[] = [];
  if (report.uploaded) notes.push(`${report.uploaded} ${noun} uploaded.`);
  const missing = report.failed + report.notAttached.length;
  if (missing) {
    notes.push(
      `${missing} ${noun} NOT uploaded${report.notAttached.length ? ` (not attached: ${report.notAttached.slice(0, 5).join(", ")}${report.notAttached.length > 5 ? "…" : ""})` : ""} — see the row messages.`,
    );
  }
  if (report.unused.length) {
    notes.push(
      `Attached file(s) not used by any imported ${owner}: ${report.unused.slice(0, 5).join(", ")}${report.unused.length > 5 ? "…" : ""}.`,
    );
  }
  if (!notes.length) return outcome;
  return {
    outcome:
      outcome.outcome === "completed" && (missing || report.unused.length)
        ? "completed_with_warnings"
        : outcome.outcome,
    reason: `${outcome.reason} ${notes.join(" ")}`,
  };
}

/** Company logos: see withImageOutcome. */
export function withLogoOutcome(
  outcome: JobTerminalStatus,
  logos: AttachedImageUploadReport,
): JobTerminalStatus {
  return withImageOutcome(outcome, logos, "logo(s)", "company");
}

export type BulkJobPhase = "running" | "finishing" | "done";

export interface BulkCommitJob {
  /** The preview this job imports — the page shows it again on return. */
  preview: BulkPreviewResult;
  /** Kept so Preview can be re-run after navigating away and back. */
  file: File | null;
  duplicateMode: BulkDuplicateMode;
  startedAt: number;
  /**
   * running   — server is writing rows (closing the tab loses the report);
   * finishing — result received, post-import steps (company logos) running;
   * done      — terminal: `outcome` is set.
   */
  phase: BulkJobPhase;
  progress: CommitProgressState | null;
  result: BulkCommitResult | null;
  outcome: JobTerminalStatus | null;
  /**
   * The commit ended without a confirmed result, so some rows may already be
   * saved. Confirm stays disabled until Preview is run again.
   */
  needsRepreview: boolean;
}

type Notify = (message: string, tone: "success" | "error") => void;

let current: BulkCommitJob | null = null;
const listeners = new Set<() => void>();

function emit(next: BulkCommitJob | null): void {
  current = next;
  for (const listener of listeners) listener();
}

function patch(job: BulkCommitJob, changes: Partial<BulkCommitJob>): void {
  // Ignore late updates from a job the admin already cleared.
  if (current?.startedAt !== job.startedAt) return;
  emit({ ...(current ?? job), ...changes });
}

export function subscribeBulkCommitJob(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function getBulkCommitJob(): BulkCommitJob | null {
  return current;
}

export function getServerBulkCommitJob(): BulkCommitJob | null {
  return null;
}

/** True while closing the tab would lose the import report. */
export function bulkCommitJobActive(): boolean {
  return current != null && current.phase !== "done";
}

/** Forget a finished job (new file, new preview, Clear). Never drops a live one. */
export function clearBulkCommitJob(): void {
  if (bulkCommitJobActive()) return;
  emit(null);
}

export interface StartBulkCommitInput {
  preview: BulkPreviewResult;
  file: File | null;
  duplicateMode: BulkDuplicateMode;
  suppressNotifications: boolean;
  logoFiles: File[];
  /** Workforce imports: candidate photo files to upload after the import. */
  photoFiles: File[];
  notify: Notify;
}

/**
 * Start the commit. Returns false (and notifies) when another import is
 * already running — two concurrent commits of one file would race each
 * other's duplicate checks and create the same records twice.
 */
export function startBulkCommit(input: StartBulkCommitInput): boolean {
  if (bulkCommitJobActive()) {
    input.notify(
      "An import is already running. Wait for it to finish before starting another.",
      "error",
    );
    return false;
  }
  const { preview } = input;
  const startedAt = Date.now();
  const job: BulkCommitJob = {
    preview,
    file: input.file,
    duplicateMode: input.duplicateMode,
    startedAt,
    phase: "running",
    progress: null,
    result: null,
    outcome: null,
    needsRepreview: false,
  };
  emit(job);
  void runBulkCommit(job, input);
  return true;
}

async function runBulkCommit(
  job: BulkCommitJob,
  input: StartBulkCommitInput,
): Promise<void> {
  const { preview, notify } = input;
  const hardDeadlineAt = job.startedAt + COMMIT_HARD_TIMEOUT_MS;

  // Live progress state accumulated from the streamed events.
  const model = phaseModelFor(preview.importType);
  const state = {
    totalRows: preview.rows.length,
    phaseKey: null as string | null,
    phaseLabel: "Starting…",
    done: 0,
    total: 0,
    completed: new Set<string>(),
    parallel: false,
  };
  const publish = () => {
    patch(job, {
      progress: {
        totalRows: state.totalRows,
        phaseKey: state.phaseKey,
        phaseLabel: state.phaseLabel,
        done: state.done,
        total: state.total,
        completed: [...state.completed],
        parallel: state.parallel,
        overallPct: computeOverallPct(
          model,
          state.completed,
          state.phaseKey,
          state.done,
          state.total,
        ),
      },
    });
  };
  publish();

  const applyEvent = (event: {
    type?: string;
    label?: string;
    event?: string;
    details?: Record<string, unknown>;
  }) => {
    if (event.type === "phase-start" && event.label) {
      state.phaseKey = event.label;
      const known = model.find((phase) => phase.key === event.label);
      state.phaseLabel = known?.label ?? event.label;
      state.parallel = Boolean(known?.parallel);
      state.done = 0;
      state.total = 0;
      publish();
    } else if (event.type === "phase-end" && event.label) {
      state.completed.add(event.label);
      if (state.phaseKey === event.label && state.total > 0) {
        state.done = state.total;
      }
      publish();
    } else if (event.type === "log") {
      const details = event.details ?? {};
      if (event.event === "start" && typeof details.rows === "number") {
        state.totalRows = details.rows;
        publish();
      } else if (event.event === "progress") {
        if (typeof details.done === "number") state.done = details.done;
        if (typeof details.total === "number") state.total = details.total;
        publish();
      }
    }
  };

  /** Error text the server itself reported — shown as-is, not reworded. */
  let serverFailure: string | null = null;
  /**
   * True only when the server answered with a refusal (auth, validation,
   * already running) — then nothing was written. Any other failure, including
   * a timeout before the response, may have reached the server and saved rows.
   */
  let refused = false;
  const openController = new AbortController();
  const openTimer = setTimeout(
    () => openController.abort(),
    COMMIT_IDLE_TIMEOUT_MS * 2,
  );
  try {
    const response = await fetch("/api/admin/bulk-upload/commit-progress", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        importType: preview.importType,
        fileName: preview.fileName,
        duplicateMode: job.duplicateMode,
        suppressNotifications: input.suppressNotifications,
        autoCreateMissingCompanies: preview.importType === "workforce",
        autoCreateMissingDepartments: preview.importType === "workforce",
        rows: preview.rows.map((row) => ({
          rowNumber: row.rowNumber,
          fields: row.fields,
          source: row.source,
        })),
      }),
      // Covers only the wait for the stream to open (the signal also
      // governs the body, so the timer is cleared on response); once open,
      // the read watchdog below takes over.
      signal: openController.signal,
    });
    clearTimeout(openTimer);
    if (!response.ok || !response.body) {
      refused = true;
      throw new Error(await readPublicApiError(response));
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let data: BulkCommitResult | null = null;
    let streamError: string | null = null;
    let terminal = false;

    const consumeLine = (line: string) => {
      const trimmed = line.trim();
      if (!trimmed) return;
      let message: {
        kind?: string;
        result?: BulkCommitResult;
        message?: string;
        event?: Record<string, unknown>;
      };
      try {
        message = JSON.parse(trimmed);
      } catch {
        return; // ignore a partial / malformed line
      }
      if (message.kind === "result" && message.result) {
        data = message.result;
        terminal = true;
      } else if (message.kind === "error") {
        streamError = message.message || "Import failed on the server.";
        terminal = true;
      } else if (message.kind === "progress" && message.event) {
        applyEvent(message.event);
      }
      // "heartbeat" lines only reset the watchdog.
    };

    // Read the NDJSON stream: update the clock as events arrive. Stop at
    // the first terminal line ("result" / "error"); a stream that goes
    // silent or overruns is cancelled by the watchdog.
    try {
      while (!terminal) {
        const { value, done } = await readWithWatchdog(reader, hardDeadlineAt);
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        let newline = buffer.indexOf("\n");
        while (newline >= 0 && !terminal) {
          consumeLine(buffer.slice(0, newline));
          buffer = buffer.slice(newline + 1);
          newline = buffer.indexOf("\n");
        }
      }
      if (buffer && !terminal) consumeLine(buffer);
    } finally {
      reader.cancel().catch(() => {});
    }

    if (streamError) {
      serverFailure = streamError;
      throw new Error(streamError);
    }
    if (!data) {
      // Stream closed with no terminal line: the server process was cut off.
      throw new JobWatchdogError(
        "The server stopped before reporting a result (the connection closed mid-import).",
      );
    }

    const result: BulkCommitResult = data;
    const outcome = bulkCommitOutcome(result);
    // The import itself is over: the spinner stops now, even if logos follow.
    const hasImageStep =
      result.importType === "company" || result.importType === "workforce";
    patch(job, {
      phase: hasImageStep ? "finishing" : "done",
      result,
      outcome,
      progress: null,
    });
    let finalOutcome = outcome;
    if (result.importType === "company") {
      const logos = await uploadAttachedCompanyLogos(result.rows, input.logoFiles);
      finalOutcome = withLogoOutcome(outcome, logos);
      patch(job, {
        phase: "done",
        result: { ...result, rows: logos.rows },
        outcome: finalOutcome,
      });
    } else if (result.importType === "workforce") {
      // Photos are optional: rows without one are untouched, and a missing
      // or failed photo only adds a note — the candidates are already saved.
      const photos = await uploadAttachedCandidatePhotos(
        result.rows,
        input.photoFiles ?? [],
      );
      finalOutcome = withImageOutcome(outcome, photos, "photo(s)", "candidate");
      patch(job, {
        phase: "done",
        result: { ...result, rows: photos.rows },
        outcome: finalOutcome,
      });
    }
    // Wording (incl. the client-approved Matrix phrasing) lives in
    // bulkCommitOutcome so the toast and the status banner always agree.
    notify(
      finalOutcome.reason,
      finalOutcome.outcome === "completed" ? "success" : "error",
    );
  } catch (error) {
    const reason = serverFailure
      ? `${serverFailure} Some rows may already be saved — run Preview again before retrying so saved rows show as duplicates.`
      : refused
        ? error instanceof Error && error.message
          ? error.message
          : "The server refused the import. Nothing was saved."
        : describeJobTransportFailure(error, "import");
    patch(job, {
      phase: "done",
      progress: null,
      outcome: { outcome: "failed", reason },
      // A refused request wrote nothing, so the same preview is safe to retry.
      needsRepreview: !refused,
    });
    notify(reason, "error");
  } finally {
    clearTimeout(openTimer);
  }
}
