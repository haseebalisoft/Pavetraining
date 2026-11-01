"use client";

import {
  useCallback,
  useEffect,
  useMemo,
  useState,
  useSyncExternalStore,
  type DragEvent,
} from "react";

import {
  clearBulkCommitJob,
  getBulkCommitJob,
  getServerBulkCommitJob,
  startBulkCommit,
  subscribeBulkCommitJob,
  type CommitProgressState,
} from "@/components/admin/bulkUpload/bulkCommitJob";
import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import {
  bulkUploadRowLimit,
  checkBulkUploadRowLimit,
} from "@/lib/services/bulkUpload/bulkUploadLimits";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import {
  JOB_OUTCOME_LABELS,
  describeJobTransportFailure,
  reviewReasonSummary,
  type JobOutcome,
  type JobTerminalStatus,
} from "@/lib/jobs/jobOutcome";
import type { StatusTone } from "@/lib/ui/status";
import {
  BULK_FOLDER_OUTCOME_LABELS,
  BULK_LINK_OUTCOME_LABELS,
  type BulkDuplicateMode,
  type BulkImportSummary,
  type BulkImportType,
  type BulkPreviewResult,
  type BulkPreviewRow,
  type BulkRowStatus,
} from "@/types/bulkUpload";

import styles from "@/components/admin/admin.module.css";

const IMPORT_OPTIONS: Array<{
  value: BulkImportType;
  label: string;
  hint: string;
  implemented: boolean;
}> = [
  {
    value: "company",
    label: "Companies",
    hint: "Use Company list.xlsx exactly. Company Logo = https image URL (optional — missing logo warns only). Preview shows every Excel column. Creates or updates Company List rows.",
    implemented: true,
  },
  {
    value: "workforce",
    label: "Workforce / Candidates",
    hint: "Use Workforce list.xlsx exactly. Missing companies, departments, training managers, and supervisors are created automatically. No row limit.",
    implemented: true,
  },
  {
    value: "trainingMatrix",
    label: "Training Matrix rows",
    hint: "Use Training matrix example.xlsx exactly. Preview shows every Excel column. Import Workforce first (match by Name / DOB).",
    implemented: true,
  },
  {
    value: "npors",
    label: "NPORS records",
    hint: "Import NPORS rows. Candidate must exist in Workforce. Pass outcomes sync to the Training Matrix.",
    implemented: true,
  },
  {
    value: "eusr",
    label: "EUSR records",
    hint: "Import EUSR rows. Candidate must exist in Workforce. Pass outcomes sync to the Training Matrix.",
    implemented: true,
  },
  {
    value: "streetworks",
    label: "Streetworks / NRSWA",
    hint: "Import Streetworks/NRSWA rows. Candidate must exist in Workforce. Pass outcomes sync to the Training Matrix.",
    implemented: true,
  },
  {
    value: "inHouse",
    label: "In-House records",
    hint: "Course column: Asbestos Awareness (or N031) / Face Fit / etc. Asbestos Awareness Pass + Expiry syncs to Training Matrix N031.",
    implemented: true,
  },
  {
    value: "nvq",
    label: "NVQ records",
    hint: "Import NVQ progress rows. Standalone — does not update the Training Matrix.",
    implemented: true,
  },
];

function statusTone(status: BulkRowStatus): StatusTone {
  switch (status) {
    case "Ready":
    case "Imported":
      return "ok";
    case "Warning":
      return "warn";
    case "Duplicate":
      return "info";
    case "Error":
      return "danger";
    case "Skipped":
      return "missing";
    default:
      return "neutral";
  }
}

function downloadTextFile(fileName: string, content: string, mime: string) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function SummaryCards({ summary }: { summary: BulkImportSummary }) {
  const items = [
    { label: "Total", value: summary.totalRows },
    { label: "Ready", value: summary.readyRows },
    { label: "Warnings", value: summary.warningRows },
    { label: "Duplicates", value: summary.duplicateRows },
    { label: "Errors", value: summary.errorRows },
    { label: "Skipped", value: summary.skippedRows },
    { label: "Imported", value: summary.importedRows },
    ...(summary.notProcessedRows
      ? [{ label: "Not processed", value: summary.notProcessedRows }]
      : []),
  ];
  return (
    <div className={styles.bulkSummaryGrid}>
      {items.map((item) => (
        <div key={item.label} className={styles.bulkSummaryCard}>
          <span className={styles.bulkSummaryLabel}>{item.label}</span>
          <strong className={styles.bulkSummaryValue}>{item.value}</strong>
        </div>
      ))}
    </div>
  );
}

function PreviewTable({
  rows,
  headers,
}: {
  rows: BulkPreviewRow[];
  headers: string[];
}) {
  if (!rows.length) {
    return <p className={styles.bulkEmpty}>No rows to show.</p>;
  }

  const dataHeaders = headers.length
    ? headers
    : [
        "Candidate Name",
        "Company Name",
        "Workforce Number",
        "Date of birth",
      ];

  const cellValue = (row: BulkPreviewRow, header: string): string => {
    const key = header.trim().toLowerCase();
    // Prefer normalized Company Size so preview shows the SharePoint choice value
    // (Small | Medium | Large | Enterprise), not a raw/casing variant from Excel.
    if (key === "company size" && row.fields.companySize?.trim()) {
      return row.fields.companySize.trim();
    }
    const fromSource = row.source?.[header];
    if (fromSource != null && String(fromSource).trim() !== "") {
      return String(fromSource);
    }
    // Fallbacks for mapped internal keys when source cell is blank/missing.
    const mapped: Record<string, string | null | undefined> = {
      "candidate name": row.fields.candidateName,
      name: row.fields.candidateName,
      "company name": row.fields.companyName ?? row.fields.company,
      company: row.fields.company ?? row.fields.companyName,
      "company size": row.fields.companySize,
      "workforce number": row.fields.workforceNumber,
      "date of birth": row.fields.dateOfBirth,
      dob: row.fields.dateOfBirth,
    };
    const hit = mapped[key] ?? row.fields[header] ?? null;
    return hit?.trim() ? hit : "—";
  };

  // Only Workforce/Matrix imports resolve a link, so the column stays hidden for
  // company/register imports instead of showing a column of dashes.
  const showLink = rows.some((row) => row.linkOutcome);
  // Only Workforce imports touch document folders.
  const showFolder = rows.some((row) => row.folderOutcome);

  return (
    <div className={`${styles.tableWrap} ${styles.bulkPreviewTableWrap}`}>
      <table className={`${styles.dataTable} ${styles.bulkPreviewTable}`}>
        <thead>
          <tr>
            <th>Row</th>
            <th>Status</th>
            {showLink ? <th>Link</th> : null}
            {showFolder ? <th>Folder</th> : null}
            {dataHeaders.map((header) => (
              <th key={header}>{header}</th>
            ))}
            <th>Messages</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={`${row.rowNumber}-${row.status}`}>
              <td>{row.rowNumber}</td>
              <td>
                <StatusBadge label={row.status} tone={statusTone(row.status)} />
              </td>
              {showLink ? (
                <td>
                  {row.linkOutcome
                    ? BULK_LINK_OUTCOME_LABELS[row.linkOutcome]
                    : "—"}
                  {row.reviewReason ? (
                    <>
                      <br />
                      <strong>{row.reviewReason.label}</strong>
                    </>
                  ) : null}
                </td>
              ) : null}
              {showFolder ? (
                <td>
                  {row.folderOutcome
                    ? BULK_FOLDER_OUTCOME_LABELS[row.folderOutcome]
                    : "—"}
                </td>
              ) : null}
              {dataHeaders.map((header) => (
                <td key={`${row.rowNumber}-${header}`}>
                  {cellValue(row, header)}
                </td>
              ))}
              <td className={styles.bulkMessageCell}>
                {row.messages.length ? row.messages.join(" ") : "—"}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/** Preview route maxDuration is 300s. */
const PREVIEW_TIMEOUT_MS = 310_000;

const OUTCOME_TONE: Record<JobOutcome, StatusTone> = {
  completed: "ok",
  completed_with_warnings: "warn",
  failed: "danger",
};

/** Persistent terminal state for the last commit — survives the toast. */
function CommitOutcomeBanner({ status }: { status: JobTerminalStatus }) {
  return (
    <div className={styles.bulkHint} role="status" aria-live="polite">
      <StatusBadge
        label={JOB_OUTCOME_LABELS[status.outcome]}
        tone={OUTCOME_TONE[status.outcome]}
      />{" "}
      {status.reason}
    </div>
  );
}

/** Circular "progress clock" shown live during a bulk commit. */
function BulkCommitClock({
  progress,
  elapsedMs,
}: {
  progress: CommitProgressState;
  elapsedMs: number;
}) {
  const radius = 52;
  const circumference = 2 * Math.PI * radius;
  const indeterminate = progress.overallPct == null;
  const shownPct = indeterminate ? 30 : progress.overallPct ?? 0;
  const offset = circumference * (1 - shownPct / 100);
  const totalSeconds = Math.floor(elapsedMs / 1000);
  const mm = String(Math.floor(totalSeconds / 60)).padStart(2, "0");
  const ss = String(totalSeconds % 60).padStart(2, "0");

  return (
    <div className={styles.commitClock} role="status" aria-live="polite">
      <div
        className={`${styles.commitClockGauge}${
          indeterminate ? ` ${styles.commitClockGaugeSpin}` : ""
        }`}
      >
        <svg viewBox="0 0 120 120" aria-hidden="true">
          <circle cx="60" cy="60" r={radius} className={styles.commitClockTrack} />
          <circle
            cx="60"
            cy="60"
            r={radius}
            className={styles.commitClockArc}
            strokeDasharray={circumference}
            strokeDashoffset={offset}
            transform="rotate(-90 60 60)"
          />
        </svg>
        <div className={styles.commitClockCenter}>
          <span className={styles.commitClockPct}>
            {indeterminate ? "…" : `${progress.overallPct}%`}
          </span>
        </div>
      </div>
      <div className={styles.commitClockMeta}>
        <p className={styles.commitClockTime}>
          <span aria-hidden="true">⏱ </span>
          {mm}:{ss} elapsed
        </p>
        <p className={styles.commitClockPhase}>{progress.phaseLabel}</p>
        {progress.total > 0 ? (
          <p className={styles.commitClockCount}>
            {progress.done} / {progress.total} uploaded
          </p>
        ) : progress.totalRows > 0 ? (
          <p className={styles.commitClockCount}>
            {progress.totalRows} record{progress.totalRows === 1 ? "" : "s"}
          </p>
        ) : null}
        {progress.parallel ? (
          <p className={styles.commitClockParallel}>Running 12 in parallel</p>
        ) : null}
        <p className={styles.commitClockHint}>
          You can open other admin pages — the import keeps running and its
          result will be here when you come back. Closing or reloading this tab
          won&apos;t stop it, but you would lose this report.
        </p>
      </div>
    </div>
  );
}

/** "Alder-Finch.PNG" → "alderfinch" — same loose match the uploader uses. */
function imageMatchKey(name: string): string {
  return name
    .trim()
    .toLowerCase()
    .replace(/\.(png|jpe?g|gif|webp|bmp|svg)$/, "")
    .replace(/[^a-z0-9]+/g, "");
}

/** Image file names a preview names in `cellField` that are not attached. */
function unattachedImageNames(
  rows: BulkPreviewRow[],
  cellField: string,
  attachedFiles: File[],
): string[] {
  const attached = new Set(attachedFiles.map((file) => imageMatchKey(file.name)));
  const named = new Set<string>();
  for (const row of rows) {
    const cell = row.fields[cellField]?.trim() ?? "";
    if (
      row.status !== "Error" &&
      /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(cell) &&
      !/^[a-z][a-z0-9+.-]*:\/\//i.test(cell) &&
      !row.messages.some((message) => message.includes("already this company's logo")) &&
      !attached.has(imageMatchKey(cell))
    ) {
      named.add(cell);
    }
  }
  return [...named];
}

export function AdminBulkUploadClient() {
  const { pushToast } = useAdminToast();
  // The commit job lives in a module store so it survives navigation; a page
  // mounted while a job exists re-attaches to it (same preview, same file).
  const job = useSyncExternalStore(
    subscribeBulkCommitJob,
    getBulkCommitJob,
    getServerBulkCommitJob,
  );
  const [importType, setImportType] = useState<BulkImportType>(
    () => getBulkCommitJob()?.preview.importType ?? "workforce",
  );
  const [file, setFile] = useState<File | null>(
    () => getBulkCommitJob()?.file ?? null,
  );
  /** Company imports: logo images matched to rows by file name. */
  const [logoFiles, setLogoFiles] = useState<File[]>([]);
  /** Workforce imports: candidate photos matched to rows by file name. */
  const [photoFiles, setPhotoFiles] = useState<File[]>([]);
  const [dragOver, setDragOver] = useState(false);
  const [suppressNotifications, setSuppressNotifications] = useState(true);
  const [duplicateMode, setDuplicateMode] = useState<BulkDuplicateMode>(
    () => getBulkCommitJob()?.duplicateMode ?? "skip",
  );
  const [previewing, setPreviewing] = useState(false);
  const [preview, setPreview] = useState<BulkPreviewResult | null>(
    () => getBulkCommitJob()?.preview ?? null,
  );
  const [statusFilter, setStatusFilter] = useState<"" | BulkRowStatus>("");

  // Job state applies only while the page shows the preview that job imports.
  const pageJob = job && job.preview === preview ? job : null;
  /** Any import still in flight (blocks a second commit and new previews). */
  const jobActive = job != null && job.phase !== "done";
  const committing = pageJob?.phase === "running";
  const commitProgress = committing ? (pageJob?.progress ?? null) : null;
  const commitResult = pageJob?.result ?? null;
  /**
   * Rows are saved; the browser is still uploading attached logos / photos.
   * Not finished yet — no final banner, and leaving would stop the uploads.
   */
  const finishingImages = pageJob?.phase === "finishing";
  /** Terminal state of the last commit: Completed / with warnings / Failed. */
  const commitOutcome = pageJob?.phase === "done" ? pageJob.outcome : null;
  const needsRepreview = pageJob?.needsRepreview ?? false;

  // Client-side clock: ticks every 500ms while committing so the elapsed time
  // keeps moving smoothly even between server progress events.
  const [nowTick, setNowTick] = useState(() => Date.now());
  useEffect(() => {
    if (!committing) return;
    const id = setInterval(() => setNowTick(Date.now()), 500);
    return () => clearInterval(id);
  }, [committing]);
  const commitElapsedMs =
    committing && pageJob ? Math.max(0, nowTick - pageJob.startedAt) : 0;

  // In-app navigation is always safe (the job outlives this page). Closing or
  // reloading the tab would lose the report, so warn — but only while a job
  // is actually running.
  useEffect(() => {
    if (!jobActive) return;
    const onBeforeUnload = (event: BeforeUnloadEvent) => {
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", onBeforeUnload);
    return () => window.removeEventListener("beforeunload", onBeforeUnload);
  }, [jobActive]);

  const selectedOption = IMPORT_OPTIONS.find((o) => o.value === importType);

  const displayRows = useMemo(() => {
    const source = commitResult?.rows ?? preview?.rows ?? [];
    if (!statusFilter) return source;
    return source.filter((row) => row.status === statusFilter);
  }, [commitResult, preview, statusFilter]);

  const activeSummary = commitResult?.summary ?? preview?.summary ?? null;
  const activeReviewReasons = reviewReasonSummary(
    commitResult?.rows ?? preview?.rows ?? [],
  );

  /**
   * Image file names in the sheet (Company Logo / Candidate Photo) with no
   * matching attached file. Checked here because only the browser knows what
   * was attached.
   */
  const missingLogoFiles = useMemo(
    () =>
      preview?.importType === "company"
        ? unattachedImageNames(preview.rows, "companyLogo", logoFiles)
        : [],
    [logoFiles, preview],
  );
  const missingPhotoFiles = useMemo(
    () =>
      preview?.importType === "workforce"
        ? unattachedImageNames(preview.rows, "candidatePhoto", photoFiles)
        : [],
    [photoFiles, preview],
  );

  /** Row cap for the currently selected import type (null = no limit). */
  const selectedRowLimit = bulkUploadRowLimit(importType);

  /** Client-side mirror of the server limit, evaluated against the preview. */
  const rowLimitError = useMemo(() => {
    if (!preview?.implemented || !preview.rows.length) return null;
    const check = checkBulkUploadRowLimit(
      preview.importType,
      preview.rows.length,
    );
    return check.ok ? null : check.error;
  }, [preview]);

  const canCommit =
    !!preview?.implemented &&
    !!preview.rows.length &&
    !commitResult &&
    !needsRepreview &&
    !jobActive &&
    !rowLimitError &&
    (preview.summary.readyRows > 0 ||
      preview.summary.warningRows > 0 ||
      preview.summary.duplicateRows > 0);

  /** Explains why Confirm import is greyed out (client-blocking clarity). */
  const commitBlockedReason = useMemo(() => {
    if (!preview?.implemented || commitResult) return null;
    if (needsRepreview) {
      return "The last import did not confirm its result, so some rows may already be saved. Run Preview validation again — saved rows will then show as Duplicates — before importing again.";
    }
    if (jobActive && !committing) {
      return "Another import is still running. Wait for it to finish.";
    }
    if (!preview.rows.length) return "No preview rows to import.";
    if (rowLimitError) return rowLimitError;
    if (
      preview.summary.readyRows === 0 &&
      preview.summary.warningRows === 0 &&
      preview.summary.duplicateRows === 0
    ) {
      if (preview.summary.errorRows > 0) {
        return "All rows are Errors (often: candidate not found in Workforce). Fix Errors or import Workforce first, then Preview again.";
      }
      return "No importable rows after validation.";
    }
    if (
      preview.summary.duplicateRows > 0 &&
      duplicateMode === "skip" &&
      preview.summary.readyRows === 0 &&
      preview.summary.warningRows === 0
    ) {
      return "All matched rows are Duplicates. Change “If duplicate found” to “Update existing” to refresh Training Matrix, or keep Skip (Confirm will only skip).";
    }
    return null;
  }, [
    commitResult,
    committing,
    duplicateMode,
    jobActive,
    needsRepreview,
    preview,
    rowLimitError,
  ]);

  const onFileChosen = (next: File | null) => {
    setFile(next);
    setPreview(null);
    clearBulkCommitJob();
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    const dropped = event.dataTransfer.files?.[0] ?? null;
    if (!dropped || jobActive) return;
    const lower = dropped.name.toLowerCase();
    if (!lower.endsWith(".xlsx") && !lower.endsWith(".csv")) {
      pushToast("Only .xlsx and .csv files are supported.", "error");
      return;
    }
    onFileChosen(dropped);
  };

  const runPreview = useCallback(async () => {
    if (!file) {
      pushToast("Choose a spreadsheet first.", "error");
      return;
    }
    setPreviewing(true);
    clearBulkCommitJob();
    try {
      const form = new FormData();
      form.set("file", file);
      form.set("importType", importType);
      form.set(
        "suppressNotifications",
        suppressNotifications ? "true" : "false",
      );
      if (importType === "workforce") {
        form.set("autoCreateMissingCompanies", "true");
        form.set("autoCreateMissingDepartments", "true");
      }
      const response = await fetch("/api/admin/bulk-upload/preview", {
        method: "POST",
        body: form,
        // Never leave "Validating…" up forever if the server stops answering.
        signal: AbortSignal.timeout(PREVIEW_TIMEOUT_MS),
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const data = (await response.json()) as BulkPreviewResult;
      setPreview(data);
      // Training Matrix re-uploads usually match existing rows as Duplicate.
      // Skip mode leaves Confirm disabled/useless — switch to Update so admin
      // can refresh expiry columns in one click.
      if (
        data.implemented &&
        data.importType === "trainingMatrix" &&
        data.summary.duplicateRows > 0 &&
        data.summary.readyRows === 0 &&
        data.summary.warningRows === 0
      ) {
        setDuplicateMode("update");
      }
      if (!data.implemented) {
        pushToast(data.message ?? "This import type is not ready yet.", "error");
      } else {
        pushToast(
          `Validated ${data.summary.totalRows} row(s). Review before importing.`,
          "success",
        );
      }
    } catch (error) {
      pushToast(describeJobTransportFailure(error, "preview"), "error");
    } finally {
      setPreviewing(false);
    }
  }, [
    file,
    importType,
    pushToast,
    suppressNotifications,
  ]);

  const runCommit = useCallback(() => {
    if (!preview?.implemented) return;
    // Mirror the server row limit before committing so oversized uploads are
    // stopped client-side with the same clear message.
    const limitCheck = checkBulkUploadRowLimit(
      preview.importType,
      preview.rows.length,
    );
    if (!limitCheck.ok && limitCheck.error) {
      pushToast(limitCheck.error, "error");
      return;
    }
    if (
      duplicateMode === "update" ||
      duplicateMode === "create"
    ) {
      const confirmed = window.confirm(
        duplicateMode === "update"
          ? "Update existing duplicate candidates with non-blank values from the spreadsheet?"
          : "Create new candidates even when duplicates were detected?",
      );
      if (!confirmed) return;
    }
    // The job lives outside this component so it survives navigation; the
    // page re-attaches to it through useSyncExternalStore.
    startBulkCommit({
      preview,
      file,
      duplicateMode,
      suppressNotifications,
      logoFiles,
      photoFiles,
      notify: pushToast,
    });
  }, [
    duplicateMode,
    file,
    logoFiles,
    photoFiles,
    preview,
    pushToast,
    suppressNotifications,
  ]);

  const downloadTemplate = useCallback(
    async (type: BulkImportType) => {
      try {
        const response = await fetch(
          `/api/admin/bulk-upload/templates?download=${encodeURIComponent(type)}`,
          { cache: "no-store" },
        );
        if (!response.ok) throw new Error(await readPublicApiError(response));

        const contentType = response.headers.get("content-type") ?? "";
        if (
          contentType.includes(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          ) ||
          contentType.includes("application/octet-stream")
        ) {
          const blob = await response.blob();
          const disposition = response.headers.get("content-disposition") ?? "";
          const match = disposition.match(/filename=\"([^\"]+)\"/);
          const fileName =
            match?.[1] ??
            (type === "workforce"
              ? "Workforce-list-template.xlsx"
              : type === "trainingMatrix"
                ? "Training-matrix-template.xlsx"
                : "template.xlsx");
          const url = URL.createObjectURL(blob);
          const anchor = document.createElement("a");
          anchor.href = url;
          anchor.download = fileName;
          anchor.click();
          URL.revokeObjectURL(url);
          return;
        }

        const data = (await response.json()) as {
          fileName: string;
          csv: string;
        };
        downloadTextFile(data.fileName, data.csv, "text/csv;charset=utf-8");
      } catch (error) {
        pushToast(
          error instanceof Error ? error.message : "Template download failed",
          "error",
        );
      }
    },
    [pushToast],
  );

  const downloadReport = useCallback(() => {
    const rows = commitResult?.rows ?? preview?.rows ?? [];
    if (!rows.length) return;
    const dataHeaders = preview?.headers?.length
      ? preview.headers
      : [
          "Candidate Name",
          "Company Name",
          "Workforce Number",
          "Date of birth",
          "Department",
        ];
    const header = [
      "Row",
      "Status",
      "Link",
      "Review Reason",
      "Folder",
      ...dataHeaders,
      "Matched Id",
      "Matrix Row Id",
      "Messages",
    ];
    const escape = (value: string) =>
      /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
    const lines = [header.map(escape).join(",")];
    for (const row of rows) {
      const cells = dataHeaders.map((col) => {
        const fromSource = row.source?.[col];
        if (fromSource != null && String(fromSource).trim() !== "") {
          return String(fromSource);
        }
        return row.fields[col] ?? "";
      });
      lines.push(
        [
          String(row.rowNumber),
          row.status,
          row.linkOutcome ? BULK_LINK_OUTCOME_LABELS[row.linkOutcome] : "",
          row.reviewReason
            ? `${row.reviewReason.label}: ${row.reviewReason.detail}`
            : "",
          row.folderOutcome ? BULK_FOLDER_OUTCOME_LABELS[row.folderOutcome] : "",
          ...cells,
          row.matchedEntityId ?? "",
          row.matrixRowId ?? "",
          row.messages.join("; "),
        ]
          .map(escape)
          .join(","),
      );
    }
    const stamp = new Date().toISOString().slice(0, 10);
    downloadTextFile(
      `pave-bulk-validation-${stamp}.csv`,
      `${lines.join("\r\n")}\r\n`,
      "text/csv;charset=utf-8",
    );
  }, [commitResult, preview]);

  const resetAll = () => {
    setFile(null);
    setPreview(null);
    clearBulkCommitJob();
    setStatusFilter("");
  };

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs
            items={[
              { label: "Admin", href: "/admin" },
              { label: "Bulk Upload" },
            ]}
          />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>Bulk Upload</h1>
          <p className={styles.subtitle}>
            Upload a spreadsheet, preview validation, then import safely.
            Notifications are suppressed by default. Duplicates are skipped
            unless you confirm update or create.
          </p>
        </div>
        <button
          type="button"
          className={styles.secondaryButton}
          disabled={jobActive}
          title={jobActive ? "Wait for the running import to finish." : undefined}
          onClick={resetAll}
        >
          Clear
        </button>
      </header>

      <div className={styles.bulkLayout}>
        <section className={styles.settingsCard}>
          <header className={styles.settingsCardHeader}>
            <h2>1. Choose import type</h2>
            <p>
              Import Companies / Workforce first, then registers or matrix.
              NPORS / EUSR / Streetworks sync Pass outcomes into the Training
              Matrix; In-House and NVQ stay standalone.
            </p>
          </header>
          <div className={styles.settingsCardBody}>
            <label className={styles.bulkField}>
              <span>Import type</span>
              <select
                className={styles.bulkSelect}
                value={importType}
                disabled={jobActive}
                onChange={(event) => {
                  setImportType(event.target.value as BulkImportType);
                  setPreview(null);
                  clearBulkCommitJob();
                }}
              >
                {IMPORT_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                    {option.implemented ? "" : " (coming next)"}
                  </option>
                ))}
              </select>
            </label>
            {selectedOption ? (
              <p className={styles.bulkHint}>{selectedOption.hint}</p>
            ) : null}
            {selectedRowLimit !== null ? (
              <p className={styles.bulkHint}>
                Limit: up to {selectedRowLimit} records per upload.
              </p>
            ) : null}

            <div className={styles.bulkTemplateRow}>
              {IMPORT_OPTIONS.map((option) => (
                <button
                  key={option.value}
                  type="button"
                  className={styles.secondaryButton}
                  onClick={() => void downloadTemplate(option.value)}
                >
                  {option.label} template
                </button>
              ))}
            </div>
          </div>
        </section>

        <section className={styles.settingsCard}>
          <header className={styles.settingsCardHeader}>
            <h2>2. Upload spreadsheet</h2>
            <p>
              Accepted formats: .xlsx and .csv. Max 15 MB. Download the exact
              client Workforce / Training Matrix Excel templates above — column
              headers must match for a clean import.
            </p>
          </header>
          <div className={styles.settingsCardBody}>
            <div
              className={`${styles.bulkDropzone} ${dragOver ? styles.bulkDropzoneActive : ""}`}
              onDragOver={(event) => {
                event.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
            >
              <p className={styles.bulkDropTitle}>
                Drag and drop your file here
              </p>
              <p className={styles.bulkHint}>or choose a file from your computer</p>
              <label className={styles.primaryButton}>
                Choose file
                <input
                  type="file"
                  accept=".xlsx,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  className={styles.bulkFileInput}
                  disabled={jobActive}
                  onChange={(event) =>
                    onFileChosen(event.target.files?.[0] ?? null)
                  }
                />
              </label>
              {file ? (
                <p className={styles.bulkFileName}>{file.name}</p>
              ) : (
                <p className={styles.bulkHint}>No file selected yet.</p>
              )}
            </div>

            {importType === "company" ? (
              <div className={styles.bulkLogoPicker}>
                <label className={styles.secondaryButton}>
                  Company logo files (optional)
                  <input
                    type="file"
                    accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml,image/bmp"
                    multiple
                    className={styles.bulkFileInput}
                    onChange={(event) =>
                      setLogoFiles(Array.from(event.target.files ?? []))
                    }
                  />
                </label>
                <p className={styles.bulkHint}>
                  {logoFiles.length > 0
                    ? `${logoFiles.length} logo file(s) selected: ${logoFiles
                        .map((logo) => logo.name)
                        .join(", ")}`
                    : "Two ways to add logos: put an image web link (https://…) in the Company Logo column, or type the logo's file name there (e.g. alder-finch.png) and pick those image files here. A file named after the company (e.g. Alder & Finch Supplies Ltd.png) also matches when the cell is blank. Logos upload right after the import."}
                </p>
                {missingLogoFiles.length > 0 ? (
                  <p className={styles.formError} role="status">
                    The spreadsheet names {missingLogoFiles.length} logo file
                    {missingLogoFiles.length === 1 ? "" : "s"} you have not
                    attached: {missingLogoFiles.slice(0, 6).join(", ")}
                    {missingLogoFiles.length > 6 ? "…" : ""}. Those companies
                    will import without a logo unless you attach them.
                  </p>
                ) : null}
              </div>
            ) : null}

            {importType === "workforce" ? (
              <div className={styles.bulkLogoPicker}>
                <label className={styles.secondaryButton}>
                  Candidate photo files (optional)
                  <input
                    type="file"
                    accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml,image/bmp"
                    multiple
                    className={styles.bulkFileInput}
                    onChange={(event) =>
                      setPhotoFiles(Array.from(event.target.files ?? []))
                    }
                  />
                </label>
                <p className={styles.bulkHint}>
                  {photoFiles.length > 0
                    ? `${photoFiles.length} photo file(s) selected: ${photoFiles
                        .map((photo) => photo.name)
                        .join(", ")}`
                    : "Photos are optional. Put the photo's file name in the Candidate Photo column (e.g. zaid.jpg) and pick those files here, or name each file after the candidate's Workforce Number (e.g. PAVE-W0001.jpg) and leave the cell blank. An https image link in the cell also works. Photos upload right after the import; a missing photo never stops a candidate importing."}
                </p>
                {missingPhotoFiles.length > 0 ? (
                  <p className={styles.formError} role="status">
                    The spreadsheet names {missingPhotoFiles.length} photo file
                    {missingPhotoFiles.length === 1 ? "" : "s"} you have not
                    attached: {missingPhotoFiles.slice(0, 6).join(", ")}
                    {missingPhotoFiles.length > 6 ? "…" : ""}. Those
                    candidates will import without a photo unless you attach
                    them.
                  </p>
                ) : null}
              </div>
            ) : null}

            <label className={styles.bulkCheckRow}>
              <input
                type="checkbox"
                checked={suppressNotifications}
                onChange={(event) =>
                  setSuppressNotifications(event.target.checked)
                }
              />
              <span>
                Suppress notifications during import (recommended). Untick to
                email each candidate&apos;s Training Manager about the NPORS /
                EUSR / Streetworks / In-House / NVQ results being imported.
              </span>
            </label>

            <div className={styles.bulkActions}>
              <button
                type="button"
                className={styles.primaryButton}
                disabled={!file || previewing || jobActive}
                onClick={() => void runPreview()}
              >
                {previewing ? "Validating…" : "Preview validation"}
              </button>
            </div>
          </div>
        </section>

        {preview || commitResult ? (
          <section className={styles.settingsCard}>
            <header className={styles.settingsCardHeader}>
              <h2>
                {commitResult
                  ? "4. Import summary"
                  : committing
                    ? "4. Importing…"
                    : "3. Preview results"}
              </h2>
              <p>
                {commitResult
                  ? commitResult.message ??
                    "Import finished. Download the validation report for your records."
                  : committing
                    ? "Nothing below has changed yet: the counts and statuses are from the preview and switch to the import result when it finishes."
                    : "Review Ready, Warning, Duplicate, and Error rows before confirming import. Invalid rows are never imported."}
              </p>
            </header>
            <div className={styles.settingsCardBody}>
              {committing && commitProgress ? (
                <BulkCommitClock
                  progress={commitProgress}
                  elapsedMs={commitElapsedMs}
                />
              ) : null}

              {finishingImages ? (
                <p className={styles.bulkHint} role="status" aria-live="polite">
                  <strong>Rows saved — uploading the attached{" "}
                  {pageJob?.preview.importType === "company" ? "logos" : "photos"}…</strong>{" "}
                  Stay on this tab until this finishes (other admin pages are
                  fine); closing it would stop the remaining uploads. The
                  imported records themselves are already saved.
                </p>
              ) : null}

              {!committing && commitOutcome ? (
                <CommitOutcomeBanner status={commitOutcome} />
              ) : null}

              {activeSummary ? <SummaryCards summary={activeSummary} /> : null}

              {activeReviewReasons ? (
                <p className={styles.bulkHint} role="status">
                  <strong>Needs Review reasons:</strong> {activeReviewReasons}.
                  See the Link and Messages columns for each row.
                </p>
              ) : null}

              {rowLimitError ? (
                <p className={styles.formError} role="alert">
                  {rowLimitError}
                </p>
              ) : null}

              {!commitResult ? (
                <>
                  <label className={styles.bulkField}>
                    <span>Duplicate handling</span>
                    <select
                      className={styles.bulkSelect}
                      value={duplicateMode}
                      onChange={(event) =>
                        setDuplicateMode(
                          event.target.value as BulkDuplicateMode,
                        )
                      }
                    >
                      <option value="skip">
                        Skip duplicates (default / safest)
                      </option>
                      <option value="update">
                        Update existing (requires confirmation)
                      </option>
                      <option value="create">
                        Create new despite duplicates (requires confirmation)
                      </option>
                    </select>
                  </label>
                  <p className={styles.bulkHint}>
                    Updates never overwrite existing fields with blanks. Soft
                    name + company matches are Warnings and can still create a
                    new candidate unless you choose Update existing.
                  </p>
                </>
              ) : null}

              <div className={styles.bulkActions}>
                <label className={styles.bulkFieldInline}>
                  <span>Filter</span>
                  <select
                    className={styles.bulkSelect}
                    value={statusFilter}
                    onChange={(event) =>
                      setStatusFilter(
                        event.target.value as "" | BulkRowStatus,
                      )
                    }
                  >
                    <option value="">All statuses</option>
                    <option value="Ready">Ready</option>
                    <option value="Warning">Warning</option>
                    <option value="Duplicate">Duplicate</option>
                    <option value="Error">Error</option>
                    <option value="Skipped">Skipped</option>
                    <option value="Imported">Imported</option>
                  </select>
                </label>
                <button
                  type="button"
                  className={styles.secondaryButton}
                  onClick={downloadReport}
                >
                  Download validation report
                </button>
                {!commitResult ? (
                  <button
                    type="button"
                    className={styles.primaryButton}
                    disabled={!canCommit || committing}
                    title={
                      !canCommit
                        ? (commitBlockedReason ?? "Nothing to import")
                        : undefined
                    }
                    onClick={runCommit}
                  >
                    {committing ? "Importing…" : "Confirm import"}
                  </button>
                ) : null}
              </div>
              {commitBlockedReason && !commitResult ? (
                <p className={styles.bulkHint} role="status">
                  {commitBlockedReason}
                </p>
              ) : null}
              {canCommit &&
              !commitResult &&
              preview.summary.duplicateRows > 0 &&
              duplicateMode === "skip" &&
              preview.summary.readyRows === 0 &&
              preview.summary.warningRows === 0 ? (
                <p className={styles.bulkHint} role="status">
                  Tip: for Training Matrix re-upload, set{" "}
                  <strong>If duplicate found</strong> to{" "}
                  <strong>Update existing</strong>, then Confirm import.
                </p>
              ) : null}

              <PreviewTable
                rows={displayRows}
                headers={preview?.headers ?? []}
              />
            </div>
          </section>
        ) : null}

        <section className={styles.settingsCard}>
          <header className={styles.settingsCardHeader}>
            <h2>Instructions</h2>
            <p>Safe import checklist for Workforce / Candidates.</p>
          </header>
          <div className={styles.settingsCardBody}>
            <ol className={styles.bulkInstructions}>
              <li>
                Make sure the company already exists under Companies (for example
                <strong> Hash</strong>).
              </li>
              <li>
                Go to <strong>Bulk Upload</strong> → choose{" "}
                <strong>Workforce / Candidates</strong> → upload{" "}
                <code>Workforce list.xlsx</code> → Preview (all Excel columns
                shown) → Confirm import.
              </li>
              <li>
                Refresh <strong>Workforce / Candidates</strong> — rows should
                appear (Company is a SharePoint lookup; the app now reads it
                correctly).
              </li>
              <li>
                Then choose <strong>Training Matrix rows</strong> → upload{" "}
                <code>Training matrix example.xlsx</code> → Preview (all Excel
                columns shown) → Confirm. Rows match Workforce by Name / DOB;
                N-code dates are stored in Category Records.
              </li>
              <li>
                Required workforce columns: Candidate Name, Company Name.
                Optional: Workforce Number, Date of birth, Department, Status.
              </li>
              <li>
                Duplicates default to Skip. Use Update only when you intend to
                overwrite non-blank fields.
              </li>
              <li>
                Notifications stay suppressed by default. Customers cannot
                access /admin/bulk-upload or the APIs.
              </li>
            </ol>
          </div>
        </section>
      </div>
    </div>
  );
}
