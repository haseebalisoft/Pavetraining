"use client";

import { useCallback, useMemo, useState, type DragEvent } from "react";

import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { StatusTone } from "@/lib/ui/status";
import type {
  BulkCommitResult,
  BulkDuplicateMode,
  BulkImportSummary,
  BulkImportType,
  BulkPreviewResult,
  BulkPreviewRow,
  BulkRowStatus,
} from "@/types/bulkUpload";

import styles from "@/components/admin/admin.module.css";

const IMPORT_OPTIONS: Array<{
  value: BulkImportType;
  label: string;
  hint: string;
}> = [
  {
    value: "workforce",
    label: "Workforce / Candidates",
    hint: "Use your Workforce List CSV. Required: Candidate Name + Company Name. Import this first.",
  },
  {
    value: "trainingMatrix",
    label: "Training Matrix rows",
    hint: "Use your Training Matrix CSV. Candidates must already exist (match Workforce Number).",
  },
  {
    value: "npors",
    label: "NPORS records",
    hint: "Coming next — template available.",
  },
  {
    value: "eusr",
    label: "EUSR records",
    hint: "Coming next — template available.",
  },
  {
    value: "streetworks",
    label: "Streetworks / NRSWA",
    hint: "Coming next — template available.",
  },
  {
    value: "inHouse",
    label: "In-House records",
    hint: "Coming next — template available.",
  },
  {
    value: "nvq",
    label: "NVQ records",
    hint: "Coming next — template available.",
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

function PreviewTable({ rows }: { rows: BulkPreviewRow[] }) {
  if (!rows.length) {
    return <p className={styles.bulkEmpty}>No rows to show.</p>;
  }
  return (
    <div className={styles.tableWrap}>
      <table className={styles.dataTable}>
        <thead>
          <tr>
            <th>Row</th>
            <th>Status</th>
            <th>Candidate</th>
            <th>Company</th>
            <th>Workforce #</th>
            <th>DOB</th>
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
              <td>{row.fields.candidateName ?? "—"}</td>
              <td>{row.fields.company ?? "—"}</td>
              <td>{row.fields.workforceNumber ?? "—"}</td>
              <td>{row.fields.dateOfBirth ?? "—"}</td>
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

export function AdminBulkUploadClient() {
  const { pushToast } = useAdminToast();
  const [importType, setImportType] = useState<BulkImportType>("workforce");
  const [file, setFile] = useState<File | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [suppressNotifications, setSuppressNotifications] = useState(true);
  const [duplicateMode, setDuplicateMode] =
    useState<BulkDuplicateMode>("skip");
  const [previewing, setPreviewing] = useState(false);
  const [committing, setCommitting] = useState(false);
  const [preview, setPreview] = useState<BulkPreviewResult | null>(null);
  const [commitResult, setCommitResult] = useState<BulkCommitResult | null>(
    null,
  );
  const [statusFilter, setStatusFilter] = useState<"" | BulkRowStatus>("");

  const selectedOption = IMPORT_OPTIONS.find((o) => o.value === importType);

  const displayRows = useMemo(() => {
    const source = commitResult?.rows ?? preview?.rows ?? [];
    if (!statusFilter) return source;
    return source.filter((row) => row.status === statusFilter);
  }, [commitResult, preview, statusFilter]);

  const activeSummary = commitResult?.summary ?? preview?.summary ?? null;

  const canCommit =
    !!preview?.implemented &&
    !!preview.rows.length &&
    !commitResult &&
    (preview.summary.readyRows > 0 ||
      preview.summary.warningRows > 0 ||
      (duplicateMode !== "skip" && preview.summary.duplicateRows > 0));

  const onFileChosen = (next: File | null) => {
    setFile(next);
    setPreview(null);
    setCommitResult(null);
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    const dropped = event.dataTransfer.files?.[0] ?? null;
    if (!dropped) return;
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
    setCommitResult(null);
    try {
      const form = new FormData();
      form.set("file", file);
      form.set("importType", importType);
      form.set(
        "suppressNotifications",
        suppressNotifications ? "true" : "false",
      );
      const response = await fetch("/api/admin/bulk-upload/preview", {
        method: "POST",
        body: form,
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const data = (await response.json()) as BulkPreviewResult;
      setPreview(data);
      if (!data.implemented) {
        pushToast(data.message ?? "This import type is not ready yet.", "error");
      } else {
        pushToast(
          `Validated ${data.summary.totalRows} row(s). Review before importing.`,
          "success",
        );
      }
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Preview failed",
        "error",
      );
    } finally {
      setPreviewing(false);
    }
  }, [file, importType, pushToast, suppressNotifications]);

  const runCommit = useCallback(async () => {
    if (!preview?.implemented) return;
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

    setCommitting(true);
    try {
      const response = await fetch("/api/admin/bulk-upload/commit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          importType: preview.importType,
          fileName: preview.fileName,
          duplicateMode,
          suppressNotifications,
          rows: preview.rows.map((row) => ({
            rowNumber: row.rowNumber,
            fields: row.fields,
          })),
        }),
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const data = (await response.json()) as BulkCommitResult;
      setCommitResult(data);
      pushToast(
        `Import finished: ${data.summary.importedRows} imported, ${data.summary.skippedRows} skipped, ${data.summary.errorRows} errors.`,
        "success",
      );
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Import failed",
        "error",
      );
    } finally {
      setCommitting(false);
    }
  }, [duplicateMode, preview, pushToast, suppressNotifications]);

  const downloadTemplate = useCallback(
    async (type: BulkImportType) => {
      try {
        const response = await fetch(
          `/api/admin/bulk-upload/templates?download=${encodeURIComponent(type)}`,
          { cache: "no-store" },
        );
        if (!response.ok) throw new Error(await readPublicApiError(response));
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
    const header = [
      "Row",
      "Status",
      "Candidate Name",
      "Company",
      "Workforce Number",
      "DOB",
      "Department",
      "Matched Id",
      "Messages",
    ];
    const escape = (value: string) =>
      /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
    const lines = [header.map(escape).join(",")];
    for (const row of rows) {
      lines.push(
        [
          String(row.rowNumber),
          row.status,
          row.fields.candidateName ?? "",
          row.fields.company ?? "",
          row.fields.workforceNumber ?? "",
          row.fields.dateOfBirth ?? "",
          row.fields.department ?? "",
          row.matchedEntityId ?? "",
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
    setCommitResult(null);
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
              Start with Workforce / Candidates. Other types expose templates
              now and will reuse the same preview → confirm flow.
            </p>
          </header>
          <div className={styles.settingsCardBody}>
            <label className={styles.bulkField}>
              <span>Import type</span>
              <select
                className={styles.bulkSelect}
                value={importType}
                onChange={(event) => {
                  setImportType(event.target.value as BulkImportType);
                  setPreview(null);
                  setCommitResult(null);
                }}
              >
                {IMPORT_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                    {option.value === "workforce" ||
                    option.value === "trainingMatrix"
                      ? ""
                      : " (coming next)"}
                  </option>
                ))}
              </select>
            </label>
            {selectedOption ? (
              <p className={styles.bulkHint}>{selectedOption.hint}</p>
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
            <p>Accepted formats: .xlsx and .csv. Max 15 MB.</p>
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

            <label className={styles.bulkCheckRow}>
              <input
                type="checkbox"
                checked={suppressNotifications}
                onChange={(event) =>
                  setSuppressNotifications(event.target.checked)
                }
              />
              <span>
                Suppress customer notifications during import (recommended)
              </span>
            </label>

            <div className={styles.bulkActions}>
              <button
                type="button"
                className={styles.primaryButton}
                disabled={!file || previewing}
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
              <h2>{commitResult ? "4. Import summary" : "3. Preview results"}</h2>
              <p>
                {commitResult
                  ? commitResult.message ??
                    "Import finished. Download the validation report for your records."
                  : "Review Ready, Warning, Duplicate, and Error rows before confirming import. Invalid rows are never imported."}
              </p>
            </header>
            <div className={styles.settingsCardBody}>
              {activeSummary ? <SummaryCards summary={activeSummary} /> : null}

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
                    onClick={() => void runCommit()}
                  >
                    {committing ? "Importing…" : "Confirm import"}
                  </button>
                ) : null}
              </div>

              <PreviewTable rows={displayRows} />
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
                <code>Workforce List (2).csv</code> → Preview → Confirm import.
              </li>
              <li>
                Refresh <strong>Workforce / Candidates</strong> — rows should
                appear (Company is a SharePoint lookup; the app now reads it
                correctly).
              </li>
              <li>
                Then choose <strong>Training Matrix rows</strong> → upload{" "}
                <code>Training Matrix (6).csv</code> → Preview → Confirm.
                Matrix rows link to existing candidates by Workforce Number.
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
