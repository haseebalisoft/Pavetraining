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
  implemented: boolean;
}> = [
  {
    value: "company",
    label: "Companies",
    hint: "Use Company list.xlsx exactly. Preview shows every Excel column. Creates or updates Company List rows.",
    implemented: true,
  },
  {
    value: "workforce",
    label: "Workforce / Candidates",
    hint: "Use Workforce list.xlsx exactly. Preview shows every Excel column. Missing companies are created on import.",
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
    const fromSource = row.source?.[header];
    if (fromSource != null && String(fromSource).trim() !== "") {
      return String(fromSource);
    }
    // Fallbacks for mapped internal keys when source cell is blank/missing.
    const key = header.trim().toLowerCase();
    const mapped: Record<string, string | null | undefined> = {
      "candidate name": row.fields.candidateName,
      name: row.fields.candidateName,
      "company name": row.fields.company,
      company: row.fields.company,
      "workforce number": row.fields.workforceNumber,
      "date of birth": row.fields.dateOfBirth,
      dob: row.fields.dateOfBirth,
    };
    const hit = mapped[key] ?? row.fields[header] ?? null;
    return hit?.trim() ? hit : "—";
  };

  return (
    <div className={`${styles.tableWrap} ${styles.bulkPreviewTableWrap}`}>
      <table className={`${styles.dataTable} ${styles.bulkPreviewTable}`}>
        <thead>
          <tr>
            <th>Row</th>
            <th>Status</th>
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
            source: row.source,
          })),
        }),
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const data = (await response.json()) as BulkCommitResult;
      setCommitResult(data);
      const firstError = data.rows.find(
        (row) => row.status === "Error" && row.messages.length,
      );
      if (data.summary.importedRows === 0 && data.summary.errorRows > 0) {
        pushToast(
          firstError?.messages.slice(-1)[0] ??
            `Import failed for ${data.summary.errorRows} row(s). Check Messages in the table.`,
          "error",
        );
      } else {
        pushToast(
          `Import finished: ${data.summary.importedRows} imported, ${data.summary.skippedRows} skipped, ${data.summary.errorRows} errors.`,
          data.summary.errorRows > 0 ? "error" : "success",
        );
      }
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
    const header = ["Row", "Status", ...dataHeaders, "Matched Id", "Messages"];
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
          ...cells,
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
                onChange={(event) => {
                  setImportType(event.target.value as BulkImportType);
                  setPreview(null);
                  setCommitResult(null);
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
