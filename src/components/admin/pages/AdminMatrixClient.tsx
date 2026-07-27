"use client";

import { useCallback, useState } from "react";
import { useSearchParams } from "next/navigation";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import styles from "@/components/admin/admin.module.css";
import { useAdminToast } from "@/components/admin/AdminToast";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { AdminMatrixRecord } from "@/lib/services/adminCrudService";
import { CLIENT_MATRIX_DISPLAY_HEADERS } from "@/lib/services/bulkUpload/clientTemplateHeaders";
import type { MatrixSyncResult } from "@/types/matrixSync";
import {
  getExpiryStatus,
  matchesExpiryFilter,
  type ExpiryFilter,
} from "@/lib/training/expiryFilters";
import type { Company } from "@/types/models";

function formatDateCell(value: string | null | undefined): string {
  if (!value?.trim()) return "—";
  if (/^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0, 10);
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return parsed.toISOString().slice(0, 10);
}

function matrixCell(
  row: AdminMatrixRecord,
  header: string,
): string {
  if (header === "Name") return row.candidateName || "—";
  if (header === "DOB") return formatDateCell(row.dateOfBirth);
  const fromColumns = row.columnValues?.[header];
  if (fromColumns?.trim()) return formatDateCell(fromColumns);
  return "—";
}

/** Match Training matrix example.xlsx column order. */
const columns: AdminColumn<AdminMatrixRecord>[] = [
  {
    key: "company",
    header: "Company",
    render: (row) => row.companyName ?? "—",
  },
  ...CLIENT_MATRIX_DISPLAY_HEADERS.map((header) => ({
    key: header,
    header,
    render: (row: AdminMatrixRecord) => matrixCell(row, header),
  })),
  {
    key: "next",
    header: "Next expiry",
    render: (row: AdminMatrixRecord) => (
      <ExpiryDateBadge date={row.nextExpiryDate} />
    ),
  },
  {
    key: "review",
    header: "Records to Review",
    render: (row: AdminMatrixRecord) => (
      <StatusBadge
        label={row.needsReview ? "Review" : "Clear"}
        tone={row.needsReview ? "warn" : "ok"}
      />
    ),
  },
  {
    key: "status",
    header: "Overall status",
    render: (row: AdminMatrixRecord) => row.overallStatus ?? "—",
  },
];

const fields: AdminFieldConfig[] = [
  { name: "candidateName", label: "Candidate name", type: "text", required: true },
  { name: "companyName", label: "Company", type: "company", required: true },
  { name: "department", label: "Department", type: "text" },
  { name: "overallStatus", label: "Overall status", type: "text" },
  { name: "needsReview", label: "Records to Review", type: "boolean" },
  { name: "matrixNotes", label: "Notes", type: "textarea" },
  { name: "nextExpiryDate", label: "Next expiry date", type: "date" },
  { name: "n001Expiry", label: "N001 expiry", type: "date" },
  { name: "n003Expiry", label: "N003 expiry", type: "date" },
  { name: "n004Expiry", label: "N004 expiry", type: "date" },
  { name: "n010Expiry", label: "N010 expiry", type: "date" },
  { name: "n020Expiry", label: "N020 expiry", type: "date" },
  { name: "n021Expiry", label: "N021 expiry", type: "date" },
  { name: "n027Expiry", label: "N027 expiry", type: "date" },
  { name: "n100Expiry", label: "N100 expiry", type: "date" },
];

function matchesFilter(
  row: AdminMatrixRecord,
  filter: string | null,
): boolean {
  if (!filter || filter === "all") return true;
  if (filter === "review") return row.needsReview;
  if (filter === "expiring") {
    return getExpiryStatus(row.nextExpiryDate).status === "urgent";
  }
  return matchesExpiryFilter(row.nextExpiryDate, filter as ExpiryFilter);
}

function SyncResultsPanel({ result }: { result: MatrixSyncResult }) {
  return (
    <div className={styles.syncPanel} role="status">
      <div className={styles.syncPanelHeader}>
        <strong>
          {result.dryRun ? "Dry run" : "Sync"} · {result.scope}
        </strong>
        <span>
          Updated {result.summary.updated} · Created {result.summary.created} ·
          Skipped {result.summary.skipped} · Errors {result.summary.errors}
          {result.summary.warnings > 0
            ? ` · Warnings ${result.summary.warnings}`
            : ""}
        </span>
      </div>
      <div className={styles.tableWrap}>
        <table className={styles.dataTable}>
          <thead>
            <tr>
              <th scope="col">Candidate</th>
              <th scope="col">Company</th>
              <th scope="col">Register source</th>
              <th scope="col">Matrix row</th>
              <th scope="col">Fields updated</th>
              <th scope="col">Warnings</th>
              <th scope="col">Errors</th>
            </tr>
          </thead>
          <tbody>
            {result.items.length === 0 ? (
              <tr>
                <td colSpan={7}>No sync items.</td>
              </tr>
            ) : (
              result.items.map((item, index) => (
                <tr key={`${item.candidateId ?? item.candidate}-${index}`}>
                  <td>{item.candidate}</td>
                  <td>{item.company}</td>
                  <td>
                    {item.registerSources.length
                      ? item.registerSources.join(", ")
                      : "—"}
                  </td>
                  <td>
                    {item.matrixRowCreated
                      ? "Created"
                      : item.matrixRowFound
                        ? `Found (#${item.matrixRowId})`
                        : "—"}
                  </td>
                  <td>
                    {item.fieldsUpdated.length
                      ? item.fieldsUpdated.join(", ")
                      : item.skipReason ?? "—"}
                  </td>
                  <td>
                    {item.warnings.length ? item.warnings.join(" · ") : "—"}
                  </td>
                  <td>
                    {item.errors.length ? item.errors.join(" · ") : "—"}
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

export function AdminMatrixClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminMatrixRecord[];
}) {
  const searchParams = useSearchParams();
  const filter = searchParams.get("filter");
  const { pushToast } = useAdminToast();
  const [syncing, setSyncing] = useState(false);
  const [syncResult, setSyncResult] = useState<MatrixSyncResult | null>(null);
  const [reloadToken, setReloadToken] = useState(0);
  const [syncCompanyId, setSyncCompanyId] = useState("");

  const runSync = useCallback(
    async (body: Record<string, unknown>, label: string) => {
      setSyncing(true);
      try {
        const response = await fetch("/api/admin/training-matrix/sync", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });
        if (!response.ok) {
          throw new Error(await readPublicApiError(response));
        }
        const payload = (await response.json()) as { result: MatrixSyncResult };
        setSyncResult(payload.result);
        const { summary } = payload.result;
        if (summary.errors > 0) {
          pushToast(
            `${label}: completed with ${summary.errors} error(s).`,
            "error",
          );
        } else {
          pushToast(
            `${label}: updated ${summary.updated}, created ${summary.created}.`,
            "success",
          );
        }
        if (!body.dryRun) {
          setReloadToken((value) => value + 1);
        }
      } catch (error) {
        pushToast(
          error instanceof Error ? error.message : `${label} failed.`,
          "error",
        );
      } finally {
        setSyncing(false);
      }
    },
    [pushToast],
  );

  return (
    <>
      <AdminCrudPage<AdminMatrixRecord>
        key={reloadToken}
        title="Training Matrix"
        description="Columns match Training matrix example.xlsx. Scroll horizontally to see all N-code expiries. Sync still updates from registers."
        columns={columns}
        fields={fields}
        companies={companies}
        initialRows={initialRows}
        enableCompanyFilter
        getCompanyName={(row) => row.companyName}
        drawerWide
        listUrl="/api/admin/training-matrix"
        createUrl="/api/admin/training-matrix"
        updateUrl={(id) => `/api/admin/training-matrix/${id}`}
        mapResponse={(payload) =>
          ((payload as { records?: AdminMatrixRecord[] }).records ?? [])
        }
        rowFilter={(row) => matchesFilter(row, filter)}
        rowClassName={(row) => (row.needsReview ? styles.reviewRow : undefined)}
        searchKeys={[
          (row) => row.candidateName,
          (row) => row.companyName,
          (row) => row.department,
          (row) => row.overallStatus,
        ]}
        toolbarExtra={
          <div className={styles.syncToolbar}>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Sync company</span>
              <select
                className={styles.select}
                value={syncCompanyId}
                onChange={(event) => setSyncCompanyId(event.target.value)}
                disabled={syncing}
              >
                <option value="">Select company…</option>
                {companies.map((company) => (
                  <option key={company.id} value={company.id}>
                    {company.companyName}
                  </option>
                ))}
              </select>
            </label>
            <button
              type="button"
              className={styles.secondaryButton}
              disabled={syncing || !syncCompanyId}
              onClick={() =>
                void runSync(
                  { companyId: syncCompanyId },
                  "Sync company records",
                )
              }
            >
              Sync company
            </button>
            <button
              type="button"
              className={styles.secondaryButton}
              disabled={syncing}
              onClick={() => void runSync({ dryRun: true }, "Matrix dry run")}
            >
              {syncing ? "Working…" : "Dry run"}
            </button>
            <button
              type="button"
              className={styles.primaryButton}
              disabled={syncing}
              onClick={() => void runSync({}, "Sync Matrix")}
            >
              {syncing ? "Syncing…" : "Sync Matrix"}
            </button>
          </div>
        }
        extraActions={(row, { reload }) => (
          <button
            type="button"
            className={styles.linkButton}
            disabled={syncing || !row.companyName}
            onClick={() => {
              void (async () => {
                await runSync(
                  {
                    candidateName: row.candidateName,
                    companyName: row.companyName,
                  },
                  `Sync ${row.candidateName}`,
                );
                await reload();
              })();
            }}
          >
            Sync candidate
          </button>
        )}
      />
      {syncResult ? <SyncResultsPanel result={syncResult} /> : null}
    </>
  );
}
