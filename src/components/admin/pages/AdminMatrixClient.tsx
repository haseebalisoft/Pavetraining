"use client";

import Link from "next/link";
import { useCallback, useState, type ReactNode } from "react";
import { useRouter, useSearchParams } from "next/navigation";

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
  EXPIRY_STATUS_LEGEND,
  getExpiryStatus,
  matchesAnyExpiryFilter,
  type ExpiryFilter,
} from "@/lib/training/expiryFilters";
import { isManualOverrideHeader } from "@/lib/training/matrixManualOverrides";
import {
  matrixStatusLabel,
  normalizeMatrixStatus,
} from "@/lib/theme/paveBrand";
import { toneForExpiryStatus } from "@/lib/ui/status";
import { formatDate } from "@/lib/utils/formatDate";
import type { Company } from "@/types/models";

function matrixCell(
  row: AdminMatrixRecord,
  header: string,
): ReactNode {
  if (header === "Name") {
    const name = row.candidateName?.trim() || "—";
    const nameNode = row.workforceId ? (
      <Link
        className={styles.matrixNameLink}
        href={`/admin/workforce/${row.workforceId}`}
      >
        {name}
      </Link>
    ) : (
      <span className={styles.matrixNameCell}>{name}</span>
    );
    // Link-health badge: Linked is the healthy default (no badge). Orphan (no
    // Workforce record) / Needs Review (ambiguous same-name) are flagged so an
    // admin can act. Orphans are hidden unless "Show all" is on.
    const status = row.matrixLinkStatus;
    const badge =
      status === "Orphan" || status === "Needs Review" ? (
        <StatusBadge
          label={status}
          tone={status === "Orphan" ? "missing" : "warn"}
        />
      ) : null;
    return (
      <span className={styles.matrixNameWrap}>
        {nameNode}
        {badge}
      </span>
    );
  }
  if (header === "DOB") {
    const dob = row.columnValues?.DOB ?? row.dateOfBirth;
    return (
      <span className={styles.matrixTextCell}>{formatDate(dob)}</span>
    );
  }
  const fromColumns = row.columnValues?.[header];
  const manual = isManualOverrideHeader(header, row.manualOverrideHeaders);
  return (
    <span className={styles.matrixExpiryCell}>
      <ExpiryDateBadge date={fromColumns} fillCell />
      {manual ? (
        <span className={styles.matrixManualBadge} title="Manually set — register sync will not overwrite">
          Manual
        </span>
      ) : null}
    </span>
  );
}

function rowExpiryDates(row: AdminMatrixRecord): Array<string | null> {
  const values = Object.entries(row.columnValues ?? {})
    .filter(([key]) => key !== "Name" && key !== "DOB")
    .map(([, value]) => value);
  return [row.nextExpiryDate, ...values];
}

/** Name first (sticky), then company, then template headers / status. */
const columns: AdminColumn<AdminMatrixRecord>[] = [
  {
    key: "Name",
    header: "Name",
    render: (row) => matrixCell(row, "Name"),
  },
  {
    key: "company",
    header: "Company",
    render: (row) => (
      <span className={styles.matrixTextCell}>
        {row.companyName?.trim() || "—"}
      </span>
    ),
  },
  {
    key: "workforceNumber",
    header: "Workforce #",
    render: (row) => (
      <span className={styles.matrixTextCell}>
        {row.workforceNumber?.trim() || "—"}
      </span>
    ),
  },
  ...CLIENT_MATRIX_DISPLAY_HEADERS.filter((header) => header !== "Name").map(
    (header) => ({
      key: header,
      header: header === "Face ift" ? "Face Fit Expiry" : header,
      render: (row: AdminMatrixRecord) => matrixCell(row, header),
    }),
  ),
  {
    key: "next",
    header: "Next expiry",
    render: (row: AdminMatrixRecord) => (
      <ExpiryDateBadge date={row.nextExpiryDate} fillCell />
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
    render: (row: AdminMatrixRecord) => {
      const key = normalizeMatrixStatus(row.overallStatus);
      const tone =
        key === "compliant"
          ? "ok"
          : key === "expiringSoon"
            ? "warn"
            : key === "expired"
              ? "danger"
              : "missing";
      return <StatusBadge label={matrixStatusLabel(key)} tone={tone} />;
    },
  },
];

const fields: AdminFieldConfig[] = [
  {
    name: "candidateName",
    label: "Candidate name",
    type: "text",
    required: true,
    section: "Candidate",
  },
  {
    name: "companyName",
    label: "Company (from Workforce)",
    type: "text",
    readOnly: true,
    section: "Candidate",
  },
  {
    name: "dateOfBirth",
    label: "DOB",
    type: "date",
    section: "Candidate",
  },
  {
    name: "cscsExpiry",
    label: "CSCS Expiry",
    type: "date",
    section: "Card expiries",
  },
  {
    name: "nrswaExpiry",
    label: "SWQR / NRSWA Expiry",
    type: "date",
    section: "Card expiries",
  },
  {
    name: "eusrExpiry",
    label: "EUSR Expiry",
    type: "date",
    section: "Card expiries",
  },
  {
    name: "ssstsExpiry",
    label: "SSSTS Expiry",
    type: "date",
    section: "Card expiries",
  },
  {
    name: "smstsExpiry",
    label: "SMSTS Expiry",
    type: "date",
    section: "Card expiries",
  },
  {
    name: "faceFitExpiry",
    label: "Face Fit Expiry",
    type: "date",
    section: "Card expiries",
  },
  {
    name: "n001Expiry",
    label: "N001 - Ind FLT",
    type: "date",
    section: "NPORS categories",
  },
  {
    name: "n003Expiry",
    label: "N003 - Reach Lift Truck",
    type: "date",
    section: "NPORS categories",
  },
  {
    name: "n004Expiry",
    label: "N004 - Lorry Mounted Lift Truck",
    type: "date",
    section: "NPORS categories",
  },
  {
    name: "n010Expiry",
    label: "N010 - Telescopic Handler",
    type: "date",
    section: "NPORS categories",
  },
  {
    name: "n020Expiry",
    label: "N020 - Tiltrotator System",
    type: "date",
    section: "NPORS categories",
  },
  {
    name: "n021Expiry",
    label: "N021 - Suction Excavator",
    type: "date",
    section: "NPORS categories",
  },
  {
    name: "n027Expiry",
    label: "N027 - Excavation Marshal - Banksperson",
    type: "date",
    section: "NPORS categories",
  },
  {
    name: "n100Expiry",
    label: "N100 - Exc Crane",
    type: "date",
    section: "NPORS categories",
  },
  {
    name: "n031Expiry",
    label: "N031 - Asbestos Awareness",
    type: "date",
    section: "In-house / other",
  },
];

function matchesFilter(
  row: AdminMatrixRecord,
  filter: string | null,
  showAll: boolean,
): boolean {
  // Orphan rows (candidate not in Workforce) are hidden by default so the
  // matrix only shows people who exist in Workforce; "Show all" reveals them
  // for auditing. Needs Review / Linked always pass this gate.
  if (!showAll && row.matrixLinkStatus === "Orphan") return false;
  if (!filter || filter === "all") return true;
  if (filter === "review" || filter === "missing") return row.needsReview;
  if (filter === "expiring") {
    return getExpiryStatus(row.nextExpiryDate).status === "urgent";
  }
  return matchesAnyExpiryFilter(rowExpiryDates(row), filter as ExpiryFilter);
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
  const router = useRouter();
  const { pushToast } = useAdminToast();
  const [syncing, setSyncing] = useState(false);
  const [syncResult, setSyncResult] = useState<MatrixSyncResult | null>(null);
  const [syncCompanyId, setSyncCompanyId] = useState("");
  // Orphan rows (matrix rows with no Workforce candidate) are hidden by default;
  // this toggle reveals them for auditing/cleanup.
  const [showAll, setShowAll] = useState(false);
  const orphanCount = initialRows.filter(
    (row) => row.matrixLinkStatus === "Orphan",
  ).length;

  function setExpiryFilter(next: string) {
    const params = new URLSearchParams(searchParams.toString());
    if (!next || next === "all") params.delete("filter");
    else params.set("filter", next);
    const qs = params.toString();
    router.replace(qs ? `?${qs}` : "?");
  }

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
          router.refresh();
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
    [pushToast, router],
  );

  return (
    <>
      <div className={styles.legendRow} aria-label="Expiry colour legend">
        {EXPIRY_STATUS_LEGEND.map((item) => (
          <StatusBadge
            key={item.status}
            label={item.label}
            tone={toneForExpiryStatus(item.status)}
          />
        ))}
      </div>
      <AdminCrudPage<AdminMatrixRecord>
        title="Training Matrix"
        description="Register sync (NPORS / EUSR / Streetworks / In-House Asbestos → N031) and direct admin edits both update this matrix. Cells marked Manual are not overwritten by register sync. Pass updates expiry when newer; Fail never extends."
        columns={columns}
        fields={fields}
        companies={companies}
        initialRows={initialRows}
        enableCompanyFilter
        getCompanyName={(row) => row.companyName}
        drawerWide
        wideTable
        stickyLeadColumns
        tableClassName={styles.matrixTable}
        listUrl="/api/admin/training-matrix"
        updateUrl={(id) => `/api/admin/training-matrix/${id}`}
        deleteUrl={(id) => `/api/admin/training-matrix/${id}`}
        allowCreate={false}
        mapResponse={(payload) =>
          ((payload as { records?: AdminMatrixRecord[] }).records ?? [])
        }
        rowFilter={(row) => matchesFilter(row, filter, showAll)}
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
              <span className={styles.fieldLabel}>Expiry filter</span>
              <select
                className={styles.select}
                value={filter || "all"}
                onChange={(event) => setExpiryFilter(event.target.value)}
              >
                <option value="all">All expiries</option>
                <option value="expired">Expired (red)</option>
                <option value="within-3m">Expiring within 3 months (0–90 days, red)</option>
                <option value="within-6m">Expiring within 6 months (0–180 days, red + amber)</option>
                <option value="6m-plus">6 months or more / in date (181+ days, green)</option>
                <option value="review">Records to Review (missing dates)</option>
              </select>
            </label>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Link status</span>
              <span className={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={showAll}
                  onChange={(event) => setShowAll(event.target.checked)}
                />
                <span>
                  Show all
                  {orphanCount > 0
                    ? ` (incl. ${orphanCount} orphan${orphanCount === 1 ? "" : "s"})`
                    : " (incl. orphans)"}
                </span>
              </span>
            </label>
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
