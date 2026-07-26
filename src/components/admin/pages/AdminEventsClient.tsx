"use client";

import { useCallback, useState } from "react";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import styles from "@/components/admin/admin.module.css";
import { useAdminToast } from "@/components/admin/AdminToast";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { AdminEventRecord } from "@/lib/services/adminCrudService";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type { Company } from "@/types/models";

function formatDateTime(value: string | null | undefined): string {
  if (!value?.trim()) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  const day = formatDisplayDate(value);
  const time = new Intl.DateTimeFormat("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
  return `${day} ${time}`;
}

function syncTone(
  status: string | null | undefined,
): "ok" | "warn" | "danger" | "neutral" {
  const normalized = (status ?? "").trim().toLowerCase();
  if (normalized === "synced") return "ok";
  if (normalized === "pending") return "warn";
  if (normalized === "failed") return "danger";
  if (normalized === "skipped") return "neutral";
  return "neutral";
}

const columns: AdminColumn<AdminEventRecord>[] = [
  { key: "title", header: "Title", render: (row) => row.title },
  { key: "company", header: "Company", render: (row) => row.company ?? "—" },
  {
    key: "start",
    header: "Start Date",
    render: (row) => formatDateTime(row.eventDate),
  },
  {
    key: "end",
    header: "End Date",
    render: (row) => formatDateTime(row.endDate),
  },
  {
    key: "location",
    header: "Location",
    render: (row) => row.location ?? "—",
  },
  {
    key: "visible",
    header: "Customer Visible",
    render: (row) => (row.customerVisible ? "Yes" : "No"),
  },
  {
    key: "doNotSync",
    header: "Do Not Sync",
    render: (row) => (row.doNotSync ? "Yes" : "No"),
  },
  {
    key: "syncStatus",
    header: "Sync Status",
    render: (row) => (
      <StatusBadge
        label={row.syncStatus?.trim() || "—"}
        tone={syncTone(row.syncStatus)}
      />
    ),
  },
  {
    key: "lastSynced",
    header: "Last Synced At",
    render: (row) => formatDateTime(row.lastSyncedAt),
  },
  {
    key: "syncError",
    header: "Sync Error",
    render: (row) =>
      row.syncStatus?.toLowerCase() === "failed"
        ? row.syncError?.trim() || "—"
        : "—",
  },
];

const fields: AdminFieldConfig[] = [
  {
    name: "title",
    label: "Event Title",
    type: "text",
    required: true,
    section: "Event details",
  },
  {
    name: "companyId",
    label: "Company",
    type: "company",
    required: true,
  },
  {
    name: "customerVisible",
    label: "Customer Visible",
    type: "boolean",
  },
  {
    name: "eventDate",
    label: "Start Date/Time",
    type: "datetime",
    required: true,
  },
  {
    name: "endDate",
    label: "End Date/Time",
    type: "datetime",
  },
  { name: "location", label: "Location", type: "text" },
  { name: "trainingAddress", label: "Training Address", type: "text" },
  { name: "description", label: "Description", type: "textarea" },
  {
    name: "doNotSync",
    label: "Do Not Sync",
    type: "boolean",
    section: "Outlook sync",
  },
  {
    name: "syncStatus",
    label: "Sync Status",
    type: "text",
    readOnly: true,
  },
  {
    name: "lastSyncedAt",
    label: "Last Synced At",
    type: "text",
    readOnly: true,
  },
  {
    name: "syncError",
    label: "Sync Error",
    type: "textarea",
    readOnly: true,
  },
];

export function AdminEventsClient({
  companies,
  initialRows,
  warnings = [],
}: {
  companies: Company[];
  initialRows: AdminEventRecord[];
  warnings?: string[];
}) {
  const { pushToast } = useAdminToast();
  const [busyId, setBusyId] = useState<string | null>(null);

  const runSync = useCallback(
    async (row: AdminEventRecord, mode: "sync" | "retry") => {
      const path =
        mode === "retry"
          ? `/api/admin/events/${row.id}/retry-sync`
          : `/api/admin/events/${row.id}/sync`;
      setBusyId(row.id);
      try {
        const response = await fetch(path, { method: "POST" });
        if (!response.ok) {
          throw new Error(await readPublicApiError(response));
        }
        const payload = (await response.json()) as {
          result?: {
            status?: string;
            error?: string | null;
            reason?: string | null;
          };
        };
        const result = payload.result;
        const status = result?.status ?? "Unknown";
        if (status === "Failed") {
          pushToast(
            result?.error || `Sync failed for “${row.title}”.`,
            "error",
          );
        } else if (status === "Skipped") {
          pushToast(result?.reason || `Sync skipped for “${row.title}”.`);
        } else {
          pushToast(
            `Outlook sync ${status.toLowerCase()} for “${row.title}”.`,
          );
        }
      } catch (error) {
        pushToast(
          error instanceof Error ? error.message : "Sync request failed.",
          "error",
        );
      } finally {
        setBusyId(null);
      }
    },
    [pushToast],
  );

  const markDoNotSync = useCallback(
    async (row: AdminEventRecord, reload: () => Promise<void>) => {
      setBusyId(row.id);
      try {
        const response = await fetch(`/api/admin/events/${row.id}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ doNotSync: true }),
        });
        if (!response.ok) {
          throw new Error(await readPublicApiError(response));
        }
        pushToast(`“${row.title}” marked Do Not Sync.`);
        await reload();
      } catch (error) {
        pushToast(
          error instanceof Error
            ? error.message
            : "Failed to mark Do Not Sync.",
          "error",
        );
      } finally {
        setBusyId(null);
      }
    },
    [pushToast],
  );

  return (
    <AdminCrudPage<AdminEventRecord>
      title="Events"
      description="Manage training events. SharePoint Events is the source of truth; Outlook sync is one-way from this portal."
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      warnings={warnings}
      enableCompanyFilter
      getCompanyName={(row) => row.company}
      drawerWide
      emptyLabel="No events found. Create an event to begin."
      listUrl="/api/admin/events"
      createUrl="/api/admin/events"
      updateUrl={(id) => `/api/admin/events/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminEventRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.title,
        (row) => row.company,
        (row) => row.trainingAddress,
        (row) => row.location,
        (row) => row.description,
        (row) => row.syncStatus,
        (row) => row.syncError,
      ]}
      extraActions={(row, { reload }) => (
        <>
          <button
            type="button"
            className={styles.linkButton}
            disabled={busyId === row.id || row.doNotSync}
            onClick={() => {
              void (async () => {
                await runSync(row, "sync");
                await reload();
              })();
            }}
          >
            {busyId === row.id ? "Working…" : "Sync now"}
          </button>
          {row.syncStatus?.toLowerCase() === "failed" ? (
            <>
              {" · "}
              <button
                type="button"
                className={styles.linkButton}
                disabled={busyId === row.id || row.doNotSync}
                onClick={() => {
                  void (async () => {
                    await runSync(row, "retry");
                    await reload();
                  })();
                }}
              >
                Retry sync
              </button>
            </>
          ) : null}
          {row.syncError?.trim() &&
          row.syncStatus?.toLowerCase() === "failed" ? (
            <>
              {" · "}
              <button
                type="button"
                className={styles.linkButton}
                onClick={() => {
                  pushToast(row.syncError || "No sync error details.", "error");
                }}
              >
                View sync error
              </button>
            </>
          ) : null}
          {!row.doNotSync ? (
            <>
              {" · "}
              <button
                type="button"
                className={styles.linkButton}
                disabled={busyId === row.id}
                onClick={() => {
                  void markDoNotSync(row, reload);
                }}
              >
                Mark Do Not Sync
              </button>
            </>
          ) : null}
        </>
      )}
    />
  );
}
