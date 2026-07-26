"use client";

import { useCallback, useEffect, useMemo, useState } from "react";

import { AdminDrawer } from "@/components/admin/AdminDrawer";
import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { LoadingState } from "@/components/ui/States";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { AuditLogRecord } from "@/types/audit";

import styles from "@/components/admin/admin.module.css";

type LogsPayload = {
  logs: AuditLogRecord[];
  configured: boolean;
  usingConsoleFallback: boolean;
  exportEnabled: boolean;
};

const ACTION_OPTIONS = [
  "",
  "LOGIN_SUCCESS",
  "LOGIN_DENIED",
  "ACCESS_DENIED",
  "ADMIN_CREATE",
  "ADMIN_UPDATE",
  "ADMIN_DELETE",
  "DOCUMENT_VIEW",
  "DOCUMENT_DOWNLOAD",
  "DOCUMENT_UPLOAD",
  "CANDIDATE_VIEW",
  "BULK_UPLOAD_PREVIEW",
  "BULK_UPLOAD_COMMIT",
  "NOTIFICATION_SENT",
  "NOTIFICATION_FAILED",
  "MATRIX_SYNC_STARTED",
  "MATRIX_SYNC_COMPLETED",
  "MATRIX_SYNC_FAILED",
  "SYSTEM_ERROR",
  "SETTINGS_UPDATE",
];

export function AdminLogsClient() {
  const { pushToast } = useAdminToast();
  const [loading, setLoading] = useState(true);
  const [payload, setPayload] = useState<LogsPayload | null>(null);
  const [search, setSearch] = useState("");
  const [action, setAction] = useState("");
  const [success, setSuccess] = useState<"all" | "true" | "false">("all");
  const [entityType, setEntityType] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [selected, setSelected] = useState<AuditLogRecord | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (search.trim()) params.set("search", search.trim());
      if (action) params.set("action", action);
      if (success !== "all") params.set("success", success);
      if (entityType.trim()) params.set("entityType", entityType.trim());
      if (from) params.set("from", from);
      if (to) params.set("to", to);

      const response = await fetch(`/api/admin/logs?${params.toString()}`, {
        cache: "no-store",
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const data = (await response.json()) as LogsPayload;
      setPayload(data);
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to load logs",
        "error",
      );
    } finally {
      setLoading(false);
    }
  }, [action, entityType, from, pushToast, search, success, to]);

  useEffect(() => {
    void load();
  }, [load]);

  const entityTypes = useMemo(() => {
    const set = new Set<string>();
    for (const row of payload?.logs ?? []) {
      if (row.entityType) set.add(row.entityType);
    }
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [payload?.logs]);

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs
            items={[
              { label: "Admin", href: "/admin" },
              { label: "Audit / Activity Log" },
            ]}
          />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>Audit / Activity Log</h1>
          <p className={styles.subtitle}>
            Track logins, admin changes, document actions, denied access, and
            system events. Export is disabled for now.
          </p>
        </div>
        <button
          type="button"
          className={styles.secondaryButton}
          onClick={() => void load()}
          disabled={loading}
        >
          Refresh
        </button>
      </header>

      <div
        className={`${styles.settingsBanner} ${
          payload?.usingConsoleFallback
            ? styles.settingsBannerWarn
            : styles.settingsBannerOk
        }`}
        role="status"
      >
        {payload?.usingConsoleFallback
          ? "Training Manager Logs list not configured — writing to console only."
          : "Reading from SharePoint Training Manager Logs."}
        {" · "}
        Retention: no automatic deletion yet.
      </div>

      <div className={styles.crudToolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>User email</span>
          <input
            className={styles.input}
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="Search email"
          />
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Action</span>
          <select
            className={styles.select}
            value={action}
            onChange={(event) => setAction(event.target.value)}
          >
            <option value="">All actions</option>
            {ACTION_OPTIONS.filter(Boolean).map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Result</span>
          <select
            className={styles.select}
            value={success}
            onChange={(event) =>
              setSuccess(event.target.value as "all" | "true" | "false")
            }
          >
            <option value="all">All</option>
            <option value="true">Success</option>
            <option value="false">Failure</option>
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Entity type</span>
          <input
            className={styles.input}
            list="audit-entity-types"
            value={entityType}
            onChange={(event) => setEntityType(event.target.value)}
            placeholder="e.g. Customer Documents"
          />
          <datalist id="audit-entity-types">
            {entityTypes.map((value) => (
              <option key={value} value={value} />
            ))}
          </datalist>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>From</span>
          <input
            className={styles.input}
            type="date"
            value={from}
            onChange={(event) => setFrom(event.target.value)}
          />
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>To</span>
          <input
            className={styles.input}
            type="date"
            value={to}
            onChange={(event) => setTo(event.target.value)}
          />
        </label>
      </div>

      {loading ? (
        <LoadingState label="Loading audit logs…" />
      ) : (payload?.logs.length ?? 0) === 0 ? (
        <div className={styles.emptyState}>
          <h2>No log entries</h2>
          <p>No audit rows matched these filters.</p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">When</th>
                <th scope="col">User</th>
                <th scope="col">Action</th>
                <th scope="col">Entity</th>
                <th scope="col">Company</th>
                <th scope="col">Result</th>
                <th scope="col">Actions</th>
              </tr>
            </thead>
            <tbody>
              {payload!.logs.map((row) => (
                <tr key={row.id}>
                  <td>{new Date(row.timestamp).toLocaleString()}</td>
                  <td>
                    <div>{row.userEmail}</div>
                    <div className={styles.mutedNote}>{row.roleType ?? "—"}</div>
                  </td>
                  <td>{row.action}</td>
                  <td>
                    <div>{row.entityType}</div>
                    <div className={styles.mutedNote}>
                      {row.entityName ?? row.entityId ?? "—"}
                    </div>
                  </td>
                  <td>{row.company ?? "—"}</td>
                  <td>
                    <StatusBadge
                      label={row.success ? "Success" : "Failed"}
                      tone={row.success ? "ok" : "danger"}
                    />
                  </td>
                  <td>
                    <button
                      type="button"
                      className={styles.linkButton}
                      onClick={() => setSelected(row)}
                    >
                      View
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <AdminDrawer
        open={Boolean(selected)}
        title="Log details"
        onClose={() => setSelected(null)}
        wide
        footer={
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={() => setSelected(null)}
          >
            Close
          </button>
        }
      >
        {selected ? (
          <div className={styles.settingsFieldGrid}>
            <Detail label="Timestamp" value={new Date(selected.timestamp).toLocaleString()} />
            <Detail label="User email" value={selected.userEmail} />
            <Detail label="Role" value={selected.roleType ?? "—"} />
            <Detail label="Company" value={selected.company ?? "—"} />
            <Detail label="Action" value={selected.action} />
            <Detail label="Entity type" value={selected.entityType} />
            <Detail label="Entity id" value={selected.entityId ?? "—"} />
            <Detail label="Entity name" value={selected.entityName ?? "—"} />
            <Detail
              label="Result"
              value={selected.success ? "Success" : "Failed"}
            />
            <Detail
              label="Error"
              value={selected.errorMessage ?? "—"}
            />
            <Detail label="IP address" value={selected.ipAddress ?? "—"} />
            <Detail label="User agent" value={selected.userAgent ?? "—"} />
            <Detail
              label="Metadata"
              value={
                selected.metadata
                  ? JSON.stringify(selected.metadata, null, 2)
                  : "—"
              }
            />
          </div>
        ) : null}
      </AdminDrawer>
    </div>
  );
}

function Detail({ label, value }: { label: string; value: string }) {
  return (
    <div className={styles.field}>
      <span className={styles.fieldLabel}>{label}</span>
      <pre className={styles.logDetailValue}>{value}</pre>
    </div>
  );
}
