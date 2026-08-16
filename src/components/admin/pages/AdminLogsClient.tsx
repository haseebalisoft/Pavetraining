"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { AdminDrawer } from "@/components/admin/AdminDrawer";
import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { LoadingState } from "@/components/ui/States";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import { formatDateTime } from "@/lib/utils/formatDate";
import type { AuditLogRecord, AuditLogSource } from "@/types/audit";

import styles from "@/components/admin/admin.module.css";

type LogTab = "all" | "admin" | "customer" | "sharepoint" | "notification";

type LogsPayload = {
  logs: AuditLogRecord[];
  total?: number;
  fetchedAt?: string;
  configured: boolean;
  usingConsoleFallback: boolean;
  exportEnabled: boolean;
};

const TABS: Array<{ id: LogTab; label: string }> = [
  { id: "all", label: "All logs" },
  { id: "admin", label: "Admin" },
  { id: "customer", label: "Customer / TM" },
  { id: "sharepoint", label: "SharePoint" },
  { id: "notification", label: "Notifications" },
];

const PAGE_SIZE = 75;
const POLL_MS = 10_000;

const SOURCE_LABEL: Record<AuditLogSource, string> = {
  admin: "Admin",
  customer: "Customer",
  sharepoint: "SharePoint",
  notification: "Notification",
};

const SOURCE_TONE: Record<AuditLogSource, "ok" | "info" | "neutral" | "warn"> = {
  admin: "ok",
  customer: "info",
  sharepoint: "neutral",
  notification: "warn",
};

export function AdminLogsClient() {
  const { pushToast } = useAdminToast();
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [payload, setPayload] = useState<LogsPayload | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  const [action, setAction] = useState("");
  const [success, setSuccess] = useState<"all" | "true" | "false">("all");
  const [entityType, setEntityType] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [tab, setTab] = useState<LogTab>("all");
  const [page, setPage] = useState(1);
  const [selected, setSelected] = useState<AuditLogRecord | null>(null);
  const [live, setLive] = useState(true);
  const requestSeq = useRef(0);

  useEffect(() => {
    const timer = window.setTimeout(() => setDebouncedSearch(search), 300);
    return () => window.clearTimeout(timer);
  }, [search]);

  const load = useCallback(
    async (silent = false) => {
      const seq = ++requestSeq.current;
      if (silent) setRefreshing(true);
      else setLoading(true);
      try {
        const params = new URLSearchParams();
        params.set("top", "8000");
        if (debouncedSearch.trim()) params.set("search", debouncedSearch.trim());
        if (action) params.set("action", action);
        if (success !== "all") params.set("success", success);
        if (entityType.trim()) params.set("entityType", entityType.trim());
        if (tab !== "all") params.set("source", tab);
        if (from) params.set("from", from);
        if (to) params.set("to", to);

        const response = await fetch(`/api/admin/logs?${params.toString()}`, {
          cache: "no-store",
        });
        if (!response.ok) throw new Error(await readPublicApiError(response));
        const data = (await response.json()) as LogsPayload;
        if (seq !== requestSeq.current) return;
        setPayload(data);
        setLoadError(null);
      } catch (error) {
        if (seq !== requestSeq.current) return;
        const message =
          error instanceof Error ? error.message : "Failed to load logs";
        setLoadError(message);
        if (!silent) pushToast(message, "error");
      } finally {
        if (seq === requestSeq.current) {
          setLoading(false);
          setRefreshing(false);
        }
      }
    },
    [action, debouncedSearch, entityType, from, pushToast, success, tab, to],
  );

  useEffect(() => {
    setPage(1);
    void load(false);
  }, [load]);

  useEffect(() => {
    if (!live) return;
    const tick = () => {
      if (document.visibilityState !== "visible") return;
      void load(true);
    };
    const timer = window.setInterval(tick, POLL_MS);
    const onVisible = () => {
      if (document.visibilityState === "visible") void load(true);
    };
    document.addEventListener("visibilitychange", onVisible);
    return () => {
      window.clearInterval(timer);
      document.removeEventListener("visibilitychange", onVisible);
    };
  }, [live, load]);

  const entityTypes = useMemo(() => {
    const set = new Set<string>();
    for (const row of payload?.logs ?? []) {
      if (row.entityType) set.add(row.entityType);
    }
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [payload?.logs]);

  const actionOptions = useMemo(() => {
    const set = new Set<string>();
    if (action) set.add(action);
    for (const row of payload?.logs ?? []) {
      if (row.action) set.add(row.action);
    }
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [action, payload?.logs]);

  const logs = payload?.logs ?? [];
  const pageCount = Math.max(1, Math.ceil(logs.length / PAGE_SIZE));
  const safePage = Math.min(page, pageCount);
  const pageRows = logs.slice((safePage - 1) * PAGE_SIZE, safePage * PAGE_SIZE);

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
            End-to-end timeline of Admin, Customer, Training Manager, and
            SharePoint Training Manager Logs — newest first.
          </p>
        </div>
        <div className={styles.logHeaderActions}>
          <label className={styles.logLiveToggle}>
            <input
              type="checkbox"
              checked={live}
              onChange={(event) => setLive(event.target.checked)}
            />
            <span
              className={
                live ? styles.logLiveDotOn : styles.logLiveDotOff
              }
              aria-hidden="true"
            />
            Live
          </label>
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={() => void load(false)}
            disabled={loading}
          >
            {refreshing ? "Updating…" : "Refresh"}
          </button>
        </div>
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
          : `Reading every SharePoint Training Manager Logs row · ${logs.length} shown`}
        {payload?.fetchedAt ? ` · Updated ${formatDateTime(payload.fetchedAt)}` : ""}
        {" · "}
        Retention: no automatic deletion yet.
      </div>

      <div
        className={`${styles.permissionTabbar} ${styles.logTabbar}`}
        role="tablist"
        aria-label="Filter logs by source"
      >
        {TABS.map((item) => (
          <button
            key={item.id}
            type="button"
            role="tab"
            aria-selected={tab === item.id}
            className={
              tab === item.id
                ? `${styles.permissionTab} ${styles.permissionTabActive}`
                : styles.permissionTab
            }
            onClick={() => setTab(item.id)}
          >
            {item.label}
          </button>
        ))}
      </div>

      <div className={styles.crudToolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="Email, action, company, role"
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
            {actionOptions.map((value) => (
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

      {loading && !payload ? (
        <LoadingState label="Loading audit logs…" />
      ) : loadError && !payload ? (
        <div
          className={`${styles.settingsBanner} ${styles.settingsBannerWarn}`}
          role="alert"
        >
          Failed to load audit logs: {loadError}
        </div>
      ) : logs.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No log entries</h2>
          <p>No audit rows matched these filters.</p>
        </div>
      ) : (
        <>
          <div className={styles.tableWrap}>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th scope="col">When</th>
                  <th scope="col">Source</th>
                  <th scope="col">User</th>
                  <th scope="col">Action</th>
                  <th scope="col">Entity</th>
                  <th scope="col">Company</th>
                  <th scope="col">Result</th>
                  <th scope="col">Actions</th>
                </tr>
              </thead>
              <tbody>
                {pageRows.map((row) => (
                  <tr key={row.id}>
                    <td>{formatDateTime(row.timestamp)}</td>
                    <td>
                      <StatusBadge
                        label={SOURCE_LABEL[row.source] ?? row.source}
                        tone={SOURCE_TONE[row.source] ?? "neutral"}
                      />
                    </td>
                    <td>
                      <div>{row.userEmail}</div>
                      <div className={styles.logCellNote}>
                        {row.roleType ?? "—"}
                      </div>
                    </td>
                    <td>{row.action}</td>
                    <td>
                      <div>{row.entityType}</div>
                      <div className={styles.logCellNote}>
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
          {pageCount > 1 ? (
            <div className={styles.logPager}>
              <button
                type="button"
                className={styles.secondaryButton}
                disabled={safePage <= 1}
                onClick={() => setPage((current) => Math.max(1, current - 1))}
              >
                Previous
              </button>
              <span className={styles.logPagerStatus}>
                Page {safePage} of {pageCount} · {logs.length} rows
              </span>
              <button
                type="button"
                className={styles.secondaryButton}
                disabled={safePage >= pageCount}
                onClick={() =>
                  setPage((current) => Math.min(pageCount, current + 1))
                }
              >
                Next
              </button>
            </div>
          ) : (
            <p className={styles.logPagerStatus}>{logs.length} rows</p>
          )}
        </>
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
            <Detail
              label="Timestamp"
              value={formatDateTime(selected.timestamp)}
            />
            <Detail
              label="Source"
              value={SOURCE_LABEL[selected.source] ?? selected.source}
            />
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
