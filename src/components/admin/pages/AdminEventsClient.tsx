"use client";

import { useCallback, useMemo, useState } from "react";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import { AdminDrawer } from "@/components/admin/AdminDrawer";
import styles from "@/components/admin/admin.module.css";
import { useAdminToast } from "@/components/admin/AdminToast";
import {
  CalendarGrid,
  type CalendarGridEvent,
  type CalendarView,
} from "@/components/calendar/CalendarGrid";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { AdminEventRecord } from "@/lib/services/adminCrudService";
import { formatDateTime } from "@/lib/utils/formatDate";
import type { Company } from "@/types/models";

type EventsViewMode = "list" | CalendarView;

const EMPTY_EVENT_WARNINGS: string[] = [];

interface EventFormState {
  title: string;
  companyId: string;
  eventDate: string;
  endDate: string;
  location: string;
  trainingAddress: string;
  description: string;
  internalNotes: string;
  customerVisible: boolean;
  bookingStatus: "Tentative" | "Confirmed";
  doNotSync: boolean;
}

function syncTone(
  status: string | null | undefined,
): "ok" | "warn" | "danger" | "neutral" {
  const normalized = (status ?? "").trim().toLowerCase();
  if (normalized === "synced") return "ok";
  if (normalized === "pending") return "warn";
  if (normalized === "failed") return "danger";
  return "neutral";
}

function toDateTimeLocal(value: string | Date | null | undefined): string {
  if (!value) return "";
  const parsed = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  const pad = (part: number) => String(part).padStart(2, "0");
  return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}T${pad(parsed.getHours())}:${pad(parsed.getMinutes())}`;
}

function createForm(date = new Date()): EventFormState {
  const start = new Date(date);
  start.setSeconds(0, 0);
  const end = new Date(start.getTime() + 60 * 60 * 1000);
  return {
    title: "",
    companyId: "",
    eventDate: toDateTimeLocal(start),
    endDate: toDateTimeLocal(end),
    location: "",
    trainingAddress: "",
    description: "",
    internalNotes: "",
    customerVisible: true,
    bookingStatus: "Tentative",
    doNotSync: false,
  };
}

function editForm(row: AdminEventRecord): EventFormState {
  return {
    title: row.title,
    companyId: row.companyId ?? "",
    eventDate: toDateTimeLocal(row.eventDate),
    endDate: toDateTimeLocal(row.endDate),
    location: row.location ?? "",
    trainingAddress: row.trainingAddress ?? "",
    description: row.description ?? "",
    internalNotes: row.internalNotes ?? "",
    customerVisible: row.customerVisible,
    bookingStatus: row.bookingStatus ?? "Tentative",
    doNotSync: row.doNotSync,
  };
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
  { key: "location", header: "Location", render: (row) => row.location ?? "—" },
  {
    key: "bookingStatus",
    header: "Booking",
    render: (row) => (
      <StatusBadge
        label={row.bookingStatus === "Confirmed" ? "Confirmed" : "Tentative"}
        tone={row.bookingStatus === "Confirmed" ? "ok" : "warn"}
      />
    ),
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
    key: "syncError",
    header: "Sync Error",
    render: (row) => row.syncError?.trim() || "—",
  },
  {
    key: "lastSynced",
    header: "Last Synced At",
    render: (row) => formatDateTime(row.lastSyncedAt),
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
  { name: "companyId", label: "Company", type: "company", required: true },
  {
    name: "eventDate",
    label: "Start Date/Time",
    type: "datetime",
    required: true,
  },
  { name: "endDate", label: "End Date/Time", type: "datetime" },
  { name: "location", label: "Location", type: "text" },
  { name: "trainingAddress", label: "Training Address", type: "text" },
  {
    name: "description",
    label: "Description customers will see",
    type: "textarea",
  },
  {
    name: "internalNotes",
    label: "Internal notes — never shown to customers",
    type: "textarea",
  },
  {
    name: "bookingStatus",
    label: "Booking status",
    type: "select",
    options: [
      { value: "Tentative", label: "Tentative (offered dates)" },
      { value: "Confirmed", label: "Confirmed (busy + email TMs)" },
    ],
    section: "Booking",
  },
  { name: "customerVisible", label: "Customer Visible", type: "boolean" },
  {
    name: "doNotSync",
    label: "Do Not Sync",
    type: "boolean",
    section: "Outlook sync",
  },
  { name: "syncStatus", label: "Sync Status", type: "text", readOnly: true },
  {
    name: "lastSyncedAt",
    label: "Last Synced At",
    type: "text",
    readOnly: true,
  },
  { name: "syncError", label: "Sync Error", type: "textarea", readOnly: true },
];

function ViewButtons({
  view,
  onChange,
}: {
  view: EventsViewMode;
  onChange: (view: EventsViewMode) => void;
}) {
  return (
    <div className={styles.syncToolbarActions} role="group" aria-label="Events view">
      {(["list", "month", "week"] as const).map((option) => (
        <button
          key={option}
          type="button"
          className={view === option ? styles.primaryButton : styles.secondaryButton}
          aria-pressed={view === option}
          onClick={() => onChange(option)}
        >
          {option[0].toUpperCase() + option.slice(1)}
        </button>
      ))}
    </div>
  );
}

export function AdminEventsClient({
  companies,
  initialRows,
  warnings = EMPTY_EVENT_WARNINGS,
}: {
  companies: Company[];
  initialRows: AdminEventRecord[];
  warnings?: string[];
}) {
  const { pushToast } = useAdminToast();
  const [view, setView] = useState<EventsViewMode>("month");
  const [rows, setRows] = useState(initialRows);
  const [cursor, setCursor] = useState(() => new Date());
  const [busyId, setBusyId] = useState<string | null>(null);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editing, setEditing] = useState<AdminEventRecord | null>(null);
  const [form, setForm] = useState<EventFormState>(() => createForm());
  const [formError, setFormError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);

  const handleRowsChange = useCallback((nextRows: AdminEventRecord[]) => {
    setRows(nextRows);
  }, []);

  const calendarEvents = useMemo<CalendarGridEvent[]>(
    () =>
      rows.map((row) => ({
        id: row.id,
        title: row.title,
        start: row.eventDate,
        end: row.endDate,
        company: row.company,
        location: row.location,
      })),
    [rows],
  );

  const reload = useCallback(async () => {
    const response = await fetch("/api/admin/events", { cache: "no-store" });
    if (!response.ok) throw new Error(await readPublicApiError(response));
    const payload = (await response.json()) as { records?: AdminEventRecord[] };
    setRows(payload.records ?? []);
  }, []);

  const runSync = useCallback(
    async (row: AdminEventRecord, mode: "sync" | "retry") => {
      const path =
        mode === "retry"
          ? `/api/admin/events/${row.id}/retry-sync`
          : `/api/admin/events/${row.id}/sync`;
      setBusyId(row.id);
      try {
        const response = await fetch(path, { method: "POST" });
        if (!response.ok) throw new Error(await readPublicApiError(response));
        const payload = (await response.json()) as {
          result?: { status?: string; error?: string | null; reason?: string | null };
        };
        const status = payload.result?.status ?? "Unknown";
        if (status === "Failed") {
          pushToast(payload.result?.error || `Sync failed for “${row.title}”.`, "error");
        } else if (status === "Skipped") {
          pushToast(payload.result?.reason || `Sync skipped for “${row.title}”.`);
        } else {
          pushToast(`Outlook sync ${status.toLowerCase()} for “${row.title}”.`);
        }
        await reload();
      } catch (error) {
        pushToast(error instanceof Error ? error.message : "Sync request failed.", "error");
      } finally {
        setBusyId(null);
      }
    },
    [pushToast, reload],
  );

  function openCreate(date = new Date()) {
    setEditing(null);
    setForm(createForm(date));
    setFormError(null);
    setDrawerOpen(true);
  }

  function openEdit(row: AdminEventRecord) {
    setEditing(row);
    setForm(editForm(row));
    setFormError(null);
    setDrawerOpen(true);
  }

  async function saveEvent() {
    if (!form.title.trim()) return setFormError("Title is required.");
    if (!form.companyId) return setFormError("Company is required.");
    if (!form.eventDate) return setFormError("Start date/time is required.");
    if (form.endDate && new Date(form.endDate) < new Date(form.eventDate)) {
      return setFormError("End date/time must be after the start.");
    }

    setSaving(true);
    setFormError(null);
    try {
      const response = await fetch(
        editing ? `/api/admin/events/${editing.id}` : "/api/admin/events",
        {
          method: editing ? "PATCH" : "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(form),
        },
      );
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const payload = (await response.json()) as {
        record?: AdminEventRecord;
        bookingNotification?: {
          attempted?: boolean;
          skipped?: boolean;
          skipReason?: string;
          recipients?: string[];
          results?: Array<{ status: string; recipientEmail: string }>;
        } | null;
      };

      const notify = payload.bookingNotification;
      const sentTo =
        notify?.results
          ?.filter((r) => r.status === "sent")
          .map((r) => r.recipientEmail) ?? [];

      if (form.bookingStatus === "Confirmed" && notify) {
        if (sentTo.length > 0) {
          pushToast(
            `${editing ? "Event updated" : "Event created"}. Confirmation emailed to ${sentTo.join(", ")} (calendar invite attached).`,
            "success",
          );
        } else if (notify.skipped || notify.attempted) {
          pushToast(
            `${editing ? "Event updated" : "Event created"}. Email not sent: ${notify.skipReason || "No Training Manager recipients."}`,
            "error",
          );
        } else {
          pushToast(editing ? "Event updated." : "Event created.", "success");
        }
      } else {
        pushToast(
          editing
            ? "Event updated."
            : "Event created. Set status to Confirmed to email Training Managers.",
          "success",
        );
      }

      setDrawerOpen(false);
      await reload();
    } catch (error) {
      const message = error instanceof Error ? error.message : "Failed to save event.";
      setFormError(message);
      pushToast(message, "error");
    } finally {
      setSaving(false);
    }
  }

  async function deleteEvent() {
    if (!editing) return;
    const confirmed = window.confirm(
      `Delete event “${editing.title}”?\n\nThis removes it from the portal and tries to remove it from Outlook. This cannot be undone.`,
    );
    if (!confirmed) return;

    setDeleting(true);
    setFormError(null);
    try {
      const response = await fetch(`/api/admin/events/${editing.id}`, {
        method: "DELETE",
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      pushToast("Event deleted.", "success");
      setDrawerOpen(false);
      setEditing(null);
      await reload();
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Failed to delete event.";
      setFormError(message);
      pushToast(message, "error");
    } finally {
      setDeleting(false);
    }
  }

  const viewButtons = <ViewButtons view={view} onChange={setView} />;

  if (view === "list") {
    return (
      <AdminCrudPage<AdminEventRecord>
        title="Events"
        description="Manage training events. SharePoint Events is the source of truth and customer-facing copy is kept separate from internal notes."
        columns={columns}
        fields={fields}
        companies={companies}
        initialRows={rows}
        warnings={warnings}
        enableCompanyFilter
        getCompanyName={(row) => row.company}
        drawerWide
        emptyLabel="No events found. Create an event to begin."
        listUrl="/api/admin/events"
        createUrl="/api/admin/events"
        updateUrl={(id) => `/api/admin/events/${id}`}
        deleteUrl={(id) => `/api/admin/events/${id}`}
        deleteConfirmExtra="This also tries to remove the event from Outlook."
        mapResponse={(payload) =>
          ((payload as { records?: AdminEventRecord[] }).records ?? [])
        }
        onRowsChange={handleRowsChange}
        toolbarExtra={viewButtons}
        searchKeys={[
          (row) => row.title,
          (row) => row.company,
          (row) => row.trainingAddress,
          (row) => row.location,
          (row) => row.description,
          (row) => row.internalNotes,
          (row) => row.syncStatus,
          (row) => row.syncError,
        ]}
        extraActions={(row, { reload: reloadList }) => (
          <>
            <button
              type="button"
              className={styles.linkButton}
              disabled={busyId === row.id || row.doNotSync}
              onClick={() => void runSync(row, "sync").then(reloadList)}
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
                  onClick={() => void runSync(row, "retry").then(reloadList)}
                >
                  Retry sync
                </button>
              </>
            ) : null}
          </>
        )}
      />
    );
  }

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs items={[{ label: "Admin", href: "/admin" }, { label: "Events" }]} />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>Events</h1>
          <p className={styles.subtitle}>
            Outlook-style calendar. Click an empty slot to create an event or an event to edit it.
          </p>
        </div>
        <button type="button" className={styles.primaryButton} onClick={() => openCreate()}>
          Add event
        </button>
      </header>

      {warnings.length ? (
        <div className={styles.schemaWarnings} role="alert">
          {warnings.map((warning) => <p key={warning}>{warning}</p>)}
        </div>
      ) : null}

      <div className={styles.crudToolbar}>{viewButtons}</div>
      <CalendarGrid
        events={calendarEvents}
        view={view}
        cursor={cursor}
        onCursorChange={setCursor}
        onSlotClick={openCreate}
        onEventClick={(event) => {
          const row = rows.find((candidate) => candidate.id === event.id);
          if (row) openEdit(row);
        }}
        emptyLabel="No events in this period. Click an empty slot to create one."
      />

      <AdminDrawer
        open={drawerOpen}
        title={editing ? `Edit ${editing.title}` : "Add event"}
        onClose={() => setDrawerOpen(false)}
        wide
        footer={
          <>
            {editing ? (
              <button
                type="button"
                className={styles.linkButtonDanger}
                onClick={() => void deleteEvent()}
                disabled={saving || deleting}
              >
                {deleting ? "Deleting…" : "Delete"}
              </button>
            ) : null}
            <button type="button" className={styles.secondaryButton} onClick={() => setDrawerOpen(false)} disabled={saving || deleting}>
              Cancel
            </button>
            <button type="button" className={styles.primaryButton} onClick={() => void saveEvent()} disabled={saving || deleting}>
              {saving ? "Saving…" : editing ? "Save changes" : "Create event"}
            </button>
          </>
        }
      >
        {formError ? <p className={styles.formError}>{formError}</p> : null}
        <div className={styles.formGrid}>
          <label className={`${styles.field} ${styles.fieldFull}`}>
            <span className={styles.fieldLabel}>Title</span>
            <input className={styles.input} value={form.title} onChange={(event) => setForm((current) => ({ ...current, title: event.target.value }))} />
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Start date/time</span>
            <input className={styles.input} type="datetime-local" value={form.eventDate} onChange={(event) => setForm((current) => ({ ...current, eventDate: event.target.value }))} />
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>End date/time</span>
            <input className={styles.input} type="datetime-local" value={form.endDate} onChange={(event) => setForm((current) => ({ ...current, endDate: event.target.value }))} />
          </label>
          <label className={`${styles.field} ${styles.fieldFull}`}>
            <span className={styles.fieldLabel}>Company</span>
            <select className={styles.select} value={form.companyId} onChange={(event) => setForm((current) => ({ ...current, companyId: event.target.value }))}>
              <option value="">Select company…</option>
              {companies.map((company) => <option key={company.id} value={company.id}>{company.companyName}</option>)}
            </select>
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Location</span>
            <input className={styles.input} value={form.location} onChange={(event) => setForm((current) => ({ ...current, location: event.target.value }))} />
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Training address</span>
            <input className={styles.input} value={form.trainingAddress} onChange={(event) => setForm((current) => ({ ...current, trainingAddress: event.target.value }))} />
          </label>
          <label className={`${styles.field} ${styles.fieldFull}`}>
            <span className={styles.fieldLabel}>Description customers will see</span>
            <textarea className={styles.input} rows={5} value={form.description} onChange={(event) => setForm((current) => ({ ...current, description: event.target.value }))} />
          </label>
          <label className={`${styles.field} ${styles.fieldFull}`}>
            <span className={styles.fieldLabel}>Internal notes — never shown to customers</span>
            <textarea className={styles.input} rows={5} value={form.internalNotes} onChange={(event) => setForm((current) => ({ ...current, internalNotes: event.target.value }))} />
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Booking status</span>
            <select
              className={styles.select}
              value={form.bookingStatus}
              onChange={(event) =>
                setForm((current) => ({
                  ...current,
                  bookingStatus: event.target.value as "Tentative" | "Confirmed",
                }))
              }
            >
              <option value="Tentative">Tentative (offered dates)</option>
              <option value="Confirmed">Confirmed (busy + email TMs)</option>
            </select>
          </label>
          <label className={styles.checkboxRow}>
            <input type="checkbox" checked={form.customerVisible} onChange={(event) => setForm((current) => ({ ...current, customerVisible: event.target.checked }))} />
            Customer visible
          </label>
          <label className={styles.checkboxRow}>
            <input type="checkbox" checked={form.doNotSync} onChange={(event) => setForm((current) => ({ ...current, doNotSync: event.target.checked }))} />
            Do not sync to Outlook
          </label>
          {editing ? (
            <div className={`${styles.fieldFull} ${styles.settingsActions}`}>
              <StatusBadge label={editing.syncStatus?.trim() || "Not synced"} tone={syncTone(editing.syncStatus)} />
              <button type="button" className={styles.secondaryButton} disabled={busyId === editing.id || editing.doNotSync} onClick={() => void runSync(editing, editing.syncStatus?.toLowerCase() === "failed" ? "retry" : "sync")}>
                {busyId === editing.id ? "Syncing…" : "Sync now"}
              </button>
              {editing.syncError?.trim() ? (
                <p className={styles.helpText}>Sync error: {editing.syncError}</p>
              ) : null}
            </div>
          ) : null}
        </div>
      </AdminDrawer>
    </div>
  );
}
