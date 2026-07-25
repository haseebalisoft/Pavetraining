"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
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
    key: "trainingAddress",
    header: "Training Address",
    render: (row) => row.trainingAddress ?? "—",
  },
  {
    key: "visible",
    header: "Customer Visible",
    render: (row) => (row.customerVisible ? "Yes" : "No"),
  },
  {
    key: "syncStatus",
    header: "Sync Status",
    render: (row) => row.syncStatus ?? "—",
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
      ]}
    />
  );
}
