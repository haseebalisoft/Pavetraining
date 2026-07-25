"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminEventRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminEventRecord>[] = [
  { key: "title", header: "Event title", render: (row) => row.title },
  { key: "company", header: "Company", render: (row) => row.company ?? "—" },
  { key: "date", header: "Date", render: (row) => row.eventDate ?? "—" },
  {
    key: "location",
    header: "Location",
    render: (row) => row.trainingAddress ?? row.location ?? "—",
  },
  {
    key: "visible",
    header: "Customer visible",
    render: (row) => (row.customerVisible ? "Yes" : "No"),
  },
];

const fields: AdminFieldConfig[] = [
  { name: "title", label: "Event title", type: "text", required: true },
  { name: "company", label: "Company", type: "company", required: true },
  { name: "eventDate", label: "Event date", type: "date" },
  { name: "endDate", label: "End date", type: "date" },
  { name: "trainingAddress", label: "Training address", type: "text" },
  { name: "location", label: "Location", type: "text" },
  { name: "description", label: "Description", type: "textarea" },
  { name: "customerVisible", label: "Customer visible", type: "boolean" },
];

export function AdminEventsClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminEventRecord[];
}) {
  return (
    <AdminCrudPage<AdminEventRecord>
      title="Events"
      description="Create and edit training events linked to companies."
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.company}
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
      ]}
    />
  );
}
