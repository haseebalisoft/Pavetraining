"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminOfferRecord } from "@/lib/services/adminCrudService";
import { formatDate } from "@/lib/utils/formatDate";

const columns: AdminColumn<AdminOfferRecord>[] = [
  { key: "title", header: "Offer title", render: (row) => row.title },
  {
    key: "category",
    header: "Category",
    render: (row) => row.category ?? "—",
  },
  {
    key: "start",
    header: "Start date",
    render: (row) => formatDate(row.startDate),
  },
  {
    key: "end",
    header: "End date",
    render: (row) => formatDate(row.endDate),
  },
  { key: "status", header: "Status", render: (row) => row.status ?? "—" },
  {
    key: "visible",
    header: "Customer visible",
    render: (row) => (row.customerVisible ? "Yes" : "No"),
  },
];

const fields: AdminFieldConfig[] = [
  { name: "title", label: "Offer title", type: "text", required: true },
  {
    name: "category",
    label: "Category",
    type: "select",
    options: [
      { value: "NPORS", label: "NPORS" },
      { value: "EUSR", label: "EUSR" },
      { value: "Streetworks", label: "Streetworks" },
      { value: "NVQ", label: "NVQ" },
      { value: "In-House", label: "In-House" },
      { value: "M4S", label: "M4S" },
      { value: "General", label: "General" },
    ],
  },
  { name: "description", label: "Short description", type: "textarea" },
  { name: "startDate", label: "Start date", type: "date" },
  { name: "endDate", label: "End date", type: "date" },
  {
    name: "status",
    label: "Status",
    type: "select",
    options: [
      { value: "Active", label: "Active" },
      { value: "Inactive", label: "Inactive" },
      { value: "Expired", label: "Expired" },
    ],
  },
  { name: "customerVisible", label: "Customer visible", type: "boolean" },
];

export function AdminOffersClient({
  initialRows,
}: {
  initialRows: AdminOfferRecord[];
}) {
  return (
    <AdminCrudPage<AdminOfferRecord>
      title="Offers"
      description="Manage promotions, categories, active dates, and customer visibility. Offers are site-wide (not company-specific)."
      columns={columns}
      fields={fields}
      initialRows={initialRows}
      emptyLabel="No offers found. Create an offer to begin."
      listUrl="/api/admin/offers"
      createUrl="/api/admin/offers"
      updateUrl={(id) => `/api/admin/offers/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminOfferRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.title,
        (row) => row.category,
        (row) => row.description,
        (row) => row.status,
      ]}
    />
  );
}
