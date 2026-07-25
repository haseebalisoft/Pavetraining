"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminOfferRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminOfferRecord>[] = [
  { key: "title", header: "Offer title", render: (row) => row.title },
  { key: "company", header: "Company", render: (row) => row.company ?? "—" },
  { key: "start", header: "Start date", render: (row) => row.startDate ?? "—" },
  { key: "end", header: "End date", render: (row) => row.endDate ?? "—" },
  { key: "status", header: "Status", render: (row) => row.status ?? "—" },
  {
    key: "visible",
    header: "Customer visible",
    render: (row) => (row.customerVisible ? "Yes" : "No"),
  },
];

const fields: AdminFieldConfig[] = [
  { name: "title", label: "Offer title", type: "text", required: true },
  { name: "company", label: "Company", type: "company", required: true },
  { name: "description", label: "Description", type: "textarea" },
  { name: "startDate", label: "Start date", type: "date" },
  { name: "endDate", label: "End date", type: "date" },
  {
    name: "status",
    label: "Status",
    type: "select",
    options: [
      { value: "Active", label: "Active" },
      { value: "Inactive", label: "Inactive" },
    ],
  },
  { name: "customerVisible", label: "Customer visible", type: "boolean" },
];

export function AdminOffersClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminOfferRecord[];
}) {
  return (
    <AdminCrudPage<AdminOfferRecord>
      title="Offers"
      description="Manage promotions, active dates, and customer visibility."
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.company}
      listUrl="/api/admin/offers"
      createUrl="/api/admin/offers"
      updateUrl={(id) => `/api/admin/offers/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminOfferRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.title,
        (row) => row.company,
        (row) => row.description,
        (row) => row.status,
      ]}
    />
  );
}
