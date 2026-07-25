"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminWorkforceRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminWorkforceRecord>[] = [
  { key: "name", header: "Candidate name", render: (row) => row.candidateName },
  { key: "company", header: "Company", render: (row) => row.companyName },
  {
    key: "workforce",
    header: "Workforce number",
    render: (row) => row.workforceNumber ?? "—",
  },
  { key: "department", header: "Department", render: (row) => row.department ?? "—" },
  { key: "dob", header: "Date of birth", render: (row) => row.dateOfBirth ?? "—" },
  { key: "status", header: "Status", render: (row) => row.status ?? "—" },
];

const fields: AdminFieldConfig[] = [
  { name: "candidateName", label: "Candidate name", type: "text", required: true },
  { name: "companyName", label: "Company", type: "company", required: true },
  { name: "workforceNumber", label: "Workforce number", type: "text" },
  { name: "department", label: "Department", type: "text" },
  { name: "dateOfBirth", label: "Date of birth", type: "date" },
  {
    name: "status",
    label: "Status",
    type: "select",
    options: [
      { value: "Active", label: "Active" },
      { value: "Inactive", label: "Inactive" },
    ],
  },
];

export function AdminWorkforceClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminWorkforceRecord[];
}) {
  return (
    <AdminCrudPage<AdminWorkforceRecord>
      title="Workforce"
      description="Manage candidates and link them to companies."
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      listUrl="/api/admin/workforce"
      createUrl="/api/admin/workforce"
      updateUrl={(id) => `/api/admin/workforce/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminWorkforceRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.candidateName,
        (row) => row.companyName,
        (row) => row.workforceNumber,
        (row) => row.department,
      ]}
    />
  );
}
