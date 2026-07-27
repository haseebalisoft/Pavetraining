"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminPermissionRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminPermissionRecord>[] = [
  { key: "email", header: "User email", render: (row) => row.userEmail },
  {
    key: "name",
    header: "Name",
    render: (row) => row.name?.trim() || "—",
  },
  { key: "role", header: "Role", render: (row) => row.roleType },
  { key: "company", header: "Company", render: (row) => row.companyName ?? "—" },
  { key: "status", header: "Status", render: (row) => row.status },
  {
    key: "flags",
    header: "Access",
    render: (row) =>
      [
        row.canView ? "View" : null,
        row.canDownload ? "Download" : null,
        row.canEdit ? "Edit" : null,
      ]
        .filter(Boolean)
        .join(" · ") || "—",
  },
];

const fields: AdminFieldConfig[] = [
  { name: "userEmail", label: "User email", type: "email", required: true },
  {
    name: "name",
    label: "Name",
    type: "text",
    required: false,
  },
  {
    name: "roleType",
    label: "Role",
    type: "select",
    required: true,
    options: [
      {
        value: "Admin",
        label: "Training Manager (Admin portal)",
      },
      {
        value: "Customer",
        label: "Supervisor (Customer portal)",
      },
    ],
  },
  { name: "companyId", label: "Company", type: "company", required: true },
  {
    name: "status",
    label: "Status",
    type: "select",
    required: true,
    options: [
      { value: "Active", label: "Active" },
      { value: "Inactive", label: "Inactive" },
    ],
  },
  { name: "canView", label: "Can view", type: "boolean" },
  { name: "canDownload", label: "Can download", type: "boolean" },
  { name: "canEdit", label: "Can edit", type: "boolean" },
];

export function AdminPermissionsClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminPermissionRecord[];
}) {
  return (
    <AdminCrudPage<AdminPermissionRecord>
      title="Permissions"
      description="Assign portal roles, companies, and access flags. Name is used by Workforce Training manager / Supervisor lookups."
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      listUrl="/api/admin/permissions"
      createUrl="/api/admin/permissions"
      updateUrl={(id) => `/api/admin/permissions/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminPermissionRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.userEmail,
        (row) => row.roleType,
        (row) => row.companyName,
        (row) => row.status,
      ]}
    />
  );
}
