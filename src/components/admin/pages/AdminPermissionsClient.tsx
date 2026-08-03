"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminPermissionRecord } from "@/lib/services/adminCrudService";
import type { AdminDepartmentRecord } from "@/lib/services/departmentService";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminPermissionRecord>[] = [
  { key: "email", header: "User email", render: (row) => row.userEmail },
  {
    key: "name",
    header: "Name",
    render: (row) => row.name?.trim() || "—",
  },
  {
    key: "role",
    header: "Role",
    render: (row) => row.roleLabel,
  },
  {
    key: "scope",
    header: "Access scope",
    render: (row) => row.accessScope?.trim() || "—",
  },
  {
    key: "depts",
    header: "Departments covered",
    render: (row) =>
      row.departmentScopes?.length ? row.departmentScopes.join(", ") : "—",
  },
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
    name: "permissionRole",
    label: "Role",
    type: "select",
    required: true,
    options: [
      {
        value: "Admin",
        label: "Training Manager (Admin + Training Matrix)",
      },
      {
        value: "Customer",
        label: "Supervisor (Customer Training Matrix)",
      },
      {
        value: "Candidate",
        label: "Candidate (own Training Matrix only)",
      },
    ],
  },
  {
    name: "accessScope",
    label: "Access scope",
    type: "select",
    required: true,
    options: [
      { value: "Full Company", label: "Full Company" },
      {
        value: "Department Only",
        label: "Department Only (recommended for Enterprise TMs)",
      },
      {
        value: "Candidate Only",
        label: "Candidate Only",
      },
    ],
  },
  { name: "companyId", label: "Company", type: "company", required: true },
  {
    name: "departmentsAllowed",
    label: "Departments covered",
    type: "multiselect",
    section: "Department coverage",
    companyScopedDepartments: true,
    departmentValueMode: "id",
  },
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
  departments,
  initialRows,
}: {
  companies: Company[];
  departments: AdminDepartmentRecord[];
  initialRows: AdminPermissionRecord[];
}) {
  return (
    <AdminCrudPage<AdminPermissionRecord>
      title="Permissions"
      description="Assign Training Manager, Supervisor, or Candidate. For Enterprise: set Access scope to Department Only and tick the departments they cover — they only see those candidates in the customer portal. Changing coverage or a candidate’s department never deletes training history."
      columns={columns}
      fields={fields}
      companies={companies}
      departments={departments}
      initialRows={initialRows}
      listUrl="/api/admin/permissions"
      createUrl="/api/admin/permissions"
      updateUrl={(id) => `/api/admin/permissions/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminPermissionRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.userEmail,
        (row) => row.name,
        (row) => row.companyName,
        (row) => row.roleLabel,
        (row) => row.departmentScopes?.join(" "),
      ]}
    />
  );
}
