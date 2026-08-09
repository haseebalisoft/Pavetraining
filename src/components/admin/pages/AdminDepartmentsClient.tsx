"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminDepartmentRecord } from "@/lib/services/departmentTypes";
import { MAX_DEPARTMENTS_PER_COMPANY } from "@/lib/services/departmentTypes";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminDepartmentRecord>[] = [
  { key: "name", header: "Department", render: (row) => row.name },
  {
    key: "company",
    header: "Company",
    render: (row) => row.companyName?.trim() || "—",
  },
];

const fields: AdminFieldConfig[] = [
  {
    name: "companyId",
    label: "Company",
    type: "company",
    required: true,
    section: "Department",
  },
  {
    name: "name",
    label: "Department name",
    type: "text",
    required: true,
    section: "Department",
    placeholder: "e.g. Civils, Rail, Plant Hire, Gas",
  },
];

export function AdminDepartmentsClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminDepartmentRecord[];
}) {
  return (
    <AdminCrudPage<AdminDepartmentRecord>
      title="Departments"
      description={`For Enterprise companies: add up to ${MAX_DEPARTMENTS_PER_COMPANY} departments (Civils, Rail, Gas, etc.). Assign candidates to one department on Workforce, and give Training Managers coverage of one or more departments under Permissions. Changing department or TM never deletes training history. Use Admin → Companies filter here, or open Workforce/Permissions to assign.`}
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      listUrl="/api/admin/departments"
      createUrl="/api/admin/departments"
      updateUrl={(id) => `/api/admin/departments/${id}`}
      deleteUrl={(id) => `/api/admin/departments/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminDepartmentRecord[] }).records ?? [])
      }
      searchKeys={[(row) => row.name, (row) => row.companyName]}
    />
  );
}
