"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import { useAdminToast } from "@/components/admin/AdminToast";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import {
  DEFAULT_DEPARTMENT_STATUS,
  MAX_DEPARTMENTS_PER_COMPANY,
  type AdminDepartmentRecord,
} from "@/lib/services/departmentTypes";
import type { Company } from "@/types/models";

import styles from "@/components/admin/admin.module.css";

function DeactivateDepartmentButton({
  row,
  reload,
}: {
  row: AdminDepartmentRecord;
  reload: () => Promise<void>;
}) {
  const { pushToast } = useAdminToast();

  if (row.status !== "Active") return null;

  async function onDeactivate() {
    if (
      !window.confirm(
        `Deactivate "${row.name}"? It will stop appearing as a choice for new Workforce / Permissions assignments, but existing history is kept.`,
      )
    ) {
      return;
    }
    try {
      const response = await fetch(`/api/admin/departments/${row.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ status: "Inactive" }),
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      pushToast("Department deactivated.", "success");
      await reload();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to deactivate",
        "error",
      );
    }
  }

  return (
    <button
      type="button"
      className={styles.linkButton}
      onClick={() => void onDeactivate()}
    >
      Deactivate
    </button>
  );
}

const columns: AdminColumn<AdminDepartmentRecord>[] = [
  { key: "name", header: "Department", render: (row) => row.name },
  {
    key: "company",
    header: "Company",
    render: (row) => row.companyName?.trim() || "—",
  },
  {
    key: "status",
    header: "Status",
    render: (row) => row.status,
  },
  {
    key: "notes",
    header: "Notes",
    render: (row) => row.notes?.trim() || "—",
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
  {
    name: "status",
    label: "Status",
    type: "select",
    required: true,
    options: [
      { value: "Active", label: "Active" },
      { value: "Inactive", label: "Inactive" },
    ],
    section: "Department",
  },
  {
    name: "notes",
    label: "Notes",
    type: "textarea",
    section: "Department",
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
      description={`For Enterprise companies: add up to ${MAX_DEPARTMENTS_PER_COMPANY} departments (Civils, Rail, Gas, etc.). Assign candidates to one department on Workforce, and give Training Managers coverage of one or more departments under Permissions. Changing department or TM never deletes training history. Deactivate a department instead of deleting it to hide it from new assignments while keeping history — use Delete only when it was created by mistake.`}
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
      getCreateDefaults={() => ({ status: DEFAULT_DEPARTMENT_STATUS })}
      mapResponse={(payload) =>
        ((payload as { records?: AdminDepartmentRecord[] }).records ?? [])
      }
      searchKeys={[(row) => row.name, (row) => row.companyName]}
      extraActions={(row, { reload }) => (
        <DeactivateDepartmentButton row={row} reload={reload} />
      )}
    />
  );
}
