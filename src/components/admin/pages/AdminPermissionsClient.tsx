"use client";

import { useMemo, useState } from "react";
import { usePathname, useRouter, useSearchParams } from "next/navigation";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminPermissionRecord } from "@/lib/services/adminCrudService";
import type { AdminDepartmentRecord } from "@/lib/services/departmentTypes";
import type { Company } from "@/types/models";

import styles from "@/components/admin/admin.module.css";

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
        label: "Customer",
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
      { value: "Full Company", label: "Full Company (all candidates in company)" },
      {
        value: "Department Only",
        label: "Department Only (Enterprise TMs — tick departments below)",
      },
      {
        value: "Candidate Only",
        label: "Candidate Only (own records)",
      },
    ],
  },
  { name: "companyId", label: "Company", type: "company", required: true },
  {
    name: "departmentsAllowed",
    label: "Departments covered",
    type: "multiselect",
    section: "Department coverage",
    // Show every active department across every company (labelled with the
    // company name so admins can tell duplicate names apart). The runtime
    // department filter is still combined with the Permission's Company
    // scope in customerAccessService, so nothing about the security model
    // changes — this is a UX-only widening of the picker.
    companyScopedDepartments: "all",
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

function mapPermissionsResponse(payload: unknown): AdminPermissionRecord[] {
  return ((payload as { records?: AdminPermissionRecord[] }).records ?? []);
}

const permissionSearchKeys: Array<
  (row: AdminPermissionRecord) => string | null | undefined
> = [
  (row) => row.userEmail,
  (row) => row.name,
  (row) => row.companyName,
  (row) => row.roleLabel,
  (row) => row.departmentScopes?.join(" "),
];

type PermissionTabId =
  | "all"
  | "Admin"
  | "Training Manager"
  | "Supervisor"
  | "Candidate";

const TABS: Array<{ id: PermissionTabId; label: string }> = [
  { id: "all", label: "All" },
  { id: "Admin", label: "Admins" },
  { id: "Training Manager", label: "Training Managers" },
  { id: "Supervisor", label: "Customers" },
  { id: "Candidate", label: "Candidates" },
];

function normalizeSpRole(value: string | null | undefined): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    // Fold legacy "TrainingManager" and "trainingmanager" into the same key
    .replace("trainingmanager", "training manager");
}

/** Map the URL ?role= param to a permissionRole form-field default. */
function rolePresetFromQuery(role: string | null): {
  permissionRole: string;
} | null {
  if (!role) return null;
  const key = role.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "trainingmanager" || key === "admin") {
    return { permissionRole: "Admin" };
  }
  if (key === "supervisor" || key === "customer") {
    return { permissionRole: "Customer" };
  }
  if (key === "candidate") {
    return { permissionRole: "Candidate" };
  }
  return null;
}

export function AdminPermissionsClient({
  companies,
  departments,
  initialRows,
}: {
  companies: Company[];
  departments: AdminDepartmentRecord[];
  initialRows: AdminPermissionRecord[];
}) {
  const router = useRouter();
  const pathname = usePathname();
  const searchParams = useSearchParams();
  const [activeTab, setActiveTab] = useState<PermissionTabId>("all");

  const rolePreset = useMemo(
    () => rolePresetFromQuery(searchParams.get("role")),
    [searchParams],
  );

  // Lock the "Role" field when the drawer was opened via an "Add Training
  // Manager" / "Add Customer" button so admins can't accidentally save the
  // wrong role from a role-preset shortcut. The lock only applies during the
  // shortcut open; picking Edit on any existing row keeps the field editable.
  const fieldsForRender = useMemo<AdminFieldConfig[]>(() => {
    if (!rolePreset) return fields;
    return fields.map((field) =>
      field.name === "permissionRole" ? { ...field, readOnly: true } : field,
    );
  }, [rolePreset]);

  const rowFilter = useMemo(() => {
    if (activeTab === "all") return undefined;
    const target = normalizeSpRole(activeTab);
    return (row: AdminPermissionRecord) => {
      const spRole = normalizeSpRole(row.sharePointRoleType);
      return spRole === target;
    };
  }, [activeTab]);

  function openAddWithRole(role: PermissionTabId) {
    // Force a tab switch so the new row shows up in the current view after
    // save, and push ?action=add&role=<value> so AdminCrudPage's existing
    // query-param handler opens the create drawer for us.
    setActiveTab(role);
    const params = new URLSearchParams(
      Array.from(searchParams.entries()).filter(([key]) => key !== "action" && key !== "role"),
    );
    params.set("action", "add");
    params.set("role", role.replace(/\s+/g, ""));
    router.replace(`${pathname}?${params.toString()}`);
  }

  const toolbarExtra = (
    <div className={styles.permissionTabbar} role="tablist" aria-label="Filter permissions by role">
      {TABS.map((tab) => (
        <button
          key={tab.id}
          type="button"
          role="tab"
          aria-selected={activeTab === tab.id}
          className={
            activeTab === tab.id
              ? `${styles.permissionTab} ${styles.permissionTabActive}`
              : styles.permissionTab
          }
          onClick={() => setActiveTab(tab.id)}
        >
          {tab.label}
        </button>
      ))}
      <span className={styles.permissionTabDivider} aria-hidden="true" />
      <button
        type="button"
        className={styles.permissionAddButton}
        onClick={() => openAddWithRole("Training Manager")}
      >
        + Add Training Manager
      </button>
      <button
        type="button"
        className={styles.permissionAddButton}
        onClick={() => openAddWithRole("Supervisor")}
      >
        + Add Customer
      </button>
    </div>
  );

  return (
    <AdminCrudPage<AdminPermissionRecord>
      title="Permissions"
      description="Fully editable by admin: Edit or Delete any row. For Enterprise TMs, set Access scope to Department Only and tick departments — or use Full Company when the company has no departments yet."
      columns={columns}
      fields={fieldsForRender}
      companies={companies}
      departments={departments}
      initialRows={initialRows}
      listUrl="/api/admin/permissions"
      createUrl="/api/admin/permissions"
      updateUrl={(id) => `/api/admin/permissions/${id}`}
      deleteUrl={(id) => `/api/admin/permissions/${id}`}
      optimistic
      deleteConfirmExtra="This removes portal access for this user. Training records are not deleted."
      drawerWide
      stickyColumnKey="email"
      wideTable
      mapResponse={mapPermissionsResponse}
      searchKeys={permissionSearchKeys}
      rowFilter={rowFilter}
      toolbarExtra={toolbarExtra}
      getCreateDefaults={() => rolePreset ?? {}}
    />
  );
}
