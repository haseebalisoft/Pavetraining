"use client";

import { useMemo, useState } from "react";
import { usePathname, useRouter, useSearchParams } from "next/navigation";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminPermissionRecord } from "@/lib/services/adminCrudService";
import {
  isDepartmentActive,
  type AdminDepartmentRecord,
} from "@/lib/services/departmentTypes";
import type { Company } from "@/types/models";

import styles from "@/components/admin/admin.module.css";

function companyDepartmentsForRow(
  row: AdminPermissionRecord,
  departments: AdminDepartmentRecord[],
): AdminDepartmentRecord[] {
  const companyId = String(row.companyId ?? "").trim();
  const companyName = String(row.companyName ?? "")
    .trim()
    .toLowerCase();
  const seen = new Set<string>();
  const matched: AdminDepartmentRecord[] = [];
  for (const department of departments) {
    if (!isDepartmentActive(department)) continue;
    const sameCompany =
      (companyId &&
        department.companyId &&
        companyId === String(department.companyId)) ||
      (companyName &&
        (department.companyName || "").trim().toLowerCase() === companyName);
    if (!sameCompany) continue;
    const key = department.name.trim().toLowerCase();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    matched.push(department);
  }
  return matched;
}

function departmentsCoveredLabel(
  row: AdminPermissionRecord,
  departments: AdminDepartmentRecord[],
): string {
  const scope = (row.accessScope || "").toLowerCase();
  if (scope.includes("candidate")) return "—";

  if (scope.includes("department")) {
    const stored = row.departmentScopes ?? [];
    if (stored.length === 0) return "—";
    const companyDepts = companyDepartmentsForRow(row, departments);
    return stored
      .map((token) => {
        const match = companyDepts.find(
          (department) =>
            department.id === token ||
            department.name.trim().toLowerCase() === token.trim().toLowerCase(),
        );
        return match?.name ?? token;
      })
      .join(", ");
  }

  // Full Company covers every active department in the company. Do not write
  // those names onto Permissions — this is display-only so the table is not blank.
  const companyDepts = companyDepartmentsForRow(row, departments);
  if (companyDepts.length === 0) return "All departments";
  return companyDepts.map((department) => department.name).join(", ");
}

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
      { value: "Admin", label: "Admin — PAVE staff, full admin portal" },
      { value: "Manager", label: "Manager — Training Manager, company matrix" },
      { value: "Supervisor", label: "Supervisor — assigned on Workforce" },
      { value: "Customer", label: "Customer — own profile only" },
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

function normalizeSpRole(value: string | null | undefined): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    // Fold legacy "TrainingManager" and "trainingmanager" into the same key
    .replace("trainingmanager", "training manager");
}

function permissionRoleFromSharePoint(
  sharePointRole: string | null | undefined,
): "Admin" | "Manager" | "Supervisor" | "Customer" | "Candidate" | null {
  const key = normalizeSpRole(sharePointRole);
  if (key === "admin") return "Admin";
  if (key === "training manager" || key === "manager") return "Manager";
  if (key === "supervisor") return "Supervisor";
  if (key === "customer" || key === "candidate") return "Customer";
  return null;
}

function roleLabelFromSharePoint(
  sharePointRole: string | null | undefined,
  fallback?: string | null,
): string {
  return permissionRoleFromSharePoint(sharePointRole) ?? fallback?.trim() ?? "—";
}

function withCorrectedRoles(
  row: AdminPermissionRecord,
): AdminPermissionRecord {
  const permissionRole =
    permissionRoleFromSharePoint(row.sharePointRoleType) ?? row.permissionRole;
  return {
    ...row,
    permissionRole,
    roleLabel: roleLabelFromSharePoint(row.sharePointRoleType, row.roleLabel),
  };
}

function mapPermissionsResponse(payload: unknown): AdminPermissionRecord[] {
  return ((payload as { records?: AdminPermissionRecord[] }).records ?? []).map(
    withCorrectedRoles,
  );
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
  | "Customer";

const TABS: Array<{ id: PermissionTabId; label: string }> = [
  { id: "all", label: "All" },
  { id: "Admin", label: "Admins" },
  { id: "Training Manager", label: "Managers" },
  { id: "Supervisor", label: "Supervisors" },
  { id: "Customer", label: "Customers" },
];

/** Map the URL ?role= param to a permissionRole form-field default. */
function rolePresetFromQuery(role: string | null): {
  permissionRole: string;
} | null {
  if (!role) return null;
  const key = role.toLowerCase().replace(/[^a-z]/g, "");
  if (key === "admin") {
    return { permissionRole: "Admin" };
  }
  if (key === "trainingmanager" || key === "manager") {
    return { permissionRole: "Manager" };
  }
  if (key === "supervisor") {
    return { permissionRole: "Supervisor" };
  }
  if (key === "customer" || key === "candidate") {
    return { permissionRole: "Customer" };
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

  const columns = useMemo<AdminColumn<AdminPermissionRecord>[]>(
    () => [
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
        render: (row) => departmentsCoveredLabel(row, departments),
      },
      {
        key: "company",
        header: "Company",
        render: (row) => row.companyName ?? "—",
      },
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
    ],
    [departments],
  );

  const rolePreset = useMemo(
    () => rolePresetFromQuery(searchParams.get("role")),
    [searchParams],
  );

  // Lock the "Role" field when the drawer was opened via an "Add Manager"
  // / "Add Supervisor" / "Add Customer" button so admins can't accidentally
  // save the wrong role from a role-preset shortcut. The lock only applies
  // during the shortcut open; picking Edit on any existing row keeps the
  // field editable.
  const fieldsForRender = useMemo<AdminFieldConfig[]>(() => {
    if (!rolePreset) return fields;
    return fields.map((field) =>
      field.name === "permissionRole" ? { ...field, readOnly: true } : field,
    );
  }, [rolePreset]);

  const rowFilter = useMemo(() => {
    if (activeTab === "all") return undefined;
    const target =
      activeTab === "Admin"
        ? "Admin"
        : activeTab === "Training Manager"
          ? "Manager"
          : activeTab === "Supervisor"
            ? "Supervisor"
            : "Customer";
    return (row: AdminPermissionRecord) => {
      const role =
        permissionRoleFromSharePoint(row.sharePointRoleType) ??
        row.permissionRole;
      return role === target;
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
        + Add Manager
      </button>
      <button
        type="button"
        className={styles.permissionAddButton}
        onClick={() => openAddWithRole("Supervisor")}
      >
        + Add Supervisor
      </button>
      <button
        type="button"
        className={styles.permissionAddButton}
        onClick={() => openAddWithRole("Customer")}
      >
        + Add Customer
      </button>
    </div>
  );

  return (
    <AdminCrudPage<AdminPermissionRecord>
      title="Permissions"
      description="Admin = PAVE staff. Manager = Training Manager (company matrix). Supervisor = named on Workforce. Customer = own profile only (email matches Workforce). Edit or Delete any row."
      columns={columns}
      fields={fieldsForRender}
      companies={companies}
      departments={departments}
      initialRows={initialRows.map(withCorrectedRoles)}
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
