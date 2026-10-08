"use client";

import Link from "next/link";
import { useMemo } from "react";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminColumnRenderContext,
  type AdminFieldConfig,
  type AdminPermissionPersonOption,
} from "@/components/admin/AdminCrudPage";
import styles from "@/components/admin/admin.module.css";
import { ImageUploadButton } from "@/components/admin/ImageUploadButton";
import {
  WorkforceInlineSelect,
  type WorkforceInlineOption,
} from "@/components/admin/WorkforceInlineSelect";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { Thumbnail } from "@/components/ui/Thumbnail";
import type {
  AdminPermissionRecord,
  AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
import type { AdminDepartmentRecord } from "@/lib/services/departmentTypes";
import { EXPIRY_STATUS_LEGEND } from "@/lib/training/expiryFilters";
import { toneForExpiryStatus } from "@/lib/ui/status";
import { allocateNextWorkforceNumber } from "@/lib/workforceNumber";
import { formatDate } from "@/lib/utils/formatDate";
import type { Company } from "@/types/models";

function text(value: string | null | undefined): string {
  return value?.trim() ? value : "—";
}

function idsEqual(
  left: string | null | undefined,
  right: string | null | undefined,
): boolean {
  const a = String(left ?? "").trim();
  const b = String(right ?? "").trim();
  return Boolean(a) && a === b;
}

function companyNameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
}

function toPersonOption(
  person: AdminPermissionPersonOption,
): WorkforceInlineOption | null {
  const value = person.name?.trim() || person.userEmail.trim();
  if (!value) return null;
  return {
    value,
    label: person.name?.trim()
      ? `${person.name.trim()} (${person.userEmail})`
      : person.userEmail,
    email: person.userEmail,
  };
}

function departmentOptionsForRow(
  row: AdminWorkforceRecord,
  departments: AdminDepartmentRecord[],
): WorkforceInlineOption[] {
  const companyId = String(row.companyId ?? "").trim();
  const companyName = companyNameKey(row.companyName);
  return departments
    .filter((dept) => {
      if (companyId && idsEqual(dept.companyId, companyId)) return true;
      if (companyName && companyNameKey(dept.companyName) === companyName) {
        return true;
      }
      return false;
    })
    .map((dept) => ({
      value: dept.name,
      label: dept.name,
    }))
    .sort((a, b) => a.label.localeCompare(b.label));
}

function personOptionsForRow(
  row: AdminWorkforceRecord,
  people: AdminPermissionPersonOption[],
  role: "Training Manager" | "Supervisor",
): WorkforceInlineOption[] {
  const companyId = String(row.companyId ?? "").trim();
  const companyName = companyNameKey(row.companyName);
  const targetKey = role.toLowerCase().replace(/\s+/g, " ");

  return people
    .filter((person) => {
      if ((person.status || "").toLowerCase() !== "active") return false;
      const spRole = (person.sharePointRoleType || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ");
      if (spRole !== targetKey) {
        if (!(role === "Training Manager" && spRole === "trainingmanager")) {
          return false;
        }
      }
      if (companyId && person.companyId && !idsEqual(person.companyId, companyId)) {
        return false;
      }
      if (
        !person.companyId &&
        companyName &&
        companyNameKey(person.companyName) !== companyName
      ) {
        return false;
      }
      return true;
    })
    .map(toPersonOption)
    .filter((option): option is WorkforceInlineOption => option !== null)
    .sort((a, b) => a.label.localeCompare(b.label));
}

function permissionPeopleFromRecords(
  rows: AdminPermissionRecord[],
): AdminPermissionPersonOption[] {
  return rows.map((row) => ({
    id: row.id,
    userEmail: row.userEmail,
    name: row.name,
    status: row.status,
    permissionRole: row.permissionRole,
    sharePointRoleType: row.sharePointRoleType,
    companyId: row.companyId,
    companyName: row.companyName,
  }));
}

function buildColumns(
  departments: AdminDepartmentRecord[],
  people: AdminPermissionPersonOption[],
): AdminColumn<AdminWorkforceRecord>[] {
  const inlineSaved =
    (ctx: AdminColumnRenderContext<AdminWorkforceRecord>) =>
    (record: AdminWorkforceRecord) => {
      ctx.patchRow(record.id, record);
    };

  return [
    {
      key: "photo",
      header: "Photo",
      render: (row) => (
        <Thumbnail
          src={row.photoUrl}
          alt={
            row.candidateName ? `${row.candidateName} photo` : "Candidate photo"
          }
          variant="person"
        />
      ),
    },
    {
      key: "workforceNumber",
      header: "Workforce Number",
      render: (row) => text(row.workforceNumber),
    },
    {
      key: "candidateName",
      header: "Candidate Name",
      render: (row) => (
        <Link
          className={styles.matrixNameLink}
          href={`/admin/workforce/${row.id}`}
        >
          {row.candidateName}
        </Link>
      ),
    },
    {
      key: "companyName",
      header: "Company Name",
      render: (row) => row.companyName,
    },
    {
      key: "companyNumber",
      header: "Company Number",
      render: (row) => text(row.companyNumber),
    },
    {
      key: "trainingManager",
      header: "Training manager",
      render: (row, ctx) => (
        <WorkforceInlineSelect
          row={row}
          field="trainingManager"
          options={personOptionsForRow(row, people, "Training Manager")}
          onSaved={inlineSaved(ctx)}
        />
      ),
    },
    {
      key: "supervisor",
      header: "Supervisor",
      render: (row, ctx) => (
        <WorkforceInlineSelect
          row={row}
          field="supervisor"
          options={personOptionsForRow(row, people, "Supervisor")}
          onSaved={inlineSaved(ctx)}
        />
      ),
    },
    {
      key: "candidateAddress",
      header: "Candidate Address",
      render: (row) => text(row.candidateAddress),
    },
    {
      key: "email",
      header: "Email",
      render: (row) => text(row.email),
    },
    {
      key: "contactNumber",
      header: "Contact number",
      render: (row) => text(row.contactNumber),
    },
    {
      key: "dateOfBirth",
      header: "Date of birth",
      render: (row) => formatDate(row.dateOfBirth),
    },
    {
      key: "niNumber",
      header: "Ni Number",
      render: (row) => text(row.niNumber),
    },
    {
      key: "nporsNumbers",
      header: "NPORS Number",
      render: (row) => text(row.nporsNumbers),
    },
    {
      key: "cscsNumber",
      header: "CSCS Number",
      render: (row) => text(row.cscsNumber),
    },
    {
      key: "cscsExpiry",
      header: "Cscs Expiry",
      render: (row) => <ExpiryDateBadge date={row.cscsExpiry} fillCell />,
    },
    {
      key: "swqrNumber",
      header: "SWQR Number",
      render: (row) => text(row.swqrNumber),
    },
    {
      key: "swqrExpiry",
      header: "Swqr Expiry",
      render: (row) => <ExpiryDateBadge date={row.swqrExpiry} fillCell />,
    },
    {
      key: "eusrNumber",
      header: "EUSR Number",
      render: (row) => text(row.eusrNumber),
    },
    {
      key: "eusrExpiry",
      header: "Eusr Expiry",
      render: (row) => <ExpiryDateBadge date={row.eusrExpiry} fillCell />,
    },
    {
      key: "inHouseCertificationNumber",
      header: "In House Certification Number",
      render: (row) => text(row.inHouseCertificationNumber),
    },
    {
      key: "department",
      header: "Department",
      render: (row, ctx) => (
        <WorkforceInlineSelect
          row={row}
          field="department"
          options={departmentOptionsForRow(row, departments)}
          onSaved={inlineSaved(ctx)}
        />
      ),
    },
    {
      key: "status",
      header: "Status",
      render: (row) => text(row.status),
    },
    {
      key: "notes",
      header: "Notes",
      render: (row) => text(row.notes),
    },
  ];
}

const fields: AdminFieldConfig[] = [
  {
    name: "candidateName",
    label: "Candidate Name",
    type: "text",
    required: true,
    section: "Candidate",
  },
  {
    name: "companyName",
    label: "Company Name",
    type: "company",
    required: true,
  },
  {
    name: "workforceNumber",
    label: "Workforce Number",
    type: "text",
    required: false,
    readOnly: true,
    placeholder: "Auto-generated (W00001, W00002, …)",
  },
  {
    name: "trainingManager",
    label: "Training manager",
    type: "select",
    /**
     * STRICT SharePoint RoleType filter — only rows where
     * `RoleType = Training Manager` (Active, matching company). "+ Add new"
     * quick-create appears next to the dropdown.
     */
    sharePointRoleTypeFilter: "Training Manager",
  },
  {
    name: "supervisor",
    label: "Supervisor",
    type: "select",
    /** STRICT SharePoint RoleType = Supervisor only. */
    sharePointRoleTypeFilter: "Supervisor",
  },
  { name: "candidateAddress", label: "Candidate Address", type: "text" },
  { name: "email", label: "Email", type: "email" },
  { name: "contactNumber", label: "Contact number", type: "text" },
  { name: "dateOfBirth", label: "Date of birth", type: "date" },
  { name: "niNumber", label: "Ni Number", type: "text" },
  { name: "nporsNumbers", label: "NPORS Number", type: "text" },
  { name: "cscsNumber", label: "CSCS Number", type: "text" },
  { name: "cscsExpiry", label: "Cscs Expiry", type: "date" },
  { name: "swqrNumber", label: "Swqr Number", type: "text" },
  { name: "swqrExpiry", label: "Swqr Expiry", type: "date" },
  { name: "eusrNumber", label: "Eusr Number", type: "text" },
  { name: "eusrExpiry", label: "Eusr Expiry", type: "date" },
  {
    name: "inHouseCertificationNumber",
    label: "In House Certification Number (optional)",
    type: "text",
    placeholder: "Optional — course-specific numbers belong on the In-House record",
  },
  {
    name: "department",
    label: "Department",
    type: "select",
    companyScopedDepartments: true,
    departmentValueMode: "name",
  },
  {
    name: "status",
    label: "Status",
    type: "select",
    options: [
      { value: "Active", label: "Active" },
      { value: "Inactive", label: "Inactive" },
    ],
  },
  { name: "notes", label: "Notes", type: "textarea" },
];

export function AdminWorkforceClient({
  companies,
  departments,
  initialRows,
  permissionPeople,
}: {
  companies: Company[];
  departments: AdminDepartmentRecord[];
  initialRows: AdminWorkforceRecord[];
  permissionPeople: AdminPermissionRecord[];
}) {
  const people = useMemo(
    () => permissionPeopleFromRecords(permissionPeople),
    [permissionPeople],
  );
  const columns = useMemo(
    () => buildColumns(departments, people),
    [departments, people],
  );

  return (
    <AdminCrudPage<AdminWorkforceRecord>
      title="Workforce"
      description="Manage candidates and link them to companies. Department, Training manager, and Supervisor can be changed inline in the table (options are limited to the candidate’s company). After you pick a company in Edit, Training manager / Supervisor list Active Permissions people for that company. Upload a candidate photo from the row actions — it appears on their profile."
      columns={columns}
      fields={fields}
      companies={companies}
      departments={departments}
      permissionPeople={people}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      getCreateDefaults={(rows) => ({
        workforceNumber: allocateNextWorkforceNumber(rows),
      })}
      toolbarExtra={
        <div className={styles.legendRow} aria-label="Expiry colour legend">
          {EXPIRY_STATUS_LEGEND.map((item) => (
            <StatusBadge
              key={item.status}
              label={item.label}
              tone={toneForExpiryStatus(item.status)}
            />
          ))}
        </div>
      }
      drawerWide
      listUrl="/api/admin/workforce"
      createUrl="/api/admin/workforce"
      updateUrl={(id) => `/api/admin/workforce/${id}`}
      deleteUrl={(id) => `/api/admin/workforce/${id}`}
      optimistic
      deleteConfirmExtra="This removes the candidate from Workforce and the Training Matrix. Past NPORS / EUSR / Streetworks / In-House / NVQ records are kept for history."
      mapResponse={(payload) =>
        ((payload as { records?: AdminWorkforceRecord[] }).records ?? [])
      }
      stickyColumnKey="candidateName"
      wideTable
      tableClassName={styles.workforceWideTable}
      extraActions={(row, { reload }) => (
        <ImageUploadButton
          uploadUrl={`/api/admin/workforce/${row.id}/photo`}
          label="Upload photo"
          onUploaded={reload}
        />
      )}
      searchKeys={[
        (row) => row.candidateName,
        (row) => row.companyName,
        (row) => row.companyNumber,
        (row) => row.workforceNumber,
        (row) => row.department,
        (row) => row.email,
        (row) => row.niNumber,
        (row) => row.cscsNumber,
        (row) => row.swqrNumber,
        (row) => row.eusrNumber,
        (row) => row.nporsNumbers,
        (row) => row.trainingManager,
        (row) => row.supervisor,
      ]}
    />
  );
}
