"use client";

import Link from "next/link";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import styles from "@/components/admin/admin.module.css";
import { ImageUploadButton } from "@/components/admin/ImageUploadButton";
import { Thumbnail } from "@/components/ui/Thumbnail";
import type {
  AdminPermissionRecord,
  AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
import type { AdminDepartmentRecord } from "@/lib/services/departmentTypes";
import { allocateNextWorkforceNumber } from "@/lib/workforceNumber";
import { formatDate } from "@/lib/utils/formatDate";
import type { Company } from "@/types/models";

function text(value: string | null | undefined): string {
  return value?.trim() ? value : "—";
}

/** Match Workforce list.xlsx column order for admin table. */
const columns: AdminColumn<AdminWorkforceRecord>[] = [
  {
    key: "photo",
    header: "Photo",
    render: (row) => (
      <Thumbnail
        src={row.photoUrl}
        alt={row.candidateName ? `${row.candidateName} photo` : "Candidate photo"}
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
      <Link className={styles.matrixNameLink} href={`/admin/workforce/${row.id}`}>
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
    render: (row) => text(row.trainingManager),
  },
  {
    key: "supervisor",
    header: "Supervisor",
    render: (row) => text(row.supervisor),
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
    render: (row) => formatDate(row.cscsExpiry),
  },
  {
    key: "swqrNumber",
    header: "SWQR Number",
    render: (row) => text(row.swqrNumber),
  },
  {
    key: "swqrExpiry",
    header: "Swqr Expiry",
    render: (row) => formatDate(row.swqrExpiry),
  },
  {
    key: "eusrNumber",
    header: "EUSR Number",
    render: (row) => text(row.eusrNumber),
  },
  {
    key: "eusrExpiry",
    header: "Eusr Expiry",
    render: (row) => formatDate(row.eusrExpiry),
  },
  {
    key: "inHouseCertificationNumber",
    header: "In House Certification Number",
    render: (row) => text(row.inHouseCertificationNumber),
  },
  {
    key: "department",
    header: "Department",
    render: (row) => text(row.department),
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
    /** Permissions Training Manager (Admin form role), scoped to selected company. */
    permissionRoleFilter: "Admin",
  },
  {
    name: "supervisor",
    label: "Supervisor",
    type: "select",
    /** Permissions Supervisor (Customer form role), scoped to selected company. */
    permissionRoleFilter: "Customer",
  },
  { name: "candidateAddress", label: "Candidate Address", type: "text" },
  { name: "email", label: "Email", type: "email" },
  { name: "contactNumber", label: "Contact number", type: "text" },
  { name: "dateOfBirth", label: "Date of birth", type: "date" },
  { name: "niNumber", label: "Ni Number", type: "text" },
  { name: "nporsNumbers", label: "NPORS Number", type: "text" },
  { name: "cscsNumber", label: "CSCS Number", type: "text" },
  { name: "cscsExpiry", label: "Cscs Expiry", type: "date" },
  { name: "swqrNumber", label: "SWQR Number", type: "text" },
  { name: "swqrExpiry", label: "Swqr Expiry", type: "date" },
  { name: "eusrNumber", label: "EUSR Number", type: "text" },
  { name: "eusrExpiry", label: "Eusr Expiry", type: "date" },
  {
    name: "inHouseCertificationNumber",
    label: "In House Certification Number",
    type: "text",
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
      { value: "inactive", label: "Inactive" },
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
  return (
    <AdminCrudPage<AdminWorkforceRecord>
      title="Workforce"
      description="Manage candidates and link them to companies. After you pick a company, Training manager / Supervisor list only Active Permissions people for that company (RoleType Training Manager / Supervisor). Department is company-scoped (Admin → Departments). Upload a candidate photo from the row actions — it appears on their profile."
      columns={columns}
      fields={fields}
      companies={companies}
      departments={departments}
      permissionPeople={permissionPeople}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      getCreateDefaults={(rows) => ({
        workforceNumber: allocateNextWorkforceNumber(rows),
      })}
      drawerWide
      listUrl="/api/admin/workforce"
      createUrl="/api/admin/workforce"
      updateUrl={(id) => `/api/admin/workforce/${id}`}
      deleteUrl={(id) => `/api/admin/workforce/${id}`}
      deleteConfirmExtra="This removes the candidate from Workforce and the Training Matrix. Past NPORS / EUSR / Streetworks / In-House / NVQ records are kept for history."
      mapResponse={(payload) =>
        ((payload as { records?: AdminWorkforceRecord[] }).records ?? [])
      }
      stickyColumnKey="candidateName"
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
