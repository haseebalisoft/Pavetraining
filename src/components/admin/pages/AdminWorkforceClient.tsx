"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { AdminWorkforceRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

function formatDateCell(value: string | null | undefined): string {
  if (!value?.trim()) return "—";
  // SharePoint often returns ISO datetime — show date only.
  if (/^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0, 10);
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return parsed.toISOString().slice(0, 10);
}

function text(value: string | null | undefined): string {
  return value?.trim() ? value : "—";
}

/** Match Workforce list.xlsx column order for admin table. */
const columns: AdminColumn<AdminWorkforceRecord>[] = [
  {
    key: "workforceNumber",
    header: "Workforce Number",
    render: (row) => text(row.workforceNumber),
  },
  {
    key: "candidateName",
    header: "Candidate Name",
    render: (row) => row.candidateName,
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
    render: (row) => formatDateCell(row.dateOfBirth),
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
    render: (row) => formatDateCell(row.cscsExpiry),
  },
  {
    key: "swqrNumber",
    header: "SWQR Number",
    render: (row) => text(row.swqrNumber),
  },
  {
    key: "swqrExpiry",
    header: "Swqr Expiry",
    render: (row) => formatDateCell(row.swqrExpiry),
  },
  {
    key: "eusrNumber",
    header: "EUSR Number",
    render: (row) => text(row.eusrNumber),
  },
  {
    key: "eusrExpiry",
    header: "Eusr Expiry",
    render: (row) => formatDateCell(row.eusrExpiry),
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
  { name: "workforceNumber", label: "Workforce Number", type: "text" },
  { name: "trainingManager", label: "Training manager", type: "text" },
  { name: "supervisor", label: "Supervisor", type: "text" },
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
  { name: "department", label: "Department", type: "text" },
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
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminWorkforceRecord[];
}) {
  return (
    <AdminCrudPage<AdminWorkforceRecord>
      title="Workforce"
      description="Manage candidates and link them to companies. Columns match the Workforce list.xlsx template."
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      drawerWide
      listUrl="/api/admin/workforce"
      createUrl="/api/admin/workforce"
      updateUrl={(id) => `/api/admin/workforce/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminWorkforceRecord[] }).records ?? [])
      }
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
