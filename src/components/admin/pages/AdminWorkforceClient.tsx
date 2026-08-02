"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import { ImageUploadButton } from "@/components/admin/ImageUploadButton";
import { Thumbnail } from "@/components/ui/Thumbnail";
import type {
  AdminPermissionRecord,
  AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
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

function buildWorkforceFields(
  trainingManagerOptions: Array<{ value: string; label: string }>,
  supervisorOptions: Array<{ value: string; label: string }>,
): AdminFieldConfig[] {
  return [
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
    {
      name: "trainingManager",
      label: "Training manager",
      type: "select",
      options: trainingManagerOptions,
    },
    {
      name: "supervisor",
      label: "Supervisor",
      type: "select",
      options: supervisorOptions,
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
    { name: "department", label: "Department", type: "text" },
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
}

function permissionOptionLabel(row: {
  name: string | null;
  userEmail: string;
}): string {
  const name = row.name?.trim();
  return name ? `${name} (${row.userEmail})` : row.userEmail;
}

export function AdminWorkforceClient({
  companies,
  initialRows,
  permissionPeople,
}: {
  companies: Company[];
  initialRows: AdminWorkforceRecord[];
  permissionPeople: AdminPermissionRecord[];
}) {
  const activePeople = permissionPeople.filter(
    (row) => (row.status || "").toLowerCase() === "active",
  );

  const withCurrentValue = (
    options: Array<{ value: string; label: string }>,
    currentValues: Array<string | null | undefined>,
  ) => {
    const seen = new Set(options.map((o) => o.value));
    const extra = currentValues
      .map((v) => v?.trim() || "")
      .filter((v) => v && !seen.has(v))
      .map((v) => ({ value: v, label: `${v} (not in active Permissions)` }));
    return [...options, ...extra];
  };

  const trainingManagerOptions = withCurrentValue(
    [
      { value: "", label: "— None —" },
      ...activePeople
        .filter((row) => row.permissionRole === "Admin")
        .map((row) => ({
          value: row.name?.trim() || row.userEmail,
          label: permissionOptionLabel(row),
        })),
    ],
    initialRows.map((row) => row.trainingManager),
  );
  const supervisorOptions = withCurrentValue(
    [
      { value: "", label: "— None —" },
      ...activePeople
        .filter((row) => row.permissionRole === "Customer")
        .map((row) => ({
          value: row.name?.trim() || row.userEmail,
          label: permissionOptionLabel(row),
        })),
    ],
    initialRows.map((row) => row.supervisor),
  );
  const fields = buildWorkforceFields(
    trainingManagerOptions,
    supervisorOptions,
  );

  return (
    <AdminCrudPage<AdminWorkforceRecord>
      title="Workforce"
      description="Manage candidates and link them to companies. Training manager / Supervisor must already exist on the Permissions list. Upload a candidate photo from the row actions — it appears on their profile."
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
