"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import type { Company } from "@/types/models";

const columns: AdminColumn<Company>[] = [
  {
    key: "number",
    header: "Company Number",
    render: (row) => row.companyNumber ?? "—",
  },
  { key: "name", header: "Company Name", render: (row) => row.companyName },
  {
    key: "size",
    header: "Company Size",
    render: (row) => row.companySize ?? "—",
  },
  {
    key: "mainContact",
    header: "Main Contact",
    render: (row) => row.mainContact ?? "—",
  },
  { key: "email", header: "Email", render: (row) => row.email ?? "—" },
  { key: "tel", header: "Tel No", render: (row) => row.telNo ?? "—" },
  { key: "status", header: "Status", render: (row) => row.status },
];

const companySizeOptions = [
  { value: "Small", label: "Small" },
  { value: "Medium", label: "Medium" },
  { value: "Large", label: "Large" },
  { value: "Enterprise", label: "Enterprise" },
  { value: "Other", label: "Other" },
];

const fields: AdminFieldConfig[] = [
  {
    name: "companyName",
    label: "Company Name",
    type: "text",
    required: true,
    section: "Company Details",
  },
  {
    name: "companyNumber",
    label: "Company Number",
    type: "text",
    required: true,
    section: "Company Details",
  },
  {
    name: "companySize",
    label: "Company Size",
    type: "select",
    options: companySizeOptions,
    section: "Company Details",
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
    section: "Company Details",
  },
  {
    name: "registeredAddress",
    label: "Registered Address",
    type: "textarea",
    section: "Registration Details",
  },
  {
    name: "companyRegNumber",
    label: "Company Reg Number",
    type: "text",
    section: "Registration Details",
  },
  {
    name: "vatNo",
    label: "VAT No",
    type: "text",
    section: "Registration Details",
  },
  {
    name: "mainContact",
    label: "Main Contact",
    type: "text",
    section: "Main Contact",
  },
  {
    name: "telNo",
    label: "Tel No",
    type: "text",
    section: "Main Contact",
  },
  {
    name: "email",
    label: "Email",
    type: "email",
    section: "Main Contact",
  },
  {
    name: "accountsContactName",
    label: "Accounts Contact Name",
    type: "text",
    section: "Accounts Contact",
  },
  {
    name: "accountsAddress",
    label: "Accounts Address",
    type: "textarea",
    section: "Accounts Contact",
  },
  {
    name: "accountsContactNumber",
    label: "Accounts Contact Number",
    type: "text",
    section: "Accounts Contact",
  },
  {
    name: "accountsEmail",
    label: "Accounts Email",
    type: "email",
    section: "Accounts Contact",
  },
  {
    name: "notesPricesAgreed",
    label: "Notes / Prices Agreed",
    type: "textarea",
    section: "Notes and Branding",
  },
  {
    name: "companyLogo",
    label: "Company Logo URL",
    type: "text",
    placeholder: "https://…",
    section: "Notes and Branding",
  },
];

export function AdminCompaniesClient({
  initialRows,
}: {
  initialRows: Company[];
}) {
  return (
    <AdminCrudPage<Company>
      title="Companies"
      description="Manage client companies, contacts, account details, and portal status."
      columns={columns}
      fields={fields}
      initialRows={initialRows}
      listUrl="/api/admin/companies"
      createUrl="/api/admin/companies"
      updateUrl={(id) => `/api/admin/companies/${id}`}
      mapResponse={(payload) =>
        (payload as { companies?: Company[] }).companies ?? []
      }
      emptyLabel="No companies found. Add your first company to begin."
      drawerWide
      searchKeys={[
        (row) => row.companyName,
        (row) => row.companyNumber,
        (row) => row.companySize,
        (row) => row.mainContact,
        (row) => row.email,
        (row) => row.telNo,
        (row) => row.status,
      ]}
    />
  );
}
