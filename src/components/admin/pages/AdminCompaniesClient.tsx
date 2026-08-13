"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import { ImageUploadButton } from "@/components/admin/ImageUploadButton";
import { Thumbnail } from "@/components/ui/Thumbnail";
import { allocateNextCompanyNumber } from "@/lib/companyNumber";
import { isDepartmentActive } from "@/lib/services/departmentTypes";
import type { Company } from "@/types/models";

function cell(value: string | null | undefined) {
  const text = value?.trim();
  return text ? text : "—";
}

/**
 * Table columns match Company list.xlsx / SharePoint Company List exactly:
 * Company Number | Company Name | Company Size | Registered Address |
 * Company Reg Number | VAT No | Tel No | Email | Main Contact |
 * Accounts Contact Name | Accounts address | Accounts Contact number |
 * Accounts email | Notes prices agreed | Company Logo | Status
 */
const baseColumns: AdminColumn<Company>[] = [
  {
    key: "companyNumber",
    header: "Company Number",
    render: (row) => cell(row.companyNumber),
  },
  {
    key: "companyName",
    header: "Company Name",
    render: (row) => row.companyName,
  },
  {
    key: "companySize",
    header: "Company Size",
    render: (row) => cell(row.companySize),
  },
  {
    key: "registeredAddress",
    header: "Registered Address",
    render: (row) => cell(row.registeredAddress),
  },
  {
    key: "companyRegNumber",
    header: "Company Reg Number",
    render: (row) => cell(row.companyRegNumber),
  },
  {
    key: "vatNo",
    header: "VAT No",
    render: (row) => cell(row.vatNo),
  },
  {
    key: "telNo",
    header: "Tel No",
    render: (row) => cell(row.telNo),
  },
  {
    key: "email",
    header: "Email",
    render: (row) => cell(row.email),
  },
  {
    key: "mainContact",
    header: "Main Contact",
    render: (row) => cell(row.mainContact),
  },
  {
    key: "accountsContactName",
    header: "Accounts Contact Name",
    render: (row) => cell(row.accountsContactName),
  },
  {
    key: "accountsAddress",
    header: "Accounts address",
    render: (row) => cell(row.accountsAddress),
  },
  {
    key: "accountsContactNumber",
    header: "Accounts Contact number",
    render: (row) => cell(row.accountsContactNumber),
  },
  {
    key: "accountsEmail",
    header: "Accounts email",
    render: (row) => cell(row.accountsEmail),
  },
  {
    key: "notesPricesAgreed",
    header: "Notes prices agreed",
    render: (row) => cell(row.notesPricesAgreed),
  },
  {
    key: "companyLogo",
    header: "Company Logo",
    render: (row) => (
      <Thumbnail
        src={row.companyLogo}
        alt={row.companyName ? `${row.companyName} logo` : "Company logo"}
        variant="company"
      />
    ),
  },
  {
    key: "status",
    header: "Status",
    render: (row) => cell(row.status),
  },
];

const companySizeOptions = [
  { value: "Small", label: "Small" },
  { value: "Medium", label: "Medium" },
  { value: "Large", label: "Large" },
  { value: "Enterprise", label: "Enterprise" },
];

/** Form fields — same column set as Company list.xlsx. */
const fields: AdminFieldConfig[] = [
  {
    name: "companyNumber",
    label: "Company Number",
    type: "text",
    required: false,
    readOnly: true,
    placeholder: "Auto-generated (C00001, C00002, …)",
    section: "Company List",
  },
  {
    name: "companyName",
    label: "Company Name",
    type: "text",
    required: true,
    section: "Company List",
  },
  {
    name: "companySize",
    label: "Company Size",
    type: "select",
    options: companySizeOptions,
    section: "Company List",
  },
  {
    name: "registeredAddress",
    label: "Registered Address",
    type: "textarea",
    section: "Company List",
  },
  {
    name: "companyRegNumber",
    label: "Company Reg Number",
    type: "text",
    section: "Company List",
  },
  {
    name: "vatNo",
    label: "VAT No",
    type: "text",
    section: "Company List",
  },
  {
    name: "telNo",
    label: "Tel No",
    type: "text",
    section: "Company List",
  },
  {
    name: "email",
    label: "Email",
    type: "email",
    section: "Company List",
  },
  {
    name: "mainContact",
    label: "Main Contact",
    type: "text",
    section: "Company List",
  },
  {
    name: "accountsContactName",
    label: "Accounts Contact Name",
    type: "text",
    section: "Company List",
  },
  {
    name: "accountsAddress",
    label: "Accounts address",
    type: "textarea",
    section: "Company List",
  },
  {
    name: "accountsContactNumber",
    label: "Accounts Contact number",
    type: "text",
    section: "Company List",
  },
  {
    name: "accountsEmail",
    label: "Accounts email",
    type: "email",
    section: "Company List",
  },
  {
    name: "notesPricesAgreed",
    label: "Notes prices agreed",
    type: "textarea",
    section: "Company List",
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
    section: "Company List",
  },
];

export function AdminCompaniesClient({
  initialRows,
}: {
  initialRows: Company[];
}) {
  const [departmentCounts, setDepartmentCounts] = useState<
    Record<string, number>
  >({});

  useEffect(() => {
    let cancelled = false;
    fetch("/api/admin/departments")
      .then((response) => (response.ok ? response.json() : null))
      .then(
        (
          payload: {
            records?: { companyId: string | null; status?: string }[];
          } | null,
        ) => {
          if (cancelled || !payload?.records) return;
          const counts: Record<string, number> = {};
          // Active-only count — clicking through to the Departments page
          // shows the full list (active + inactive) with Status visible.
          for (const record of payload.records) {
            if (!record.companyId) continue;
            if (!isDepartmentActive(record)) continue;
            counts[record.companyId] = (counts[record.companyId] ?? 0) + 1;
          }
          setDepartmentCounts(counts);
        },
      )
      .catch(() => {
        // Non-critical — the count column just shows 0 if this fails.
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const columns = useMemo<AdminColumn<Company>[]>(
    () => [
      ...baseColumns,
      {
        key: "departments",
        header: "Departments",
        render: (row) => (
          <Link
            href={`/admin/departments?company=${encodeURIComponent(row.companyName)}`}
          >
            {departmentCounts[row.id] ?? 0} — view
          </Link>
        ),
      },
    ],
    [departmentCounts],
  );

  return (
    <AdminCrudPage<Company>
      title="Companies"
      description="Company List — same columns as SharePoint. Upload the company logo from the row actions; it shows on that company’s customer portal only."
      columns={columns}
      fields={fields}
      initialRows={initialRows}
      listUrl="/api/admin/companies"
      createUrl="/api/admin/companies"
      updateUrl={(id) => `/api/admin/companies/${id}`}
      deleteUrl={(id) => `/api/admin/companies/${id}`}
      bulkDeleteUrl="/api/admin/companies/bulk-delete"
      enableBulkDelete
      // Optimistic UI: delete row from the table instantly; SharePoint
      // cascade (12 lists) runs in the background. Silent restore + review
      // pill if it fails. Non-destructive fallback: refreshing the page
      // shows real SharePoint state.
      optimistic
      optimisticRowLabel={(row) =>
        row.companyName?.trim() || row.companyNumber?.trim() || `#${row.id}`
      }
      getCreateDefaults={(rows) => ({
        companyNumber: allocateNextCompanyNumber(rows),
      })}
      deleteConfirmExtra={
        "This also deletes related Customer Documents, Training Matrix rows, NPORS/EUSR/Streetworks/In-House registers, NVQ records, Events, Workforce candidates, and Permissions for that company. Its Audit / Activity Log history is kept."
      }
      mapResponse={(payload) =>
        (payload as { companies?: Company[] }).companies ?? []
      }
      emptyLabel="No companies found. Add your first company to begin."
      drawerWide
      wideTable
      stickyColumnKey="companyName"
      extraActions={(row, { reload }) => (
        <ImageUploadButton
          uploadUrl={`/api/admin/companies/${row.id}/logo`}
          label="Upload logo"
          onUploaded={reload}
        />
      )}
      searchKeys={[
        (row) => row.companyNumber,
        (row) => row.companyName,
        (row) => row.companySize,
        (row) => row.registeredAddress,
        (row) => row.companyRegNumber,
        (row) => row.vatNo,
        (row) => row.telNo,
        (row) => row.email,
        (row) => row.mainContact,
        (row) => row.accountsContactName,
        (row) => row.accountsAddress,
        (row) => row.accountsContactNumber,
        (row) => row.accountsEmail,
        (row) => row.notesPricesAgreed,
        (row) => row.status,
      ]}
    />
  );
}
