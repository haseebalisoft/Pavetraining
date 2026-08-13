"use client";

import { useMemo, useState } from "react";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
  type AdminWorkforceOption,
} from "@/components/admin/AdminCrudPage";
import styles from "@/components/admin/admin.module.css";
import { StatusBadge } from "@/components/ui/StatusBadge";
import type { AdminNvqRecord } from "@/lib/services/adminCrudService";
import { getNvqStageOptions } from "@/lib/training/nvqOptions";
import { formatDate } from "@/lib/utils/formatDate";
import type { Company } from "@/types/models";

function text(value: string | null | undefined): string {
  return value?.trim() ? value : "—";
}

function yesNo(value: boolean): string {
  return value ? "Yes" : "No";
}

/** Match NVQ Register list / CSV column order for the admin table. */
const columns: AdminColumn<AdminNvqRecord>[] = [
  {
    key: "candidateName",
    header: "Candidate Name",
    render: (row) => row.candidateName,
  },
  {
    key: "workforceNumber",
    header: "Workforce Number",
    render: (row) => text(row.workforceNumber),
  },
  {
    key: "companyNumber",
    header: "Company Number",
    render: (row) => text(row.companyNumber),
  },
  {
    key: "companyName",
    header: "NVQ Company",
    render: (row) => text(row.companyName),
  },
  {
    key: "niNumber",
    header: "NI Number",
    render: (row) => text(row.niNumber),
  },
  {
    key: "ulnNumber",
    header: "ULN Number",
    render: (row) => text(row.ulnNumber),
  },
  {
    key: "nvqTitle",
    header: "NVQ Title",
    render: (row) => text(row.nvqTitle),
  },
  {
    key: "boltOn",
    header: "Bolt-on NVQ",
    render: (row) => text(row.boltOn),
  },
  {
    key: "poNumber",
    header: "PO Number",
    render: (row) => text(row.poNumber),
  },
  {
    key: "cardSchemeCategory",
    header: "Card Scheme Category",
    render: (row) => text(row.cardSchemeCategory),
  },
  {
    key: "cardExtensionDateNeeded",
    header: "Card Extension Date Needed",
    render: (row) => formatDate(row.cardExtensionDateNeeded),
  },
  {
    key: "siteAddress",
    header: "Site Address",
    render: (row) => text(row.siteAddress),
  },
  {
    key: "siteContact",
    header: "Site Contact Name/Number",
    render: (row) => text(row.siteContact),
  },
  {
    key: "englishUnderstandingConfirmed",
    header: "English Understanding Confirmed",
    render: (row) => yesNo(row.englishUnderstandingConfirmed),
  },
  {
    key: "tcAcknowledged",
    header: "T&C Acknowledged",
    render: (row) => yesNo(row.tcAcknowledged),
  },
  {
    key: "gdprConsent",
    header: "GDPR Consent",
    render: (row) => yesNo(row.gdprConsent),
  },
  {
    key: "dateRegistered",
    header: "Date Registered",
    render: (row) => formatDate(row.dateRegistered),
  },
  {
    key: "inductionDate",
    header: "Date Induction Booked",
    render: (row) => formatDate(row.inductionDate),
  },
  {
    key: "stageOfNvq",
    header: "Stage of NVQ",
    render: (row) => text(row.stageOfNvq),
  },
  {
    key: "notes",
    header: "Notes",
    render: (row) => text(row.notes),
  },
  {
    key: "completedDate",
    header: "Completed Date",
    render: (row) => formatDate(row.completedDate),
  },
  {
    key: "certificationDate",
    header: "Certification Date",
    render: (row) => formatDate(row.certificationDate),
  },
  {
    key: "customerUpdateNotes",
    header: "Customer Update Notes",
    render: (row) => text(row.customerUpdateNotes),
  },
  {
    key: "status",
    header: "Status",
    render: (row) => (
      <StatusBadge
        label={row.status}
        tone={row.status === "Completed" ? "ok" : "info"}
      />
    ),
  },
  {
    key: "trainingOutcome",
    header: "Training Outcome",
    render: (row) => text(row.trainingOutcome),
  },
  {
    key: "outcomeDate",
    header: "Outcome Date",
    render: (row) => formatDate(row.outcomeDate),
  },
  {
    key: "assessorTrainer",
    header: "Assessor Trainer",
    render: (row) => text(row.assessorTrainer),
  },
  {
    key: "customerVisible",
    header: "Customer Visible",
    render: (row) => yesNo(row.customerVisible),
  },
  {
    key: "outcomeNotes",
    header: "Outcome Notes",
    render: (row) => text(row.outcomeNotes),
  },
];

const fields: AdminFieldConfig[] = [
  {
    name: "companyName",
    label: "NVQ Company",
    type: "company",
    required: true,
    section: "Candidate",
  },
  {
    name: "companyNumber",
    label: "Company Number",
    type: "text",
    readOnly: true,
    section: "Candidate",
    placeholder: "From Company List (auto)",
  },
  {
    name: "candidateName",
    label: "Candidate Name",
    type: "workforce",
    required: true,
    section: "Candidate",
  },
  {
    name: "workforceNumber",
    label: "Workforce Number",
    type: "text",
    readOnly: true,
    section: "Candidate",
    placeholder: "From Workforce (auto)",
  },
  {
    name: "niNumber",
    label: "NI Number",
    type: "text",
    section: "Candidate",
  },
  {
    name: "ulnNumber",
    label: "ULN Number",
    type: "text",
    section: "Candidate",
  },
  {
    name: "nvqTitle",
    label: "NVQ Title",
    type: "text",
    section: "NVQ",
  },
  {
    name: "boltOn",
    label: "Bolt-on NVQ",
    type: "text",
    section: "NVQ",
  },
  {
    name: "poNumber",
    label: "PO Number",
    type: "text",
    section: "NVQ",
  },
  {
    name: "cardSchemeCategory",
    label: "Card Scheme Category",
    type: "text",
    section: "NVQ",
  },
  {
    name: "cardExtensionDateNeeded",
    label: "Card Extension Date Needed",
    type: "date",
    section: "NVQ",
  },
  {
    name: "siteAddress",
    label: "Site Address",
    type: "textarea",
    section: "Site",
  },
  {
    name: "siteContact",
    label: "Site Contact Name/Number",
    type: "text",
    section: "Site",
  },
  {
    name: "englishUnderstandingConfirmed",
    label: "English Understanding Confirmed",
    type: "boolean",
    section: "Consents",
  },
  {
    name: "tcAcknowledged",
    label: "T&C Acknowledged",
    type: "boolean",
    section: "Consents",
  },
  {
    name: "gdprConsent",
    label: "GDPR Consent",
    type: "boolean",
    section: "Consents",
  },
  {
    name: "dateRegistered",
    label: "Date Registered",
    type: "date",
    section: "Progress",
  },
  {
    name: "inductionDate",
    label: "Date Induction Booked",
    type: "date",
    section: "Progress",
  },
  {
    name: "stageOfNvq",
    label: "Stage of NVQ",
    type: "select",
    section: "Progress",
    options: getNvqStageOptions(),
  },
  {
    name: "notes",
    label: "Notes",
    type: "textarea",
    section: "Progress",
  },
  {
    name: "completedDate",
    label: "Completed Date",
    type: "date",
    section: "Progress",
  },
  {
    name: "certificationDate",
    label: "Certification Date",
    type: "date",
    section: "Progress",
  },
  {
    name: "customerUpdateNotes",
    label: "Customer Update Notes",
    type: "textarea",
    section: "Progress",
  },
  {
    name: "trainingOutcome",
    label: "Training Outcome",
    type: "select",
    section: "Outcome",
    options: [
      { value: "Pass", label: "Pass" },
      { value: "Fail", label: "Fail" },
    ],
  },
  {
    name: "outcomeDate",
    label: "Outcome Date",
    type: "date",
    section: "Outcome",
  },
  {
    name: "assessorTrainer",
    label: "Assessor Trainer",
    type: "text",
    section: "Outcome",
  },
  {
    name: "outcomeNotes",
    label: "Outcome Notes",
    type: "textarea",
    section: "Outcome",
  },
  {
    name: "customerVisible",
    label: "Customer Visible",
    type: "boolean",
    section: "Outcome",
  },
];

type NvqTab = "Active" | "Completed";

export function AdminNvqClient({
  companies,
  workforce,
  initialRows,
}: {
  companies: Company[];
  workforce: AdminWorkforceOption[];
  initialRows: AdminNvqRecord[];
}) {
  const [tab, setTab] = useState<NvqTab>("Active");

  const rowFilter = useMemo(
    () => (row: AdminNvqRecord) => row.status === tab,
    [tab],
  );

  return (
    <AdminCrudPage<AdminNvqRecord>
      title="NVQ Register"
      description="Columns match the SharePoint NVQ Register list. Company Number and Workforce Number auto-fill from Company / Workforce so records cannot be mixed between companies or candidates. Completed Date blank = Active; filled = Completed."
      columns={columns}
      fields={fields}
      companies={companies}
      workforce={workforce}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      listUrl="/api/admin/nvq"
      createUrl="/api/admin/nvq"
      updateUrl={(id) => `/api/admin/nvq/${id}`}
      deleteUrl={(id) => `/api/admin/nvq/${id}`}
      optimistic
      mapResponse={(payload) =>
        ((payload as { records?: AdminNvqRecord[] }).records ?? [])
      }
      rowFilter={rowFilter}
      rowClassName={(row) =>
        row.status === "Completed" ? styles.completedRow : undefined
      }
      drawerWide
      wideTable
      stickyColumnKey="candidateName"
      toolbarExtra={
        <div
          className={styles.syncToolbarActions}
          role="tablist"
          aria-label="NVQ status"
        >
          {(["Active", "Completed"] as const).map((option) => (
            <button
              key={option}
              type="button"
              role="tab"
              aria-selected={tab === option}
              className={
                tab === option ? styles.primaryButton : styles.secondaryButton
              }
              onClick={() => setTab(option)}
            >
              {option}
            </button>
          ))}
        </div>
      }
      searchKeys={[
        (row) => row.candidateName,
        (row) => row.workforceNumber,
        (row) => row.companyNumber,
        (row) => row.companyName,
        (row) => row.niNumber,
        (row) => row.ulnNumber,
        (row) => row.nvqTitle,
        (row) => row.boltOn,
        (row) => row.poNumber,
        (row) => row.stageOfNvq,
        (row) => row.notes,
        (row) => row.customerUpdateNotes,
        (row) => row.siteAddress,
        (row) => row.siteContact,
      ]}
    />
  );
}
