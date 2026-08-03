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

const columns: AdminColumn<AdminNvqRecord>[] = [
  { key: "name", header: "Candidate name", render: (row) => row.candidateName },
  { key: "title", header: "NVQ title", render: (row) => row.nvqTitle ?? "—" },
  { key: "company", header: "Company", render: (row) => row.companyName ?? "—" },
  { key: "boltOn", header: "Bolt on", render: (row) => row.boltOn ?? "—" },
  { key: "stage", header: "Stage", render: (row) => row.stageOfNvq ?? "—" },
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
    key: "registered",
    header: "Date registered",
    render: (row) => formatDate(row.dateRegistered),
  },
  {
    key: "induction",
    header: "Induction date",
    render: (row) => formatDate(row.inductionDate),
  },
  {
    key: "completed",
    header: "Completed date",
    render: (row) => formatDate(row.completedDate),
  },
  {
    key: "visible",
    header: "Customer visible",
    render: (row) => (row.customerVisible ? "Yes" : "No"),
  },
];

const fields: AdminFieldConfig[] = [
  {
    name: "companyName",
    label: "Company",
    type: "company",
    required: true,
    section: "Candidate",
  },
  {
    name: "candidateName",
    label: "Candidate",
    type: "workforce",
    required: true,
    section: "Candidate",
  },
  {
    name: "workforceNumber",
    label: "Workforce number (from Workforce)",
    type: "text",
    readOnly: true,
    section: "Candidate",
  },
  { name: "nvqTitle", label: "NVQ title", type: "text", section: "NVQ" },
  { name: "boltOn", label: "Bolt on", type: "text", section: "NVQ" },
  {
    name: "dateRegistered",
    label: "Date registered",
    type: "date",
    section: "NVQ",
  },
  {
    name: "inductionDate",
    label: "Induction date",
    type: "date",
    section: "NVQ",
  },
  {
    name: "stageOfNvq",
    label: "Stage of NVQ",
    type: "select",
    section: "NVQ",
    options: getNvqStageOptions(),
  },
  {
    name: "notes",
    label: "Notes / Customer update notes",
    type: "textarea",
    section: "NVQ",
  },
  {
    name: "completedDate",
    label: "Completed date",
    type: "date",
    section: "NVQ",
  },
  {
    name: "trainingOutcome",
    label: "Training outcome",
    type: "select",
    section: "Outcome",
    options: [
      { value: "Pass", label: "Pass" },
      { value: "Fail", label: "Fail" },
    ],
  },
  {
    name: "outcomeDate",
    label: "Outcome date",
    type: "date",
    section: "Outcome",
  },
  {
    name: "assessorTrainer",
    label: "Assessor / trainer",
    type: "text",
    section: "Outcome",
  },
  {
    name: "outcomeNotes",
    label: "Outcome notes",
    type: "textarea",
    section: "Outcome",
  },
  {
    name: "customerVisible",
    label: "Customer visible",
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
      title="NVQ"
      description="Select a company to scope candidates. Completed date blank = Active; filled = Completed (green). Admin sees full NVQ fields; customers only see customer-visible progress."
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
      mapResponse={(payload) =>
        ((payload as { records?: AdminNvqRecord[] }).records ?? [])
      }
      rowFilter={rowFilter}
      rowClassName={(row) =>
        row.status === "Completed" ? styles.completedRow : undefined
      }
      toolbarExtra={
        <div className={styles.syncToolbarActions} role="tablist" aria-label="NVQ status">
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
        (row) => row.nvqTitle,
        (row) => row.companyName,
        (row) => row.boltOn,
        (row) => row.stageOfNvq,
        (row) => row.notes,
      ]}
    />
  );
}
