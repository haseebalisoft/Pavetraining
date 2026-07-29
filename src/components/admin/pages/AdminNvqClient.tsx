"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
  type AdminWorkforceOption,
} from "@/components/admin/AdminCrudPage";
import styles from "@/components/admin/admin.module.css";
import { StatusBadge } from "@/components/ui/StatusBadge";
import type { AdminNvqRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminNvqRecord>[] = [
  { key: "name", header: "Candidate name", render: (row) => row.candidateName },
  { key: "title", header: "NVQ title", render: (row) => row.nvqTitle ?? "—" },
  { key: "company", header: "Company", render: (row) => row.companyName ?? "—" },
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
    key: "completed",
    header: "Completed date",
    render: (row) => row.completedDate ?? "—",
  },
];

const fields: AdminFieldConfig[] = [
  {
    name: "candidateName",
    label: "Candidate",
    type: "workforce",
    required: true,
    section: "Candidate",
  },
  {
    name: "companyName",
    label: "Company",
    type: "company",
    required: true,
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
    type: "text",
    section: "NVQ",
  },
  { name: "notes", label: "Notes", type: "textarea", section: "NVQ" },
  {
    name: "completedDate",
    label: "Completed date",
    type: "date",
    section: "NVQ",
  },
  {
    name: "customerVisible",
    label: "Customer visible",
    type: "boolean",
    section: "NVQ",
  },
];

export function AdminNvqClient({
  companies,
  workforce,
  initialRows,
}: {
  companies: Company[];
  workforce: AdminWorkforceOption[];
  initialRows: AdminNvqRecord[];
}) {
  return (
    <AdminCrudPage<AdminNvqRecord>
      title="NVQ"
      description="Pick a Workforce candidate to auto-fill name and company. Track NVQ stages, notes, and completion dates."
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
      rowClassName={(row) =>
        row.status === "Completed" ? styles.completedRow : undefined
      }
      searchKeys={[
        (row) => row.candidateName,
        (row) => row.nvqTitle,
        (row) => row.companyName,
        (row) => row.stageOfNvq,
        (row) => row.notes,
      ]}
    />
  );
}
