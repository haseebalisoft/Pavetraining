"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
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
  { name: "candidateName", label: "Candidate name", type: "text", required: true },
  { name: "companyName", label: "Company", type: "company", required: true },
  { name: "nvqTitle", label: "NVQ title", type: "text" },
  { name: "boltOn", label: "Bolt on", type: "text" },
  { name: "dateRegistered", label: "Date registered", type: "date" },
  { name: "inductionDate", label: "Induction date", type: "date" },
  { name: "stageOfNvq", label: "Stage of NVQ", type: "text" },
  { name: "notes", label: "Notes", type: "textarea" },
  { name: "completedDate", label: "Completed date", type: "date" },
  { name: "customerVisible", label: "Customer visible", type: "boolean" },
];

export function AdminNvqClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminNvqRecord[];
}) {
  return (
    <AdminCrudPage<AdminNvqRecord>
      title="NVQ"
      description="Track NVQ stages, notes, and completion dates."
      columns={columns}
      fields={fields}
      companies={companies}
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
