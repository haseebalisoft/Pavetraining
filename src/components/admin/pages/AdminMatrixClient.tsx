"use client";

import { useSearchParams } from "next/navigation";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import styles from "@/components/admin/admin.module.css";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { StatusBadge } from "@/components/ui/StatusBadge";
import type { AdminMatrixRecord } from "@/lib/services/adminCrudService";
import {
  getExpiryStatus,
  matchesExpiryFilter,
  type ExpiryFilter,
} from "@/lib/training/expiryFilters";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminMatrixRecord>[] = [
  { key: "name", header: "Candidate name", render: (row) => row.candidateName },
  { key: "company", header: "Company", render: (row) => row.companyName ?? "—" },
  { key: "department", header: "Department", render: (row) => row.department ?? "—" },
  {
    key: "next",
    header: "Next expiry",
    render: (row) => <ExpiryDateBadge date={row.nextExpiryDate} />,
  },
  {
    key: "review",
    header: "Records to Review",
    render: (row) => (
      <StatusBadge
        label={row.needsReview ? "Review" : "Clear"}
        tone={row.needsReview ? "warn" : "ok"}
      />
    ),
  },
  {
    key: "status",
    header: "Overall status",
    render: (row) => row.overallStatus ?? "—",
  },
];

const fields: AdminFieldConfig[] = [
  { name: "candidateName", label: "Candidate name", type: "text", required: true },
  { name: "companyName", label: "Company", type: "company", required: true },
  { name: "department", label: "Department", type: "text" },
  { name: "overallStatus", label: "Overall status", type: "text" },
  { name: "needsReview", label: "Records to Review", type: "boolean" },
  { name: "matrixNotes", label: "Notes", type: "textarea" },
  { name: "nextExpiryDate", label: "Next expiry date", type: "date" },
  { name: "n001Expiry", label: "N001 expiry", type: "date" },
  { name: "n003Expiry", label: "N003 expiry", type: "date" },
  { name: "n004Expiry", label: "N004 expiry", type: "date" },
  { name: "n010Expiry", label: "N010 expiry", type: "date" },
  { name: "n020Expiry", label: "N020 expiry", type: "date" },
  { name: "n021Expiry", label: "N021 expiry", type: "date" },
  { name: "n027Expiry", label: "N027 expiry", type: "date" },
  { name: "n100Expiry", label: "N100 expiry", type: "date" },
];

function matchesFilter(
  row: AdminMatrixRecord,
  filter: string | null,
): boolean {
  if (!filter || filter === "all") return true;
  if (filter === "review") return row.needsReview;
  if (filter === "expiring") {
    return getExpiryStatus(row.nextExpiryDate).status === "urgent";
  }
  return matchesExpiryFilter(row.nextExpiryDate, filter as ExpiryFilter);
}

export function AdminMatrixClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminMatrixRecord[];
}) {
  const searchParams = useSearchParams();
  const filter = searchParams.get("filter");

  return (
    <AdminCrudPage<AdminMatrixRecord>
      title="Training Matrix"
      description="Review and edit matrix rows, notes, and Records to Review flags."
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      listUrl="/api/admin/training-matrix"
      createUrl="/api/admin/training-matrix"
      updateUrl={(id) => `/api/admin/training-matrix/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminMatrixRecord[] }).records ?? [])
      }
      rowFilter={(row) => matchesFilter(row, filter)}
      rowClassName={(row) => (row.needsReview ? styles.reviewRow : undefined)}
      searchKeys={[
        (row) => row.candidateName,
        (row) => row.companyName,
        (row) => row.department,
        (row) => row.overallStatus,
      ]}
    />
  );
}
