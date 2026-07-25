"use client";

import { useMemo, useState } from "react";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import styles from "@/components/admin/admin.module.css";
import { useAdminToast } from "@/components/admin/AdminToast";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type {
  AdminDocumentRecord,
  DocumentMetadataStatus,
} from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

const columns: AdminColumn<AdminDocumentRecord>[] = [
  {
    key: "name",
    header: "File name",
    render: (row) => (
      <>
        {row.name}
        {row.isFolder ? (
          <span className={styles.muted}> (folder)</span>
        ) : null}
      </>
    ),
  },
  {
    key: "company",
    header: "Company",
    render: (row) => row.company ?? "—",
  },
  {
    key: "candidate",
    header: "Candidate",
    render: (row) => row.candidate ?? "—",
  },
  {
    key: "type",
    header: "Document Type",
    render: (row) => row.documentType ?? "—",
  },
  {
    key: "visible",
    header: "Customer Visible",
    render: (row) => (row.customerVisible ? "Yes" : "No"),
  },
  {
    key: "notification",
    header: "Notification Sent",
    render: (row) => (row.notificationSent ? "Yes" : "No"),
  },
  {
    key: "modified",
    header: "Modified Date",
    render: (row) =>
      row.modifiedDate ? formatDisplayDate(row.modifiedDate) : "—",
  },
  {
    key: "modifiedBy",
    header: "Modified By",
    render: (row) => row.modifiedBy ?? "—",
  },
  {
    key: "metadata",
    header: "Metadata Status",
    render: (row) => (
      <span className={metadataStatusClass(row.metadataStatus)}>
        {row.metadataStatus}
      </span>
    ),
  },
];

const fields: AdminFieldConfig[] = [
  { name: "name", label: "File name", type: "text", required: true },
  { name: "company", label: "Company", type: "company" },
  { name: "candidate", label: "Candidate", type: "text" },
  { name: "documentType", label: "Document type", type: "text" },
  { name: "customerVisible", label: "Customer visible", type: "boolean" },
  { name: "notificationSent", label: "Notification sent", type: "boolean" },
];

function metadataStatusClass(status: DocumentMetadataStatus): string {
  if (status === "Complete") return styles.metadataComplete;
  if (status === "Hidden from Customer") return styles.metadataHidden;
  return styles.metadataMissing;
}

function DocumentActions({
  row,
  reload,
}: {
  row: AdminDocumentRecord;
  reload: () => Promise<void>;
}) {
  const { pushToast } = useAdminToast();
  const [busy, setBusy] = useState(false);

  async function setVisibility(customerVisible: boolean) {
    setBusy(true);
    try {
      const response = await fetch(`/api/admin/documents/${row.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ customerVisible }),
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      pushToast(
        customerVisible
          ? "Document marked customer visible."
          : "Document hidden from customer.",
        "success",
      );
      await reload();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to update visibility",
        "error",
      );
    } finally {
      setBusy(false);
    }
  }

  return (
    <span className={styles.actionGroup}>
      {row.previewPath ? (
        <a
          className={styles.linkButton}
          href={row.previewPath}
          target="_blank"
          rel="noreferrer"
        >
          Preview
        </a>
      ) : null}
      {row.downloadPath ? (
        <a className={styles.linkButton} href={row.downloadPath}>
          Download
        </a>
      ) : (
        <span className={styles.muted}>No file</span>
      )}
      {!row.customerVisible ? (
        <button
          type="button"
          className={styles.linkButton}
          disabled={busy}
          onClick={() => void setVisibility(true)}
        >
          Mark Customer Visible
        </button>
      ) : (
        <button
          type="button"
          className={styles.linkButton}
          disabled={busy}
          onClick={() => void setVisibility(false)}
        >
          Hide from Customer
        </button>
      )}
    </span>
  );
}

export function AdminDocumentsClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminDocumentRecord[];
}) {
  const [documentTypeFilter, setDocumentTypeFilter] = useState("");
  const [candidateFilter, setCandidateFilter] = useState("");
  const [visibilityFilter, setVisibilityFilter] = useState("");

  const documentTypes = useMemo(() => {
    const values = new Set<string>();
    for (const row of initialRows) {
      if (row.documentType?.trim()) values.add(row.documentType.trim());
    }
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  }, [initialRows]);

  const candidates = useMemo(() => {
    const values = new Set<string>();
    for (const row of initialRows) {
      if (row.candidate?.trim()) values.add(row.candidate.trim());
    }
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  }, [initialRows]);

  return (
    <AdminCrudPage<AdminDocumentRecord>
      title="Documents"
      description="Manage Customer Documents metadata, visibility, and company links. Records with missing metadata stay listed so they can be fixed."
      columns={columns}
      fields={fields}
      companies={companies}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.company}
      allowCreate={false}
      editLabel="Edit metadata"
      listUrl="/api/admin/documents"
      updateUrl={(id) => `/api/admin/documents/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminDocumentRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.name,
        (row) => row.documentType,
        (row) => row.company,
        (row) => row.candidate,
        (row) => row.metadataStatus,
        (row) => row.modifiedBy,
      ]}
      rowClassName={(row) =>
        row.metadataStatus === "Complete" ? undefined : styles.rowNeedsMetadata
      }
      rowFilter={(row) => {
        if (
          documentTypeFilter &&
          (row.documentType ?? "").trim().toLowerCase() !==
            documentTypeFilter.trim().toLowerCase()
        ) {
          return false;
        }
        if (
          candidateFilter &&
          (row.candidate ?? "").trim().toLowerCase() !==
            candidateFilter.trim().toLowerCase()
        ) {
          return false;
        }
        if (visibilityFilter === "yes" && !row.customerVisible) return false;
        if (visibilityFilter === "no" && row.customerVisible) return false;
        return true;
      }}
      toolbarExtra={
        <>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Document type</span>
            <select
              className={styles.select}
              value={documentTypeFilter}
              onChange={(event) => setDocumentTypeFilter(event.target.value)}
            >
              <option value="">All types</option>
              {documentTypes.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Candidate</span>
            <select
              className={styles.select}
              value={candidateFilter}
              onChange={(event) => setCandidateFilter(event.target.value)}
            >
              <option value="">All candidates</option>
              {candidates.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Visibility</span>
            <select
              className={styles.select}
              value={visibilityFilter}
              onChange={(event) => setVisibilityFilter(event.target.value)}
            >
              <option value="">All</option>
              <option value="yes">Customer visible</option>
              <option value="no">Hidden from customer</option>
            </select>
          </label>
        </>
      }
      extraActions={(row, { reload }) => (
        <DocumentActions row={row} reload={reload} />
      )}
    />
  );
}
