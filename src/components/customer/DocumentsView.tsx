"use client";

import { useMemo, useState } from "react";

import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type { CustomerDocumentRecord } from "@/types/models";

import styles from "./customer.module.css";

interface Props {
  companyName: string;
  canDownload: boolean;
  records: CustomerDocumentRecord[];
}

function cell(value: string | null | undefined) {
  if (!value?.trim()) {
    return <span className={styles.muted}>—</span>;
  }
  return value;
}

export function DocumentsView({ companyName, canDownload, records }: Props) {
  const [search, setSearch] = useState("");
  const [documentTypeFilter, setDocumentTypeFilter] = useState("");
  const [candidateFilter, setCandidateFilter] = useState("");

  const documentTypes = useMemo(() => {
    const values = new Set<string>();
    for (const row of records) {
      if (row.documentType?.trim()) values.add(row.documentType.trim());
    }
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  }, [records]);

  const candidates = useMemo(() => {
    const values = new Set<string>();
    for (const row of records) {
      if (row.candidate?.trim()) values.add(row.candidate.trim());
    }
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  }, [records]);

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    return records.filter((row) => {
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
      if (!query) return true;
      return [row.name, row.documentType, row.candidate]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [records, search, documentTypeFilter, candidateFilter]);

  return (
    <div>
      <CustomerPageHeader
        breadcrumbs={[
          { label: "Customer", href: "/customer" },
          { label: "Documents" },
        ]}
        title="Documents"
        subtitle="Company documents shared with your organisation."
      />

      <p className={styles.companyMeta}>
        Showing documents for <strong>{companyName}</strong>
        {" · "}
        <StatusBadge
          label={canDownload ? "Downloads enabled" : "View only"}
          tone={canDownload ? "ok" : "neutral"}
        />
      </p>

      <div className={styles.toolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            placeholder="Search by file name, type, or candidate…"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
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
      </div>

      <p className={styles.resultCount}>
        {filtered.length} of {records.length} document
        {records.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No documents found</h2>
          <p>
            {records.length === 0
              ? "No documents have been shared with your company yet."
              : "Try adjusting your search or filters to find matching documents."}
          </p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">Document name</th>
                <th scope="col">Document type</th>
                <th scope="col">Candidate</th>
                <th scope="col">Modified date</th>
                <th scope="col">Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row) => (
                <tr key={row.id}>
                  <td>{cell(row.name)}</td>
                  <td>{cell(row.documentType)}</td>
                  <td>{cell(row.candidate)}</td>
                  <td>
                    {row.uploadedDate ? (
                      formatDisplayDate(row.uploadedDate)
                    ) : (
                      <span className={styles.muted}>—</span>
                    )}
                  </td>
                  <td>
                    <span className={styles.actionLinks}>
                      {row.viewPath ? (
                        <a
                          className={styles.link}
                          href={row.viewPath}
                          target="_blank"
                          rel="noreferrer"
                        >
                          View
                        </a>
                      ) : null}
                      {row.canDownload && row.downloadPath ? (
                        <a className={styles.link} href={row.downloadPath}>
                          Download
                        </a>
                      ) : null}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
