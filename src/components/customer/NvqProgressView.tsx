"use client";

import { useMemo, useState } from "react";

import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type { CustomerNvqRecord } from "@/types/models";

import styles from "./customer.module.css";

type NvqTab = "Active" | "Completed";

interface Props {
  companyName: string;
  records: CustomerNvqRecord[];
}

function cell(value: string | null | undefined) {
  if (!value?.trim()) {
    return <span className={styles.muted}>—</span>;
  }
  return value;
}

function dateCell(value: string | null | undefined) {
  if (!value?.trim()) {
    return <span className={styles.muted}>—</span>;
  }
  return formatDisplayDate(value);
}

export function NvqProgressView({ companyName, records }: Props) {
  const [tab, setTab] = useState<NvqTab>("Active");
  const [search, setSearch] = useState("");

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    return records.filter((row) => {
      if (row.status !== tab) {
        return false;
      }
      if (!query) {
        return true;
      }
      return [
        row.candidateName,
        row.nvqTitle,
        row.boltOn,
        row.stageOfNvq,
        row.notes,
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [records, tab, search]);

  const activeCount = records.filter((row) => row.status === "Active").length;
  const completedCount = records.filter(
    (row) => row.status === "Completed",
  ).length;

  return (
    <div>
      <CustomerPageHeader
        breadcrumbs={[
          { label: "Customer", href: "/customer" },
          { label: "NVQ Progress" },
        ]}
        title="NVQ Progress"
        subtitle="Track active and completed NVQ programmes for your workforce."
      />

      <p className={styles.companyMeta}>
        Showing records for <strong>{companyName}</strong>
      </p>

      <div className={styles.tabs} role="tablist" aria-label="NVQ status">
        <button
          type="button"
          role="tab"
          aria-selected={tab === "Active"}
          className={`${styles.tab} ${tab === "Active" ? styles.tabActive : ""}`}
          onClick={() => setTab("Active")}
        >
          Active ({activeCount})
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={tab === "Completed"}
          className={`${styles.tab} ${tab === "Completed" ? styles.tabActive : ""}`}
          onClick={() => setTab("Completed")}
        >
          Completed ({completedCount})
        </button>
      </div>

      <div className={styles.toolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            placeholder="Search candidates, titles, stages…"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
      </div>

      <p className={styles.resultCount}>
        {filtered.length} {tab.toLowerCase()} record
        {filtered.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No {tab.toLowerCase()} NVQs</h2>
          <p>
            {records.length === 0
              ? "There are no customer-visible NVQ records for your company yet."
              : `No ${tab.toLowerCase()} NVQ records match your search.`}
          </p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">Candidate Name</th>
                <th scope="col">NVQ Title</th>
                <th scope="col">Status</th>
                <th scope="col">Bolt On</th>
                <th scope="col">Date Registered</th>
                <th scope="col">Induction Date</th>
                <th scope="col">Stage of NVQ</th>
                <th scope="col">Notes</th>
                <th scope="col">Completed Date</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row) => (
                <tr
                  key={row.id}
                  className={
                    row.status === "Completed" ? styles.completedRow : undefined
                  }
                >
                  <td>{cell(row.candidateName)}</td>
                  <td>{cell(row.nvqTitle)}</td>
                  <td>
                    <StatusBadge
                      label={row.status}
                      tone={row.status === "Completed" ? "ok" : "info"}
                    />
                  </td>
                  <td>{cell(row.boltOn)}</td>
                  <td>{dateCell(row.dateRegistered)}</td>
                  <td>{dateCell(row.inductionDate)}</td>
                  <td>{cell(row.stageOfNvq)}</td>
                  <td>{cell(row.notes)}</td>
                  <td>{dateCell(row.completedDate)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
