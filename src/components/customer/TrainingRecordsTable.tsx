"use client";

import Link from "next/link";
import { useMemo, useState, type ReactNode } from "react";

import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { StatusBadge } from "@/components/ui/StatusBadge";
import {
  formatDisplayDate,
  matchesExpiryFilter,
  type ExpiryFilter,
} from "@/lib/training/expiryFilters";
import type { CustomerOutcome } from "@/types/models";

import styles from "./trainingRecords.module.css";
import customerStyles from "./customer.module.css";

export type OutcomeFilter = "all" | CustomerOutcome;

export interface TrainingRecordColumn<T> {
  key: string;
  header: string;
  render: (row: T) => ReactNode;
}

interface TrainingRecordsTableProps<T extends { id: string }> {
  title: string;
  description: string;
  companyName: string;
  records: T[];
  columns: TrainingRecordColumn<T>[];
  getSearchText: (row: T) => string;
  getOutcome: (row: T) => CustomerOutcome | null;
  getExpiry: (row: T) => string | null;
  getWorkforceId: (row: T) => string | null;
}

const EXPIRY_OPTIONS: { value: ExpiryFilter; label: string }[] = [
  { value: "all", label: "All expiries" },
  { value: "within-3m", label: "Expiring within 3 months" },
  { value: "within-6m", label: "Expiring within 6 months" },
  { value: "within-9m", label: "Expiring within 9 months" },
  { value: "expired", label: "Expired" },
  { value: "urgent", label: "Urgent (0–90 days)" },
  { value: "upcoming", label: "Upcoming (91–270 days)" },
  { value: "valid", label: "Valid (271+ days)" },
  { value: "missing", label: "Records to Review" },
];

function displayCell(value: string | null | undefined): ReactNode {
  if (!value?.trim()) {
    return <span className={styles.muted}>—</span>;
  }
  return value;
}

export function formatExpiryCell(expiry: string | null): ReactNode {
  return <ExpiryDateBadge date={expiry} />;
}

export function formatOutcomeCell(outcome: CustomerOutcome | null): ReactNode {
  if (!outcome) {
    return <span className={styles.muted}>—</span>;
  }

  return (
    <StatusBadge
      label={outcome}
      tone={outcome === "Pass" ? "ok" : "danger"}
    />
  );
}

export function formatTextCell(value: string | null | undefined): ReactNode {
  return displayCell(value);
}

export function formatDateCell(value: string | null | undefined): ReactNode {
  if (!value?.trim()) {
    return <span className={styles.muted}>—</span>;
  }
  return formatDisplayDate(value);
}

export function TrainingRecordsTable<T extends { id: string }>({
  title,
  description,
  companyName,
  records,
  columns,
  getSearchText,
  getOutcome,
  getExpiry,
  getWorkforceId,
}: TrainingRecordsTableProps<T>) {
  const [search, setSearch] = useState("");
  const [outcomeFilter, setOutcomeFilter] = useState<OutcomeFilter>("all");
  const [expiryFilter, setExpiryFilter] = useState<ExpiryFilter>("all");

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();

    return records.filter((row) => {
      if (query && !getSearchText(row).toLowerCase().includes(query)) {
        return false;
      }

      const outcome = getOutcome(row);
      if (outcomeFilter !== "all" && outcome !== outcomeFilter) {
        return false;
      }

      if (!matchesExpiryFilter(getExpiry(row), expiryFilter)) {
        return false;
      }

      return true;
    });
  }, [
    records,
    search,
    outcomeFilter,
    expiryFilter,
    getSearchText,
    getOutcome,
    getExpiry,
  ]);

  return (
    <div>
      <CustomerPageHeader
        breadcrumbs={[
          { label: "Customer", href: "/customer" },
          { label: "Training Records", href: "/customer/training-records" },
          { label: title },
        ]}
        eyebrow="Training Records"
        title={title}
        subtitle={description}
      />

      <p className={customerStyles.companyMeta}>
        Showing records for <strong>{companyName}</strong>
      </p>

      <div className={styles.toolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            placeholder="Search candidates and details…"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>

        <label className={styles.field}>
          <span className={styles.fieldLabel}>Outcome</span>
          <select
            className={styles.select}
            value={outcomeFilter}
            onChange={(event) =>
              setOutcomeFilter(event.target.value as OutcomeFilter)
            }
          >
            <option value="all">All outcomes</option>
            <option value="Pass">Pass</option>
            <option value="Fail">Fail</option>
          </select>
        </label>

        <label className={styles.field}>
          <span className={styles.fieldLabel}>Expiry</span>
          <select
            className={styles.select}
            value={expiryFilter}
            onChange={(event) =>
              setExpiryFilter(event.target.value as ExpiryFilter)
            }
          >
            {EXPIRY_OPTIONS.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
      </div>

      <p className={styles.resultCount}>
        {filtered.length} of {records.length} record
        {records.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.empty}>
          <h2>No records found</h2>
          <p>
            {records.length === 0
              ? "There are no customer-visible records for your company in this register yet."
              : "Try adjusting search or filters to see matching training records."}
          </p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                {columns.map((column) => (
                  <th key={column.key} scope="col">
                    {column.header}
                  </th>
                ))}
                <th scope="col">Profile</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row) => {
                const workforceId = getWorkforceId(row);
                return (
                  <tr key={row.id}>
                    {columns.map((column) => (
                      <td key={column.key}>{column.render(row)}</td>
                    ))}
                    <td>
                      {workforceId ? (
                        <Link
                          className={styles.profileLink}
                          href={`/customer/candidates/${workforceId}`}
                        >
                          View profile
                        </Link>
                      ) : (
                        <span className={styles.profileDisabled}>
                          Unavailable
                        </span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
