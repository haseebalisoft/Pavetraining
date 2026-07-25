"use client";

import Link from "next/link";
import { useMemo, useState } from "react";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import {
  formatDisplayDate,
  daysUntilExpiry,
} from "@/lib/training/expiryFilters";
import { toneForExpiry } from "@/lib/ui/status";
import type { CustomerMatrixRecord } from "@/types/models";

import styles from "./customer.module.css";

type MatrixFilter = "all" | "review" | "expired" | "expiring";

interface Props {
  companyName: string;
  records: CustomerMatrixRecord[];
  initialFilter?: MatrixFilter;
}

function cell(value: string | null | undefined) {
  if (!value?.trim()) {
    return <span className={styles.muted}>—</span>;
  }
  return value;
}

function expiryCell(value: string | null | undefined) {
  if (!value?.trim()) {
    return <StatusBadge label="Missing" tone="missing" />;
  }
  return (
    <StatusBadge
      label={formatDisplayDate(value) ?? value}
      tone={toneForExpiry(value)}
    />
  );
}

export function TrainingMatrixView({
  companyName,
  records,
  initialFilter = "all",
}: Props) {
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<MatrixFilter>(initialFilter);

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    return records.filter((row) => {
      if (filter === "review" && !row.needsReview) return false;
      if (filter === "expired") {
        const days = daysUntilExpiry(row.nextExpiryDate);
        if (days === null || days >= 0) return false;
      }
      if (filter === "expiring") {
        const days = daysUntilExpiry(row.nextExpiryDate);
        if (days === null || days < 0 || days > 90) return false;
      }
      if (!query) return true;
      return [row.candidateName, row.department, row.overallStatus]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [records, search, filter]);

  return (
    <div>
      <header className={styles.pageHeader}>
        <Breadcrumbs
          items={[
            { label: "Customer", href: "/customer" },
            { label: "Training Matrix" },
          ]}
        />
        <p className={styles.eyebrow}>Customer</p>
        <h1 className={styles.title}>Training Matrix</h1>
        <p className={styles.subtitle}>
          Competency overview for your workforce, including next expiry and
          Records to Review.
        </p>
      </header>

      <p className={styles.companyMeta}>
        Showing matrix rows for <strong>{companyName}</strong>
      </p>

      <div className={styles.toolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            placeholder="Search candidates, departments…"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Filter</span>
          <select
            className={styles.select}
            value={filter}
            onChange={(event) =>
              setFilter(event.target.value as MatrixFilter)
            }
          >
            <option value="all">All rows</option>
            <option value="review">Records to Review</option>
            <option value="expired">Expired</option>
            <option value="expiring">Expiring in 3 months</option>
          </select>
        </label>
      </div>

      <p className={styles.resultCount}>
        {filtered.length} of {records.length} row
        {records.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No matrix rows</h2>
          <p>
            {records.length === 0
              ? "There are no training matrix rows for your company yet."
              : "No rows match your current search or filter."}
          </p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">Candidate</th>
                <th scope="col">Department</th>
                <th scope="col">Next expiry</th>
                <th scope="col">Records to Review</th>
                <th scope="col">Overall status</th>
                <th scope="col">N001</th>
                <th scope="col">N010</th>
                <th scope="col">N020</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row) => (
                <tr
                  key={row.id}
                  className={row.needsReview ? styles.reviewRow : undefined}
                >
                  <td>{cell(row.candidateName)}</td>
                  <td>{cell(row.department)}</td>
                  <td>{expiryCell(row.nextExpiryDate)}</td>
                  <td>
                    <StatusBadge
                      label={row.needsReview ? "Review" : "Clear"}
                      tone={row.needsReview ? "warn" : "ok"}
                    />
                  </td>
                  <td>{cell(row.overallStatus)}</td>
                  <td>{expiryCell(row.n001Expiry)}</td>
                  <td>{expiryCell(row.n010Expiry)}</td>
                  <td>{expiryCell(row.n020Expiry)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <p className={styles.companyMeta} style={{ marginTop: "1rem" }}>
        Looking for a person?{" "}
        <Link className={styles.link} href="/customer/candidates">
          Open candidates
        </Link>
      </p>
    </div>
  );
}
