"use client";

import Link from "next/link";
import { useMemo, useState } from "react";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDate } from "@/lib/utils/formatDate";
import type { WorkforceCandidate } from "@/types/models";

import styles from "./customer.module.css";

interface Props {
  companyName: string;
  candidates: WorkforceCandidate[];
}

export function CandidatesView({ companyName, candidates }: Props) {
  const [search, setSearch] = useState("");

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) return candidates;
    return candidates.filter((row) =>
      [
        row.candidateName,
        row.companyName,
        row.department,
        row.workforceNumber,
        row.status,
        row.trainingManager,
        row.supervisor,
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query),
    );
  }, [candidates, search]);

  return (
    <div>
      <header className={styles.pageHeader}>
        <Breadcrumbs
          items={[
            { label: "Customer", href: "/customer" },
            { label: "Candidates" },
          ]}
        />
        <p className={styles.eyebrow}>Customer</p>
        <h1 className={styles.title}>Candidates</h1>
        <p className={styles.subtitle}>
          People in your organisation. Open a profile for workforce details.
        </p>
      </header>

      <p className={styles.companyMeta}>
        Showing candidates for <strong>{companyName}</strong>
      </p>

      <div className={styles.toolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            placeholder="Search name, department, manager…"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
      </div>

      <p className={styles.resultCount}>
        {filtered.length} of {candidates.length} candidate
        {candidates.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No candidates found</h2>
          <p>
            {candidates.length === 0
              ? "There are no workforce candidates for your company yet."
              : "No candidates match your search."}
          </p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">Candidate</th>
                <th scope="col">Company</th>
                <th scope="col">Training Manager</th>
                <th scope="col">Supervisor</th>
                <th scope="col">Status</th>
                <th scope="col">Profile</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row) => (
                <tr key={row.id}>
                  <td>
                    <div className={styles.candidateNameCell}>
                      <span>{row.candidateName}</span>
                      {row.dateOfBirth?.trim() ? (
                        <span className={styles.dobSecondary}>
                          DOB {formatDate(row.dateOfBirth)}
                        </span>
                      ) : null}
                    </div>
                  </td>
                  <td>{row.companyName?.trim() || companyName}</td>
                  <td>
                    {row.trainingManager?.trim() ? (
                      row.trainingManager
                    ) : (
                      <span className={styles.muted}>—</span>
                    )}
                  </td>
                  <td>
                    {row.supervisor?.trim() ? (
                      row.supervisor
                    ) : (
                      <span className={styles.muted}>—</span>
                    )}
                  </td>
                  <td>
                    <StatusBadge
                      label={row.status?.trim() || "Unknown"}
                      tone={
                        (row.status ?? "").toLowerCase() === "active"
                          ? "ok"
                          : "neutral"
                      }
                    />
                  </td>
                  <td>
                    <Link
                      className={styles.link}
                      href={`/customer/candidates/${row.id}`}
                    >
                      View Profile
                    </Link>
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
