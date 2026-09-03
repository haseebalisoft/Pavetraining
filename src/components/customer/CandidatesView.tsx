"use client";

import Link from "next/link";
import { useMemo, useState, type ReactNode } from "react";

import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDate } from "@/lib/utils/formatDate";
import type { WorkforceCandidate } from "@/types/models";

import styles from "./customer.module.css";

/** Workforce row + matrix card/expiry enrichment for the customer mini-matrix list. */
export type CustomerCandidateListRow = WorkforceCandidate & {
  nporsNumber: string | null;
  nporsCategories: string | null;
  nporsExpiry: string | null;
  inHouseCourse: string | null;
  inHouseExpiry: string | null;
};

interface Props {
  companyName: string;
  candidates: CustomerCandidateListRow[];
  /** When access scope filtered the company workforce down to zero. */
  accessEmptyHint?: string | null;
}

function mutedDash() {
  return <span className={styles.muted}>—</span>;
}

function NumberExpiryCell({
  number,
  expiry,
  extra,
}: {
  number: string | null | undefined;
  expiry: string | null | undefined;
  extra?: string | null;
}) {
  const hasNumber = Boolean(number?.trim());
  const hasExtra = Boolean(extra?.trim());
  const hasExpiry = Boolean(expiry?.trim());
  if (!hasNumber && !hasExtra && !hasExpiry) return mutedDash();

  return (
    <div className={styles.cardExpiryCell}>
      {hasNumber ? <span className={styles.cardNumber}>{number}</span> : null}
      {hasExtra ? <span className={styles.cardExtra}>{extra}</span> : null}
      {hasExpiry ? <ExpiryDateBadge date={expiry} /> : mutedDash()}
    </div>
  );
}

function textOrDash(value: string | null | undefined): ReactNode {
  if (!value?.trim()) return mutedDash();
  return value;
}

export function CandidatesView({
  companyName,
  candidates,
  accessEmptyHint = null,
}: Props) {
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
        row.cscsNumber,
        row.swqrNumber,
        row.eusrNumber,
        row.nporsNumber,
        row.nporsNumbers,
        row.nporsCategories,
        row.inHouseCertificationNumber,
        row.inHouseCourse,
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
          Your workforce with key card numbers and expiries. Open a profile for
          full training records and downloadable certificates.
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
            placeholder="Search name, department, manager, card number…"
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
              ? accessEmptyHint ||
                "There are no workforce candidates for your company yet."
              : "No candidates match your search."}
          </p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table className={`${styles.dataTable} ${styles.candidatesWideTable}`}>
            <thead>
              <tr>
                <th scope="col">Candidate</th>
                <th scope="col">Department</th>
                <th scope="col">Training Manager</th>
                <th scope="col">Supervisor</th>
                <th scope="col">CSCS</th>
                <th scope="col">Streetworks (SWQR)</th>
                <th scope="col">EUSR</th>
                <th scope="col">NPORS</th>
                <th scope="col">In-House</th>
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
                  <td>{textOrDash(row.department)}</td>
                  <td>{textOrDash(row.trainingManager)}</td>
                  <td>{textOrDash(row.supervisor)}</td>
                  <td>
                    <NumberExpiryCell
                      number={row.cscsNumber}
                      expiry={row.cscsExpiry}
                    />
                  </td>
                  <td>
                    <NumberExpiryCell
                      number={row.swqrNumber}
                      expiry={row.swqrExpiry}
                    />
                  </td>
                  <td>
                    <NumberExpiryCell
                      number={row.eusrNumber}
                      expiry={row.eusrExpiry}
                    />
                  </td>
                  <td>
                    <NumberExpiryCell
                      number={row.nporsNumber ?? row.nporsNumbers}
                      expiry={row.nporsExpiry}
                      extra={row.nporsCategories}
                    />
                  </td>
                  <td>
                    <NumberExpiryCell
                      number={row.inHouseCertificationNumber}
                      expiry={row.inHouseExpiry}
                      extra={row.inHouseCourse}
                    />
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
