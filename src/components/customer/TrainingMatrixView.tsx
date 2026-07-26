"use client";

import Link from "next/link";
import { usePathname, useRouter, useSearchParams } from "next/navigation";
import { useCallback, useMemo, useTransition } from "react";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { StatusBadge } from "@/components/ui/StatusBadge";
import {
  EXPIRY_STATUS_LEGEND,
  formatDisplayDate,
  matchesAnyExpiryFilter,
  type ExpiryFilter,
} from "@/lib/training/expiryFilters";
import type { CustomerMatrixRecord } from "@/types/models";

import styles from "./customer.module.css";

type MatrixExpiryFilter =
  | "all"
  | "within-3m"
  | "within-6m"
  | "within-9m"
  | "expired"
  | "valid"
  | "missing"
  | "urgent"
  | "upcoming"
  | "review";

const TRAINING_CATEGORIES = [
  { value: "", label: "All categories" },
  { value: "npors", label: "NPORS" },
  { value: "cscs", label: "CSCS" },
  { value: "swqr", label: "SWQR / Streetworks" },
  { value: "eusr", label: "EUSR" },
  { value: "in-house", label: "In-House" },
] as const;

function cell(value: string | null | undefined) {
  if (!value?.trim()) {
    return <span className={styles.muted}>—</span>;
  }
  return value;
}

function uniqueSorted(values: Array<string | null | undefined>): string[] {
  const set = new Set<string>();
  for (const value of values) {
    const trimmed = value?.trim();
    if (trimmed) set.add(trimmed);
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b));
}

function rowExpiryDates(row: CustomerMatrixRecord): Array<string | null> {
  return [
    row.nextExpiryDate,
    row.nporsExpiry,
    row.cscsExpiry,
    row.swqrExpiry,
    row.eusrExpiry,
    row.inHouseExpiry,
    row.n001Expiry,
    row.n003Expiry,
    row.n004Expiry,
    row.n010Expiry,
    row.n020Expiry,
    row.n021Expiry,
    row.n027Expiry,
    row.n100Expiry,
  ];
}

function matchesCategory(
  row: CustomerMatrixRecord,
  category: string,
): boolean {
  if (!category) return true;
  if (category === "npors") {
    return Boolean(row.nporsCategories?.trim() || row.nporsExpiry?.trim());
  }
  if (category === "cscs") return Boolean(row.cscsExpiry?.trim());
  if (category === "swqr") return Boolean(row.swqrExpiry?.trim());
  if (category === "eusr") return Boolean(row.eusrExpiry?.trim());
  if (category === "in-house") return Boolean(row.inHouseExpiry?.trim());
  return true;
}

function matchesMatrixFilter(
  row: CustomerMatrixRecord,
  filter: MatrixExpiryFilter,
): boolean {
  if (filter === "all") return true;
  if (filter === "review") return row.needsReview;
  return matchesAnyExpiryFilter(rowExpiryDates(row), filter as ExpiryFilter);
}

function parseFilter(raw: string | null): MatrixExpiryFilter {
  if (
    raw === "within-3m" ||
    raw === "within-6m" ||
    raw === "within-9m" ||
    raw === "expired" ||
    raw === "valid" ||
    raw === "missing" ||
    raw === "urgent" ||
    raw === "upcoming" ||
    raw === "review" ||
    raw === "all"
  ) {
    return raw;
  }
  if (raw === "expiring" || raw === "expiring-3m") return "within-3m";
  if (raw === "expiring-6m") return "within-6m";
  if (raw === "expiring-9m") return "within-9m";
  return "all";
}

interface Props {
  companyName: string;
  records: CustomerMatrixRecord[];
  /** When true, this is the customer home landing page. */
  isLanding?: boolean;
}

export function TrainingMatrixView({
  companyName,
  records,
  isLanding = false,
}: Props) {
  const router = useRouter();
  const pathname = usePathname();
  const searchParams = useSearchParams();
  const [, startTransition] = useTransition();

  const search = searchParams.get("q") ?? "";
  const department = searchParams.get("department") ?? "";
  const trainingManager = searchParams.get("trainingManager") ?? "";
  const supervisor = searchParams.get("supervisor") ?? "";
  const category = searchParams.get("category") ?? "";
  const filter = parseFilter(searchParams.get("filter"));

  const setParams = useCallback(
    (patch: Record<string, string | null>) => {
      const next = new URLSearchParams(searchParams.toString());
      for (const [key, value] of Object.entries(patch)) {
        if (!value) next.delete(key);
        else next.set(key, value);
      }
      const query = next.toString();
      startTransition(() => {
        router.replace(query ? `${pathname}?${query}` : pathname, {
          scroll: false,
        });
      });
    },
    [pathname, router, searchParams],
  );

  const clearFilters = useCallback(() => {
    startTransition(() => {
      router.replace(pathname, { scroll: false });
    });
  }, [pathname, router]);

  const departments = useMemo(
    () => uniqueSorted(records.map((row) => row.department)),
    [records],
  );
  const managers = useMemo(
    () => uniqueSorted(records.map((row) => row.trainingManager)),
    [records],
  );
  const supervisors = useMemo(
    () => uniqueSorted(records.map((row) => row.supervisor)),
    [records],
  );

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    return records.filter((row) => {
      if (!matchesMatrixFilter(row, filter)) return false;
      if (
        department &&
        (row.department ?? "").trim().toLowerCase() !==
          department.trim().toLowerCase()
      ) {
        return false;
      }
      if (
        trainingManager &&
        (row.trainingManager ?? "").trim().toLowerCase() !==
          trainingManager.trim().toLowerCase()
      ) {
        return false;
      }
      if (
        supervisor &&
        (row.supervisor ?? "").trim().toLowerCase() !==
          supervisor.trim().toLowerCase()
      ) {
        return false;
      }
      if (!matchesCategory(row, category)) return false;
      if (!query) return true;
      return [
        row.candidateName,
        row.department,
        row.trainingManager,
        row.supervisor,
        row.nporsCategories,
        row.overallStatus,
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [
    records,
    search,
    filter,
    department,
    trainingManager,
    supervisor,
    category,
  ]);

  const returnQuery = searchParams.toString();
  const profileHref = (row: CustomerMatrixRecord) => {
    if (!row.candidateId) return null;
    const params = new URLSearchParams();
    if (returnQuery) params.set("return", `/customer?${returnQuery}`);
    else params.set("return", "/customer");
    const qs = params.toString();
    return `/customer/candidates/${row.candidateId}${qs ? `?${qs}` : ""}`;
  };

  const hasActiveFilters = Boolean(
    search ||
      department ||
      trainingManager ||
      supervisor ||
      category ||
      (filter && filter !== "all"),
  );

  return (
    <div>
      <header className={styles.pageHeader}>
        <Breadcrumbs
          items={
            isLanding
              ? [{ label: "Customer" }, { label: "Training Matrix" }]
              : [
                  { label: "Customer", href: "/customer" },
                  { label: "Training Matrix" },
                ]
          }
        />
        <p className={styles.eyebrow}>Customer</p>
        <h1 className={styles.title}>Training Matrix</h1>
        <p className={styles.subtitle}>
          Workforce competency overview combining candidate details with NPORS,
          CSCS, SWQR, EUSR, and In-House expiry information.
        </p>
      </header>

      <p className={styles.companyMeta}>
        Showing matrix rows for <strong>{companyName}</strong>
      </p>

      <div className={styles.expiryLegend} role="region" aria-label="Expiry colour legend">
        {EXPIRY_STATUS_LEGEND.map((item) => (
          <div key={item.status} className={styles.expiryLegendItem}>
            <StatusBadge label={item.label} tone={
              item.status === "missing"
                ? "missing"
                : item.status === "valid"
                  ? "ok"
                  : item.status === "upcoming"
                    ? "warn"
                    : "danger"
            } />
            <span className={styles.expiryLegendText}>{item.description}</span>
          </div>
        ))}
      </div>

      <div className={styles.toolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Candidate search</span>
          <input
            className={styles.input}
            type="search"
            placeholder="Search candidates…"
            value={search}
            onChange={(event) =>
              setParams({ q: event.target.value ? event.target.value : null })
            }
          />
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Department</span>
          <select
            className={styles.select}
            value={department}
            onChange={(event) =>
              setParams({ department: event.target.value || null })
            }
          >
            <option value="">All departments</option>
            {departments.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Training Manager</span>
          <select
            className={styles.select}
            value={trainingManager}
            onChange={(event) =>
              setParams({ trainingManager: event.target.value || null })
            }
          >
            <option value="">All managers</option>
            {managers.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Supervisor</span>
          <select
            className={styles.select}
            value={supervisor}
            onChange={(event) =>
              setParams({ supervisor: event.target.value || null })
            }
          >
            <option value="">All supervisors</option>
            {supervisors.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Training category</span>
          <select
            className={styles.select}
            value={category}
            onChange={(event) =>
              setParams({ category: event.target.value || null })
            }
          >
            {TRAINING_CATEGORIES.map((option) => (
              <option key={option.value || "all"} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Expiry filter</span>
          <select
            className={styles.select}
            value={filter}
            onChange={(event) =>
              setParams({
                filter:
                  event.target.value === "all" ? null : event.target.value,
              })
            }
          >
            <option value="all">All expiries</option>
            <option value="within-3m">Expiring within 3 months</option>
            <option value="within-6m">Expiring within 6 months</option>
            <option value="within-9m">Expiring within 9 months</option>
            <option value="expired">Expired</option>
            <option value="valid">Valid</option>
            <option value="missing">Records to Review</option>
            <option value="urgent">Urgent (status)</option>
            <option value="upcoming">Upcoming (status)</option>
            <option value="review">Flagged for review</option>
          </select>
        </label>
        <div className={styles.field}>
          <span className={styles.fieldLabel}>&nbsp;</span>
          <button
            type="button"
            className={styles.secondaryButton}
            disabled={!hasActiveFilters}
            onClick={clearFilters}
          >
            Clear filters
          </button>
        </div>
      </div>

      <p className={styles.resultCount}>
        {filtered.length} of {records.length} candidate
        {records.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No matrix rows</h2>
          <p>
            {records.length === 0
              ? "There are no training matrix candidates for your access yet."
              : "No rows match your current search or filters."}
          </p>
        </div>
      ) : (
        <>
          <div className={`${styles.tableWrap} ${styles.matrixDesktop}`}>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th scope="col">Candidate</th>
                  <th scope="col">Date of birth</th>
                  <th scope="col">Department</th>
                  <th scope="col">Training Manager</th>
                  <th scope="col">Supervisor</th>
                  <th scope="col">NPORS categories</th>
                  <th scope="col">NPORS expiry</th>
                  <th scope="col">CSCS expiry</th>
                  <th scope="col">SWQR expiry</th>
                  <th scope="col">EUSR expiry</th>
                  <th scope="col">In-House expiry</th>
                  <th scope="col">Next expiry</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((row) => {
                  const href = profileHref(row);
                  return (
                    <tr
                      key={row.id}
                      className={`${row.needsReview ? styles.reviewRow : ""} ${href ? styles.clickableRow : ""}`.trim()}
                      tabIndex={href ? 0 : undefined}
                      onClick={() => {
                        if (href) router.push(href);
                      }}
                      onKeyDown={(event) => {
                        if (!href) return;
                        if (event.key === "Enter" || event.key === " ") {
                          event.preventDefault();
                          router.push(href);
                        }
                      }}
                      aria-label={
                        href
                          ? `Open profile for ${row.candidateName}`
                          : undefined
                      }
                    >
                      <td>
                        {href ? (
                          <Link
                            className={styles.link}
                            href={href}
                            onClick={(event) => event.stopPropagation()}
                          >
                            {row.candidateName}
                          </Link>
                        ) : (
                          cell(row.candidateName)
                        )}
                      </td>
                      <td>{formatDisplayDate(row.dateOfBirth)}</td>
                      <td>{cell(row.department)}</td>
                      <td>{cell(row.trainingManager)}</td>
                      <td>{cell(row.supervisor)}</td>
                      <td>{cell(row.nporsCategories)}</td>
                      <td>
                        <ExpiryDateBadge date={row.nporsExpiry} />
                      </td>
                      <td>
                        <ExpiryDateBadge date={row.cscsExpiry} />
                      </td>
                      <td>
                        <ExpiryDateBadge date={row.swqrExpiry} />
                      </td>
                      <td>
                        <ExpiryDateBadge date={row.eusrExpiry} />
                      </td>
                      <td>
                        <ExpiryDateBadge date={row.inHouseExpiry} />
                      </td>
                      <td>
                        <ExpiryDateBadge date={row.nextExpiryDate} />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <ul className={styles.matrixMobileList}>
            {filtered.map((row) => {
              const href = profileHref(row);
              const card = (
                <>
                  <p className={styles.matrixCardTitle}>{row.candidateName}</p>
                  <p className={styles.matrixCardMeta}>
                    {row.department?.trim() || "No department"}
                    {row.trainingManager
                      ? ` · TM: ${row.trainingManager}`
                      : ""}
                  </p>
                  <dl className={styles.matrixCardGrid}>
                    <div>
                      <dt>Date of birth</dt>
                      <dd>{formatDisplayDate(row.dateOfBirth)}</dd>
                    </div>
                    <div>
                      <dt>Supervisor</dt>
                      <dd>{row.supervisor?.trim() || "—"}</dd>
                    </div>
                    <div>
                      <dt>NPORS</dt>
                      <dd>
                        <ExpiryDateBadge date={row.nporsExpiry} />
                      </dd>
                    </div>
                    <div>
                      <dt>CSCS</dt>
                      <dd>
                        <ExpiryDateBadge date={row.cscsExpiry} />
                      </dd>
                    </div>
                    <div>
                      <dt>SWQR</dt>
                      <dd>
                        <ExpiryDateBadge date={row.swqrExpiry} />
                      </dd>
                    </div>
                    <div>
                      <dt>EUSR</dt>
                      <dd>
                        <ExpiryDateBadge date={row.eusrExpiry} />
                      </dd>
                    </div>
                    <div>
                      <dt>In-House</dt>
                      <dd>
                        <ExpiryDateBadge date={row.inHouseExpiry} />
                      </dd>
                    </div>
                    <div>
                      <dt>Next expiry</dt>
                      <dd>
                        <ExpiryDateBadge date={row.nextExpiryDate} />
                      </dd>
                    </div>
                  </dl>
                  {row.nporsCategories ? (
                    <p className={styles.matrixCardCats}>
                      Categories: {row.nporsCategories}
                    </p>
                  ) : null}
                </>
              );

              return (
                <li key={row.id} className={styles.matrixCard}>
                  {href ? (
                    <Link className={styles.matrixCardLink} href={href}>
                      {card}
                    </Link>
                  ) : (
                    card
                  )}
                </li>
              );
            })}
          </ul>
        </>
      )}
    </div>
  );
}
