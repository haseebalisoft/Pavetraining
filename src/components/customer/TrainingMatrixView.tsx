"use client";

import Link from "next/link";
import { usePathname, useRouter, useSearchParams } from "next/navigation";
import { useCallback, useMemo, useRef, useState, useTransition } from "react";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { SlideOverPanel } from "@/components/ui/SlideOverPanel";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { LiveTrainingRefresh } from "@/components/training/LiveTrainingRefresh";
import {
  EXPIRY_STATUS_LEGEND,
  getExpiryStatus,
  matchesAnyExpiryFilter,
  type ExpiryFilter,
} from "@/lib/training/expiryFilters";
import { formatDate } from "@/lib/utils/formatDate";
import type { CustomerMatrixRecord } from "@/types/models";
import type { StatusTone } from "@/lib/ui/status";

import styles from "./customer.module.css";

type MatrixExpiryFilter =
  | "all"
  | "within-3m"
  | "within-6m"
  | "6m-plus"
  | "9m-plus"
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
  { value: "swqr", label: "Streetworks (SWQR)" },
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
    ...(row.eusrCategoryRows ?? []).map((cell) => cell.expiry),
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
    return Boolean(
      row.nporsCategories?.trim() ||
        row.nporsExpiry?.trim() ||
        row.nporsNumber?.trim(),
    );
  }
  if (category === "cscs") {
    return Boolean(row.cscsExpiry?.trim() || row.cscsNumber?.trim());
  }
  if (category === "swqr") {
    return Boolean(row.swqrExpiry?.trim() || row.swqrNumber?.trim());
  }
  if (category === "eusr") {
    return Boolean(
      row.eusrExpiry?.trim() ||
        row.eusrNumber?.trim() ||
        (row.eusrCategoryRows ?? []).some(
          (cell) => cell.expiry?.trim() || cell.trainingDate?.trim(),
        ),
    );
  }
  if (category === "in-house") {
    return Boolean(row.inHouseExpiry?.trim() || row.inHouseCourse?.trim());
  }
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

function overallTone(row: CustomerMatrixRecord): StatusTone {
  if (row.needsReview) return "missing";
  const status = getExpiryStatus(row.nextExpiryDate).status;
  if (status === "missing") return "missing";
  if (status === "valid") return "ok";
  if (status === "upcoming") return "warn";
  return "danger";
}

function CertCell({
  number,
  expiry,
}: {
  number: string | null;
  expiry: string | null;
}) {
  return (
    <div className={styles.certCell}>
      {number?.trim() ? (
        <span className={styles.certNumber}>{number}</span>
      ) : (
        <span className={styles.muted}>—</span>
      )}
      <ExpiryDateBadge date={expiry} />
    </div>
  );
}

function EusrCertCell({ row }: { row: CustomerMatrixRecord }) {
  const categories = row.eusrCategoryRows ?? [];
  if (categories.length > 0) {
    return (
      <div className={styles.eusrCategoryStack}>
        {row.eusrNumber?.trim() ? (
          <span className={styles.certNumber}>{row.eusrNumber}</span>
        ) : null}
        {categories.map((cell) => (
          <div key={cell.category} className={styles.eusrCategoryItem}>
            <span className={styles.eusrCategoryName}>{cell.category}</span>
            {cell.trainingDate?.trim() ? (
              <span className={styles.eusrCategoryDates}>
                Trained {formatDate(cell.trainingDate)}
              </span>
            ) : null}
            <ExpiryDateBadge date={cell.expiry} />
          </div>
        ))}
      </div>
    );
  }
  return <CertCell number={row.eusrNumber} expiry={row.eusrExpiry} />;
}

function CandidateNameCell({
  row,
  href,
}: {
  row: CustomerMatrixRecord;
  href: string | null;
}) {
  const name = href ? (
    <Link
      className={styles.link}
      href={href}
      onClick={(event) => event.stopPropagation()}
    >
      {row.candidateName}
    </Link>
  ) : (
    cell(row.candidateName)
  );

  return (
    <div className={styles.candidateNameCell}>
      {name}
      {row.dateOfBirth?.trim() ? (
        <span className={styles.dobSecondary}>
          DOB {formatDate(row.dateOfBirth)}
        </span>
      ) : null}
    </div>
  );
}

function parseFilter(raw: string | null): MatrixExpiryFilter {
  if (
    raw === "within-3m" ||
    raw === "within-6m" ||
    raw === "6m-plus" ||
    raw === "9m-plus" ||
    raw === "expired" ||
    raw === "valid" ||
    raw === "missing" ||
    raw === "urgent" ||
    raw === "upcoming" ||
    raw === "review" ||
    raw === "all"
  ) {
    return raw === "9m-plus" ? "6m-plus" : raw;
  }
  if (raw === "expiring" || raw === "expiring-3m") return "within-3m";
  if (raw === "expiring-6m") return "within-6m";
  if (raw === "within-9m" || raw === "expiring-9m") return "6m-plus";
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
        row.companyName,
        row.department,
        row.trainingManager,
        row.supervisor,
        row.nporsCategories,
        row.nporsNumber,
        row.cscsNumber,
        row.swqrNumber,
        row.eusrNumber,
        row.inHouseCourse,
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
    if (returnQuery) {
      params.set("return", `/customer?${returnQuery}`);
    } else {
      params.set("return", "/customer");
    }
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

  /** Advanced filters only (excludes search) — drives the Filters (N) badge. */
  const advancedFilterCount = [
    department,
    trainingManager,
    supervisor,
    category,
    filter !== "all" ? filter : "",
  ].filter(Boolean).length;

  const [filtersOpen, setFiltersOpen] = useState(false);
  const [legendOpen, setLegendOpen] = useState(false);
  const filtersTriggerRef = useRef<HTMLButtonElement>(null);

  const filterFields = (
    <>
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
              filter: event.target.value === "all" ? null : event.target.value,
            })
          }
        >
          <option value="all">All expiries</option>
          <option value="expired">Expired</option>
          <option value="within-3m">Expiring within 3 months</option>
          <option value="within-6m">Expiring within 6 months</option>
          <option value="6m-plus">6–9 months and beyond / in date</option>
          <option value="review">Records to Review</option>
        </select>
      </label>
    </>
  );

  return (
    <div className={styles.matrixPage}>
      <LiveTrainingRefresh />
      <header className={`${styles.pageHeader} ${styles.matrixHeader}`}>
        <div className={styles.matrixBreadcrumbs}>
          <Breadcrumbs
            items={[
              { label: "Customer", href: "/customer" },
              { label: "Training Matrix" },
            ]}
          />
        </div>
        <div className={styles.matrixTitleRow}>
          <p className={styles.eyebrow}>Customer</p>
          <h1 className={styles.title}>Training Matrix</h1>
        </div>
        <p className={`${styles.subtitle} ${styles.matrixSubtitleDesktop}`}>
          Workforce competency overview for your company — NPORS, CSCS,
          Streetworks (SWQR), EUSR, and In-House Training.
        </p>
      </header>

      <p className={styles.companyMeta}>
        Showing matrix rows for <strong>{companyName}</strong>
      </p>

      <div className={styles.matrixActions}>
        <button
          type="button"
          className={styles.secondaryButton}
          disabled
          title="PDF Snapshot export is coming soon. Spreadsheet export is not available for customers."
        >
          PDF Snapshot (coming soon)
        </button>
      </div>

      <div className={styles.expiryLegendDesktop} role="region" aria-label="Expiry colour legend">
        {EXPIRY_STATUS_LEGEND.map((item) => (
          <div key={item.status} className={styles.expiryLegendItem}>
            <StatusBadge
              label={item.label}
              tone={
                item.status === "missing"
                  ? "missing"
                  : item.status === "valid"
                    ? "ok"
                    : item.status === "upcoming"
                      ? "warn"
                      : "danger"
              }
            />
            <span className={styles.expiryLegendText}>{item.description}</span>
          </div>
        ))}
      </div>

      <div className={styles.legendMobile}>
        <button
          type="button"
          className={styles.legendToggle}
          aria-expanded={legendOpen}
          onClick={() => setLegendOpen((value) => !value)}
        >
          What do the colours mean?
          <span aria-hidden="true">{legendOpen ? "▴" : "▾"}</span>
        </button>
        {legendOpen ? (
          <div className={styles.expiryLegend} role="region" aria-label="Expiry colour legend">
            {EXPIRY_STATUS_LEGEND.map((item) => (
              <div key={item.status} className={styles.expiryLegendItem}>
                <StatusBadge
                  label={item.label}
                  tone={
                    item.status === "missing"
                      ? "missing"
                      : item.status === "valid"
                        ? "ok"
                        : item.status === "upcoming"
                          ? "warn"
                          : "danger"
                  }
                />
                <span className={styles.expiryLegendText}>{item.description}</span>
              </div>
            ))}
          </div>
        ) : null}
      </div>

      <div className={styles.filterPanel}>
        <div className={styles.mobileFilterBar}>
          <label className={`${styles.field} ${styles.mobileSearchField}`}>
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
          <button
            ref={filtersTriggerRef}
            type="button"
            className={styles.filtersChip}
            aria-haspopup="dialog"
            aria-expanded={filtersOpen}
            onClick={() => setFiltersOpen(true)}
          >
            Filters
            {advancedFilterCount > 0 ? (
              <span className={styles.filtersChipCount}>{advancedFilterCount}</span>
            ) : null}
          </button>
        </div>

        <div className={styles.toolbarDesktop}>
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
          {filterFields}
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
      </div>

      <SlideOverPanel
        open={filtersOpen}
        title="Filters"
        onClose={() => setFiltersOpen(false)}
        returnFocusRef={filtersTriggerRef}
        footer={
          <div className={styles.filterSheetActions}>
            <button
              type="button"
              className={styles.secondaryButton}
              disabled={!hasActiveFilters}
              onClick={clearFilters}
            >
              Clear filters
            </button>
            <button
              type="button"
              className={styles.primaryButton}
              onClick={() => setFiltersOpen(false)}
            >
              Apply
            </button>
          </div>
        }
      >
        <div className={styles.filterSheetBody}>{filterFields}</div>
      </SlideOverPanel>

      <p className={styles.resultCount}>
        {filtered.length} of {records.length} candidate
        {records.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <span className={styles.emptyStateIcon} aria-hidden="true">
            <svg viewBox="0 0 24 24" fill="none">
              <rect x="3" y="4" width="18" height="16" rx="2" />
              <path d="M3 9h18M9 9v11M15 9v11" />
            </svg>
          </span>
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
                  <th scope="col">Company</th>
                  <th scope="col">Training Manager</th>
                  <th scope="col">Supervisor</th>
                  <th scope="col">CSCS</th>
                  <th scope="col">Streetworks (SWQR)</th>
                  <th scope="col">EUSR</th>
                  <th scope="col">NPORS</th>
                  <th scope="col">In-House</th>
                  <th scope="col">Next expiry</th>
                  <th scope="col">Overall status</th>
                  <th scope="col">Records to Review</th>
                  <th scope="col">Profile</th>
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
                        <CandidateNameCell row={row} href={href} />
                      </td>
                      <td>{cell(row.companyName)}</td>
                      <td>{cell(row.trainingManager)}</td>
                      <td>{cell(row.supervisor)}</td>
                      <td>
                        <CertCell
                          number={row.cscsNumber}
                          expiry={row.cscsExpiry}
                        />
                      </td>
                      <td>
                        <CertCell
                          number={row.swqrNumber}
                          expiry={row.swqrExpiry}
                        />
                      </td>
                      <td>
                        <EusrCertCell row={row} />
                      </td>
                      <td>
                        <div className={styles.certCell}>
                          {row.nporsNumber?.trim() ||
                          row.nporsCategories?.trim() ? (
                            <span className={styles.certNumber}>
                              {[row.nporsNumber, row.nporsCategories]
                                .filter((v) => v?.trim())
                                .join(" · ")}
                            </span>
                          ) : (
                            <span className={styles.muted}>—</span>
                          )}
                          <ExpiryDateBadge date={row.nporsExpiry} />
                        </div>
                      </td>
                      <td>
                        <div className={styles.certCell}>
                          {row.inHouseCourse?.trim() ? (
                            <span className={styles.certNumber}>
                              {row.inHouseCourse}
                            </span>
                          ) : (
                            <span className={styles.muted}>—</span>
                          )}
                          <ExpiryDateBadge date={row.inHouseExpiry} />
                        </div>
                      </td>
                      <td>
                        <ExpiryDateBadge date={row.nextExpiryDate} />
                      </td>
                      <td>
                        <StatusBadge
                          label={row.overallStatus ?? "Records to Review"}
                          tone={overallTone(row)}
                        />
                      </td>
                      <td>
                        {row.needsReview ? (
                          <StatusBadge
                            label="Records to Review"
                            tone="missing"
                          />
                        ) : (
                          <span className={styles.muted}>—</span>
                        )}
                      </td>
                      <td>
                        {href ? (
                          <Link
                            className={styles.link}
                            href={href}
                            onClick={(event) => event.stopPropagation()}
                          >
                            View Profile
                          </Link>
                        ) : (
                          <span className={styles.muted}>—</span>
                        )}
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
                  {row.dateOfBirth?.trim() ? (
                    <p className={styles.dobSecondary}>
                      DOB {formatDate(row.dateOfBirth)}
                    </p>
                  ) : null}
                  <p className={styles.matrixCardMeta}>
                    {row.companyName?.trim() || companyName}
                    {row.trainingManager
                      ? ` · TM: ${row.trainingManager}`
                      : ""}
                    {row.supervisor ? ` · Sup: ${row.supervisor}` : ""}
                  </p>
                  <dl className={styles.matrixCardGrid}>
                    <div>
                      <dt>Overall status</dt>
                      <dd>
                        <StatusBadge
                          label={row.overallStatus ?? "Records to Review"}
                          tone={overallTone(row)}
                        />
                      </dd>
                    </div>
                    <div>
                      <dt>CSCS</dt>
                      <dd>
                        <CertCell
                          number={row.cscsNumber}
                          expiry={row.cscsExpiry}
                        />
                      </dd>
                    </div>
                    <div>
                      <dt>Streetworks</dt>
                      <dd>
                        <CertCell
                          number={row.swqrNumber}
                          expiry={row.swqrExpiry}
                        />
                      </dd>
                    </div>
                    <div>
                      <dt>EUSR</dt>
                      <dd>
                        <EusrCertCell row={row} />
                      </dd>
                    </div>
                    <div>
                      <dt>NPORS</dt>
                      <dd>
                        <CertCell
                          number={
                            [row.nporsNumber, row.nporsCategories]
                              .filter((v) => v?.trim())
                              .join(" · ") || null
                          }
                          expiry={row.nporsExpiry}
                        />
                      </dd>
                    </div>
                    <div>
                      <dt>In-House</dt>
                      <dd>
                        <CertCell
                          number={row.inHouseCourse}
                          expiry={row.inHouseExpiry}
                        />
                      </dd>
                    </div>
                    <div>
                      <dt>Next expiry</dt>
                      <dd>
                        <ExpiryDateBadge date={row.nextExpiryDate} />
                      </dd>
                    </div>
                  </dl>
                  {href ? (
                    <p className={styles.matrixCardCats}>View Profile</p>
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
