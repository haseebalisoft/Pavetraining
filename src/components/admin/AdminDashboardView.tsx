import Link from "next/link";

import { AdminCompanySelector } from "@/components/admin/AdminCompanySelector";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import type {
  AdminDashboardPayload,
  Company,
} from "@/types/models";

import styles from "./admin.module.css";

const QUICK_ACTIONS = [
  {
    href: "/admin/companies?action=add",
    title: "Add company",
    description: "Create a new company record in SharePoint.",
  },
  {
    href: "/admin/workforce?action=add",
    title: "Add candidate",
    description: "Register a workforce candidate.",
  },
  {
    href: "/admin/documents?action=add",
    title: "Upload document",
    description: "Add a customer-visible document.",
  },
  {
    href: "/admin/events?action=add",
    title: "Create event",
    description: "Schedule a training event.",
  },
  {
    href: "/admin/permissions",
    title: "Manage permissions",
    description: "Control portal access and download rights.",
  },
  {
    href: "/admin/training-matrix?filter=review",
    title: "Review training matrix",
    description: "Open Records to Review.",
  },
] as const;

interface AdminDashboardViewProps {
  email: string;
  companies: Company[];
  dashboard: AdminDashboardPayload;
}

export function AdminDashboardView({
  email,
  companies,
  dashboard,
}: AdminDashboardViewProps) {
  const { stats, warnings, selectedCompanyId, selectedCompanyName } = dashboard;

  const cards: Array<{
    label: string;
    value: number;
    tone?: "warn" | "ok";
  }> = [
    { label: "Total companies", value: stats.totalCompanies },
    { label: "Active companies", value: stats.activeCompanies, tone: "ok" },
    { label: "Total candidates", value: stats.totalCandidates },
    { label: "Expired training", value: stats.expiredTraining, tone: "warn" },
    {
      label: "Expiring within 3 months",
      value: stats.expiringWithin3Months,
      tone: "warn",
    },
    {
      label: "Records to Review",
      value: stats.recordsToReview,
      tone: "warn",
    },
    { label: "Active NVQs", value: stats.activeNvqs },
    { label: "Completed NVQs", value: stats.completedNvqs, tone: "ok" },
    {
      label: "Documents pending visibility",
      value: stats.documentsPendingVisibility,
      tone: "warn",
    },
    { label: "Upcoming events", value: stats.upcomingEvents },
  ];

  const attention = [
    {
      href: "/admin/training-matrix?filter=expired",
      title: "Expired training",
      value: stats.expiredTraining,
      detail: "Matrix rows past next expiry",
    },
    {
      href: "/admin/training-matrix?filter=expiring",
      title: "Expiring within 3 months",
      value: stats.expiringWithin3Months,
      detail: "Renewals due in the next 90 days",
    },
    {
      href: "/admin/training-matrix?filter=review",
      title: "Records to Review",
      value: stats.recordsToReview,
      detail: "Matrix rows flagged for attention",
    },
  ].filter((item) => item.value > 0);

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs items={[{ label: "Admin" }, { label: "Dashboard" }]} />
          <p className={styles.eyebrow}>Admin · {email}</p>
          <h1 className={styles.title}>Operations dashboard</h1>
          <p className={styles.subtitle}>
            Cross-company training health from SharePoint. Filter by company to
            focus operational metrics.
          </p>
        </div>
      </header>

      <div className={styles.toolbar}>
        <AdminCompanySelector
          companies={companies}
          selectedCompanyId={selectedCompanyId}
        />
      </div>

      {selectedCompanyName ? (
        <p className={styles.filterChip}>
          Filtered to <strong>{selectedCompanyName}</strong>
          <Link className={styles.clearFilter} href="/admin">
            Clear
          </Link>
        </p>
      ) : null}

      {attention.length > 0 ? (
        <section className={styles.attentionGrid} aria-label="Expiry warnings">
          {attention.map((item) => (
            <Link key={item.href} href={item.href} className={styles.attentionCard}>
              <strong>
                {item.title} · {item.value}
              </strong>
              <span>{item.detail}</span>
            </Link>
          ))}
        </section>
      ) : null}

      <section className={styles.statsGrid} aria-label="Dashboard statistics">
        {cards.map((card) => (
          <article
            key={card.label}
            className={`${styles.statCard} ${
              card.tone === "warn"
                ? styles.statWarn
                : card.tone === "ok"
                  ? styles.statOk
                  : ""
            }`}
          >
            <p className={styles.statLabel}>{card.label}</p>
            <p className={styles.statValue}>{card.value}</p>
          </article>
        ))}
      </section>

      <div className={styles.sectionGrid}>
        <section className={styles.panel} aria-label="Quick actions">
          <h2 className={styles.panelTitle}>Quick actions</h2>
          <div className={styles.actionsGrid}>
            {QUICK_ACTIONS.map((action) => (
              <Link
                key={action.href}
                className={styles.actionLink}
                href={action.href}
              >
                <strong>{action.title}</strong>
                <span>{action.description}</span>
              </Link>
            ))}
          </div>
        </section>

        <section className={styles.panel} aria-label="Data quality warnings">
          <h2 className={styles.panelTitle}>Data quality warnings</h2>
          {warnings.length === 0 ? (
            <p className={styles.emptyNote}>
              No missing company, visibility, or training address issues
              detected in the current scope.
            </p>
          ) : (
            <div className={styles.warningList}>
              {warnings.map((warning) => (
                <article key={warning.id} className={styles.warningItem}>
                  <strong>
                    {warning.source}
                    {warning.candidateName ? ` · ${warning.candidateName}` : ""}
                  </strong>
                  <p>{warning.detail}</p>
                </article>
              ))}
            </div>
          )}
        </section>
      </div>
    </div>
  );
}
