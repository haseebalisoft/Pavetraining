import Link from "next/link";

import { AdminCompanySelector } from "@/components/admin/AdminCompanySelector";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type {
  AdminDashboardPayload,
  Company,
} from "@/types/models";

import styles from "./admin.module.css";

const QUICK_ACTIONS = [
  {
    href: "/admin/companies?action=add",
    title: "Add Company",
    description: "Create a new company record in SharePoint.",
  },
  {
    href: "/admin/workforce?action=add",
    title: "Add Candidate",
    description: "Register a workforce candidate.",
  },
  {
    href: "/admin/documents",
    title: "Upload Document",
    description: "Open documents to manage uploads and visibility.",
  },
  {
    href: "/admin/events?action=add",
    title: "Create Booking",
    description: "Schedule a training event / booking.",
  },
  {
    href: "/admin/offers",
    title: "Manage offers",
    description: "Offers and promotions for customers.",
  },
  {
    href: "/admin/permissions",
    title: "Manage permissions",
    description: "Control portal access and download rights.",
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
  const {
    stats,
    warnings,
    selectedCompanyId,
    selectedCompanyName,
    upcomingExpiries,
    recentDocuments,
    upcomingBookings,
    recentActivity,
  } = dashboard;

  const cards: Array<{
    label: string;
    value: number;
    tone?: "warn" | "ok";
  }> = [
    { label: "Active companies", value: stats.activeCompanies, tone: "ok" },
    { label: "Active candidates", value: stats.activeCandidates, tone: "ok" },
    {
      label: "Expiring within 3 months",
      value: stats.expiringWithin3Months,
      tone: "warn",
    },
    {
      label: "Expiring within 6 months",
      value: stats.expiringWithin6Months,
      tone: "warn",
    },
    { label: "Upcoming bookings", value: stats.upcomingEvents },
    {
      label: "Documents uploaded recently",
      value: stats.documentsUploadedRecently,
    },
    { label: "NVQs in progress", value: stats.activeNvqs },
    {
      label: "Access invitations pending",
      value: stats.accessInvitationsPending,
      tone: stats.accessInvitationsPending > 0 ? "warn" : undefined,
    },
  ];

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
        <section className={styles.panel} aria-label="Upcoming training expiries">
          <div className={styles.panelHeaderRow}>
            <h2 className={styles.panelTitle}>Upcoming training expiries</h2>
            <Link className={styles.panelLink} href="/admin/training-matrix">
              Open matrix
            </Link>
          </div>
          {upcomingExpiries.length === 0 ? (
            <p className={styles.emptyNote}>No urgent or upcoming expiries.</p>
          ) : (
            <ul className={styles.dashList}>
              {upcomingExpiries.map((row) => (
                <li key={row.id} className={styles.dashListItem}>
                  <div>
                    <strong>{row.candidateName}</strong>
                    <span className={styles.dashMeta}>
                      {[row.companyName, formatDisplayDate(row.nextExpiryDate)]
                        .filter(Boolean)
                        .join(" · ")}
                    </span>
                  </div>
                  <StatusBadge label={row.statusLabel} tone={row.statusTone} />
                </li>
              ))}
            </ul>
          )}
        </section>

        <section className={styles.panel} aria-label="Recent document uploads">
          <div className={styles.panelHeaderRow}>
            <h2 className={styles.panelTitle}>Recent document uploads</h2>
            <Link className={styles.panelLink} href="/admin/documents">
              Open documents
            </Link>
          </div>
          {recentDocuments.length === 0 ? (
            <p className={styles.emptyNote}>No recent documents found.</p>
          ) : (
            <ul className={styles.dashList}>
              {recentDocuments.map((row) => (
                <li key={row.id} className={styles.dashListItem}>
                  <div>
                    <strong>{row.name}</strong>
                    <span className={styles.dashMeta}>
                      {[
                        row.company,
                        row.candidate,
                        formatDisplayDate(row.modifiedDate),
                      ]
                        .filter(Boolean)
                        .join(" · ")}
                    </span>
                  </div>
                  <StatusBadge
                    label={row.customerVisible ? "Visible" : "Hidden"}
                    tone={row.customerVisible ? "ok" : "warn"}
                  />
                </li>
              ))}
            </ul>
          )}
        </section>

        <section className={styles.panel} aria-label="Upcoming company events">
          <div className={styles.panelHeaderRow}>
            <h2 className={styles.panelTitle}>Upcoming company events</h2>
            <Link className={styles.panelLink} href="/admin/events">
              Open bookings
            </Link>
          </div>
          {upcomingBookings.length === 0 ? (
            <p className={styles.emptyNote}>No upcoming bookings.</p>
          ) : (
            <ul className={styles.dashList}>
              {upcomingBookings.map((row) => (
                <li key={row.id} className={styles.dashListItem}>
                  <div>
                    <strong>{row.title}</strong>
                    <span className={styles.dashMeta}>
                      {[
                        row.company,
                        formatDisplayDate(row.eventDate),
                        row.location,
                      ]
                        .filter(Boolean)
                        .join(" · ")}
                    </span>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </section>

        <section className={styles.panel} aria-label="Recent portal activity">
          <div className={styles.panelHeaderRow}>
            <h2 className={styles.panelTitle}>Recent portal activity</h2>
            <Link className={styles.panelLink} href="/admin/logs">
              Open audit log
            </Link>
          </div>
          {recentActivity.length === 0 ? (
            <p className={styles.emptyNote}>
              No recent portal activity yet. New admin actions will appear here.
            </p>
          ) : (
            <ul className={styles.dashList}>
              {recentActivity.map((row) => (
                <li key={row.id} className={styles.dashListItem}>
                  <div>
                    <strong>{row.title}</strong>
                    <span className={styles.dashMeta}>
                      {[row.userEmail, formatDisplayDate(row.timestamp)]
                        .filter(Boolean)
                        .join(" · ")}
                    </span>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </section>
      </div>

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
