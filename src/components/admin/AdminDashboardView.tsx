import Link from "next/link";

import { AdminCompanySelector } from "@/components/admin/AdminCompanySelector";
import {
  AdminHubHeroSlider,
  type HubHeroSlide,
} from "@/components/admin/AdminHubHeroSlider";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDate } from "@/lib/utils/formatDate";
import type {
  AdminDashboardPayload,
  Company,
} from "@/types/models";

import styles from "./admin.module.css";

const ACTION_TILES = [
  {
    href: "/admin/companies?action=add",
    title: "Add Company",
    hint: "New company record",
    icon: "C",
  },
  {
    href: "/admin/workforce?action=add",
    title: "Add Candidate",
    hint: "Workforce register",
    icon: "W",
  },
  {
    href: "/admin/documents",
    title: "Documents",
    hint: "Upload & visibility",
    icon: "D",
  },
  {
    href: "/admin/events?action=add",
    title: "Create Booking",
    hint: "Calendar event",
    icon: "B",
  },
  {
    href: "/admin/training-matrix",
    title: "Matrix",
    hint: "Training expiries",
    icon: "M",
  },
  {
    href: "/admin/training-records",
    title: "Registers",
    hint: "NPORS · EUSR · more",
    icon: "R",
  },
  {
    href: "/admin/offers",
    title: "Offers",
    hint: "Promotions",
    icon: "O",
  },
  {
    href: "/admin/permissions",
    title: "Permissions",
    hint: "Portal access",
    icon: "P",
  },
] as const;

const RESOURCE_TILES = [
  {
    href: "/admin/training-matrix",
    title: "Training Matrix",
    description: "Wide expiry grid and sync",
    tone: "lime" as const,
  },
  {
    href: "/admin/training-records",
    title: "Training Registers",
    description: "NPORS, EUSR, Streetworks, In-House",
    tone: "charcoal" as const,
  },
  {
    href: "/admin/documents",
    title: "Customer Documents",
    description: "Folders, uploads, visibility",
    tone: "forest" as const,
  },
  {
    href: "/admin/events",
    title: "Calendar / Bookings",
    description: "Events and Outlook sync",
    tone: "slate" as const,
  },
  {
    href: "/admin/bulk-upload",
    title: "Bulk Upload",
    description: "Import spreadsheets",
    tone: "moss" as const,
  },
  {
    href: "/admin/logs",
    title: "Audit Log",
    description: "Portal activity history",
    tone: "ink" as const,
  },
] as const;

const RESOURCE_TONES: Record<
  (typeof RESOURCE_TILES)[number]["tone"],
  string
> = {
  lime: styles.hubTone_lime,
  charcoal: styles.hubTone_charcoal,
  forest: styles.hubTone_forest,
  slate: styles.hubTone_slate,
  moss: styles.hubTone_moss,
  ink: styles.hubTone_ink,
};

function displayNameFromEmail(email: string): string {
  const local = email.split("@")[0]?.trim() || email.trim();
  const parts = local.split(/[.\-_+\s]+/).filter(Boolean);
  if (parts.length === 0) return "Admin";
  return parts
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

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
    selectedCompanyId,
    selectedCompanyName,
    upcomingExpiries,
    recentDocuments,
    upcomingBookings,
  } = dashboard;

  const welcomeName = displayNameFromEmail(email);

  const heroSlides: HubHeroSlide[] = [
    {
      id: "welcome",
      eyebrow: "PAVE Training · Admin",
      title: `Welcome, ${welcomeName}`,
      subtitle:
        "Operations hub for companies, workforce, matrix, bookings, and customer access — powered by SharePoint.",
      ctaLabel: "Browse companies",
      ctaHref: "/admin/companies",
      metricLabel: "Active companies",
      metricValue: stats.activeCompanies,
    },
    {
      id: "matrix",
      eyebrow: "Training matrix",
      title: "Stay ahead of expiries",
      subtitle:
        "Review candidates nearing expiry, open the matrix grid, and keep compliance on track.",
      ctaLabel: "Open matrix",
      ctaHref: "/admin/training-matrix",
      metricLabel: "Expiring in 3 months",
      metricValue: stats.expiringWithin3Months,
    },
    {
      id: "documents",
      eyebrow: "Customer documents",
      title: "Files ready for customers",
      subtitle:
        "Upload certificates, control visibility, and keep company folders tidy for the customer portal.",
      ctaLabel: "Open documents",
      ctaHref: "/admin/documents",
      metricLabel: "Recent uploads",
      metricValue: stats.documentsUploadedRecently,
    },
    {
      id: "bookings",
      eyebrow: "Calendar & bookings",
      title: "Plan the next sessions",
      subtitle:
        "Create bookings, check upcoming company events, and keep the operations calendar current.",
      ctaLabel: "Open calendar",
      ctaHref: "/admin/events",
      metricLabel: "Upcoming bookings",
      metricValue: stats.upcomingEvents,
    },
    {
      id: "access",
      eyebrow: "Portal access",
      title: "Manage who can sign in",
      subtitle:
        "Invite customers, review permissions, and clear pending access invitations quickly.",
      ctaLabel: "Open permissions",
      ctaHref: "/admin/permissions",
      metricLabel: "Invitations pending",
      metricValue: stats.accessInvitationsPending,
    },
  ];

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
    <div className={styles.hubPage}>
      <section className={styles.hubHero} aria-label="Welcome">
        <div className={styles.hubHeroGlow} aria-hidden />
        <div className={styles.hubHeroInner}>
          <AdminHubHeroSlider slides={heroSlides} />
          <div className={styles.hubHeroPanel}>
            <div>
              <p className={styles.hubHeroPanelLabel}>Operations</p>
              <p className={styles.hubHeroPanelHint}>
                Filter metrics by company, then use the tiles below for everyday
                admin tasks — matrix, registers, documents, and bookings.
              </p>
            </div>
            <div>
              <p className={styles.hubHeroPanelLabel}>Focus company</p>
              <AdminCompanySelector
                companies={companies}
                selectedCompanyId={selectedCompanyId}
              />
              {selectedCompanyName ? (
                <p className={styles.filterChip}>
                  Filtered to <strong>{selectedCompanyName}</strong>
                  <Link className={styles.clearFilter} href="/admin">
                    Clear
                  </Link>
                </p>
              ) : (
                <p className={styles.hubHeroPanelHint}>
                  Showing all companies until you pick one.
                </p>
              )}
            </div>
          </div>
        </div>
      </section>

      <section className={styles.hubActionGrid} aria-label="Quick actions">
        {ACTION_TILES.map((tile) => (
          <Link
            key={tile.href}
            className={styles.hubActionTile}
            href={tile.href}
          >
            <span className={styles.hubActionIcon} aria-hidden>
              {tile.icon}
            </span>
            <span className={styles.hubActionTitle}>{tile.title}</span>
            <span className={styles.hubActionHint}>{tile.hint}</span>
          </Link>
        ))}
      </section>

      <section className={styles.hubResources} aria-label="Top resources">
        <div className={styles.hubResourcesHeader}>
          <h2 className={styles.hubResourcesTitle}>Top resources</h2>
          <p className={styles.hubResourcesSubtitle}>
            Jump into the lists you use most.
          </p>
        </div>
        <div className={styles.hubResourceGrid}>
          {RESOURCE_TILES.map((tile) => (
            <Link
              key={tile.href}
              className={`${styles.hubResourceTile} ${RESOURCE_TONES[tile.tone]}`}
              href={tile.href}
            >
              <strong>{tile.title}</strong>
              <span>{tile.description}</span>
            </Link>
          ))}
        </div>
      </section>

      <section className={styles.statsGrid} aria-label="Dashboard statistics">
        {cards.map((card) => (
          <article
            key={card.label}
            className={`${styles.statCard} ${styles.statCardElevated} ${
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
        <section
          className={`${styles.panel} ${styles.panelElevated}`}
          aria-label="Upcoming training expiries"
        >
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
                    <strong title={row.candidateName}>{row.candidateName}</strong>
                    <span className={styles.dashMeta}>
                      {[row.companyName, formatDate(row.nextExpiryDate)]
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

        <section
          className={`${styles.panel} ${styles.panelElevated}`}
          aria-label="Recent document uploads"
        >
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
                    <strong title={row.name}>{row.name}</strong>
                    <span className={styles.dashMeta}>
                      {[
                        row.company,
                        row.candidate,
                        formatDate(row.modifiedDate),
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

        <section
          className={`${styles.panel} ${styles.panelElevated}`}
          aria-label="Upcoming company events"
        >
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
                        formatDate(row.eventDate),
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
      </div>
    </div>
  );
}
