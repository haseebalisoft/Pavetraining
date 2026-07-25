import Link from "next/link";

import { CustomerCompanyProfileCard } from "@/components/customer/CustomerCompanyProfileCard";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import type { CustomerCompanyProfile, DashboardStats } from "@/types/models";

import styles from "./customer.module.css";

const SECTIONS = [
  {
    href: "/customer/training-matrix",
    title: "Training Matrix",
    description: "Workforce competency overview and upcoming renewals.",
  },
  {
    href: "/customer/candidates",
    title: "Candidates",
    description: "Browse people in your organisation.",
  },
  {
    href: "/customer/training-records",
    title: "Training Records",
    description: "NPORS, EUSR, Streetworks, and In-House registers.",
  },
  {
    href: "/customer/nvq-progress",
    title: "NVQ Progress",
    description: "Active and completed NVQ programmes.",
  },
  {
    href: "/customer/documents",
    title: "Documents",
    description: "Shared company documents and downloads.",
  },
  {
    href: "/customer/events",
    title: "Events",
    description: "Upcoming and visible training events.",
  },
  {
    href: "/customer/offers",
    title: "Offers",
    description: "Current promotions shared with your company.",
  },
] as const;

interface CustomerDashboardViewProps {
  companyName: string;
  email: string;
  permissionStatus: string;
  accessScope: string;
  canDownload: boolean;
  stats: DashboardStats;
  companyProfile?: CustomerCompanyProfile | null;
}

export function CustomerDashboardView({
  companyName,
  email,
  permissionStatus,
  accessScope,
  canDownload,
  stats,
  companyProfile = null,
}: CustomerDashboardViewProps) {
  const warningCards = [
    {
      href: "/customer/training-matrix?filter=expired",
      title: "Expired training",
      value: stats.expiredCount,
      detail: "Matrix rows with an expired next date",
      show: stats.expiredCount > 0,
    },
    {
      href: "/customer/training-matrix?filter=expiring",
      title: "Expiring within 3 months",
      value: stats.expiringSoonCount,
      detail: "Renewals due in the next 90 days",
      show: stats.expiringSoonCount > 0,
    },
    {
      href: "/customer/training-matrix?filter=review",
      title: "Records to Review",
      value: stats.needsReviewCount,
      detail: "Matrix rows flagged for attention",
      show: stats.needsReviewCount > 0,
    },
  ].filter((card) => card.show);

  return (
    <div>
      <header className={styles.pageHeader}>
        <Breadcrumbs items={[{ label: "Customer" }, { label: "Dashboard" }]} />
        <p className={styles.eyebrow}>Customer portal</p>
        <h1 className={styles.title}>Dashboard</h1>
        <p className={styles.subtitle}>
          A clear view of training health for {companyName}. Only records marked
          visible to your organisation are included.
        </p>
      </header>

      {companyProfile ? (
        <CustomerCompanyProfileCard company={companyProfile} />
      ) : null}

      <section className={styles.metaGrid} aria-label="Account summary">
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Company</p>
          <p className={styles.metaValue}>{companyName}</p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Signed in</p>
          <p className={styles.metaValue}>{email}</p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Permission status</p>
          <p className={styles.metaValue}>
            <StatusBadge
              label={permissionStatus}
              tone={
                permissionStatus.toLowerCase() === "active" ? "ok" : "warn"
              }
            />
          </p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Access scope</p>
          <p className={styles.metaValue}>{accessScope}</p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Downloads</p>
          <p className={styles.metaValue}>
            <StatusBadge
              label={canDownload ? "Enabled" : "View only"}
              tone={canDownload ? "ok" : "neutral"}
            />
          </p>
        </div>
      </section>

      <section className={styles.statsGrid} aria-label="Training overview">
        <article className={styles.statCard}>
          <p className={styles.statLabel}>Candidates</p>
          <p className={styles.statValue}>{stats.workforceCount}</p>
        </article>
        <article className={styles.statCard}>
          <p className={styles.statLabel}>Matrix rows</p>
          <p className={styles.statValue}>{stats.trainingMatrixCount}</p>
        </article>
        <article
          className={`${styles.statCard} ${stats.expiredCount ? styles.statDanger : ""}`}
        >
          <p className={styles.statLabel}>Expired</p>
          <p className={styles.statValue}>{stats.expiredCount}</p>
        </article>
        <article
          className={`${styles.statCard} ${stats.expiringSoonCount ? styles.statWarn : ""}`}
        >
          <p className={styles.statLabel}>Expiring (3 months)</p>
          <p className={styles.statValue}>{stats.expiringSoonCount}</p>
        </article>
        <article
          className={`${styles.statCard} ${stats.needsReviewCount ? styles.statWarn : styles.statOk}`}
        >
          <p className={styles.statLabel}>Records to Review</p>
          <p className={styles.statValue}>{stats.needsReviewCount}</p>
        </article>
        <article className={styles.statCard}>
          <p className={styles.statLabel}>NVQ programmes</p>
          <p className={styles.statValue}>{stats.nvqCount}</p>
        </article>
        <article className={styles.statCard}>
          <p className={styles.statLabel}>Documents</p>
          <p className={styles.statValue}>{stats.documentsCount}</p>
        </article>
        <article className={styles.statCard}>
          <p className={styles.statLabel}>Upcoming events</p>
          <p className={styles.statValue}>{stats.upcomingEventsCount}</p>
        </article>
      </section>

      {warningCards.length > 0 ? (
        <section className={styles.panel} aria-label="Expiry warnings">
          <h2 className={styles.panelTitle}>Attention needed</h2>
          <div className={styles.warningList}>
            {warningCards.map((card) => (
              <Link key={card.href} href={card.href} className={styles.warningCard}>
                <strong>
                  {card.title} · {card.value}
                </strong>
                <span>{card.detail}</span>
              </Link>
            ))}
          </div>
        </section>
      ) : null}

      <section
        className={styles.cardGrid}
        aria-label="Customer portal sections"
        style={{ marginTop: "1.35rem" }}
      >
        {SECTIONS.map((section) => (
          <Link key={section.href} href={section.href} className={styles.navCard}>
            <h2>{section.title}</h2>
            <p>{section.description}</p>
            <span className={styles.navCardCta}>Open →</span>
          </Link>
        ))}
      </section>
    </div>
  );
}
