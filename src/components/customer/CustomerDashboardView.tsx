import Link from "next/link";

import { CustomerCompanyProfileCard } from "@/components/customer/CustomerCompanyProfileCard";
import { CustomerOfferSlider } from "@/components/customer/CustomerOfferSlider";
import { CustomerUpcomingEvents } from "@/components/customer/CustomerUpcomingEvents";
import { LiveTrainingRefresh } from "@/components/training/LiveTrainingRefresh";
import { formatDate } from "@/lib/utils/formatDate";
import type {
  CustomerCompanyProfile,
  CustomerDocumentRecord,
  CustomerEventRecord,
  CustomerOfferRecord,
  DashboardStats,
} from "@/types/models";

import { DashboardIcon, type DashboardIconName } from "./DashboardIcons";
import styles from "./customerDashboard.module.css";

interface CustomerDashboardViewProps {
  companyName: string;
  stats: DashboardStats;
  offers: CustomerOfferRecord[];
  upcomingEvents: CustomerEventRecord[];
  companyProfile?: CustomerCompanyProfile | null;
  recentDocuments: CustomerDocumentRecord[];
  profileShortcut: { href: string; label: string };
}

type StatTone = "danger" | "warn" | "ok" | "neutral";

interface StatCard {
  label: string;
  value: number;
  icon: DashboardIconName;
  tone: StatTone;
  hint?: string;
  href: string;
}

const TONE_CLASS: Record<StatTone, string> = {
  danger: styles.toneDanger,
  warn: styles.toneWarn,
  ok: styles.toneOk,
  neutral: styles.toneNeutral,
};

function plural(count: number, word: string): string {
  return `${count} ${word}${count === 1 ? "" : "s"}`;
}

function percent(part: number, total: number): number {
  if (total <= 0) return 0;
  return Math.round((part / total) * 100);
}

function fileExtension(name: string): string {
  const match = /\.([a-z0-9]{2,5})$/i.exec(name.trim());
  return match ? match[1].toUpperCase() : "DOC";
}

function initials(name: string): string {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("");
}

function ComplianceCard({ stats }: { stats: DashboardStats }) {
  const total = stats.trainingMatrixCount;
  const expired = stats.expiredCount;
  const urgent = stats.expiringSoonCount;
  const upcoming = stats.upcomingExpiryCount;
  const other = Math.max(total - expired - urgent - upcoming, 0);

  const segments = [
    { label: "In date / no renewal due", value: other, color: "var(--pave-lime)" },
    { label: "Due 91–270 days", value: upcoming, color: "#c9e8ad" },
    { label: "Due 0–90 days", value: urgent, color: "var(--pave-amber)" },
    { label: "Expired", value: expired, color: "var(--pave-red)" },
  ];

  let cursor = 0;
  const stops = segments
    .filter((segment) => segment.value > 0)
    .map((segment) => {
      const start = cursor;
      cursor += (segment.value / total) * 100;
      return `${segment.color} ${start}% ${cursor}%`;
    });
  const ring =
    total > 0 && stops.length > 0
      ? `conic-gradient(${stops.join(", ")})`
      : "conic-gradient(var(--pave-border) 0 100%)";

  return (
    <article className={styles.complianceCard} aria-label="Compliance split">
      <div className={styles.complianceRing} style={{ background: ring }}>
        <div className={styles.complianceRingInner}>
          <strong>{total > 0 ? `${percent(total - expired, total)}%` : "—"}</strong>
          <span>not expired</span>
        </div>
      </div>
      <h3>Compliance split</h3>
      <ul className={styles.complianceLegend}>
        {segments.map((segment) => (
          <li key={segment.label}>
            <span
              className={styles.legendDot}
              style={{ background: segment.color }}
              aria-hidden="true"
            />
            <span>{segment.label}</span>
            <strong>{percent(segment.value, total)}%</strong>
          </li>
        ))}
      </ul>
    </article>
  );
}

function RecentDocuments({ documents }: { documents: CustomerDocumentRecord[] }) {
  return (
    <section className={styles.splitColumn} aria-labelledby="recent-documents">
      <div className={styles.sectionHeader}>
        <h2 id="recent-documents">Recent documents</h2>
        <Link href="/customer/documents" className={styles.viewAllLink}>
          View all <DashboardIcon name="arrow" />
        </Link>
      </div>

      <div className={styles.listCard}>
        {documents.length === 0 ? (
          <div className={styles.listEmpty}>
            <strong>No documents yet</strong>
            <span>Documents shared with your company will appear here.</span>
          </div>
        ) : (
          documents.map((doc) => {
            const ext = fileExtension(doc.name);
            const actionHref = doc.downloadPath ?? doc.viewPath;
            return (
              <div key={doc.id} className={styles.docRow}>
                <span
                  className={
                    ext === "PDF" ? styles.docIconPdf : styles.docIconOther
                  }
                  aria-hidden="true"
                >
                  <DashboardIcon name="file" />
                  <small>{ext}</small>
                </span>
                <span className={styles.docBody}>
                  <span className={styles.docName} title={doc.name}>
                    {doc.name}
                  </span>
                  <span className={styles.docType}>
                    {doc.documentType ?? "Document"}
                  </span>
                </span>
                {doc.candidate ? (
                  <span className={styles.docCandidate}>
                    <span className={styles.miniAvatar} aria-hidden="true">
                      {initials(doc.candidate)}
                    </span>
                    <span>{doc.candidate}</span>
                  </span>
                ) : (
                  <span className={styles.docCandidate}>Company</span>
                )}
                <span className={styles.docDate}>
                  {formatDate(doc.uploadedDate, "—")}
                </span>
                {actionHref ? (
                  <a
                    href={actionHref}
                    className={styles.iconButton}
                    target={doc.downloadPath ? undefined : "_blank"}
                    rel={doc.downloadPath ? undefined : "noreferrer"}
                    aria-label={`${doc.downloadPath ? "Download" : "Open"} ${doc.name}`}
                  >
                    <DashboardIcon name={doc.downloadPath ? "download" : "eye"} />
                  </a>
                ) : (
                  <span />
                )}
              </div>
            );
          })
        )}
      </div>
    </section>
  );
}

export function CustomerDashboardView({
  stats,
  offers,
  upcomingEvents,
  companyProfile = null,
  recentDocuments,
  profileShortcut,
}: CustomerDashboardViewProps) {
  const nextEvent = upcomingEvents.find((event) => event.eventDate);

  const statCards: StatCard[] = [
    {
      label: "Candidates",
      value: stats.workforceCount,
      icon: "users",
      tone: "ok",
      hint: "In your access",
      href: profileShortcut.href,
    },
    {
      label: "Matrix rows",
      value: stats.trainingMatrixCount,
      icon: "table",
      tone: "ok",
      hint: "Training Matrix",
      href: "/customer/training-matrix",
    },
    {
      label: "Expired",
      value: stats.expiredCount,
      icon: "shield",
      tone: stats.expiredCount > 0 ? "danger" : "ok",
      hint: stats.expiredCount > 0 ? "Need action" : "All clear",
      href: "/customer/training-matrix?filter=expired",
    },
    {
      label: "Urgent 0–90 days",
      value: stats.expiringSoonCount,
      icon: "clock",
      tone: stats.expiringSoonCount > 0 ? "danger" : "ok",
      hint: stats.expiringSoonCount > 0 ? "Book renewals" : "None due",
      href: "/customer/training-matrix?filter=urgent",
    },
    {
      label: "Upcoming 91–270",
      value: stats.upcomingExpiryCount,
      icon: "calendar",
      tone: stats.upcomingExpiryCount > 0 ? "warn" : "ok",
      hint: "Plan ahead",
      href: "/customer/training-matrix?filter=upcoming",
    },
    {
      label: "Records to review",
      value: stats.needsReviewCount,
      icon: "check",
      tone: stats.needsReviewCount > 0 ? "warn" : "ok",
      hint: stats.needsReviewCount > 0 ? "Awaiting review" : "Up to date",
      href: "/customer/training-matrix?filter=review",
    },
    {
      label: "NVQ programmes",
      value: stats.nvqCount,
      icon: "message",
      tone: "ok",
      hint: "NVQ Progress",
      href: "/customer/nvq-progress",
    },
    {
      label: "Documents",
      value: stats.documentsCount,
      icon: "file",
      tone: "ok",
      hint: "Shared with you",
      href: "/customer/documents",
    },
    {
      label: "Upcoming events",
      value: stats.upcomingEventsCount,
      icon: "calendar",
      tone: "ok",
      hint: nextEvent
        ? `Next: ${formatDate(nextEvent.eventDate, "TBC")}`
        : "None booked",
      href: "/customer/events",
    },
  ];

  const isOwnProfile = profileShortcut.href !== "/customer/candidates";
  const quickLinks: Array<{
    href: string;
    title: string;
    description: string;
    icon: DashboardIconName;
  }> = [
    {
      href: "/customer/training-matrix",
      title: "Training Matrix",
      description: "Check company compliance",
      icon: "table",
    },
    {
      href: "/customer/courses",
      title: "PAVE Training Courses",
      description: "Browse available courses",
      icon: "message",
    },
    {
      href: profileShortcut.href,
      title: profileShortcut.label,
      description: isOwnProfile
        ? "Your training and certificates"
        : `Manage your ${plural(stats.workforceCount, "candidate")}`,
      icon: "users",
    },
    {
      href: "/customer/training-records",
      title: "Training Records",
      description: "Open all training registers",
      icon: "shield",
    },
    {
      href: "/customer/nvq-progress",
      title: "NVQ Progress",
      description: "Track programmes and stages",
      icon: "check",
    },
    {
      href: "/customer/documents",
      title: "Documents",
      description: "Certificates and card scans",
      icon: "file",
    },
    {
      href: "/customer/events",
      title: "Events",
      description: "Upcoming courses and bookings",
      icon: "calendar",
    },
    {
      href: "/customer/offers",
      title: "Offers",
      description: "See customer-only savings",
      icon: "message",
    },
    {
      href: "/customer/support",
      title: "Support",
      description: "Speak to the PAVE team",
      icon: "help",
    },
  ];

  return (
    <div className={styles.dashboardPage}>
      <LiveTrainingRefresh />

      <header className={styles.pageHeader}>
        <nav className={styles.breadcrumbs} aria-label="Breadcrumb">
          <Link href="/customer/dashboard">Customer</Link>
          <DashboardIcon name="chevron" />
          <span aria-current="page">Dashboard</span>
        </nav>
        <p className={styles.eyebrow}>Customer</p>
        <h1 className={styles.pageTitle}>Dashboard</h1>
        <p className={styles.pageSubtitle}>
          Your company training, compliance and upcoming activity at a glance.
        </p>
      </header>

      {companyProfile ? (
        <CustomerCompanyProfileCard company={companyProfile} />
      ) : null}

      <section aria-labelledby="matrix-summary">
        <div className={styles.sectionHeader}>
          <h2 id="matrix-summary">Training Matrix summary</h2>
          <Link href="/customer/training-matrix" className={styles.viewAllLink}>
            Open matrix <DashboardIcon name="arrow" />
          </Link>
        </div>
        <div className={styles.summaryGrid}>
          <div className={styles.statsGrid}>
            {statCards.map((card) => (
              <Link
                key={card.label}
                href={card.href}
                className={`${styles.statCard} ${TONE_CLASS[card.tone]}`}
              >
                <span className={styles.statTop}>
                  <span className={styles.statIcon}>
                    <DashboardIcon name={card.icon} />
                  </span>
                  {card.hint ? (
                    <span className={styles.statHint}>{card.hint}</span>
                  ) : null}
                </span>
                <strong className={styles.statValue}>{card.value}</strong>
                <span className={styles.statLabel}>{card.label}</span>
              </Link>
            ))}
          </div>
          <ComplianceCard stats={stats} />
        </div>
      </section>

      <div className={styles.splitRow}>
        <CustomerUpcomingEvents events={upcomingEvents} />
        <RecentDocuments documents={recentDocuments} />
      </div>

      <CustomerOfferSlider offers={offers} />

      <section aria-labelledby="quick-links">
        <div className={styles.sectionHeader}>
          <h2 id="quick-links">Quick links</h2>
        </div>
        <div className={styles.quickGrid}>
          {quickLinks.map((link) => (
            <Link key={link.title} href={link.href} className={styles.quickCard}>
              <span className={styles.quickIcon}>
                <DashboardIcon name={link.icon} />
              </span>
              <span className={styles.quickTitle}>{link.title}</span>
              <span className={styles.quickDescription}>{link.description}</span>
              <span className={styles.quickCta}>
                Open <DashboardIcon name="arrow" />
              </span>
            </Link>
          ))}
        </div>
      </section>
    </div>
  );
}
