import Link from "next/link";

import { CustomerCompanyProfileCard } from "@/components/customer/CustomerCompanyProfileCard";
import { CustomerOfferSlider } from "@/components/customer/CustomerOfferSlider";
import { CustomerUpcomingEvents } from "@/components/customer/CustomerUpcomingEvents";
import type {
  CustomerCompanyProfile,
  CustomerEventRecord,
  CustomerOfferRecord,
  DashboardStats,
} from "@/types/models";

import styles from "./customer.module.css";
import dashStyles from "./customerDashboard.module.css";

const SECTIONS = [
  {
    href: "/customer",
    title: "Training Matrix",
    description: "Workforce competency overview and upcoming renewals.",
  },
  {
    href: "/customer/courses",
    title: "PAVE Training Courses",
    description: "Everything PAVE can deliver — NPORS, Streetworks, EUSR, NVQ and more.",
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
  {
    href: "/customer/support",
    title: "Support",
    description: "Contact PAVE Training with a form or call us.",
  },
] as const;

interface CustomerDashboardViewProps {
  companyName: string;
  stats: DashboardStats;
  offers: CustomerOfferRecord[];
  upcomingEvents: CustomerEventRecord[];
  companyProfile?: CustomerCompanyProfile | null;
}

type StatTone = "danger" | "warn" | "ok" | undefined;

function toneClass(tone: StatTone): string {
  if (tone === "danger") return styles.statDanger;
  if (tone === "warn") return styles.statWarn;
  if (tone === "ok") return styles.statOk;
  return "";
}

export function CustomerDashboardView({
  stats,
  offers,
  upcomingEvents,
  companyProfile = null,
}: CustomerDashboardViewProps) {
  const warningCards = [
    {
      href: "/customer?filter=expired",
      title: "Expired training",
      value: stats.expiredCount,
      detail: "Matrix rows with an expired next date",
      show: stats.expiredCount > 0,
    },
    {
      href: "/customer?filter=within-3m",
      title: "Urgent",
      value: stats.expiringSoonCount,
      detail: "Renewals due in the next 90 days",
      show: stats.expiringSoonCount > 0,
    },
    {
      href: "/customer?filter=within-6m",
      title: "Upcoming",
      value: stats.upcomingExpiryCount,
      detail: "Renewals due in 91–180 days",
      show: stats.upcomingExpiryCount > 0,
    },
    {
      href: "/customer?filter=review",
      title: "Records to Review",
      value: stats.needsReviewCount,
      detail: "Matrix rows flagged for attention",
      show: stats.needsReviewCount > 0,
    },
  ].filter((card) => card.show);

  const overviewCards: Array<{
    label: string;
    value: number;
    tone?: StatTone;
  }> = [
    { label: "Candidates", value: stats.workforceCount },
    { label: "Matrix rows", value: stats.trainingMatrixCount },
    {
      label: "Expired",
      value: stats.expiredCount,
      tone: stats.expiredCount > 0 ? "danger" : undefined,
    },
    {
      label: "Urgent (0–90 days)",
      value: stats.expiringSoonCount,
      tone: stats.expiringSoonCount > 0 ? "danger" : undefined,
    },
    {
      label: "Upcoming (91–270 days)",
      value: stats.upcomingExpiryCount,
      tone: stats.upcomingExpiryCount > 0 ? "warn" : undefined,
    },
    {
      label: "Records to Review",
      value: stats.needsReviewCount,
      tone: stats.needsReviewCount > 0 ? "warn" : "ok",
    },
    { label: "NVQ programmes", value: stats.nvqCount },
    { label: "Documents", value: stats.documentsCount },
    { label: "Upcoming events", value: stats.upcomingEventsCount },
  ];

  return (
    <div className={dashStyles.dashboardPage}>
      <h1 className={styles.srOnly}>Dashboard</h1>

      {/* Reference order: offers hero, then upcoming events */}
      <CustomerOfferSlider offers={offers} />
      <CustomerUpcomingEvents events={upcomingEvents} />

      {companyProfile ? (
        <CustomerCompanyProfileCard company={companyProfile} />
      ) : null}

      <section className={styles.statsSection} aria-label="Training overview">
        <div className={dashStyles.sectionHeader}>
          <h2>Training overview</h2>
        </div>
        <div className={styles.statsGrid}>
          {overviewCards.map((card) => (
            <article
              key={card.label}
              className={`${styles.statCard} ${toneClass(card.tone)}`}
            >
              <p className={styles.statLabel}>{card.label}</p>
              <p className={styles.statValue}>{card.value}</p>
            </article>
          ))}
        </div>
      </section>

      {warningCards.length > 0 ? (
        <section className={styles.panel} aria-label="Expiry warnings">
          <div className={dashStyles.sectionHeader}>
            <h2>Attention needed</h2>
          </div>
          <div className={styles.warningList}>
            {warningCards.map((card) => (
              <Link
                key={card.href}
                href={card.href}
                className={styles.warningCard}
              >
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
        className={styles.cardGridSection}
        aria-label="Customer portal sections"
      >
        <div className={dashStyles.sectionHeader}>
          <h2>Quick links</h2>
        </div>
        <div className={styles.cardGrid}>
          {SECTIONS.map((section) => (
            <Link
              key={section.href}
              href={section.href}
              className={styles.navCard}
            >
              <h2>{section.title}</h2>
              <p>{section.description}</p>
              <span className={styles.navCardCta}>Open →</span>
            </Link>
          ))}
        </div>
      </section>
    </div>
  );
}
