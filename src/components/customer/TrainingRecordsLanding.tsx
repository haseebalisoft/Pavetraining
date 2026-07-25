import Link from "next/link";

import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";

import styles from "./customer.module.css";

const SECTIONS = [
  {
    href: "/customer/training-records/npors",
    title: "NPORS Training",
    description:
      "View plant and machinery qualifications, categories, and expiry dates for your workforce.",
  },
  {
    href: "/customer/training-records/eusr",
    title: "EUSR Training",
    description:
      "Review EUSR registrations, categories, outcomes, and upcoming renewals.",
  },
  {
    href: "/customer/training-records/streetworks",
    title: "Streetworks Training",
    description:
      "Access SWQR and streetworks course records for operatives in your company.",
  },
  {
    href: "/customer/training-records/in-house",
    title: "In-House Training",
    description:
      "Browse in-house certificates, course dates, and pass or fail outcomes.",
  },
] as const;

interface TrainingRecordsLandingProps {
  companyName: string;
}

export function TrainingRecordsLanding({
  companyName,
}: TrainingRecordsLandingProps) {
  return (
    <div>
      <CustomerPageHeader
        breadcrumbs={[
          { label: "Customer", href: "/customer" },
          { label: "Training Records" },
        ]}
        title="Training Records"
        subtitle="Browse company-scoped qualification registers. Only customer-visible records for your organisation are shown."
      />

      <p className={styles.companyMeta}>
        Showing records for <strong>{companyName}</strong>
      </p>

      <section className={styles.cardGrid} aria-label="Training record sections">
        {SECTIONS.map((section) => (
          <Link
            key={section.href}
            className={styles.navCard}
            href={section.href}
          >
            <h2>{section.title}</h2>
            <p>{section.description}</p>
            <span className={styles.navCardCta}>Open register →</span>
          </Link>
        ))}
      </section>
    </div>
  );
}
