import Link from "next/link";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";

import styles from "../../../components/admin/admin.module.css";

const SECTIONS = [
  {
    href: "/admin/training-records/npors",
    title: "NPORS Training",
    description: "Add and edit NPORS records, outcomes, and visibility.",
  },
  {
    href: "/admin/training-records/eusr",
    title: "EUSR Training",
    description: "Manage EUSR registrations and customer visibility.",
  },
  {
    href: "/admin/training-records/streetworks",
    title: "Streetworks Training",
    description: "Maintain Streetworks / SWQR training records.",
  },
  {
    href: "/admin/training-records/in-house",
    title: "In-House Training",
    description: "Manage in-house certificates and outcomes.",
  },
] as const;

export default function AdminTrainingRecordsHubPage() {
  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs
            items={[
              { label: "Admin", href: "/admin" },
              { label: "Training Records" },
            ]}
          />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>Training Records</h1>
          <p className={styles.subtitle}>
            Choose a register to add or edit training records.
          </p>
        </div>
      </header>
      <section className={styles.hubGrid} aria-label="Training record registers">
        {SECTIONS.map((section) => (
          <Link key={section.href} className={styles.hubCard} href={section.href}>
            <h2>{section.title}</h2>
            <p>{section.description}</p>
          </Link>
        ))}
      </section>
    </div>
  );
}
