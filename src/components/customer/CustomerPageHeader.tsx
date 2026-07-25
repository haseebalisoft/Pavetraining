import { Breadcrumbs, type BreadcrumbItem } from "@/components/ui/Breadcrumbs";

import styles from "./customer.module.css";

interface PageHeaderProps {
  breadcrumbs: BreadcrumbItem[];
  eyebrow?: string;
  title: string;
  subtitle: string;
}

export function CustomerPageHeader({
  breadcrumbs,
  eyebrow = "Customer",
  title,
  subtitle,
}: PageHeaderProps) {
  return (
    <header className={styles.pageHeader}>
      <Breadcrumbs items={breadcrumbs} />
      <p className={styles.eyebrow}>{eyebrow}</p>
      <h1 className={styles.title}>{title}</h1>
      <p className={styles.subtitle}>{subtitle}</p>
    </header>
  );
}
