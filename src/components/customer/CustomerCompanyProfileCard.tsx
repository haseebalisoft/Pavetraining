import { StatusBadge } from "@/components/ui/StatusBadge";
import type { CustomerCompanyProfile } from "@/types/models";

import styles from "./customer.module.css";

function valueOrDash(value: string | null | undefined): string {
  const trimmed = value?.trim();
  return trimmed ? trimmed : "—";
}

interface CustomerCompanyProfileCardProps {
  company: CustomerCompanyProfile;
}

export function CustomerCompanyProfileCard({
  company,
}: CustomerCompanyProfileCardProps) {
  return (
    <section className={styles.companyProfile} aria-label="Your company">
      <div className={styles.companyProfileHeader}>
        <div>
          <p className={styles.metaLabel}>Your company</p>
          <h2 className={styles.companyProfileTitle}>{company.companyName}</h2>
        </div>
        {company.companyLogo ? (
          // eslint-disable-next-line @next/next/no-img-element
          <img
            className={styles.companyLogo}
            src={company.companyLogo}
            alt={`${company.companyName} logo`}
          />
        ) : null}
      </div>

      <div className={styles.companyProfileGrid}>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Company number</p>
          <p className={styles.metaValue}>
            {valueOrDash(company.companyNumber)}
          </p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Company size</p>
          <p className={styles.metaValue}>{valueOrDash(company.companySize)}</p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Status</p>
          <p className={styles.metaValue}>
            <StatusBadge
              label={company.status}
              tone={
                company.status.toLowerCase() === "active" ? "ok" : "neutral"
              }
            />
          </p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Company reg number</p>
          <p className={styles.metaValue}>
            {valueOrDash(company.companyRegNumber)}
          </p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>VAT no</p>
          <p className={styles.metaValue}>{valueOrDash(company.vatNo)}</p>
        </div>
        <div className={`${styles.metaItem} ${styles.metaItemWide}`}>
          <p className={styles.metaLabel}>Registered address</p>
          <p className={styles.metaValue}>
            {valueOrDash(company.registeredAddress)}
          </p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Main contact</p>
          <p className={styles.metaValue}>
            {valueOrDash(company.mainContact)}
          </p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Tel no</p>
          <p className={styles.metaValue}>{valueOrDash(company.telNo)}</p>
        </div>
        <div className={styles.metaItem}>
          <p className={styles.metaLabel}>Email</p>
          <p className={styles.metaValue}>{valueOrDash(company.email)}</p>
        </div>
      </div>
    </section>
  );
}
