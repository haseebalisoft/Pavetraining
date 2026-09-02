import { StatusBadge } from "@/components/ui/StatusBadge";
import type { CustomerCompanyProfile } from "@/types/models";

import styles from "./customer.module.css";

function valueOrDash(value: string | null | undefined): string {
  const trimmed = value?.trim();
  return trimmed ? trimmed : "—";
}

function TelValue({ value }: { value: string | null }) {
  if (!value?.trim()) return "—";
  return (
    <a
      className={styles.companyHeroLink}
      href={`tel:${value.replace(/\s+/g, "")}`}
    >
      {value}
    </a>
  );
}

function EmailValue({ value }: { value: string | null }) {
  const email = value?.trim();
  if (!email) return "—";
  return (
    <a className={styles.companyHeroLink} href={`mailto:${email}`}>
      {email}
    </a>
  );
}

interface CustomerCompanyProfileCardProps {
  company: CustomerCompanyProfile;
}

/**
 * Customer Dashboard hero — large logo, then name, address, main contact,
 * tel no. and email. Remaining company metadata sits below.
 */
export function CustomerCompanyProfileCard({
  company,
}: CustomerCompanyProfileCardProps) {
  const initials = company.companyName
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 3)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("");

  const statusTone =
    company.status.toLowerCase() === "active" ? "ok" : "neutral";

  return (
    <section
      className={styles.companyHero}
      aria-label={`${company.companyName} — company profile`}
    >
      <div className={styles.companyHeroBand}>
        {company.companyLogo ? (
          // eslint-disable-next-line @next/next/no-img-element
          <img
            className={styles.companyHeroLogo}
            src={company.companyLogo}
            alt={`${company.companyName} logo`}
          />
        ) : (
          <div className={styles.companyHeroLogoPlaceholder} aria-hidden="true">
            {initials || "—"}
          </div>
        )}

        <div className={styles.companyHeroText}>
          <h1 className={styles.companyHeroTitle}>{company.companyName}</h1>

          <dl className={styles.companyHeroPrimary}>
            <div className={styles.companyHeroContactItem}>
              <dt className={styles.metaLabel}>Address</dt>
              <dd className={styles.companyHeroContactValue}>
                {valueOrDash(company.registeredAddress)}
              </dd>
            </div>
            <div className={styles.companyHeroContactItem}>
              <dt className={styles.metaLabel}>Main contact</dt>
              <dd className={styles.companyHeroContactValue}>
                {valueOrDash(company.mainContact)}
              </dd>
            </div>
            <div className={styles.companyHeroContactItem}>
              <dt className={styles.metaLabel}>Tel no.</dt>
              <dd className={styles.companyHeroContactValue}>
                <TelValue value={company.telNo} />
              </dd>
            </div>
            <div className={styles.companyHeroContactItem}>
              <dt className={styles.metaLabel}>Email</dt>
              <dd className={styles.companyHeroContactValue}>
                <EmailValue value={company.email} />
              </dd>
            </div>
          </dl>
        </div>
      </div>

      <div className={styles.companyHeroSecondary}>
        <span className={styles.companyHeroStatusRow}>
          <StatusBadge label={company.status} tone={statusTone} />
        </span>
        {company.companyNumber?.trim() ? (
          <span>
            <strong>Company no.</strong> {company.companyNumber.trim()}
          </span>
        ) : null}
        {company.companyRegNumber?.trim() ? (
          <span>
            <strong>Reg no.</strong> {company.companyRegNumber.trim()}
          </span>
        ) : null}
        {company.vatNo?.trim() ? (
          <span>
            <strong>VAT no.</strong> {company.vatNo.trim()}
          </span>
        ) : null}
        {company.companySize?.trim() ? (
          <span>
            <strong>Company size</strong> {company.companySize.trim()}
          </span>
        ) : null}
      </div>
    </section>
  );
}
