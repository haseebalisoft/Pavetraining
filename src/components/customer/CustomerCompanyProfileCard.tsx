import Link from "next/link";

import type { CustomerCompanyProfile } from "@/types/models";

import { DashboardIcon, type DashboardIconName } from "./DashboardIcons";
import styles from "./customerDashboard.module.css";

function companyInitials(name: string): string {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 3)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("");
}

function DetailItem({
  icon,
  label,
  children,
}: {
  icon: DashboardIconName;
  label: string;
  children: React.ReactNode;
}) {
  return (
    <div className={styles.profileDetail}>
      <span className={styles.profileDetailIcon}>
        <DashboardIcon name={icon} />
      </span>
      <div>
        <dt>{label}</dt>
        <dd>{children}</dd>
      </div>
    </div>
  );
}

interface CustomerCompanyProfileCardProps {
  company: CustomerCompanyProfile;
}

/**
 * Customer Dashboard company card — real company logo (or initials tile),
 * name + status, address / tel / main contact / email, and support actions.
 */
export function CustomerCompanyProfileCard({
  company,
}: CustomerCompanyProfileCardProps) {
  const isActive = company.status.trim().toLowerCase() === "active";
  const tel = company.telNo?.trim();
  const email = company.email?.trim();

  return (
    <section
      className={styles.profileCard}
      aria-label={`${company.companyName} — company profile`}
    >
      <div className={styles.profileLogoWrap}>
        {company.companyLogo ? (
          // eslint-disable-next-line @next/next/no-img-element
          <img
            className={styles.profileLogo}
            src={company.companyLogo}
            alt={`${company.companyName} logo`}
          />
        ) : (
          <div className={styles.profileLogoPlaceholder} aria-hidden="true">
            {companyInitials(company.companyName) || "—"}
          </div>
        )}
      </div>

      <div className={styles.profileBody}>
        <p className={styles.eyebrow}>Company profile</p>
        <div className={styles.profileTitleRow}>
          <h2 className={styles.profileTitle}>{company.companyName}</h2>
          {company.status.trim() ? (
            <span
              className={
                isActive ? styles.statusPillActive : styles.statusPillNeutral
              }
            >
              {company.status}
            </span>
          ) : null}
        </div>

        <dl className={styles.profileDetails}>
          <DetailItem icon="pin" label="Address">
            {company.registeredAddress?.trim() || "—"}
          </DetailItem>
          <DetailItem icon="person" label="Main contact">
            {company.mainContact?.trim() || "—"}
          </DetailItem>
          <DetailItem icon="phone" label="Tel">
            {tel ? <a href={`tel:${tel.replace(/\s+/g, "")}`}>{tel}</a> : "—"}
          </DetailItem>
          <DetailItem icon="mail" label="Email">
            {email ? <a href={`mailto:${email}`}>{email}</a> : "—"}
          </DetailItem>
        </dl>

        <p className={styles.profileMeta}>
          {company.companyNumber?.trim() ? (
            <span>Company No: {company.companyNumber.trim()}</span>
          ) : null}
          {company.companyRegNumber?.trim() ? (
            <span>Reg No: {company.companyRegNumber.trim()}</span>
          ) : null}
          {company.vatNo?.trim() ? <span>VAT No: {company.vatNo.trim()}</span> : null}
          {company.companySize?.trim() ? (
            <span>Size: {company.companySize.trim()}</span>
          ) : null}
        </p>
      </div>

      <div className={styles.profileActions}>
        <Link href="/customer/support" className={styles.buttonOutline}>
          Edit request
        </Link>
        <Link href="/customer/support" className={styles.buttonPrimary}>
          Contact support
        </Link>
      </div>
    </section>
  );
}
