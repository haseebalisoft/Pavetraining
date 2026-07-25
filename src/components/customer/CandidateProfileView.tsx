import Link from "next/link";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type { WorkforceCandidate } from "@/types/models";

import styles from "./customer.module.css";

interface Props {
  candidate: WorkforceCandidate;
}

function Item({
  label,
  value,
}: {
  label: string;
  value: string | null | undefined;
}) {
  return (
    <div className={styles.profileCard}>
      <p className={styles.metaLabel}>{label}</p>
      <p className={styles.metaValue}>
        {value?.trim() ? value : <span className={styles.muted}>—</span>}
      </p>
    </div>
  );
}

export function CandidateProfileView({ candidate }: Props) {
  return (
    <div>
      <header className={styles.pageHeader}>
        <Breadcrumbs
          items={[
            { label: "Customer", href: "/customer" },
            { label: "Candidates", href: "/customer/candidates" },
            { label: candidate.candidateName },
          ]}
        />
        <p className={styles.eyebrow}>Candidate profile</p>
        <h1 className={styles.title}>{candidate.candidateName}</h1>
        <p className={styles.subtitle}>
          Workforce details for your organisation. Sensitive fields stay
          de-emphasized.
        </p>
      </header>

      <p className={styles.companyMeta}>
        <StatusBadge
          label={candidate.status?.trim() || "Unknown"}
          tone={
            (candidate.status ?? "").toLowerCase() === "active" ? "ok" : "neutral"
          }
        />{" "}
        · {candidate.companyName}
      </p>

      <section className={styles.profileGrid} aria-label="Candidate details">
        <Item label="Workforce number" value={candidate.workforceNumber} />
        <Item label="Department" value={candidate.department} />
        <Item label="Training manager" value={candidate.trainingManager} />
        <Item label="Supervisor" value={candidate.supervisor} />
        <Item
          label="Date of birth"
          value={
            candidate.dateOfBirth
              ? formatDisplayDate(candidate.dateOfBirth)
              : null
          }
        />
        <Item label="CSCS number" value={candidate.cscsNumber} />
        <Item label="SWQR number" value={candidate.swqrNumber} />
        <Item label="EUSR number" value={candidate.eusrNumber} />
        <Item label="NPORS numbers" value={candidate.nporsNumbers} />
        <Item
          label="In-house certification"
          value={candidate.inHouseCertificationNumber}
        />
      </section>

      <p className={styles.companyMeta} style={{ marginTop: "1.35rem" }}>
        <Link className={styles.link} href="/customer/training-matrix">
          View training matrix
        </Link>
        {" · "}
        <Link className={styles.link} href="/customer/training-records">
          View training records
        </Link>
        {" · "}
        <Link className={styles.link} href="/customer/candidates">
          Back to candidates
        </Link>
      </p>
    </div>
  );
}
