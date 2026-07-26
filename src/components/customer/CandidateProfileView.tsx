import Link from "next/link";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type {
  CustomerMatrixRecord,
  WorkforceCandidate,
} from "@/types/models";

import styles from "./customer.module.css";

interface Props {
  candidate: WorkforceCandidate;
  matrixRow?: CustomerMatrixRecord | null;
  /** Preserves Training Matrix filters when returning from a row click. */
  matrixReturnHref?: string;
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

function ExpiryItem({
  label,
  date,
}: {
  label: string;
  date: string | null | undefined;
}) {
  return (
    <div className={styles.profileCard}>
      <p className={styles.metaLabel}>{label}</p>
      <div className={styles.metaValue}>
        <ExpiryDateBadge date={date} />
      </div>
    </div>
  );
}

function safeReturnHref(value: string | null | undefined): string {
  if (!value?.trim()) return "/customer";
  if (!value.startsWith("/customer")) return "/customer";
  return value;
}

export function CandidateProfileView({
  candidate,
  matrixRow = null,
  matrixReturnHref = "/customer",
}: Props) {
  const backToMatrix = safeReturnHref(matrixReturnHref);

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

      <section
        className={styles.profileGrid}
        aria-label="Expiry summary"
        style={{ marginTop: "1.25rem" }}
      >
        <ExpiryItem
          label="Next expiry"
          date={matrixRow?.nextExpiryDate ?? null}
        />
        <ExpiryItem
          label="NPORS expiry"
          date={matrixRow?.nporsExpiry ?? null}
        />
        <ExpiryItem
          label="CSCS expiry"
          date={matrixRow?.cscsExpiry ?? candidate.cscsExpiry}
        />
        <ExpiryItem
          label="SWQR expiry"
          date={matrixRow?.swqrExpiry ?? candidate.swqrExpiry}
        />
        <ExpiryItem
          label="EUSR expiry"
          date={matrixRow?.eusrExpiry ?? candidate.eusrExpiry}
        />
        <ExpiryItem
          label="In-House expiry"
          date={matrixRow?.inHouseExpiry ?? null}
        />
        <ExpiryItem label="N001 expiry" date={matrixRow?.n001Expiry} />
        <ExpiryItem label="N010 expiry" date={matrixRow?.n010Expiry} />
        <ExpiryItem label="N020 expiry" date={matrixRow?.n020Expiry} />
        <ExpiryItem label="N003 expiry" date={matrixRow?.n003Expiry} />
        <ExpiryItem label="N100 expiry" date={matrixRow?.n100Expiry} />
      </section>

      <p className={styles.companyMeta} style={{ marginTop: "1.35rem" }}>
        <Link className={styles.link} href={backToMatrix}>
          Back to training matrix
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
