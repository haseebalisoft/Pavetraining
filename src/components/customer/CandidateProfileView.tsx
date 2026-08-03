"use client";

import Link from "next/link";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { StatusBadge } from "@/components/ui/StatusBadge";
import {
  formatExpiryCell,
  formatOutcomeCell,
  formatTextCell,
  TrainingRecordsTable,
  type TrainingRecordColumn,
} from "@/components/customer/TrainingRecordsTable";
import { formatDate } from "@/lib/utils/formatDate";
import type {
  CustomerDocumentRecord,
  CustomerEusrRecord,
  CustomerInHouseRecord,
  CustomerMatrixRecord,
  CustomerNporsRecord,
  CustomerNvqRecord,
  CustomerStreetworksRecord,
  WorkforceCandidate,
} from "@/types/models";

import styles from "./customer.module.css";

interface Props {
  candidate: WorkforceCandidate;
  matrixRow?: CustomerMatrixRecord | null;
  /** Preserves Training Matrix filters when returning from a row click. */
  matrixReturnHref?: string;
  nporsRecords?: CustomerNporsRecord[];
  eusrRecords?: CustomerEusrRecord[];
  streetworksRecords?: CustomerStreetworksRecord[];
  inHouseRecords?: CustomerInHouseRecord[];
  nvqRecords?: CustomerNvqRecord[];
  documents?: CustomerDocumentRecord[];
  /** Admin audit profile uses the same layout with admin nav/copy. */
  variant?: "customer" | "admin";
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

function safeReturnHref(
  value: string | null | undefined,
  variant: "customer" | "admin",
): string {
  const fallback = variant === "admin" ? "/admin/workforce" : "/customer";
  const prefix = variant === "admin" ? "/admin" : "/customer";
  if (!value?.trim()) return fallback;
  if (!value.startsWith(prefix)) return fallback;
  return value;
}

const nporsColumns: TrainingRecordColumn<CustomerNporsRecord>[] = [
  {
    key: "nporsNumber",
    header: "NPORS Number",
    render: (row) => formatTextCell(row.nporsNumber),
  },
  {
    key: "trainingDate",
    header: "Training Date",
    render: (row) => formatDate(row.trainingDate),
  },
  {
    key: "trainingAddress",
    header: "Training Address",
    render: (row) => formatTextCell(row.trainingAddress),
  },
  {
    key: "noviceOrEwt",
    header: "Novice or EWT",
    render: (row) => formatTextCell(row.noviceOrEwt),
  },
  {
    key: "nporsCategory",
    header: "NPORS Category",
    render: (row) => formatTextCell(row.nporsCategory),
  },
  {
    key: "outcome",
    header: "Pass / Fail",
    render: (row) => formatOutcomeCell(row.outcome),
  },
];

const eusrColumns: TrainingRecordColumn<CustomerEusrRecord>[] = [
  {
    key: "eusrNumber",
    header: "EUSR Number",
    render: (row) => formatTextCell(row.eusrNumber),
  },
  {
    key: "eusrCategory",
    header: "EUSR Category",
    render: (row) => formatTextCell(row.eusrCategory),
  },
  {
    key: "cardType",
    header: "Card Type",
    render: (row) => formatTextCell(row.cardType),
  },
  {
    key: "trainingDate",
    header: "Training Date",
    render: (row) => formatDate(row.trainingDate),
  },
  {
    key: "trainingAddress",
    header: "Training Address",
    render: (row) => formatTextCell(row.trainingAddress),
  },
  {
    key: "outcome",
    header: "Pass / Fail",
    render: (row) => formatOutcomeCell(row.outcome),
  },
];

const streetworksColumns: TrainingRecordColumn<CustomerStreetworksRecord>[] = [
  {
    key: "swqrNumber",
    header: "SWQR Number",
    render: (row) => formatTextCell(row.swqrNumber),
  },
  {
    key: "trainingDate",
    header: "Training Date",
    render: (row) => {
      const from = formatDate(row.trainingDate);
      if (row.trainingDateEnd?.trim()) {
        return `${from} → ${formatDate(row.trainingDateEnd)}`;
      }
      return from;
    },
  },
  {
    key: "trainingAddress",
    header: "Training Address",
    render: (row) => formatTextCell(row.trainingAddress),
  },
  {
    key: "course",
    header: "Course",
    render: (row) => formatTextCell(row.course),
  },
  {
    key: "streetworksCategory",
    header: "Streetworks Category",
    render: (row) => formatTextCell(row.streetworksCategory),
  },
  {
    key: "outcome",
    header: "Pass / Fail",
    render: (row) => formatOutcomeCell(row.outcome),
  },
];

const inHouseColumns: TrainingRecordColumn<CustomerInHouseRecord>[] = [
  {
    key: "course",
    header: "Certificate Category",
    render: (row) => formatTextCell(row.course),
  },
  {
    key: "trainingDate",
    header: "Training Date",
    render: (row) => formatDate(row.trainingDate),
  },
  {
    key: "trainingAddress",
    header: "Training Address",
    render: (row) => formatTextCell(row.trainingAddress),
  },
  {
    key: "expiry",
    header: "Expiry Date",
    render: (row) => formatExpiryCell(row.expiry),
  },
  {
    key: "outcome",
    header: "Pass / Fail",
    render: (row) => formatOutcomeCell(row.outcome),
  },
];

function ProfileNvqSection({
  records,
  emptyLabel,
}: {
  records: CustomerNvqRecord[];
  emptyLabel: string;
}) {
  const active = records.filter((row) => row.status === "Active");
  const completed = records.filter((row) => row.status === "Completed");

  return (
    <section className={styles.profileSection} aria-label="NVQ Progress">
      <h2 className={styles.profileSectionTitle}>NVQ Progress</h2>
      <p className={styles.profileSectionMeta}>
        Active {active.length} · Completed {completed.length}
      </p>
      {records.length === 0 ? (
        <p className={styles.muted}>{emptyLabel}</p>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">NVQ Title</th>
                <th scope="col">Status</th>
                <th scope="col">Bolt On</th>
                <th scope="col">Date Registered</th>
                <th scope="col">Induction Date</th>
                <th scope="col">Stage of NVQ</th>
                <th scope="col">Notes</th>
                <th scope="col">Completed Date</th>
              </tr>
            </thead>
            <tbody>
              {records.map((row) => (
                <tr
                  key={row.id}
                  className={
                    row.status === "Completed" ? styles.completedRow : undefined
                  }
                >
                  <td>{row.nvqTitle?.trim() || "—"}</td>
                  <td>
                    <StatusBadge
                      label={row.status}
                      tone={row.status === "Completed" ? "ok" : "info"}
                    />
                  </td>
                  <td>{row.boltOn?.trim() || "—"}</td>
                  <td>
                    {row.dateRegistered
                      ? formatDate(row.dateRegistered)
                      : "—"}
                  </td>
                  <td>
                    {row.inductionDate
                      ? formatDate(row.inductionDate)
                      : "—"}
                  </td>
                  <td>{row.stageOfNvq?.trim() || "—"}</td>
                  <td>{row.notes?.trim() || "—"}</td>
                  <td>
                    {row.completedDate
                      ? formatDate(row.completedDate)
                      : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

function ProfileDocumentsSection({
  records,
  emptyLabel,
}: {
  records: CustomerDocumentRecord[];
  emptyLabel: string;
}) {
  return (
    <section
      className={styles.profileSection}
      aria-label="Documents"
    >
      <h2 className={styles.profileSectionTitle}>Documents</h2>
      {records.length === 0 ? (
        <p className={styles.muted}>{emptyLabel}</p>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">Document name</th>
                <th scope="col">Document type</th>
                <th scope="col">Modified date</th>
                <th scope="col">View</th>
                <th scope="col">Download</th>
              </tr>
            </thead>
            <tbody>
              {records.map((row) => (
                <tr key={row.id}>
                  <td>{row.name}</td>
                  <td>{row.documentType?.trim() || "—"}</td>
                  <td>
                    {row.uploadedDate
                      ? formatDate(row.uploadedDate)
                      : "—"}
                  </td>
                  <td>
                    {row.viewPath ? (
                      <a
                        className={styles.link}
                        href={row.viewPath}
                        target="_blank"
                        rel="noreferrer"
                      >
                        View
                      </a>
                    ) : (
                      <span className={styles.muted}>—</span>
                    )}
                  </td>
                  <td>
                    {row.canDownload && row.downloadPath ? (
                      <a className={styles.link} href={row.downloadPath}>
                        Download
                      </a>
                    ) : (
                      <span className={styles.muted}>—</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

export function CandidateProfileView({
  candidate,
  matrixRow = null,
  matrixReturnHref,
  nporsRecords = [],
  eusrRecords = [],
  streetworksRecords = [],
  inHouseRecords = [],
  nvqRecords = [],
  documents = [],
  variant = "customer",
}: Props) {
  const isAdmin = variant === "admin";
  const backHref = safeReturnHref(
    matrixReturnHref ?? (isAdmin ? "/admin/workforce" : "/customer"),
    variant,
  );
  const companyName = candidate.companyName;

  return (
    <div>
      <header className={styles.pageHeader}>
        <Breadcrumbs
          items={
            isAdmin
              ? [
                  { label: "Admin", href: "/admin" },
                  { label: "Workforce", href: "/admin/workforce" },
                  { label: candidate.candidateName },
                ]
              : [
                  { label: "Customer", href: "/customer" },
                  { label: "Training Matrix", href: backHref },
                  { label: candidate.candidateName },
                ]
          }
        />
        <div className={styles.profileHero}>
          {candidate.photoUrl ? (
            // eslint-disable-next-line @next/next/no-img-element
            <img
              className={styles.candidatePhoto}
              src={candidate.photoUrl}
              alt={candidate.candidateName}
            />
          ) : (
            <div className={styles.candidatePhotoPlaceholder} aria-hidden>
              {candidate.candidateName
                .split(/\s+/)
                .filter(Boolean)
                .slice(0, 2)
                .map((part) => part[0]?.toUpperCase() ?? "")
                .join("")}
            </div>
          )}
          <div>
            <p className={styles.eyebrow}>
              {isAdmin ? "Admin Candidate Profile" : "Candidate Profile"}
            </p>
            <h1 className={styles.title}>{candidate.candidateName}</h1>
            {candidate.dateOfBirth?.trim() ? (
              <p className={styles.dobSecondary}>
                DOB {formatDate(candidate.dateOfBirth)}
              </p>
            ) : null}
            <p className={styles.subtitle}>
              {isAdmin
                ? "Full training history for audit — includes records hidden from customers."
                : "Training summary and customer-visible records for this candidate."}
            </p>
          </div>
        </div>
      </header>

      <p className={styles.companyMeta}>
        <StatusBadge
          label={candidate.status?.trim() || "Unknown"}
          tone={
            (candidate.status ?? "").toLowerCase() === "active" ? "ok" : "neutral"
          }
        />{" "}
        · {companyName}
      </p>

      <section className={styles.profileGrid} aria-label="Candidate details">
        <Item label="Company name" value={candidate.companyName} />
        <Item label="Department" value={candidate.department} />
        <Item label="Training manager" value={candidate.trainingManager} />
        <Item label="Supervisor" value={candidate.supervisor} />
        <Item
          label="Workforce number"
          value={candidate.workforceNumber}
        />
        <Item
          label="Date of birth"
          value={
            candidate.dateOfBirth
              ? formatDate(candidate.dateOfBirth)
              : null
          }
        />
        <Item label="NPORS number" value={candidate.nporsNumbers} />
        <Item label="CSCS number" value={candidate.cscsNumber} />
        <Item label="SWQR number" value={candidate.swqrNumber} />
        <Item label="EUSR number" value={candidate.eusrNumber} />
        <Item
          label="In-House certification number"
          value={candidate.inHouseCertificationNumber}
        />
      </section>

      <section
        className={styles.profileGrid}
        aria-label="Training expiry summary"
        style={{ marginTop: "1.25rem" }}
      >
        <ExpiryItem
          label="In-House / Asbestos Awareness expiry"
          date={matrixRow?.inHouseExpiry ?? null}
        />
        <ExpiryItem
          label="CSCS expiry"
          date={matrixRow?.cscsExpiry ?? candidate.cscsExpiry}
        />
        <ExpiryItem
          label="Streetworks (SWQR) expiry"
          date={matrixRow?.swqrExpiry ?? candidate.swqrExpiry}
        />
        <ExpiryItem
          label="EUSR expiry"
          date={matrixRow?.eusrExpiry ?? candidate.eusrExpiry}
        />
        <ExpiryItem
          label="NPORS expiry"
          date={matrixRow?.nporsExpiry ?? null}
        />
        <ExpiryItem
          label="Next expiry"
          date={matrixRow?.nextExpiryDate ?? null}
        />
      </section>

      <div className={styles.profileSection}>
        <TrainingRecordsTable
          title="NPORS Training"
          description="Plant and machinery qualifications for this candidate."
          companyName={companyName}
          records={nporsRecords}
          columns={nporsColumns}
          embedded
          getSearchText={(row) =>
            [
              row.nporsNumber,
              row.trainingAddress,
              row.noviceOrEwt,
              row.nporsCategory,
              row.outcome,
            ]
              .filter(Boolean)
              .join(" ")
          }
          getOutcome={(row) => row.outcome}
          getExpiry={(row) => row.expiry}
          getWorkforceId={(row) => row.workforceId}
        />
      </div>

      <div className={styles.profileSection}>
        <TrainingRecordsTable
          title="EUSR Training"
          description="EUSR registrations for this candidate."
          companyName={companyName}
          records={eusrRecords}
          columns={eusrColumns}
          embedded
          getSearchText={(row) =>
            [row.eusrNumber, row.eusrCategory, row.cardType, row.trainingAddress, row.outcome]
              .filter(Boolean)
              .join(" ")
          }
          getOutcome={(row) => row.outcome}
          getExpiry={(row) => row.expiry}
          getWorkforceId={(row) => row.workforceId}
        />
      </div>

      <div className={styles.profileSection}>
        <TrainingRecordsTable
          title="Streetworks Training"
          description="SWQR and streetworks course records for this candidate."
          companyName={companyName}
          records={streetworksRecords}
          columns={streetworksColumns}
          embedded
          getSearchText={(row) =>
            [
              row.swqrNumber,
              row.trainingAddress,
              row.course,
              row.streetworksCategory,
              row.outcome,
            ]
              .filter(Boolean)
              .join(" ")
          }
          getOutcome={(row) => row.outcome}
          getExpiry={(row) => row.expiry}
          getWorkforceId={(row) => row.workforceId}
        />
      </div>

      <div className={styles.profileSection}>
        <TrainingRecordsTable
          title="In-House Training"
          description="In-house certificates for this candidate."
          companyName={companyName}
          records={inHouseRecords}
          columns={inHouseColumns}
          embedded
          getSearchText={(row) =>
            [row.course, row.trainingAddress, row.outcome]
              .filter(Boolean)
              .join(" ")
          }
          getOutcome={(row) => row.outcome}
          getExpiry={(row) => row.expiry}
          getWorkforceId={(row) => row.workforceId}
        />
      </div>

      <ProfileNvqSection
        records={nvqRecords}
        emptyLabel={
          isAdmin ? "No NVQ records." : "No customer-visible NVQ records."
        }
      />
      <ProfileDocumentsSection
        records={documents}
        emptyLabel={
          isAdmin ? "No documents." : "No customer-visible documents."
        }
      />

      <p className={styles.companyMeta} style={{ marginTop: "1.35rem" }}>
        {isAdmin ? (
          <>
            <Link className={styles.link} href={backHref}>
              Back to Workforce
            </Link>
            {" · "}
            <Link className={styles.link} href="/admin/training-matrix">
              Training Matrix
            </Link>
          </>
        ) : (
          <>
            <Link className={styles.link} href={backHref}>
              Back to Training Matrix
            </Link>
            {" · "}
            <Link className={styles.link} href="/customer/training-records">
              View training records
            </Link>
          </>
        )}
      </p>
    </div>
  );
}
