"use client";

import Link from "next/link";

import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { StatusBadge } from "@/components/ui/StatusBadge";
import {
  formatDateCell,
  formatExpiryCell,
  formatOutcomeCell,
  formatTextCell,
  TrainingRecordsTable,
  type TrainingRecordColumn,
} from "@/components/customer/TrainingRecordsTable";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
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

const nporsColumns: TrainingRecordColumn<CustomerNporsRecord>[] = [
  {
    key: "nporsNumber",
    header: "NPORS Number",
    render: (row) => formatTextCell(row.nporsNumber),
  },
  {
    key: "trainingDate",
    header: "Training Date",
    render: (row) => formatDateCell(row.trainingDate),
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
    key: "trainingDate",
    header: "Training Date",
    render: (row) => formatDateCell(row.trainingDate),
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
    render: (row) => formatDateCell(row.trainingDate),
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
    header: "Course",
    render: (row) => formatTextCell(row.course),
  },
  {
    key: "trainingDate",
    header: "Training Date",
    render: (row) => formatDateCell(row.trainingDate),
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

function ProfileNvqSection({ records }: { records: CustomerNvqRecord[] }) {
  const active = records.filter((row) => row.status === "Active");
  const completed = records.filter((row) => row.status === "Completed");

  return (
    <section className={styles.profileSection} aria-label="NVQ Progress">
      <h2 className={styles.profileSectionTitle}>NVQ Progress</h2>
      <p className={styles.profileSectionMeta}>
        Active {active.length} · Completed {completed.length}
      </p>
      {records.length === 0 ? (
        <p className={styles.muted}>No customer-visible NVQ records.</p>
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
                      ? formatDisplayDate(row.dateRegistered)
                      : "—"}
                  </td>
                  <td>
                    {row.inductionDate
                      ? formatDisplayDate(row.inductionDate)
                      : "—"}
                  </td>
                  <td>{row.stageOfNvq?.trim() || "—"}</td>
                  <td>{row.notes?.trim() || "—"}</td>
                  <td>
                    {row.completedDate
                      ? formatDisplayDate(row.completedDate)
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
}: {
  records: CustomerDocumentRecord[];
}) {
  return (
    <section
      className={styles.profileSection}
      aria-label="Customer-visible documents"
    >
      <h2 className={styles.profileSectionTitle}>Documents</h2>
      {records.length === 0 ? (
        <p className={styles.muted}>No customer-visible documents.</p>
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
                      ? formatDisplayDate(row.uploadedDate)
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
  matrixReturnHref = "/customer",
  nporsRecords = [],
  eusrRecords = [],
  streetworksRecords = [],
  inHouseRecords = [],
  nvqRecords = [],
  documents = [],
}: Props) {
  const backToMatrix = safeReturnHref(matrixReturnHref);
  const companyName = candidate.companyName;

  return (
    <div>
      <header className={styles.pageHeader}>
        <Breadcrumbs
          items={[
            { label: "Customer", href: "/customer" },
            { label: "Training Matrix", href: backToMatrix },
            { label: candidate.candidateName },
          ]}
        />
        <p className={styles.eyebrow}>Candidate Profile</p>
        <h1 className={styles.title}>{candidate.candidateName}</h1>
        {candidate.dateOfBirth?.trim() ? (
          <p className={styles.dobSecondary}>
            DOB {formatDisplayDate(candidate.dateOfBirth)}
          </p>
        ) : null}
        <p className={styles.subtitle}>
          Training summary and customer-visible records for this candidate.
        </p>
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
        <Item label="Department" value={candidate.department} />
      </section>

      <section
        className={styles.profileGrid}
        aria-label="Training expiry summary"
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
          label="Streetworks (SWQR) expiry"
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
            [row.eusrNumber, row.eusrCategory, row.trainingAddress, row.outcome]
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

      <ProfileNvqSection records={nvqRecords} />
      <ProfileDocumentsSection records={documents} />

      <p className={styles.companyMeta} style={{ marginTop: "1.35rem" }}>
        <Link className={styles.link} href={backToMatrix}>
          Back to Training Matrix
        </Link>
        {" · "}
        <Link className={styles.link} href="/customer/training-records">
          View training records
        </Link>
      </p>
    </div>
  );
}
