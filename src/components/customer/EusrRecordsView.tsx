"use client";

import {
  formatExpiryCell,
  formatOutcomeCell,
  formatTextCell,
  TrainingRecordsTable,
  type TrainingRecordColumn,
} from "@/components/customer/TrainingRecordsTable";
import { expandEusrRecordsForDisplay } from "@/lib/training/eusrOptions";
import { formatDate } from "@/lib/utils/formatDate";
import type { CustomerEusrRecord } from "@/types/models";

const columns: TrainingRecordColumn<CustomerEusrRecord>[] = [
  {
    key: "candidateName",
    header: "Candidate Name",
    render: (row) => formatTextCell(row.candidateName),
  },
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
    header: "Outcome Pass/Fail",
    render: (row) => formatOutcomeCell(row.outcome),
  },
  {
    key: "expiry",
    header: "Expiry",
    render: (row) => formatExpiryCell(row.expiry),
  },
];

interface Props {
  companyName: string;
  records: CustomerEusrRecord[];
}

export function EusrRecordsView({ companyName, records }: Props) {
  return (
    <TrainingRecordsTable
      title="EUSR Training"
      description="EUSR registrations and outcomes for your company workforce."
      companyName={companyName}
      records={expandEusrRecordsForDisplay(records)}
      columns={columns}
      getSearchText={(row) =>
        [
          row.candidateName,
          row.eusrNumber,
          row.eusrCategory,
          row.cardType,
          row.trainingAddress,
          row.outcome,
        ]
          .filter(Boolean)
          .join(" ")
      }
      getOutcome={(row) => row.outcome}
      getExpiry={(row) => row.expiry}
      getWorkforceId={(row) => row.workforceId}
    />
  );
}
