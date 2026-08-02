"use client";

import {
  formatExpiryCell,
  formatOutcomeCell,
  formatTextCell,
  TrainingRecordsTable,
  type TrainingRecordColumn,
} from "@/components/customer/TrainingRecordsTable";
import { formatDate } from "@/lib/utils/formatDate";
import type { CustomerNporsRecord } from "@/types/models";

const columns: TrainingRecordColumn<CustomerNporsRecord>[] = [
  {
    key: "candidateName",
    header: "Candidate Name",
    render: (row) => formatTextCell(row.candidateName),
  },
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
  records: CustomerNporsRecord[];
}

export function NporsRecordsView({ companyName, records }: Props) {
  return (
    <TrainingRecordsTable
      title="NPORS Training"
      description="Plant and machinery qualifications for your company workforce."
      companyName={companyName}
      records={records}
      columns={columns}
      getSearchText={(row) =>
        [
          row.candidateName,
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
  );
}
