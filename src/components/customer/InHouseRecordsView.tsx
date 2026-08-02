"use client";

import {
  formatExpiryCell,
  formatOutcomeCell,
  formatTextCell,
  TrainingRecordsTable,
  type TrainingRecordColumn,
} from "@/components/customer/TrainingRecordsTable";
import { formatDate } from "@/lib/utils/formatDate";
import type { CustomerInHouseRecord } from "@/types/models";

const columns: TrainingRecordColumn<CustomerInHouseRecord>[] = [
  {
    key: "candidateName",
    header: "Candidate Name",
    render: (row) => formatTextCell(row.candidateName),
  },
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
    header: "Outcome Pass/Fail",
    render: (row) => formatOutcomeCell(row.outcome),
  },
];

interface Props {
  companyName: string;
  records: CustomerInHouseRecord[];
}

export function InHouseRecordsView({ companyName, records }: Props) {
  return (
    <TrainingRecordsTable
      title="In-House Training"
      description="In-house certificates and outcomes for your company workforce."
      companyName={companyName}
      records={records}
      columns={columns}
      getSearchText={(row) =>
        [
          row.candidateName,
          row.course,
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
