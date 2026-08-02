"use client";

import {
  formatExpiryCell,
  formatOutcomeCell,
  formatTextCell,
  TrainingRecordsTable,
  type TrainingRecordColumn,
} from "@/components/customer/TrainingRecordsTable";
import { formatDate } from "@/lib/utils/formatDate";
import type { CustomerStreetworksRecord } from "@/types/models";

const columns: TrainingRecordColumn<CustomerStreetworksRecord>[] = [
  {
    key: "candidateName",
    header: "Candidate Name",
    render: (row) => formatTextCell(row.candidateName),
  },
  {
    key: "swqrNumber",
    header: "SWQR Number",
    render: (row) => formatTextCell(row.swqrNumber),
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
  records: CustomerStreetworksRecord[];
}

export function StreetworksRecordsView({ companyName, records }: Props) {
  return (
    <TrainingRecordsTable
      title="Streetworks Training"
      description="SWQR and streetworks course records for your company workforce."
      companyName={companyName}
      records={records}
      columns={columns}
      getSearchText={(row) =>
        [
          row.candidateName,
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
  );
}
