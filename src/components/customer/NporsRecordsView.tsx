"use client";

import {
  formatExpiryCell,
  formatOutcomeCell,
  formatTextCell,
  TrainingRecordsTable,
  type TrainingRecordColumn,
} from "@/components/customer/TrainingRecordsTable";
import { splitNporsCategory } from "@/lib/training/nporsCategoryOptions";
import { formatDate } from "@/lib/utils/formatDate";
import type { CustomerNporsRecord } from "@/types/models";

const columns: TrainingRecordColumn<CustomerNporsRecord>[] = [
  {
    key: "candidateName",
    header: "Candidate Name",
    render: (row) => formatTextCell(row.candidateName),
  },
  {
    key: "categoryName",
    header: "Category name",
    render: (row) => {
      const split = splitNporsCategory(row.nporsCategory);
      return formatTextCell(split.categoryName || row.nporsCategory);
    },
  },
  {
    key: "nNumber",
    header: "N number",
    render: (row) => {
      const split = splitNporsCategory(row.nporsCategory);
      return formatTextCell(split.nNumber || row.nporsNumber);
    },
  },
  {
    key: "expiry",
    header: "Expiry date",
    render: (row) => formatExpiryCell(row.expiry),
  },
  {
    key: "trainingDate",
    header: "Training Date",
    render: (row) => formatDate(row.trainingDate),
  },
  {
    key: "nporsNumber",
    header: "NPORS card number",
    render: (row) => formatTextCell(row.nporsNumber),
  },
  {
    key: "outcome",
    header: "Outcome Pass/Fail",
    render: (row) => formatOutcomeCell(row.outcome),
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
