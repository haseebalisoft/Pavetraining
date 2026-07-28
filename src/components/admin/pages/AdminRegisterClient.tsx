"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
} from "@/components/admin/AdminCrudPage";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { getInHouseCourseOptions } from "@/lib/training/inHouseCourseOptions";
import type { AdminTrainingRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

type RegisterKind = "npors" | "eusr" | "streetworks" | "in-house";

const titles: Record<RegisterKind, string> = {
  npors: "NPORS Training",
  eusr: "EUSR Training",
  streetworks: "Streetworks Training",
  "in-house": "In-House Training",
};

function fieldsFor(kind: RegisterKind): AdminFieldConfig[] {
  const sharedStart: AdminFieldConfig[] = [
    { name: "candidateName", label: "Candidate name", type: "text", required: true },
    { name: "companyName", label: "Company", type: "company", required: true },
  ];
  const sharedEnd: AdminFieldConfig[] = [
    { name: "trainingDate", label: "Training date", type: "date" },
    { name: "trainingAddress", label: "Training address", type: "text" },
    {
      name: "trainingOutcome",
      label: "Outcome",
      type: "select",
      options: [
        { value: "Pass", label: "Pass" },
        { value: "Fail", label: "Fail" },
      ],
    },
    { name: "expiry", label: "Expiry", type: "date" },
    { name: "customerVisible", label: "Customer visible", type: "boolean" },
  ];

  if (kind === "npors") {
    return [
      ...sharedStart,
      { name: "nporsNumber", label: "NPORS number", type: "text" },
      { name: "noviceOrEwt", label: "Novice or EWT", type: "text" },
      { name: "nporsCategory", label: "NPORS category", type: "text" },
      ...sharedEnd,
    ];
  }
  if (kind === "eusr") {
    return [
      ...sharedStart,
      { name: "eusrNumber", label: "EUSR number", type: "text" },
      { name: "eusrCategory", label: "EUSR category", type: "text" },
      ...sharedEnd,
    ];
  }
  if (kind === "streetworks") {
    return [
      ...sharedStart,
      { name: "swqrNumber", label: "SWQR number", type: "text" },
      {
        name: "course",
        label: "Course",
        type: "select",
        options: [
          { value: "Operative", label: "Operative" },
          { value: "Supervisor", label: "Supervisor" },
        ],
      },
      { name: "streetworksCategory", label: "Streetworks category", type: "text" },
      ...sharedEnd,
    ];
  }
  return [
    ...sharedStart,
    {
      name: "course",
      label: "Course",
      type: "select",
      options: getInHouseCourseOptions(),
    },
    ...sharedEnd,
  ];
}

function columnsFor(kind: RegisterKind): AdminColumn<AdminTrainingRecord>[] {
  const base: AdminColumn<AdminTrainingRecord>[] = [
    { key: "name", header: "Candidate name", render: (row) => row.candidateName },
    { key: "company", header: "Company", render: (row) => row.companyName },
  ];
  if (kind === "npors") {
    base.push({
      key: "number",
      header: "NPORS number",
      render: (row) => row.nporsNumber ?? "—",
    });
  }
  if (kind === "eusr") {
    base.push({
      key: "number",
      header: "EUSR number",
      render: (row) => row.eusrNumber ?? "—",
    });
  }
  if (kind === "streetworks") {
    base.push({
      key: "number",
      header: "SWQR number",
      render: (row) => row.swqrNumber ?? "—",
    });
  }
  if (kind === "in-house") {
    base.push({
      key: "course",
      header: "Course",
      render: (row) => row.course ?? "—",
    });
  }
  base.push(
    {
      key: "outcome",
      header: "Outcome",
      render: (row) => row.trainingOutcome ?? "—",
    },
    {
      key: "visible",
      header: "Customer visible",
      render: (row) => (row.customerVisible ? "Yes" : "No"),
    },
    {
      key: "expiry",
      header: "Expiry",
      render: (row) => <ExpiryDateBadge date={row.expiry} />,
    },
  );
  return base;
}

export function AdminRegisterClient({
  kind,
  companies,
  initialRows,
}: {
  kind: RegisterKind;
  companies: Company[];
  initialRows: AdminTrainingRecord[];
}) {
  return (
    <AdminCrudPage<AdminTrainingRecord>
      title={titles[kind]}
      description="Add and edit training records, outcomes, visibility, and addresses."
      columns={columnsFor(kind)}
      fields={fieldsFor(kind)}
      companies={companies}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      listUrl={`/api/admin/training-records/${kind}`}
      createUrl={`/api/admin/training-records/${kind}`}
      updateUrl={(id) => `/api/admin/training-records/${kind}/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminTrainingRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.candidateName,
        (row) => row.companyName,
        (row) => row.nporsNumber,
        (row) => row.eusrNumber,
        (row) => row.swqrNumber,
        (row) => row.course,
        (row) => row.trainingAddress,
      ]}
    />
  );
}
