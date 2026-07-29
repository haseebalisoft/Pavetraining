"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
  type AdminWorkforceOption,
} from "@/components/admin/AdminCrudPage";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import { getInHouseCourseOptions } from "@/lib/training/inHouseCourseOptions";
import { getNporsCategoryOptions } from "@/lib/training/nporsCategoryOptions";
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
  const people: AdminFieldConfig[] = [
    {
      name: "candidateName",
      label: "Candidate",
      type: "workforce",
      required: true,
      section: "Candidate",
    },
    {
      name: "companyName",
      label: "Company",
      type: "company",
      required: true,
      section: "Candidate",
    },
  ];

  const outcome: AdminFieldConfig[] = [
    {
      name: "trainingDate",
      label: "Training date",
      type: "date",
      section: "Training",
    },
    {
      name: "trainingAddress",
      label: "Training address",
      type: "text",
      section: "Training",
    },
    {
      name: "trainingOutcome",
      label: "Outcome",
      type: "select",
      section: "Outcome",
      options: [
        { value: "Pass", label: "Pass" },
        { value: "Fail", label: "Fail" },
      ],
    },
    {
      name: "outcomeDate",
      label: "Outcome date",
      type: "date",
      section: "Outcome",
    },
    { name: "expiry", label: "Expiry", type: "date", section: "Outcome" },
    {
      name: "assessorTrainer",
      label: "Assessor / trainer",
      type: "text",
      section: "Outcome",
    },
    {
      name: "outcomeNotes",
      label: "Outcome notes",
      type: "textarea",
      section: "Outcome",
    },
    {
      name: "notes",
      label: "Notes",
      type: "textarea",
      section: "Outcome",
    },
    {
      name: "customerVisible",
      label: "Customer visible",
      type: "boolean",
      section: "Outcome",
    },
  ];

  if (kind === "npors") {
    return [
      ...people,
      {
        name: "nporsNumber",
        label: "NPORS number",
        type: "text",
        section: "Training",
      },
      {
        name: "noviceOrEwt",
        label: "Novice or EWT",
        type: "select",
        section: "Training",
        options: [
          { value: "Novice", label: "Novice" },
          { value: "Ewt", label: "Ewt" },
        ],
      },
      {
        name: "nporsCategory",
        label: "NPORS category",
        type: "select",
        section: "Training",
        options: getNporsCategoryOptions(),
      },
      ...outcome,
    ];
  }
  if (kind === "eusr") {
    return [
      ...people,
      {
        name: "eusrNumber",
        label: "EUSR number",
        type: "text",
        section: "Training",
      },
      {
        name: "eusrCategory",
        label: "EUSR category",
        type: "text",
        section: "Training",
      },
      {
        name: "cardType",
        label: "Card type",
        type: "text",
        section: "Training",
      },
      ...outcome,
    ];
  }
  if (kind === "streetworks") {
    return [
      ...people,
      {
        name: "swqrNumber",
        label: "SWQR number",
        type: "text",
        section: "Training",
      },
      {
        name: "course",
        label: "Course",
        type: "select",
        section: "Training",
        options: [
          { value: "Operative", label: "Operative" },
          { value: "Supervisor", label: "Supervisor" },
        ],
      },
      {
        name: "streetworksCategory",
        label: "Streetworks category",
        type: "text",
        section: "Training",
      },
      ...outcome.filter((field) => field.name !== "notes"),
    ];
  }
  return [
    ...people,
    {
      name: "certificateCategory",
      label: "Certificate category",
      type: "text",
      section: "Training",
    },
    {
      name: "course",
      label: "Course",
      type: "select",
      section: "Training",
      options: getInHouseCourseOptions(),
    },
    ...outcome,
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
    base.push({
      key: "category",
      header: "Category",
      render: (row) => row.nporsCategory ?? "—",
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
  workforce,
  initialRows,
}: {
  kind: RegisterKind;
  companies: Company[];
  workforce: AdminWorkforceOption[];
  initialRows: AdminTrainingRecord[];
}) {
  return (
    <AdminCrudPage<AdminTrainingRecord>
      title={titles[kind]}
      description="Pick a Workforce candidate to auto-fill name and company (same as SharePoint). Saving a Pass/Fail record updates the Training Matrix for that candidate."
      columns={columnsFor(kind)}
      fields={fieldsFor(kind)}
      companies={companies}
      workforce={workforce}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      drawerWide
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
        (row) => row.nporsCategory,
        (row) => row.eusrNumber,
        (row) => row.swqrNumber,
        (row) => row.course,
        (row) => row.trainingAddress,
        (row) => row.assessorTrainer,
      ]}
    />
  );
}
