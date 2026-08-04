"use client";

import {
  AdminCrudPage,
  type AdminColumn,
  type AdminFieldConfig,
  type AdminWorkforceOption,
} from "@/components/admin/AdminCrudPage";
import { ExpiryDateBadge } from "@/components/ui/ExpiryDateBadge";
import {
  getEusrCardTypeOptions,
  getEusrCategoryOptions,
} from "@/lib/training/eusrOptions";
import { getInHouseCourseOptions } from "@/lib/training/inHouseCourseOptions";
import { getNporsCategoryOptions } from "@/lib/training/nporsCategoryOptions";
import {
  getStreetworksCategoryOptions,
  getStreetworksCourseOptions,
} from "@/lib/training/streetworksOptions";
import type { AdminTrainingRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

type RegisterKind = "npors" | "eusr" | "streetworks" | "in-house";

const titles: Record<RegisterKind, string> = {
  npors: "NPORS Training",
  eusr: "EUSR Training",
  streetworks: "Streetworks Training",
  "in-house": "In-House Training",
};

const descriptions: Record<RegisterKind, string> = {
  npors:
    "Select company, then candidate — Workforce / NPORS numbers fill from Workforce. Pass updates the Training Matrix and profile.",
  eusr:
    "Select company, then candidate — Workforce / EUSR numbers fill from Workforce. Pass updates the Training Matrix and profile.",
  streetworks:
    "Select company, then candidate — Workforce / SWQR numbers fill from Workforce. Pass updates the Training Matrix and profile.",
  "in-house":
    "Select company, then candidate — certification number fills from Workforce. Asbestos Awareness Pass syncs to N031 on the Matrix; other courses stay standalone.",
};

const workforceNumberField: AdminFieldConfig = {
  name: "workforceNumber",
  label: "Workforce number (from Workforce)",
  type: "text",
  readOnly: true,
  section: "Candidate",
};

function fieldsFor(
  kind: RegisterKind,
  nporsCategoryOptions?: Array<{ value: string; label: string }>,
): AdminFieldConfig[] {
  const people: AdminFieldConfig[] = [
    {
      name: "companyName",
      label: "Company",
      type: "company",
      required: true,
      section: "Candidate",
    },
    {
      name: "candidateName",
      label: "Candidate",
      type: "workforce",
      required: true,
      section: "Candidate",
    },
    workforceNumberField,
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
    const categoryOptions = (
      nporsCategoryOptions?.length
        ? nporsCategoryOptions
        : getNporsCategoryOptions()
    ).filter((option) => {
      const label = option.label.trim();
      const value = option.value.trim();
      // Prefer full "N001 - Title" labels; drop incomplete code-only text.
      if (!label) return false;
      if (/^N\d+[A-Z]?$/i.test(label) && label.toUpperCase() === value.toUpperCase()) {
        return false;
      }
      return true;
    });

    return [
      ...people,
      {
        name: "nporsNumber",
        label: "NPORS number (from Workforce)",
        type: "text",
        readOnly: true,
        section: "Training",
      },
      {
        name: "noviceOrEwt",
        label: "Novice or EWT",
        type: "select",
        section: "Training",
        options: [
          { value: "Novice", label: "Novice" },
          { value: "Ewt", label: "EWT" },
        ],
      },
      {
        name: "nporsCategory",
        label: "NPORS category",
        type: "multiselect",
        section: "Training",
        options: categoryOptions.length
          ? categoryOptions
          : getNporsCategoryOptions(),
      },
      ...outcome,
    ];
  }
  if (kind === "eusr") {
    return [
      ...people,
      {
        name: "eusrNumber",
        label: "EUSR number (from Workforce)",
        type: "text",
        readOnly: true,
        section: "Training",
      },
      {
        name: "eusrCategory",
        label: "EUSR category",
        type: "multiselect",
        section: "Training",
        options: getEusrCategoryOptions(),
      },
      {
        name: "cardType",
        label: "Card type",
        type: "select",
        section: "Training",
        options: getEusrCardTypeOptions(),
      },
      ...outcome,
    ];
  }
  if (kind === "streetworks") {
    return [
      ...people,
      {
        name: "swqrNumber",
        label: "SWQR number (from Workforce)",
        type: "text",
        readOnly: true,
        section: "Training",
      },
      {
        name: "course",
        label: "Course",
        type: "select",
        section: "Training",
        options: getStreetworksCourseOptions(),
      },
      {
        name: "streetworksCategory",
        label: "Streetworks category (units)",
        type: "multiselect",
        section: "Training",
        options: getStreetworksCategoryOptions(),
      },
      {
        name: "trainingDate",
        label: "Training date (from)",
        type: "date",
        section: "Training",
      },
      {
        name: "trainingDateEnd",
        label: "Training date (to)",
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
        name: "customerVisible",
        label: "Customer visible",
        type: "boolean",
        section: "Outcome",
      },
    ];
  }
  return [
    ...people,
    {
      name: "inHouseCertificationNumber",
      label: "In-House certification number (from Workforce)",
      type: "text",
      readOnly: true,
      section: "Training",
    },
    {
      name: "certificateCategory",
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
    {
      key: "workforceNumber",
      header: "Workforce number",
      render: (row) => row.workforceNumber ?? "—",
    },
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
    base.push({
      key: "novice",
      header: "Novice / EWT",
      render: (row) => row.noviceOrEwt ?? "—",
    });
  }
  if (kind === "eusr") {
    base.push({
      key: "number",
      header: "EUSR number",
      render: (row) => row.eusrNumber ?? "—",
    });
    base.push({
      key: "category",
      header: "Category",
      render: (row) => row.eusrCategory ?? "—",
    });
    base.push({
      key: "card",
      header: "Card type",
      render: (row) => row.cardType ?? "—",
    });
  }
  if (kind === "streetworks") {
    base.push({
      key: "number",
      header: "SWQR number",
      render: (row) => row.swqrNumber ?? "—",
    });
    base.push({
      key: "course",
      header: "Course",
      render: (row) => row.course ?? "—",
    });
    base.push({
      key: "category",
      header: "Category",
      render: (row) => row.streetworksCategory ?? "—",
    });
  }
  if (kind === "in-house") {
    base.push({
      key: "certCategory",
      header: "Course",
      render: (row) => row.certificateCategory ?? "—",
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
  nporsCategoryOptions,
}: {
  kind: RegisterKind;
  companies: Company[];
  workforce: AdminWorkforceOption[];
  initialRows: AdminTrainingRecord[];
  /** Live from SharePoint NPORS Categories list when kind=npors. */
  nporsCategoryOptions?: Array<{ value: string; label: string }>;
}) {
  return (
    <AdminCrudPage<AdminTrainingRecord>
      title={titles[kind]}
      description={
        kind === "npors"
          ? `${descriptions.npors} Categories load from SharePoint NPORS Categories (N number + full title).`
          : descriptions[kind]
      }
      columns={columnsFor(kind)}
      fields={fieldsFor(kind, nporsCategoryOptions)}
      companies={companies}
      workforce={workforce}
      initialRows={initialRows}
      enableCompanyFilter
      getCompanyName={(row) => row.companyName}
      drawerWide
      listUrl={`/api/admin/training-records/${kind}`}
      createUrl={`/api/admin/training-records/${kind}`}
      updateUrl={(id) => `/api/admin/training-records/${kind}/${id}`}
      deleteUrl={(id) => `/api/admin/training-records/${kind}/${id}`}
      mapResponse={(payload) =>
        ((payload as { records?: AdminTrainingRecord[] }).records ?? [])
      }
      searchKeys={[
        (row) => row.candidateName,
        (row) => row.companyName,
        (row) => row.workforceNumber,
        (row) => row.nporsNumber,
        (row) => row.eusrNumber,
        (row) => row.swqrNumber,
        (row) => row.nporsCategory,
        (row) => row.eusrCategory,
        (row) => row.course,
        (row) => row.certificateCategory,
      ]}
    />
  );
}
