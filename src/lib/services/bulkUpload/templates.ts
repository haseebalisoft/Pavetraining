import type {
  BulkImportColumn,
  BulkImportTemplate,
  BulkImportType,
} from "@/types/bulkUpload";
import {
  CLIENT_COMPANY_HEADERS,
  CLIENT_MATRIX_CATEGORY_COLUMNS,
  CLIENT_MATRIX_META_HEADERS,
  CLIENT_WORKFORCE_HEADERS,
} from "@/lib/services/bulkUpload/clientTemplateHeaders";

const COMPANY_COLUMNS: BulkImportColumn[] = CLIENT_COMPANY_HEADERS.map(
  (label) => ({
    key: label,
    label,
    required: label === "Company Number" || label === "Company Name",
  }),
);

const WORKFORCE_COLUMNS: BulkImportColumn[] = CLIENT_WORKFORCE_HEADERS.map(
  (label) => ({
    key: label,
    label,
    required: label === "Candidate Name" || label === "Company Name",
  }),
);

const MATRIX_COLUMNS: BulkImportColumn[] = [
  ...CLIENT_MATRIX_META_HEADERS.map((label) => ({
    key: label,
    label,
    required: label === "Name",
  })),
  ...CLIENT_MATRIX_CATEGORY_COLUMNS.map((column) => ({
    key: column.code,
    label: column.header,
  })),
];

const NPORS_COLUMNS: BulkImportColumn[] = [
  { key: "candidateName", label: "Candidate Name", required: true },
  { key: "company", label: "Company", required: true },
  { key: "nporsNumber", label: "NPORS Number" },
  { key: "nporsCategory", label: "NPORS Category" },
  { key: "noviceOrEwt", label: "Novice or EWT" },
  { key: "startDate", label: "Start Date" },
  { key: "expiry", label: "Expiry" },
  { key: "outcome", label: "Outcome" },
];

const EUSR_COLUMNS: BulkImportColumn[] = [
  { key: "candidateName", label: "Candidate Name", required: true },
  { key: "company", label: "Company", required: true },
  { key: "eusrNumber", label: "EUSR Number" },
  { key: "eusrCategory", label: "EUSR Category" },
  { key: "startDate", label: "Start Date" },
  { key: "expiry", label: "Expiry" },
  { key: "outcome", label: "Outcome" },
];

const STREETWORKS_COLUMNS: BulkImportColumn[] = [
  { key: "candidateName", label: "Candidate Name", required: true },
  { key: "company", label: "Company", required: true },
  { key: "swqrNumber", label: "SWQR Number" },
  { key: "course", label: "Course" },
  { key: "streetworksCategory", label: "Streetworks Category" },
  { key: "startDate", label: "Start Date" },
  { key: "expiry", label: "Expiry" },
  { key: "outcome", label: "Outcome" },
];

const IN_HOUSE_COLUMNS: BulkImportColumn[] = [
  { key: "candidateName", label: "Candidate Name", required: true },
  { key: "company", label: "Company", required: true },
  { key: "course", label: "Course" },
  { key: "certificateCategory", label: "Certificate Category" },
  { key: "startDate", label: "Start Date" },
  { key: "expiry", label: "Expiry" },
  { key: "outcome", label: "Outcome" },
];

const NVQ_COLUMNS: BulkImportColumn[] = [
  { key: "candidateName", label: "Candidate Name", required: true },
  { key: "company", label: "Company", required: true },
  { key: "nvqTitle", label: "NVQ Title" },
  { key: "boltOn", label: "Bolt On" },
  { key: "dateRegistered", label: "Date Registered" },
  { key: "inductionDate", label: "Date Induction Booked" },
  { key: "stageOfNvq", label: "Stage of NVQ" },
  { key: "notes", label: "Notes" },
  { key: "completedDate", label: "Completed Date" },
];

export const BULK_IMPORT_TEMPLATES: BulkImportTemplate[] = [
  {
    importType: "company",
    label: "Companies",
    description:
      "Use the exact client Company list.xlsx headers. Creates or updates SharePoint Company List rows.",
    fileName: "Company-list-template.xlsx",
    columns: COMPANY_COLUMNS,
    implemented: true,
  },
  {
    importType: "workforce",
    label: "Workforce / Candidates",
    description:
      "Use the exact client Workforce list.xlsx headers. Missing companies are created automatically on import.",
    fileName: "Workforce-list-template.xlsx",
    columns: WORKFORCE_COLUMNS,
    implemented: true,
  },
  {
    importType: "trainingMatrix",
    label: "Training Matrix",
    description:
      "Use the exact client Training matrix example.xlsx headers. N-code dates write to Training Matrix + Category Records. Import Workforce first.",
    fileName: "Training-matrix-template.xlsx",
    columns: MATRIX_COLUMNS,
    implemented: true,
  },
  {
    importType: "npors",
    label: "NPORS records",
    description:
      "Import NPORS training records. Candidate must already exist in Workforce. Pass outcomes sync into the Training Matrix.",
    fileName: "pave-npors-template.csv",
    columns: NPORS_COLUMNS,
    implemented: true,
  },
  {
    importType: "eusr",
    label: "EUSR records",
    description:
      "Import EUSR training records. Candidate must already exist in Workforce. Pass outcomes sync into the Training Matrix.",
    fileName: "pave-eusr-template.csv",
    columns: EUSR_COLUMNS,
    implemented: true,
  },
  {
    importType: "streetworks",
    label: "Streetworks / NRSWA records",
    description:
      "Import Streetworks/NRSWA records. Candidate must already exist in Workforce. Pass outcomes sync into the Training Matrix.",
    fileName: "pave-streetworks-template.csv",
    columns: STREETWORKS_COLUMNS,
    implemented: true,
  },
  {
    importType: "inHouse",
    label: "In-House records",
    description:
      "Import in-house certificates (standalone — does not update the Training Matrix).",
    fileName: "pave-in-house-template.csv",
    columns: IN_HOUSE_COLUMNS,
    implemented: true,
  },
  {
    importType: "nvq",
    label: "NVQ records",
    description:
      "Import NVQ progress records (standalone — does not update the Training Matrix).",
    fileName: "pave-nvq-template.csv",
    columns: NVQ_COLUMNS,
    implemented: true,
  },
];

export function getBulkImportTemplate(
  importType: BulkImportType,
): BulkImportTemplate | null {
  return (
    BULK_IMPORT_TEMPLATES.find((row) => row.importType === importType) ?? null
  );
}

export function getClientExcelTemplatePath(
  importType: BulkImportType,
): string | null {
  if (importType === "company") {
    return "bulk-templates/Company-list-template.xlsx";
  }
  if (importType === "workforce") {
    return "bulk-templates/Workforce-list-template.xlsx";
  }
  if (importType === "trainingMatrix") {
    return "bulk-templates/Training-matrix-template.xlsx";
  }
  return null;
}

export function buildTemplateCsv(importType: BulkImportType): string {
  const template = getBulkImportTemplate(importType);
  if (!template) {
    throw new Error(`Unknown import type: ${importType}`);
  }
  const header = template.columns.map((c) => escapeCsv(c.label)).join(",");
  return `${header}\r\n`;
}

function escapeCsv(value: string): string {
  if (/[",\r\n]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}
