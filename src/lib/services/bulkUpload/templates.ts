import type {
  BulkImportColumn,
  BulkImportTemplate,
  BulkImportType,
} from "@/types/bulkUpload";

const WORKFORCE_COLUMNS: BulkImportColumn[] = [
  { key: "workforceNumber", label: "Workforce Number" },
  { key: "candidateName", label: "Candidate Name", required: true },
  { key: "company", label: "Company Name", required: true },
  { key: "companyNumber", label: "Company Number" },
  { key: "trainingManager", label: "Training manager" },
  { key: "supervisor", label: "Supervisor" },
  { key: "department", label: "Department" },
  { key: "dateOfBirth", label: "Date of birth" },
  { key: "email", label: "Email" },
  { key: "status", label: "Status" },
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
  { key: "startDate", label: "Start Date" },
  { key: "stageOfNvq", label: "Stage of NVQ" },
  { key: "completedDate", label: "Completed Date" },
  { key: "notes", label: "Notes" },
];

const MATRIX_COLUMNS: BulkImportColumn[] = [
  { key: "candidateName", label: "Candidate Name", required: true },
  { key: "workforceNumber", label: "Workforce Number" },
  { key: "dateOfBirth", label: "Date of birth" },
  { key: "n001Expiry", label: "N001 Expiry" },
  { key: "n003Expiry", label: "N003 Expiry" },
  { key: "n004Expiry", label: "N004 Expiry" },
  { key: "n010Expiry", label: "N010 Expiry" },
  { key: "n020Expiry", label: "N020 Expiry" },
  { key: "n021Expiry", label: "N021 Expiry" },
  { key: "n027Expiry", label: "N027 Expiry" },
  { key: "n100Expiry", label: "N100 Expiry" },
  { key: "overallStatus", label: "Overall Status" },
  { key: "needsReview", label: "Needs Review" },
  { key: "matrixNotes", label: "Matrix Notes" },
];

export const BULK_IMPORT_TEMPLATES: BulkImportTemplate[] = [
  {
    importType: "workforce",
    label: "Workforce / Candidates",
    description:
      "Import candidate records. Duplicates are detected by workforce number or name + DOB + company.",
    fileName: "pave-workforce-template.csv",
    columns: WORKFORCE_COLUMNS,
    implemented: true,
  },
  {
    importType: "trainingMatrix",
    label: "Training Matrix rows",
    description:
      "Import matrix expiry/status rows. Candidates must already exist in Workforce (match by Workforce Number or name).",
    fileName: "pave-training-matrix-template.csv",
    columns: MATRIX_COLUMNS,
    implemented: true,
  },
  {
    importType: "npors",
    label: "NPORS records",
    description: "Import NPORS training records (coming next).",
    fileName: "pave-npors-template.csv",
    columns: NPORS_COLUMNS,
    implemented: false,
  },
  {
    importType: "eusr",
    label: "EUSR records",
    description: "Import EUSR training records (coming next).",
    fileName: "pave-eusr-template.csv",
    columns: EUSR_COLUMNS,
    implemented: false,
  },
  {
    importType: "streetworks",
    label: "Streetworks / NRSWA records",
    description: "Import Streetworks/NRSWA records (coming next).",
    fileName: "pave-streetworks-template.csv",
    columns: STREETWORKS_COLUMNS,
    implemented: false,
  },
  {
    importType: "inHouse",
    label: "In-House records",
    description: "Import in-house certificates (coming next).",
    fileName: "pave-in-house-template.csv",
    columns: IN_HOUSE_COLUMNS,
    implemented: false,
  },
  {
    importType: "nvq",
    label: "NVQ records",
    description: "Import NVQ progress records (coming next).",
    fileName: "pave-nvq-template.csv",
    columns: NVQ_COLUMNS,
    implemented: false,
  },
];

export function getBulkImportTemplate(
  importType: BulkImportType,
): BulkImportTemplate | null {
  return (
    BULK_IMPORT_TEMPLATES.find((row) => row.importType === importType) ?? null
  );
}

export function buildTemplateCsv(importType: BulkImportType): string {
  const template = getBulkImportTemplate(importType);
  if (!template) {
    throw new Error(`Unknown import type: ${importType}`);
  }
  const header = template.columns.map((c) => escapeCsv(c.label)).join(",");
  const example = template.columns
    .map((c) => {
      if (c.key === "candidateName") return escapeCsv("Jane Example");
      if (c.key === "company") return escapeCsv("Example Company Ltd");
      if (c.key === "department") return escapeCsv("Operations");
      if (c.key === "dateOfBirth" || c.key === "DOB" || c.key === "dob") {
        return escapeCsv("1990-01-15");
      }
      if (c.key === "workforceNumber") return escapeCsv("WF-1001");
      if (c.key === "status") return escapeCsv("Active");
      return "";
    })
    .join(",");
  return `${header}\r\n${example}\r\n`;
}

function escapeCsv(value: string): string {
  if (/[",\r\n]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}
