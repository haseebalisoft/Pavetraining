/**
 * SharePoint list titles and internal field names for Pave Training Operation Admin.
 * Aligned with the live site schema (see spfx/src/shared/schema/sharepointSchema.ts).
 */
export const LIST_TITLES = {
  company: "Company List",
  workforce: "Workforce List",
  npors: "NPORS Register",
  nvq: "NVQ Register",
  nrswa: "NRSWA Register",
  /** Live title on site (not "In House Register"). */
  inHouse: "In-House Certificates Register",
  eusr: "EUSR Register",
  events: "Events",
  /** Prefer site document library; Customer Documents used as fallback. */
  documents: "Documents",
  customerDocuments: "Customer Documents",
} as const;

export const FIELDS = {
  company: {
    id: "ID",
    number: "CompanyNumber",
    name: "CompanyName",
    size: "CompanySize",
    address: "RegisteredAddress",
    regNo: "CompanyRegNumber",
    vatNo: "VATNo",
    tel: "TelNo",
    email: "Email",
    mainContact: "MainContact",
  },
  workforce: {
    id: "ID",
    name: "CandidateName",
    role: "Department",
    phone: "Trainingmanager",
    email: "Email",
    status: "Status",
  },
  npors: {
    id: "ID",
    operator: "CandidateName",
    category: "NPORSCategory",
    expiryDate: "Expiry",
    status: "TrainingOutcome",
  },
  eusr: {
    id: "ID",
    operator: "CandidateName",
    category: "EusrCategory",
    expiryDate: "Expiry",
    status: "TrainingOutcome",
  },
  nrswa: {
    id: "ID",
    operator: "CandidateName",
    category: "Course",
    expiryDate: "Expirydate",
    status: "TrainingOutcome",
  },
  inHouse: {
    id: "ID",
    operator: "CandidateName",
    category: "CourseCategory",
    expiryDate: "ExpiryDate",
    status: "TrainingOutcome",
  },
  nvq: {
    id: "ID",
    operator: "CandidateName",
    category: "NvqTitle",
    expiryDate: "CompletedDate",
    status: "StageofNvq",
    dateRegistered: "DateRegistered",
    completedDate: "CompletedDate",
  },
  events: {
    id: "ID",
    title: "Title",
    start: "EventDate",
    end: "EndDate",
    location: "Location",
  },
  documents: {
    id: "ID",
    name: "FileLeafRef",
    url: "FileRef",
    modified: "Modified",
  },
} as const;
