/**
 * Central SharePoint schema for the PAVE Training Portal.
 *
 * Site: https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin
 *
 * Rules:
 * - Backend code must import list names and internal field names from here.
 * - Do not hard-code SharePoint list/field names in components or ad-hoc services.
 * - `fields` values are SharePoint internal names (including encoded names).
 * - `labels` are UI-friendly display names only.
 */

export const SHAREPOINT_SITE = {
  url: "https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin",
  hostname: "pavetraining.sharepoint.com",
  serverRelativePath: "/sites/PaveTrainingOperationAdmin",
} as const;

export type SharePointFieldMap = Readonly<Record<string, string>>;

export interface SharePointListDefinition<TFields extends SharePointFieldMap> {
  /** Stable app key used in code. */
  key: string;
  /** Exact SharePoint list title. */
  listName: string;
  /** Customer/admin facing label when different from listName. */
  displayName: string;
  /** Environment variable that holds the Graph list ID. */
  listIdEnvVar: string;
  /** Internal field name map. */
  fields: TFields;
  /** Optional UI labels keyed by the same property keys as `fields`. */
  labels: Readonly<Record<keyof TFields & string, string>>;
}

const companyFields = {
  id: "ID",
  title: "Title",
  companyNumber: "CompanyNumber",
  companyName: "CompanyName",
  companySize: "CompanySize",
  registeredAddress: "RegisteredAddress",
  companyRegNumber: "CompanyRegNumber",
  vatNo: "VATNo",
  telNo: "TelNo",
  email: "Email",
  mainContact: "MainContact",
  accountsContactName: "AccountsContactName",
  /** SharePoint internal name is `Accountsaddress` (lowercase a). */
  accountsAddress: "Accountsaddress",
  /** SharePoint internal name is `AccountsContactnumber` (lowercase n). */
  accountsContactNumber: "AccountsContactnumber",
  /** SharePoint internal name is `Accountsemail` (lowercase e). */
  accountsEmail: "Accountsemail",
  /** SharePoint internal name is `Notespricesagreed` (lowercase p/a). */
  notesPricesAgreed: "Notespricesagreed",
  companyLogo: "CompanyLogo",
  status: "Status",
} as const;

const workforceFields = {
  id: "ID",
  candidateName: "CandidateName",
  companyName: "CompanyName",
  workforceNumber: "WorkforceNumber",
  dateOfBirth: "Dateofbirth",
  department: "Department",
  status: "Status",
  /** Live internal name is Trainingmanager (lowercase m). */
  trainingManager: "Trainingmanager",
  supervisor: "Supervisor",
  email: "Email",
  cscsNumber: "CSCSNumber",
  swqrNumber: "SWQRNumber",
  eusrNumber: "EUSRNumber",
  nporsNumbers: "NPORSNumbers",
  inHouseCertificationNumber: "InHouseCertificationNumber",
} as const;

const trainingMatrixFields = {
  id: "ID",
  candidateName: "CandidateName",
  matrixCompany: "MatrixCompany",
  companyName: "Company_x0020_Name",
  department: "Department",
  overallStatus: "OverallStatus",
  needsReview: "NeedsReview",
  matrixNotes: "MatrixNotes",
  nextExpiryDate: "NextExpiryDate",
  n001Expiry: "N001Expiry",
  n003Expiry: "N003Expiry",
  n004Expiry: "N004Expiry",
  n010Expiry: "N010Expiry",
  n020Expiry: "N020Expiry",
  n021Expiry: "N021Expiry",
  n027Expiry: "N027Expiry",
  n100Expiry: "N100Expiry",
} as const;

const nporsRegisterFields = {
  id: "ID",
  candidateName: "CandidateName",
  companyName: "CompanyName",
  nporsNumber: "NPORSNumber",
  trainingDate: "TrainingDate",
  trainingAddress: "TrainingAddress",
  noviceOrEwt: "NoviceorEwt",
  nporsCategory: "NPORSCategory",
  expiry: "Expiry",
  trainingOutcome: "TrainingOutcome",
  outcomeDate: "OutcomeDate",
  assessorTrainer: "AssessorTrainer",
  customerVisible: "CustomerVisible",
  outcomeNotes: "OutcomeNotes",
  notes: "Notes",
} as const;

const eusrRegisterFields = {
  id: "ID",
  candidateName: "CandidateName",
  companyName: "CompanyName",
  eusrNumber: "EUSRNumber",
  eusrCategory: "EusrCategory",
  cardType: "CardType",
  trainingDate: "TrainingDate",
  trainingAddress: "TrainingAddress",
  expiry: "Expiry",
  trainingOutcome: "TrainingOutcome",
  outcomeDate: "OutcomeDate",
  assessorTrainer: "AssessorTrainer",
  customerVisible: "CustomerVisible",
  outcomeNotes: "OutcomeNotes",
  notes: "Notes",
} as const;

const nrswaRegisterFields = {
  id: "ID",
  candidateName: "CandidateName",
  companyName: "CompanyName",
  swqrNumber: "SWQRNumber",
  streetworksCategory: "StreetworksCategory",
  course: "Course",
  trainingDate: "TrainingDate",
  trainingAddress: "TrainingAddress",
  expiryDate: "Expirydate",
  trainingOutcome: "TrainingOutcome",
  outcomeDate: "OutcomeDate",
  assessorTrainer: "AssessorTrainer",
  customerVisible: "CustomerVisible",
  outcomeNotes: "OutcomeNotes",
} as const;

const inHouseCertificatesFields = {
  id: "ID",
  candidateName: "CandidateName",
  companyName: "CompanyName",
  certificateCategory: "CertificateCategory",
  courseCategory: "CourseCategory",
  courseDate: "CourseDate",
  trainingAddress: "TrainingAddress",
  expiryDate: "ExpiryDate",
  trainingOutcome: "TrainingOutcome",
  outcomeDate: "OutcomeDate",
  assessorTrainer: "AssessorTrainer",
  customerVisible: "CustomerVisible",
  outcomeNotes: "OutcomeNotes",
  notes: "Notes",
} as const;

const nvqRegisterFields = {
  id: "ID",
  candidateName: "CandidateName",
  nvqCompany: "NVQCompany",
  companyName: "Company_x0020_Name",
  nvqTitle: "NvqTitle",
  boltonNvq: "BoltonNvq",
  dateRegistered: "DateRegistered",
  dateInductionBooked: "DateinductionBooked",
  stageOfNvq: "StageofNvq",
  customerUpdateNotes: "CustomerUpdateNotes",
  completedDate: "CompletedDate",
  customerVisible: "CustomerVisible",
  trainingOutcome: "TrainingOutcome",
  outcomeDate: "OutcomeDate",
  assessorTrainer: "AssessorTrainer",
  outcomeNotes: "OutcomeNotes",
} as const;

const customerDocumentsFields = {
  id: "ID",
  title: "Title",
  company: "Company",
  candidate: "Candidate",
  documentType: "DocumentType",
  customerVisible: "CustomerVisible",
  notificationSent: "NotificationSent",
  fileRef: "FileRef",
  fileLeafRef: "FileLeafRef",
  fsObjType: "FSObjType",
  modified: "Modified",
  editor: "Editor",
} as const;

const eventsFields = {
  id: "ID",
  title: "Title",
  /** Lookup to Company List — do not use a legacy `Company` field. */
  eventCompany: "EventCompany",
  /** Graph companion for EventCompany lookup writes. */
  eventCompanyLookupId: "EventCompanyLookupId",
  customerVisible: "Customer_x0020_Visible",
  trainingAddress: "TrainingAddress",
  eventDate: "EventDate",
  endDate: "EndDate",
  description: "Description",
  location: "Location",
  outlookEventId: "OutlookEventId",
  outlookCalendarId: "OutlookCalendarId",
  outlookICalUid: "OutlookICalUid",
  syncStatus: "SyncStatus",
  syncDirection: "SyncDirection",
  lastSyncedAt: "LastSyncedAt",
  lastSyncSource: "LastSyncSource",
  syncError: "SyncError",
  doNotSync: "DoNotSync",
} as const;

const offersPromotionsFields = {
  id: "ID",
  title: "Title",
  category: "Category",
  customerVisible: "CustomerVisible",
  startDate: "StartDate",
  endDate: "EndDate",
  /** Live SharePoint field is ShortDescription (not Description). */
  shortDescription: "ShortDescription",
  status: "Status",
} as const;

const permissionsFields = {
  id: "ID",
  userEmail: "UserEmail",
  roleType: "RoleType",
  status: "Status",
  company: "Company",
  /** Graph expand companion for lookup field `Company`. */
  companyLookupId: "CompanyLookupId",
  accessScope: "AccessScope",
  /** Person/display name — used for CandidateOnly and Supervisor matching. */
  name: "Name",
  /** Multi-choice department scopes for Supervisor. */
  departments: "Departments",
  /** Lookup multi to Departments list. */
  departmentsAllowed: "DepartmentsAllowed",
  canView: "CanView",
  canDownload: "CanDownload",
  canEdit: "CanEdit",
} as const;

const trainingCourseCategoriesFields = {
  id: "ID",
  title: "Title",
  categoryCode: "CategoryCode",
  courseName: "CourseName",
  courseType: "CourseType",
  source: "Source",
  active: "Active",
  customerVisible: "CustomerVisible",
  displayOrder: "DisplayOrder",
  notes: "Notes",
} as const;

const trainingManagerLogsFields = {
  id: "ID",
  title: "Title",
  /** Encoded internal name on live list. */
  userEmail: "User_x0020_Email",
  listName: "ListName",
  itemsId: "ItemsId",
  areaViewed: "Area_x0020_Viewed",
  timestamp: "Timestamp",
  notes: "Notes",
  company: "Company",
  role: "Role",
} as const;

export const SHAREPOINT_LISTS = {
  company: {
    key: "company",
    listName: "Company List",
    displayName: "Company",
    listIdEnvVar: "SHAREPOINT_COMPANY_LIST_ID",
    fields: companyFields,
    labels: {
      id: "ID",
      title: "Title",
      companyNumber: "Company number",
      companyName: "Company name",
      companySize: "Company size",
      registeredAddress: "Registered address",
      companyRegNumber: "Company reg number",
      vatNo: "VAT no",
      telNo: "Tel no",
      email: "Email",
      mainContact: "Main contact",
      accountsContactName: "Accounts contact name",
      accountsAddress: "Accounts address",
      accountsContactNumber: "Accounts contact number",
      accountsEmail: "Accounts email",
      notesPricesAgreed: "Notes / prices agreed",
      companyLogo: "Company logo",
      status: "Status",
    },
  },

  workforce: {
    key: "workforce",
    listName: "Workforce List",
    displayName: "Workforce",
    listIdEnvVar: "SHAREPOINT_WORKFORCE_LIST_ID",
    fields: workforceFields,
    labels: {
      id: "ID",
      candidateName: "Candidate name",
      companyName: "Company name",
      workforceNumber: "Workforce number",
      dateOfBirth: "Date of birth",
      department: "Department",
      status: "Status",
      trainingManager: "Training manager",
      supervisor: "Supervisor",
      email: "Email",
      cscsNumber: "CSCS number",
      swqrNumber: "SWQR number",
      eusrNumber: "EUSR number",
      nporsNumbers: "NPORS numbers",
      inHouseCertificationNumber: "In-house certification number",
    },
  },

  trainingMatrix: {
    key: "trainingMatrix",
    listName: "Training Matrix",
    displayName: "Training Matrix",
    listIdEnvVar: "SHAREPOINT_TRAINING_MATRIX_LIST_ID",
    fields: trainingMatrixFields,
    labels: {
      id: "ID",
      candidateName: "Candidate name",
      matrixCompany: "Matrix company",
      companyName: "Company name",
      department: "Department",
      overallStatus: "Overall status",
      needsReview: "Needs review",
      matrixNotes: "Matrix notes",
      nextExpiryDate: "Next expiry date",
      n001Expiry: "N001 expiry",
      n003Expiry: "N003 expiry",
      n004Expiry: "N004 expiry",
      n010Expiry: "N010 expiry",
      n020Expiry: "N020 expiry",
      n021Expiry: "N021 expiry",
      n027Expiry: "N027 expiry",
      n100Expiry: "N100 expiry",
    },
  },

  nporsRegister: {
    key: "nporsRegister",
    listName: "NPORS Register",
    displayName: "NPORS",
    listIdEnvVar: "SHAREPOINT_NPORS_REGISTER_LIST_ID",
    fields: nporsRegisterFields,
    labels: {
      id: "ID",
      candidateName: "Candidate name",
      companyName: "Company name",
      nporsNumber: "NPORS number",
      trainingDate: "Training date",
      trainingAddress: "Training address",
      noviceOrEwt: "Novice or EWT",
      nporsCategory: "NPORS category",
      expiry: "Expiry",
      trainingOutcome: "Training outcome",
      outcomeDate: "Outcome date",
      assessorTrainer: "Assessor / trainer",
      customerVisible: "Customer visible",
      outcomeNotes: "Outcome notes",
      notes: "Notes",
    },
  },

  eusrRegister: {
    key: "eusrRegister",
    listName: "EUSR Register",
    displayName: "EUSR",
    listIdEnvVar: "SHAREPOINT_EUSR_REGISTER_LIST_ID",
    fields: eusrRegisterFields,
    labels: {
      id: "ID",
      candidateName: "Candidate name",
      companyName: "Company name",
      eusrNumber: "EUSR number",
      eusrCategory: "EUSR category",
      cardType: "Card type",
      trainingDate: "Training date",
      trainingAddress: "Training address",
      expiry: "Expiry",
      trainingOutcome: "Training outcome",
      outcomeDate: "Outcome date",
      assessorTrainer: "Assessor / trainer",
      customerVisible: "Customer visible",
      outcomeNotes: "Outcome notes",
      notes: "Notes",
    },
  },

  nrswaRegister: {
    key: "nrswaRegister",
    listName: "NRSWA Register",
    displayName: "Streetworks Training",
    listIdEnvVar: "SHAREPOINT_NRSWA_REGISTER_LIST_ID",
    fields: nrswaRegisterFields,
    labels: {
      id: "ID",
      candidateName: "Candidate name",
      companyName: "Company name",
      swqrNumber: "SWQR number",
      streetworksCategory: "Streetworks category",
      course: "Course",
      trainingDate: "Training date",
      trainingAddress: "Training address",
      expiryDate: "Expiry date",
      trainingOutcome: "Training outcome",
      outcomeDate: "Outcome date",
      assessorTrainer: "Assessor / trainer",
      customerVisible: "Customer visible",
      outcomeNotes: "Outcome notes",
    },
  },

  inHouseCertificates: {
    key: "inHouseCertificates",
    listName: "In-House Certificates Register",
    displayName: "In-House Certificates",
    listIdEnvVar: "SHAREPOINT_IN_HOUSE_CERTIFICATES_LIST_ID",
    fields: inHouseCertificatesFields,
    labels: {
      id: "ID",
      candidateName: "Candidate name",
      companyName: "Company name",
      certificateCategory: "Certificate category",
      courseCategory: "Course category",
      courseDate: "Course date",
      trainingAddress: "Training address",
      expiryDate: "Expiry date",
      trainingOutcome: "Training outcome",
      outcomeDate: "Outcome date",
      assessorTrainer: "Assessor / trainer",
      customerVisible: "Customer visible",
      outcomeNotes: "Outcome notes",
      notes: "Notes",
    },
  },

  nvqRegister: {
    key: "nvqRegister",
    listName: "NVQ Register",
    displayName: "NVQ",
    listIdEnvVar: "SHAREPOINT_NVQ_REGISTER_LIST_ID",
    fields: nvqRegisterFields,
    labels: {
      id: "ID",
      candidateName: "Candidate name",
      nvqCompany: "NVQ company",
      companyName: "Company name",
      nvqTitle: "NVQ title",
      boltonNvq: "Bolton NVQ",
      dateRegistered: "Date registered",
      dateInductionBooked: "Date induction booked",
      stageOfNvq: "Stage of NVQ",
      customerUpdateNotes: "Customer update notes",
      completedDate: "Completed date",
      customerVisible: "Customer visible",
      trainingOutcome: "Training outcome",
      outcomeDate: "Outcome date",
      assessorTrainer: "Assessor / trainer",
      outcomeNotes: "Outcome notes",
    },
  },

  customerDocuments: {
    key: "customerDocuments",
    listName: "Customer Documents",
    displayName: "Customer Documents",
    listIdEnvVar: "SHAREPOINT_CUSTOMER_DOCUMENTS_LIST_ID",
    fields: customerDocumentsFields,
    labels: {
      id: "ID",
      title: "Title",
      company: "Company",
      candidate: "Candidate",
      documentType: "Document type",
      customerVisible: "Customer visible",
      notificationSent: "Notification sent",
      fileRef: "File path",
      fileLeafRef: "File name",
      fsObjType: "Object type",
      modified: "Modified date",
      editor: "Modified by",
    },
  },

  events: {
    key: "events",
    listName: "Events",
    displayName: "Events",
    listIdEnvVar: "SHAREPOINT_EVENTS_LIST_ID",
    fields: eventsFields,
    labels: {
      id: "ID",
      title: "Title",
      eventCompany: "Company",
      eventCompanyLookupId: "Company lookup ID",
      customerVisible: "Customer visible",
      trainingAddress: "Training address",
      eventDate: "Start date / time",
      endDate: "End date / time",
      description: "Description",
      location: "Location",
      outlookEventId: "Outlook event ID",
      outlookCalendarId: "Outlook calendar ID",
      outlookICalUid: "Outlook iCal UID",
      syncStatus: "Sync status",
      syncDirection: "Sync direction",
      lastSyncedAt: "Last synced at",
      lastSyncSource: "Last sync source",
      syncError: "Sync error",
      doNotSync: "Do not sync",
    },
  },

  offersPromotions: {
    key: "offersPromotions",
    listName: "Offers / Promotions",
    displayName: "Offers / Promotions",
    /** GUID avoids REST 404 — slash in title breaks getbytitle URLs. */
    listId: "8e887fc7-0404-47e3-977c-e6e24e0b85c6",
    listIdEnvVar: "SHAREPOINT_OFFERS_PROMOTIONS_LIST_ID",
    fields: offersPromotionsFields,
    labels: {
      id: "ID",
      title: "Title",
      category: "Category",
      customerVisible: "Customer visible",
      startDate: "Start date",
      endDate: "End date",
      shortDescription: "Short description",
      status: "Status",
    },
  },

  permissions: {
    key: "permissions",
    listName: "Permissions List",
    displayName: "Permissions",
    listIdEnvVar: "SHAREPOINT_PERMISSIONS_LIST_ID",
    fields: permissionsFields,
    labels: {
      id: "ID",
      userEmail: "User email",
      roleType: "Role type",
      status: "Status",
      company: "Company",
      companyLookupId: "Company lookup ID",
      accessScope: "Access scope",
      name: "Name",
      departments: "Departments",
      departmentsAllowed: "Departments allowed",
      canView: "Can view",
      canDownload: "Can download",
      canEdit: "Can edit",
    },
  },

  trainingCourseCategories: {
    key: "trainingCourseCategories",
    listName: "Training Course Categories",
    displayName: "Training Course Categories",
    listIdEnvVar: "SHAREPOINT_TRAINING_COURSE_CATEGORIES_LIST_ID",
    fields: trainingCourseCategoriesFields,
    labels: {
      id: "ID",
      title: "Title",
      categoryCode: "Category code",
      courseName: "Course name",
      courseType: "Course type",
      source: "Source",
      active: "Active",
      customerVisible: "Customer visible",
      displayOrder: "Display order",
      notes: "Notes",
    },
  },

  trainingManagerLogs: {
    key: "trainingManagerLogs",
    listName: "Training Manager Logs",
    displayName: "Training Manager Logs",
    listIdEnvVar: "SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID",
    fields: trainingManagerLogsFields,
    labels: {
      id: "ID",
      title: "Title",
      userEmail: "User email",
      listName: "List name",
      itemsId: "Items ID",
      areaViewed: "Area viewed",
      timestamp: "Timestamp",
      notes: "Notes",
      company: "Company",
      role: "Role",
    },
  },
} as const;

export type SharePointListKey = keyof typeof SHAREPOINT_LISTS;

export type SharePointListFields<K extends SharePointListKey> =
  (typeof SHAREPOINT_LISTS)[K]["fields"];

export function getSharePointList<K extends SharePointListKey>(listKey: K) {
  return SHAREPOINT_LISTS[listKey];
}

export function getSharePointListName(listKey: SharePointListKey): string {
  return SHAREPOINT_LISTS[listKey].listName;
}

export function getSharePointDisplayName(listKey: SharePointListKey): string {
  return SHAREPOINT_LISTS[listKey].displayName;
}

export function getSharePointFields<K extends SharePointListKey>(
  listKey: K,
): SharePointListFields<K> {
  return SHAREPOINT_LISTS[listKey].fields;
}

export function getSharePointFieldInternalNames(
  listKey: SharePointListKey,
): string[] {
  const fields = SHAREPOINT_LISTS[listKey].fields as Record<string, string>;
  const names: string[] = [];
  for (const key in fields) {
    if (Object.prototype.hasOwnProperty.call(fields, key)) {
      names.push(fields[key]);
    }
  }
  return names;
}

export function getSharePointFieldLabel(
  listKey: SharePointListKey,
  fieldKey: string,
): string {
  const labels = SHAREPOINT_LISTS[listKey].labels as Record<string, string>;
  return labels[fieldKey] ?? fieldKey;
}
