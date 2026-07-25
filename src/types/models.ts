/**
 * Strongly typed domain models for the PAVE Training Portal.
 * Field values are mapped from SharePoint using `src/lib/schema/sharepointSchema.ts`.
 */

export type RoleType = "Admin" | "Customer";
export type AccessScope = string;
export type PermissionStatus = "Active" | "Inactive" | string;

/** SharePoint person/lookup style value when returned as an object. */
export interface SharePointLookupValue {
  lookupId: string;
  lookupValue: string;
}

export interface Company {
  id: string;
  title: string;
  companyName: string;
  companyNumber: string | null;
  companySize: string | null;
  registeredAddress: string | null;
  companyRegNumber: string | null;
  vatNo: string | null;
  telNo: string | null;
  email: string | null;
  mainContact: string | null;
  accountsContactName: string | null;
  accountsAddress: string | null;
  accountsContactNumber: string | null;
  accountsEmail: string | null;
  notesPricesAgreed: string | null;
  companyLogo: string | null;
  status: string;
}

/** Customer-visible company profile (excludes internal commercial notes). */
export interface CustomerCompanyProfile {
  id: string;
  companyName: string;
  companyNumber: string | null;
  companySize: string | null;
  registeredAddress: string | null;
  companyRegNumber: string | null;
  vatNo: string | null;
  telNo: string | null;
  email: string | null;
  mainContact: string | null;
  companyLogo: string | null;
  status: string;
}

export interface WorkforceCandidate {
  id: string;
  candidateName: string;
  companyName: string;
  workforceNumber: string | null;
  dateOfBirth: string | null;
  department: string | null;
  status: string | null;
  trainingManager: string | null;
  supervisor: string | null;
  cscsNumber: string | null;
  swqrNumber: string | null;
  eusrNumber: string | null;
  nporsNumbers: string | null;
  inHouseCertificationNumber: string | null;
}

export interface TrainingMatrixRow {
  id: string;
  candidateName: string;
  matrixCompany: string | null;
  companyName: string | null;
  department: string | null;
  overallStatus: string | null;
  needsReview: boolean;
  matrixNotes: string | null;
  nextExpiryDate: string | null;
  n001Expiry: string | null;
  n003Expiry: string | null;
  n004Expiry: string | null;
  n010Expiry: string | null;
  n020Expiry: string | null;
  n021Expiry: string | null;
  n027Expiry: string | null;
  n100Expiry: string | null;
}

export interface TrainingRecordBase {
  id: string;
  candidateName: string;
  companyName: string;
  trainingDate: string | null;
  trainingAddress: string | null;
  trainingOutcome: string | null;
  outcomeDate: string | null;
  assessorTrainer: string | null;
  customerVisible: boolean;
  outcomeNotes: string | null;
}

export interface NporsTrainingRecord extends TrainingRecordBase {
  nporsNumber: string | null;
  noviceOrEwt: string | null;
  nporsCategory: string | null;
  expiry: string | null;
  notes: string | null;
}

export interface EusrTrainingRecord extends TrainingRecordBase {
  eusrNumber: string | null;
  eusrCategory: string | null;
  cardType: string | null;
  expiry: string | null;
  notes: string | null;
}

export interface StreetworksTrainingRecord extends TrainingRecordBase {
  swqrNumber: string | null;
  streetworksCategory: string | null;
  course: string | null;
  expiryDate: string | null;
}

export interface InHouseTrainingRecord extends TrainingRecordBase {
  certificateCategory: string | null;
  courseCategory: string | null;
  courseDate: string | null;
  expiryDate: string | null;
  notes: string | null;
}

export interface NvqRecord {
  id: string;
  candidateName: string;
  nvqCompany: string | null;
  companyName: string | null;
  nvqTitle: string | null;
  boltonNvq: string | null;
  dateRegistered: string | null;
  dateInductionBooked: string | null;
  stageOfNvq: string | null;
  customerUpdateNotes: string | null;
  completedDate: string | null;
  customerVisible: boolean;
  trainingOutcome: string | null;
  outcomeDate: string | null;
  assessorTrainer: string | null;
  outcomeNotes: string | null;
}

export interface CustomerDocument {
  id: string;
  title: string;
  company: string | null;
  candidate: string | null;
  documentType: string | null;
  customerVisible: boolean;
  notificationSent: boolean;
  fileRef: string | null;
  fileLeafRef: string | null;
  fsObjType: string | number | null;
}

export interface TrainingEvent {
  id: string;
  title: string;
  eventCompany: string | null;
  customerVisible: boolean;
  trainingAddress: string | null;
  eventDate: string | null;
  endDate: string | null;
  description: string | null;
  location: string | null;
  outlookEventId: string | null;
  syncStatus: string | null;
  lastSyncedAt: string | null;
}

export interface OfferPromotion {
  id: string;
  title: string;
  company: string | null;
  customerVisible: boolean;
  startDate: string | null;
  endDate: string | null;
  description: string | null;
  status: string | null;
}

export interface PermissionProfile {
  id: string;
  userEmail: string;
  roleType: RoleType;
  status: PermissionStatus;
  companyId: string;
  companyDisplayName?: string;
  accessScope: AccessScope;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
}

export interface TrainingCourseCategory {
  id: string;
  title: string;
  categoryCode: string | null;
  courseName: string | null;
  courseType: string | null;
  source: string | null;
  active: boolean;
  customerVisible: boolean;
  displayOrder: number | null;
  notes: string | null;
}

export interface MeResponse {
  loggedInEmail: string;
  role: RoleType;
  redirectTo: "/admin" | "/customer";
  companyId: string | null;
  companyName: string | null;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
  accessScope: AccessScope | null;
}

export interface CustomerContext {
  loggedInEmail: string;
  role: "Customer";
  companyId: string;
  companyName: string;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
  accessScope: AccessScope;
  permissionStatus: "Active";
}

export interface AdminContext {
  loggedInEmail: string;
  role: "Admin";
  /** Reserved for future admin company selection; never sourced from customer query params. */
  selectedCompanyId: string | null;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
  accessScope: AccessScope;
}

export interface AdminDashboardStats {
  totalCompanies: number;
  activeCompanies: number;
  totalCandidates: number;
  expiredTraining: number;
  expiringWithin3Months: number;
  recordsToReview: number;
  activeNvqs: number;
  completedNvqs: number;
  documentsPendingVisibility: number;
  upcomingEvents: number;
}

export type AdminWarningIssue =
  | "CompanyName"
  | "CustomerVisible"
  | "TrainingAddress";

export interface AdminDataWarning {
  id: string;
  source: string;
  candidateName: string | null;
  issues: AdminWarningIssue[];
  detail: string;
}

export interface AdminDashboardPayload {
  selectedCompanyId: string | null;
  selectedCompanyName: string | null;
  stats: AdminDashboardStats;
  warnings: AdminDataWarning[];
}

export interface DashboardStats {
  workforceCount: number;
  trainingMatrixCount: number;
  needsReviewCount: number;
  expiringSoonCount: number;
  expiredCount: number;
  documentsCount: number;
  upcomingEventsCount: number;
  activeOffersCount: number;
  nporsCount: number;
  eusrCount: number;
  streetworksCount: number;
  inHouseCount: number;
  nvqCount: number;
}

/** Customer-visible training matrix summary (no admin notes). */
export interface CustomerMatrixRecord {
  id: string;
  candidateName: string;
  department: string | null;
  overallStatus: string | null;
  needsReview: boolean;
  nextExpiryDate: string | null;
  n001Expiry: string | null;
  n003Expiry: string | null;
  n004Expiry: string | null;
  n010Expiry: string | null;
  n020Expiry: string | null;
  n021Expiry: string | null;
  n027Expiry: string | null;
  n100Expiry: string | null;
}

/** Customer-facing Pass/Fail only — other SharePoint outcome values are hidden. */
export type CustomerOutcome = "Pass" | "Fail";

export interface CustomerTrainingRecordBase {
  id: string;
  candidateName: string;
  workforceId: string | null;
  trainingDate: string | null;
  trainingAddress: string | null;
  outcome: CustomerOutcome | null;
  expiry: string | null;
}

export interface CustomerNporsRecord extends CustomerTrainingRecordBase {
  nporsNumber: string | null;
  noviceOrEwt: string | null;
  nporsCategory: string | null;
}

export interface CustomerEusrRecord extends CustomerTrainingRecordBase {
  eusrNumber: string | null;
  eusrCategory: string | null;
}

export interface CustomerStreetworksRecord extends CustomerTrainingRecordBase {
  swqrNumber: string | null;
  course: string | null;
  streetworksCategory: string | null;
}

export interface CustomerInHouseRecord extends CustomerTrainingRecordBase {
  course: string | null;
}

export type CustomerNvqStatus = "Active" | "Completed";

export interface CustomerNvqRecord {
  id: string;
  candidateName: string;
  nvqTitle: string | null;
  boltOn: string | null;
  dateRegistered: string | null;
  inductionDate: string | null;
  stageOfNvq: string | null;
  notes: string | null;
  completedDate: string | null;
  status: CustomerNvqStatus;
}

export interface CustomerDocumentRecord {
  id: string;
  name: string;
  documentType: string | null;
  candidate: string | null;
  /** Modified / uploaded date for display. */
  uploadedDate: string | null;
  canDownload: boolean;
  /** Inline preview path — always set for customer-visible files. */
  viewPath: string | null;
  /** Download path only when CanDownload is true. */
  downloadPath: string | null;
}

export interface CustomerEventRecord {
  id: string;
  title: string;
  eventDate: string | null;
  endDate: string | null;
  trainingAddress: string | null;
  location: string | null;
  description: string | null;
  company: string | null;
}

export interface CustomerOfferRecord {
  id: string;
  title: string;
  description: string | null;
  startDate: string | null;
  endDate: string | null;
  status: string | null;
}
