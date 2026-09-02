/**
 * Strongly typed domain models for the PAVE Training Portal.
 * Field values are mapped from SharePoint using `src/lib/schema/sharepointSchema.ts`.
 */

export type RoleType = "Admin" | "Customer";
export type AccessScope = string;
export type PermissionStatus = "Active" | "Inactive" | string;

/** Customer-side role from Permissions List (Wayne brief). */
export type CustomerRoleType =
  | "TrainingManager"
  | "Supervisor"
  | "Candidate";

/** Normalized AccessScope for server-side filtering. */
export type NormalizedAccessScope =
  | "All"
  | "Company"
  | "Department"
  | "AssignedCandidates"
  | "CandidateOnly";

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
  /** SharePoint Company list id when the lookup is set. */
  companyId: string | null;
  /** App media URL for the company logo when set. */
  companyLogo: string | null;
  workforceNumber: string | null;
  dateOfBirth: string | null;
  department: string | null;
  status: string | null;
  trainingManager: string | null;
  supervisor: string | null;
  /** Workforce Email — used for CandidateOnly matching. */
  email: string | null;
  cscsNumber: string | null;
  swqrNumber: string | null;
  eusrNumber: string | null;
  nporsNumbers: string | null;
  inHouseCertificationNumber: string | null;
  cscsExpiry: string | null;
  swqrExpiry: string | null;
  eusrExpiry: string | null;
  /** App media URL for Candidate Photo when set. */
  photoUrl: string | null;
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
  niNumber: string | null;
  ulnNumber: string | null;
  nvqTitle: string | null;
  boltonNvq: string | null;
  poNumber: string | null;
  cardSchemeCategory: string | null;
  cardExtensionDateNeeded: string | null;
  siteAddress: string | null;
  siteContact: string | null;
  englishUnderstandingConfirmed: boolean;
  tcAcknowledged: boolean;
  gdprConsent: boolean;
  dateRegistered: string | null;
  dateInductionBooked: string | null;
  stageOfNvq: string | null;
  notes: string | null;
  completedDate: string | null;
  certificationDate: string | null;
  customerUpdateNotes: string | null;
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
  eventCompanyId: string | null;
  customerVisible: boolean;
  trainingAddress: string | null;
  eventDate: string | null;
  endDate: string | null;
  /** Customer-facing copy; this is the only description exposed publicly. */
  description: string | null;
  /** Admin-only notes. Never include in customer API payloads. */
  internalNotes: string | null;
  location: string | null;
  outlookEventId: string | null;
  outlookCalendarId: string | null;
  outlookICalUid: string | null;
  syncStatus: string | null;
  syncDirection: string | null;
  lastSyncedAt: string | null;
  lastSyncSource: string | null;
  syncError: string | null;
  doNotSync: boolean;
}

export interface OfferPromotion {
  id: string;
  title: string;
  category: string | null;
  customerVisible: boolean;
  startDate: string | null;
  endDate: string | null;
  description: string | null;
  status: string | null;
  image: string | null;
  ctaLabel: string | null;
  ctaLink: string | null;
}

export interface PermissionProfile {
  id: string;
  userEmail: string;
  /** Admin | Customer — portal routing bucket (legacy compatible). */
  roleType: RoleType;
  /** Raw SharePoint RoleType choice text. */
  sharePointRoleType: string;
  /** Customer-side role when user may use the customer portal. */
  customerRole: CustomerRoleType | null;
  /** UI label e.g. "Training Manager". */
  roleLabel: string;
  status: PermissionStatus;
  companyId: string;
  companyDisplayName?: string;
  /** Raw SharePoint AccessScope choice. */
  accessScope: AccessScope;
  /** Normalized scope used by customer filters. */
  normalizedAccessScope: NormalizedAccessScope;
  /** From Permissions.Departments / DepartmentsAllowed. */
  departmentScopes: string[];
  /** CandidateOnly — Permissions.Name or matched workforce name. */
  candidateScopeName: string | null;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
  /** PAVE Admin portal (includes legacy Training Manager). */
  canAccessAdmin: boolean;
  /** Customer portal (TM / Supervisor / Candidate). */
  canAccessCustomer: boolean;
  /** Defaults true when SharePoint field missing. */
  receiveDocumentNotifications: boolean;
  receiveExpiryNotifications: boolean;
  customerNotificationsEnabled: boolean;
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
  redirectTo: "/admin" | "/customer/dashboard";
  companyId: string | null;
  companyName: string | null;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
  accessScope: AccessScope | null;
  customerRole?: CustomerRoleType | null;
  roleLabel?: string;
  normalizedAccessScope?: NormalizedAccessScope;
}

export interface CustomerContext {
  loggedInEmail: string;
  role: "Customer";
  customerRole: CustomerRoleType;
  roleLabel: string;
  companyId: string;
  companyName: string;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
  accessScope: AccessScope;
  normalizedAccessScope: NormalizedAccessScope;
  departmentScopes: string[];
  candidateScopeName: string | null;
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
  /**
   * Raw SharePoint RoleType value ("Admin", "Training Manager", …). Preserved
   * so the top-nav can distinguish a pure Admin from a Training Manager who
   * also has `canAccessAdmin`.
   */
  sharePointRoleType: string;
  /** Customer sub-role for Training Managers who ALSO have admin access; null for pure Admins. */
  customerRole: CustomerRoleType | null;
  /** Human label ("Admin" / "Training Manager") for badges + diagnostics. */
  roleLabel: string;
  /**
   * True iff this user should have full SharePoint-Admin power. Either:
   *   - literal SharePoint RoleType === "Admin" (customerRole === null), OR
   *   - the email is on the hardcoded/protected admin list.
   * Training Managers with only `canAccessAdmin === true` are NOT SharePoint admins.
   * Used to gate admin-only nav items and pages (Bulk Upload, Permissions writes).
   */
  isSharePointAdmin: boolean;
  /**
   * True iff `loggedInEmail` is on the hardcoded protected-admin list (see
   * `protectedAdmins.ts`). Rendered in the admin debug panel so we can quickly
   * tell whether a user is being kept as admin by that safety net.
   */
  isAlwaysAdminEmail: boolean;
}

export interface AdminDashboardExpiryRow {
  id: string;
  candidateName: string;
  companyName: string | null;
  nextExpiryDate: string | null;
  statusLabel: string;
  statusTone: "danger" | "warn" | "ok" | "missing";
}

export interface AdminDashboardDocumentRow {
  id: string;
  name: string;
  company: string | null;
  candidate: string | null;
  modifiedDate: string | null;
  customerVisible: boolean;
}

export interface AdminDashboardEventRow {
  id: string;
  title: string;
  company: string | null;
  eventDate: string | null;
  location: string | null;
}

export interface AdminDashboardActivityRow {
  id: string;
  title: string;
  userEmail: string | null;
  timestamp: string | null;
  detail: string | null;
}

export interface AdminDashboardStats {
  totalCompanies: number;
  activeCompanies: number;
  totalCandidates: number;
  activeCandidates: number;
  expiredTraining: number;
  expiringWithin3Months: number;
  expiringWithin6Months: number;
  recordsToReview: number;
  activeNvqs: number;
  completedNvqs: number;
  documentsPendingVisibility: number;
  documentsUploadedRecently: number;
  upcomingEvents: number;
  accessInvitationsPending: number;
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
  upcomingExpiries: AdminDashboardExpiryRow[];
  recentDocuments: AdminDashboardDocumentRow[];
  upcomingBookings: AdminDashboardEventRow[];
  recentActivity: AdminDashboardActivityRow[];
}

export interface DashboardStats {
  workforceCount: number;
  trainingMatrixCount: number;
  needsReviewCount: number;
  expiringSoonCount: number;
  /** Matrix rows with next expiry in 91–270 days (Upcoming). */
  upcomingExpiryCount: number;
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
export interface CustomerEusrCategoryCell {
  category: string;
  trainingDate: string | null;
  expiry: string | null;
}

export interface CustomerMatrixRecord {
  id: string;
  /** Workforce list item id when resolved; used for profile links. */
  candidateId: string | null;
  candidateName: string;
  companyName: string | null;
  dateOfBirth: string | null;
  department: string | null;
  trainingManager: string | null;
  supervisor: string | null;
  overallStatus: string | null;
  needsReview: boolean;
  nextExpiryDate: string | null;
  /** Active NPORS category codes / labels for the candidate. */
  nporsCategories: string | null;
  /** Earliest NPORS-related expiry (matrix N* columns and/or register). */
  nporsExpiry: string | null;
  nporsNumber: string | null;
  cscsNumber: string | null;
  cscsExpiry: string | null;
  swqrNumber: string | null;
  swqrExpiry: string | null;
  eusrNumber: string | null;
  eusrExpiry: string | null;
  /** Per-category EUSR training + expiry when categories were completed separately. */
  eusrCategoryRows: CustomerEusrCategoryCell[];
  inHouseCourse: string | null;
  inHouseExpiry: string | null;
  n001Expiry: string | null;
  n003Expiry: string | null;
  n004Expiry: string | null;
  n010Expiry: string | null;
  n020Expiry: string | null;
  n021Expiry: string | null;
  n027Expiry: string | null;
  n100Expiry: string | null;
  /** Full matrix expiry cells keyed by display header. */
  columnValues?: Record<string, string | null>;
  /** Per-header training dates from the matrix TrainingDates JSON. */
  categoryTrainingDates?: Record<string, string | null>;
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
  cardType: string | null;
}

export interface CustomerStreetworksRecord extends CustomerTrainingRecordBase {
  swqrNumber: string | null;
  course: string | null;
  streetworksCategory: string | null;
  /** Multi-day course end date when recorded by admin. */
  trainingDateEnd: string | null;
}

export interface CustomerInHouseRecord extends CustomerTrainingRecordBase {
  course: string | null;
  certificationNumber: string | null;
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
  ulnNumber: string | null;
  cardSchemeCategory: string | null;
  cardExtensionDateNeeded: string | null;
}

export interface CustomerDocumentRecord {
  id: string;
  name: string;
  documentType: string | null;
  candidate: string | null;
  candidateId?: string | null;
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
  category: string | null;
  description: string | null;
  startDate: string | null;
  endDate: string | null;
  status: string | null;
  image: string | null;
  ctaLabel: string | null;
  ctaLink: string | null;
}
