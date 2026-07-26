import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { listAuditLogs } from "@/lib/services/auditLogService";
import {
  getAllCompanies,
  getCompanyById,
  isActiveCompanyStatus,
} from "@/lib/services/companyService";
import { COMPANY_LEVEL_FOLDERS } from "@/lib/services/customerDocumentsFolderService";
import {
  asBoolean,
  asLookupOrString,
  asNullableString,
  asString,
  extractLookupId,
  getListItemsByKey,
  type SharePointFields,
  type SharePointListItem,
} from "@/lib/services/sharePointListService";
import { getExpiryStatus } from "@/lib/training/expiryFilters";
import type {
  AdminDashboardActivityRow,
  AdminDashboardDocumentRow,
  AdminDashboardEventRow,
  AdminDashboardExpiryRow,
  AdminDashboardPayload,
  AdminDashboardStats,
  AdminDataWarning,
  AdminWarningIssue,
} from "@/types/models";

const workforceFields = getSharePointFields("workforce");
const matrixFields = getSharePointFields("trainingMatrix");
const nporsFields = getSharePointFields("nporsRegister");
const eusrFields = getSharePointFields("eusrRegister");
const streetworksFields = getSharePointFields("nrswaRegister");
const inHouseFields = getSharePointFields("inHouseCertificates");
const nvqFields = getSharePointFields("nvqRegister");
const documentFields = getSharePointFields("customerDocuments");
const eventFields = getSharePointFields("events");
const permissionFields = getSharePointFields("permissions");

const FOLDER_NAME_BLOCKLIST = new Set(
  [
    ...COMPANY_LEVEL_FOLDERS,
    "NVQ Documents",
    "Training Documents",
    "Certificates",
    "Forms",
    "Card Scans",
    "Other Documents",
  ].map((name) => name.trim().toLowerCase()),
);

function matchesCompanyFilter(
  value: string | null | undefined,
  companyName: string | null,
): boolean {
  if (!companyName) {
    return true;
  }
  return (value ?? "").trim().toLowerCase() === companyName.trim().toLowerCase();
}

function companyFromFields(
  fields: SharePointFields,
  companyNameById: Map<string, string>,
  ...keys: string[]
): string | null {
  for (const key of keys) {
    const value = asLookupOrString(fields[key]);
    if (value) return value;
    const lookupId = extractLookupId(fields, key);
    if (lookupId) {
      const resolved = companyNameById.get(lookupId);
      if (resolved) return resolved;
    }
  }
  return null;
}

function candidateFromFields(
  fields: SharePointFields,
  workforceNameById: Map<string, string>,
  ...keys: string[]
): string | null {
  for (const key of keys) {
    const value = asLookupOrString(fields[key]) ?? asString(fields[key]);
    if (value) return value;
    const lookupId = extractLookupId(fields, key);
    if (lookupId) {
      const resolved = workforceNameById.get(lookupId);
      if (resolved) return resolved;
    }
  }
  return null;
}

function pushWarning(
  warnings: AdminDataWarning[],
  item: {
    id: string;
    source: string;
    candidateName: string | null;
    issues: AdminWarningIssue[];
  },
) {
  if (item.issues.length === 0 || warnings.length >= 25) {
    return;
  }

  warnings.push({
    id: `${item.source}-${item.id}`,
    source: item.source,
    candidateName: item.candidateName,
    issues: item.issues,
    detail: `Missing ${item.issues
      .map((issue) => {
        if (issue === "CompanyName") return "company";
        if (issue === "CustomerVisible") return "customer visibility";
        if (issue === "TrainingAddress") return "training address";
        return issue;
      })
      .join(", ")}`,
  });
}

function collectTrainingWarnings(
  warnings: AdminDataWarning[],
  source: string,
  items: SharePointListItem[],
  companyName: string | null,
  companyNameById: Map<string, string>,
  workforceNameById: Map<string, string>,
  map: {
    companyField: string;
    candidateField: string;
    addressField?: string;
    visibleField?: string;
  },
) {
  for (const item of items) {
    const rowCompany = companyFromFields(
      item.fields,
      companyNameById,
      map.companyField,
    );
    if (!matchesCompanyFilter(rowCompany, companyName)) {
      continue;
    }

    const issues: AdminWarningIssue[] = [];
    if (!rowCompany) {
      issues.push("CompanyName");
    }
    if (map.visibleField && !asBoolean(item.fields[map.visibleField])) {
      const raw = item.fields[map.visibleField];
      if (raw === undefined || raw === null || raw === "") {
        issues.push("CustomerVisible");
      }
    }
    if (map.addressField && !asString(item.fields[map.addressField])) {
      issues.push("TrainingAddress");
    }

    pushWarning(warnings, {
      id: item.id,
      source,
      candidateName: candidateFromFields(
        item.fields,
        workforceNameById,
        map.candidateField,
      ),
      issues,
    });
  }
}

function isSharePointFolder(fields: SharePointFields): boolean {
  const fs = fields[documentFields.fsObjType];
  if (fs === 1 || fs === "1") return true;
  const contentType = asString(fields.ContentType) ?? "";
  if (/folder/i.test(contentType)) return true;
  const fileRef = asNullableString(fields[documentFields.fileRef]) ?? "";
  if (fileRef.endsWith("/")) return true;
  return false;
}

function isNoiseDocumentName(name: string | null | undefined): boolean {
  const normalized = (name ?? "").trim().toLowerCase();
  if (!normalized) return true;
  return FOLDER_NAME_BLOCKLIST.has(normalized);
}

function expiryTone(
  status: ReturnType<typeof getExpiryStatus>["status"],
): AdminDashboardExpiryRow["statusTone"] {
  if (status === "expired" || status === "urgent") return "danger";
  if (status === "upcoming") return "warn";
  if (status === "valid") return "ok";
  return "missing";
}

function formatActivityTitle(
  action: string,
  entityType: string,
  entityName: string | null,
): string {
  const cleanAction = action.replace(/_/g, " ").trim();
  if (entityName?.trim()) {
    return `${cleanAction} · ${entityName.trim()}`.slice(0, 120);
  }
  if (entityType.trim() && entityType !== "Unknown") {
    return `${cleanAction} · ${entityType}`.slice(0, 120);
  }
  return cleanAction.slice(0, 120) || "Portal event";
}

/**
 * Aggregates admin dashboard metrics from SharePoint.
 * Optional companyId filters operational stats to one company (admins only).
 */
export async function getAdminDashboard(
  selectedCompanyId?: string | null,
): Promise<AdminDashboardPayload> {
  const companies = await getAllCompanies();
  const selectedCompany = selectedCompanyId
    ? ((await getCompanyById(selectedCompanyId)) ??
      companies.find((company) => company.id === selectedCompanyId) ??
      null)
    : null;
  const companyName = selectedCompany?.companyName ?? null;
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );

  const [
    workforce,
    matrix,
    npors,
    eusr,
    streetworks,
    inHouse,
    nvq,
    documents,
    events,
    permissions,
    auditLogs,
  ] = await Promise.all([
    getListItemsByKey("workforce", { top: 5000 }),
    getListItemsByKey("trainingMatrix", { top: 5000 }),
    getListItemsByKey("nporsRegister", { top: 5000 }),
    getListItemsByKey("eusrRegister", { top: 5000 }),
    getListItemsByKey("nrswaRegister", { top: 5000 }),
    getListItemsByKey("inHouseCertificates", { top: 5000 }),
    getListItemsByKey("nvqRegister", { top: 5000 }),
    getListItemsByKey("customerDocuments", { top: 5000 }),
    getListItemsByKey("events", { top: 5000 }),
    getListItemsByKey("permissions", { top: 5000 }).catch(() => []),
    listAuditLogs({ top: 40 }).catch(() => []),
  ]);

  const workforceNameById = new Map<string, string>();
  const workforceCompanyById = new Map<string, string>();
  for (const item of workforce) {
    const name =
      asString(item.fields[workforceFields.candidateName]) ??
      asLookupOrString(item.fields[workforceFields.candidateName]);
    if (!name) continue;
    workforceNameById.set(item.id, name);
    const company = companyFromFields(
      item.fields,
      companyNameById,
      workforceFields.companyName,
    );
    if (company) workforceCompanyById.set(item.id, company);
  }

  const filteredWorkforce = workforce.filter((item) =>
    matchesCompanyFilter(
      companyFromFields(
        item.fields,
        companyNameById,
        workforceFields.companyName,
      ),
      companyName,
    ),
  );

  const activeCandidates = filteredWorkforce.filter((item) => {
    const status = (asNullableString(item.fields[workforceFields.status]) ?? "")
      .trim()
      .toLowerCase();
    return !status || status === "active";
  }).length;

  const filteredMatrix = matrix.filter((item) => {
    const candidateId = extractLookupId(item.fields, matrixFields.candidateName);
    const rowCompany =
      companyFromFields(
        item.fields,
        companyNameById,
        matrixFields.matrixCompany,
        matrixFields.companyName,
      ) ?? (candidateId ? (workforceCompanyById.get(candidateId) ?? null) : null);
    return matchesCompanyFilter(rowCompany, companyName);
  });

  let expiredTraining = 0;
  let expiringWithin3Months = 0;
  let expiringWithin6Months = 0;
  let recordsToReview = 0;
  const upcomingExpiries: AdminDashboardExpiryRow[] = [];

  for (const item of filteredMatrix) {
    const nextExpiry = asNullableString(item.fields[matrixFields.nextExpiryDate]);
    const expiry = getExpiryStatus(nextExpiry);
    if (expiry.status === "expired") {
      expiredTraining += 1;
    }
    if (expiry.status === "urgent") {
      expiringWithin3Months += 1;
    }
    if (expiry.status === "upcoming") {
      expiringWithin6Months += 1;
    }
    if (asBoolean(item.fields[matrixFields.needsReview])) {
      recordsToReview += 1;
    }
    if (
      expiry.status === "expired" ||
      expiry.status === "urgent" ||
      expiry.status === "upcoming"
    ) {
      const candidateId = extractLookupId(
        item.fields,
        matrixFields.candidateName,
      );
      const candidateName = candidateFromFields(
        item.fields,
        workforceNameById,
        matrixFields.candidateName,
      );
      // Skip unresolved rows — never show "Unknown" placeholders.
      if (!candidateName) continue;

      upcomingExpiries.push({
        id: item.id,
        candidateName,
        companyName:
          companyFromFields(
            item.fields,
            companyNameById,
            matrixFields.matrixCompany,
            matrixFields.companyName,
          ) ??
          (candidateId
            ? (workforceCompanyById.get(candidateId) ?? null)
            : null),
        nextExpiryDate: nextExpiry,
        statusLabel: expiry.label,
        statusTone: expiryTone(expiry.status),
      });
    }
  }

  upcomingExpiries.sort((a, b) => {
    const aTime = a.nextExpiryDate
      ? new Date(a.nextExpiryDate).getTime()
      : Number.POSITIVE_INFINITY;
    const bTime = b.nextExpiryDate
      ? new Date(b.nextExpiryDate).getTime()
      : Number.POSITIVE_INFINITY;
    return aTime - bTime;
  });

  const filteredNvq = nvq.filter((item) =>
    matchesCompanyFilter(
      companyFromFields(
        item.fields,
        companyNameById,
        nvqFields.companyName,
        nvqFields.nvqCompany,
      ),
      companyName,
    ),
  );

  let activeNvqs = 0;
  let completedNvqs = 0;
  for (const item of filteredNvq) {
    const completed = asNullableString(item.fields[nvqFields.completedDate]);
    if (completed?.trim()) {
      completedNvqs += 1;
    } else {
      activeNvqs += 1;
    }
  }

  const now = Date.now();
  const thirtyDaysAgo = now - 30 * 86_400_000;

  const filteredDocuments = documents.filter((item) => {
    if (isSharePointFolder(item.fields)) return false;
    const name =
      asString(item.fields[documentFields.fileLeafRef]) ??
      asString(item.fields[documentFields.title]);
    if (isNoiseDocumentName(name)) return false;
    const rowCompany = companyFromFields(
      item.fields,
      companyNameById,
      documentFields.company,
    );
    return matchesCompanyFilter(rowCompany, companyName);
  });

  const documentsPendingVisibility = filteredDocuments.filter(
    (item) => !asBoolean(item.fields[documentFields.customerVisible]),
  ).length;

  const recentDocuments: AdminDashboardDocumentRow[] = filteredDocuments
    .map((item) => {
      const modifiedDate =
        item.lastModifiedDateTime ??
        asNullableString(item.fields[documentFields.modified]) ??
        item.createdDateTime ??
        null;
      const name =
        asString(item.fields[documentFields.fileLeafRef]) ??
        asString(item.fields[documentFields.title]);
      if (!name || isNoiseDocumentName(name)) return null;
      return {
        id: item.id,
        name,
        company: companyFromFields(
          item.fields,
          companyNameById,
          documentFields.company,
        ),
        candidate: candidateFromFields(
          item.fields,
          workforceNameById,
          documentFields.candidate,
        ),
        modifiedDate,
        customerVisible: asBoolean(item.fields[documentFields.customerVisible]),
        _sort: modifiedDate ? new Date(modifiedDate).getTime() : 0,
      };
    })
    .filter((row): row is NonNullable<typeof row> => row !== null)
    .sort((a, b) => b._sort - a._sort)
    .slice(0, 8)
    .map(({ _sort: _ignored, ...row }) => row);

  const documentsUploadedRecently = filteredDocuments.filter((item) => {
    const modified =
      item.lastModifiedDateTime ??
      asNullableString(item.fields[documentFields.modified]) ??
      item.createdDateTime;
    if (!modified) return false;
    const time = new Date(modified).getTime();
    return !Number.isNaN(time) && time >= thirtyDaysAgo;
  }).length;

  const upcomingBookings: AdminDashboardEventRow[] = events
    .map((item) => {
      const eventDate = asNullableString(item.fields[eventFields.eventDate]);
      const time = eventDate ? new Date(eventDate).getTime() : NaN;
      return {
        id: item.id,
        title: asString(item.fields[eventFields.title]) ?? "Event",
        company: companyFromFields(
          item.fields,
          companyNameById,
          eventFields.eventCompany,
        ),
        eventDate,
        location: asNullableString(item.fields[eventFields.location]),
        _time: time,
      };
    })
    .filter((row) => {
      if (!matchesCompanyFilter(row.company, companyName)) return false;
      return !Number.isNaN(row._time) && row._time >= now;
    })
    .sort((a, b) => a._time - b._time)
    .slice(0, 8)
    .map(({ _time: _ignored, ...row }) => row);

  const upcomingEvents = upcomingBookings.length;

  const accessInvitationsPending = permissions.filter((item) => {
    const rowCompany = companyFromFields(
      item.fields,
      companyNameById,
      permissionFields.company,
    );
    if (!matchesCompanyFilter(rowCompany, companyName)) return false;
    const status = (asNullableString(item.fields[permissionFields.status]) ?? "")
      .trim()
      .toLowerCase();
    return status === "pending" || status === "invited" || status === "inactive";
  }).length;

  const recentActivity: AdminDashboardActivityRow[] = auditLogs
    .filter((row) => {
      const action = (row.action ?? "").trim();
      if (!action || action.toLowerCase() === "activity") return false;
      if (
        !row.userEmail ||
        row.userEmail === "unknown" ||
        row.userEmail === "system"
      ) {
        if (!row.entityType || row.entityType === "Unknown") return false;
      }
      if (companyName && row.company) {
        return matchesCompanyFilter(row.company, companyName);
      }
      return true;
    })
    .slice(0, 8)
    .map((row) => ({
      id: row.id,
      title: formatActivityTitle(row.action, row.entityType, row.entityName),
      userEmail: row.userEmail || null,
      timestamp: row.timestamp,
      detail: row.errorMessage,
    }));

  const stats: AdminDashboardStats = {
    totalCompanies: companies.length,
    activeCompanies: companies.filter((company) =>
      isActiveCompanyStatus(company.status),
    ).length,
    totalCandidates: filteredWorkforce.length,
    activeCandidates,
    expiredTraining,
    expiringWithin3Months,
    expiringWithin6Months,
    recordsToReview,
    activeNvqs,
    completedNvqs,
    documentsPendingVisibility,
    documentsUploadedRecently,
    upcomingEvents,
    accessInvitationsPending,
  };

  const warnings: AdminDataWarning[] = [];
  collectTrainingWarnings(
    warnings,
    "NPORS",
    npors,
    companyName,
    companyNameById,
    workforceNameById,
    {
      companyField: nporsFields.companyName,
      candidateField: nporsFields.candidateName,
      addressField: nporsFields.trainingAddress,
      visibleField: nporsFields.customerVisible,
    },
  );
  collectTrainingWarnings(
    warnings,
    "EUSR",
    eusr,
    companyName,
    companyNameById,
    workforceNameById,
    {
      companyField: eusrFields.companyName,
      candidateField: eusrFields.candidateName,
      addressField: eusrFields.trainingAddress,
      visibleField: eusrFields.customerVisible,
    },
  );
  collectTrainingWarnings(
    warnings,
    "Streetworks",
    streetworks,
    companyName,
    companyNameById,
    workforceNameById,
    {
      companyField: streetworksFields.companyName,
      candidateField: streetworksFields.candidateName,
      addressField: streetworksFields.trainingAddress,
      visibleField: streetworksFields.customerVisible,
    },
  );
  collectTrainingWarnings(
    warnings,
    "In-House",
    inHouse,
    companyName,
    companyNameById,
    workforceNameById,
    {
      companyField: inHouseFields.companyName,
      candidateField: inHouseFields.candidateName,
      addressField: inHouseFields.trainingAddress,
      visibleField: inHouseFields.customerVisible,
    },
  );

  for (const item of filteredMatrix) {
    const issues: AdminWarningIssue[] = [];
    const candidateId = extractLookupId(item.fields, matrixFields.candidateName);
    const rowCompany =
      companyFromFields(
        item.fields,
        companyNameById,
        matrixFields.matrixCompany,
        matrixFields.companyName,
      ) ?? (candidateId ? (workforceCompanyById.get(candidateId) ?? null) : null);
    if (!rowCompany) {
      issues.push("CompanyName");
    }
    pushWarning(warnings, {
      id: item.id,
      source: "Training Matrix",
      candidateName: candidateFromFields(
        item.fields,
        workforceNameById,
        matrixFields.candidateName,
      ),
      issues,
    });
  }

  return {
    selectedCompanyId: selectedCompany?.id ?? null,
    selectedCompanyName: selectedCompany?.companyName ?? null,
    stats,
    warnings,
    upcomingExpiries: upcomingExpiries.slice(0, 8),
    recentDocuments,
    upcomingBookings,
    recentActivity,
  };
}
