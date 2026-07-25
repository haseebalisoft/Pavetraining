import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  getAllCompanies,
  getCompanyById,
  isActiveCompanyStatus,
} from "@/lib/services/companyService";
import {
  asBoolean,
  asLookupOrString,
  asNullableString,
  asString,
  getListItemsByKey,
  type SharePointFields,
  type SharePointListItem,
} from "@/lib/services/sharePointListService";
import { daysUntilExpiry } from "@/lib/training/expiryFilters";
import type {
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
  ...keys: string[]
): string | null {
  for (const key of keys) {
    const value = asLookupOrString(fields[key]);
    if (value) {
      return value;
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
  map: {
    companyField: string;
    candidateField: string;
    addressField?: string;
    visibleField?: string;
  },
) {
  for (const item of items) {
    const rowCompany = companyFromFields(item.fields, map.companyField);
    if (!matchesCompanyFilter(rowCompany, companyName)) {
      continue;
    }

    const issues: AdminWarningIssue[] = [];
    if (!rowCompany) {
      issues.push("CompanyName");
    }
    if (map.visibleField && !asBoolean(item.fields[map.visibleField])) {
      // Missing or false CustomerVisible both warrant a review warning when field exists
      const raw = item.fields[map.visibleField];
      if (raw === undefined || raw === null || raw === "") {
        issues.push("CustomerVisible");
      }
    }
    if (
      map.addressField &&
      !asString(item.fields[map.addressField])
    ) {
      issues.push("TrainingAddress");
    }

    pushWarning(warnings, {
      id: item.id,
      source,
      candidateName: asNullableString(item.fields[map.candidateField]),
      issues,
    });
  }
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
  ]);

  const filteredWorkforce = workforce.filter((item) =>
    matchesCompanyFilter(
      companyFromFields(item.fields, workforceFields.companyName),
      companyName,
    ),
  );

  const filteredMatrix = matrix.filter((item) =>
    matchesCompanyFilter(
      companyFromFields(
        item.fields,
        matrixFields.companyName,
        matrixFields.matrixCompany,
      ),
      companyName,
    ),
  );

  let expiredTraining = 0;
  let expiringWithin3Months = 0;
  let recordsToReview = 0;

  for (const item of filteredMatrix) {
    const nextExpiry = asNullableString(item.fields[matrixFields.nextExpiryDate]);
    const days = daysUntilExpiry(nextExpiry);
    if (days !== null && days < 0) {
      expiredTraining += 1;
    }
    if (days !== null && days >= 0 && days <= 90) {
      expiringWithin3Months += 1;
    }
    if (asBoolean(item.fields[matrixFields.needsReview])) {
      recordsToReview += 1;
    }
  }

  const filteredNvq = nvq.filter((item) =>
    matchesCompanyFilter(
      companyFromFields(
        item.fields,
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

  const documentsPendingVisibility = documents.filter((item) => {
    const rowCompany = companyFromFields(item.fields, documentFields.company);
    if (!matchesCompanyFilter(rowCompany, companyName)) {
      return false;
    }
    return !asBoolean(item.fields[documentFields.customerVisible]);
  }).length;

  const now = Date.now();
  const upcomingEvents = events.filter((item) => {
    const rowCompany = companyFromFields(item.fields, eventFields.eventCompany);
    if (!matchesCompanyFilter(rowCompany, companyName)) {
      return false;
    }
    if (!asBoolean(item.fields[eventFields.customerVisible])) {
      // Still count admin-visible upcoming events even if not customer visible
    }
    const eventDate = asNullableString(item.fields[eventFields.eventDate]);
    if (!eventDate) {
      return false;
    }
    const time = new Date(eventDate).getTime();
    return !Number.isNaN(time) && time >= now;
  }).length;

  const stats: AdminDashboardStats = {
    totalCompanies: companies.length,
    activeCompanies: companies.filter((company) =>
      isActiveCompanyStatus(company.status),
    ).length,
    totalCandidates: filteredWorkforce.length,
    expiredTraining,
    expiringWithin3Months,
    recordsToReview,
    activeNvqs,
    completedNvqs,
    documentsPendingVisibility,
    upcomingEvents,
  };

  const warnings: AdminDataWarning[] = [];
  collectTrainingWarnings(warnings, "NPORS", npors, companyName, {
    companyField: nporsFields.companyName,
    candidateField: nporsFields.candidateName,
    addressField: nporsFields.trainingAddress,
    visibleField: nporsFields.customerVisible,
  });
  collectTrainingWarnings(warnings, "EUSR", eusr, companyName, {
    companyField: eusrFields.companyName,
    candidateField: eusrFields.candidateName,
    addressField: eusrFields.trainingAddress,
    visibleField: eusrFields.customerVisible,
  });
  collectTrainingWarnings(warnings, "Streetworks", streetworks, companyName, {
    companyField: streetworksFields.companyName,
    candidateField: streetworksFields.candidateName,
    addressField: streetworksFields.trainingAddress,
    visibleField: streetworksFields.customerVisible,
  });
  collectTrainingWarnings(warnings, "In-House", inHouse, companyName, {
    companyField: inHouseFields.companyName,
    candidateField: inHouseFields.candidateName,
    addressField: inHouseFields.trainingAddress,
    visibleField: inHouseFields.customerVisible,
  });

  for (const item of filteredMatrix) {
    const issues: AdminWarningIssue[] = [];
    const rowCompany = companyFromFields(
      item.fields,
      matrixFields.companyName,
      matrixFields.matrixCompany,
    );
    if (!rowCompany) {
      issues.push("CompanyName");
    }
    pushWarning(warnings, {
      id: item.id,
      source: "Training Matrix",
      candidateName: asNullableString(item.fields[matrixFields.candidateName]),
      issues,
    });
  }

  return {
    selectedCompanyId: selectedCompany?.id ?? null,
    selectedCompanyName: selectedCompany?.companyName ?? null,
    stats,
    warnings,
  };
}
