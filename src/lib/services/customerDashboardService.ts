import "server-only";

import { listAdminMatrix } from "@/lib/services/adminCrudService";
import {
  filterRowsByCandidateAccess,
  getAllowedCandidateNames,
  getAllowedWorkforceForCustomer,
} from "@/lib/services/customerAccessService";
import {
  getCustomerDocumentRecords,
  getCustomerEventRecords,
  getCustomerNvqRecords,
  getCustomerOfferRecords,
} from "@/lib/services/customerPortalService";
import { daysUntilExpiry } from "@/lib/training/expiryFilters";
import type {
  CustomerContext,
  CustomerMatrixRecord,
  DashboardStats,
} from "@/types/models";

/**
 * Customer-facing matrix rows for a company.
 * Omits admin-only review notes noise; keeps needsReview for the dashboard.
 * Applies Supervisor / Candidate access scope when context is provided.
 */
export async function getCustomerMatrixRecords(
  companyName: string,
  context?: CustomerContext,
): Promise<CustomerMatrixRecord[]> {
  const rows = await listAdminMatrix(companyName);
  let mapped: CustomerMatrixRecord[] = rows.map((row) => ({
    id: row.id,
    candidateName: row.candidateName,
    department: row.department,
    overallStatus: row.overallStatus,
    needsReview: row.needsReview,
    nextExpiryDate: row.nextExpiryDate,
    n001Expiry: row.n001Expiry,
    n003Expiry: row.n003Expiry,
    n004Expiry: row.n004Expiry,
    n010Expiry: row.n010Expiry,
    n020Expiry: row.n020Expiry,
    n021Expiry: row.n021Expiry,
    n027Expiry: row.n027Expiry,
    n100Expiry: row.n100Expiry,
  }));

  if (context) {
    const allowedNames = await getAllowedCandidateNames(context);
    mapped = filterRowsByCandidateAccess(mapped, allowedNames, context);
  }

  return mapped;
}

/**
 * Dashboard stats — only loads lists needed for the cards.
 * Training-register counts are deferred (open Training Records for detail)
 * so first paint does not wait on NPORS/EUSR/Streetworks/In-House.
 */
export async function getCustomerDashboard(
  context: CustomerContext,
): Promise<DashboardStats> {
  const [workforce, matrix, documents, events, offers, nvq] = await Promise.all([
    getAllowedWorkforceForCustomer(context),
    getCustomerMatrixRecords(context.companyName, context),
    getCustomerDocumentRecords(context.companyId, false, context),
    getCustomerEventRecords(context.companyId),
    getCustomerOfferRecords(context.companyId),
    getCustomerNvqRecords(context.companyId, context),
  ]);

  let needsReviewCount = 0;
  let expiringSoonCount = 0;
  let expiredCount = 0;

  for (const row of matrix) {
    if (row.needsReview) {
      needsReviewCount += 1;
    }
    const days = daysUntilExpiry(row.nextExpiryDate);
    if (days !== null && days < 0) {
      expiredCount += 1;
    }
    if (days !== null && days >= 0 && days <= 90) {
      expiringSoonCount += 1;
    }
  }

  const now = Date.now();
  const upcomingEventsCount = events.filter((event) => {
    if (!event.eventDate) return false;
    const time = new Date(event.eventDate).getTime();
    return !Number.isNaN(time) && time >= now;
  }).length;

  return {
    workforceCount: workforce.length,
    trainingMatrixCount: matrix.length,
    needsReviewCount,
    expiringSoonCount,
    expiredCount,
    documentsCount: documents.length,
    upcomingEventsCount,
    activeOffersCount: offers.length,
    // Detailed register counts live on Training Records pages.
    nporsCount: 0,
    eusrCount: 0,
    streetworksCount: 0,
    inHouseCount: 0,
    nvqCount: nvq.length,
  };
}
