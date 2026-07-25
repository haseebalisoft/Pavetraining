import "server-only";

import { listAdminMatrix } from "@/lib/services/adminCrudService";
import {
  getCustomerDocumentRecords,
  getCustomerEventRecords,
  getCustomerNvqRecords,
  getCustomerOfferRecords,
} from "@/lib/services/customerPortalService";
import {
  getCustomerEusrRecords,
  getCustomerInHouseRecords,
  getCustomerNporsRecords,
  getCustomerStreetworksRecords,
} from "@/lib/services/customerTrainingRecordsService";
import { daysUntilExpiry } from "@/lib/training/expiryFilters";
import { getWorkforceByCompanyName } from "@/lib/services/workforceService";
import type {
  CustomerMatrixRecord,
  DashboardStats,
} from "@/types/models";

/**
 * Customer-facing matrix rows for a company.
 * Omits admin-only review notes noise; keeps needsReview for the dashboard.
 */
export async function getCustomerMatrixRecords(
  companyName: string,
): Promise<CustomerMatrixRecord[]> {
  const rows = await listAdminMatrix(companyName);
  return rows.map((row) => ({
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
}

export async function getCustomerDashboard(
  companyId: string,
  companyName: string,
): Promise<DashboardStats> {
  const [
    workforce,
    matrix,
    documents,
    events,
    offers,
    nvq,
    npors,
    eusr,
    streetworks,
    inHouse,
  ] = await Promise.all([
    getWorkforceByCompanyName(companyName),
    getCustomerMatrixRecords(companyName),
    getCustomerDocumentRecords(companyId, false),
    getCustomerEventRecords(companyId),
    getCustomerOfferRecords(companyId),
    getCustomerNvqRecords(companyId),
    getCustomerNporsRecords(companyId),
    getCustomerEusrRecords(companyId),
    getCustomerStreetworksRecords(companyId),
    getCustomerInHouseRecords(companyId),
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
    nporsCount: npors.length,
    eusrCount: eusr.length,
    streetworksCount: streetworks.length,
    inHouseCount: inHouse.length,
    nvqCount: nvq.length,
  };
}
