import { redirect } from "next/navigation";

import { CustomerDashboardView } from "@/components/customer/CustomerDashboardView";
import { auth } from "@/auth";
import {
  getCompanyById,
  toCustomerCompanyProfile,
} from "@/lib/services/companyService";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerDashboardContent } from "@/lib/services/customerDashboardService";
import type {
  CustomerCompanyProfile,
  CustomerContext,
  CustomerEventRecord,
  CustomerOfferRecord,
  DashboardStats,
} from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerDashboardPage() {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();

  if (!email) {
    redirect("/login");
  }

  let context: CustomerContext;
  try {
    context = await getCustomerContext(email);
  } catch {
    redirect("/access-denied");
  }

  // These two reads are independent — run them concurrently so the dashboard
  // waits for the slower of the two, not the sum of both.
  const [companyResult, statsResult] = await Promise.allSettled([
    getCompanyById(context.companyId),
    getCustomerDashboardContent(context),
  ]);

  let companyProfile: CustomerCompanyProfile | null = null;
  if (companyResult.status === "fulfilled" && companyResult.value) {
    companyProfile = toCustomerCompanyProfile(companyResult.value);
  }

  let stats: DashboardStats;
  let offers: CustomerOfferRecord[] = [];
  let upcomingEvents: CustomerEventRecord[] = [];
  if (statsResult.status === "fulfilled") {
    stats = statsResult.value.stats;
    offers = statsResult.value.offers;
    upcomingEvents = statsResult.value.upcomingEvents;
  } else {
    stats = {
      workforceCount: 0,
      trainingMatrixCount: 0,
      needsReviewCount: 0,
      expiringSoonCount: 0,
      upcomingExpiryCount: 0,
      expiredCount: 0,
      documentsCount: 0,
      upcomingEventsCount: 0,
      activeOffersCount: 0,
      nporsCount: 0,
      eusrCount: 0,
      streetworksCount: 0,
      inHouseCount: 0,
      nvqCount: 0,
    };
  }

  return (
    <CustomerDashboardView
      companyName={context.companyName}
      stats={stats}
      offers={offers}
      upcomingEvents={upcomingEvents}
      companyProfile={companyProfile}
    />
  );
}
