import { AdminDashboardView } from "@/components/admin/AdminDashboardView";
import { getAdminDashboard } from "@/lib/services/adminDashboardService";
import { getAllCompanies } from "@/lib/services/companyService";
import { auth } from "@/auth";

export const dynamic = "force-dynamic";

export default async function AdminDashboardPage({
  searchParams,
}: {
  searchParams: Promise<{ companyId?: string }>;
}) {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase() ?? "";
  const params = await searchParams;
  const companyId = params.companyId?.trim() || null;

  const [companies, dashboard] = await Promise.all([
    getAllCompanies(),
    getAdminDashboard(companyId),
  ]);

  return (
    <AdminDashboardView
      email={email}
      companies={companies}
      dashboard={dashboard}
    />
  );
}
